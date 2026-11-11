"""Evento UPS: importación autorizada, credenciales y control de acceso.
NO integra automáticamente el sistema de inscripciones de la universidad.
"""
import base64, csv, hashlib, io, os, re, secrets, unicodedata, json, html
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path
from urllib.parse import quote

import qrcode
import httpx
from fastapi import FastAPI, File, Header, HTTPException, UploadFile, BackgroundTasks, Request
from fastapi.responses import FileResponse, StreamingResponse, Response
from openpyxl import load_workbook
from pydantic import BaseModel, Field
from sqlalchemy import create_engine, MetaData, Table, Column, Integer, String, Text, UniqueConstraint, select, func, update, insert, delete, and_, inspect, text
from sqlalchemy.exc import IntegrityError

ROOT = Path(__file__).resolve().parent
EVENT_ID = os.getenv('QR_EVENTO_ID','ups-2207')
EVENT = 'Conferencia: Software Libre, Libertad Digital y su Relación con el Machine Learning'
URL_UPS = 'https://appwfp.ups.edu.ec/foc-webins/index.xhtml?prog=2207'
ADMIN_KEY = os.getenv('QR_ADMIN_KEY','')
DB_URL = os.getenv('DATABASE_URL',f'sqlite:///{(ROOT / "asistencia.db").as_posix()}')
# Render entrega URLs postgres:// o postgresql://; usar el controlador psycopg moderno.
if DB_URL.startswith('postgres://'):
    DB_URL = 'postgresql+psycopg://' + DB_URL[len('postgres://'):]
elif DB_URL.startswith('postgresql://'):
    DB_URL = 'postgresql+psycopg://' + DB_URL[len('postgresql://'):]
engine = create_engine(DB_URL, pool_pre_ping=True, pool_recycle=300)

meta = MetaData()
people = Table('participantes', meta,
    Column('id', Integer, primary_key=True, autoincrement=True),
    Column('evento_id', String(80), nullable=False), Column('cedula', String(50), nullable=False),
    Column('nombre', String(200), nullable=False), Column('email', String(255), nullable=False, default=''),
    Column('telefono', String(60), nullable=False, default=''), Column('tipo', String(100), nullable=False, default=''),
    Column('token_hash', String(64), nullable=False, unique=True), Column('token', String(120), nullable=False),
    Column('registrado_en', String(40), nullable=False), Column('asistio_en', String(40)),
    Column('correo_enviado_en', String(40)),
    Column('email_estado', String(30), nullable=False, server_default='pendiente'),
    Column('email_error', Text),
    Column('email_message_id', String(255)),
    Column('email_intentos', Integer, nullable=False, server_default='0'),
    UniqueConstraint('evento_id','cedula', name='uq_evento_cedula'))
attempts = Table('envios', meta,
    Column('id', Integer, primary_key=True, autoincrement=True),
    Column('participante_id', Integer, nullable=False),
    Column('evento_id', String(80), nullable=False),
    Column('destinatario', String(255), nullable=False),
    Column('estado', String(30), nullable=False),
    Column('fecha_utc', String(40), nullable=False),
    Column('message_id', String(255)),
    Column('detalle', Text),
)
meta.create_all(engine)
# Migración compatible con la tabla creada por la versión v3: preserva los QR y asistencia.
old_columns = {col['name'] for col in inspect(engine).get_columns('participantes')}
new_columns = {
    'email_estado': "VARCHAR(30) NOT NULL DEFAULT 'pendiente'",
    'email_error': 'TEXT',
    'email_message_id': 'VARCHAR(255)',
    'email_intentos': 'INTEGER NOT NULL DEFAULT 0',
}
with engine.begin() as conn:
    for name, ddl in new_columns.items():
        if name not in old_columns:
            conn.execute(text(f'ALTER TABLE participantes ADD COLUMN {name} {ddl}'))
    conn.execute(text("UPDATE participantes SET email_estado='aceptado' WHERE correo_enviado_en IS NOT NULL AND email_estado='pendiente'"))

app = FastAPI(title='Evento UPS · Gestión de QR', version='4.0')

def now(): return datetime.now(timezone.utc).isoformat(timespec='seconds')
def guard(key):
    if not ADMIN_KEY or not key or not secrets.compare_digest(ADMIN_KEY,key):
        raise HTTPException(401,'Acceso no autorizado: configura QR_ADMIN_KEY y proporciona la clave.')
def norm(s):
    s=unicodedata.normalize('NFKD',str(s or '')).encode('ascii','ignore').decode().lower().strip()
    return re.sub('[^a-z0-9]','',s)
def getfield(row,*names):
    for n in names:
        v=row.get(norm(n))
        if v is not None and str(v).strip(): return str(v).strip()
    return ''
def qr_payload(token): return f'EVT1:{EVENT_ID}:{token}'
def qr_bytes(token):
    img=qrcode.make(qr_payload(token)); out=io.BytesIO(); img.save(out,'PNG');out.seek(0);return out

def asrow(row): return dict(row._mapping)
def get_person(cid):
    with engine.connect() as c:
        row=c.execute(select(people).where(and_(people.c.id==cid,people.c.evento_id==EVENT_ID))).first()
    if row is None: raise HTTPException(404,'Participante no encontrado')
    return asrow(row)

BREVO_URL = 'https://api.brevo.com/v3/smtp/email'
MAX_DAILY = max(1, min(int(os.getenv('EMAIL_DAILY_LIMIT', '200')), 10000))


def mail_configured():
    return bool(os.getenv('BREVO_API_KEY') and os.getenv('BREVO_SENDER_EMAIL'))


def is_test_address(value):
    addr=(value or '').strip().lower()
    return (not addr or addr.endswith(('@example.com','@example.org','.test','.invalid')))


def claim_mail(pid, *, allow_retry=False):
    """Reserva atómicamente el envío. Nunca reintenta estados inciertos automáticamente."""
    person=get_person(pid)
    if is_test_address(person['email']):
        return False, 'Correo ficticio o ausente. No se enviará a example.com, example.org, .test o .invalid.'
    with engine.begin() as conn:
        # No permite sobrepasar el techo diario por uso ordinario.
        today = datetime.now(timezone.utc).date().isoformat()
        daily = conn.execute(select(func.count()).select_from(attempts).where(and_(
            attempts.c.evento_id==EVENT_ID, attempts.c.fecha_utc>=today))).scalar_one()
        if daily >= MAX_DAILY:
            return False, 'Límite diario de intentos alcanzado. Continúa mañana o revisa el plan de Brevo.'
        allowed = ['pendiente'] + (['error', 'rebotado'] if allow_retry else [])
        res = conn.execute(update(people).where(and_(
            people.c.id==pid, people.c.evento_id==EVENT_ID,
            people.c.email!='', people.c.email_estado.in_(allowed)
        )).values(email_estado='en_proceso', email_error=None,
                 email_intentos=people.c.email_intentos+1))
        if res.rowcount != 1: return False, 'El participante no tiene correo o ya fue procesado.'
        row=conn.execute(select(people.c.email).where(people.c.id==pid)).first()
        conn.execute(insert(attempts).values(participante_id=pid, evento_id=EVENT_ID,
                        destinatario=row.email, estado='en_proceso', fecha_utc=now()))
    return True, None


def finish_mail(pid, state, detail=None, message_id=None):
    with engine.begin() as conn:
        # Evita pisar un estado definitivo si ya llegó un webhook.
        conn.execute(update(people).where(and_(people.c.id==pid, people.c.email_estado=='en_proceso'))
                     .values(email_estado=state, email_error=(detail or '')[:600] or None,
                             email_message_id=message_id,
                             **({'correo_enviado_en': now()} if state=='aceptado' else {})))
        aid=conn.execute(select(attempts.c.id).where(and_(attempts.c.participante_id==pid,
                      attempts.c.evento_id==EVENT_ID, attempts.c.estado=='en_proceso'))
                         .order_by(attempts.c.id.desc()).limit(1)).scalar_one_or_none()
        if aid:
            conn.execute(update(attempts).where(attempts.c.id==aid).values(
                estado=state, detalle=(detail or '')[:600] or None, message_id=message_id))


def send_brevo(pid):
    person=get_person(pid)
    greeting=html.escape(person['nombre'])
    title=html.escape(EVENT)
    plain=(f"Hola {person['nombre']}:\n\nAdjuntamos tu credencial QR para el evento «{EVENT}».\n"
           '23 de noviembre de 2026, 10:00–13:00. Guayaquil – Centenario.\n'
           'Presenta esta credencial junto con tu documento de identificación.\n'
           'Si no solicitaste tu inscripción, contacta con los organizadores.\n')
    markup=(f'<html><body style="font-family:Arial,sans-serif;color:#172a45">'
            '<div style="max-width:590px;margin:auto;padding:24px;border:1px solid #ddd;border-radius:12px">'
            '<h2 style="color:#183d75">Credencial de ingreso · Evento UPS</h2>'
            f'<p>Hola <b>{greeting}</b>:</p><p>Adjuntamos tu código QR para la conferencia <b>{title}</b>.</p>'
            '<p><b>Fecha:</b> 23 de noviembre de 2026<br><b>Horario:</b> 10:00–13:00<br>'
            '<b>Lugar:</b> Guayaquil – Centenario</p>'
            '<p>Presenta tu código QR y documento de identificación al ingresar.</p>'
            '<p style="font-size:12px;color:#677">Si recibiste esto por error, contacta con la organización.</p>'
            '</div></body></html>')
    payload={
        'sender':{'name':os.getenv('BREVO_SENDER_NAME','Organización del evento'),
                  'email':os.environ['BREVO_SENDER_EMAIL']},
        'to':[{'email':person['email'],'name':person['nombre']}],
        'subject':'Tu código QR · Conferencia UPS · 23 de noviembre de 2026',
        'htmlContent': markup,
        'attachment':[{'name':'credencial_qr.png',
                       'content':base64.b64encode(qr_bytes(person['token']).read()).decode('ascii')}],
        'tags':['evento-ups-2207'],
    }
    try:
        response=httpx.post(BREVO_URL, json=payload,
                            headers={'api-key':os.environ['BREVO_API_KEY'],'accept':'application/json'},
                            timeout=12)
        if response.status_code==201:
            message_id=response.json().get('messageId')
            finish_mail(pid,'aceptado', message_id=message_id)
            return True
        # Solo guarda una descripción, jamás claves ni payloads.
        error=f'Brevo HTTP {response.status_code}. Revisa créditos, remitente y correo.'
        finish_mail(pid,'error', error)
    except (httpx.TimeoutException, httpx.NetworkError):
        # Brevo pudo recibir el mensaje aunque nuestra app no obtuviera respuesta.
        finish_mail(pid,'incierto','No se recibió confirmación. Revisar logs de Brevo antes de reintentar.')
    except Exception:
        finish_mail(pid,'error','Fallo local de envío. Consultar registros del servidor.')
    return False


def process_batch(ids):
    for pid in ids:
        send_brevo(pid)

@app.get('/')
def home(): return FileResponse(ROOT/'static'/'index.html')
@app.get('/escanear')
def scanner(): return FileResponse(ROOT/'static'/'escanear.html')
@app.get('/api/evento')
def event(): return {'evento':EVENT,'evento_id':EVENT_ID,'inscripcion_oficial':URL_UPS,'base_de_datos':'PostgreSQL persistente' if DB_URL.startswith('postgresql+') else ('Azure SQL' if DB_URL.startswith('mssql') else 'SQLite temporal (solo pruebas)')}

@app.post('/api/importar')
async def importer(archivo:UploadFile=File(...), x_admin_key:str|None=Header(None)):
    guard(x_admin_key)
    filename=(archivo.filename or '').lower()
    if not filename.endswith(('.xlsx','.csv')): raise HTTPException(400,'Formato permitido: .xlsx o .csv')
    raw=await archivo.read(8_000_001)
    if len(raw)>8_000_000: raise HTTPException(413,'Máximo 8 MB')
    try:
        if filename.endswith('.xlsx'):
            book=load_workbook(io.BytesIO(raw),read_only=True,data_only=True)
            sheet=book.active
            rows=sheet.iter_rows(values_only=True)
            titles=next(rows,None)
            if not titles: raise ValueError('Excel vacío')
            records=({norm(k):('' if v is None else str(v).strip()) for k,v in zip(titles,values)} for values in rows)
        else:
            try: body=raw.decode('utf-8-sig')
            except UnicodeDecodeError: body=raw.decode('cp1252')
            try: dialect=csv.Sniffer().sniff(body[:2048],delimiters=',;\t')
            except csv.Error: dialect=csv.excel
            reader=csv.DictReader(io.StringIO(body),dialect=dialect)
            records=({norm(k):str(v or '').strip() for k,v in r.items() if k is not None} for r in reader)
        added=existing=updated=protected=invalid=0; errors=[]
        with engine.begin() as conn:
            for idx,row in enumerate(records,start=2):
                doc=getfield(row,'identificacion','cedula','documento','numero de identificacion','numero de documento')
                name=getfield(row,'nombre completo','participante','nombres','nombre')
                surname=getfield(row,'apellidos','apellido')
                if surname and surname.lower() not in name.lower(): name=f'{name} {surname}'.strip()
                email=getfield(row,'correo personal','correo electronico','correo','email','e-mail')
                phone=getfield(row,'celular','telefono','telefono movil','whatsapp','numero de celular')
                typ=getfield(row,'tipo inscripcion','tipo de inscripcion')
                if not (5<=len(doc)<=50 and 2<=len(name)<=200) or (email and not re.fullmatch(r'[^\s@]+@[^\s@]+\.[^\s@]+',email)):
                    invalid+=1
                    if len(errors)<20:errors.append(idx)
                    continue
                old=conn.execute(select(people.c.id,people.c.nombre,people.c.email,people.c.telefono,people.c.tipo,
                    people.c.asistio_en,people.c.email_estado,people.c.email_intentos).where(and_(
                    people.c.evento_id==EVENT_ID,people.c.cedula==doc))).first()
                if old:
                    # No se cambia el destinatario de un QR ya enviado ni una asistencia existente.
                    if old.asistio_en or old.email_estado!='pendiente' or old.email_intentos:
                        protected+=1
                        continue
                    changes={}
                    for field,value in [('nombre',name),('email',email),('telefono',phone),('tipo',typ)]:
                        # Las celdas vacías del Excel no borran datos existentes por accidente.
                        if value and value != getattr(old,field): changes[field]=value
                    if changes:
                        conn.execute(update(people).where(people.c.id==old.id).values(**changes))
                        updated+=1
                    else:
                        existing+=1
                    continue
                token=secrets.token_urlsafe(32)
                conn.execute(insert(people).values(evento_id=EVENT_ID,cedula=doc,nombre=name,email=email,
                    telefono=phone,tipo=typ,token_hash=hashlib.sha256(token.encode()).hexdigest(),token=token,registrado_en=now()))
                added+=1
        return {'mensaje':'Archivo procesado correctamente','nuevos':added,'existentes':existing,'actualizados':updated,'protegidos':protected,'invalidos':invalid,'filas_invalidas':errors}
    except (ValueError,StopIteration) as e: raise HTTPException(400,'No se pudo leer el archivo: '+str(e)) from e


@app.delete('/api/participantes/{pid}')
def remove_person(pid:int, x_admin_key:str|None=Header(None)):
    guard(x_admin_key)
    with engine.begin() as conn:
        p=conn.execute(select(people.c.id,people.c.nombre,people.c.cedula,people.c.asistio_en,
              people.c.email_estado,people.c.email_intentos).where(and_(people.c.id==pid,
              people.c.evento_id==EVENT_ID)).with_for_update()).first()
        if p is None:raise HTTPException(404,'Participante no encontrado')
        if p.asistio_en or p.email_estado!='pendiente' or p.email_intentos:
            raise HTTPException(409,'No se puede eliminar: ya tiene ingreso o historial de envío. Contacta al administrador para una corrección auditada.')
        conn.execute(delete(people).where(and_(people.c.id==pid,people.c.evento_id==EVENT_ID)))
    return {'eliminado':True,'nombre':p.nombre,'cedula':p.cedula}

class DeleteTests(BaseModel):
    confirmacion:str=Field(min_length=1)

@app.post('/api/eliminar-ejemplos')
def remove_tests(body:DeleteTests,x_admin_key:str|None=Header(None)):
    guard(x_admin_key)
    if body.confirmacion != 'ELIMINAR EJEMPLOS':
        raise HTTPException(400,'Debes confirmar exactamente ELIMINAR EJEMPLOS')
    removed=[]
    with engine.begin() as conn:
        rows=conn.execute(select(people.c.id,people.c.cedula,people.c.email,
            people.c.asistio_en,people.c.email_estado,people.c.email_intentos).where(
            people.c.evento_id==EVENT_ID).with_for_update()).all()
        for p in rows:
            # Solo registros identificados explícitamente con TEST, sin ingresos ni envíos.
            if not p.cedula.upper().startswith('TEST'):continue
            if p.asistio_en or p.email_estado!='pendiente' or p.email_intentos:continue
            conn.execute(delete(people).where(and_(people.c.id==p.id,people.c.evento_id==EVENT_ID)))
            removed.append(p.cedula)
    return {'eliminados':len(removed),'documentos':removed}

@app.get('/api/participantes')
def list_people(x_admin_key:str|None=Header(None)):
    guard(x_admin_key)
    with engine.connect() as c:
        count=c.execute(select(func.count()).select_from(people).where(people.c.evento_id==EVENT_ID)).scalar_one()
        arrived=c.execute(select(func.count()).select_from(people).where(and_(people.c.evento_id==EVENT_ID,people.c.asistio_en.is_not(None)))).scalar_one()
        rows=c.execute(select(people.c.id,people.c.nombre,people.c.cedula,people.c.email,people.c.telefono,people.c.tipo,people.c.asistio_en,people.c.correo_enviado_en,people.c.email_estado,people.c.email_error,people.c.email_intentos).where(people.c.evento_id==EVENT_ID).order_by(people.c.id.desc()).limit(3000)).all()
    with engine.connect() as c:
        states=dict(c.execute(select(people.c.email_estado,func.count()).where(people.c.evento_id==EVENT_ID)
                      .group_by(people.c.email_estado)).all())
    return {'total':count,'asistieron':arrived,'pendientes':count-arrived,
            'envios':states,'limite_diario':MAX_DAILY,'brevo_configurado':mail_configured(),
            'participantes':[asrow(p) for p in rows]}

@app.get('/api/qr/{pid}')
def download_qr(pid:int,x_admin_key:str|None=Header(None)):
    guard(x_admin_key);p=get_person(pid)
    return StreamingResponse(qr_bytes(p['token']),media_type='image/png',headers={'Cache-Control':'no-store','Content-Disposition':f'attachment; filename="entrada_{pid}.png"'})

@app.post('/api/enviar-correo/{pid}')
def email_one(pid:int, background_tasks:BackgroundTasks, x_admin_key:str|None=Header(None)):
    guard(x_admin_key)
    if not mail_configured(): raise HTTPException(409,'Configura BREVO_API_KEY y BREVO_SENDER_EMAIL en Render.')
    get_person(pid)
    ok, reason=claim_mail(pid)
    if not ok: raise HTTPException(409,reason)
    background_tasks.add_task(send_brevo,pid)
    return {'mensaje':'Envío programado. Consulta el estado para confirmar la aceptación de Brevo.', 'id':pid}

class BulkRequest(BaseModel):
    limite:int=Field(default=20,ge=1,le=20)

@app.post('/api/enviar-pendientes')
def email_pending(req:BulkRequest, background_tasks:BackgroundTasks, x_admin_key:str|None=Header(None)):
    guard(x_admin_key)
    if not mail_configured(): raise HTTPException(409,'Configura BREVO_API_KEY y BREVO_SENDER_EMAIL en Render.')
    with engine.connect() as c:
        ids=c.execute(select(people.c.id).where(and_(people.c.evento_id==EVENT_ID,
                people.c.email!='',people.c.email_estado=='pendiente'))
                .order_by(people.c.id).limit(3000)).scalars().all()
    claimed=[]
    for pid in ids:
        if len(claimed)>=req.limite: break
        if is_test_address(get_person(pid)['email']): continue
        ok, reason=claim_mail(pid)
        if not ok: break
        claimed.append(pid)
    if claimed: background_tasks.add_task(process_batch,claimed)
    return {'programados':len(claimed),'mensaje':'En proceso. Actualiza las estadísticas luego de unos segundos.',
            'limite_diario':MAX_DAILY,'nota':'Programado no significa entregado. Comprueba estado aceptado/entregado.'}

@app.post('/api/reintentar-correo/{pid}')
def retry_mail(pid:int, background_tasks:BackgroundTasks, x_admin_key:str|None=Header(None)):
    guard(x_admin_key)
    if not mail_configured(): raise HTTPException(409,'Brevo no configurado.')
    get_person(pid)
    ok,reason=claim_mail(pid,allow_retry=True)
    if not ok: raise HTTPException(409,reason+' No reintentar estados inciertos sin revisar Brevo.')
    background_tasks.add_task(send_brevo,pid)
    return {'mensaje':'Reintento programado. Actualiza el panel para revisar el resultado.'}

@app.get('/api/envios')
def logs_mail(x_admin_key:str|None=Header(None)):
    guard(x_admin_key)
    with engine.connect() as c:
        rows=c.execute(select(attempts.c.fecha_utc,attempts.c.destinatario,attempts.c.estado,
                     attempts.c.message_id,attempts.c.detalle).where(attempts.c.evento_id==EVENT_ID)
                     .order_by(attempts.c.id.desc()).limit(100)).all()
    return {'envios':[asrow(x) for x in rows]}

@app.post('/api/brevo/webhook/{secret}')
async def brevo_webhook(secret:str, request:Request):
    expected=os.getenv('BREVO_WEBHOOK_SECRET','')
    if not expected or not secrets.compare_digest(secret,expected):
        raise HTTPException(403,'Webhook no autorizado')
    content=await request.body()
    if len(content)>64000: raise HTTPException(413,'Payload demasiado grande')
    try: data=json.loads(content)
    except (ValueError,UnicodeError): raise HTTPException(400,'JSON inválido')
    if not isinstance(data,dict): raise HTTPException(400,'Formato inválido')
    msg_id=data.get('message-id') or data.get('messageId')
    evt=str(data.get('event','')).lower()
    new={'delivered':'entregado','hard_bounce':'rebotado','soft_bounce':'error',
         'invalid_email':'rebotado','blocked':'error','error':'error',
         'spam':'error','deferred':'error'}.get(evt)
    if not msg_id or not new: return {'procesado':False}
    with engine.begin() as c:
        p=c.execute(select(people.c.id,people.c.email_estado).where(and_(
            people.c.evento_id==EVENT_ID,people.c.email_message_id==msg_id))).first()
        if not p: return {'procesado':False}
        if p.email_estado=='entregado' and new!='rebotado':return {'procesado':True}
        c.execute(update(people).where(people.c.id==p.id).values(email_estado=new,
                 email_error=('Brevo reportó '+evt) if new in ('rebotado','error') else None))
        c.execute(update(attempts).where(attempts.c.message_id==msg_id).values(estado=new))
    return {'procesado':True}

@app.get('/api/whatsapp/{pid}')
def whatsapp_manual(pid:int,x_admin_key:str|None=Header(None)):
    guard(x_admin_key);p=get_person(pid)
    phone=re.sub('[^0-9]','',p['telefono'])
    if not phone: raise HTTPException(400,'No hay número de celular para este participante.')
    return {'url':f'https://wa.me/{phone}?text='+quote(f'Hola {p["nombre"]}, contactamos desde la organización del evento UPS sobre tu credencial QR. Por seguridad, te la enviaremos por el canal de entrega autorizado.'),'nota':'Abre un chat manual. No envía automáticamente el QR. Requiere consentimiento y número con código de país.'}

@app.get('/api/exportar-asistencia')
def export_attendance(x_admin_key:str|None=Header(None)):
    guard(x_admin_key)
    with engine.connect() as c:
        records=c.execute(select(people.c.cedula,people.c.nombre,people.c.email,people.c.telefono,people.c.asistio_en).where(people.c.evento_id==EVENT_ID)).all()
    out=io.StringIO();w=csv.writer(out);w.writerow(['Identificacion','Nombre','Correo','Celular','Ingreso UTC'])
    for r in records: w.writerow(r)
    return Response('\ufeff'+out.getvalue(),media_type='text/csv; charset=utf-8',headers={'Content-Disposition':'attachment; filename=asistencia_ups.csv','Cache-Control':'no-store'})

class Scan(BaseModel):qr:str=Field(min_length=15,max_length=500)
@app.post('/api/verificar')
def verify(scan:Scan,x_admin_key:str|None=Header(None)):
    guard(x_admin_key)
    parts=scan.qr.strip().split(':')
    if len(parts)!=3 or parts[0]!='EVT1' or parts[1]!=EVENT_ID: raise HTTPException(400,'QR ajeno al evento o formato no reconocido')
    token=parts[2]
    if not 30<=len(token)<=100: raise HTTPException(400,'QR inválido')
    digest=hashlib.sha256(token.encode()).hexdigest()
    # UPDATE atómico: dos escáneres simultáneos no pueden registrar dos primeras entradas.
    with engine.begin() as c:
        ts=now()
        res=c.execute(update(people).where(and_(people.c.evento_id==EVENT_ID,people.c.token_hash==digest,people.c.asistio_en.is_(None))).values(asistio_en=ts))
        p=c.execute(select(people.c.nombre,people.c.cedula,people.c.asistio_en).where(and_(people.c.evento_id==EVENT_ID,people.c.token_hash==digest))).first()
        if not p: raise HTTPException(404,'QR no registrado en la lista oficial importada')
        valid=res.rowcount==1
    return {'estado':'valido' if valid else 'duplicado','nombre':p.nombre,'cedula':p.cedula,'asistio_en':p.asistio_en,'mensaje':'Asistencia registrada' if valid else 'Esta persona ya ingresó previamente'}

@app.get('/healthz')
def healthz():
    with engine.connect() as c: c.execute(text('SELECT 1'))
    return {'ok': True}

class Lookup(BaseModel):
    cedula:str=Field(min_length=1,max_length=50)

@app.post('/api/consultar')
def lookup_person(body:Lookup,x_admin_key:str|None=Header(None)):
    guard(x_admin_key)
    with engine.connect() as c:
        p=c.execute(select(people.c.nombre,people.c.cedula,people.c.asistio_en).where(
            and_(people.c.evento_id==EVENT_ID,people.c.cedula==body.cedula))).first()
    if not p:raise HTTPException(404,'No consta en el listado importado')
    return asrow(p)
