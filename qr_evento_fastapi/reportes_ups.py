"""Reportes XLSX y consulta diaria para el sistema Eventos UPS.
Integración: instalar_reportes(app, engine, people, attempts, EVENT_ID, guard)
"""
from datetime import date, datetime, timezone
from io import BytesIO
from zoneinfo import ZoneInfo
from collections import Counter
from fastapi import Header, HTTPException
from fastapi.responses import StreamingResponse
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from sqlalchemy import select

ECUADOR = ZoneInfo('America/Guayaquil')


def _fecha_local(iso):
    if not iso:
        return None
    try:
        dt = datetime.fromisoformat(str(iso).replace('Z', '+00:00'))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(ECUADOR)
    except (ValueError, TypeError):
        return None


def _dia(iso):
    dt = _fecha_local(iso)
    return dt.date().isoformat() if dt else ''


def _fecha_hora(iso):
    dt = _fecha_local(iso)
    return dt.strftime('%Y-%m-%d %H:%M:%S') if dt else ''


def _limpiar_excel(valor):
    if valor is None:
        return ''
    s = str(valor).replace('\x00', '')
    # Evita que datos proporcionados por terceros se interpreten como fórmulas.
    if s.lstrip().startswith(('=', '+', '-', '@')):
        s = "'" + s
    return s[:32000]


def _hoja(wb, titulo, columnas, filas):
    ws = wb.create_sheet(titulo)
    ws.append(columnas)
    for fila in filas:
        ws.append([_limpiar_excel(c) for c in fila])
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 28
    for cell in ws[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='176B68')
        cell.alignment = Alignment(vertical='center', wrap_text=True)
    for i, col in enumerate(ws.columns, start=1):
        # Limitar ancho para mantener Excel manejable, pero se puede ampliar.
        length = max((len(str(c.value or '')) for c in list(col)[:150]), default=12)
        ws.column_dimensions[get_column_letter(i)].width = min(max(length + 3, 14), 43)
    return ws


def instalar_reportes(app, engine, people, attempts, event_id, guard):
    def recopilar(fecha, modo):
        try:
            date.fromisoformat(fecha)
        except (ValueError, TypeError):
            raise HTTPException(400, 'La fecha debe tener formato AAAA-MM-DD.')
        if modo not in ('evento', 'actividad'):
            raise HTTPException(400, 'Modo válido: evento o actividad.')
        with engine.connect() as conn:
            personas = [dict(x._mapping) for x in conn.execute(
                select(people).where(people.c.evento_id == event_id)
                .order_by(people.c.id)
            ).all()]
            envios = [dict(x._mapping) for x in conn.execute(
                select(attempts).where(attempts.c.evento_id == event_id)
                .order_by(attempts.c.id)
            ).all()]
        envios_por_persona = {}
        for intento in envios:
            envios_por_persona.setdefault(intento['participante_id'], []).append(intento)
        if modo == 'actividad':
            personas = [p for p in personas if (
                _dia(p.get('registrado_en')) == fecha or
                _dia(p.get('asistio_en')) == fecha or
                any(_dia(e.get('fecha_utc')) == fecha for e in envios_por_persona.get(p['id'], []))
            )]
            envios = [e for e in envios if _dia(e.get('fecha_utc')) == fecha]
        # En modo evento: muestra TODOS los inscritos del evento, aun si faltaron.
        # 'Ingresó' refiere estrictamente al día seleccionado.
        return personas, envios, envios_por_persona

    def construir_filas(personas, fecha, envios_por_persona):
        filas = []
        for p in personas:
            rapido = p.get('tipo') == 'Registro rápido'
            ingreso = _dia(p.get('asistio_en')) == fecha
            intentos = envios_por_persona.get(p['id'], [])
            ultimo = intentos[-1] if intentos else {}
            filas.append([
                p.get('cedula'), p.get('nombre'), p.get('email'), p.get('telefono'),
                'Registro rápido' if rapido else 'Importación Excel/CSV',
                p.get('tipo') or '',
                _fecha_hora(p.get('registrado_en')),
                'Ingresó' if ingreso else 'No ingresó',
                _fecha_hora(p.get('asistio_en')),
                p.get('email_estado') or 'pendiente',
                p.get('email_intentos') or 0,
                _fecha_hora(p.get('correo_enviado_en')),
                _fecha_hora(ultimo.get('fecha_utc')),
                ultimo.get('estado') or '',
                ultimo.get('message_id') or '',
                p.get('email_error') or ultimo.get('detalle') or ''
            ])
        return filas

    encabezados = [
        'Identificación', 'Nombre completo', 'Correo', 'Celular',
        'Origen de inscripción', 'Tipo de inscripción',
        'Fecha de registro (Ecuador)', 'Asistencia día seleccionado',
        'Fecha/hora de ingreso (Ecuador)', 'Estado actual del correo',
        'Intentos de correo', 'Fecha de correo aceptado (Ecuador)',
        'Último intento (Ecuador)', 'Estado último intento',
        'ID del mensaje Brevo', 'Observaciones del correo'
    ]

    @app.get('/api/reporte-dia')
    def reporte_dia(fecha: str, modo: str = 'evento', x_admin_key: str | None = Header(None)):
        guard(x_admin_key)
        personas, envios, agrupados = recopilar(fecha, modo)
        ingresaron = sum(_dia(p.get('asistio_en')) == fecha for p in personas)
        rapido = sum(p.get('tipo') == 'Registro rápido' for p in personas)
        estados = Counter(str(e.get('estado') or 'desconocido') for e in envios)
        return {
            'fecha': fecha, 'modo': modo, 'personas': len(personas),
            'ingresaron_ese_dia': ingresaron,
            'no_ingresaron_ese_dia': len(personas)-ingresaron,
            'registro_rapido': rapido, 'importacion': len(personas)-rapido,
            'intentos_correo': len(envios), 'estados_intentos_correo': dict(estados)
        }

    @app.get('/api/exportar-excel')
    def exportar_excel(fecha: str, modo: str = 'evento', x_admin_key: str | None = Header(None)):
        guard(x_admin_key)
        personas, envios, agrupados = recopilar(fecha, modo)
        filas = construir_filas(personas, fecha, agrupados)
        wb = Workbook()
        wb.remove(wb.active)
        _hoja(wb, 'Todos', encabezados, filas)
        _hoja(wb, 'Ingresaron', encabezados, [r for r in filas if r[7] == 'Ingresó'])
        _hoja(wb, 'No ingresaron', encabezados, [r for r in filas if r[7] == 'No ingresó'])
        _hoja(wb, 'Registro rapido', encabezados, [r for r in filas if r[4] == 'Registro rápido'])
        _hoja(wb, 'Correos', ['ID participante', 'Correo', 'Fecha intento Ecuador', 'Estado', 'ID Brevo', 'Detalle'], [
            [e.get('participante_id'), e.get('destinatario'), _fecha_hora(e.get('fecha_utc')),
             e.get('estado'), e.get('message_id'), e.get('detalle')]
            for e in envios
        ])
        _hoja(wb, 'Resumen', ['Indicador', 'Valor'], [
            ['Evento', event_id], ['Fecha consultada', fecha], ['Modalidad', modo],
            ['Total participantes del reporte', len(filas)],
            ['Ingresaron ese día', sum(r[7] == 'Ingresó' for r in filas)],
            ['No ingresaron ese día', sum(r[7] == 'No ingresó' for r in filas)],
            ['Registro rápido', sum(r[4] == 'Registro rápido' for r in filas)],
            ['Importación Excel/CSV', sum(r[4] != 'Registro rápido' for r in filas)],
            ['Intentos de correo incluidos', len(envios)]
        ])
        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        nombre = f'evento_ups_{fecha}_{modo}.xlsx'
        return StreamingResponse(
            buffer,
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={'Content-Disposition': f'attachment; filename="{nombre}"', 'Cache-Control': 'no-store'}
        )
