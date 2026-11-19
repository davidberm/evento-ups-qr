
# importaciones_ups.py
# Historial e importación incremental para Eventos UPS.

import csv
import hashlib
import io
import re
import secrets
from collections import Counter, defaultdict
from datetime import datetime, timezone
from zoneinfo import ZoneInfo

from fastapi import File, Header, HTTPException, UploadFile
from fastapi.responses import StreamingResponse
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from pydantic import BaseModel
from sqlalchemy import (
    MetaData, Table, Column, Integer, String, Text,
    ForeignKey, select, insert, update, delete, and_, text
)

ECUADOR = ZoneInfo("America/Guayaquil")
MAX_BYTES = 8_000_000


def ahora():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def normalizar_titulo(valor):
    import unicodedata
    valor = unicodedata.normalize("NFKD", str(valor or ""))
    valor = valor.encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]", "", valor)


def normalizar_documento(valor):
    return re.sub(r"[\s.\-]", "", str(valor or "")).upper()


def campo(fila, *nombres):
    for nombre in nombres:
        valor = fila.get(normalizar_titulo(nombre), "")
        if str(valor).strip():
            return str(valor).strip()
    return ""


def leer_archivo(contenido, nombre):
    try:
        if nombre.lower().endswith(".xlsx"):
            libro = load_workbook(
                io.BytesIO(contenido),
                read_only=True,
                data_only=True
            )
            try:
                hoja = libro.active
                filas = hoja.iter_rows(values_only=True)
                cabecera = next(filas, None)
                if not cabecera:
                    raise ValueError("Excel sin encabezados")

                titulos = [normalizar_titulo(x) for x in cabecera]
                resultado = [
                    dict(zip(
                        titulos,
                        [str(x).strip() if x is not None else ""
                         for x in valores]
                    ))
                    for valores in filas
                ]
            finally:
                libro.close()

        elif nombre.lower().endswith(".csv"):
            try:
                contenido_texto = contenido.decode("utf-8-sig")
            except UnicodeDecodeError:
                contenido_texto = contenido.decode("cp1252")

            try:
                dialecto = csv.Sniffer().sniff(
                    contenido_texto[:4096],
                    delimiters=",;\t"
                )
            except csv.Error:
                dialecto = csv.excel

            lector = csv.DictReader(
                io.StringIO(contenido_texto),
                dialect=dialecto
            )
            if not lector.fieldnames:
                raise ValueError("CSV sin encabezados")

            resultado = [
                {
                    normalizar_titulo(k): str(v or "").strip()
                    for k, v in fila.items()
                    if k
                }
                for fila in lector
            ]
        else:
            raise ValueError("Solo se admiten Excel y CSV")

    except (ValueError, OSError, TypeError) as error:
        raise HTTPException(400, str(error)) from error

    if len(resultado) > 10000:
        raise HTTPException(400, "Máximo 10000 filas por archivo")

    return resultado


def preparar(filas):
    resultado = []
    documentos = set()

    for numero, fila in enumerate(filas, start=2):
        documento = normalizar_documento(campo(
            fila, "cedula", "identificacion",
            "documento", "numero de identificacion",
            "numero de documento"
        ))

        nombre = campo(
            fila, "nombre completo", "participante",
            "nombres", "nombre"
        )
        apellido = campo(fila, "apellidos", "apellido")
        if apellido and normalizar_titulo(apellido) not in normalizar_titulo(nombre):
            nombre = f"{nombre} {apellido}".strip()

        correo = campo(
            fila, "correo personal", "correo electronico",
            "correo", "email", "e-mail"
        ).lower()

        participante = {
            "fila": numero,
            "cedula": documento,
            "nombre": nombre,
            "email": correo,
            "telefono": campo(
                fila, "telefono", "celular",
                "telefono movil", "whatsapp"
            ),
            "tipo": campo(
                fila, "tipo inscripcion",
                "tipo de inscripcion"
            ),
        }

        invalido = (
            not re.fullmatch(r"[A-Z0-9]{5,50}", documento)
            or (documento.isdigit() and len(documento) != 10)
            or not 2 <= len(nombre) <= 200
            or (
                correo
                and not re.fullmatch(
                    r"[^\s@]+@[^\s@]+\.[^\s@]+", correo
                )
            )
        )

        if invalido:
            participante["estado"] = "invalido"
        elif documento in documentos:
            participante["estado"] = "repetido"
        else:
            participante["estado"] = "pendiente"
            documentos.add(documento)

        resultado.append(participante)

    return resultado


class Confirmacion(BaseModel):
    confirmacion: str


def instalar_importaciones(
    app, engine, people, attempts, event_id, guard
):
    metadata = MetaData()

    lotes = Table(
        "importaciones_ups", metadata,
        Column("id", Integer, primary_key=True),
        Column("evento_id", String(80), nullable=False),
        Column("archivo", String(255), nullable=False),
        Column("sha256", String(64), nullable=False),
        Column("creada_en", String(40), nullable=False),
        Column("estado", String(30), nullable=False),
        Column("total", Integer, nullable=False),
        Column("nuevos", Integer, nullable=False),
        Column("existentes", Integer, nullable=False),
        Column("repetidos", Integer, nullable=False),
        Column("conflictos", Integer, nullable=False),
        Column("invalidos", Integer, nullable=False),
        Column("eliminados", Integer, nullable=False, default=0),
    )

    detalles = Table(
        "importaciones_ups_detalles", metadata,
        Column("id", Integer, primary_key=True),
        Column("lote_id", Integer, ForeignKey("importaciones_ups.id")),
        Column("participante_id", Integer),
        Column("fila", Integer, nullable=False),
        Column("documento", String(50)),
        Column("accion", String(40), nullable=False),
    )

    metadata.create_all(engine, tables=[lotes, detalles])

    async def recibir(archivo):
        nombre = (archivo.filename or "").replace("\\", "/").split("/")[-1]
        if not nombre.lower().endswith((".xlsx", ".csv")):
            raise HTTPException(400, "Selecciona Excel o CSV")

        contenido = await archivo.read(MAX_BYTES + 1)
        if len(contenido) > MAX_BYTES:
            raise HTTPException(413, "Archivo superior a 8 MB")

        sha = hashlib.sha256(contenido).hexdigest()
        filas = preparar(leer_archivo(contenido, nombre))
        return nombre[:255], sha, filas

    def clasificar(conexion, filas):
        originales = conexion.execute(
            select(
                people.c.id, people.c.cedula,
                people.c.email
            ).where(people.c.evento_id == event_id)
        ).all()

        por_documento = {}
        por_correo = defaultdict(list)

        for p in originales:
            por_documento.setdefault(
                normalizar_documento(p.cedula), p
            )
            if p.email:
                por_correo[p.email.strip().lower()].append(p)

        salida = []

        for fila in filas:
            p = dict(fila)

            if p["estado"] == "pendiente":
                existente = por_documento.get(p["cedula"])

                if existente:
                    p["estado"] = "existente"
                    p["participante_id"] = existente.id

                elif p["email"] and p["email"] in por_correo:
                    p["estado"] = "conflicto"

                else:
                    p["estado"] = "nuevo"

            salida.append(p)

        return salida

    def resumen(filas):
        estados = Counter(x["estado"] for x in filas)
        return {
            "filas": len(filas),
            "nuevos": estados["nuevo"],
            "existentes": estados["existente"],
            "repetidos": estados["repetido"],
            "conflictos": estados["conflicto"],
            "invalidos": estados["invalido"],
        }

    @app.post("/api/importaciones/analizar")
    async def analizar(
        archivo: UploadFile = File(...),
        x_admin_key: str | None = Header(None)
    ):
        guard(x_admin_key)
        nombre, sha, filas = await recibir(archivo)

        with engine.connect() as conexion:
            previo = conexion.execute(
                select(lotes.c.id).where(and_(
                    lotes.c.evento_id == event_id,
                    lotes.c.sha256 == sha,
                    lotes.c.estado == "completada"
                )).limit(1)
            ).scalar_one_or_none()

            clasificados = clasificar(conexion, filas)

        return {
            "archivo": nombre,
            "sha256": sha,
            "ya_importado": previo is not None,
            **resumen(clasificados),
        }

    @app.post("/api/importaciones/confirmar")
    async def confirmar(
        archivo: UploadFile = File(...),
        sha256: str = "",
        x_admin_key: str | None = Header(None)
    ):
        guard(x_admin_key)
        nombre, sha, filas = await recibir(archivo)

        if not sha256 or not secrets.compare_digest(sha256, sha):
            raise HTTPException(
                409, "El archivo cambió. Analízalo nuevamente."
            )

        with engine.begin() as conexion:
            if engine.dialect.name == "postgresql":
                # Bloquea importaciones simultáneas de este evento.
                conexion.execute(
                    text("SELECT pg_advisory_xact_lock(hashtext(:evento))"),
                    {"evento": event_id}
                )

            previo = conexion.execute(
                select(lotes.c.id).where(and_(
                    lotes.c.evento_id == event_id,
                    lotes.c.sha256 == sha,
                    lotes.c.estado == "completada"
                ))
            ).scalar_one_or_none()

            if previo:
                raise HTTPException(
                    409,
                    f"Archivo ya importado en la carga #{previo}"
                )

            clasificados = clasificar(conexion, filas)
            cifras = resumen(clasificados)

            creado = conexion.execute(
                insert(lotes).values(
                    evento_id=event_id,
                    archivo=nombre,
                    sha256=sha,
                    creada_en=ahora(),
                    estado="completada",
                    total=len(filas),
                    nuevos=0,
                    existentes=cifras["existentes"],
                    repetidos=cifras["repetidos"],
                    conflictos=cifras["conflictos"],
                    invalidos=cifras["invalidos"],
                    eliminados=0
                )
            )
            lote_id = creado.inserted_primary_key[0]
            nuevos = 0

            for p in clasificados:
                accion = p["estado"]
                participante_id = p.get("participante_id")

                if accion == "nuevo":
                    token = secrets.token_urlsafe(32)

                    nuevo = conexion.execute(
                        insert(people).values(
                            evento_id=event_id,
                            cedula=p["cedula"],
                            nombre=p["nombre"],
                            email=p["email"],
                            telefono=p["telefono"],
                            tipo=p["tipo"],
                            token=token,
                            token_hash=hashlib.sha256(
                                token.encode()
                            ).hexdigest(),
                            registrado_en=ahora()
                        )
                    )
                    participante_id = nuevo.inserted_primary_key[0]
                    nuevos += 1
                    accion = "creado"

                conexion.execute(
                    insert(detalles).values(
                        lote_id=lote_id,
                        participante_id=participante_id,
                        fila=p["fila"],
                        documento=p["cedula"],
                        accion=accion
                    )
                )

            conexion.execute(
                update(lotes).where(lotes.c.id == lote_id)
                .values(nuevos=nuevos)
            )

        return {
            "importacion_id": lote_id,
            "nuevos": nuevos,
            "existentes": cifras["existentes"],
            "repetidos": cifras["repetidos"],
            "conflictos": cifras["conflictos"],
            "invalidos": cifras["invalidos"],
        }

    @app.get("/api/importaciones")
    def historial(x_admin_key: str | None = Header(None)):
        guard(x_admin_key)

        with engine.connect() as conexion:
            registros = conexion.execute(
                select(lotes)
                .where(lotes.c.evento_id == event_id)
                .order_by(lotes.c.id.desc())
                .limit(100)
            ).all()

        return {
            "importaciones": [
                dict(x._mapping) for x in registros
            ]
        }

    @app.get("/api/importaciones/anteriores")
    def anteriores(x_admin_key: str | None = Header(None)):
        guard(x_admin_key)

        with engine.connect() as conexion:
            ya_asociados = select(
                detalles.c.participante_id
            ).where(and_(
                detalles.c.accion == "creado",
                detalles.c.participante_id.is_not(None)
            ))

            registros = conexion.execute(
                select(
                    people.c.id,
                    people.c.nombre,
                    people.c.cedula,
                    people.c.registrado_en,
                    people.c.email_estado,
                    people.c.asistio_en
                )
                .where(and_(
                    people.c.evento_id == event_id,
                    ~people.c.id.in_(ya_asociados)
                ))
                .order_by(people.c.registrado_en.desc())
            ).all()

        grupos = defaultdict(list)

        for p in registros:
            try:
                fecha = datetime.fromisoformat(
                    p.registrado_en.replace("Z", "+00:00")
                )
                hora = fecha.astimezone(ECUADOR).strftime(
                    "%Y-%m-%d %H:%M"
                )
            except (ValueError, AttributeError):
                hora = "Fecha desconocida"

            grupos[hora].append({
                "id": p.id,
                "nombre": p.nombre,
                "cedula": p.cedula,
                "email_estado": p.email_estado,
                "asistio": bool(p.asistio_en),
            })

        return {
            "total": len(registros),
            "grupos": [
                {
                    "hora": hora,
                    "cantidad": len(personas),
                    "participantes": personas
                }
                for hora, personas in sorted(
                    grupos.items(), reverse=True
                )
            ]
        }

    def revisar_deshacer(conexion, lote_id):
        creados = conexion.execute(
            select(detalles.c.participante_id)
            .where(and_(
                detalles.c.lote_id == lote_id,
                detalles.c.accion == "creado"
            ))
        ).scalars().all()

        eliminables = []
        protegidos = []

        for pid in creados:
            participante = conexion.execute(
                select(people).where(and_(
                    people.c.id == pid,
                    people.c.evento_id == event_id
                ))
            ).first()

            if not participante:
                continue

            envios = conexion.execute(
                select(attempts.c.id).where(
                    attempts.c.participante_id == pid
                ).limit(1)
            ).first()

            otras_cargas = conexion.execute(
                select(detalles.c.id).where(and_(
                    detalles.c.participante_id == pid,
                    detalles.c.lote_id != lote_id
                )).limit(1)
            ).first()

            actividad = (
                envios is not None
                or otras_cargas is not None
                or participante.asistio_en is not None
                or participante.correo_enviado_en is not None
                or participante.email_estado != "pendiente"
                or int(participante.email_intentos or 0) > 0
            )

            if actividad:
                protegidos.append(pid)
            else:
                eliminables.append(pid)

        return eliminables, protegidos

    @app.get("/api/importaciones/{lote_id}")
    def detalle(
        lote_id: int,
        x_admin_key: str | None = Header(None)
    ):
        guard(x_admin_key)

        with engine.connect() as conexion:
            lote = conexion.execute(
                select(lotes).where(and_(
                    lotes.c.id == lote_id,
                    lotes.c.evento_id == event_id
                ))
            ).first()

            if not lote:
                raise HTTPException(404, "Carga no encontrada")

            filas = conexion.execute(
                select(detalles)
                .where(detalles.c.lote_id == lote_id)
                .order_by(detalles.c.fila)
            ).all()

        return {
            "importacion": dict(lote._mapping),
            "detalles": [dict(x._mapping) for x in filas]
        }

    @app.get("/api/importaciones/{lote_id}/simular")
    def simular(
        lote_id: int,
        x_admin_key: str | None = Header(None)
    ):
        guard(x_admin_key)

        with engine.connect() as conexion:
            lote = conexion.execute(
                select(lotes).where(and_(
                    lotes.c.id == lote_id,
                    lotes.c.evento_id == event_id
                ))
            ).first()

            if not lote:
                raise HTTPException(404, "Carga no encontrada")

            eliminables, protegidos = revisar_deshacer(
                conexion, lote_id
            )

        return {
            "eliminables": len(eliminables),
            "protegidos": len(protegidos),
            "estado": lote.estado
        }

    @app.post("/api/importaciones/{lote_id}/deshacer")
    def deshacer(
        lote_id: int,
        cuerpo: Confirmacion,
        x_admin_key: str | None = Header(None)
    ):
        guard(x_admin_key)

        if cuerpo.confirmacion != f"DESHACER {lote_id}":
            raise HTTPException(
                400, f"Debes escribir DESHACER {lote_id}"
            )

        with engine.begin() as conexion:
            lote = conexion.execute(
                select(lotes).where(and_(
                    lotes.c.id == lote_id,
                    lotes.c.evento_id == event_id
                )).with_for_update()
            ).first()

            if not lote:
                raise HTTPException(404, "Carga no encontrada")

            if lote.estado != "completada":
                raise HTTPException(409, "Carga ya deshecha")

            eliminables, protegidos = revisar_deshacer(
                conexion, lote_id
            )

            eliminados = 0

            for pid in eliminables:
                p = conexion.execute(
                    select(people).where(and_(
                        people.c.id == pid,
                        people.c.evento_id == event_id
                    )).with_for_update()
                ).first()

                if not p:
                    continue

                # Segunda comprobación con la fila bloqueada.
                correos = conexion.execute(
                    select(attempts.c.id).where(
                        attempts.c.participante_id == pid
                    ).limit(1)
                ).first()

                otras = conexion.execute(
                    select(detalles.c.id).where(and_(
                        detalles.c.participante_id == pid,
                        detalles.c.lote_id != lote_id
                    )).limit(1)
                ).first()

                if (
                    correos or otras or p.asistio_en
                    or p.correo_enviado_en
                    or p.email_estado != "pendiente"
                    or int(p.email_intentos or 0) > 0
                ):
                    protegidos.append(pid)
                    continue

                conexion.execute(
                    delete(people).where(and_(
                        people.c.id == pid,
                        people.c.evento_id == event_id
                    ))
                )
                eliminados += 1

            conexion.execute(
                update(lotes)
                .where(lotes.c.id == lote_id)
                .values(
                    estado="deshecha_parcial"
                    if protegidos else "deshecha",
                    eliminados=eliminados
                )
            )

        return {
            "eliminados": eliminados,
            "protegidos": len(set(protegidos))
        }

    @app.get("/api/importaciones/{lote_id}/nuevos")
    def descargar_nuevos(
        lote_id: int,
        formato: str = "xlsx",
        x_admin_key: str | None = Header(None)
    ):
        guard(x_admin_key)

        if formato not in ("xlsx", "csv"):
            raise HTTPException(400, "Formato inválido")

        with engine.connect() as conexion:
            lote = conexion.execute(
                select(lotes.c.id).where(and_(
                    lotes.c.id == lote_id,
                    lotes.c.evento_id == event_id
                ))
            ).first()

            if not lote:
                raise HTTPException(404, "Carga no encontrada")

            registros = conexion.execute(
                select(
                    people.c.cedula,
                    people.c.nombre,
                    people.c.email,
                    people.c.telefono,
                    people.c.tipo,
                    people.c.email_estado
                ).select_from(
                    detalles.join(
                        people,
                        detalles.c.participante_id == people.c.id
                    )
                ).where(and_(
                    detalles.c.lote_id == lote_id,
                    detalles.c.accion == "creado",
                    people.c.evento_id == event_id
                ))
            ).all()

        encabezados = [
            "Cédula", "Nombre", "Correo",
            "Teléfono", "Tipo", "Estado correo"
        ]

        def seguro(v):
            valor = str(v or "")
            if valor.lstrip().startswith(("=", "+", "-", "@")):
                return "'" + valor
            return valor

        datos = [[seguro(v) for v in fila] for fila in registros]

        if formato == "csv":
            salida = io.StringIO()
            escritor = csv.writer(salida, delimiter=";")
            escritor.writerow(encabezados)
            escritor.writerows(datos)

            contenido = io.BytesIO(
                salida.getvalue().encode("utf-8-sig")
            )
            medio = "text/csv; charset=utf-8"

        else:
            libro = Workbook()
            hoja = libro.active
            hoja.title = "Solo nuevos"
            hoja.append(encabezados)

            for fila in datos:
                hoja.append(fila)

            hoja.freeze_panes = "A2"
            hoja.auto_filter.ref = hoja.dimensions

            for celda in hoja[1]:
                celda.font = Font(bold=True, color="FFFFFF")
                celda.fill = PatternFill(
                    "solid", fgColor="176B68"
                )

            for columna, ancho in {
                "A": 22, "B": 42, "C": 40,
                "D": 20, "E": 25, "F": 23
            }.items():
                hoja.column_dimensions[columna].width = ancho

            contenido = io.BytesIO()
            libro.save(contenido)
            contenido.seek(0)
            medio = (
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )

        nombre = f"ups_carga_{lote_id}_nuevos.{formato}"

        return StreamingResponse(
            contenido,
            media_type=medio,
            headers={
                "Content-Disposition":
                    f'attachment; filename="{nombre}"',
                "Cache-Control": "no-store"
            }
        )
