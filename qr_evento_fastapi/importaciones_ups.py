
# Historial e importación incremental para Eventos UPS.

import xlrd
import csv
import hashlib
import io
import re
import secrets
import unicodedata

from collections import Counter, defaultdict
from datetime import datetime, timezone
from zoneinfo import ZoneInfo

from fastapi import File, Header, HTTPException, UploadFile
from fastapi.responses import StreamingResponse
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from pydantic import BaseModel
from sqlalchemy import (
    MetaData, Table, Column, Integer, String,
    ForeignKey, select, insert, update, delete,
    and_, text
)

ECUADOR = ZoneInfo("America/Guayaquil")
MAX_BYTES = 8_000_000


def ahora():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def normalizar_titulo(valor):
    valor = unicodedata.normalize("NFKD", str(valor or ""))
    valor = valor.encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]", "", valor)


def normalizar_documento(valor):
    return re.sub(r"[\s.\-]", "", str(valor or "")).upper()


def convertir_valor(valor):
    """
    Convierte valores de Excel a texto sin dejar identificaciones
    numéricas como 940077514.0.
    """
    if valor is None:
        return ""

    if isinstance(valor, bool):
        return str(valor)

    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))

    return str(valor).strip()


def campo(fila, *nombres):
    for nombre in nombres:
        valor = fila.get(normalizar_titulo(nombre), "")
        if str(valor).strip():
            return str(valor).strip()
    return ""


def leer_archivo(contenido, nombre):
    nombre = nombre.lower()

    try:
        if nombre.endswith(".xlsx"):
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

                resultado = []

                for valores in filas:
                    resultado.append({
                        titulo: convertir_valor(valor)
                        for titulo, valor in zip(titulos, valores)
                        if titulo
                    })

            finally:
                libro.close()

        elif nombre.endswith(".xls"):
            libro = xlrd.open_workbook(file_contents=contenido)
            hoja = libro.sheet_by_index(0)

            if hoja.nrows < 1:
                raise ValueError("Excel sin encabezados")

            titulos = [
                normalizar_titulo(hoja.cell_value(0, columna))
                for columna in range(hoja.ncols)
            ]

            resultado = []

            for numero_fila in range(1, hoja.nrows):
                fila = {}

                for columna, titulo in enumerate(titulos):
                    if not titulo:
                        continue

                    celda = hoja.cell(numero_fila, columna)
                    valor = celda.value

                    if celda.ctype == xlrd.XL_CELL_NUMBER:
                        if float(valor).is_integer():
                            valor = str(int(valor))
                        else:
                            valor = str(valor)

                    elif celda.ctype == xlrd.XL_CELL_DATE:
                        valor = xlrd.xldate_as_datetime(
                            valor, libro.datemode
                        ).isoformat()

                    else:
                        valor = str(valor or "").strip()

                    fila[titulo] = valor

                resultado.append(fila)

        elif nombre.endswith(".csv"):
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
            raise ValueError("Solo se admiten archivos XLS, XLSX y CSV")

    except HTTPException:
        raise
    except Exception as error:
        raise HTTPException(
            400,
            f"No se pudo leer el archivo: {error}"
        ) from error

    if len(resultado) > 10000:
        raise HTTPException(400, "Máximo 10000 filas por archivo")

    return resultado


def correo_valido(correo):
    if not correo:
        return True

    return bool(re.fullmatch(
        r"[^\s@]+@[^\s@]+\.[^\s@]+",
        correo
    ))


def preparar(filas):
    resultado = []
    documentos = set()

    for numero, fila in enumerate(filas, start=2):
        documento = normalizar_documento(campo(
            fila,
            "cedula",
            "identificacion",
            "documento",
            "numero de identificacion",
            "numero de documento"
        ))

        nombre = campo(
            fila,
            "nombre completo",
            "participante",
            "nombres",
            "nombre"
        )

        apellido = campo(fila, "apellidos", "apellido")

        if (
            apellido
            and normalizar_titulo(apellido)
            not in normalizar_titulo(nombre)
        ):
            nombre = f"{nombre} {apellido}".strip()

        correo = campo(
            fila,
            "correo personal",
            "correo electronico",
            "correo",
            "email",
            "e-mail",
            "email personal"
        ).lower()

        correo_institucional = campo(
            fila,
            "correo institucional",
            "email institucional",
            "correo ups",
            "email ups",
            "correo institucional ups",
            "email institucional ups"
        ).lower()

        participante = {
            "fila": numero,
            "cedula": documento,
            "nombre": nombre,
            "email": correo,
            "email_institucional": correo_institucional,
            "telefono": campo(
                fila,
                "telefono",
                "celular",
                "telefono movil",
                "whatsapp"
            ),
            "tipo": campo(
                fila,
                "tipo inscripcion",
                "tipo de inscripcion"
            ),
        }

        correos = [
            x for x in (correo, correo_institucional) if x
        ]

        documento_valido = bool(
            re.fullmatch(r"[A-Z0-9]{5,50}", documento)
        )

        invalido = (
            not documento_valido
            or not 2 <= len(nombre) <= 200
            or not correos
            or not all(correo_valido(x) for x in correos)
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
        "importaciones_ups",
        metadata,
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
        "importaciones_ups_detalles",
        metadata,
        Column("id", Integer, primary_key=True),
        Column(
            "lote_id",
            Integer,
            ForeignKey("importaciones_ups.id")
        ),
        Column("participante_id", Integer),
        Column("fila", Integer, nullable=False),
        Column("documento", String(50)),
        Column("accion", String(40), nullable=False),
    )

    metadata.create_all(engine, tables=[lotes, detalles])

    async def recibir(archivo):
        nombre = (
            (archivo.filename or "")
            .replace("\\", "/")
            .split("/")[-1]
        )

        if not nombre.lower().endswith((".xlsx", ".xls", ".csv")):
            raise HTTPException(
                400,
                "Selecciona un archivo XLS, XLSX o CSV"
            )

        contenido = await archivo.read(MAX_BYTES + 1)

        if len(contenido) > MAX_BYTES:
            raise HTTPException(413, "Archivo superior a 8 MB")

        sha = hashlib.sha256(contenido).hexdigest()
        filas = preparar(leer_archivo(contenido, nombre))

        return nombre[:255], sha, filas

    def clasificar(conexion, filas):
        originales = conexion.execute(
            select(
                people.c.id,
                people.c.cedula,
                people.c.email,
                people.c.email_institucional
            ).where(people.c.evento_id == event_id)
        ).all()

        por_documento = {}
        por_correo = defaultdict(set)

        for persona in originales:
            documento = normalizar_documento(persona.cedula)

            por_documento.setdefault(documento, persona)

            for correo in (
                persona.email,
                persona.email_institucional
            ):
                correo = (correo or "").strip().lower()

                if correo:
                    por_correo[correo].add(documento)

        # Detectar correos compartidos por documentos distintos
        # dentro del mismo archivo.
        propietarios_archivo = defaultdict(set)

        for fila in filas:
            if fila["estado"] != "pendiente":
                continue

            for correo in (
                fila["email"],
                fila["email_institucional"]
            ):
                if correo:
                    propietarios_archivo[correo].add(
                        fila["cedula"]
                    )

        salida = []

        for fila in filas:
            participante = dict(fila)

            if participante["estado"] == "pendiente":
                documento = participante["cedula"]
                correos = [
                    x for x in (
                        participante["email"],
                        participante["email_institucional"]
                    ) if x
                ]

                existente = por_documento.get(documento)

                conflicto = False

                for correo in correos:
                    documentos_bd = por_correo.get(correo, set())

                    if any(
                        documento_bd != documento
                        for documento_bd in documentos_bd
                    ):
                        conflicto = True

                    documentos_archivo = propietarios_archivo.get(
                        correo, set()
                    )

                    if any(
                        otro != documento
                        for otro in documentos_archivo
                    ):
                        conflicto = True

                if conflicto:
                    participante["estado"] = "conflicto"

                elif existente:
                    participante["estado"] = "existente"
                    participante["participante_id"] = existente.id

                else:
                    participante["estado"] = "nuevo"

            salida.append(participante)

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
                409,
                "El archivo cambió. Analízalo nuevamente."
            )

        with engine.begin() as conexion:
            if engine.dialect.name == "postgresql":
                conexion.execute(
                    text(
                        "SELECT pg_advisory_xact_lock(hashtext(:evento))"
                    ),
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

            for participante in clasificados:
                accion = participante["estado"]
                participante_id = participante.get("participante_id")

                if accion == "nuevo":
                    token = secrets.token_urlsafe(32)

                    nuevo = conexion.execute(
                        insert(people).values(
                            evento_id=event_id,
                            cedula=participante["cedula"],
                            nombre=participante["nombre"],
                            email=participante["email"],
                            email_institucional=(
                                participante["email_institucional"]
                            ),
                            telefono=participante["telefono"],
                            tipo=participante["tipo"],
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

                elif accion == "existente":
                    # Completar únicamente los correos vacíos.
                    # No se reemplazan los datos que ya existen.
                    actual = conexion.execute(
                        select(
                            people.c.email,
                            people.c.email_institucional
                        ).where(and_(
                            people.c.id == participante_id,
                            people.c.evento_id == event_id
                        )).with_for_update()
                    ).first()

                    if actual:
                        cambios = {}

                        if (
                            not (actual.email or "").strip()
                            and participante["email"]
                        ):
                            cambios["email"] = participante["email"]

                        if (
                            not (actual.email_institucional or "").strip()
                            and participante["email_institucional"]
                        ):
                            cambios["email_institucional"] = (
                                participante["email_institucional"]
                            )

                        if cambios:
                            conexion.execute(
                                update(people)
                                .where(and_(
                                    people.c.id == participante_id,
                                    people.c.evento_id == event_id
                                ))
                                .values(**cambios)
                            )

                conexion.execute(
                    insert(detalles).values(
                        lote_id=lote_id,
                        participante_id=participante_id,
                        fila=participante["fila"],
                        documento=participante["cedula"],
                        accion=accion
                    )
                )

            conexion.execute(
                update(lotes)
                .where(lotes.c.id == lote_id)
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

        for persona in registros:
            try:
                fecha = datetime.fromisoformat(
                    persona.registrado_en.replace("Z", "+00:00")
                )

                hora = fecha.astimezone(ECUADOR).strftime(
                    "%Y-%m-%d %H:%M"
                )

            except (ValueError, AttributeError):
                hora = "Fecha desconocida"

            grupos[hora].append({
                "id": persona.id,
                "nombre": persona.nombre,
                "cedula": persona.cedula,
                "email_estado": persona.email_estado,
                "asistio": bool(persona.asistio_en),
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
                select(attempts.c.id)
                .where(attempts.c.participante_id == pid)
                .limit(1)
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
                400,
                f"Debes escribir DESHACER {lote_id}"
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
                participante = conexion.execute(
                    select(people).where(and_(
                        people.c.id == pid,
                        people.c.evento_id == event_id
                    )).with_for_update()
                ).first()

                if not participante:
                    continue

                correos = conexion.execute(
                    select(attempts.c.id)
                    .where(attempts.c.participante_id == pid)
                    .limit(1)
                ).first()

                otras = conexion.execute(
                    select(detalles.c.id).where(and_(
                        detalles.c.participante_id == pid,
                        detalles.c.lote_id != lote_id
                    )).limit(1)
                ).first()

                if (
                    correos
                    or otras
                    or participante.asistio_en
                    or participante.correo_enviado_en
                    or participante.email_estado != "pendiente"
                    or int(participante.email_intentos or 0) > 0
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
                    estado=(
                        "deshecha_parcial"
                        if protegidos else "deshecha"
                    ),
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
                    people.c.email_institucional,
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
            "Identificación",
            "Nombre",
            "Correo personal",
            "Correo institucional",
            "Teléfono",
            "Tipo",
            "Estado correo"
        ]

        def seguro(valor):
            texto_valor = str(valor or "")

            if texto_valor.lstrip().startswith(
                ("=", "+", "-", "@")
            ):
                return "'" + texto_valor

            return texto_valor

        datos = [
            [seguro(valor) for valor in fila]
            for fila in registros
        ]

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
                celda.font = Font(
                    bold=True,
                    color="FFFFFF"
                )

                celda.fill = PatternFill(
                    "solid",
                    fgColor="176B68"
                )

            for columna, ancho in {
                "A": 22,
                "B": 42,
                "C": 40,
                "D": 40,
                "E": 20,
                "F": 25,
                "G": 23
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
