# Eventos UPS · QR, Brevo API y PostgreSQL (v4)

Sistema independiente, no enlazado al formulario oficial UPS. Deben importar un CSV o Excel proporcionado por responsables autorizados. **No cargar datos reales ni activar envíos sin autorización, pruebas y un plan de tratamiento de datos.** La página de validación continúa en `/escanear`.

## Cambios incluidos

- Importación CSV / XLSX, generación de un QR aleatorio por persona, detección de duplicados de identificación y escaneo único de asistencia.
- Envío **individual** y por **lotes de 20** por API HTTPS de Brevo; QR PNG adjunto y correo en HTML + texto.
- Historial persistente de intentos y estados `pendiente`, `en_proceso`, `aceptado`, `entregado`, `error`, `rebotado`, `incierto`.
- La aceptación por Brevo no implica entrega; `entregado` solo después de webhook. Los timeout de red se marcan `incierto` y **nunca se reintentan automáticamente**.
- Límite diario prudente configurable (200 intentos por defecto); no garantiza cuotas del proveedor. Envía en tandas durante varios días.
- Direcciones terminadas en `@example.com`, `@example.org`, `.test`, `.invalid` se excluyen de envíos para proteger las pruebas.
- PostgreSQL persistente compatible con `DATABASE_URL`; SQLite sigue disponible únicamente para desarrollo. Migración de columnas para una base v3 existente en **el mismo servidor de datos**.
- Ruta `/healthz` para Render; endpoint de webhook protegido con secreto. `/api/consultar` corregido y mensaje de importación corregido.

## 1. Render: despliegue de archivos actualizados

Tu GitHub tiene `sistema_evento_ups_qr_v3/qr_evento_fastapi/`; **reemplaza el contenido de esa carpeta** con los archivos del ZIP v4 (no anides otra carpeta). Conserva la configuración de Render:

- Root Directory: `sistema_evento_ups_qr_v3/qr_evento_fastapi`
- Build Command: `pip install -r requirements.txt`
- Start Command: `uvicorn main:app --host 0.0.0.0 --port $PORT`
- Health Check Path: `/healthz`
- Python: usa versión compatible; `psycopg[binary]` instala un controlador PostgreSQL.

En Render > Environment crea:

| Variable | Valor |
|---|---|
| `QR_ADMIN_KEY` | clave fuerte existente (no cambiar sin coordinar con el equipo) |
| `QR_EVENTO_ID` | `ups-2207` (igual al existente) |
| `DATABASE_URL` | **Internal Database URL** de PostgreSQL de Render, si se aloja en la misma región |
| `BREVO_API_KEY` | clave **API v3**, nunca clave SMTP |
| `BREVO_SENDER_EMAIL` | remitente verificado en Brevo |
| `BREVO_SENDER_NAME` | `Organización Evento UPS` |
| `EMAIL_DAILY_LIMIT` | `200` para comenzar |
| `BREVO_WEBHOOK_SECRET` | cadena aleatoria larga, privada |

### PostgreSQL persistente

En Render selecciona **New > Postgres**, pon nombre (por ejemplo `evento-ups-db`) y la **misma región** del Web Service. Copia *Internal Database URL* en `DATABASE_URL` si ambos están en Render. PostgreSQL conserva filas cuando reinicies/despliegues el Web Service. **No** uses SQLite local de Render para datos oficiales: se pierde durante despliegues. **No se migra automáticamente la información desde SQLite de Render a PostgreSQL.** Si solo tienes los 20 registros de prueba, importa de nuevo el CSV en PostgreSQL y genera nuevos QR; los códigos anteriores no funcionarán. Si ya repartiste QR reales, primero se necesita una migración segura de tokens y asistencias desde la base anterior.

La base PostgreSQL Free de Render **caduca a los 30 días y no ofrece copias de seguridad gestionadas**. Para el evento oficial recomienda un servicio de base con backups o Azure SQL asignado por TI; si usas Render Free solo para prototipos, no lo consideres almacenamiento duradero de producción.

## 2. Brevo: preparar el remitente

En https://app.brevo.com/ crea/accede a la cuenta que autorizará la UPS, verifica un remitente/dominio y genera una **API key v3**. Guarda la clave **solo** en Render > Environment; no en GitHub, CSV ni frontend. Se usa `POST https://api.brevo.com/v3/smtp/email` con encabezado `api-key`. Brevo devuelve `messageId` cuando acepta la solicitud; los estados finales se consultan mediante eventos/webhooks. No necesita conocer ni acceder a las contraseñas de Gmail/Hotmail de participantes. Una cuenta de dominio institucional verificada ayuda a la entregabilidad.

### Seguimiento real de entrega (opcional pero recomendado)

En Brevo crea un webhook **transaccional** para los eventos de entrega y rebote. URL:

`https://TU-SERVICIO.onrender.com/api/brevo/webhook/TU_SECRETO_LARGO`

El secreto debe coincidir exactamente con `BREVO_WEBHOOK_SECRET`. Mantén secreta esa URL y no la publiques en documentación compartida. Habilita eventos `delivered`, `hard_bounce`, `invalid_email` y errores; dependiendo de los nombres de eventos que remita Brevo, se traducen a `entregado`, `rebotado` o `error`. Un webhook sin coincidencia de `messageId` se ignora.

## 3. Ensayo seguro

1. Confirma que `/healthz` devuelve `{"ok":true}` y entra al panel con `QR_ADMIN_KEY`.
2. Importa **solo un participante de prueba con tu dirección real** y otros destinatarios ficticios de `example.com`.
3. Pulsa **Enviar QR** en tu persona de prueba, refresca el panel: pasa de `en_proceso` a `aceptado` si Brevo recibe la solicitud.
4. Confirma en tu bandeja (incluido spam). Si configuras webhook y Brevo informa entrega, el estado se convierte en `entregado`.
5. Comprueba que el segundo clic no reenvía un correo `aceptado`; prueba el QR desde `/escanear` y verifica aviso al segundo ingreso.
6. Antes del uso real, solicita aprobación, comprueba cuota diaria real, permisos, backups, revocación de códigos, roles individuales y plan offline de contingencia.

## 4. Importante para 1.200 personas

Con `EMAIL_DAILY_LIMIT=200`, son **al menos seis días** enviando a 200 personas por día, siempre que la cuenta Brevo permita ese volumen y no existan fallos. Cada clic prepara como máximo 20 intentos; el proceso continua en segundo plano dentro del Web Service. **Render Free puede suspender el proceso**: si queda `en_proceso`, no pulses de nuevo hasta verificar si Brevo envió realmente el mensaje. Para operación fiable en producción, usar un trabajador de tareas persistente y monitorizado, no confiar solo en tareas en memoria.

El lote **no es un programador automático diario**: el organizador debe iniciar tandas desde el panel en los días deseados. El historial muestra los últimos 100 intentos. Si se necesita programación sin intervención, debe incorporarse un worker/cola y un disparador horario con controles de cuota.

## 5. Seguridad y alcance

- La clave única de administrador es válida **solo para demostración**; para producción incorporar cuentas individuales, roles, MFA y auditoría.
- No se almacena contraseña de correo de participantes. El QR contiene token aleatorio, no cédula. El servidor muestra los datos solo con clave de organizador.
- Si se importan filas repetidas, no se renuevan QR ni se modifican asistentes existentes.
- WhatsApp abre un chat manual sin adjuntar el QR. Integración oficial de WhatsApp Business todavía **no está implementada**.
- Las limitaciones del plan de correo pueden cambiar; verifica el límite y la aprobación del proveedor antes de repartir 1.200 credenciales.

## 6. Arranque local

```powershell
py -m pip install -r requirements.txt
$env:QR_ADMIN_KEY='CLAVE_PRIVADA_LARGA'
$env:BREVO_API_KEY='SOLO_SI_VAS_A_PROBAR_CORREO'
$env:BREVO_SENDER_EMAIL='remitente@dominio-verificado.com'
py -m uvicorn main:app --reload
```

Sin `DATABASE_URL` usa SQLite local. Para producción, la base de datos debe estar configurada antes de la primera importación.


## V5: Corrección de datos y eliminación segura de pruebas
- Al importar Excel/CSV, las identificaciones nuevas se agregan; las existentes se actualizan solo si no tienen historial de envío ni asistencia. No se eliminan registros ausentes del archivo.
- Botón **Eliminar** para participantes sin envíos ni asistencia. Protegido por `QR_ADMIN_KEY`.
- Botón **Borrar registros TEST**: elimina exclusivamente documentos cuyo valor comienza por `TEST` que no tengan envíos ni asistencias. Requiere confirmación explícita.
- El campo `protegidos` informa los registros que necesitan corrección manual auditada.
- No ejecuta borrados masivos por mera importación de Excel. Respaldar PostgreSQL antes de eliminar datos.
