# Sistema de Control de Asistencia — Bot de Telegram

Sistema completo de registro de asistencia para empresas, basado en un bot de Telegram. Los empleados fichan desde el celular con un mensaje de texto; el administrador recibe reportes automáticos en Excel y puede monitorear el estado en tiempo real desde un panel web.

---

## Cómo funciona

### Para el empleado

El empleado abre el chat del bot en Telegram y escribe un mensaje natural. No necesita aprender comandos.

**Registrar entrada:**
```
llegué
```
```
entro
```
```
buenos días
```
```
/entro
```

**Registrar salida:**
```
me voy
```
```
salgo
```
```
chau
```
```
/salgo
```

**Ver estado del día:**
```
/estado
```

El bot responde confirmando el registro con nombre, hora y — al salir — el desglose de horas trabajadas:

```
Salida registrada
Nombre: Juan Pérez
Hora:   16:08
Hs. normales:    9.00
Hs. extra 50%:   0.00
```

---

### Primer uso (registro automático)

La primera vez que un empleado escribe al bot, el sistema le pide su nombre:

```
Bot: Hola! Para registrarte necesito tu nombre completo. ¿Cómo te llamás?
Empleado: Juan Pérez
Bot: Bienvenido/a, Juan Pérez! Ya estás registrado en el sistema de asistencia.
```

A partir de ese momento, el bot reconoce al empleado por su cuenta de Telegram (ligada al número de celular). No necesita identificarse nunca más.

---

## Cálculo de horas — Convenio

| Franja horaria | Tipo |
|----------------|------|
| Lun–Vie 07:00–16:00 | Horas normales |
| Lun–Vie antes de 07:00 o después de 16:00 | Horas extra 50% |
| Sábado 07:00–13:00 | Horas extra 50% |
| Sábado 13:00 en adelante | Horas extra 100% |
| Domingo | Horas extra 100% |
| Feriados nacionales | Horas extra 100% |

### Tolerancia de entrada y salida

El sistema aplica un margen de **15 minutos** (configurable) en los límites de cada zona:

| Situación | Resultado |
|-----------|-----------|
| Entra a las 07:08 | Se cuenta como 07:00 — dentro del margen |
| Entra a las 07:20 | Se cuenta desde las 07:20 — fuera del margen, pierde 20 min |
| Sale a las 15:52 | Se cuenta como 16:00 — jornada completa |
| Sale a las 15:38 | Se cuenta desde las 15:38 — fuera del margen |
| Sábado: sale 10:53 | Se cuenta como 11:00 — todo al 50%, sin extra 100% |

El margen se configura en el archivo `.env` con `GRACE_MINUTES=15`.

---

## Reporte Excel

Cada quincena (día 15 y último día del mes) el bot envía automáticamente un archivo `.xlsx` al administrador por Telegram y por email.

### Estructura del archivo

**Hoja "Resumen"** — una fila por empleado:

| Empleado | Hs. Normales | Hs. Extra 50% | Hs. Extra 100% |
|----------|-------------|---------------|----------------|
| Juan Pérez | 180.00 | 12.50 | 0.00 |
| María García | 162.00 | 8.00 | 16.00 |
| **TOTAL GENERAL** | **342.00** | **20.50** | **16.00** |

**Una hoja por empleado** — todos los días del período:

| Fecha | Día | Estado | Entrada | Salida | Hs. Normales | Hs. Extra 50% | Hs. Extra 100% |
|-------|-----|--------|---------|--------|-------------|---------------|----------------|
| 01/06 | Lunes | Trabajó | 07:05 | 16:10 | 9.00 | 0.00 | 0.00 |
| 07/06 | Domingo | Fin de semana | — | — | — | — | — |
| 15/06 | Lunes | Feriado: … | — | — | — | — | — |
| 20/06 | Sábado | Trabajó | 07:00 | 13:00 | 0.00 | 4.00 | 2.00 |

### Código de colores

| Color | Significado |
|-------|-------------|
| Gris | Fin de semana |
| Dorado/amarillo | Feriado nacional |
| Azul claro | Vacación |
| Verde claro | Enfermedad |
| Naranja claro | Licencia |
| Amarillo | Sin salida registrada |
| Rojo claro | Ausente (día hábil sin registro) |
| Dorado (fila) | Día con horas extra |

---

## Panel web del administrador

Acceso: `http://localhost:8501` desde la PC de la oficina.

### Pestaña "Hoy"

- Métricas en tiempo real: cuántos empleados están dentro, cuántos no ficharon, cuántos tienen salida pendiente.
- Tabla con el estado actual de cada empleado (entrada, salida, horas acumuladas del día).
- Botón Actualizar para refrescar los datos.

### Pestaña "Empleados"

- Lista completa de empleados registrados con sus turnos asignados.
- Selector para ver los últimos 15 fichajes de cualquier empleado.

### Pestaña "Reportes"

- Selector de período: mes actual, año completo, o rango personalizado.
- Botón para generar y descargar el Excel directamente desde el navegador.
- Estadísticas rápidas del mes: total de horas por empleado.

---

## Comandos del administrador (por Telegram)

### Reportes

| Comando | Acción |
|---------|--------|
| `/reporte` | Genera y envía el Excel del año en curso |

### Feriados

| Comando | Ejemplo | Acción |
|---------|---------|--------|
| `/feriado DD/MM nombre` | `/feriado 25/12 Navidad` | Agrega un feriado |
| `/borrarf DD/MM` | `/borrarf 25/12` | Elimina un feriado |
| `/feriados` | `/feriados` | Lista los feriados del mes actual |

### Ausencias

| Comando | Ejemplo | Acción |
|---------|---------|--------|
| `/ausencia Nombre tipo DD/MM` | `/ausencia Juan Pérez vacacion 15/07` | Registra una ausencia |
| `/ausencias` | `/ausencias` | Lista las ausencias del mes actual |

**Tipos de ausencia:** `vacacion`, `enfermedad`, `licencia`, `justificada`, `injustificada`

### Corrección de registros

| Comando | Ejemplo | Acción |
|---------|---------|--------|
| `/corregir Nombre entrada HH:MM DD/MM` | `/corregir Juan entrada 08:30 25/06` | Corrige hora de entrada |
| `/corregir Nombre salida HH:MM DD/MM` | `/corregir Juan salida 17:00 25/06` | Corrige hora de salida |
| `/borrar Nombre DD/MM` | `/borrar Juan 25/06` | Elimina el registro de un día |
| `/ver Nombre` | `/ver Juan` | Muestra los últimos 10 fichajes |

### Turnos

| Comando | Ejemplo | Acción |
|---------|---------|--------|
| `/turno Nombre HH:MM HH:MM` | `/turno Juan Pérez 08:00 17:00` | Asigna turno a un empleado |

### Empleados

| Comando | Acción |
|---------|--------|
| `/empleados` | Lista todos los empleados registrados |

---

## Notificaciones automáticas

| Hora | Qué hace |
|------|----------|
| 08:00 | Confirma al admin que el bot está activo |
| 09:00–12:00 | Avisa al admin qué empleados no ficharon entrada (respeta el turno individual de cada uno); envía recordatorio directo a cada empleado faltante |
| 18:30 | Recuerda a los empleados con entrada abierta que registren su salida; avisa al admin |
| 18:00 (día 15 y último del mes) | Genera y envía el reporte de quincena automáticamente |
| 23:00 | Backup automático de la base de datos |

---

## Arquitectura del sistema

```
attendance-bot/
├── bot.py            # Bot de Telegram — lógica de comandos y mensajes
├── database.py       # Base de datos SQLite — empleados, asistencia, feriados, ausencias, turnos
├── reports.py        # Generación de XLSX y envío de email
├── admin_panel.py    # Panel web Streamlit (localhost:8501)
├── iniciar_bot.bat           # Arranca el bot (con auto-reinicio)
├── iniciar_bot_oculto.vbs    # Versión silenciosa para el inicio de Windows
├── iniciar_panel.bat         # Arranca el panel web
├── iniciar_panel_oculto.vbs  # Versión silenciosa para el inicio de Windows
├── attendance.db     # Base de datos (generada al primer arranque)
├── backups/          # Copias diarias de la base de datos
├── .env              # Variables de entorno (token, email, configuración)
└── requirements.txt  # Dependencias Python
```

### Base de datos (SQLite)

| Tabla | Contenido |
|-------|-----------|
| `employees` | Empleados registrados (Telegram ID, nombre) |
| `attendance` | Registros de entrada/salida con horas calculadas |
| `holidays` | Feriados nacionales (pre-cargados 2026 + editables) |
| `absences` | Ausencias por empleado (vacación, enfermedad, etc.) |
| `shifts` | Turno de cada empleado (hora de entrada y salida) |

---

## Configuración (.env)

```env
# Bot de Telegram
TELEGRAM_TOKEN=tu_token_de_botfather
ADMIN_IDS=tu_id_de_telegram

# Base de datos
DB_PATH=C:\ruta\attendance.db

# Tolerancia de entrada/salida en minutos
GRACE_MINUTES=15

# Email (opcional)
EMAIL_FROM=tucuenta@gmail.com
EMAIL_PASSWORD=xxxx xxxx xxxx xxxx
EMAIL_TO=destino@gmail.com

# Geolocalización (opcional)
REQUIRE_LOCATION=false
OFFICE_LAT=-32.9468
OFFICE_LON=-60.6393
OFFICE_RADIUS_METERS=300
```

---

## Instalación

### Requisitos
- Windows 10/11
- Python 3.10+
- Cuenta de Telegram

### Pasos

**1. Crear el bot en Telegram**

Buscar `@BotFather` en Telegram → `/newbot` → guardar el token.

Para obtener tu ID de admin: buscar `@userinfobot` → enviar cualquier mensaje → copiar el `Id`.

**2. Instalar dependencias**

```bash
py -m venv venv
venv\Scripts\pip install -r requirements.txt
```

**3. Configurar el archivo .env**

Copiar `.env.example` a `.env` y completar los valores.

**4. Agregar al inicio de Windows**

Copiar `iniciar_bot_oculto.vbs` y `iniciar_panel_oculto.vbs` a la carpeta:
```
%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\
```

**5. Arrancar manualmente la primera vez**

```bash
venv\Scripts\python.exe bot.py
venv\Scripts\streamlit.exe run admin_panel.py --server.port 8501
```

El panel queda disponible en `http://localhost:8501`.

---

## Feriados nacionales Argentina 2026 (pre-cargados)

| Fecha | Feriado |
|-------|---------|
| 01/01 | Año Nuevo |
| 16/02 | Carnaval |
| 17/02 | Carnaval |
| 24/03 | Día de la Memoria |
| 02/04 | Día del Veterano de Malvinas |
| 03/04 | Viernes Santo |
| 01/05 | Día del Trabajador |
| 25/05 | Revolución de Mayo |
| 20/06 | Paso a la Inmortalidad del Gral. Güemes |
| 09/07 | Día de la Independencia |
| 17/08 | Paso a la Inmortalidad del Gral. San Martín |
| 12/10 | Día de la Diversidad Cultural |
| 20/11 | Día de la Soberanía Nacional |
| 08/12 | Inmaculada Concepción |
| 25/12 | Navidad |

---

## Stack tecnológico

| Componente | Tecnología |
|-----------|-----------|
| Bot de mensajería | python-telegram-bot 21.6 |
| Base de datos | SQLite (via Python sqlite3) |
| Panel web | Streamlit 1.58 |
| Reportes Excel | openpyxl 3.1 |
| Scheduler interno | APScheduler (via python-telegram-bot job-queue) |
| Email | smtplib (Gmail SMTP SSL) |
| Lenguaje | Python 3.14 |
| Hosting | PC de oficina (Windows, inicio automático) |

---

## Repositorio

[github.com/Lucho-code/attendance-bot](https://github.com/Lucho-code/attendance-bot) (privado)
