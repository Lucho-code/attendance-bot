# Bot de Control de Asistencia - Guía de instalación

## Paso 1: Crear el bot en Telegram (5 minutos)

1. Abrir Telegram y buscar **@BotFather**
2. Enviar `/newbot`
3. Ponerle un nombre, por ejemplo: `Control de Asistencia Empresa`
4. Ponerle un usuario (debe terminar en `bot`), por ejemplo: `asistencia_miempresa_bot`
5. BotFather te da un **token** (parece: `123456789:ABCdef...`). Guardarlo.

## Paso 2: Saber tu ID de administrador

1. En Telegram buscar **@userinfobot**
2. Enviar cualquier mensaje
3. Te responde con tu `Id:` — ese es tu ADMIN_ID

## Paso 3: Deploy en Railway

### 3a. Crear cuenta en Railway
Ir a [railway.app](https://railway.app) y crear cuenta gratuita con GitHub.

### 3b. Subir el código a GitHub
```bash
# En la carpeta attendance-bot:
git init
git add .
git commit -m "Bot de asistencia"
# Crear repo en github.com y luego:
git remote add origin https://github.com/TU_USUARIO/attendance-bot.git
git push -u origin main
```

### 3c. Crear proyecto en Railway
1. En Railway: **New Project → Deploy from GitHub repo**
2. Seleccionar el repo `attendance-bot`
3. Railway lo detecta automáticamente

### 3d. Configurar variables de entorno
En Railway, ir a tu proyecto → **Variables** → agregar:

| Variable         | Valor                          |
|-----------------|-------------------------------|
| `TELEGRAM_TOKEN` | El token de BotFather         |
| `ADMIN_IDS`      | Tu ID de Telegram             |
| `DB_PATH`        | `/data/attendance.db`         |

### 3e. Crear volumen para la base de datos (importante)
Sin esto, los datos se borran cada vez que Railway reinicia el bot.

1. En tu proyecto Railway → **Add Volume**
2. Mount path: `/data`
3. Hacer **Redeploy**

---

## Uso del bot

### Empleados
| Acción          | Qué escribir                                      |
|----------------|--------------------------------------------------|
| Registrarse     | `/start` (solo la primera vez)                   |
| Registrar entrada | `/entro` o escribir "llegué", "buenos días"  |
| Registrar salida  | `/salgo` o escribir "me voy", "chau"          |
| Ver estado hoy  | `/estado`                                        |

### Administrador
| Acción                  | Comando        |
|------------------------|----------------|
| Descargar CSV completo  | `/reporte`     |
| Ver empleados registrados | `/empleados` |

---

## Estructura del CSV exportado

```
Nombre,Fecha,Entrada,Salida,Horas_Trabajadas
Juan Pérez,2026-06-25,08:30,17:30,9.00
María García,2026-06-25,09:00,Sin salida,
```

El archivo se abre correctamente en Excel (incluye marca BOM para caracteres en español).

---

## Costos

| Servicio | Costo      |
|----------|-----------|
| Telegram Bot API | Gratis |
| Railway (Hobby plan) | ~$5 USD/mes o gratis con créditos iniciales |
| Base de datos SQLite | Gratis (incluida) |

---

## Preguntas frecuentes

**¿Qué pasa si un empleado no registra salida?**
Aparece como "Sin salida" en el CSV. Podés editar la base de datos manualmente si es necesario.

**¿Pueden registrarse varias entradas el mismo día?**
No, el bot detecta si ya registraste entrada y avisa.

**¿Se pueden agregar más admins después?**
Sí, agregá los IDs separados por coma en `ADMIN_IDS`: `123456789,987654321`
