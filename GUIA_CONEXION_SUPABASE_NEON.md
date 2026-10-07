# 🚀 GUÍA PASO A PASO: CONEXIÓN EN LA NUBE CON SUPABASE O NEON

Esta guía te explica cómo conectar el sistema de facturación a una base de datos PostgreSQL en la nube (**Supabase** o **Neon**) sin costo alguno.

> **¿Por qué conectar a la nube?**  
> En el plan gratuito de plataformas como Render, cuando la aplicación entra en suspensión por inactividad, los archivos locales se reinician. Al conectar una base de datos en la nube (PostgreSQL), **toda tu información (empresas, prefacturas, facturas, órdenes, fechas y valores) queda protegida permanentemente y nunca se borra**.

---

## 🛠️ OPCIÓN 1: SUPABASE (Recomendada: Base de Datos + Almacenamiento de Archivos)

Supabase te ofrece una base de datos PostgreSQL de alto rendimiento con 500 MB de almacenamiento gratuito y 1 GB para guardar archivos (PDFs de facturas e informes en Excel).

### Paso 1: Crear la cuenta y el proyecto
1. Ingresa a [https://supabase.com/](https://supabase.com/) y haz clic en **Start your project** (puedes ingresar con tu cuenta de GitHub o Google).
2. Haz clic en **"New project"**.
3. Diligencia los siguientes campos:
   - **Name:** `control-facturacion` (o el nombre que prefieras).
   - **Database Password:** Crea una contraseña segura (guárdala en un bloc de notas, la necesitarás).
   - **Region:** Selecciona la región más cercana (ejemplo: `East US (North Virginia)` o `South America (São Paulo)`).
   - **Pricing Plan:** Free ($0/mes).
4. Haz clic en **"Create new project"** y espera ~1 minuto a que termine de aprovisionar.

### Paso 2: Obtener la cadena de conexión (`DATABASE_URL`)
1. En el menú lateral izquierdo de Supabase, haz clic en el ícono de engranaje **Project Settings** (abajo a la izquierda).
2. Ve a la sección **Database**.
3. Baja hasta el apartado **Connection String**.
4. Selecciona la pestaña **URI**.
5. Copia la URL que aparece, la cual tiene esta estructura:
   ```text
   postgresql://postgres:[YOUR-PASSWORD]@db.xxxxxxxxxxxxxxxxxxxx.supabase.co:5432/postgres
   ```
   > ⚠️ **Importante:** Reemplaza `[YOUR-PASSWORD]` por la contraseña que creaste en el Paso 1. Si tu contraseña tiene caracteres especiales como `@`, `#` o `/`, usa la versión codificada en URL o una contraseña alfanumérica.

### Paso 3 (Opcional pero recomendado): Activar Supabase Storage para PDFs e Informes
1. En el menú lateral izquierdo, haz clic en **Storage**.
2. Haz clic en **"New bucket"**.
3. Asigna el nombre: `facturas-docs`.
4. Marca la casilla **"Public bucket"** (esto permite que los botones "Ver PDF Factura" y "Ver Informe" abran los documentos directamente en el navegador).
5. Haz clic en **Save**.
6. Ve a **Project Settings > API**:
   - Copia la **Project URL** (ejemplo: `https://xxxxxxxxxxxxxxxxxxxx.supabase.co`).
   - Copia la clave **service_role** o **anon public**.

---

## ⚡ OPCIÓN 2: NEON (Alternativa Ultra-Rápida Serverless Postgres)

Si solo deseas la base de datos PostgreSQL sin almacenamiento de archivos externo:

1. Ingresa a [https://neon.tech/](https://neon.tech/) e inicia sesión con GitHub o Google.
2. Haz clic en **"Create project"**.
3. Nómbralo `control-facturacion` y haz clic en **Create Project**.
4. En el dashboard principal aparecerá inmediatamente la cadena de conexión:
   ```text
   postgresql://tu_usuario:tu_password@ep-xxxx-xxxx.us-east-2.aws.neon.tech/neondb?sslmode=require
   ```
5. Cópiala directamente.

---

## 🌐 PASO 4: CONFIGURAR EN RENDER.COM

Una vez tengas tu cadena de conexión:

1. Ingresa a tu cuenta de [https://dashboard.render.com/](https://dashboard.render.com/).
2. Haz clic sobre tu servicio web (ejemplo: `facturacion-aeropuertos`).
3. En el menú de la izquierda, haz clic en **Environment**.
4. Haz clic en **"Add Environment Variable"** y agrega:
   - **Key:** `DATABASE_URL`
   - **Value:** `postgresql://postgres:tu_password@db.xxxx.supabase.co:5432/postgres` (la URL que copiaste).
5. *(Opcional si configuraste Supabase Storage)*:
   - **Key:** `SUPABASE_URL` | **Value:** `https://xxxx.supabase.co`
   - **Key:** `SUPABASE_KEY` | **Value:** `tu_clave_anon_o_service_role`
   - **Key:** `SUPABASE_BUCKET` | **Value:** `facturas-docs`
6. Haz clic en **"Save Changes"**.
7. Render iniciará automáticamente un nuevo despliegue (*Deploying...*).

---

## 🔍 PASO 5: VERIFICACIÓN AUTOMÁTICA EN EL SISTEMA

El sistema cuenta con autodetección y auto-inicialización inteligente:

1. **Auto-Migración:** En el primer arranque contra Supabase o Neon, el servidor creará automáticamente todas las tablas relacionales (`clients`, `billing_records`, `system_users`, `uploaded_files_archive`).
2. **Siembra Maestra:** Sembrará automáticamente las 41 empresas maestras oficiales y los usuarios configurados (`master`, `Sol`, `Edwar`), **totalmente limpios de datos preliminares de prueba ($0 prefacturado y $0 facturado)**.
3. **Indicador Visual en Vivo:**  
   Al entrar a la aplicación web, mira el badge en el encabezado superior:
   - Si dice: `Cloud: PostgreSQL (Supabase/Neon) 🟢` -> **¡Conectado exitosamente a la nube permanente!**
   - Si dice: `Almacenamiento Local (SQLite) 🟡` -> Significa que no se configuró `DATABASE_URL` y el sistema está funcionando en modo local seguro.

---

## 🛡️ RESPALDO Y TOLERANCIA A FALLOS (FALLBACK)

- **Cero tiempo de inactividad:** Si por algún motivo temporal la nube no responde o no tiene internet, el sistema activa automáticamente el motor local SQLite (`facturacion.db`), garantizando que Sol y Edwar nunca se queden sin poder trabajar.
- **Disaster Recovery de Archivos:** Incluso con archivos locales, todos los PDFs y hojas de cálculo se almacenan simultáneamente como binarios cifrados en la base de datos (`uploaded_files_archive`). Si un archivo se borra del disco de Render, el sistema lo reconstruye automáticamente en el siguiente clic.
