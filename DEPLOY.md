# Guía completa de despliegue — Portal Consalud EFIKA

---

## PASO 1 — Crear app en Azure (10 min, solo una vez)

> Debes hacer esto con la cuenta **ignacio.rivera@efika.cl** (la de EFIKA)

1. Ve a: https://portal.azure.com
2. Busca y abre **"Registros de aplicaciones"** (App registrations)
3. Clic en **"+ Nuevo registro"**
   - Nombre: `EFIKA Portal Sync`
   - Tipo de cuenta: *Cuentas solo en este directorio*
   - Clic **Registrar**

4. En la app creada, copia y guarda:
   - **Id. de aplicación (cliente)** → es tu `AZURE_CLIENT_ID`
   - **Id. de directorio (inquilino)** → debe ser `bdabe570-1ae7-4cd4-8deb-7243d62c4520`

5. Ve a **"Certificados y secretos"** → **"+ Nuevo secreto de cliente"**
   - Descripción: `sync-portal`
   - Expiración: 24 meses
   - Clic **Agregar** → copia el **Valor** (solo se muestra una vez) → es tu `AZURE_CLIENT_SECRET`

6. Ve a **"Permisos de API"** → **"+ Agregar un permiso"**
   - Microsoft Graph → Permisos de aplicación
   - Busca y selecciona: `Files.Read.All` y `Sites.Read.All`
   - Clic **Agregar permisos**
   - Clic **"Conceder consentimiento de administrador para EFIKA"** → Sí

---

## PASO 2 — Crear repositorio en GitHub (5 min)

1. Ve a https://github.com/new
   - Nombre: `efika-consalud-portal`
   - Visibilidad: **Público** (necesario para GitHub Pages gratis)
   - Clic **Create repository**

2. En tu computador, abre PowerShell en la carpeta `efika-consalud-portal` y ejecuta:

```powershell
cd C:\Users\ignac\OneDrive\Escritorio\efika-consalud-portal

git init
git add .
git commit -m "Portal Consalud EFIKA v1"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/efika-consalud-portal.git
git push -u origin main
```

> Reemplaza `TU_USUARIO` con tu usuario de GitHub

---

## PASO 3 — Configurar secretos en GitHub (3 min)

1. Ve a tu repo en GitHub → **Settings** → **Secrets and variables** → **Actions**
2. Agrega estos 3 secretos (botón **"New repository secret"**):

| Nombre                | Valor                                    |
|-----------------------|------------------------------------------|
| `AZURE_TENANT_ID`     | `bdabe570-1ae7-4cd4-8deb-7243d62c4520`  |
| `AZURE_CLIENT_ID`     | (el que copiaste en Paso 1)              |
| `AZURE_CLIENT_SECRET` | (el que copiaste en Paso 1)              |

---

## PASO 4 — Activar GitHub Pages (2 min)

1. En el repo → **Settings** → **Pages**
2. Source: **Deploy from a branch**
3. Branch: **main** → **/ (root)**
4. Clic **Save**

Tu portal estará en:
**`https://TU_USUARIO.github.io/efika-consalud-portal`**
(disponible en ~2 minutos)

---

## Resultado final

- **Los analistas** editan el Excel en SharePoint normalmente
- **Cada 2 horas**, GitHub Actions descarga el Excel automáticamente
- **data.js** se actualiza y el portal muestra los nuevos datos
- **Nadie necesita correr ningún script** — todo es automático

También puedes forzar una actualización manual entrando al repo en GitHub →
**Actions** → **Sync SharePoint → Portal Consalud** → **Run workflow**

---

## Credenciales del portal

| Usuario   | Contraseña    |
|-----------|---------------|
| consalud  | Consalud2026  |
| victor    | Efika2026     |
| ignacio   | Efika2026     |
| efika     | Efika2026     |
| admin     | Efika2026     |

---

## Actualización manual (opcional)

Si quieres actualizar datos ahora sin esperar las 2 horas:

```powershell
cd C:\Users\ignac\OneDrive\Escritorio\efika-consalud-portal
python sync_data.py --push
```

O para usar un Excel descargado manualmente:

```powershell
python sync_data.py "C:\ruta\al\archivo.xlsx" --push
```
