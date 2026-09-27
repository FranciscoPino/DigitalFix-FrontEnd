# frontend-digitalfix

Aplicación React + Vite + TypeScript para operaciones de terreno DigitalFix. Usa MSAL React para Microsoft Entra ID y consume exclusivamente el BFF.

## Arquitectura y patrón

```text
React views -> api.ts (MSAL token + fetch) -> BFF :8080 -> APIs de dominio
```

La UI separa navegación por rol y funciones por vistas. `api.ts` centraliza autenticación HTTP, timeout y retry limitado: reintenta GET en red/timeout/429/5xx; no reintenta POST/PUT para evitar operaciones duplicadas. Los roles se leen del claim `roles` (lista) y se normalizan sin distinguir mayúsculas.

## Vistas y roles

| Rol | Vistas/acciones |
|---|---|
| Admin | Órdenes, catálogo técnico, repuestos/stock, reportería y auditoría |
| Supervisor | Órdenes, cambios de estado y catálogo técnico |
| Cliente | Crear y consultar sus órdenes |
| Auditor | Resumen y auditoría de solo lectura |

Las vistas de reportería/auditoría usan datos de demostración; la integración Kafka queda fuera del frontend actual. La seguridad efectiva debe aplicarla también el BFF.

## Configurar Entra ID

Copia `.env.example` a `.env` y completa Tenant ID, Client ID del SPA y API Client ID. Registra `http://localhost:5173` como Redirect URI SPA y concede el scope delegado `api://<API_CLIENT_ID>/access_as_user`.

```env
VITE_AZURE_TENANT_ID=<TENANT_ID>
VITE_AZURE_CLIENT_ID=<SPA_CLIENT_ID>
VITE_AZURE_API_CLIENT_ID=<API_CLIENT_ID>
VITE_API_SCOPE=api://<API_CLIENT_ID>/access_as_user
VITE_REDIRECT_URI=http://localhost:5173
VITE_API_BASE_URL=http://localhost:8080
```

Asigna App Roles `Admin`, `Supervisor`, `Cliente` y `Auditor` a los usuarios/grupos en la Enterprise Application de la API.

## Ejecutar y probar

```powershell
npm install
npm run dev -- --host localhost
```

Abre `http://localhost:5173/`, inicia sesión y prueba creación/consulta de órdenes, cambios de estado, servicios y administración de repuestos. Deben estar levantados BFF y microservicios. Para verificar el bundle: `npm run build`.
