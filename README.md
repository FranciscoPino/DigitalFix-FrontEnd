# frontend-digitalfix

Aplicación React + Vite + TypeScript para operaciones de terreno DigitalFix. Usa MSAL React para Microsoft Entra ID y consume exclusivamente el BFF.

## Flujo de la aplicación

```text
Login (Entra ID) ─► Roles del token ─► Menú por rol ─► Vistas ─► api.ts ─► BFF ─► Microservicios ─► RabbitMQ ─► Notify
```

1. **Inicio de sesión**: MSAL redirige a Microsoft Entra ID. El token trae el claim `roles`, que se normaliza (`Admin`, `Supervisor`, `Cliente`, `Auditor`).
2. **Menú por rol**: solo se muestran las vistas permitidas (ver tabla abajo).
3. **Llamadas a la API**: cada vista usa `api.ts`, que obtiene el token en silencio, agrega `Authorization: Bearer` y llama al BFF. Aplica timeout y reintenta solo GET (red/timeout/429/5xx); POST/PUT no se reintentan para evitar duplicados.
4. **BFF → microservicios**: el BFF enruta a workorders, catalog, report y audit.
5. **Eventos**: al cambiar el estado de una orden, workorders publica comandos en RabbitMQ y notify los consume (email, despacho, informe PDF).

### Ciclo de vida de una orden

```text
CREATED ─► ASSIGNED ─► EN_ROUTE ─► IN_PROGRESS ─► CLOSED
   (cualquier estado excepto CANCELLED) ─► CANCELLED
```

| Paso en la UI | Quién | Qué ocurre | Mensajes RabbitMQ generados |
|---|---|---|---|
| Crear orden (servicio, dirección, descripción) | Admin, Supervisor, Cliente | `POST /api/workorders`; queda en `CREATED` | ninguno |
| Cambiar estado → `ASSIGNED` (exige elegir técnico) | Admin, Supervisor | `PUT /api/workorders/{id}/status` | `email.send` + `dispatch.ticket` |
| Cambiar estado → `EN_ROUTE` | Admin, Supervisor | técnico en camino | `email.send` |
| Cambiar estado → `IN_PROGRESS` | Admin, Supervisor | inicio del trabajo | `email.send` |
| Cambiar estado → `CLOSED` | Admin, Supervisor | cierre de la orden | `email.send` + `report.gen` |
| Cambiar estado → `CANCELLED` | Admin, Supervisor | cancelación | `email.send` |

En **Órdenes de trabajo**, el botón *Cambiar estado* abre un diálogo con la siguiente transición o la cancelación, y permite agregar una observación opcional. Un Cliente solo ve sus propias órdenes (filtradas por su usuario); Admin y Supervisor ven todas.

## Vistas y roles

| Rol | Vistas/acciones |
|---|---|
| Admin | Resumen, órdenes (crear y cambiar estado), catálogo técnico, repuestos/stock, reportería y auditoría |
| Supervisor | Resumen, órdenes (crear y cambiar estado) y catálogo técnico (lectura) |
| Cliente | Resumen, crear y consultar sus órdenes |
| Auditor | Resumen y auditoría de solo lectura |

Resumen y reportería usan datos de demostración. La seguridad efectiva debe aplicarla también el BFF y los microservicios.

## Demo de RabbitMQ (solo clase)

En **Órdenes de trabajo**, debajo de la tabla y solo para Admin/Supervisor, hay un bloque plegable **"Demo RabbitMQ (solo clase)"**. Permite elegir una cola (`q.cmd.email`, `q.cmd.dispatch`, `q.cmd.report`) y un tipo de mensaje inválido (JSON inválido o envelope sin `type`/`eventId`).

```text
Botón "Enviar a DLQ" ─► BFF POST /api/workorders/demo/dlq ─► workorders publica en cmd.direct
  ─► notify rechaza (NACK sin requeue) ─► el mensaje queda en q.cmd.<cola>.dlq
```

El panel muestra la DLQ de destino. Los detalles se ven en las consolas de workorders y notify, y en `http://localhost:15672` (Queues). No debe habilitarse en producción.

## Configurar Entra ID

Copia `.env.example` a `.env` y completa Tenant ID, Client ID del SPA y API Client ID. Registra `http://localhost:5173` como Redirect URI SPA y concede el scope delegado `api://<API_CLIENT_ID>/access_as_user`.

```env
VITE_AZURE_TENANT_ID=<TENANT_ID>
VITE_AZURE_CLIENT_ID=<SPA_CLIENT_ID>
VITE_AZURE_API_CLIENT_ID=<API_CLIENT_ID>
VITE_API_SCOPE=api://<API_CLIENT_ID>/access_as_user
VITE_REDIRECT_URI=http://localhost:5173
VITE_API_BASE_URL=https://kvl97jinbb.execute-api.us-east-1.amazonaws.com/dev
```

Asigna App Roles `Admin`, `Supervisor`, `Cliente` y `Auditor` a los usuarios/grupos en la Enterprise Application de la API.

Por defecto las llamadas van al stage `/dev` de API Gateway, que valida el JWT y enruta al BFF en EC2. Para desarrollo local contra el BFF desplegado, configura `CORS_ALLOWED_ORIGINS` en los GitHub Actions Variables del repositorio BFF con `http://localhost:5173`.

### Probar el flujo completo en local (con RabbitMQ)

Las acciones del front solo llegan a tu RabbitMQ local si apuntan al BFF local. Crea `.env.local` (ignorado por git y con prioridad sobre `.env`):

```env
VITE_API_BASE_URL=http://localhost:8080
```

Levanta RabbitMQ (5672/15672), `ms-digitalfix-notify` (8086), `ms-digitalfix-workorders` (8081), `ms-digitalfix-catalog` (8082) y `ms-digitalfix-bff` (8080), y luego el front. Con el BFF local sin seguridad (`SECURITY_ENABLED=false`) los endpoints no validan el token. Borra `.env.local` para volver a la nube.

## Ejecutar y probar

```powershell
npm install
npm run dev -- --host localhost
```

Abre `http://localhost:5173/`, inicia sesión y recorre el flujo: crear orden → asignar técnico → en camino → en ejecución → cerrar. En la consola de workorders aparece `LLAMADA A NOTIFY` y en la de notify el mensaje recibido, consumido y procesado en cada cambio. Para verificar el bundle: `npm run build`.
