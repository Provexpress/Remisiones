# Portal de remisiones abiertas · Provexpress

Aplicación web para revisar las remisiones pendientes por facturar de Provexpress. Lee `Remisiones.xlsx` directamente desde SharePoint después de autenticar al usuario con Microsoft 365, cruza la hoja `Base` con `Grupos` y recalcula todos los indicadores en el navegador.

- Producción: https://remisiones-provexpress-projects.vercel.app
- Repositorio: https://github.com/Provexpress/Remisiones

## Tecnologías

- React 19 + TypeScript + Vite.
- Microsoft Authentication Library (MSAL Browser 5).
- Microsoft Graph para descargar el Excel compartido.
- ExcelJS para leer el libro sin enviarlo a otro servicio.
- Recharts para indicadores y tendencias.
- Configuración de compilación lista para Vercel.

El archivo y los tokens permanecen en el navegador. No se incluyó `Remisiones.xlsx` en los recursos públicos de la aplicación.

## Funcionalidad incluida

- Inicio de sesión con cuentas del tenant de Provexpress.
- Descarga automática de `Remisiones.xlsx` desde el vínculo de SharePoint suministrado.
- Actualización manual y automática cada 10 minutos.
- Apertura de un Excel local como respaldo.
- KPIs del corte: pendiente, número de remisiones, clientes, valor de mercancía, IVA, antigüedad, vencidas y cantidad cero.
- Comparación con el corte anterior.
- Tendencia por periodo, nuevas remisiones, retiros estimados y reducción bruta.
- Distribución por antigüedad, director y ejecutivo.
- Cruce tolerante a nombres abreviados entre `Base` y `Grupos`.
- Aviso visible para nombres sin grupo.
- Tabla operativa con búsqueda, filtros, paginación y exportación CSV.
- Diseño adaptable a escritorio, tableta y móvil.

## Estructura esperada del Excel

La aplicación busca estas hojas:

- `Base-SIS`: Hoja principal y simplificada. El operador únicamente copia y pega la exportación directa de SIS sin requerir columna de fecha ni cálculos adicionales. Contiene las columnas `Empleado`, `NIT`, `Empresa`, `Vr. Mercancia`, `Vr. IVA`, `Vr. Total`, `Emision`, `Dias`, `Documento`, `Pedido` y `Cantidad`. El portal toma automáticamente la fecha y hora de modificación del Excel/SharePoint como corte.
- `Base`: Fuente alternativa o histórica. Si contiene `Fecha_Corte`, se lee para reconstruir múltiples cortes.
- `Grupos`: Bloques con el formato `Grupo N — Director(a): Nombre`, seguidos por los ejecutivos de ese grupo.
- `Diario`: Resumen consolidado opcional de gestión diaria que alimenta el historial de entradas y retiros.

El portal almacena además el historial diario en el almacenamiento local del navegador (`localStorage`), permitiendo conservar la evolución histórica de días anteriores incluso cuando el operador sobreescribe `Base-SIS` diariamente.

## Ejecución local

```powershell
npm install
npm run dev
```

Abrir `http://localhost:5173`. Para el inicio Microsoft también se debe registrar esta URI de redirección como plataforma **Single-page application**:

```text
http://localhost:5173/redirect.html
```

## Configuración de Microsoft Entra ID

Los identificadores de la aplicación se configuran mediante variables de entorno. Los valores del entorno de Provexpress se obtuvieron del proyecto local `C:\Proyectos\ForeCast` y no se publican en este repositorio.

En **Microsoft Entra admin center → App registrations → la aplicación → Authentication**:

1. Agregar la plataforma **Single-page application**.
2. Registrar `http://localhost:5173/redirect.html`.
3. Después de crear el proyecto Vercel, registrar también `https://NOMBRE-PROYECTO.vercel.app/redirect.html`.
4. Si se usa un dominio propio, registrar igualmente `https://DOMINIO/redirect.html`.

MSAL Browser 5 usa `redirect.html` como puente para el flujo emergente. La URI debe coincidir exactamente con la registrada en Entra ID. Referencia: https://learn.microsoft.com/en-us/entra/msal/javascript/browser/redirect-bridge

En **API permissions**, confirmar estos permisos delegados de Microsoft Graph:

```text
User.Read
Files.Read.All
```

El proyecto `ForeCast` ya solicita `Files.Read.All`; se conserva ese permiso para que los usuarios autorizados puedan leer el archivo compartido. La aplicación no solicita permisos de escritura. Microsoft documenta `Files.Read` como permiso mínimo para descargar contenido y `Files.Read.All` como opción de mayor alcance: https://learn.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0

Los usuarios también deben tener acceso al Excel en SharePoint/OneDrive. La autenticación en el tenant no reemplaza los permisos del archivo.

## Variables de entorno

Copiar `.env.example` a `.env.local` para desarrollo si se desea cambiar la configuración. En Vercel, configurar las mismas variables en **Project Settings → Environment Variables**:

```text
VITE_AZURE_CLIENT_ID
VITE_AZURE_TENANT_ID
VITE_SHAREPOINT_FILE_URL
```

Opcionalmente se pueden usar identificadores directos de Graph, más estables que un vínculo compartido:

```text
VITE_GRAPH_DRIVE_ID
VITE_GRAPH_ITEM_ID
```

Las variables `VITE_*` se incorporan al frontend durante la compilación y no deben contener secretos. El control de acceso real lo realizan Entra ID, el token delegado y SharePoint. Vercel exige el prefijo `VITE_` para exponer variables a una aplicación Vite: https://vercel.com/docs/frameworks/frontend/vite

## Despliegue en Vercel

El archivo `vercel.json` ya declara Vite, `npm run build`, la carpeta `dist` y cabeceras de seguridad.

```powershell
npm install
npm test
npm run build
npx vercel
npx vercel --prod
```

Después del primer despliegue:

1. Copiar el dominio definitivo de Vercel.
2. Registrar su `/redirect.html` en la aplicación de Entra ID.
3. Configurar las variables para Production y Preview.
4. Volver a desplegar; los cambios de variables solo aplican a despliegues nuevos: https://vercel.com/docs/environment-variables/managing-environment-variables

## Verificación

```powershell
npm test
npm run build
```

Las pruebas siempre ejecutan un libro sintético seguro. Si `Remisiones.xlsx` existe localmente, también concilian los valores principales del dashboard, las 674 filas, el corte y los cruces de nombres con directores. El archivo real está excluido de Git.
