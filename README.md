# Excel AI Assistant

Complemento de Excel (Add-in) con integración de Azure OpenAI para chat conversacional, análisis de datos y generación de fórmulas.

## Características

- **Chat conversacional**: Interfaz de chat en panel lateral con historial de sesión
- **Lectura de datos**: Lee celdas seleccionadas, rangos y tablas completas
- **Inserción de respuestas**: Escribe respuestas del modelo directamente en celdas
- **Generación de fórmulas**: Crea fórmulas de Excel a partir de lenguaje natural
- **Análisis de datos**: Analiza datos seleccionados con IA
- **Diseño Fluent UI**: Interfaz profesional estilo Microsoft

## Requisitos Previos

### Software
- **Node.js** 18.x o superior
- **npm** 9.x o superior
- **Visual Studio 2022** (opcional, para desarrollo avanzado)
- **Microsoft Excel** 2016 o superior (Desktop o Web)

### Azure OpenAI
- Suscripción de Azure activa
- Recurso de Azure OpenAI creado
- Deployment de GPT-4 configurado
- API Key y Endpoint disponibles

## Instalación

### 1. Clonar/Descargar el proyecto

```bash
cd c:\Proyectos\complementoexcelia
```

### 2. Instalar dependencias

```bash
npm install
```

### 3. Configurar credenciales de Azure OpenAI

Edita el archivo `src/config/config.ts` y reemplaza los valores placeholder:

```typescript
export const config: AppConfig = {
  azureOpenAI: {
    endpoint: "https://tu-recurso.openai.azure.com/",
    apiKey: "tu-api-key-aqui",
    deploymentName: "gpt-4-deployment",
    apiVersion: "2024-02-15-preview",
  },
  // ...
};
```

**Alternativa con variables de entorno:**

Crea un archivo `.env` basado en `.env.example`:

```bash
cp .env.example .env
```

Y configura las variables:

```env
AZURE_OPENAI_ENDPOINT=https://tu-recurso.openai.azure.com/
AZURE_OPENAI_API_KEY=tu-api-key-aqui
AZURE_OPENAI_DEPLOYMENT=gpt-4-deployment
```

### 4. Generar certificados SSL para desarrollo local

```bash
npx office-addin-dev-certs install
```

Esto genera certificados auto-firmados necesarios para HTTPS en localhost.

## Desarrollo Local

### Iniciar el servidor de desarrollo

```bash
npm run dev-server
```

El servidor se iniciará en `https://localhost:3000`.

### Cargar el Add-in en Excel

#### Opción A: Excel Desktop (Windows)

1. Abre Excel
2. Ve a **Insertar** > **Obtener complementos** > **Mis complementos**
3. Haz clic en **Cargar mi complemento**
4. Selecciona el archivo `manifest.xml` del proyecto

#### Opción B: Excel Web

1. Abre Excel en [office.com](https://www.office.com)
2. Ve a **Insertar** > **Complementos de Office**
3. Selecciona **Cargar mi complemento**
4. Sube el archivo `manifest.xml`

#### Opción C: Comando automático (solo Windows)

```bash
npm run start:desktop
```

### Construir para producción

```bash
npm run build
```

Los archivos se generan en la carpeta `dist/`.

## Estructura del Proyecto

```
complementoexcelia/
├── src/
│   ├── config/
│   │   └── config.ts          # Configuración de Azure OpenAI
│   ├── services/
│   │   ├── azureOpenAIService.ts  # Servicio de comunicación con Azure
│   │   └── excelService.ts        # Servicio de interacción con Excel
│   └── taskpane/
│       ├── taskpane.html      # HTML del panel lateral
│       ├── taskpane.css       # Estilos Fluent UI
│       └── taskpane.ts        # Lógica principal
├── assets/                    # Iconos del add-in
├── dist/                      # Archivos compilados (generado)
├── manifest.xml               # Configuración del add-in
├── package.json               # Dependencias y scripts
├── tsconfig.json              # Configuración TypeScript
├── webpack.config.js          # Configuración de bundling
└── README.md
```

## Publicar en Azure Web Apps

### 1. Crear una Azure Web App

```bash
# Instalar Azure CLI si no lo tienes
# https://docs.microsoft.com/cli/azure/install-azure-cli

# Iniciar sesión
az login

# Crear grupo de recursos
az group create --name rg-excel-addin --location westeurope

# Crear App Service Plan
az appservice plan create --name asp-excel-addin --resource-group rg-excel-addin --sku B1

# Crear Web App
az webapp create --name excel-ai-assistant-tuempresa --resource-group rg-excel-addin --plan asp-excel-addin --runtime "NODE:18-lts"
```

### 2. Configurar variables de entorno en Azure

```bash
az webapp config appsettings set --name excel-ai-assistant-tuempresa --resource-group rg-excel-addin --settings \
  AZURE_OPENAI_ENDPOINT="https://tu-recurso.openai.azure.com/" \
  AZURE_OPENAI_API_KEY="tu-api-key" \
  AZURE_OPENAI_DEPLOYMENT="gpt-4-deployment"
```

### 3. Desplegar la aplicación

```bash
# Construir el proyecto
npm run build

# Desplegar usando Azure CLI
az webapp deployment source config-zip --resource-group rg-excel-addin --name excel-ai-assistant-tuempresa --src dist.zip
```

**O usando GitHub Actions:**

Configura un workflow en `.github/workflows/azure-deploy.yml` para CI/CD automático.

### 4. Actualizar el manifest.xml para producción

Cambia todas las URLs de `https://localhost:3000` a tu dominio de Azure:

```xml
<SourceLocation DefaultValue="https://excel-ai-assistant-tuempresa.azurewebsites.net/taskpane.html"/>
```

## Distribuir via Microsoft 365 Admin Center

### Para distribución interna en tu organización:

1. **Accede al Admin Center**
   - Ve a [admin.microsoft.com](https://admin.microsoft.com)
   - Inicia sesión como administrador

2. **Navega a Integrated Apps**
   - Configuración > Integrated apps > Upload custom apps

3. **Sube el manifest**
   - Selecciona "Upload custom app"
   - Elige "Office Add-in"
   - Sube tu `manifest.xml` actualizado con URLs de producción

4. **Configura la distribución**
   - Selecciona los usuarios o grupos que tendrán acceso
   - Revisa y despliega

5. **Los usuarios encontrarán el add-in**
   - En Excel: Insertar > Complementos de Office > Admin Managed

### Requisitos para distribución:
- Licencias Microsoft 365 Business/Enterprise
- Rol de administrador global o administrador de aplicaciones
- El manifest debe apuntar a URLs HTTPS válidas

## Uso del Add-in

1. **Abrir el panel**: Clic en "Excel AI" > "Abrir Chat" en el ribbon
2. **Leer datos**: Selecciona celdas y haz clic en "Leer selección"
3. **Hacer preguntas**: Escribe tu pregunta en el campo de texto
4. **Insertar respuestas**: Selecciona una celda y haz clic en "Insertar respuesta"

### Ejemplos de uso:

```
Usuario: "Crea una fórmula para sumar todas las ventas mayores a 1000"
Asistente: =SUMAR.SI(B:B,">1000")

Usuario: "Analiza estos datos de ventas y dame insights"
Asistente: [Análisis detallado de los datos cargados]

Usuario: "¿Cómo puedo buscar un valor en otra tabla?"
Asistente: [Explicación de BUSCARV/XLOOKUP con ejemplos]
```

## Solución de Problemas

### El add-in no carga
- Verifica que el servidor de desarrollo esté corriendo
- Asegúrate de haber instalado los certificados SSL
- Comprueba que la URL en el manifest sea correcta

### Error de API Key
- Verifica las credenciales en `config.ts`
- Asegúrate de que el deployment exista en Azure OpenAI
- Comprueba los permisos de la API Key

### Error "Host not allowed"
- Añade tu dominio a `<AppDomains>` en el manifest
- Para desarrollo, asegúrate de usar `https://localhost:3000`

### Excel Web no encuentra el add-in
- Verifica que uses HTTPS (no HTTP)
- Limpia la caché del navegador
- Recarga el add-in

## Seguridad

- **No subir credenciales a repositorios públicos**
- Usa variables de entorno en producción
- Considera usar Azure Key Vault para secretos
- El `.gitignore` ya excluye archivos sensibles

## Licencia

Uso interno - Todos los derechos reservados

## Soporte

Para soporte interno, contacta a: [tu-email@empresa.com]
