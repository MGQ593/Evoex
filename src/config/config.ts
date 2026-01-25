/**
 * Configuración de Azure OpenAI
 *
 * IMPORTANTE: No subir este archivo con credenciales reales a repositorios públicos.
 * Para producción, usar variables de entorno o Azure Key Vault.
 */

// Importar variables de entorno desde archivo generado por webpack
import { ENV } from "./env.generated";

/**
 * Definición de un modelo disponible
 */
export interface ModelDefinition {
  id: string;
  name: string;
  description: string;
}

/**
 * Parsea los modelos desde la variable de entorno AVAILABLE_MODELS
 * Formato: JSON array de objetos {id, name, description}
 * Ejemplo: [{"id":"gpt-4o","name":"GPT-4o","description":"Descripción"}]
 */
function parseModelsFromEnv(): ModelDefinition[] {
  const envModels = ENV.AVAILABLE_MODELS;

  if (!envModels) {
    console.warn("[Config] AVAILABLE_MODELS no está definida en .env");
    return [];
  }

  try {
    const parsed = JSON.parse(envModels);
    if (Array.isArray(parsed) && parsed.length > 0) {
      return parsed
        .filter((m: any) => m.id && m.name && typeof m.id === "string" && typeof m.name === "string")
        .map((m: any) => ({
          id: m.id,
          name: m.name,
          description: m.description || "",
        }));
    }
  } catch (e) {
    console.error("[Config] Error parseando AVAILABLE_MODELS:", e);
  }

  return [];
}

/**
 * Lista de modelos disponibles (SOLO desde variable de entorno)
 */
export const availableModels: ModelDefinition[] = parseModelsFromEnv();

/**
 * Modelo predeterminado (desde variable de entorno o el primero disponible)
 */
export const defaultModelId = ENV.DEFAULT_MODEL_ID || availableModels[0]?.id || "";

/**
 * URL del proxy SearXNG
 */
export const searxngProxyUrl = ENV.SEARXNG_PROXY_URL || "";

export interface AzureOpenAIConfig {
  endpoint: string;
  apiKey: string;
  deploymentName: string;
  apiVersion: string;
}

export interface AppConfig {
  azureOpenAI: AzureOpenAIConfig;
  maxTokens: number;
  temperature: number;
  systemPrompt: string;
}

/**
 * Configuración de la aplicación
 */
export const config: AppConfig = {
  azureOpenAI: {
    endpoint: ENV.AZURE_OPENAI_ENDPOINT || "",
    apiKey: ENV.AZURE_OPENAI_API_KEY || "",
    deploymentName: defaultModelId,
    apiVersion: "2024-12-01-preview",
  },

  maxTokens: 16000,
  temperature: 0.7,

  systemPrompt: `Eres un asistente experto en Microsoft Excel integrado como complemento.
Tu rol es ayudar a los usuarios con:
- Análisis de datos y hojas de cálculo
- Generación de fórmulas de Excel
- Explicación de funciones y características de Excel
- Interpretación de datos proporcionados
- Sugerencias para mejorar la organización de datos

Cuando el usuario te proporcione datos de celdas o rangos de Excel:
1. Analiza el contexto y estructura de los datos
2. Proporciona respuestas claras y accionables
3. Cuando generes fórmulas, asegúrate de que sean compatibles con Excel
4. Explica brevemente qué hace cada fórmula que generes

Si el usuario pide una fórmula, devuélvela en un formato claro que pueda copiar directamente.
Responde siempre en español a menos que el usuario escriba en otro idioma.`,
};

/**
 * Obtiene la definición de un modelo por su ID
 */
export function getModelById(modelId: string): ModelDefinition | undefined {
  return availableModels.find((m) => m.id === modelId);
}

/**
 * Valida que la configuración tenga valores reales
 */
export function validateConfig(): { isValid: boolean; errors: string[] } {
  const errors: string[] = [];

  if (!config.azureOpenAI.endpoint) {
    errors.push("AZURE_OPENAI_ENDPOINT no está configurado en .env");
  }

  if (!config.azureOpenAI.apiKey) {
    errors.push("AZURE_OPENAI_API_KEY no está configurado en .env");
  }

  if (!config.azureOpenAI.deploymentName) {
    errors.push("El nombre del deployment/modelo no está configurado");
  }

  return {
    isValid: errors.length === 0,
    errors,
  };
}
