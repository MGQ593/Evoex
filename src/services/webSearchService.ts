/**
 * Servicio de búsqueda web usando SearXNG
 *
 * Utiliza una instancia self-hosted de SearXNG como metabuscador
 * para proporcionar resultados de búsqueda web al asistente AI.
 */

import { searxngProxyUrl } from "../config/config";

export interface SearchResult {
  title: string;
  url: string;
  content: string;
  engine: string;
  publishedDate?: string;
}

export interface SearchResponse {
  query: string;
  results: SearchResult[];
  suggestions: string[];
  infoboxes?: Array<{
    infobox: string;
    content: string;
    urls?: Array<{ title: string; url: string }>;
  }>;
}

interface SearXNGResponse {
  query: string;
  results: Array<{
    title: string;
    url: string;
    content: string;
    engine: string;
    publishedDate?: string;
    category?: string;
  }>;
  suggestions?: string[];
  infoboxes?: Array<{
    infobox: string;
    content: string;
    urls?: Array<{ title: string; url: string }>;
  }>;
}

/**
 * Configuración del servicio de búsqueda
 * Usa nuestro proxy propio que tiene CORS habilitado
 */
const SEARXNG_CONFIG = {
  // URL del proxy desde configuración
  baseUrl: searxngProxyUrl,
  defaultCategories: ["general"],
  defaultEngines: ["google", "bing", "duckduckgo"],
  maxResults: 5,
  language: "es",
  safesearch: 1,
};

/**
 * Realiza una búsqueda web usando SearXNG
 */
export async function searchWeb(
  query: string,
  options: {
    categories?: string[];
    engines?: string[];
    maxResults?: number;
    language?: string;
  } = {}
): Promise<SearchResponse> {
  const {
    categories = SEARXNG_CONFIG.defaultCategories,
    maxResults = SEARXNG_CONFIG.maxResults,
    language = SEARXNG_CONFIG.language,
  } = options;

  // Construir URL de búsqueda usando nuestro proxy
  const params = new URLSearchParams({
    q: query,
    format: "json",
    language: language,
    safesearch: SEARXNG_CONFIG.safesearch.toString(),
    categories: categories.join(","),
  });

  const url = `${SEARXNG_CONFIG.baseUrl}/search?${params.toString()}`;

  const response = await fetch(url, {
    method: "GET",
    headers: {
      Accept: "application/json",
    },
  });

  if (!response.ok) {
    throw new Error(`Error en búsqueda: ${response.status} ${response.statusText}`);
  }

  const data: SearXNGResponse = await response.json();

  // Limitar resultados
  const limitedResults = data.results.slice(0, maxResults).map((r) => ({
    title: r.title,
    url: r.url,
    content: r.content || "",
    engine: r.engine,
    publishedDate: r.publishedDate,
  }));

  return {
    query: data.query,
    results: limitedResults,
    suggestions: data.suggestions || [],
    infoboxes: data.infoboxes,
  };
}

/**
 * Formatea los resultados de búsqueda para incluirlos en el contexto del AI
 */
export function formatSearchResultsForContext(searchResponse: SearchResponse): string {
  if (!searchResponse.results.length) {
    return "No se encontraron resultados de búsqueda web.";
  }

  let context = `\n📡 **Resultados de búsqueda web para "${searchResponse.query}":**\n\n`;

  searchResponse.results.forEach((result, index) => {
    context += `**${index + 1}. ${result.title}**\n`;
    context += `   URL: ${result.url}\n`;
    if (result.content) {
      context += `   ${result.content.substring(0, 300)}${result.content.length > 300 ? "..." : ""}\n`;
    }
    if (result.publishedDate) {
      context += `   Fecha: ${result.publishedDate}\n`;
    }
    context += "\n";
  });

  // Agregar infoboxes si existen
  if (searchResponse.infoboxes && searchResponse.infoboxes.length > 0) {
    context += "\n📋 **Información destacada:**\n";
    searchResponse.infoboxes.forEach((box) => {
      context += `- ${box.infobox}: ${box.content.substring(0, 200)}...\n`;
    });
  }

  return context;
}

/**
 * Detecta si una consulta realmente necesita búsqueda web
 * Evita búsquedas innecesarias para saludos simples o comandos de Excel
 */
export function shouldSearchWeb(query: string): boolean {
  // Primero, descartar saludos y mensajes muy cortos
  const greetings = /^(hola|hi|hey|buenos días|buenas tardes|buenas noches|saludos|hello|gracias|ok|vale|sí|no|bien|perfecto)[\s!.?]*$/i;
  if (greetings.test(query.trim())) {
    return false;
  }

  // Descartar mensajes muy cortos (menos de 10 caracteres)
  if (query.trim().length < 10) {
    return false;
  }

  // Patrones que SÍ requieren búsqueda web
  const webSearchPatterns = [
    /\b(busca|buscar|búsqueda|search|googlea)\b/i,
    /\b(qué es|que es|what is|quién es|quien es)\b/i,
    /\b(cómo se hace|como se hace|how to)\b/i,
    /\b(últimas noticias|noticias de|news about|actualidad)\b/i,
    /\b(precio actual|cotización|stock price)\b/i,
    /\b(en internet|en la web|online|en google)\b/i,
    /\b(información sobre|info about|datos de)\b/i,
    /\b(2024|2025|2026)\b/, // Fechas actuales sugieren info actualizada
    /\bhttps?:\/\//i, // URLs directas
  ];

  return webSearchPatterns.some((pattern) => pattern.test(query));
}

/**
 * Detecta si una consulta debería ir a la Knowledge Base (RAG)
 * Es más permisivo que shouldSearchWeb - cualquier pregunta sustancial aplica
 */
export function shouldQueryRag(query: string): boolean {
  // Descartar saludos simples
  const greetings = /^(hola|hi|hey|buenos días|buenas tardes|buenas noches|saludos|hello|gracias|ok|vale|sí|no|bien|perfecto)[\s!.?]*$/i;
  if (greetings.test(query.trim())) {
    return false;
  }

  // Descartar mensajes muy cortos (menos de 5 caracteres)
  if (query.trim().length < 5) {
    return false;
  }

  // Comandos de Excel que no necesitan RAG
  const excelCommands = /^(suma|promedio|contar|ordenar|filtrar|crea|genera|escribe|pon|calcula)\s+(en|la|el|una|un)?\s*(celda|columna|fila|rango|tabla)/i;
  if (excelCommands.test(query.trim())) {
    return false;
  }

  // Todo lo demás puede beneficiarse del contexto de RAG
  return true;
}

/**
 * Verifica si el servicio de búsqueda está disponible
 * Usa el endpoint /health de nuestro proxy
 */
export async function checkSearXNGHealth(): Promise<boolean> {
  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 5000);

    const response = await fetch(`${SEARXNG_CONFIG.baseUrl}/health`, {
      method: "GET",
      headers: { Accept: "application/json" },
      signal: controller.signal,
    });

    clearTimeout(timeoutId);

    return response.ok;
  } catch {
    return false;
  }
}

// ===== Web Content Fetching =====

export interface WebTable {
  headers: string[];
  rows: string[][];
}

export interface DownloadLink {
  text: string;
  url: string;
  type: string;
}

export interface WebContentResponse {
  type: "html" | "json";
  url: string;
  title?: string;
  textContent?: string[];
  tables?: WebTable[];
  downloadLinks?: DownloadLink[];
  data?: unknown; // For JSON responses
  fetchedAt: string;
  error?: string;
}

/**
 * Obtiene y parsea el contenido de una URL
 * Extrae texto, tablas y enlaces de descarga
 * @param url URL a obtener
 * @param useJsRendering Si es true, usa Puppeteer para renderizar JavaScript (más lento pero funciona con sitios dinámicos)
 */
export async function fetchWebContent(url: string, useJsRendering: boolean = false): Promise<WebContentResponse> {
  try {
    const endpoint = useJsRendering ? "/fetch-js" : "/fetch";
    const fetchUrl = `${SEARXNG_CONFIG.baseUrl}${endpoint}?url=${encodeURIComponent(url)}`;

    const controller = new AbortController();
    // Timeout más largo para Puppeteer
    const timeout = useJsRendering ? 60000 : 30000;
    const timeoutId = setTimeout(() => controller.abort(), timeout);

    const response = await fetch(fetchUrl, {
      method: "GET",
      headers: { Accept: "application/json" },
      signal: controller.signal,
    });

    clearTimeout(timeoutId);

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    const data: WebContentResponse = await response.json();
    return data;
  } catch (error) {
    return {
      type: "html",
      url,
      fetchedAt: new Date().toISOString(),
      error: error instanceof Error ? error.message : "Error desconocido",
    };
  }
}

/**
 * Formatea el contenido web para incluirlo en el contexto del AI
 */
export function formatWebContentForContext(content: WebContentResponse): string {
  if (content.error) {
    return `⚠️ Error al obtener contenido de ${content.url}: ${content.error}`;
  }

  let context = `\n📄 **Contenido de "${content.title || content.url}":**\n\n`;

  // Agregar texto principal (limitado)
  if (content.textContent && content.textContent.length > 0) {
    context += "**Texto extraído:**\n";
    content.textContent.slice(0, 15).forEach((text) => {
      context += `• ${text.substring(0, 500)}${text.length > 500 ? "..." : ""}\n`;
    });
    context += "\n";
  }

  // Agregar tablas
  if (content.tables && content.tables.length > 0) {
    context += "**Tablas encontradas:**\n";
    content.tables.slice(0, 3).forEach((table, index) => {
      context += `\nTabla ${index + 1}:\n`;
      if (table.headers.length > 0) {
        context += `Columnas: ${table.headers.join(" | ")}\n`;
      }
      table.rows.slice(0, 10).forEach((row) => {
        context += `${row.join(" | ")}\n`;
      });
      if (table.rows.length > 10) {
        context += `... (${table.rows.length - 10} filas más)\n`;
      }
    });
    context += "\n";
  }

  // Agregar enlaces de descarga
  if (content.downloadLinks && content.downloadLinks.length > 0) {
    context += "**Archivos disponibles para descarga:**\n";
    content.downloadLinks.forEach((link) => {
      context += `• [${link.type.toUpperCase()}] ${link.text}: ${link.url}\n`;
    });
  }

  return context;
}

/**
 * Detecta si el mensaje contiene una URL para hacer fetch
 */
export function extractUrlFromMessage(message: string): string | null {
  // eslint-disable-next-line no-useless-escape
  const urlPattern = /https?:\/\/[^\s<>"{}|\\^`\[\]]+/gi;
  const matches = message.match(urlPattern);
  return matches ? matches[0] : null;
}

/**
 * Obtiene contenido web de forma inteligente
 * Primero intenta con fetch estático, si falla o no obtiene datos útiles,
 * reintenta con Puppeteer (JavaScript rendering)
 */
export async function fetchWebContentSmart(url: string): Promise<WebContentResponse> {
  // Primero intentar con fetch estático (más rápido)
  const staticResult = await fetchWebContent(url, false);

  // Si hubo error o no se obtuvo contenido útil, intentar con Puppeteer
  const hasUsefulContent =
    !staticResult.error &&
    ((staticResult.textContent && staticResult.textContent.length > 3) ||
      (staticResult.tables && staticResult.tables.length > 0));

  if (!hasUsefulContent) {
    const jsResult = await fetchWebContent(url, true);

    // Retornar el resultado de Puppeteer si tiene más contenido
    if (!jsResult.error && (jsResult.textContent?.length || 0) > (staticResult.textContent?.length || 0)) {
      return jsResult;
    }
  }

  return staticResult;
}

// ===== Open WebUI RAG Integration =====

export interface RAGFile {
  id: string;
  filename: string;
  meta?: Record<string, unknown>;
  created_at?: string;
}

export interface RAGUploadResponse {
  success: boolean;
  message: string;
  file?: RAGFile;
  originalUrl: string;
  error?: string;
}

export interface RAGChunk {
  content: string;
  metadata?: Record<string, unknown>;
  score?: number | null;
  source: string;
  filename?: string;
  fileId?: string;
  index?: number;
}

export interface RAGRetrieveResponse {
  success: boolean;
  method?: "vector_retrieval" | "file_content";
  query?: string;
  chunks?: RAGChunk[];
  knowledgeBaseId?: string;
  knowledgeBaseName?: string;
  totalChunks?: number;
  error?: string;
}

/**
 * Sube un PDF a Open WebUI RAG
 * @param pdfUrl URL del PDF a descargar y subir
 * @param apiKey API key de Open WebUI (opcional - el proxy usa la configurada en el servidor)
 * @param filename Nombre opcional para el archivo
 */
export async function uploadPdfToRag(
  pdfUrl: string,
  apiKey?: string,
  filename?: string
): Promise<RAGUploadResponse> {
  try {
    // El proxy usa la API key de sus variables de entorno si no se proporciona una
    const body: Record<string, string> = { pdfUrl };
    if (apiKey) body.apiKey = apiKey;
    if (filename) body.filename = filename;

    const response = await fetch(`${SEARXNG_CONFIG.baseUrl}/upload-to-rag`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.message || `HTTP ${response.status}`);
    }

    const result: RAGUploadResponse = await response.json();
    return result;
  } catch (error) {
    return {
      success: false,
      message: "Error subiendo PDF a RAG",
      originalUrl: pdfUrl,
      error: error instanceof Error ? error.message : "Error desconocido",
    };
  }
}

/**
 * Detecta si hay enlaces PDF en los downloadLinks y ofrece subirlos a RAG
 */
export function getPdfLinksFromContent(content: WebContentResponse): DownloadLink[] {
  if (!content.downloadLinks) return [];
  return content.downloadLinks.filter((link) => link.type.toLowerCase() === "pdf");
}

/**
 * Obtiene fragmentos relevantes de la Knowledge Base sin generación
 * Esto permite usar los fragmentos con el modelo principal (Azure OpenAI) del add-in
 * @param query Pregunta o consulta
 * @param topK Número máximo de fragmentos a retornar (default: 8)
 */
export async function retrieveFromRag(query: string, topK: number = 8): Promise<RAGRetrieveResponse> {
  try {
    const response = await fetch(`${SEARXNG_CONFIG.baseUrl}/retrieve-only`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ query, topK }),
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.message || `HTTP ${response.status}`);
    }

    const result: RAGRetrieveResponse = await response.json();
    return result;
  } catch (error) {
    return {
      success: false,
      error: error instanceof Error ? error.message : "Error desconocido",
    };
  }
}

/**
 * Formatea los fragmentos de RAG para incluirlos en el contexto del AI
 */
export function formatRagChunksForContext(retrieveResponse: RAGRetrieveResponse): string {
  if (!retrieveResponse.success || !retrieveResponse.chunks || retrieveResponse.chunks.length === 0) {
    return "";
  }

  let context = `\n📚 **Información de Knowledge Base${retrieveResponse.knowledgeBaseName ? ` (${retrieveResponse.knowledgeBaseName})` : ""}:**\n\n`;

  retrieveResponse.chunks.forEach((chunk, index) => {
    const source = chunk.filename || chunk.metadata?.source || `Fragmento ${index + 1}`;
    context += `**[${source}]**\n`;
    context += `${chunk.content.substring(0, 800)}${chunk.content.length > 800 ? "..." : ""}\n\n`;
  });

  return context;
}
