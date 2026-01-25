/**
 * Excel AI Assistant - Lógica principal (Simplificada)
 */

import "./taskpane.css";
import { azureOpenAIService, AzureOpenAIError, StructuredResponse } from "../services/azureOpenAIService";
import {
  excelService,
  ExcelServiceError,
  SelectionInfo,
  ExcelAction,
  ActionResult
} from "../services/excelService";
import { validateConfig, availableModels, getModelById, defaultModelId, searxngProxyUrl } from "../config/config";
import {
  searchWeb,
  formatSearchResultsForContext,
  checkSearXNGHealth,
  shouldSearchWeb,
  shouldQueryRag,
  fetchWebContentSmart,
  formatWebContentForContext,
  extractUrlFromMessage,
  getPdfLinksFromContent,
  uploadPdfToRag,
  retrieveFromRag,
  formatRagChunksForContext,
  DownloadLink,
  WebContentResponse
} from "../services/webSearchService";
import { getUserInfo, getGreeting, UserInfo } from "../services/userService";

/* global Office, document */

// ===== Estado =====

type EditMode = "auto" | "ask";

type FileAction = "context" | "import" | "rag" | "extract" | "extract-multi";

interface AttachedFile {
  name: string;
  type: string;
  size: number;
  content?: string; // Texto extraído del archivo
  base64?: string;  // Contenido en base64 para enviar al proxy
  processing?: boolean;
  error?: string;
  action?: FileAction; // Qué hacer con el archivo
  sheets?: Record<string, string>; // Para archivos Excel - datos por hoja
  sheetNames?: string[]; // Nombres de las hojas
}

interface AppState {
  currentSelection: SelectionInfo | null;
  currentExcelContext: string | null;
  isProcessing: boolean;
  pendingActions: ExcelAction[];
  lastActionResults: ActionResult[];
  editMode: EditMode;
  lastUserMessage: string;
  webSearchEnabled: boolean;
  ragEnabled: boolean;
  attachedFiles: AttachedFile[];
  currentUser: UserInfo | null;
}

const state: AppState = {
  currentSelection: null,
  currentExcelContext: null,
  isProcessing: false,
  pendingActions: [],
  lastActionResults: [],
  editMode: (localStorage.getItem("editMode") as EditMode) || "auto",
  lastUserMessage: "",
  webSearchEnabled: localStorage.getItem("webSearchEnabled") === "true",
  ragEnabled: localStorage.getItem("ragEnabled") === "true",
  attachedFiles: [],
  currentUser: null
};

// URL del proxy para procesamiento de archivos
const PROXY_URL = searxngProxyUrl;

/**
 * Detecta si el usuario pregunta específicamente por los datos seleccionados
 */
function asksForSelectedData(message: string): boolean {
  const patterns = [
    /\b(datos?|valores?|información|contenido|celdas?)\s+(seleccionad[oa]s?|marcad[oa]s?|elegid[oa]s?)\b/i,
    /\b(selección|lo seleccionado|que seleccioné|que marqué)\b/i,
    /\b(dame|muéstrame|dime|qué hay en)\s+(la|el|los|las)?\s*(selección|seleccionado)\b/i,
    /\bselected\s+(data|cells?|range)\b/i,
    /\b(analiza|procesa|usa|trabaja con)\s+(la|el|los|las)?\s*(selección|datos seleccionados)\b/i,
  ];
  return patterns.some(p => p.test(message));
}

// ===== Elementos del DOM =====

const el = {
  messagesContainer: () => document.getElementById("messagesContainer") as HTMLElement,
  userInput: () => document.getElementById("userInput") as HTMLTextAreaElement,
  sendBtn: () => document.getElementById("sendBtn") as HTMLButtonElement,
  clearHistoryBtn: () => document.getElementById("clearHistoryBtn") as HTMLButtonElement,
  toast: () => document.getElementById("toast") as HTMLElement,
  modelSelector: () => document.getElementById("modelSelector") as HTMLSelectElement,
  selectionText: () => document.getElementById("selectionText") as HTMLElement,
  pendingActions: () => document.getElementById("pendingActions") as HTMLElement,
  pendingCount: () => document.getElementById("pendingCount") as HTMLElement,
  acceptAllBtn: () => document.getElementById("acceptAllBtn") as HTMLButtonElement,
  editModeToggle: () => document.getElementById("editModeToggle") as HTMLElement,
  editModeIcon: () => document.getElementById("editModeIcon") as HTMLElement,
  editModeText: () => document.getElementById("editModeText") as HTMLElement,
  moreOptionsBtn: () => document.getElementById("moreOptionsBtn") as HTMLButtonElement,
  buttonsLeft: () => document.querySelector(".buttons-left") as HTMLElement,
  webActiveIndicator: () => document.getElementById("webActiveIndicator") as HTMLElement,
  ragActiveIndicator: () => document.getElementById("ragActiveIndicator") as HTMLElement,
  fileInput: () => document.getElementById("fileInput") as HTMLInputElement,
  attachedFilesContainer: () => document.getElementById("attachedFilesContainer") as HTMLElement,
};

// Referencia al popup menu (se crea dinámicamente)
let optionsPopup: HTMLElement | null = null;

// ===== Toast =====

function showToast(message: string, type: "success" | "error" | "info" = "info"): void {
  const toast = el.toast();
  const msgEl = toast.querySelector(".toast-msg") as HTMLElement;

  toast.classList.remove("success", "error", "info");
  toast.classList.add(type);
  msgEl.textContent = message;

  toast.classList.remove("hidden");
  toast.classList.add("visible");

  setTimeout(() => {
    toast.classList.remove("visible");
    setTimeout(() => toast.classList.add("hidden"), 200);
  }, 2500);
}

// ===== Loading =====

let loadingMessageElement: HTMLElement | null = null;

function setLoading(loading: boolean): void {
  state.isProcessing = loading;
  const container = el.messagesContainer();

  if (loading) {
    // Crear mensaje de carga en el chat
    if (!loadingMessageElement) {
      loadingMessageElement = document.createElement("div");
      loadingMessageElement.className = "message assistant loading-message";
      loadingMessageElement.innerHTML = `
        <div class="loading-dots">
          <div class="loading-dot"></div>
          <div class="loading-dot"></div>
          <div class="loading-dot"></div>
        </div>
      `;
      container.appendChild(loadingMessageElement);
      container.scrollTop = container.scrollHeight;
    }
  } else {
    // Remover mensaje de carga
    if (loadingMessageElement) {
      loadingMessageElement.remove();
      loadingMessageElement = null;
    }
  }

  updateInputState();
  updateAcceptButtonState();
}

// ===== Input State =====

function updateInputState(): void {
  const input = el.userInput();
  const sendBtn = el.sendBtn();
  const hasText = input.value.trim().length > 0;

  sendBtn.disabled = !hasText || state.isProcessing;
  el.modelSelector().disabled = state.isProcessing;
}

function autoResizeTextarea(): void {
  const input = el.userInput();
  input.style.height = "auto";
  input.style.height = Math.min(input.scrollHeight, 100) + "px";
}

// ===== Selection =====

function updateSelection(selection: SelectionInfo): void {
  state.currentSelection = selection;
  const text = el.selectionText();

  if (selection.isSingleCell) {
    text.textContent = `${selection.address} selected`;
  } else {
    const cells = selection.rowCount * selection.columnCount;
    text.textContent = `${selection.address} (${cells} cells)`;
  }
}

// ===== Model Selector =====

function initializeModelSelector(): void {
  const selector = el.modelSelector();
  selector.innerHTML = "";

  availableModels.forEach((model) => {
    const option = document.createElement("option");
    option.value = model.id;
    option.textContent = model.name;
    if (model.id === defaultModelId) {
      option.selected = true;
    }
    selector.appendChild(option);
  });
}

function handleModelChange(): void {
  const selector = el.modelSelector();
  const success = azureOpenAIService.setModel(selector.value);

  if (success) {
    showToast(`Modelo: ${azureOpenAIService.getCurrentModelName()}`, "info");
  }
}

// ===== Edit Mode =====

function updateEditModeUI(): void {
  const toggle = el.editModeToggle();
  const icon = el.editModeIcon();
  const text = el.editModeText();

  if (state.editMode === "auto") {
    toggle.classList.remove("ask-mode");
    toggle.classList.add("auto-mode");
    icon.className = "ms-Icon ms-Icon--CheckMark";
    text.textContent = "Auto";
    toggle.title = "Ejecutar automáticamente";
  } else {
    toggle.classList.remove("auto-mode");
    toggle.classList.add("ask-mode");
    icon.className = "ms-Icon ms-Icon--EditMirrored";
    text.textContent = "Confirmar";
    toggle.title = "Pedir confirmación antes de editar";
  }
}

function toggleEditMode(): void {
  state.editMode = state.editMode === "auto" ? "ask" : "auto";
  localStorage.setItem("editMode", state.editMode);
  updateEditModeUI();

  const modeText = state.editMode === "auto" ? "Ejecución automática" : "Pedir confirmación";
  showToast(modeText, "info");
}

// ===== Options Popup Menu =====

function createOptionsPopup(): HTMLElement {
  const popup = document.createElement("div");
  popup.className = "options-popup";
  popup.innerHTML = `
    <div class="options-popup-item" id="attachOption">
      <i class="ms-Icon ms-Icon--Attach" aria-hidden="true"></i>
      <span>Adjuntar archivos</span>
    </div>
    <div class="options-popup-item" id="webSearchOption">
      <i class="ms-Icon ms-Icon--Globe" aria-hidden="true"></i>
      <span>Buscar en la web</span>
      <i class="ms-Icon ms-Icon--CheckMark check-icon" style="display: none;" aria-hidden="true"></i>
    </div>
    <div class="options-popup-item" id="ragOption">
      <i class="ms-Icon ms-Icon--Database" aria-hidden="true"></i>
      <span>RAG</span>
      <i class="ms-Icon ms-Icon--CheckMark check-icon" style="display: none;" aria-hidden="true"></i>
    </div>
    <div class="options-popup-item" id="refreshIndexOption">
      <i class="ms-Icon ms-Icon--Sync" aria-hidden="true"></i>
      <span>Actualizar índice</span>
    </div>
  `;
  return popup;
}

function showOptionsPopup(): void {
  if (optionsPopup) {
    hideOptionsPopup();
    return;
  }

  const moreBtn = el.moreOptionsBtn();
  if (!moreBtn) return;

  // Crear popup
  optionsPopup = createOptionsPopup();
  
  // Añadir al body para que no se corte
  document.body.appendChild(optionsPopup);
  
  // Calcular posición basada en el botón
  const btnRect = moreBtn.getBoundingClientRect();
  const popupHeight = 150; // altura aproximada del popup
  
  optionsPopup.style.position = "fixed";
  optionsPopup.style.left = `${btnRect.left}px`;
  optionsPopup.style.bottom = `${window.innerHeight - btnRect.top + 8}px`;

  // Actualizar estado visual del item de web search
  updatePopupWebSearchState();
  updatePopupRagState();

  // Event listeners para las opciones
  const attachOption = optionsPopup.querySelector("#attachOption");
  const webSearchOption = optionsPopup.querySelector("#webSearchOption");
  const ragOption = optionsPopup.querySelector("#ragOption");
  const refreshIndexOption = optionsPopup.querySelector("#refreshIndexOption");

  attachOption?.addEventListener("click", (e) => {
    e.stopPropagation();
    hideOptionsPopup();
    openFileDialog();
  });

  webSearchOption?.addEventListener("click", async (e) => {
    e.stopPropagation();
    await toggleWebSearch();
    updatePopupWebSearchState();
  });

  ragOption?.addEventListener("click", async (e) => {
    e.stopPropagation();
    await toggleRag();
    updatePopupRagState();
  });

  refreshIndexOption?.addEventListener("click", async (e) => {
    e.stopPropagation();
    hideOptionsPopup();
    await refreshDataIndex(true);
  });

  // Cerrar al hacer click fuera
  setTimeout(() => {
    document.addEventListener("click", handleOutsideClick);
  }, 10);
}

function hideOptionsPopup(): void {
  if (optionsPopup) {
    optionsPopup.remove();
    optionsPopup = null;
  }
  document.removeEventListener("click", handleOutsideClick);
}

function handleOutsideClick(e: MouseEvent): void {
  const moreBtn = el.moreOptionsBtn();
  if (optionsPopup && !optionsPopup.contains(e.target as Node) && e.target !== moreBtn) {
    hideOptionsPopup();
  }
}

function updatePopupWebSearchState(): void {
  if (!optionsPopup) return;
  
  const webSearchOption = optionsPopup.querySelector("#webSearchOption");
  const checkIcon = webSearchOption?.querySelector(".check-icon") as HTMLElement;
  
  if (webSearchOption && checkIcon) {
    if (state.webSearchEnabled) {
      webSearchOption.classList.add("active");
      checkIcon.style.display = "block";
    } else {
      webSearchOption.classList.remove("active");
      checkIcon.style.display = "none";
    }
  }
}

function updatePopupRagState(): void {
  if (!optionsPopup) return;
  
  const ragOption = optionsPopup.querySelector("#ragOption");
  const checkIcon = ragOption?.querySelector(".check-icon") as HTMLElement;
  
  if (ragOption && checkIcon) {
    if (state.ragEnabled) {
      ragOption.classList.add("active");
      checkIcon.style.display = "block";
    } else {
      ragOption.classList.remove("active");
      checkIcon.style.display = "none";
    }
  }
}

// ===== Web Search Toggle =====

function updateWebSearchUI(): void {
  // Actualizar popup si está abierto
  updatePopupWebSearchState();
  
  // Actualizar indicador externo
  const indicator = el.webActiveIndicator();
  if (indicator) {
    indicator.style.display = state.webSearchEnabled ? "flex" : "none";
  }
}

async function toggleWebSearch(): Promise<void> {
  // Si estamos activando, verificar que el servicio esté disponible
  if (!state.webSearchEnabled) {
    const isAvailable = await checkSearXNGHealth();
    if (!isAvailable) {
      showToast("Servicio de búsqueda no disponible", "error");
      return;
    }
  }

  state.webSearchEnabled = !state.webSearchEnabled;
  localStorage.setItem("webSearchEnabled", state.webSearchEnabled.toString());
  updateWebSearchUI();
}

// ===== RAG Toggle =====

function updateRagUI(): void {
  // Actualizar popup si está abierto
  updatePopupRagState();
  
  // Actualizar indicador externo
  const indicator = el.ragActiveIndicator();
  if (indicator) {
    indicator.style.display = state.ragEnabled ? "flex" : "none";
  }
}

async function toggleRag(): Promise<void> {
  // Si estamos activando, verificar que el servicio esté disponible
  if (!state.ragEnabled) {
    const isAvailable = await checkSearXNGHealth();
    if (!isAvailable) {
      showToast("Servicio RAG no disponible", "error");
      return;
    }
  }

  state.ragEnabled = !state.ragEnabled;
  localStorage.setItem("ragEnabled", state.ragEnabled.toString());
  updateRagUI();
}

// ===== File Attachment =====

/**
 * Tipos de archivo soportados
 */
const SUPPORTED_FILE_TYPES = [
  ".pdf",
  ".docx",
  ".xlsx",
  ".xls",
  ".csv",
  ".txt",
  ".md",
  ".json",
  ".xml"
];

const SUPPORTED_MIME_TYPES = [
  "application/pdf",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-excel",
  "text/csv",
  "text/plain",
  "text/markdown",
  "application/json",
  "application/xml",
  "text/xml"
];

/**
 * Crea el input file oculto
 */
function createFileInput(): void {
  // Verificar si ya existe
  if (document.getElementById("fileInput")) return;

  const input = document.createElement("input");
  input.type = "file";
  input.id = "fileInput";
  input.multiple = true;
  input.accept = SUPPORTED_FILE_TYPES.join(",");
  input.style.display = "none";
  input.addEventListener("change", handleFileSelection);
  document.body.appendChild(input);
}

/**
 * Abre el diálogo de selección de archivos
 */
function openFileDialog(): void {
  const input = el.fileInput();
  if (input) {
    input.click();
  }
}

/**
 * Maneja la selección de archivos
 */
async function handleFileSelection(e: Event): Promise<void> {
  const input = e.target as HTMLInputElement;
  const files = input.files;

  if (!files || files.length === 0) return;

  // Si hay múltiples archivos, mostrar diálogo de acción masiva
  if (files.length > 1) {
    const result = await showMultiFileActionDialog(Array.from(files));
    if (result) {
      const fileList = Array.from(files);

      // Si es extracción múltiple, procesar todos y luego enviar prompt consolidado
      if (result.action === "extract-multi") {
        showToast(`Procesando ${fileList.length} archivos...`, "info");

        // Procesar todos como contexto primero
        for (const file of fileList) {
          await processAttachedFile(file, "context");
        }

        // Esperar un momento para asegurar que todos están procesados
        await new Promise(resolve => setTimeout(resolve, 500));

        // Enviar prompt automático para extraer datos consolidados (silencioso)
        const extractPrompt = `Extrae los datos de los ${fileList.length} documentos adjuntos a una tabla de Excel.

INSTRUCCIONES OBLIGATORIAS:
1. FILA 1 = ENCABEZADOS: Usa los nombres EXACTOS de los campos del documento (ej: "RUC", "Razón Social", "Fecha de Emisión", "Subtotal 12%")
2. FILAS 2+: Una fila por cada documento con sus datos
3. Los RUC y números de factura deben ir como TEXTO (no números) para evitar notación científica
4. Incluye TODOS los campos: emisor, receptor, fechas, números de documento, conceptos, subtotales, impuestos, totales
5. Ajusta el ancho de las columnas al contenido`;

        // Enviar silenciosamente con mensaje simplificado
        const displayMsg = `📄 Extrayendo datos de ${fileList.length} documentos...`;
        setTimeout(() => {
          sendMessageSilent(extractPrompt, displayMsg);
        }, 300);
      } else {
        // Procesar normalmente con la acción seleccionada
        for (const file of fileList) {
          await processAttachedFile(file, result.action as FileAction);
        }
      }
    }
  } else {
    // Un solo archivo - mostrar diálogo individual
    const file = files[0];
    const action = await showFileActionDialog(file);
    if (action) {
      await processAttachedFile(file, action);
    }
  }

  // Limpiar el input para permitir seleccionar el mismo archivo de nuevo
  input.value = "";
}

/**
 * Muestra diálogo para múltiples archivos
 */
function showMultiFileActionDialog(files: File[]): Promise<{ action: FileAction } | null> {
  return new Promise((resolve) => {
    const overlay = document.createElement("div");
    overlay.className = "file-action-overlay";

    const totalSize = files.reduce((sum, f) => sum + f.size, 0);
    const fileTypes = [...new Set(files.map(f => f.name.split(".").pop()?.toUpperCase()))].join(", ");

    const dialog = document.createElement("div");
    dialog.className = "file-action-dialog";
    dialog.innerHTML = `
      <div class="file-action-header">
        <i class="ms-Icon ms-Icon--Documentation" aria-hidden="true"></i>
        <div class="file-action-title">
          <span class="file-action-name">${files.length} archivos seleccionados</span>
          <span class="file-action-size">${fileTypes} - ${formatFileSize(totalSize)} total</span>
        </div>
      </div>
      <div class="multi-file-list">
        ${files.slice(0, 5).map(f => `
          <div class="multi-file-item">
            <i class="ms-Icon ${getFileIcon(f.type || f.name.split(".").pop() || "")}" aria-hidden="true"></i>
            <span>${truncateFilename(f.name, 25)}</span>
          </div>
        `).join("")}
        ${files.length > 5 ? `<div class="multi-file-more">...y ${files.length - 5} más</div>` : ""}
      </div>
      <p class="file-action-question">¿Qué deseas hacer con todos los archivos?</p>
      <div class="file-action-options">
        <button class="file-action-btn primary" data-action="extract-multi">
          <i class="ms-Icon ms-Icon--ExcelDocument" aria-hidden="true"></i>
          <div class="file-action-btn-text">
            <span class="file-action-btn-title">Extraer datos a Excel</span>
            <span class="file-action-btn-desc">Crea una tabla consolidada con los datos de todos los archivos</span>
          </div>
        </button>
        <button class="file-action-btn" data-action="context">
          <i class="ms-Icon ms-Icon--ChatBot" aria-hidden="true"></i>
          <div class="file-action-btn-text">
            <span class="file-action-btn-title">Usar como contexto</span>
            <span class="file-action-btn-desc">Adjuntar al chat para consultar manualmente</span>
          </div>
        </button>
        <button class="file-action-btn" data-action="rag">
          <i class="ms-Icon ms-Icon--CloudUpload" aria-hidden="true"></i>
          <div class="file-action-btn-text">
            <span class="file-action-btn-title">Subir a Knowledge Base</span>
            <span class="file-action-btn-desc">Para consultar en futuras conversaciones</span>
          </div>
        </button>
      </div>
      <button class="file-action-cancel">Cancelar</button>
    `;

    overlay.appendChild(dialog);
    document.body.appendChild(overlay);

    // Event listeners
    const buttons = dialog.querySelectorAll(".file-action-btn");
    buttons.forEach(btn => {
      btn.addEventListener("click", () => {
        const action = (btn as HTMLElement).dataset.action as FileAction;
        overlay.remove();
        resolve({ action });
      });
    });

    const cancelBtn = dialog.querySelector(".file-action-cancel");
    cancelBtn?.addEventListener("click", () => {
      overlay.remove();
      resolve(null);
    });

    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) {
        overlay.remove();
        resolve(null);
      }
    });
  });
}

/**
 * Muestra diálogo preguntando qué hacer con el archivo
 */
function showFileActionDialog(file: File): Promise<FileAction | null> {
  return new Promise((resolve) => {
    // Crear overlay
    const overlay = document.createElement("div");
    overlay.className = "file-action-overlay";

    // Determinar si el archivo puede importarse a Excel (tipo compatible y <20MB)
    const ext = file.name.split(".").pop()?.toLowerCase();
    const MAX_IMPORT_SIZE = 20 * 1024 * 1024; // 20MB
    const canImport = ["xlsx", "xls", "csv", "txt"].includes(ext || "") && file.size <= MAX_IMPORT_SIZE;

    // Archivos grandes (>500KB) no deberían usarse como contexto (consume muchos tokens)
    const MAX_CONTEXT_SIZE = 500 * 1024; // 500KB
    const isLargeFile = file.size > MAX_CONTEXT_SIZE;

    // Archivos muy grandes (>20MB) solo pueden ir a RAG
    const isVeryLargeFile = file.size > MAX_IMPORT_SIZE;

    const dialog = document.createElement("div");
    dialog.className = "file-action-dialog";
    dialog.innerHTML = `
      <div class="file-action-header">
        <i class="ms-Icon ${getFileIcon(file.type || ext || "")}" aria-hidden="true"></i>
        <div class="file-action-title">
          <span class="file-action-name">${truncateFilename(file.name, 30)}</span>
          <span class="file-action-size">${formatFileSize(file.size)}</span>
        </div>
      </div>
      ${isVeryLargeFile ? `
      <p class="file-action-warning">
        <i class="ms-Icon ms-Icon--Warning" aria-hidden="true"></i>
        Archivo muy grande (${formatFileSize(file.size)}): solo se puede subir a Knowledge Base
      </p>
      ` : isLargeFile ? `
      <p class="file-action-warning">
        <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i>
        Archivo grande: se recomienda subirlo a Knowledge Base
      </p>
      ` : ""}
      <p class="file-action-question">¿Qué deseas hacer con este archivo?</p>
      <div class="file-action-options">
        ${!isLargeFile ? `
        <button class="file-action-btn" data-action="context">
          <i class="ms-Icon ms-Icon--ChatBot" aria-hidden="true"></i>
          <div class="file-action-btn-text">
            <span class="file-action-btn-title">Usar como contexto</span>
            <span class="file-action-btn-desc">El modelo analizará el contenido para responder preguntas</span>
          </div>
        </button>
        ` : ""}
        ${canImport ? `
        <button class="file-action-btn" data-action="import">
          <i class="ms-Icon ms-Icon--ExcelDocument" aria-hidden="true"></i>
          <div class="file-action-btn-text">
            <span class="file-action-btn-title">Importar a Excel</span>
            <span class="file-action-btn-desc">Copiar los datos a una nueva hoja de cálculo</span>
          </div>
        </button>
        ` : ""}
        <button class="file-action-btn primary" data-action="extract">
          <i class="ms-Icon ms-Icon--Robot" aria-hidden="true"></i>
          <div class="file-action-btn-text">
            <span class="file-action-btn-title">Extraer datos con IA</span>
            <span class="file-action-btn-desc">Sube al RAG y usa IA para extraer datos a Excel automáticamente</span>
          </div>
        </button>
        <button class="file-action-btn" data-action="rag">
          <i class="ms-Icon ms-Icon--Database" aria-hidden="true"></i>
          <div class="file-action-btn-text">
            <span class="file-action-btn-title">Subir a Knowledge Base${isVeryLargeFile ? " (única opción)" : ""}</span>
            <span class="file-action-btn-desc">Guardar en el RAG para consultas futuras${isLargeFile && !isVeryLargeFile ? " (recomendado)" : ""}</span>
          </div>
        </button>
      </div>
      <button class="file-action-cancel">Cancelar</button>
    `;

    overlay.appendChild(dialog);
    document.body.appendChild(overlay);

    // Event listeners
    const buttons = dialog.querySelectorAll(".file-action-btn");
    buttons.forEach(btn => {
      btn.addEventListener("click", () => {
        const action = (btn as HTMLElement).dataset.action as FileAction;
        overlay.remove();
        resolve(action);
      });
    });

    const cancelBtn = dialog.querySelector(".file-action-cancel");
    cancelBtn?.addEventListener("click", () => {
      overlay.remove();
      resolve(null);
    });

    // Cerrar al hacer clic fuera
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) {
        overlay.remove();
        resolve(null);
      }
    });
  });
}

/**
 * Procesa un archivo adjunto según la acción seleccionada
 */
async function processAttachedFile(file: File, action: FileAction): Promise<void> {
  // Límites de tamaño según la acción
  const SIZE_LIMITS: Record<FileAction, number> = {
    context: 500 * 1024,           // 500KB para contexto (consume tokens)
    import: 20 * 1024 * 1024,      // 20MB para importar a Excel
    rag: 50 * 1024 * 1024,         // 50MB para RAG (el servidor lo soporta)
    extract: 50 * 1024 * 1024,     // 50MB para extracción con IA (usa RAG)
    "extract-multi": 500 * 1024    // 500KB por archivo para extracción múltiple (usa contexto)
  };

  const maxSize = SIZE_LIMITS[action] || SIZE_LIMITS.import;
  const maxSizeLabel = (action === "rag" || action === "extract") ? "50MB" : action === "context" ? "500KB" : "20MB";

  if (file.size > maxSize) {
    showToast(`Archivo muy grande para ${action === "rag" ? "RAG" : action === "context" ? "contexto" : "importar"}: ${file.name} (max ${maxSizeLabel})`, "error");
    return;
  }

  // Verificar tipo de archivo
  const ext = "." + file.name.split(".").pop()?.toLowerCase();
  if (!SUPPORTED_FILE_TYPES.includes(ext)) {
    showToast(`Tipo no soportado: ${ext}`, "error");
    return;
  }

  // Para importar a Excel, procesar directamente sin agregar al estado
  if (action === "import") {
    await importFileToExcel(file);
    return;
  }

  // Agregar archivo al estado con estado de procesando
  const attachedFile: AttachedFile = {
    name: file.name,
    type: file.type || ext,
    size: file.size,
    processing: true,
    action: action
  };

  state.attachedFiles.push(attachedFile);
  renderAttachedFiles();

  try {
    // Leer archivo como base64
    const base64 = await readFileAsBase64(file);
    attachedFile.base64 = base64;

    // Enviar al proxy para extraer texto
    const response = await fetch(`${PROXY_URL}/upload-file`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        file: base64,
        filename: file.name,
        mimeType: file.type,
        uploadToRag: action === "rag" || action === "extract" // Subir a RAG si es rag o extract
      })
    });

    const result = await response.json();

    if (!response.ok) {
      throw new Error(result.error || `Error del servidor: ${response.status}`);
    }

    if (result.success && result.text) {
      attachedFile.content = result.text;
      attachedFile.sheets = result.sheets || undefined;
      attachedFile.sheetNames = result.sheetNames || undefined;
      attachedFile.processing = false;

      if ((action === "rag" || action === "extract") && result.uploadedToRag) {
        showToast(`📚 ${file.name} subido a Knowledge Base`, "success");
        // Quitar del estado ya que fue a RAG
        const idx = state.attachedFiles.indexOf(attachedFile);
        if (idx >= 0) state.attachedFiles.splice(idx, 1);

        // Si es "extract", enviar automáticamente el prompt para extraer datos
        if (action === "extract") {
          const fileNameWithoutExt = file.name.replace(/\.[^/.]+$/, "");
          const extractPrompt = `Extrae los datos del archivo "${file.name}" en una nueva hoja llamada "${fileNameWithoutExt}".

INSTRUCCIONES:
1. FILA 1 = ENCABEZADOS con los nombres EXACTOS de los campos del documento
2. FILA 2 = Los datos extraídos
3. RUC y números de factura como TEXTO (no números)
4. Incluye TODOS los campos: emisor, receptor, fechas, números, conceptos, subtotales, IVA, totales
5. Ajusta el ancho de columnas`;

          // Enviar silenciosamente
          const displayMsg = `📄 Extrayendo datos de ${file.name}...`;
          setTimeout(() => sendMessageSilent(extractPrompt, displayMsg), 500);
        }
      }
    } else if (result.success && result.uploadedToRag) {
      attachedFile.processing = false;
      showToast(`📚 ${file.name} subido a Knowledge Base`, "success");
      // Quitar del estado
      const idx = state.attachedFiles.indexOf(attachedFile);
      if (idx >= 0) state.attachedFiles.splice(idx, 1);

      // Si es "extract", enviar automáticamente el prompt para extraer datos
      if (action === "extract") {
        const fileNameWithoutExt = file.name.replace(/\.[^/.]+$/, "");
        const extractPrompt = `Extrae los datos del archivo "${file.name}" en una nueva hoja llamada "${fileNameWithoutExt}".

INSTRUCCIONES:
1. FILA 1 = ENCABEZADOS con los nombres EXACTOS de los campos del documento
2. FILA 2 = Los datos extraídos
3. RUC y números de factura como TEXTO (no números)
4. Incluye TODOS los campos: emisor, receptor, fechas, números, conceptos, subtotales, IVA, totales
5. Ajusta el ancho de columnas`;

        // Enviar silenciosamente
        const displayMsg = `📄 Extrayendo datos de ${file.name}...`;
        setTimeout(() => sendMessageSilent(extractPrompt, displayMsg), 500);
      }
    } else {
      throw new Error(result.error || "No se pudo extraer texto");
    }
  } catch (error) {
    console.error(`Error procesando ${file.name}:`, error);
    attachedFile.processing = false;
    attachedFile.error = error instanceof Error ? error.message : "Error desconocido";
    showToast(`Error: ${attachedFile.error}`, "error");
  }

  renderAttachedFiles();
}

/**
 * Importa un archivo directamente a Excel
 */
async function importFileToExcel(file: File): Promise<void> {
  showToast(`Importando ${file.name}...`, "info");

  try {
    // Leer archivo como base64
    const base64 = await readFileAsBase64(file);

    // Enviar al proxy para extraer datos
    const response = await fetch(`${PROXY_URL}/upload-file`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        file: base64,
        filename: file.name,
        mimeType: file.type,
        uploadToRag: false
      })
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      throw new Error(result.error || `Error del servidor: ${response.status}`);
    }

    const ext = file.name.split(".").pop()?.toLowerCase();

    // Si es Excel con múltiples hojas, preguntar cuál importar
    if (result.sheetNames && result.sheetNames.length > 1) {
      const sheetName = await selectSheetToImport(result.sheetNames);
      if (!sheetName) {
        showToast("Importación cancelada", "info");
        return;
      }
      await importCsvDataToExcel(result.sheets[sheetName], `${file.name} - ${sheetName}`);
    } else if (result.sheets && result.sheetNames && result.sheetNames.length === 1) {
      // Una sola hoja
      await importCsvDataToExcel(result.sheets[result.sheetNames[0]], file.name);
    } else if (ext === "csv" || ext === "txt") {
      // CSV o TXT
      await importCsvDataToExcel(result.text, file.name);
    } else {
      // Otros formatos - importar como texto
      await importTextToExcel(result.text, file.name);
    }

    showToast(`✅ ${file.name} importado a Excel`, "success");
  } catch (error) {
    console.error("Error importando archivo:", error);
    showToast(`Error: ${error instanceof Error ? error.message : "Error desconocido"}`, "error");
  }
}

/**
 * Muestra selector de hoja para archivos Excel con múltiples hojas
 */
function selectSheetToImport(sheetNames: string[]): Promise<string | null> {
  return new Promise((resolve) => {
    const overlay = document.createElement("div");
    overlay.className = "file-action-overlay";

    const dialog = document.createElement("div");
    dialog.className = "file-action-dialog";
    dialog.innerHTML = `
      <div class="file-action-header">
        <i class="ms-Icon ms-Icon--ExcelDocument" aria-hidden="true"></i>
        <div class="file-action-title">
          <span class="file-action-name">Seleccionar hoja</span>
        </div>
      </div>
      <p class="file-action-question">¿Qué hoja deseas importar?</p>
      <div class="file-action-options sheet-options">
        ${sheetNames.map(name => `
          <button class="file-action-btn sheet-btn" data-sheet="${name}">
            <i class="ms-Icon ms-Icon--Table" aria-hidden="true"></i>
            <span>${name}</span>
          </button>
        `).join("")}
      </div>
      <button class="file-action-cancel">Cancelar</button>
    `;

    overlay.appendChild(dialog);
    document.body.appendChild(overlay);

    const buttons = dialog.querySelectorAll(".sheet-btn");
    buttons.forEach(btn => {
      btn.addEventListener("click", () => {
        const sheet = (btn as HTMLElement).dataset.sheet || null;
        overlay.remove();
        resolve(sheet);
      });
    });

    const cancelBtn = dialog.querySelector(".file-action-cancel");
    cancelBtn?.addEventListener("click", () => {
      overlay.remove();
      resolve(null);
    });

    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) {
        overlay.remove();
        resolve(null);
      }
    });
  });
}

/**
 * Importa datos CSV a una nueva hoja de Excel
 */
async function importCsvDataToExcel(csvData: string, sheetName: string): Promise<void> {
  // Parsear CSV
  const rows = csvData.split("\n").filter(row => row.trim());
  if (rows.length === 0) {
    throw new Error("No hay datos para importar");
  }

  // Crear nombre de hoja válido (max 31 caracteres, sin caracteres especiales)
  const cleanSheetName = sheetName
    // eslint-disable-next-line no-useless-escape
    .replace(/[\\/*?:\[\]]/g, "")
    .substring(0, 31);

  // Convertir CSV a matriz 2D
  const data: string[][] = rows.map(row => {
    // Manejar campos con comas entre comillas
    const cells: string[] = [];
    let current = "";
    let inQuotes = false;

    for (let i = 0; i < row.length; i++) {
      const char = row[i];
      if (char === '"') {
        inQuotes = !inQuotes;
      } else if (char === "," && !inQuotes) {
        cells.push(current.trim());
        current = "";
      } else {
        current += char;
      }
    }
    cells.push(current.trim());
    return cells;
  });

  // Crear nueva hoja y escribir datos
  await excelService.createSheetWithData(cleanSheetName, data);
}

/**
 * Importa texto a una nueva hoja de Excel (una celda por línea)
 */
async function importTextToExcel(text: string, sheetName: string): Promise<void> {
  const lines = text.split("\n").filter(line => line.trim());
  if (lines.length === 0) {
    throw new Error("No hay datos para importar");
  }

  const cleanSheetName = sheetName
    // eslint-disable-next-line no-useless-escape
    .replace(/[\\/*?:\[\]]/g, "")
    .substring(0, 31);

  // Convertir líneas a matriz 2D (una columna)
  const data: string[][] = lines.map(line => [line]);

  await excelService.createSheetWithData(cleanSheetName, data);
}

/**
 * Lee un archivo como base64
 */
function readFileAsBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    if (file.size === 0) {
      reject(new Error(`El archivo "${file.name}" está vacío (0 bytes). Por favor verifica que el archivo tenga contenido.`));
      return;
    }

    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result as string;

      // Remover el prefijo "data:...;base64,"
      const commaIndex = result.indexOf(",");
      if (commaIndex === -1) {
        reject(new Error("Formato de archivo no válido"));
        return;
      }
      const base64 = result.substring(commaIndex + 1);

      if (!base64 || base64.length === 0) {
        reject(new Error("No se pudo leer el contenido del archivo"));
        return;
      }
      resolve(base64);
    };
    reader.onerror = () => {
      reject(reader.error);
    };
    reader.readAsDataURL(file);
  });
}

/**
 * Renderiza la vista previa de archivos adjuntos
 */
function renderAttachedFiles(): void {
  let container = el.attachedFilesContainer();

  // Crear contenedor si no existe
  if (!container) {
    container = document.createElement("div");
    container.id = "attachedFilesContainer";
    container.className = "attached-files-container";

    // Insertar antes del área de input
    const inputContainer = document.querySelector(".input-container");
    if (inputContainer) {
      inputContainer.insertBefore(container, inputContainer.firstChild);
    }
  }

  // Si no hay archivos, ocultar contenedor
  if (state.attachedFiles.length === 0) {
    container.style.display = "none";
    container.innerHTML = "";
    return;
  }

  container.style.display = "flex";

  // Renderizar archivos
  container.innerHTML = state.attachedFiles.map((file, index) => {
    const icon = getFileIcon(file.type);
    const sizeStr = formatFileSize(file.size);
    const statusClass = file.processing ? "processing" : (file.error ? "error" : "ready");
    const statusIcon = file.processing ? "ms-Icon--Sync spinning" : (file.error ? "ms-Icon--ErrorBadge" : "ms-Icon--CheckMark");

    return `
      <div class="attached-file ${statusClass}" data-index="${index}">
        <i class="ms-Icon ${icon}" aria-hidden="true"></i>
        <div class="attached-file-info">
          <span class="attached-file-name" title="${file.name}">${truncateFilename(file.name, 20)}</span>
          <span class="attached-file-size">${sizeStr}</span>
        </div>
        <i class="ms-Icon ${statusIcon} attached-file-status" aria-hidden="true"></i>
        <button class="attached-file-remove" data-index="${index}" title="Quitar archivo">
          <i class="ms-Icon ms-Icon--Cancel" aria-hidden="true"></i>
        </button>
      </div>
    `;
  }).join("");

  // Agregar listeners para remover archivos
  container.querySelectorAll(".attached-file-remove").forEach(btn => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const index = parseInt((btn as HTMLElement).dataset.index || "0");
      removeAttachedFile(index);
    });
  });
}

/**
 * Obtiene el icono según el tipo de archivo
 */
function getFileIcon(type: string): string {
  if (type.includes("pdf")) return "ms-Icon--PDF";
  if (type.includes("word") || type.includes("docx")) return "ms-Icon--WordDocument";
  if (type.includes("excel") || type.includes("xlsx") || type.includes("xls")) return "ms-Icon--ExcelDocument";
  if (type.includes("csv")) return "ms-Icon--Table";
  if (type.includes("json")) return "ms-Icon--Code";
  return "ms-Icon--TextDocument";
}

/**
 * Formatea el tamaño de archivo
 */
function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/**
 * Trunca el nombre de archivo si es muy largo
 */
function truncateFilename(name: string, maxLength: number): string {
  if (name.length <= maxLength) return name;
  const ext = name.split(".").pop() || "";
  const baseName = name.substring(0, name.length - ext.length - 1);
  const truncated = baseName.substring(0, maxLength - ext.length - 4) + "...";
  return truncated + "." + ext;
}

/**
 * Elimina un archivo adjunto
 */
function removeAttachedFile(index: number): void {
  state.attachedFiles.splice(index, 1);
  renderAttachedFiles();
}

/**
 * Limpia todos los archivos adjuntos
 */
function clearAttachedFiles(): void {
  state.attachedFiles = [];
  renderAttachedFiles();
}

/**
 * Obtiene el contexto de archivos adjuntos para el prompt
 */
function getAttachedFilesContext(): string {
  const readyFiles = state.attachedFiles.filter(f => f.content && !f.processing && !f.error);

  if (readyFiles.length === 0) return "";

  const contexts = readyFiles.map(file => {
    // Limitar el contenido a 15000 caracteres por archivo
    const content = file.content!.substring(0, 15000);
    const truncated = file.content!.length > 15000 ? "\n[... contenido truncado ...]" : "";
    return `[ARCHIVO ADJUNTO: ${file.name}]\n${content}${truncated}`;
  });

  return "\n\n" + contexts.join("\n\n");
}

// ===== Data Index =====

let isIndexing = false;

async function refreshDataIndex(showNotification: boolean = true): Promise<void> {
  if (isIndexing) return;

  isIndexing = true;

  try {
    if (showNotification) {
      showToast("Actualizando índice...", "info");
    }

    // Usar el índice ligero (solo metadatos, sin muestreo)
    const lightweightIndex = await excelService.buildLightweightIndex(true);

    if (lightweightIndex && showNotification) {
      const colCount = lightweightIndex.totalColumns;
      const rowCount = lightweightIndex.totalRows;
      showToast(`Índice: ${colCount} cols × ${rowCount} filas`, "success");
    } else if (!lightweightIndex && showNotification) {
      // Intentar verificar si hay conexión con Excel
      try {
        const usedRange = await excelService.getUsedRangeInfo();
        if (usedRange) {
          showToast(`Datos detectados pero índice falló. Reintenta.`, "info");
        } else {
          showToast("Hoja vacía - sin datos para indexar", "info");
        }
      } catch {
        showToast("Sin conexión con Excel", "error");
      }
    }
  } catch (error) {
    if (showNotification) {
      showToast("Error al indexar datos", "error");
    }
  } finally {
    isIndexing = false;
  }
}

// ===== Format Message =====

function formatContent(content: string): string {
  let formatted = content
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");

  // Code blocks (antes de tablas para no interferir)
  formatted = formatted.replace(
    /```(\w*)\n?([\s\S]*?)```/g,
    '<pre><code>$2</code></pre>'
  );

  // Markdown tables - detectar y convertir
  formatted = convertMarkdownTables(formatted);

  // Opciones clickeables - detectar listas numeradas/con letras
  formatted = convertOptionsToButtons(formatted);

  // Inline code
  formatted = formatted.replace(/`([^`]+)`/g, "<code>$1</code>");

  // Bold
  formatted = formatted.replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>");

  // Italic
  formatted = formatted.replace(/\*([^*]+)\*/g, "<em>$1</em>");

  return formatted;
}

/**
 * Convierte listas de opciones en botones clickeables
 */
function convertOptionsToButtons(content: string): string {
  const lines = content.split('\n');
  const result: string[] = [];
  let i = 0;
  let optionGroup: { number: string; text: string }[] = [];
  let questionLine = '';

  while (i < lines.length) {
    const line = lines[i];
    const trimmed = line.trim();

    // Detectar pregunta (línea que termina con ?)
    if (trimmed.endsWith('?') && !trimmed.startsWith('|')) {
      // Si había opciones pendientes, cerrarlas primero
      if (optionGroup.length > 0) {
        result.push(createOptionButtons(optionGroup, questionLine));
        optionGroup = [];
        questionLine = '';
      }
      questionLine = trimmed;
      result.push(line);
      i++;
      continue;
    }

    // Detectar opciones numeradas: "1. texto", "1) texto", "1- texto"
    const numberedMatch = trimmed.match(/^(\d+)[.\-)\s]+(.+)$/);
    if (numberedMatch) {
      optionGroup.push({ number: numberedMatch[1], text: numberedMatch[2].trim() });
      i++;
      continue;
    }

    // Detectar opciones con letras: "a) texto", "a. texto", "A) texto"
    const letterMatch = trimmed.match(/^([a-zA-Z])[.\-)]\s*(.+)$/);
    if (letterMatch && optionGroup.length > 0) {
      // Solo si ya hay opciones detectadas (para evitar falsos positivos)
      optionGroup.push({ number: letterMatch[1], text: letterMatch[2].trim() });
      i++;
      continue;
    }

    // Detectar opciones con viñetas: "- texto", "• texto", "* texto" (solo si parece opción corta)
    const bulletMatch = trimmed.match(/^[-•*]\s+(.{3,50})$/);
    if (bulletMatch && isLikelyOption(bulletMatch[1])) {
      optionGroup.push({ number: '•', text: bulletMatch[1].trim() });
      i++;
      continue;
    }

    // Si no es opción y teníamos opciones acumuladas, renderizarlas
    if (optionGroup.length >= 2) {
      result.push(createOptionButtons(optionGroup, questionLine));
      optionGroup = [];
      questionLine = '';
    } else if (optionGroup.length === 1) {
      // Solo una opción, mantenerla como texto normal
      const opt = optionGroup[0];
      result.push(`${opt.number}. ${opt.text}`);
      optionGroup = [];
    }

    result.push(line);
    i++;
  }

  // Si quedaron opciones al final
  if (optionGroup.length >= 2) {
    result.push(createOptionButtons(optionGroup, questionLine));
  } else if (optionGroup.length === 1) {
    const opt = optionGroup[0];
    result.push(`${opt.number}. ${opt.text}`);
  }

  return result.join('\n');
}

/**
 * Determina si un texto parece ser una opción seleccionable
 */
function isLikelyOption(text: string): boolean {
  // Opciones típicas son cortas y empiezan con mayúscula o son verbos
  if (text.length > 60) return false;
  if (text.length < 3) return false;
  if (text.includes('.') && text.split('.').length > 2) return false;

  // Excluir términos técnicos que NO son opciones clickeables
  const technicalPatterns = [
    /\b(RUC|NIT|TEXTO|JSON|XML|API|URL|PDF|Excel|CSV|KB|RAG)\b/i,
    /\b(columna|fila|celda|encabezado|formato|número|fecha)\b/i,
    /\b(como|para|con|sin|que|los|las|del|desde|hasta)\b/i,
    /\b(INSTRUCCIONES|IMPORTANTE|NOTA|OBLIGATORIO)\b/i
  ];

  if (technicalPatterns.some(p => p.test(text))) return false;

  // Patrones comunes de opciones REALES
  const optionPatterns = [
    /^(Sí|No|Si|Yes|Aceptar|Cancelar|Continuar|Salir|Ver más|Crear|Editar|Eliminar|Agregar)/i,
    /^(Opción|Option|Alternativa)\s*\d/i,
    /^(En una nueva hoja|En la hoja actual|Descargar)/i
  ];

  return optionPatterns.some(p => p.test(text));
}

/**
 * Crea HTML de botones para un grupo de opciones
 */
function createOptionButtons(options: { number: string; text: string }[], question: string): string {
  if (options.length < 2) return '';

  let html = '<div class="option-buttons-container">';

  options.forEach(opt => {
    // Limpiar markdown del texto para enviar (sin asteriscos)
    const cleanText = opt.text.replace(/\*\*([^*]+)\*\*/g, '$1').replace(/\*([^*]+)\*/g, '$1');

    // Procesar markdown para mostrar (convertir **texto** a <strong>)
    const displayContent = opt.text
      .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
      .replace(/\*([^*]+)\*/g, '<em>$1</em>');

    const displayText = opt.number !== '•' ? `${opt.number}. ${displayContent}` : displayContent;

    html += `<button class="option-btn" data-option-text="${escapeHtml(cleanText)}">${displayText}</button>`;
  });

  html += '</div>';
  return html;
}

/**
 * Escapa HTML para usar en atributos
 */
function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/**
 * Convierte tablas markdown a HTML con estilos bonitos
 */
function convertMarkdownTables(content: string): string {
  const lines = content.split('\n');
  const result: string[] = [];
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];

    // Detectar inicio de tabla (línea que empieza y termina con |)
    if (line.trim().startsWith('|') && line.trim().endsWith('|')) {
      // Verificar si la siguiente línea es el separador
      const nextLine = lines[i + 1];
      if (nextLine && /^\|[\s\-:|]+\|$/.test(nextLine.trim())) {
        // Es una tabla - procesar
        const tableLines: string[] = [line];
        let j = i + 1;

        // Recoger todas las líneas de la tabla
        while (j < lines.length && lines[j].trim().startsWith('|') && lines[j].trim().endsWith('|')) {
          tableLines.push(lines[j]);
          j++;
        }

        // Convertir a HTML
        const tableHtml = markdownTableToHtml(tableLines);
        result.push(tableHtml);
        i = j;
        continue;
      }
    }

    result.push(line);
    i++;
  }

  return result.join('\n');
}

/**
 * Convierte líneas de tabla markdown a HTML
 */
function markdownTableToHtml(lines: string[]): string {
  if (lines.length < 2) return lines.join('\n');

  // Parsear headers (primera línea)
  const headerCells = parseTableRow(lines[0]);

  // La segunda línea es el separador - la saltamos pero podemos usarla para alineación
  const alignments = parseAlignments(lines[1]);

  // El resto son filas de datos
  const dataRows = lines.slice(2).map(line => parseTableRow(line));

  // Detectar si la primera columna es un ranking
  const isRankingTable = detectRankingColumn(headerCells[0], dataRows.map(row => row[0]));

  // Construir HTML
  let html = '<div class="chat-table-wrapper"><table class="chat-table">';

  // Header
  html += '<thead><tr>';
  headerCells.forEach((cell, idx) => {
    const align = alignments[idx] || 'left';
    html += `<th style="text-align:${align}">${formatTableCell(cell)}</th>`;
  });
  html += '</tr></thead>';

  // Body
  html += '<tbody>';
  dataRows.forEach((row, rowIdx) => {
    html += '<tr>';
    row.forEach((cell, cellIdx) => {
      const align = alignments[cellIdx] || 'left';
      // Solo aplicar formato de ranking si es la primera columna Y es una tabla de ranking
      const isRankCol = cellIdx === 0 && isRankingTable;
      const formattedCell = isRankCol ? formatRankCell(cell, rowIdx) : formatTableCell(cell);
      html += `<td style="text-align:${align}">${formattedCell}</td>`;
    });
    html += '</tr>';
  });
  html += '</tbody>';

  html += '</table></div>';
  return html;
}

/**
 * Detecta si una columna es de ranking basándose en el header y los valores
 */
function detectRankingColumn(header: string, values: string[]): boolean {
  const headerLower = header.toLowerCase().trim();

  // Patrones de headers que indican ranking
  const rankingHeaders = [
    'posición', 'posicion', 'pos', 'pos.',
    'rank', 'ranking', '#', 'no.', 'n°', 'nº',
    'puesto', 'lugar', 'orden', 'top'
  ];

  const headerIsRanking = rankingHeaders.some(rh =>
    headerLower === rh || headerLower.includes(rh)
  );

  if (!headerIsRanking) return false;

  // Verificar que los valores son números secuenciales (1, 2, 3...) o (1°, 2°, 3°...)
  let expectedPos = 1;
  for (const val of values) {
    const cleaned = val.trim().replace(/[°º]/g, '');
    const num = parseInt(cleaned);
    if (isNaN(num) || num !== expectedPos) {
      return false;
    }
    expectedPos++;
  }

  return true;
}

/**
 * Parsea una fila de tabla markdown
 */
function parseTableRow(line: string): string[] {
  const trimmed = line.trim();
  const withoutBorders = trimmed.substring(1, trimmed.length - 1);
  return withoutBorders.split('|').map(cell => cell.trim());
}

/**
 * Parsea la línea de alineación para determinar left/center/right
 */
function parseAlignments(line: string): string[] {
  const cells = parseTableRow(line);
  return cells.map(cell => {
    const trimmed = cell.trim();
    const hasLeft = trimmed.startsWith(':');
    const hasRight = trimmed.endsWith(':');

    if (hasLeft && hasRight) return 'center';
    if (hasRight) return 'right';
    return 'left';
  });
}

/**
 * Formatea el contenido de una celda de tabla
 */
function formatTableCell(cell: string): string {
  let formatted = cell;
  // Bold
  formatted = formatted.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  // Números grandes con separador de miles
  formatted = formatted.replace(/\b(\d{4,})\b/g, (match) => {
    const num = parseInt(match);
    return num.toLocaleString('es-ES');
  });
  return formatted;
}

/**
 * Formatea celdas de ranking con medallas para top 3
 */
function formatRankCell(cell: string, rowIndex: number): string {
  const trimmed = cell.trim();

  // Si ya tiene emoji, dejarlo
  if (/[\u{1F3C6}\u{1F947}\u{1F948}\u{1F949}]/u.test(trimmed)) {
    return trimmed;
  }

  // Detectar si es un número de posición (1°, 2°, etc.)
  const posMatch = trimmed.match(/^(\d+)[°º]?$/);
  if (posMatch) {
    const pos = parseInt(posMatch[1]);
    if (pos === 1) return `<span class="rank-medal gold">🥇 1°</span>`;
    if (pos === 2) return `<span class="rank-medal silver">🥈 2°</span>`;
    if (pos === 3) return `<span class="rank-medal bronze">🥉 3°</span>`;
    return `${pos}°`;
  }

  return trimmed;
}

// ===== Actions Card =====

function createActionsCard(actions: ExcelAction[], results?: ActionResult[]): string {
  if (!actions || actions.length === 0) return "";

  let html = '<div class="actions-card">';

  actions.forEach((action, index) => {
    const result = results?.[index];
    const statusClass = result ? (result.success ? "success" : "error") : "pending";
    const description = action.description || `${action.type}`;

    // Show error message if failed
    const errorInfo = result && !result.success && result.error
      ? `<br><strong style="color:#d32f2f">Error:</strong> ${result.error}`
      : '';

    html += `<div class="action-item"><div class="action-row" data-index="${index}"><div class="action-indicator ${statusClass}"></div><span class="action-text">${description}</span><span class="action-range">${action.range}</span><i class="ms-Icon ms-Icon--ChevronDown action-chevron"></i></div><div class="action-details"><strong>Tipo:</strong> ${action.type}<br><strong>Rango:</strong> ${action.range}${result ? `<br><strong>Estado:</strong> ${result.success ? 'Completado' : 'Falló'}` : ''}${errorInfo}</div></div>`;
  });

  html += '</div>';
  return html;
}

function setupActionRowListeners(container: HTMLElement): void {
  const rows = container.querySelectorAll(".action-row");
  rows.forEach((row) => {
    row.addEventListener("click", () => {
      row.classList.toggle("expanded");
    });
  });
}

// ===== Messages =====

interface WebSource {
  title: string;
  url: string;
}

interface DetectedPdf {
  text: string;
  url: string;
  uploading?: boolean;
  uploaded?: boolean;
  error?: string;
}

/**
 * Crea el badge de fuentes web con lista expandible
 */
function createWebSourcesBadge(sources: WebSource[]): string {
  if (!sources || sources.length === 0) return "";

  const sourcesList = sources.map(s =>
    `<a href="${s.url}" target="_blank" rel="noopener noreferrer" class="web-source-item">
      <i class="ms-Icon ms-Icon--Link" aria-hidden="true"></i>
      <span>${s.title || s.url}</span>
    </a>`
  ).join("");

  return `<div class="web-sources-container">
    <div class="web-sources-badge" title="Clic para ver fuentes">
      <i class="ms-Icon ms-Icon--Globe" aria-hidden="true"></i>
      <span>${sources.length} Fuente${sources.length > 1 ? "s" : ""}</span>
      <i class="ms-Icon ms-Icon--ChevronDown web-sources-chevron" aria-hidden="true"></i>
    </div>
    <div class="web-sources-list">${sourcesList}</div>
  </div>`;
}

/**
 * Configura los listeners para expandir/colapsar fuentes web
 */
function setupWebSourcesListeners(container: HTMLElement): void {
  const badges = container.querySelectorAll(".web-sources-badge");
  badges.forEach((badge) => {
    badge.addEventListener("click", () => {
      const parent = badge.parentElement;
      if (parent) {
        parent.classList.toggle("expanded");
      }
    });
  });
}

/**
 * Crea el badge de PDFs detectados con opción de subir al RAG
 */
function createPdfUploadBadge(pdfs: DetectedPdf[]): string {
  if (!pdfs || pdfs.length === 0) return "";

  const pdfList = pdfs.map((pdf, index) => {
    // Limpiar el nombre para usarlo como filename
    const cleanName = (pdf.text || "documento")
      .replace(/[<>:"/\\|?*]/g, "") // Remover caracteres inválidos
      .replace(/\s+/g, "_") // Espacios a guiones bajos
      .substring(0, 100); // Limitar longitud
    return `<div class="pdf-item" data-pdf-index="${index}" data-pdf-url="${pdf.url}">
      <div class="pdf-info">
        <i class="ms-Icon ms-Icon--PDF" aria-hidden="true"></i>
        <span class="pdf-name">${pdf.text || "PDF"}</span>
      </div>
      <button class="pdf-upload-btn" data-pdf-url="${pdf.url}" data-pdf-name="${cleanName}" title="Subir a Knowledge Base">
        <i class="ms-Icon ms-Icon--CloudUpload" aria-hidden="true"></i>
        <span>Subir a RAG</span>
      </button>
    </div>`;
  }).join("");

  return `<div class="pdf-upload-container">
    <div class="pdf-upload-badge" title="PDFs detectados - Clic para expandir">
      <i class="ms-Icon ms-Icon--PDF" aria-hidden="true"></i>
      <span>${pdfs.length} PDF${pdfs.length > 1 ? "s" : ""} detectado${pdfs.length > 1 ? "s" : ""}</span>
      <i class="ms-Icon ms-Icon--ChevronDown pdf-chevron" aria-hidden="true"></i>
    </div>
    <div class="pdf-list">${pdfList}</div>
  </div>`;
}

/**
 * Configura los listeners para PDFs detectados
 */
function setupPdfUploadListeners(container: HTMLElement): void {
  // Toggle expandir/colapsar
  const badges = container.querySelectorAll(".pdf-upload-badge");
  badges.forEach((badge) => {
    badge.addEventListener("click", () => {
      const parent = badge.parentElement;
      if (parent) {
        parent.classList.toggle("expanded");
      }
    });
  });

  // Botones de subida
  const uploadBtns = container.querySelectorAll(".pdf-upload-btn");
  uploadBtns.forEach((btn) => {
    btn.addEventListener("click", async (e) => {
      e.stopPropagation();
      const pdfUrl = (btn as HTMLElement).dataset.pdfUrl;
      const pdfName = (btn as HTMLElement).dataset.pdfName;
      if (pdfUrl) {
        await handlePdfUpload(btn as HTMLElement, pdfUrl, pdfName);
      }
    });
  });
}

/**
 * Maneja la subida de un PDF al RAG
 * El proxy ya tiene la API Key configurada, no se necesita del cliente
 */
async function handlePdfUpload(btn: HTMLElement, pdfUrl: string, pdfName?: string): Promise<void> {
  // Cambiar estado del botón
  const originalHtml = btn.innerHTML;
  btn.innerHTML = '<i class="ms-Icon ms-Icon--Sync spinning" aria-hidden="true"></i><span>Subiendo...</span>';
  btn.classList.add("uploading");
  (btn as HTMLButtonElement).disabled = true;

  try {
    // El proxy usa la API key configurada en sus variables de entorno
    // Pasar el nombre del PDF para que se guarde con un nombre descriptivo
    const result = await uploadPdfToRag(pdfUrl, undefined, pdfName);

    if (result.success) {
      btn.innerHTML = '<i class="ms-Icon ms-Icon--CheckMark" aria-hidden="true"></i><span>Subido</span>';
      btn.classList.remove("uploading");
      btn.classList.add("uploaded");
      showToast(`PDF subido: ${result.file?.filename || "OK"}`, "success");
    } else {
      throw new Error(result.error || "Error desconocido");
    }
  } catch (error) {
    console.error("Error subiendo PDF:", error);
    btn.innerHTML = '<i class="ms-Icon ms-Icon--ErrorBadge" aria-hidden="true"></i><span>Error</span>';
    btn.classList.remove("uploading");
    btn.classList.add("error");
    (btn as HTMLButtonElement).disabled = false;

    const errorMsg = error instanceof Error ? error.message : "Error desconocido";
    showToast(`Error: ${errorMsg}`, "error");

    // Restaurar después de 3 segundos
    setTimeout(() => {
      btn.innerHTML = originalHtml;
      btn.classList.remove("error");
    }, 3000);
  }
}

/**
 * Efecto typewriter para mostrar texto progresivamente
 */
async function typewriterEffect(element: HTMLElement, content: string, actions?: ExcelAction[], results?: ActionResult[], webSources?: WebSource[], detectedPdfs?: DetectedPdf[]): Promise<void> {
  const formattedContent = formatContent(content);
  const actionsHtml = (actions && actions.length > 0) ? createActionsCard(actions, results) : "";
  const sourcesHtml = createWebSourcesBadge(webSources || []);
  const pdfsHtml = createPdfUploadBadge(detectedPdfs || []);

  // Crear un elemento temporal para parsear el HTML formateado
  const tempDiv = document.createElement("div");
  tempDiv.innerHTML = formattedContent;

  // Obtener el texto plano para el efecto
  const plainText = tempDiv.textContent || "";

  // Velocidad de escritura (ms por caracter)
  const speed = 15;
  const minChunkSize = 3;

  let displayedText = "";
  let i = 0;

  return new Promise((resolve) => {
    function typeNextChunk(): void {
      if (i < plainText.length) {
        // Escribir varios caracteres a la vez para mayor fluidez
        const chunkSize = Math.min(minChunkSize, plainText.length - i);
        displayedText += plainText.substring(i, i + chunkSize);
        i += chunkSize;

        // Mostrar el texto con formato parcial
        element.innerHTML = formatContent(displayedText) + '<span class="typing-cursor">|</span>';

        // Scroll al final
        const container = el.messagesContainer();
        container.scrollTop = container.scrollHeight;

        setTimeout(typeNextChunk, speed);
      } else {
        // Finalizado: mostrar contenido completo con acciones, fuentes y PDFs
        element.innerHTML = formattedContent + actionsHtml + sourcesHtml + pdfsHtml;

        // Setup listeners
        if (actions && actions.length > 0) {
          setupActionRowListeners(element.parentElement as HTMLElement);
        }
        if (webSources && webSources.length > 0) {
          setupWebSourcesListeners(element);
        }
        if (detectedPdfs && detectedPdfs.length > 0) {
          setupPdfUploadListeners(element);
        }

        resolve();
      }
    }

    typeNextChunk();
  });
}

function addMessage(role: "user" | "assistant", content: string, actions?: ExcelAction[], results?: ActionResult[], webSources?: WebSource[], detectedPdfs?: DetectedPdf[]): void {
  const container = el.messagesContainer();

  // Remove welcome
  const welcome = container.querySelector(".welcome");
  if (welcome) {
    welcome.remove();
  }

  const messageDiv = document.createElement("div");
  messageDiv.className = `message ${role}`;

  const bubbleDiv = document.createElement("div");
  bubbleDiv.className = "message-bubble";

  if (role === "assistant") {
    // Efecto typewriter para mensajes del asistente
    bubbleDiv.innerHTML = '<span class="typing-cursor">|</span>';
    messageDiv.appendChild(bubbleDiv);
    container.appendChild(messageDiv);
    container.scrollTop = container.scrollHeight;

    // Iniciar efecto typewriter
    typewriterEffect(bubbleDiv, content, actions, results, webSources, detectedPdfs);
  } else {
    // Mensajes del usuario aparecen inmediatamente
    bubbleDiv.innerHTML = formatContent(content);
    messageDiv.appendChild(bubbleDiv);
    container.appendChild(messageDiv);
    container.scrollTop = container.scrollHeight;
  }
}

function clearMessages(): void {
  const container = el.messagesContainer();
  container.innerHTML = getWelcomeHTML();
}

/**
 * Genera el HTML del mensaje de bienvenida personalizado
 */
function getWelcomeHTML(): string {
  const user = state.currentUser;
  const greeting = user ? getGreeting(user.firstName) : 'Hola';

  return `
    <div class="welcome">
      <div class="welcome-icon">
        <i class="ms-Icon ms-Icon--ExcelDocument" aria-hidden="true"></i>
      </div>
      <p class="welcome-greeting">${greeting}</p>
      <p class="welcome-text">
        Tu asistente de IA para análisis de datos, fórmulas y automatización en Excel.
      </p>
      <p class="welcome-examples-text">
        <strong>Prueba:</strong> "Crea un calendario", "Fórmula de IVA 15%", "Analiza mis datos"
      </p>
    </div>
  `;
}

// ===== Pending Actions =====

function showPendingActions(actions: ExcelAction[]): void {
  state.pendingActions = actions;
  el.pendingCount().textContent = `${actions.length} ediciones pendientes`;
  el.pendingActions().classList.remove("hidden");
  updateAcceptButtonState();
}

function hidePendingActions(): void {
  state.pendingActions = [];
  el.pendingActions().classList.add("hidden");
  updateAcceptButtonState();
}

function updateAcceptButtonState(): void {
  const btn = el.acceptAllBtn();
  if (btn) {
    btn.disabled = state.isProcessing || state.pendingActions.length === 0;
  }
}

async function acceptAllActions(): Promise<void> {
  if (state.pendingActions.length === 0) return;

  setLoading(true);

  try {
    const results = await excelService.executeActions(state.pendingActions);
    state.lastActionResults = results;

    const successCount = results.filter(r => r.success).length;
    const errorCount = results.length - successCount;

    // Update the last message with results
    const lastMessage = el.messagesContainer().querySelector(".message.assistant:last-child .message-bubble");
    if (lastMessage) {
      const actionsCard = lastMessage.querySelector(".actions-card");
      if (actionsCard) {
        actionsCard.outerHTML = createActionsCard(state.pendingActions, results);
        setupActionRowListeners(lastMessage as HTMLElement);
      }
    }

    if (errorCount === 0) {
      showToast(`${successCount} acciones completadas`, "success");

      // Highlight first range
      if (results.length > 0 && results[0].success) {
        try {
          await excelService.highlightRange(results[0].action.range, 1500);
        } catch {
          // Ignore highlight errors
        }
      }
      
      // Verificación post-ejecución: comprobar si hay fórmulas con resultados anómalos
      const formulaActions = state.pendingActions.filter(a => a.type === "formula");
      if (formulaActions.length > 0) {
        // Dar tiempo a Excel para calcular las fórmulas
        await new Promise(resolve => setTimeout(resolve, 500));
        
        // Verificar resultados
        const verificationResult = await verifyFormulaResults(formulaActions);
        if (verificationResult.needsCorrection) {
          showToast("⚠️ Detectados resultados anómalos, solicitando corrección...", "info");
          setLoading(false);
          hidePendingActions();
          
          // Solicitar corrección con el feedback
          await requestResultCorrection(formulaActions, verificationResult.feedback, state.lastUserMessage || "");
          return;
        }
      }
    } else {
      showToast(`${successCount} OK, ${errorCount} errores - reintentando...`, "info");

      // Solicitar corrección al modelo
      // NO llamamos hidePendingActions() aquí ni en finally
      // porque requestErrorCorrection se encarga del flujo completo
      setLoading(false); // Liberar loading antes de la corrección
      requestErrorCorrection(results); // Sin await - maneja su propio flujo
      return; // Salir sin ejecutar finally
    }

  } catch (error) {
    console.error("Error ejecutando acciones:", error);
    showToast("Error al ejecutar acciones", "error");
    hidePendingActions();
    setLoading(false);
  }

  // Solo limpiar si no hubo errores (no entramos al else anterior)
  hidePendingActions();
  setLoading(false);
}

/**
 * Solicita al modelo que corrija las acciones que fallaron
 * Maneja su propio ciclo de loading y ejecución
 */
async function requestErrorCorrection(results: ActionResult[]): Promise<void> {
  const failedActions = results.filter(r => !r.success);

  if (failedActions.length === 0) {
    return;
  }

  // Construir mensaje de error para el modelo
  const errorDetails = failedActions.map(r =>
    `- Acción "${r.action.type}" en ${r.action.range}: ${r.error}`
  ).join("\n");

  const correctionRequest = `[ERROR EN ACCIONES ANTERIORES]
Las siguientes acciones fallaron:
${errorDetails}

ERRORES COMUNES Y SOLUCIONES:
1. "matriz de entrada no coincide con el tamaño" → Las fórmulas UNIQUE/SORT/FILTER van en UNA SOLA CELDA (ej: "A2"), NO en rangos (ej: "A2:A100"). Excel las derrama automáticamente.
2. "Ya existe un recurso con el mismo nombre" → La hoja ya existe, usa un nombre diferente o activa la hoja existente.
3. Referencia a hoja incorrecta → Usa el nombre exacto de la hoja del índice.

Por favor, genera las acciones CORREGIDAS para completar la tarea.`;

  setLoading(true);

  try {
    // Obtener contexto actualizado (por si la hoja cambió)
    let context = "";
    const usedRange = await excelService.getUsedRangeInfo();
    if (usedRange) {
      const lightweightIndex = await excelService.buildLightweightIndex(true); // Forzar rebuild
      if (lightweightIndex) {
        context = excelService.formatLightweightIndexForContext(lightweightIndex);
      }
    }

    addMessage("user", "⚠️ Corrigiendo errores...");

    const response = await azureOpenAIService.sendMessageStructured(
      correctionRequest,
      context || undefined
    );

    if (response.actions && response.actions.length > 0) {
      addMessage("assistant", response.message, response.actions);
      showPendingActions(response.actions);

      // Auto-ejecutar la corrección
      if (state.editMode === "auto") {
        // Ejecutar directamente, no con setTimeout
        // Pequeña pausa para que se renderice la UI
        await new Promise(resolve => setTimeout(resolve, 300));
        await acceptAllActions();
      } else {
        setLoading(false);
      }
    } else {
      addMessage("assistant", response.message);
      setLoading(false);
    }
  } catch (error) {
    console.error("Error solicitando corrección:", error);
    showToast("Error al solicitar corrección", "error");
    setLoading(false);
  }
}

/**
 * Verifica si los resultados de las fórmulas parecen correctos
 * Detecta anomalías como: todos ceros, #ERROR, valores inesperados
 */
async function verifyFormulaResults(formulaActions: ExcelAction[]): Promise<{ needsCorrection: boolean; feedback: string }> {
  try {
    // Leer los valores resultantes de las fórmulas
    const anomalies: string[] = [];
    
    for (const action of formulaActions) {
      try {
        // Leer el rango donde se escribió la fórmula
        const data = await excelService.readRange(action.range, action.sheetName);
        const values = Object.values(data.cells);
        
        // Verificar anomalías
        const numericValues = values.filter(v => typeof v === "number") as number[];
        const stringValues = values.filter(v => typeof v === "string") as string[];
        
        // Detectar errores de Excel
        const errorValues = stringValues.filter(v => 
          v.startsWith("#") || v.includes("ERROR") || v.includes("¡")
        );
        if (errorValues.length > 0) {
          anomalies.push(`- Fórmula en ${action.range}: Errores de Excel detectados: ${errorValues.slice(0, 3).join(", ")}`);
          continue;
        }
        
        // Detectar si TODOS los valores numéricos son cero (anómalo para sumas/conteos)
        if (numericValues.length > 5 && numericValues.every(v => v === 0)) {
          const formulaText = action.formula || (action.formulas ? JSON.stringify(action.formulas[0]) : "desconocida");
          anomalies.push(`- Fórmula en ${action.range}: TODOS los valores son 0 (${numericValues.length} ceros). Fórmula: ${formulaText.substring(0, 100)}`);
          continue;
        }
        
        // Detectar si hay demasiados ceros (más del 90%)
        if (numericValues.length > 10) {
          const zeroCount = numericValues.filter(v => v === 0).length;
          const zeroPercentage = (zeroCount / numericValues.length) * 100;
          if (zeroPercentage > 90) {
            anomalies.push(`- Fórmula en ${action.range}: ${zeroPercentage.toFixed(0)}% de los valores son 0 (${zeroCount}/${numericValues.length}). Posible error en referencias.`);
          }
        }
        
      } catch (e) {
        console.warn(`No se pudo verificar ${action.range}:`, e);
      }
    }
    
    if (anomalies.length > 0) {
      return {
        needsCorrection: true,
        feedback: anomalies.join("\n")
      };
    }
    
    return { needsCorrection: false, feedback: "" };
    
  } catch (error) {
    console.error("Error verificando resultados:", error);
    return { needsCorrection: false, feedback: "" };
  }
}

/**
 * Solicita corrección cuando los resultados parecen incorrectos
 */
async function requestResultCorrection(
  formulaActions: ExcelAction[], 
  feedback: string, 
  originalRequest: string
): Promise<void> {
  setLoading(true);
  
  const correctionRequest = `[VERIFICACIÓN DE RESULTADOS - ANOMALÍAS DETECTADAS]

Las fórmulas se ejecutaron pero los resultados parecen INCORRECTOS:
${feedback}

SOLICITUD ORIGINAL DEL USUARIO: "${originalRequest}"

PROBLEMAS COMUNES Y SOLUCIONES:
1. **Todos los valores en 0**: Las referencias de columna son incorrectas. Verifica:
   - ¿La columna de fechas tiene formato de fecha real o es texto?
   - ¿La columna de valores tiene números o texto?
   - ¿Las referencias apuntan a la hoja correcta?

2. **SUMIF/COUNTIF devuelve 0**: 
   - El criterio no coincide (mayúsculas vs minúsculas, espacios extra)
   - La columna de criterio tiene formato diferente al esperado
   - Usa SUMIFS con referencia a la misma hoja: SUMIFS('NombreHoja'!EU:EU,'NombreHoja'!AJ:AJ,A2)

3. **Fórmulas de fecha**:
   - YEAR() solo funciona con fechas reales, no con texto
   - Usa TEXT(AJ2,"YYYY") si la fecha es texto
   - O verifica el formato real de la columna

ACCIÓN REQUERIDA:
1. Analiza qué puede estar mal con las referencias
2. Genera NUEVAS fórmulas corregidas
3. Si no estás seguro del formato de los datos, primero haz una acción "calc" para verificar el tipo de datos:
   =TYPE(AJ2) → 1=número/fecha, 2=texto
   =LEFT(AJ2,10) → ver qué contiene

Por favor, genera las acciones CORREGIDAS.`;

  try {
    // Obtener contexto actualizado
    let context = "";
    const lightweightIndex = await excelService.buildLightweightIndex(true);
    if (lightweightIndex) {
      context = excelService.formatLightweightIndexForContext(lightweightIndex);
    }

    addMessage("assistant", "⚠️ Los resultados parecen incorrectos (muchos ceros). Analizando y corrigiendo...");

    const response = await azureOpenAIService.sendMessageStructured(
      correctionRequest,
      context || undefined
    );

    if (response.actions && response.actions.length > 0) {
      addMessage("assistant", response.message, response.actions);
      showPendingActions(response.actions);

      if (state.editMode === "auto") {
        await new Promise(resolve => setTimeout(resolve, 300));
        await acceptAllActions();
      } else {
        setLoading(false);
      }
    } else {
      addMessage("assistant", response.message);
      setLoading(false);
    }
  } catch (error) {
    console.error("Error solicitando corrección de resultados:", error);
    showToast("Error al solicitar corrección", "error");
    setLoading(false);
  }
}

// ===== Send Message =====

/**
 * Envía un mensaje de forma silenciosa (sin mostrarlo en el chat)
 * Útil para prompts automáticos de extracción
 */
async function sendMessageSilent(message: string, displayMessage?: string): Promise<void> {
  // Protección contra doble ejecución
  if (!message || state.isProcessing) {
    return;
  }

  const validation = validateConfig();
  if (!validation.isValid) {
    showToast("Configura tu API Key en config.ts", "error");
    return;
  }

  // Marcar como procesando ANTES de cualquier operación async
  state.isProcessing = true;

  // Mostrar mensaje simplificado en lugar del prompt completo
  if (displayMessage) {
    addMessage("user", displayMessage);
  }

  setLoading(true);

  try {
    // Obtener contexto de archivos adjuntos
    const attachedFilesContext = getAttachedFilesContext();
    let context = attachedFilesContext || "";

    // Agregar contexto de Excel
    const usedRange = await excelService.getUsedRangeInfo();
    if (usedRange) {
      context += `\n[Hoja tiene datos en: ${usedRange.address}. Para nuevo contenido, usar siguiente fila disponible.]`;
    } else {
      context += `\n[Hoja vacía. Empezar en A1.]`;
    }

    // Enviar al modelo
    const response = await azureOpenAIService.sendMessageStructured(message, context || undefined);

    // Limpiar archivos adjuntos después de enviar
    state.attachedFiles = [];
    renderAttachedFiles();

    // Procesar respuesta normalmente
    if (response.actions && response.actions.length > 0) {
      addMessage("assistant", response.message, response.actions);
      showPendingActions(response.actions);

      // Auto-ejecutar si está en modo automático
      if (state.editMode === "auto") {
        await new Promise(resolve => setTimeout(resolve, 300));
        await acceptAllActions();
      }
    } else {
      addMessage("assistant", response.message);
    }

  } catch (error) {
    console.error("Error en mensaje silencioso:", error);
    addMessage("assistant", "Error al procesar la solicitud. Por favor intenta de nuevo.");
  } finally {
    setLoading(false);
  }
}

async function sendMessage(): Promise<void> {
  const input = el.userInput();
  const userMessage = input.value.trim();

  if (!userMessage || state.isProcessing) return;

  const validation = validateConfig();
  if (!validation.isValid) {
    showToast("Configura tu API Key en config.ts", "error");
    return;
  }

  // Guardar mensaje para posibles correcciones posteriores
  state.lastUserMessage = userMessage;

  // Clear input
  input.value = "";
  autoResizeTextarea();
  updateInputState();
  hidePendingActions();

  addMessage("user", userMessage);
  setLoading(true);

  try {
    // Agregar contexto de selección y rango usado automáticamente
    let context = state.currentExcelContext || "";

    // Obtener información del rango usado (área con datos)
    const usedRange = await excelService.getUsedRangeInfo();
    if (usedRange) {
      const nextFreeRow = usedRange.lastRow + 2; // Dejar una fila de espacio
      context += `\n[Hoja tiene datos en: ${usedRange.address}. Última fila: ${usedRange.lastRow}. Para nuevo contenido, usar fila ${nextFreeRow} o columna después de ${usedRange.lastColumn}.]`;

      // Usar el índice ligero (solo metadatos, sin muestreo)
      const lightweightIndex = await excelService.buildLightweightIndex();
      if (lightweightIndex) {
        // Incluir el índice ligero formateado
        context += `\n${excelService.formatLightweightIndexForContext(lightweightIndex)}`;
      } else {
        // Fallback a solo encabezados si no se pudo crear índice
        const { headers, columnLetters } = await excelService.getHeaders();
        if (headers.length > 0) {
          const headerInfo = headers.slice(0, 50).map((h, i) => `${columnLetters[i]}:${h}`).join(", ");
          context += `\n[ENCABEZADOS: ${headerInfo}]`;
        }
      }
    } else {
      context += `\n[Hoja vacía. Puedes empezar en A1.]`;
    }

    // Agregar información de la selección
    // Si el usuario pregunta específicamente por "datos seleccionados", incluir los datos completos
    const wantsSelectedData = asksForSelectedData(userMessage);

    if (state.currentSelection) {
      const sel = state.currentSelection;
      const cellCount = sel.rowCount * sel.columnCount;

      if (wantsSelectedData && cellCount > 1) {
        // El usuario pregunta por datos seleccionados y tiene un rango seleccionado
        try {
          const selectedDataText = await excelService.getSelectedRangeAsText();
          context += `\n\n[DATOS SELECCIONADOS POR EL USUARIO - Rango ${sel.address}]\n${selectedDataText}\n[FIN DATOS SELECCIONADOS]`;
        } catch {
          context += `\n[Selección: ${sel.address} (${cellCount} celdas) - No se pudieron obtener los datos]`;
        }
      } else if (sel.hasContent) {
        // Selección normal con contenido - solo informar la dirección
        context += `\n[Celda seleccionada: ${sel.address} - CONTIENE DATOS: "${sel.firstCellValue}". NO sobrescribir.]`;
      } else {
        // Selección vacía
        context += `\n[Celda seleccionada: ${sel.address} - VACÍA. Puedes usar como punto de inicio.]`;
      }
    }

    // Búsqueda web si está habilitada
    let webSources: WebSource[] = [];
    let detectedPdfs: DetectedPdf[] = [];
    if (state.webSearchEnabled) {
      try {
        // Primero verificar si el usuario proporcionó una URL directamente
        const directUrl = extractUrlFromMessage(userMessage);

        if (directUrl) {
          // Fetch directo del contenido de la URL (intenta estático, luego Puppeteer si es necesario)
          const webContent = await fetchWebContentSmart(directUrl);

          if (!webContent.error) {
            webSources = [{ title: webContent.title || directUrl, url: directUrl }];
            const webContext = formatWebContentForContext(webContent);
            context += `\n${webContext}`;

            // Detectar PDFs para ofrecer subida al RAG
            const pdfLinks = getPdfLinksFromContent(webContent);
            if (pdfLinks.length > 0) {
              detectedPdfs = pdfLinks.map(link => ({
                text: link.text,
                url: link.url
              }));
            }
          }
        } else if (shouldSearchWeb(userMessage)) {
          // Solo buscar si el mensaje realmente lo requiere (no saludos simples)
          const searchResults = await searchWeb(userMessage, { maxResults: 5 });

          if (searchResults.results.length > 0) {
            webSources = searchResults.results.map(r => ({ title: r.title, url: r.url }));
            const webContext = formatSearchResultsForContext(searchResults);
            context += `\n${webContext}`;
          }
        }
      } catch {
        // Continuar sin resultados de búsqueda
      }
    }

    // Agregar contexto de archivos adjuntos
    const attachedFilesContext = getAttachedFilesContext();
    if (attachedFilesContext) {
      context += attachedFilesContext;
    }

    // Consulta RAG (Knowledge Base) si está habilitada
    // Usa retrieve-only para obtener fragmentos relevantes sin necesidad de modelo en Open WebUI
    let ragContext = "";
    if (state.ragEnabled && shouldQueryRag(userMessage)) {
      // Solo consultar RAG si el mensaje es una consulta real (no saludos simples)
      try {
        const ragResult = await retrieveFromRag(userMessage, 8); // Top 8 fragmentos

        if (ragResult.success && ragResult.chunks && ragResult.chunks.length > 0) {
          ragContext = formatRagChunksForContext(ragResult);
          context += ragContext;
        }
      } catch {
        // Continuar sin resultados de RAG
      }
    }

    let response: StructuredResponse = await azureOpenAIService.sendMessageStructured(
      userMessage,
      context || undefined
    );

    // Verificar si hay acciones de conteo por categoría (countByCategory) - MÁS EFICIENTE
    if (response.actions && response.actions.some(a => a.type === "countByCategory")) {
      const countActions = response.actions.filter(a => a.type === "countByCategory");
      const otherActions = response.actions.filter(a => a.type !== "countByCategory");

      addMessage("assistant", `📊 Contando por categoría...`, countActions);

      const countResults: string[] = [];
      const countActionResults: ActionResult[] = [];

      for (const countAction of countActions) {
        try {
          if (countAction.categoryColumn) {
            const results = await excelService.getUniqueCountsByCategory(
              countAction.categoryColumn,
              countAction.filterColumn,
              countAction.filterValue,
              countAction.sheetName
            );
            
            // Formatear resultados para mostrar
            const total = results.reduce((sum, r) => sum + r.count, 0);
            const formattedResults = results.map(r => `• ${r.category}: ${r.count}`).join("\n");
            
            countResults.push(`[CONTEO POR CATEGORÍA: ${countAction.description || ""}]\n${formattedResults}\n\n📊 TOTAL: ${total}`);
            
            countActionResults.push({
              success: true,
              action: countAction,
              message: `${results.length} categorías, total: ${total}`
            });
          } else {
            countResults.push(`[ERROR: Falta columna de categoría]`);
            countActionResults.push({
              success: false,
              action: countAction,
              message: "Sin columna",
              error: "No se especificó la columna de categorías"
            });
          }
        } catch (error) {
          const errorMsg = error instanceof Error ? error.message : "Error desconocido";
          countResults.push(`[ERROR: ${errorMsg}]`);
          countActionResults.push({
            success: false,
            action: countAction,
            message: "Error",
            error: errorMsg
          });
        }
      }

      // Actualizar tarjeta
      const lastMessage = el.messagesContainer().querySelector(".message.assistant:last-child .message-bubble");
      if (lastMessage) {
        const actionsCard = lastMessage.querySelector(".actions-card");
        if (actionsCard) {
          actionsCard.outerHTML = createActionsCard(countActions, countActionResults);
          setupActionRowListeners(lastMessage as HTMLElement);
        }
      }

      // Construir respuesta con los resultados
      const followUpPrompt = `[RESULTADOS DE CONTEO POR CATEGORÍA - DATOS 100% REALES]
${countResults.join("\n\n")}

Ahora responde la pregunta original del usuario: "${userMessage}"
Presenta estos datos de forma clara y legible. NO inventes cifras adicionales.

${otherActions.length > 0 ? `Nota: También habías propuesto estas acciones: ${JSON.stringify(otherActions)}` : ""}`;

      response = await azureOpenAIService.sendMessageStructured(
        followUpPrompt,
        context || undefined
      );
    }

    // Verificar si hay acciones de promedio por categoría (avgByCategory)
    if (response.actions && response.actions.some(a => a.type === "avgByCategory")) {
      const avgActions = response.actions.filter(a => a.type === "avgByCategory");
      const otherActions = response.actions.filter(a => a.type !== "avgByCategory");

      addMessage("assistant", `📊 Calculando promedios por categoría...`, avgActions);

      const avgResults: string[] = [];
      const avgActionResults: ActionResult[] = [];

      for (const avgAction of avgActions) {
        try {
          if (avgAction.categoryColumn && avgAction.valueColumn) {
            const results = await excelService.getAverageByCategory(
              avgAction.categoryColumn,
              avgAction.valueColumn,
              avgAction.filterColumn,
              avgAction.filterValue,
              avgAction.sheetName
            );
            
            // Formatear resultados para mostrar (limitar a top 30 para legibilidad)
            const topResults = results.slice(0, 30);
            const formattedResults = topResults.map(r => `• ${r.category}: ${r.average} (${r.count} registros)`).join("\n");
            const globalAvg = results.length > 0 
              ? Math.round(results.reduce((sum, r) => sum + r.average * r.count, 0) / results.reduce((sum, r) => sum + r.count, 0) * 100) / 100
              : 0;
            
            avgResults.push(`[PROMEDIO POR CATEGORÍA: ${avgAction.description || ""}]\n${formattedResults}${results.length > 30 ? `\n... y ${results.length - 30} categorías más` : ""}\n\n📊 PROMEDIO GLOBAL: ${globalAvg} | Total categorías: ${results.length}`);
            
            avgActionResults.push({
              success: true,
              action: avgAction,
              message: `${results.length} categorías, promedio global: ${globalAvg}`
            });
          } else {
            avgResults.push(`[ERROR: Falta columna de categoría o valores]`);
            avgActionResults.push({
              success: false,
              action: avgAction,
              message: "Sin columnas",
              error: "No se especificó categoryColumn o valueColumn"
            });
          }
        } catch (error) {
          const errorMsg = error instanceof Error ? error.message : "Error desconocido";
          avgResults.push(`[ERROR: ${errorMsg}]`);
          avgActionResults.push({
            success: false,
            action: avgAction,
            message: "Error",
            error: errorMsg
          });
        }
      }

      // Actualizar tarjeta
      const lastMessage = el.messagesContainer().querySelector(".message.assistant:last-child .message-bubble");
      if (lastMessage) {
        const actionsCard = lastMessage.querySelector(".actions-card");
        if (actionsCard) {
          actionsCard.outerHTML = createActionsCard(avgActions, avgActionResults);
          setupActionRowListeners(lastMessage as HTMLElement);
        }
      }

      // Construir respuesta con los resultados
      const followUpPrompt = `[RESULTADOS DE PROMEDIO POR CATEGORÍA - DATOS 100% REALES]
${avgResults.join("\n\n")}

Ahora responde la pregunta original del usuario: "${userMessage}"
Presenta estos datos de forma clara y legible. NO inventes cifras.

${otherActions.length > 0 ? `Nota: También habías propuesto estas acciones: ${JSON.stringify(otherActions)}` : ""}`;

      response = await azureOpenAIService.sendMessageStructured(
        followUpPrompt,
        context || undefined
      );
    }

    // Verificar si hay acciones de cálculo (calc) - procesar primero
    if (response.actions && response.actions.some(a => a.type === "calc")) {
      const calcActions = response.actions.filter(a => a.type === "calc");
      const otherActions = response.actions.filter(a => a.type !== "calc");

      // Mostrar mensaje con tarjeta de acciones de cálculo
      addMessage("assistant", `🧮 Calculando con datos reales...`, calcActions);

      // Ejecutar cálculos en hoja oculta
      const calcResults: string[] = [];
      const calcActionResults: ActionResult[] = [];
      const lightweightIndex = await excelService.buildLightweightIndex();
      const sheetName = lightweightIndex?.sheetName || "Hoja1";

      for (const calcAction of calcActions) {
        try {
          // Aceptar tanto calcFormulas como formulas (por compatibilidad)
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const formulas = calcAction.calcFormulas || (calcAction as any).formulas;
          
          if (formulas && formulas.length > 0) {
            // Timeout de 30 segundos para evitar que se quede colgado
            const timeoutPromise = new Promise<never>((_, reject) => 
              setTimeout(() => reject(new Error("Timeout: el cálculo tardó más de 30 segundos")), 30000)
            );
            
            const results = await Promise.race([
              excelService.executeCalcFormulas(formulas, sheetName),
              timeoutPromise
            ]);
            
            // Detectar errores de Excel en los resultados
            const excelErrors = ["#SPILL!", "#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#N/A", "#NULL!", "#NUM!", "#GETTING_DATA"];
            const hasErrors = results.some(r => 
              r.result != null && typeof r.result === "string" && excelErrors.some(err => String(r.result).includes(err))
            );
            
            const errorResults = results.filter(r => 
              r.result != null && typeof r.result === "string" && excelErrors.some(err => String(r.result).includes(err))
            );
            
            if (hasErrors && errorResults.length > 0) {
              // Guardar las fórmulas con error para retroalimentación
              const errorFeedback = errorResults.map(r => 
                `- ${r.formula} → ERROR: ${r.result}`
              ).join("\n");
              
              calcResults.push(`[ERRORES EN FÓRMULAS - NECESITA CORRECCIÓN]\n${errorFeedback}\n\nNota: #SPILL! indica que la fórmula UNIQUE o similar intenta devolver demasiados valores. Usa fórmulas COUNTIFS individuales para cada zona en lugar de UNIQUE.`);
              
              calcActionResults.push({
                success: false,
                action: calcAction,
                message: `Errores: ${errorResults.length}/${results.length}`,
                error: `Errores de Excel: ${errorResults.map(r => r.result).join(", ")}`
              });
            } else {
              const formattedResults = results.map(r => `${r.formula} = ${r.result}`).join("\n");
              calcResults.push(`[CÁLCULOS EJECUTADOS]\n${formattedResults}`);

              // Marcar como exitoso
              calcActionResults.push({
                success: true,
                action: calcAction,
                message: `Calculado: ${results.length} fórmulas`
              });
            }
          } else {
            calcResults.push(`[ERROR: Acción calc sin fórmulas definidas]`);
            calcActionResults.push({
              success: false,
              action: calcAction,
              message: "Sin fórmulas",
              error: "No se definieron fórmulas para calcular"
            });
          }
        } catch (error) {
          const errorMsg = error instanceof Error ? error.message : "Error desconocido";
          calcResults.push(`[ERROR en cálculo: ${errorMsg}]`);

          calcActionResults.push({
            success: false,
            action: calcAction,
            message: "Error",
            error: errorMsg
          });
          
          // Mostrar error al usuario pero continuar
          showToast(`Error en cálculo: ${errorMsg}`, "error");
        }
      }

      // Actualizar la tarjeta con los resultados
      const lastMessage = el.messagesContainer().querySelector(".message.assistant:last-child .message-bubble");
      if (lastMessage) {
        const actionsCard = lastMessage.querySelector(".actions-card");
        if (actionsCard) {
          actionsCard.outerHTML = createActionsCard(calcActions, calcActionResults);
          setupActionRowListeners(lastMessage as HTMLElement);
        }
      }

      // Limpiar hoja de cálculos (ignorar errores)
      try {
        await excelService.clearCalcSheet();
      } catch {
        // Ignorar error de limpieza
      }

      // Detectar si hay errores de Excel que necesitan corrección
      const hasExcelErrors = calcResults.some(r => r.includes("ERRORES EN FÓRMULAS") || r.includes("#SPILL!"));

      // Si hay errores de Excel, pedir al modelo que corrija las fórmulas
      if (hasExcelErrors) {
        addMessage("assistant", "⚠️ Se detectaron errores en las fórmulas. Recalculando con estrategia alternativa...");
        
        const retryPrompt = `[ERROR EN CÁLCULOS - NECESITA CORRECCIÓN]
${calcResults.join("\n\n")}

PROBLEMA: Las fórmulas anteriores fallaron. Errores comunes:
- #SPILL!: La fórmula UNIQUE intenta devolver demasiados valores únicos (hay muchas zonas)
- #VALUE!: Tipos de datos incompatibles

SOLUCIÓN REQUERIDA: 
En lugar de usar UNIQUE para obtener zonas y luego contar, usa una estrategia diferente:

OPCIÓN 1 (RECOMENDADA): Si conoces las zonas principales del negocio, usa COUNTIFS directos:
=COUNTIFS(columna_zona, "NOMBRE_ZONA_1", columna_status, "ANULADO")
=COUNTIFS(columna_zona, "NOMBRE_ZONA_2", columna_status, "ANULADO")
... etc para cada zona

OPCIÓN 2: Contar el TOTAL de contratos anulados:
=COUNTIF(columna_status, "ANULADO")

OPCIÓN 3: Si el usuario lo acepta, crear una Tabla Dinámica (pivot) que mostrará todas las zonas automáticamente.

Genera NUEVAS fórmulas para responder: "${userMessage}"
Usa la acción "calc" con fórmulas corregidas. NO uses UNIQUE si puede haber muchos valores únicos.`;

        // Obtener fórmulas corregidas
        const retryResponse = await azureOpenAIService.sendMessageStructured(
          retryPrompt,
          context || undefined
        );
        
        // Si hay nuevas acciones calc, procesarlas
        if (retryResponse.actions && retryResponse.actions.some(a => a.type === "calc")) {
          const retryCalcActions = retryResponse.actions.filter(a => a.type === "calc");
          addMessage("assistant", "🔄 Reintentando con fórmulas corregidas...", retryCalcActions);
          
          // Ejecutar las fórmulas corregidas
          const retryResults: string[] = [];
          const retryActionResults: ActionResult[] = [];
          
          for (const retryAction of retryCalcActions) {
            try {
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const formulas = retryAction.calcFormulas || (retryAction as any).formulas;
              if (formulas && formulas.length > 0) {
                const results = await excelService.executeCalcFormulas(formulas, sheetName);
                
                // Verificar si hay errores en el reintento
                const retryHasErrors = results.some(r => 
                  r.result != null && typeof r.result === "string" && 
                  ["#SPILL!", "#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#N/A"].some(err => String(r.result).includes(err))
                );
                
                if (retryHasErrors) {
                  const formattedResults = results.map(r => `${r.formula} = ${r.result}`).join("\n");
                  retryResults.push(`[CÁLCULOS CON ERRORES]\n${formattedResults}`);
                  retryActionResults.push({
                    success: false,
                    action: retryAction,
                    message: "Aún hay errores",
                    error: "Las fórmulas corregidas también fallaron"
                  });
                } else {
                  const formattedResults = results.map(r => `${r.formula} = ${r.result}`).join("\n");
                  retryResults.push(`[CÁLCULOS CORREGIDOS]\n${formattedResults}`);
                  retryActionResults.push({
                    success: true,
                    action: retryAction,
                    message: `Calculado: ${results.length} fórmulas`
                  });
                }
              }
            } catch (e) {
              retryResults.push(`[ERROR en reintento: ${e instanceof Error ? e.message : "Error"}]`);
              retryActionResults.push({
                success: false,
                action: retryAction,
                message: "Error",
                error: e instanceof Error ? e.message : "Error desconocido"
              });
            }
          }
          
          // Actualizar la tarjeta del reintento con los resultados
          const retryMessage = el.messagesContainer().querySelector(".message.assistant:last-child .message-bubble");
          if (retryMessage) {
            const retryActionsCard = retryMessage.querySelector(".actions-card");
            if (retryActionsCard) {
              retryActionsCard.outerHTML = createActionsCard(retryCalcActions, retryActionResults);
              setupActionRowListeners(retryMessage as HTMLElement);
            }
          }
          
          // Limpiar de nuevo
          try { await excelService.clearCalcSheet(); } catch (e) { /* ignorar */ }
          
          // Usar los resultados corregidos
          calcResults.length = 0; // Limpiar resultados anteriores
          calcResults.push(...retryResults);
        } else if (retryResponse.actions && retryResponse.actions.some(a => a.type === "pivotTable")) {
          // El modelo sugirió usar tabla dinámica
          addMessage("assistant", retryResponse.message || "Sugiero crear una Tabla Dinámica para ver los datos por zona. ¿Deseas que la cree?", retryResponse.actions);
          setLoading(false);
          return;
        }
      }

      // Si no hay resultados válidos, informar al usuario
      if (calcResults.length === 0 || calcResults.every(r => r.includes("ERROR"))) {
        addMessage("assistant", "❌ No se pudieron ejecutar los cálculos. Puede que la columna solicitada no exista en los datos.");
        setLoading(false);
        return;
      }

      // Construir prompt de seguimiento con los resultados
      const followUpPrompt = `[RESULTADOS DE CÁLCULOS - DATOS 100% REALES]
${calcResults.join("\n\n")}

Ahora responde la pregunta original del usuario: "${userMessage}"
Usa SOLO los datos calculados arriba. NO inventes cifras.

${otherActions.length > 0 ? `Nota: También habías propuesto estas acciones: ${JSON.stringify(otherActions)}` : ""}`;

      // Obtener respuesta final con los datos calculados
      response = await azureOpenAIService.sendMessageStructured(
        followUpPrompt,
        context || undefined
      );
    }

    // Verificar si hay acciones de lectura (read)
    if (response.actions && response.actions.some(a => a.type === "read")) {
      // Procesar lecturas primero
      const readActions = response.actions.filter(a => a.type === "read");
      const otherActions = response.actions.filter(a => a.type !== "read");

      // Mostrar mensaje indicando que está leyendo
      addMessage("assistant", `📖 Leyendo datos: ${readActions.map(r => r.range).join(", ")}...`);

      // Ejecutar lecturas
      const readResults: string[] = [];
      for (const readAction of readActions) {
        try {
          const readData = await excelService.readRange(readAction.range, readAction.sheetName);
          // Formatear resultado similar a Claude for Sheets
          const cellsPreview = Object.entries(readData.cells)
            .slice(0, 100) // Limitar a 100 celdas para no sobrecargar
            .map(([addr, val]) => `"${addr}": ${JSON.stringify(val)}`)
            .join(", ");

          readResults.push(`[LECTURA: ${readAction.range} en "${readData.sheetName}"]
Dimensión: ${readData.dimension}
Celdas: {${cellsPreview}${Object.keys(readData.cells).length > 100 ? ", ..." : ""}}`);
        } catch (error) {
          readResults.push(`[ERROR leyendo ${readAction.range}: ${error instanceof Error ? error.message : "Error desconocido"}]`);
        }
      }

      // Construir prompt de seguimiento con los datos leídos
      const followUpPrompt = `[DATOS LEÍDOS]
${readResults.join("\n\n")}

Basándote en estos datos, ahora genera las acciones necesarias para completar la tarea original: "${userMessage}"

${otherActions.length > 0 ? `Nota: Ya habías propuesto estas acciones adicionales que aún son válidas: ${JSON.stringify(otherActions)}` : ""}`;

      // Obtener respuesta final con las acciones
      response = await azureOpenAIService.sendMessageStructured(
        followUpPrompt,
        context || undefined
      );
    }

    // If there are actions, show pending bar first
    if (response.actions && response.actions.length > 0) {
      // Filtrar acciones de lectura ya que fueron procesadas
      const executableActions = response.actions.filter(a => a.type !== "read");

      if (executableActions.length > 0) {
        addMessage("assistant", response.message, executableActions, undefined, webSources, detectedPdfs);
        showPendingActions(executableActions);

        // Execute based on edit mode setting
        if (state.editMode === "auto") {
          // Auto mode: ejecutar directamente sin setTimeout para evitar race conditions
          // Pequeña pausa para que la UI se renderice
          setLoading(false); // Liberar loading temporalmente para la UI
          await new Promise(resolve => setTimeout(resolve, 300));

          // Verificar que aún hay acciones pendientes antes de ejecutar
          if (state.pendingActions.length > 0) {
            await acceptAllActions();
          }
          return; // Salir temprano, acceptAllActions maneja su propio loading
        }
        // If "ask" mode, just show the pending bar and wait for user to click
      } else {
        addMessage("assistant", response.message, undefined, undefined, webSources, detectedPdfs);
      }
    } else {
      addMessage("assistant", response.message, undefined, undefined, webSources, detectedPdfs);
    }

    // Clear context after use
    state.currentExcelContext = null;

    // Clear attached files after sending
    if (state.attachedFiles.length > 0) {
      clearAttachedFiles();
    }

  } catch (error) {
    let message = "Error al comunicarse con el modelo";

    if (error instanceof AzureOpenAIError) {
      message = error.message;
      if (error.statusCode === 401) message = "API Key inválida";
      else if (error.statusCode === 404) message = `Modelo no encontrado`;
      else if (error.statusCode === 429) message = "Límite de solicitudes excedido";
    }

    showToast(message, "error");
    addMessage("assistant", `Lo siento, ocurrió un error: ${message}`);
  } finally {
    setLoading(false);
  }
}

// ===== Clear History =====

function clearHistory(): void {
  azureOpenAIService.clearHistory();
  hidePendingActions();
  clearMessages();
  showToast("Nueva conversación", "info");
}

// ===== Event Listeners =====

function setupEventListeners(): void {
  // Send button
  el.sendBtn().addEventListener("click", sendMessage);

  // Input
  const input = el.userInput();
  input.addEventListener("input", () => {
    autoResizeTextarea();
    updateInputState();
  });

  input.addEventListener("keydown", (e: KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  });

  // Clear history
  el.clearHistoryBtn().addEventListener("click", clearHistory);

  // Model selector
  el.modelSelector().addEventListener("change", handleModelChange);

  // Accept all
  el.acceptAllBtn().addEventListener("click", acceptAllActions);

  // Edit mode toggle
  el.editModeToggle().addEventListener("click", toggleEditMode);

  // More options button (+)
  el.moreOptionsBtn().addEventListener("click", (e) => {
    e.stopPropagation();
    showOptionsPopup();
  });

  // Option buttons and welcome example buttons (delegated event listener)
  el.messagesContainer().addEventListener("click", (e) => {
    const target = e.target as HTMLElement;

    // Option buttons
    if (target.classList.contains("option-btn")) {
      const optionText = target.dataset.optionText;
      if (optionText) {
        handleOptionClick(optionText);
      }
    }
  });
}

/**
 * Maneja el click en un botón de opción
 */
function handleOptionClick(optionText: string): void {
  const input = el.userInput();
  input.value = optionText;
  autoResizeTextarea();
  updateInputState();
  // Enviar automáticamente
  sendMessage();
}

// ===== User Info =====

/**
 * Carga la información del usuario y actualiza la UI
 */
async function loadUserInfo(): Promise<void> {
  try {
    const user = await getUserInfo();
    state.currentUser = user;

    // Actualizar el mensaje de bienvenida si está visible
    const welcomeGreeting = document.querySelector('.welcome-greeting');
    if (welcomeGreeting) {
      welcomeGreeting.textContent = getGreeting(user.firstName);
    }

  } catch {
    // Ignorar errores de carga de usuario
  }
}

// ===== Initialize =====

async function initialize(): Promise<void> {
  try {
    // Esperar a que Office esté listo
    await excelService.waitForOffice();

    // Cargar información del usuario (en paralelo)
    loadUserInfo();

    initializeModelSelector();
    setupEventListeners();
    updateEditModeUI();
    updateWebSearchUI();
    updateRagUI();
    createFileInput(); // Crear input file oculto para adjuntar archivos

    // Selection listener
    excelService.onSelectionChanged(updateSelection);
    await excelService.startSelectionListener();

    updateInputState();

    const validation = validateConfig();
    if (!validation.isValid) {
      setTimeout(() => {
        showToast("Configura tu API Key", "info");
      }, 1000);
    }

    // Auto-indexar datos al iniciar (silenciosamente)
    setTimeout(async () => {
      try {
        await refreshDataIndex(false); // false = sin notificación
      } catch {
        // Ignorar errores de indexación inicial
      }
    }, 1500); // Esperar 1.5s para que Excel esté completamente listo

  } catch {
    // Intentar inicializar la UI aunque Office falle
    try {
      initializeModelSelector();
      setupEventListeners();
      updateEditModeUI();
      updateWebSearchUI();
      updateRagUI();
      updateInputState();
    } catch {
      // Ignorar errores de inicialización de UI
    }
  }
}

// Start - SOLO funciona dentro de Excel
function startApp(): void {
  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady((info: { host: string | null; platform: string | null }) => {
      // Verificar que estamos dentro de Excel
      if (info.host === Office.HostType.Excel) {
        initialize();
      } else {
        // No estamos en Excel - mostrar error
        showOfficeRequiredError();
      }
    });
  } else {
    // Office.js no está disponible - no es un add-in válido
    showOfficeRequiredError();
  }
}

/**
 * Si se accede fuera de Excel, redirigir a la página principal
 */
function showOfficeRequiredError(): void {
  window.location.href = "https://www.chevyplan.com.ec";
}

startApp();
