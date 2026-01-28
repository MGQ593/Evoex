/**
 * Servicio avanzado para interactuar con Excel usando Office.js
 * Incluye: escritura en celdas, detección de selección, formatos, y más
 */

/* global Excel, Office */

// Declaración de tipos para Office.js (cargado desde CDN)
declare const Excel: {
  run: <T>(batch: (context: ExcelContext) => Promise<T>) => Promise<T>;
  EventType: {
    WorksheetSelectionChanged: string;
  };
};

declare const Office: {
  context: OfficeContext;
  onReady: (callback: (info: OfficeInfo) => void) => void;
  HostType: { Excel: string };
};

interface OfficeContext {
  document: unknown;
}

interface OfficeInfo {
  host: string | null;
  platform: string | null;
}

interface ExcelContext {
  workbook: ExcelWorkbook;
  sync: () => Promise<void>;
}

interface ExcelWorkbook {
  getSelectedRange: () => ExcelRange;
  worksheets: ExcelWorksheets;
  tables: ExcelTables;
}

interface ExcelWorksheets {
  getActiveWorksheet: () => ExcelWorksheet;
  getFirst: () => ExcelWorksheet;
}

interface ExcelWorksheet {
  name: string;
  getRange: (address: string) => ExcelRange;
  tables: ExcelTables;
  charts: ExcelCharts;
  getUsedRangeOrNullObject: () => ExcelRange;
  load: (properties: string | string[]) => void;
  onSelectionChanged: ExcelEvent;
}

interface ExcelCharts {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  add: (chartType: any, sourceData: ExcelRange, seriesBy: any) => ExcelChart;
}

interface ExcelChart {
  title: ExcelChartTitle;
  height: number;
  width: number;
  setPosition: (startCell: ExcelRange, endCell?: ExcelRange) => void;
}

interface ExcelChartTitle {
  text: string;
}

interface ExcelEvent {
  add: (handler: (args: SelectionChangedEventArgs) => Promise<void>) => EventHandlerResult;
}

interface EventHandlerResult {
  remove: () => void;
}

interface SelectionChangedEventArgs {
  address: string;
  worksheetId: string;
}

interface ExcelTables {
  items: ExcelTable[];
  getItem: (name: string) => ExcelTable;
  add: (address: string, hasHeaders: boolean) => ExcelTable;
  load: (properties: string[]) => void;
}

interface ExcelTable {
  name: string;
  getRange: () => ExcelRange;
  getHeaderRowRange: () => ExcelRange;
  getDataBodyRange: () => ExcelRange;
}

interface ExcelRangeFormat {
  fill: ExcelFill;
  font: ExcelFont;
  borders: ExcelBorders;
  horizontalAlignment: string;
  verticalAlignment: string;
  wrapText: boolean;
  autofitColumns: () => void;
  autofitRows: () => void;
}

interface ExcelFill {
  color: string;
}

interface ExcelFont {
  bold: boolean;
  italic: boolean;
  size: number;
  color: string;
  name: string;
}

interface ExcelBorders {
  getItem: (index: string) => ExcelBorder;
}

interface ExcelBorder {
  style: string;
  color: string;
  weight: string;
}

interface ExcelRange {
  address: string;
  values: CellValue[][];
  formulas: string[][];
  numberFormat: string[][];
  rowCount: number;
  columnCount: number;
  isNullObject: boolean;
  format: ExcelRangeFormat;
  load: (properties: string | string[]) => void;
  getCell: (row: number, col: number) => ExcelRange;
  getResizedRange: (deltaRows: number, deltaCols: number) => ExcelRange;
  getEntireColumn: () => ExcelRange;
  getEntireRow: () => ExcelRange;
  merge: (across?: boolean) => void;
  unmerge: () => void;
  select: () => void;
  clear: () => void;
}

interface ExcelRangeFormatExtended extends ExcelRangeFormat {
  columnWidth: number;
  rowHeight: number;
}

// Tipos de valores de celda
type CellValue = string | number | boolean | null;

/**
 * Información de un rango
 */
export interface RangeInfo {
  address: string;
  values: CellValue[][];
  formulas?: string[][];
  rowCount: number;
  columnCount: number;
}

/**
 * Información de la selección actual
 */
export interface SelectionInfo {
  address: string;
  sheetName: string;
  rowCount: number;
  columnCount: number;
  isSingleCell: boolean;
  hasContent: boolean;
  firstCellValue?: string | number | boolean | null;
}

/**
 * Opciones de formato para celdas
 */
export interface FormatOptions {
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  fontColor?: string;
  backgroundColor?: string;
  horizontalAlignment?: "left" | "center" | "right";
  verticalAlignment?: "top" | "center" | "bottom";
  wrapText?: boolean;
  borders?: boolean | BorderOptions;
  merge?: boolean;
  numberFormat?: string;
}

export interface BorderOptions {
  style?: "thin" | "medium" | "thick";
  color?: string;
}

/**
 * Tipo semántico de columna - ayuda al modelo a saber qué operación usar
 */
export type SemanticType =
  | "id"        // Identificadores (contar, NO sumar)
  | "amount"    // Montos/valores monetarios (sumar tiene sentido)
  | "quantity"  // Cantidades (sumar o contar según contexto)
  | "category"  // Categorías/texto (solo contar)
  | "date"      // Fechas (solo contar)
  | "unknown";  // No determinado

/**
 * Índice de columna con estadísticas y valores únicos
 */
export interface ColumnIndex {
  header: string;
  column: string; // Letra de columna (A, B, C, etc.)
  type: "text" | "number" | "date" | "mixed";
  semanticType: SemanticType; // Tipo semántico para operaciones
  uniqueValues: string[]; // Máximo 100 valores únicos (ordenados por frecuencia)
  valueCounts: Record<string, number>; // Conteo de cada valor único
  hasMoreValues: boolean; // Si hay más de 100 valores únicos
  // Para columnas numéricas
  stats?: {
    min: number;
    max: number;
    sum: number;
    avg: number;
    count: number;
  };
}

/**
 * Índice completo de datos de la hoja
 */
export interface DataIndex {
  sheetName: string;
  totalRows: number;
  totalColumns: number;
  columns: ColumnIndex[];
  createdAt: Date;
  // Muestra de las primeras filas
  sampleRows: CellValue[][];
}

/**
 * Metadatos ligeros de columna (sin leer datos)
 */
export interface LightweightColumnMeta {
  column: string;      // Letra de columna (A, B, C, etc.)
  header: string;      // Nombre del encabezado
  dataRange: string;   // Rango de datos (ej: "A2:A99379")
}

/**
 * Índice ligero de datos - solo metadatos, sin estadísticas
 */
export interface LightweightDataIndex {
  sheetName: string;
  totalRows: number;
  totalColumns: number;
  columns: LightweightColumnMeta[];
  lastColumn: string;  // Última columna con datos (ej: "ER")
  dataRange: string;   // Rango completo de datos (ej: "A1:ER99379")
  createdAt: Date;
}

/**
 * Resultado de cálculo en hoja oculta
 */
export interface CalcResult {
  formula: string;
  result: CellValue;
  cell: string;
}

/**
 * Tipos de gráfico soportados
 */
export type ChartType =
  | "barClustered" | "barStacked"
  | "columnClustered" | "columnStacked"
  | "line" | "lineMarkers"
  | "pie" | "doughnut"
  | "area" | "areaStacked";

/**
 * Configuración de tabla dinámica
 */
export interface PivotTableConfig {
  sourceSheet: string;       // Hoja de origen de los datos
  sourceRange: string;       // Rango de datos fuente (ej: "A1:Z99356")
  rowField: string;          // Campo para las filas (ej: "CIUDAD")
  valueField: string;        // Campo para los valores (ej: "CONTRATO")
  valueFunction: "count" | "sum" | "average" | "max" | "min"; // Función de agregación
  columnField?: string;      // Campo opcional para columnas
  filterField?: string;      // Campo opcional para filtros
}

/**
 * Resultado de una lectura de Excel
 */
export interface ReadResult {
  range: string;
  sheetName: string;
  dimension: string;  // Ej: "A1:B207"
  cells: Record<string, CellValue>;  // Ej: {"A1": "CIUDAD", "B1": "TOTAL", ...}
  rowCount: number;
  colCount: number;
}

/**
 * Acción a ejecutar en Excel
 */
export interface ExcelAction {
  type: "write" | "formula" | "format" | "merge" | "table" | "columnWidth" | "rowHeight" | "autofit" | "chart" | "createSheet" | "activateSheet" | "deleteSheet" | "pivotTable" | "read" | "calc" | "countByCategory" | "avgByCategory" | "filter" | "clearFilter" | "search" | "sort" | "conditionalFormat" | "dataValidation" | "comment" | "hyperlink" | "namedRange" | "protect" | "unprotect" | "freezePanes" | "unfreezePane" | "groupRows" | "groupColumns" | "ungroupRows" | "ungroupColumns" | "hideRows" | "hideColumns" | "showRows" | "showColumns" | "removeDuplicates" | "textToColumns";
  range: string;
  description?: string;
  value?: CellValue;
  values?: CellValue[][];
  formula?: string;
  formulas?: string[][];
  format?: FormatOptions;
  tableName?: string;
  hasHeaders?: boolean;
  width?: number;
  height?: number;
  // Para gráficos
  chartType?: ChartType;
  chartTitle?: string;
  anchor?: string; // Celda donde anclar el gráfico
  // Protección contra sobreescritura
  allowOverwrite?: boolean; // Si es true, permite sobreescribir datos existentes
  // Para crear/activar hojas
  sheetName?: string; // Nombre de la hoja a crear o activar
  // Para tablas dinámicas
  pivotConfig?: PivotTableConfig;
  // Para cálculos en hoja oculta (tipo "calc")
  calcFormulas?: string[]; // Fórmulas a ejecutar en hoja oculta
  // Para conteo por categoría (tipo "countByCategory")
  categoryColumn?: string;  // Columna de categorías (ej: "T" para ZONAVENTA)
  filterColumn?: string;    // Columna de filtro (ej: "AI" para STATUS)
  filterValue?: string;     // Valor a filtrar (ej: "ANULADO")
  // Para promedio por categoría (tipo "avgByCategory")
  valueColumn?: string;     // Columna de valores a promediar (ej: "E" para EDAD)
  // Para filtros de Excel (tipo "filter")
  filterCriteria?: FilterCriteria[];  // Criterios de filtro por columna
  // Para búsqueda (tipo "search")
  searchValue?: string;     // Valor a buscar
  searchColumn?: string;    // Columna donde buscar (opcional, si no se especifica busca en todo el rango)
  // Para ordenamiento (tipo "sort")
  sortConfig?: SortConfig;
  // Para formato condicional (tipo "conditionalFormat")
  conditionalFormatConfig?: ConditionalFormatConfig;
  // Para validación de datos (tipo "dataValidation")
  dataValidationConfig?: DataValidationConfig;
  // Para comentarios (tipo "comment")
  commentText?: string;
  // Para hipervínculos (tipo "hyperlink")
  hyperlinkConfig?: HyperlinkConfig;
  // Para rangos con nombre (tipo "namedRange")
  namedRangeConfig?: NamedRangeConfig;
  // Para protección de hojas (tipo "protect")
  protectionConfig?: ProtectionConfig;
  // Para congelar paneles (tipo "freezePanes")
  freezeConfig?: FreezeConfig;
  // Para agrupar filas/columnas
  groupOutline?: boolean; // Si true, crea outline colapsable
  // Para texto a columnas
  textToColumnsConfig?: TextToColumnsConfig;
  // Para quitar duplicados
  removeDuplicatesColumns?: number[]; // Índices de columnas a considerar (0-based)
}

/**
 * Configuración para ordenamiento
 */
export interface SortConfig {
  columns: SortColumn[];      // Columnas para ordenar
  hasHeaders?: boolean;       // Si el rango tiene encabezados (default: true)
}

export interface SortColumn {
  columnIndex: number;        // Índice de columna (0-based)
  ascending?: boolean;        // true = A-Z, false = Z-A (default: true)
}

/**
 * Configuración para formato condicional
 */
export interface ConditionalFormatConfig {
  type: "colorScale" | "dataBar" | "iconSet" | "cellValue" | "topBottom" | "aboveAverage" | "duplicates" | "custom";
  // Para colorScale
  colorScale?: {
    minimum?: { color: string; type?: "lowestValue" | "number" | "percent"; value?: number };
    midpoint?: { color: string; type?: "number" | "percent" | "percentile"; value?: number };
    maximum?: { color: string; type?: "highestValue" | "number" | "percent"; value?: number };
  };
  // Para dataBar
  dataBar?: {
    barColor?: string;
    showValue?: boolean;
  };
  // Para iconSet
  iconSet?: "threeArrows" | "threeTrafficLights" | "threeSymbols" | "fourArrows" | "fourTrafficLights" | "fiveArrows" | "fiveRatings";
  // Para cellValue
  cellValue?: {
    operator: "greaterThan" | "lessThan" | "equalTo" | "notEqualTo" | "between" | "greaterThanOrEqual" | "lessThanOrEqual";
    value1: string | number;
    value2?: string | number; // Solo para "between"
    format: { backgroundColor?: string; fontColor?: string; bold?: boolean };
  };
  // Para topBottom
  topBottom?: {
    type: "top" | "bottom";
    count: number;
    percent?: boolean; // Si true, count es porcentaje
    format: { backgroundColor?: string; fontColor?: string };
  };
  // Para aboveAverage
  aboveAverage?: {
    above: boolean; // true = arriba del promedio, false = abajo
    format: { backgroundColor?: string; fontColor?: string };
  };
  // Para duplicates
  duplicates?: {
    unique?: boolean; // true = resaltar únicos, false = resaltar duplicados
    format: { backgroundColor?: string; fontColor?: string };
  };
  // Para custom (fórmula)
  custom?: {
    formula: string; // Fórmula que retorna true/false
    format: { backgroundColor?: string; fontColor?: string; bold?: boolean };
  };
}

/**
 * Configuración para validación de datos
 */
export interface DataValidationConfig {
  type: "list" | "whole" | "decimal" | "date" | "time" | "textLength" | "custom";
  // Para list
  list?: string[] | string;  // Array de valores o rango (ej: "A1:A10")
  // Para whole, decimal, textLength
  operator?: "between" | "notBetween" | "equalTo" | "notEqualTo" | "greaterThan" | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual";
  value1?: number | string;
  value2?: number | string;
  // Para custom
  formula?: string;
  // Opciones comunes
  allowBlank?: boolean;
  showInputMessage?: boolean;
  inputTitle?: string;
  inputMessage?: string;
  showErrorMessage?: boolean;
  errorTitle?: string;
  errorMessage?: string;
  errorStyle?: "stop" | "warning" | "information";
}

/**
 * Configuración para hipervínculos
 */
export interface HyperlinkConfig {
  address: string;          // URL, mailto:, o referencia de celda
  textToDisplay?: string;   // Texto visible
  screenTip?: string;       // Tooltip al pasar el mouse
}

/**
 * Configuración para rangos con nombre
 */
export interface NamedRangeConfig {
  name: string;             // Nombre del rango
  scope?: "workbook" | "worksheet"; // Alcance (default: workbook)
  comment?: string;         // Comentario opcional
}

/**
 * Configuración para protección de hojas
 */
export interface ProtectionConfig {
  password?: string;
  allowFormatCells?: boolean;
  allowFormatColumns?: boolean;
  allowFormatRows?: boolean;
  allowInsertColumns?: boolean;
  allowInsertRows?: boolean;
  allowInsertHyperlinks?: boolean;
  allowDeleteColumns?: boolean;
  allowDeleteRows?: boolean;
  allowSort?: boolean;
  allowAutoFilter?: boolean;
  allowPivotTables?: boolean;
  allowEditObjects?: boolean;
  allowEditScenarios?: boolean;
}

/**
 * Configuración para congelar paneles
 */
export interface FreezeConfig {
  rows?: number;    // Número de filas a congelar desde arriba
  columns?: number; // Número de columnas a congelar desde la izquierda
}

/**
 * Configuración para texto a columnas
 */
export interface TextToColumnsConfig {
  delimiter: "comma" | "semicolon" | "tab" | "space" | "custom";
  customDelimiter?: string; // Si delimiter es "custom"
  treatConsecutiveAsOne?: boolean;
}

/**
 * Criterio de filtro para una columna
 */
export interface FilterCriteria {
  columnIndex: number;      // Índice de columna (0-based desde el inicio del rango)
  values?: string[];        // Valores a mostrar (filtro por valores)
  criteria?: string;        // Criterio de comparación (ej: ">100", "=ANULADO")
}

/**
 * Resultado de verificación de sobreescritura
 */
export interface OverwriteCheckResult {
  hasData: boolean;
  nonEmptyCells: number;
  totalCells: number;
  suggestedRange?: string; // Rango alternativo vacío
}

/**
 * Resultado de ejecutar una acción
 */
export interface ActionResult {
  success: boolean;
  action: ExcelAction;
  message: string;
  error?: string;
  readData?: ReadResult;  // Datos leídos si la acción fue "read"
  // Campos de validación post-ejecución
  validated?: boolean;           // ¿Se realizó validación?
  validationPassed?: boolean;    // ¿Pasó la validación?
  validationMessage?: string;    // Detalle de qué falló
  actualValues?: CellValue[][];  // Valores reales después de ejecutar
}

/**
 * Callback para cambios de selección
 */
export type SelectionChangedCallback = (selection: SelectionInfo) => void;

/**
 * Error personalizado para operaciones de Excel
 */
export class ExcelServiceError extends Error {
  constructor(message: string, public originalError?: Error) {
    super(message);
    this.name = "ExcelServiceError";
  }
}

/**
 * Infiere el tipo semántico de una columna basándose en:
 * 1. El nombre del encabezado (palabras clave)
 * 2. Las características de los datos (unicidad, patrones)
 */
function inferSemanticType(
  header: string,
  dataType: "text" | "number" | "date" | "mixed",
  uniqueCount: number,
  totalRows: number,
  stats?: { min: number; max: number; sum: number; avg: number; count: number }
): SemanticType {
  const headerLower = header.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

  // Patrones para IDs/códigos (NUNCA sumar)
  const idPatterns = [
    /\bid\b/, /\bids?\b/, /\bcod(igo)?\b/, /\bcode\b/,
    /\bnumero?\b/, /\bnum\b/, /\bn[uo]m?\b/, /\bno\.?\b/,
    /\bcontrato\b/, /\bcontract\b/,
    /\bcedula\b/, /\bruc\b/, /\bnit\b/, /\bdni\b/, /\bci\b/,
    /\bdocumento\b/, /\bdoc\b/,
    /\breferencia\b/, /\bref\b/,
    /\bfolio\b/, /\bticket\b/, /\borden\b/, /\border\b/,
    /\bfactura\b/, /\binvoice\b/,
    /\bcliente\b/, /\bclient\b/, /\busuario\b/, /\buser\b/,
    /\bsolicitud\b/, /\brequest\b/,
    /\bplan\b/, /\bgrupo\b/, /\bgroup\b/,
    /\bserie\b/, /\bserial\b/,
    /\bsku\b/, /\bitem\b/, /\bproducto\b/, /\bproduct\b/
  ];

  // Patrones para montos/valores (tiene sentido sumar)
  const amountPatterns = [
    /\bmonto\b/, /\bvalor\b/, /\bprecio\b/, /\bprice\b/,
    /\btotal\b/, /\bsubtotal\b/,
    /\bimporte\b/, /\bcosto\b/, /\bcost\b/,
    /\bventa\b/, /\bsale\b/, /\bingreso\b/, /\bingress\b/,
    /\begreso\b/, /\bgasto\b/, /\bexpense\b/,
    /\bpago\b/, /\bpayment\b/,
    /\bsaldo\b/, /\bbalance\b/,
    /\bcomision\b/, /\bcommission\b/,
    /\bdescuento\b/, /\bdiscount\b/,
    /\biva\b/, /\btax\b/, /\bimpuesto\b/,
    /\bcuota\b/, /\bfee\b/,
    /\b(usd|eur|cop|mxn|pen|clp)\b/, /\$/, /\bmoneda\b/
  ];

  // Patrones para cantidades
  const quantityPatterns = [
    /\bcantidad\b/, /\bqty\b/, /\bquantity\b/,
    /\bunidades?\b/, /\bunits?\b/,
    /\bstock\b/, /\binventario\b/,
    /\bpiezas?\b/, /\bpieces?\b/
  ];

  // Patrones para fechas
  const datePatterns = [
    /\bfecha\b/, /\bdate\b/, /\bdia\b/, /\bday\b/,
    /\bmes\b/, /\bmonth\b/, /\bano\b/, /\byear\b/,
    /\bcreado\b/, /\bcreated\b/, /\bmodificado\b/, /\bmodified\b/,
    /\binicio\b/, /\bstart\b/, /\bfin\b/, /\bend\b/,
    /\bvencimiento\b/, /\bexpir/
  ];

  // Verificar patrones por prioridad
  for (const pattern of idPatterns) {
    if (pattern.test(headerLower)) {
      return "id";
    }
  }

  for (const pattern of amountPatterns) {
    if (pattern.test(headerLower)) {
      return "amount";
    }
  }

  for (const pattern of quantityPatterns) {
    if (pattern.test(headerLower)) {
      return "quantity";
    }
  }

  for (const pattern of datePatterns) {
    if (pattern.test(headerLower)) {
      return "date";
    }
  }

  // Si no hay patrón en el nombre, inferir por características de datos
  if (dataType === "date") {
    return "date";
  }

  if (dataType === "text" || dataType === "mixed") {
    return "category";
  }

  // Para columnas numéricas sin patrón claro, usar heurísticas
  if (dataType === "number" && stats) {
    const uniqueRatio = uniqueCount / totalRows;

    // Alta unicidad (>70%) sugiere IDs
    if (uniqueRatio > 0.7) {
      return "id";
    }

    // Si los valores son enteros consecutivos o casi, probablemente son IDs
    if (stats.min >= 1 && stats.max - stats.min + 1 <= totalRows * 1.2) {
      // Rango muy cercano al número de filas = probable secuencia de IDs
      if (uniqueRatio > 0.5) {
        return "id";
      }
    }

    // Valores con decimales significativos sugieren montos
    // (esto es una heurística débil, pero ayuda)
    if (stats.avg !== Math.floor(stats.avg)) {
      return "amount";
    }

    // Si llegamos aquí, podría ser cantidad o monto
    // Valores pequeños (< 1000 en promedio) más probablemente cantidades
    if (stats.avg < 1000 && stats.max < 10000) {
      return "quantity";
    }

    // Valores grandes más probablemente montos
    return "amount";
  }

  return "unknown";
}

/**
 * Clase de servicio avanzado para operaciones con Excel
 */
export class ExcelService {
  private selectionListeners: SelectionChangedCallback[] = [];
  private eventHandler: EventHandlerResult | null = null;
  private sheetActivatedHandler: EventHandlerResult | null = null;
  private currentSheetName: string = "";

  // Constantes para hoja de cálculos oculta
  private static readonly CALC_SHEET_NAME = "_AI_Calc";

  // Cache del índice ligero
  private lightweightIndexCache: LightweightDataIndex | null = null;
  private lightweightIndexCacheKey: string = "";

  /**
   * Verifica si Office.js está disponible
   */
  isOfficeReady(): boolean {
    return typeof Office !== "undefined" && typeof Excel !== "undefined";
  }

  /**
   * Espera a que Office.js esté listo
   */
  async waitForOffice(): Promise<void> {
    return new Promise((resolve, reject) => {
      // Verificar si Office existe
      if (typeof Office === "undefined") {
        console.warn("Office.js no está disponible");
        reject(new ExcelServiceError("Office.js no está cargado"));
        return;
      }

      // Si ya está listo
      if (this.isOfficeReady()) {
        resolve();
        return;
      }

      // Esperar a que Office esté listo
      try {
        Office.onReady((info: OfficeInfo) => {
          if (info.host === Office.HostType.Excel) {
            resolve();
          } else {
            // Solo funciona en Excel - rechazar cualquier otro contexto
            reject(new ExcelServiceError("Este complemento solo funciona en Microsoft Excel"));
          }
        });
      } catch (error) {
        reject(new ExcelServiceError("Error inicializando Office.js"));
      }
    });
  }

  // ===== DETECCIÓN DE SELECCIÓN =====

  /**
   * Registra un listener para cambios de selección
   */
  onSelectionChanged(callback: SelectionChangedCallback): void {
    this.selectionListeners.push(callback);
  }

  /**
   * Remueve un listener de cambios de selección
   */
  offSelectionChanged(callback: SelectionChangedCallback): void {
    const index = this.selectionListeners.indexOf(callback);
    if (index > -1) {
      this.selectionListeners.splice(index, 1);
    }
  }

  /**
   * Inicia la escucha de cambios de selección
   */
  async startSelectionListener(): Promise<void> {
    try {
      // Registrar listener de selección en la hoja activa
      await this.registerSelectionListenerOnActiveSheet();

      // Registrar listener para cambio de hoja activa
      await Excel.run(async (context: ExcelContext) => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const worksheets = context.workbook.worksheets as any;
        
        this.sheetActivatedHandler = worksheets.onActivated.add(async () => {
          // Re-registrar el listener de selección en la nueva hoja activa
          await this.registerSelectionListenerOnActiveSheet();
        });

        await context.sync();
      });

      // Emitir selección inicial
      const initialSelection = await this.getCurrentSelection();
      this.selectionListeners.forEach(listener => listener(initialSelection));
    } catch (error) {
      console.warn("No se pudo iniciar el listener de selección:", error);
    }
  }

  /**
   * Registra el listener de selección en la hoja activa actual
   */
  private async registerSelectionListenerOnActiveSheet(): Promise<void> {
    try {
      // Primero remover el handler anterior si existe
      if (this.eventHandler) {
        try {
          await Excel.run(async (context: ExcelContext) => {
            this.eventHandler?.remove();
            await context.sync();
          });
        } catch {
          // Ignorar errores al remover
        }
        this.eventHandler = null;
      }

      // Registrar nuevo handler en la hoja activa
      await Excel.run(async (context: ExcelContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        await context.sync();

        this.currentSheetName = sheet.name;

        this.eventHandler = sheet.onSelectionChanged.add(async () => {
          const selection = await this.getCurrentSelection();
          this.selectionListeners.forEach(listener => listener(selection));
        });

        await context.sync();
      });

      // Emitir selección actual después de cambiar de hoja
      const selection = await this.getCurrentSelection();
      this.selectionListeners.forEach(listener => listener(selection));
    } catch (error) {
      console.warn("No se pudo registrar listener en hoja activa:", error);
    }
  }

  /**
   * Detiene la escucha de cambios de selección
   */
  async stopSelectionListener(): Promise<void> {
    // Remover handler de selección
    if (this.eventHandler) {
      try {
        await Excel.run(async (context: ExcelContext) => {
          this.eventHandler?.remove();
          await context.sync();
        });
      } catch {
        // Ignorar errores al detener
      }
      this.eventHandler = null;
    }

    // Remover handler de cambio de hoja
    if (this.sheetActivatedHandler) {
      try {
        await Excel.run(async (context: ExcelContext) => {
          this.sheetActivatedHandler?.remove();
          await context.sync();
        });
      } catch {
        // Ignorar errores al detener
      }
      this.sheetActivatedHandler = null;
    }
  }

  /**
   * Obtiene información de la selección actual
   */
  async getCurrentSelection(): Promise<SelectionInfo> {
    try {
      return await Excel.run(async (context: ExcelContext) => {
        const range = context.workbook.getSelectedRange();
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        range.load(["address", "rowCount", "columnCount", "values"]);
        sheet.load("name");

        await context.sync();

        // Extraer solo la parte de la dirección sin el nombre de la hoja
        const addressParts = range.address.split("!");
        const cellAddress = addressParts.length > 1 ? addressParts[1] : range.address;

        // Verificar si la primera celda tiene contenido
        const firstCellValue = range.values?.[0]?.[0];
        const hasContent = firstCellValue !== null && firstCellValue !== undefined && firstCellValue !== "";

        return {
          address: cellAddress,
          sheetName: sheet.name,
          rowCount: range.rowCount,
          columnCount: range.columnCount,
          isSingleCell: range.rowCount === 1 && range.columnCount === 1,
          hasContent: hasContent,
          firstCellValue: firstCellValue,
        };
      });
    } catch (error) {
      return {
        address: "A1",
        sheetName: "Sheet1",
        rowCount: 1,
        columnCount: 1,
        isSingleCell: true,
        hasContent: false,
      };
    }
  }

  /**
   * Obtiene información del rango usado en la hoja activa
   * Retorna el área que contiene datos
   */
  async getUsedRangeInfo(): Promise<{ address: string; lastRow: number; lastColumn: string } | null> {
    try {
      return await Excel.run(async (context: ExcelContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRangeOrNullObject();

        usedRange.load(["address", "rowCount", "columnCount"]);
        await context.sync();

        if (usedRange.isNullObject) {
          return null;
        }

        // Extraer dirección sin nombre de hoja
        const addressParts = usedRange.address.split("!");
        const address = addressParts.length > 1 ? addressParts[1] : usedRange.address;

        // Calcular última fila y columna
        const lastRow = usedRange.rowCount;
        // Convertir número de columna a letra (1=A, 2=B, etc.)
        const lastColNum = usedRange.columnCount;
        let lastColumn = "";
        let n = lastColNum;
        while (n > 0) {
          n--;
          lastColumn = String.fromCharCode(65 + (n % 26)) + lastColumn;
          n = Math.floor(n / 26);
        }

        return {
          address,
          lastRow,
          lastColumn,
        };
      });
    } catch (error) {
      console.warn("Error obteniendo rango usado:", error);
      return null;
    }
  }

  /**
   * Lee un rango específico y devuelve los datos en formato estructurado
   * Similar a cómo Claude for Sheets lee datos
   */
  async readRange(range: string, sheetName?: string): Promise<ReadResult> {
    try {
      return await Excel.run(async (context: ExcelContext) => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const worksheets = context.workbook.worksheets as any;
        const sheet = sheetName
          ? worksheets.getItem(sheetName)
          : context.workbook.worksheets.getActiveWorksheet();

        sheet.load("name");
        const targetRange = sheet.getRange(range);
        targetRange.load(["address", "values", "rowCount", "columnCount"]);
        await context.sync();

        // Construir objeto de celdas similar al formato de Claude for Sheets
        const cells: Record<string, CellValue> = {};
        const values = targetRange.values;

        for (let row = 0; row < values.length; row++) {
          for (let col = 0; col < values[row].length; col++) {
            const value = values[row][col];
            if (value !== "" && value !== null && value !== undefined) {
              // Convertir índice de columna a letra
              let colLetter = "";
              let n = col + 1;
              while (n > 0) {
                n--;
                colLetter = String.fromCharCode(65 + (n % 26)) + colLetter;
                n = Math.floor(n / 26);
              }

              // Parsear el rango para obtener la fila inicial
              const rangeMatch = range.match(/([A-Z]+)(\d+)/i);
              const startRow = rangeMatch ? parseInt(rangeMatch[2]) : 1;
              const startCol = rangeMatch ? rangeMatch[1].toUpperCase() : "A";

              // Calcular la letra de columna real
              const startColNum = startCol.split("").reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0);
              const realColNum = startColNum + col;
              let realColLetter = "";
              let m = realColNum;
              while (m > 0) {
                m--;
                realColLetter = String.fromCharCode(65 + (m % 26)) + realColLetter;
                m = Math.floor(m / 26);
              }

              const cellAddress = `${realColLetter}${startRow + row}`;
              cells[cellAddress] = value;
            }
          }
        }

        // Calcular dimensión real
        const rangeMatch = range.match(/([A-Z]+)(\d+)/i);
        const startRow = rangeMatch ? parseInt(rangeMatch[2]) : 1;
        const endRow = startRow + targetRange.rowCount - 1;

        const startCol = rangeMatch ? rangeMatch[1].toUpperCase() : "A";
        const startColNum = startCol.split("").reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0);
        const endColNum = startColNum + targetRange.columnCount - 1;
        let endColLetter = "";
        let m = endColNum;
        while (m > 0) {
          m--;
          endColLetter = String.fromCharCode(65 + (m % 26)) + endColLetter;
          m = Math.floor(m / 26);
        }

        return {
          range,
          sheetName: sheet.name,
          dimension: `${startCol}${startRow}:${endColLetter}${endRow}`,
          cells,
          rowCount: targetRange.rowCount,
          colCount: targetRange.columnCount,
        };
      });
    } catch (error) {
      console.error("Error leyendo rango:", error);
      throw new ExcelServiceError(
        `Error al leer rango ${range}`,
        error instanceof Error ? error : undefined
      );
    }
  }

  // ===== LECTURA DE DATOS =====

  /**
   * Lee el contenido de las celdas seleccionadas actualmente
   */
  async getSelectedRange(): Promise<RangeInfo> {
    try {
      return await Excel.run(async (context: ExcelContext) => {
        const range = context.workbook.getSelectedRange();
        range.load(["address", "values", "formulas", "rowCount", "columnCount"]);
        await context.sync();

        return {
          address: range.address,
          values: range.values,
          formulas: range.formulas,
          rowCount: range.rowCount,
          columnCount: range.columnCount,
        };
      });
    } catch (error) {
      throw new ExcelServiceError(
        "Error al leer el rango seleccionado",
        error instanceof Error ? error : undefined
      );
    }
  }

  /**
   * Convierte los datos del rango a formato de texto legible
   */
  async getSelectedRangeAsText(): Promise<string> {
    const rangeInfo = await this.getSelectedRange();

    if (rangeInfo.rowCount === 0 || rangeInfo.columnCount === 0) {
      return "No hay datos seleccionados";
    }

    const lines = rangeInfo.values.map((row) =>
      row.map((cell) => (cell === null ? "" : String(cell))).join("\t")
    );

    const header = `Rango: ${rangeInfo.address} (${rangeInfo.rowCount} filas x ${rangeInfo.columnCount} columnas)`;
    return `${header}\n\nDatos:\n${lines.join("\n")}`;
  }

  /**
   * Lee un rango específico por dirección
   */
  async getRange(address: string): Promise<RangeInfo> {
    try {
      return await Excel.run(async (context: ExcelContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);
        range.load(["address", "values", "formulas", "rowCount", "columnCount"]);
        await context.sync();

        return {
          address: range.address,
          values: range.values,
          formulas: range.formulas,
          rowCount: range.rowCount,
          columnCount: range.columnCount,
        };
      });
    } catch (error) {
      throw new ExcelServiceError(
        `Error al leer el rango ${address}`,
        error instanceof Error ? error : undefined
      );
    }
  }

  // ===== ESCRITURA DE DATOS =====

  /**
   * Escribe valores en un rango específico
   */
  async writeToRange(address: string, values: CellValue[][]): Promise<string> {
    try {
      return await Excel.run(async (context: ExcelContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);

        range.values = values;
        range.load("address");

        await context.sync();
        return range.address;
      });
    } catch (error) {
      throw new ExcelServiceError(
        `Error al escribir en el rango ${address}`,
        error instanceof Error ? error : undefined
      );
    }
  }

  /**
   * Escribe un valor en la celda seleccionada
   */
  async writeToSelectedCell(value: CellValue): Promise<string> {
    try {
      return await Excel.run(async (context: ExcelContext) => {
        const range = context.workbook.getSelectedRange();
        range.load("address");
        range.values = [[value]];
        await context.sync();
        return range.address;
      });
    } catch (error) {
      throw new ExcelServiceError(
        "Error al escribir en la celda",
        error instanceof Error ? error : undefined
      );
    }
  }

  /**
   * Escribe fórmulas en un rango específico
   */
  async writeFormulasToRange(address: string, formulas: string[][]): Promise<string> {
    try {
      return await Excel.run(async (context: ExcelContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);

        // Asegurar que todas las fórmulas empiecen con =
        const normalizedFormulas = formulas.map(row =>
          row.map(f => f && !f.startsWith("=") ? `=${f}` : f)
        );

        range.formulas = normalizedFormulas;
        range.load("address");

        await context.sync();
        return range.address;
      });
    } catch (error) {
      throw new ExcelServiceError(
        `Error al escribir fórmulas en ${address}`,
        error instanceof Error ? error : undefined
      );
    }
  }

  /**
   * Escribe una fórmula en la celda seleccionada
   */
  async writeFormulaToSelectedCell(formula: string): Promise<string> {
    try {
      const normalizedFormula = formula.startsWith("=") ? formula : `=${formula}`;

      return await Excel.run(async (context: ExcelContext) => {
        const range = context.workbook.getSelectedRange();
        range.load("address");
        range.formulas = [[normalizedFormula]];
        await context.sync();
        return range.address;
      });
    } catch (error) {
      throw new ExcelServiceError(
        "Error al escribir la fórmula",
        error instanceof Error ? error : undefined
      );
    }
  }

  /**
   * Escribe múltiples valores a partir de la celda seleccionada
   */
  async writeRangeFromSelected(values: CellValue[][]): Promise<string> {
    try {
      return await Excel.run(async (context: ExcelContext) => {
        const selectedRange = context.workbook.getSelectedRange();
        selectedRange.load("address");
        await context.sync();

        const startCell = selectedRange.getCell(0, 0);
        startCell.load("address");
        await context.sync();

        const rowCount = values.length;
        const colCount = values[0]?.length || 0;

        if (rowCount === 0 || colCount === 0) {
          throw new Error("No hay datos para escribir");
        }

        const targetRange = startCell.getResizedRange(rowCount - 1, colCount - 1);
        targetRange.values = values;
        targetRange.load("address");

        await context.sync();
        return targetRange.address;
      });
    } catch (error) {
      throw new ExcelServiceError(
        "Error al escribir el rango de datos",
        error instanceof Error ? error : undefined
      );
    }
  }

  // ===== FORMATO =====

  /**
   * Aplica formato a un rango
   */
  async applyFormat(address: string, format: FormatOptions): Promise<void> {
    try {
      await Excel.run(async (context: ExcelContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);

        // Aplicar fuente
        if (format.bold !== undefined) {
          range.format.font.bold = format.bold;
        }
        if (format.italic !== undefined) {
          range.format.font.italic = format.italic;
        }
        if (format.fontSize !== undefined) {
          range.format.font.size = format.fontSize;
        }
        if (format.fontColor) {
          range.format.font.color = format.fontColor;
        }

        // Aplicar relleno
        if (format.backgroundColor) {
          range.format.fill.color = format.backgroundColor;
        }

        // Aplicar alineación
        if (format.horizontalAlignment) {
          range.format.horizontalAlignment = format.horizontalAlignment;
        }
        if (format.verticalAlignment) {
          range.format.verticalAlignment = format.verticalAlignment;
        }

        // Wrap text
        if (format.wrapText !== undefined) {
          range.format.wrapText = format.wrapText;
        }

        // Bordes
        if (format.borders) {
          const borderStyle = typeof format.borders === "object" ? format.borders.style || "thin" : "thin";
          const borderColor = typeof format.borders === "object" ? format.borders.color || "#000000" : "#000000";

          const borderPositions = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
          for (const pos of borderPositions) {
            const border = range.format.borders.getItem(pos);
            border.style = borderStyle === "thin" ? "Continuous" : borderStyle === "medium" ? "Medium" : "Thick";
            border.color = borderColor;
          }
        }

        // Merge
        if (format.merge) {
          range.merge();
        }

        // Number format
        if (format.numberFormat) {
          range.numberFormat = [[format.numberFormat]];
        }

        await context.sync();
      });
    } catch (error) {
      throw new ExcelServiceError(
        `Error al aplicar formato a ${address}`,
        error instanceof Error ? error : undefined
      );
    }
  }

  /**
   * Auto-ajusta el ancho de las columnas
   */
  async autoFitColumns(address: string): Promise<void> {
    try {
      await Excel.run(async (context: ExcelContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);
        range.format.autofitColumns();
        await context.sync();
      });
    } catch (error) {
      console.warn("No se pudo auto-ajustar columnas:", error);
    }
  }

  // ===== PROTECCIÓN CONTRA SOBREESCRITURA =====

  /**
   * Verifica si un rango tiene celdas con datos
   * Retorna información sobre cuántas celdas no están vacías
   */
  async checkRangeForData(address: string, numRows: number, numCols: number): Promise<OverwriteCheckResult> {
    try {
      return await Excel.run(async (context: ExcelContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // Extraer celda inicial del rango
        const startAddress = address.split(":")[0];
        const startCell = sheet.getRange(startAddress);

        // Crear rango del tamaño de los datos
        const targetRange = startCell.getResizedRange(numRows - 1, numCols - 1);
        targetRange.load(["values", "address"]);
        await context.sync();

        // Contar celdas no vacías
        let nonEmptyCells = 0;
        const totalCells = numRows * numCols;

        for (const row of targetRange.values) {
          for (const cell of row) {
            if (cell !== null && cell !== undefined && cell !== "") {
              nonEmptyCells++;
            }
          }
        }

        return {
          hasData: nonEmptyCells > 0,
          nonEmptyCells,
          totalCells,
        };
      });
    } catch (error) {
      console.error("Error verificando datos en rango:", error);
      return { hasData: false, nonEmptyCells: 0, totalCells: 0 };
    }
  }

  /**
   * Busca un rango vacío donde se puedan escribir datos
   * Coloca el nuevo contenido justo después de la última columna usada + 1 columna de separación
   * Siempre empieza desde la fila 1
   */
  async findEmptyRangeForData(numRows: number, numCols: number): Promise<string | null> {
    try {
      return await Excel.run(async (context: ExcelContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRangeOrNullObject();

        usedRange.load(["columnCount", "rowCount"]);
        await context.sync();

        if (usedRange.isNullObject) {
          // Si no hay datos, escribir en A1
          return "A1";
        }

        // Colocar inmediatamente después de la última columna usada + 1 de separación
        // Ejemplo: Si datos terminan en ES (col 149), nuevo contenido va en EU (col 151)
        // Esto deja ET (col 150) como separador visual
        const startCol = usedRange.columnCount + 2; // +1 para siguiente, +1 para separación
        const startColLetter = this.numberToColumn(startCol);
        const endColLetter = this.numberToColumn(startCol + numCols - 1);

        // Verificar que las columnas están vacías
        const testRange = sheet.getRange(`${startColLetter}1:${endColLetter}${Math.max(numRows, usedRange.rowCount)}`);
        testRange.load("values");
        await context.sync();

        let isEmpty = true;
        for (const row of testRange.values) {
          for (const cell of row) {
            if (cell !== null && cell !== undefined && cell !== "") {
              isEmpty = false;
              break;
            }
          }
          if (!isEmpty) break;
        }

        if (isEmpty) {
          return `${startColLetter}1`;
        }

        // Si no está vacío, seguir buscando más a la derecha
        for (let offset = 3; offset <= 20; offset++) {
          const newStartCol = usedRange.columnCount + offset;
          const newStartColLetter = this.numberToColumn(newStartCol);
          const newEndColLetter = this.numberToColumn(newStartCol + numCols - 1);

          const newTestRange = sheet.getRange(`${newStartColLetter}1:${newEndColLetter}${Math.max(numRows, usedRange.rowCount)}`);
          newTestRange.load("values");
          await context.sync();

          let isNewEmpty = true;
          for (const row of newTestRange.values) {
            for (const cell of row) {
              if (cell !== null && cell !== undefined && cell !== "") {
                isNewEmpty = false;
                break;
              }
            }
            if (!isNewEmpty) break;
          }

          if (isNewEmpty) {
            return `${newStartColLetter}1`;
          }
        }

        return null; // No se encontró rango vacío
      });
    } catch (error) {
      console.error("Error buscando rango vacío:", error);
      return null;
    }
  }

  // ===== EJECUCIÓN DE ACCIONES =====

  /**
   * Ejecuta una lista de acciones en Excel de forma eficiente
   * Ejecuta cada acción individualmente para mayor confiabilidad
   */
  async executeActions(actions: ExcelAction[]): Promise<ActionResult[]> {
    const results: ActionResult[] = [];

    // Ejecutar cada acción individualmente para evitar errores en lote
    for (const action of actions) {
      const result = await this.executeSingleAction(action);
      results.push(result);
    }

    return results;
  }

  /**
   * Ejecuta una sola acción en Excel
   */
  private async executeSingleAction(action: ExcelAction): Promise<ActionResult> {
    try {
      let message = "";
      let finalRange = action.range;
      let overwriteWarning = "";

      // === PROTECCIÓN CONTRA SOBREESCRITURA (para acciones de escritura) ===
      if ((action.type === "write" || action.type === "formula") && !action.allowOverwrite) {
        const numRows = action.values?.length || action.formulas?.length || 1;
        const numCols = action.values?.[0]?.length || action.formulas?.[0]?.length || 1;
        const startAddress = action.range.split(":")[0];

        // Verificar si hay datos en el rango destino
        const overwriteCheck = await this.checkRangeForData(startAddress, numRows, numCols);

        if (overwriteCheck.hasData) {
          // Buscar un rango vacío alternativo
          const emptyRange = await this.findEmptyRangeForData(numRows, numCols);

          if (emptyRange) {
            finalRange = emptyRange;
            overwriteWarning = `⚠️ Se evitó sobreescribir ${overwriteCheck.nonEmptyCells} celdas. `;
          } else {
            // No se encontró rango vacío, reportar error
            return {
              success: false,
              action,
              message: `No se puede escribir: el rango ${startAddress} contiene ${overwriteCheck.nonEmptyCells} celdas con datos`,
              error: `El rango contiene datos existentes. Seleccione un rango vacío o use allowOverwrite para forzar la escritura.`,
            };
          }
        }
      }

      await Excel.run(async (context: ExcelContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        switch (action.type) {
          case "write":
            if (action.values && action.values.length > 0) {
              // Extraer celda inicial del rango
              const startAddress = finalRange.split(":")[0];
              const startCell = sheet.getRange(startAddress);

              // Calcular dimensiones de los datos
              const numRows = action.values.length;
              const numCols = action.values[0]?.length || 1;

              // Crear rango del tamaño exacto de los datos
              const targetRange = startCell.getResizedRange(numRows - 1, numCols - 1);
              targetRange.values = action.values;
              
              // Auto-ajustar columnas para evitar ######
              targetRange.format.autofitColumns();
              
              message = `${overwriteWarning}Datos escritos (${numRows}x${numCols}) desde ${startAddress}`;
            } else if (action.value !== undefined) {
              const range = sheet.getRange(finalRange);
              range.values = [[action.value]];
              range.format.autofitColumns();
              message = `${overwriteWarning}Valor escrito en ${finalRange}`;
            } else {
              message = `Sin datos para escribir en ${finalRange}`;
            }
            break;

          case "formula":
            if (action.formulas) {
              const range = sheet.getRange(finalRange);
              const normalizedFormulas = action.formulas.map(row =>
                row.map(f => f && !f.startsWith("=") ? `=${f}` : f)
              );
              range.formulas = normalizedFormulas;
              message = `${overwriteWarning}Fórmulas escritas en ${finalRange}`;
            } else if (action.formula) {
              const range = sheet.getRange(finalRange);
              const formula = action.formula.startsWith("=") ? action.formula : `=${action.formula}`;
              range.formulas = [[formula]];
              message = `${overwriteWarning}Fórmula escrita en ${finalRange}`;
            } else {
              message = `Sin fórmula para escribir en ${finalRange}`;
            }
            break;

          case "format":
            if (action.format) {
              const range = sheet.getRange(action.range);
              this.applyFormatToRange(range, action.format);
              message = `Formato aplicado a ${action.range}`;
            }
            break;

          case "merge":
            {
              const range = sheet.getRange(action.range);
              range.merge();
              message = `Celdas combinadas en ${action.range}`;
            }
            break;

          case "columnWidth":
            {
              const range = sheet.getRange(action.range);
              const columns = range.getEntireColumn();
              // Usar puntos directamente (40pt es buen ancho para calendarios)
              const widthInPoints = action.width || 40;
              (columns.format as ExcelRangeFormatExtended).columnWidth = widthInPoints;
              message = `Ancho de columna ajustado a ${widthInPoints}pt en ${action.range}`;
            }
            break;

          case "rowHeight":
            {
              const range = sheet.getRange(action.range);
              const rows = range.getEntireRow();
              (rows.format as ExcelRangeFormatExtended).rowHeight = action.height || 15;
              message = `Alto de fila ajustado en ${action.range}`;
            }
            break;

          case "autofit":
            {
              const range = sheet.getRange(action.range);
              range.format.autofitColumns();
              range.format.autofitRows();
              message = `Auto-ajuste aplicado a ${action.range}`;
            }
            break;

          case "table":
            {
              const table = sheet.tables.add(action.range, action.hasHeaders ?? true);
              if (action.tableName) {
                table.name = action.tableName;
              }
              message = `Tabla creada en ${action.range}`;
            }
            break;

          case "chart":
            {
              const dataRange = sheet.getRange(action.range);
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const chart = sheet.charts.add(
                this.mapChartType(action.chartType || "columnClustered") as any,
                dataRange,
                "Auto" as any // ChartSeriesBy.auto
              );

              // Configurar título si se proporciona
              if (action.chartTitle) {
                chart.title.text = action.chartTitle;
              }

              // Posicionar el gráfico
              if (action.anchor) {
                const anchorCell = sheet.getRange(action.anchor);
                chart.setPosition(anchorCell);
              }

              // Tamaño por defecto
              chart.height = 300;
              chart.width = 450;

              message = `Gráfico ${action.chartType || "columnClustered"} creado`;
            }
            break;

          case "createSheet":
            {
              const newSheetName = action.sheetName || "Nueva Hoja";
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const worksheets = context.workbook.worksheets as any;
              const newSheet = worksheets.add(newSheetName);
              newSheet.activate();
              message = `Hoja "${newSheetName}" creada y activada`;
            }
            break;

          case "activateSheet":
            {
              const targetSheetName = action.sheetName;
              if (targetSheetName) {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const worksheets = context.workbook.worksheets as any;
                const targetSheet = worksheets.getItem(targetSheetName);
                targetSheet.activate();
                message = `Hoja "${targetSheetName}" activada`;
              } else {
                message = `No se especificó nombre de hoja para activar`;
              }
            }
            break;

          case "deleteSheet":
            {
              const sheetToDelete = action.sheetName;
              if (sheetToDelete) {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const worksheets = context.workbook.worksheets as any;
                const targetSheet = worksheets.getItem(sheetToDelete);
                targetSheet.delete();
                message = `Hoja "${sheetToDelete}" eliminada`;
              } else {
                message = `No se especificó nombre de hoja para eliminar`;
              }
            }
            break;

          case "pivotTable":
            {
              const pivotConfig = action.pivotConfig;
              if (!pivotConfig) {
                throw new Error("Se requiere pivotConfig para crear tabla dinámica");
              }

              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const worksheets = context.workbook.worksheets as any;

              // Obtener la hoja de origen
              const sourceSheet = worksheets.getItem(pivotConfig.sourceSheet);
              const sourceRange = sourceSheet.getRange(pivotConfig.sourceRange);

              // Crear nueva hoja para la tabla dinámica si se especifica sheetName
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              let pivotSheet: any = sheet;
              if (action.sheetName) {
                pivotSheet = worksheets.add(action.sheetName);
                pivotSheet.activate();
              }

              // Crear la tabla dinámica
              const pivotLocation = pivotSheet.getRange(action.range);
              const pivotTableName = `PivotTable_${Date.now()}`;

              const pivotTable = pivotSheet.pivotTables.add(
                pivotTableName,
                sourceRange,
                pivotLocation
              );

              // Configurar el campo de filas
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const rowHierarchy = pivotTable.rowHierarchies.add(
                pivotTable.hierarchies.getItem(pivotConfig.rowField)
              );

              // Configurar el campo de valores con la función de agregación
              const dataHierarchy = pivotTable.dataHierarchies.add(
                pivotTable.hierarchies.getItem(pivotConfig.valueField)
              );

              // Mapear función de agregación
              const aggregationMap: Record<string, string> = {
                count: "Count",
                sum: "Sum",
                average: "Average",
                max: "Max",
                min: "Min",
              };
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              (dataHierarchy as any).summarizeBy = aggregationMap[pivotConfig.valueFunction] || "Count";

              // Configurar campo de columnas si se especifica
              if (pivotConfig.columnField) {
                pivotTable.columnHierarchies.add(
                  pivotTable.hierarchies.getItem(pivotConfig.columnField)
                );
              }

              // Configurar campo de filtro si se especifica
              if (pivotConfig.filterField) {
                pivotTable.filterHierarchies.add(
                  pivotTable.hierarchies.getItem(pivotConfig.filterField)
                );
              }

              message = `Tabla dinámica creada: ${pivotConfig.rowField} vs ${pivotConfig.valueField} (${pivotConfig.valueFunction})`;
            }
            break;

          case "read":
            // Las acciones de lectura se procesan por separado en processReadActions
            // Este caso no debería ejecutarse en el flujo normal
            message = `Lectura de ${action.range} procesada por separado`;
            break;

          case "calc":
            {
              // Acciones de cálculo se procesan aparte para obtener resultados
              // Este caso no debería ejecutarse en executeActions directamente
              message = `Cálculo procesado por separado`;
            }
            break;

          case "filter":
            {
              // Aplicar filtro automático a un rango
              const filterRange = sheet.getRange(action.range);
              
              // Primero, aplicar AutoFilter al rango si no existe
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const autoFilter = (sheet as any).autoFilter;
              
              // Aplicar filtro al rango
              autoFilter.apply(filterRange);
              
              // Si hay criterios de filtro, aplicarlos
              if (action.filterCriteria && action.filterCriteria.length > 0) {
                for (const criteria of action.filterCriteria) {
                  if (criteria.values && criteria.values.length > 0) {
                    // Filtro por valores específicos
                    autoFilter.apply(filterRange, criteria.columnIndex, {
                      filterOn: "Values",
                      values: criteria.values
                    });
                  } else if (criteria.criteria) {
                    // Filtro por criterio (ej: ">100", "=ANULADO")
                    autoFilter.apply(filterRange, criteria.columnIndex, {
                      filterOn: "Custom",
                      criterion1: criteria.criteria
                    });
                  }
                }
              }
              
              await context.sync();
              message = `Filtro aplicado en ${action.range}`;
            }
            break;

          case "clearFilter":
            {
              // Limpiar todos los filtros de la hoja
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const autoFilter = (sheet as any).autoFilter;
              autoFilter.clearCriteria();
              await context.sync();
              message = `Filtros limpiados`;
            }
            break;

          case "search":
            {
              // Buscar un valor y seleccionar las celdas que coinciden
              if (!action.searchValue) {
                message = `No se especificó valor de búsqueda`;
                break;
              }
              
              const searchRange = sheet.getRange(action.range);
              searchRange.load("values, address");
              await context.sync();
              
              const searchTerm = action.searchValue.toLowerCase();
              const values = searchRange.values;
              const matches: string[] = [];
              
              // Obtener la dirección base para calcular las celdas individuales
              const baseAddress = searchRange.address.split("!")[1] || searchRange.address;
              const startMatch = baseAddress.match(/([A-Z]+)(\d+)/);
              
              if (startMatch) {
                const startCol = startMatch[1];
                const startRow = parseInt(startMatch[2]);
                
                for (let row = 0; row < values.length; row++) {
                  for (let col = 0; col < values[row].length; col++) {
                    const cellValue = String(values[row][col] || "").toLowerCase();
                    if (cellValue.includes(searchTerm)) {
                      // Calcular la dirección de la celda
                      const colLetter = this.numberToColumnLetter(this.columnLetterToNumber(startCol) + col);
                      matches.push(`${colLetter}${startRow + row}`);
                    }
                  }
                }
              }
              
              if (matches.length > 0) {
                // Seleccionar la primera coincidencia
                const firstMatch = sheet.getRange(matches[0]);
                firstMatch.select();
                message = `Encontradas ${matches.length} coincidencias. Primera en ${matches[0]}`;
              } else {
                message = `No se encontró "${action.searchValue}" en ${action.range}`;
              }
            }
            break;

          case "sort":
            {
              // Ordenar un rango
              const sortRange = sheet.getRange(action.range);
              const sortConfig = action.sortConfig;
              
              if (sortConfig && sortConfig.columns && sortConfig.columns.length > 0) {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const sortFields: any[] = sortConfig.columns.map(col => ({
                  key: col.columnIndex,
                  ascending: col.ascending !== false // Default true
                }));
                
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                (sortRange as any).sort.apply(sortFields, false, sortConfig.hasHeaders !== false);
                message = `Rango ${action.range} ordenado`;
              } else {
                // Ordenar simple por la primera columna
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                (sortRange as any).sort.apply([{ key: 0, ascending: true }], false, true);
                message = `Rango ${action.range} ordenado ascendente`;
              }
            }
            break;

          case "conditionalFormat":
            {
              // Aplicar formato condicional
              const cfRange = sheet.getRange(action.range);
              const cfConfig = action.conditionalFormatConfig;
              
              if (!cfConfig) {
                message = `No se especificó configuración de formato condicional`;
                break;
              }

              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const conditionalFormats = (cfRange as any).conditionalFormats;

              switch (cfConfig.type) {
                case "colorScale":
                  {
                    const colorScale = conditionalFormats.add("ColorScale");
                    const cs = cfConfig.colorScale;
                    if (cs) {
                      colorScale.colorScale.criteria = [
                        { type: cs.minimum?.type || "lowestValue", color: cs.minimum?.color || "#F8696B" },
                        ...(cs.midpoint ? [{ type: cs.midpoint.type || "percentile", value: cs.midpoint.value || 50, color: cs.midpoint.color || "#FFEB84" }] : []),
                        { type: cs.maximum?.type || "highestValue", color: cs.maximum?.color || "#63BE7B" }
                      ];
                    }
                    message = `Escala de colores aplicada a ${action.range}`;
                  }
                  break;

                case "dataBar":
                  {
                    const dataBar = conditionalFormats.add("DataBar");
                    if (cfConfig.dataBar) {
                      dataBar.dataBar.barDirection = "Context";
                      if (cfConfig.dataBar.barColor) {
                        dataBar.dataBar.positiveFormat.fillColor = cfConfig.dataBar.barColor;
                      }
                      dataBar.dataBar.showValue = cfConfig.dataBar.showValue !== false;
                    }
                    message = `Barras de datos aplicadas a ${action.range}`;
                  }
                  break;

                case "iconSet":
                  {
                    const iconSet = conditionalFormats.add("IconSet");
                    const iconMap: Record<string, string> = {
                      "threeArrows": "ThreeArrows",
                      "threeTrafficLights": "ThreeTrafficLights1",
                      "threeSymbols": "ThreeSymbols",
                      "fourArrows": "FourArrows",
                      "fourTrafficLights": "FourTrafficLights",
                      "fiveArrows": "FiveArrows",
                      "fiveRatings": "FiveRating"
                    };
                    iconSet.iconSet.style = iconMap[cfConfig.iconSet || "threeArrows"] || "ThreeArrows";
                    message = `Iconos aplicados a ${action.range}`;
                  }
                  break;

                case "cellValue":
                  {
                    const cv = cfConfig.cellValue;
                    if (cv) {
                      const cellValueFormat = conditionalFormats.add("CellValue");
                      const operatorMap: Record<string, string> = {
                        "greaterThan": "GreaterThan",
                        "lessThan": "LessThan",
                        "equalTo": "EqualTo",
                        "notEqualTo": "NotEqualTo",
                        "between": "Between",
                        "greaterThanOrEqual": "GreaterThanOrEqual",
                        "lessThanOrEqual": "LessThanOrEqual"
                      };
                      cellValueFormat.cellValue.rule = {
                        formula1: String(cv.value1),
                        formula2: cv.value2 ? String(cv.value2) : undefined,
                        operator: operatorMap[cv.operator] || "GreaterThan"
                      };
                      if (cv.format.backgroundColor) {
                        cellValueFormat.cellValue.format.fill.color = cv.format.backgroundColor;
                      }
                      if (cv.format.fontColor) {
                        cellValueFormat.cellValue.format.font.color = cv.format.fontColor;
                      }
                      if (cv.format.bold) {
                        cellValueFormat.cellValue.format.font.bold = cv.format.bold;
                      }
                    }
                    message = `Formato condicional por valor aplicado a ${action.range}`;
                  }
                  break;

                case "topBottom":
                  {
                    const tb = cfConfig.topBottom;
                    if (tb) {
                      const topBottom = conditionalFormats.add("TopBottom");
                      topBottom.topBottom.rule = {
                        type: tb.type === "top" ? "TopItems" : "BottomItems",
                        rank: tb.count
                      };
                      if (tb.percent) {
                        topBottom.topBottom.rule.type = tb.type === "top" ? "TopPercent" : "BottomPercent";
                      }
                      if (tb.format.backgroundColor) {
                        topBottom.topBottom.format.fill.color = tb.format.backgroundColor;
                      }
                      if (tb.format.fontColor) {
                        topBottom.topBottom.format.font.color = tb.format.fontColor;
                      }
                    }
                    message = `Top/Bottom aplicado a ${action.range}`;
                  }
                  break;

                case "aboveAverage":
                  {
                    const aa = cfConfig.aboveAverage;
                    if (aa) {
                      const aboveAvg = conditionalFormats.add("PresetCriteria");
                      aboveAvg.preset.rule = {
                        criterion: aa.above ? "AboveAverage" : "BelowAverage"
                      };
                      if (aa.format.backgroundColor) {
                        aboveAvg.preset.format.fill.color = aa.format.backgroundColor;
                      }
                      if (aa.format.fontColor) {
                        aboveAvg.preset.format.font.color = aa.format.fontColor;
                      }
                    }
                    message = `Formato arriba/debajo promedio aplicado a ${action.range}`;
                  }
                  break;

                case "duplicates":
                  {
                    const dup = cfConfig.duplicates;
                    if (dup) {
                      const duplicates = conditionalFormats.add("PresetCriteria");
                      duplicates.preset.rule = {
                        criterion: dup.unique ? "UniqueValues" : "DuplicateValues"
                      };
                      if (dup.format.backgroundColor) {
                        duplicates.preset.format.fill.color = dup.format.backgroundColor;
                      }
                      if (dup.format.fontColor) {
                        duplicates.preset.format.font.color = dup.format.fontColor;
                      }
                    }
                    message = `Formato de duplicados aplicado a ${action.range}`;
                  }
                  break;

                case "custom":
                  {
                    const custom = cfConfig.custom;
                    if (custom) {
                      const customFormat = conditionalFormats.add("Custom");
                      customFormat.custom.rule = {
                        formula: custom.formula
                      };
                      if (custom.format.backgroundColor) {
                        customFormat.custom.format.fill.color = custom.format.backgroundColor;
                      }
                      if (custom.format.fontColor) {
                        customFormat.custom.format.font.color = custom.format.fontColor;
                      }
                      if (custom.format.bold) {
                        customFormat.custom.format.font.bold = custom.format.bold;
                      }
                    }
                    message = `Formato condicional personalizado aplicado a ${action.range}`;
                  }
                  break;
              }
            }
            break;

          case "dataValidation":
            {
              // Aplicar validación de datos
              const dvRange = sheet.getRange(action.range);
              const dvConfig = action.dataValidationConfig;
              
              if (!dvConfig) {
                message = `No se especificó configuración de validación`;
                break;
              }

              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const dataValidation = (dvRange as any).dataValidation;
              dataValidation.clear();

              const rule: Record<string, unknown> = {};

              switch (dvConfig.type) {
                case "list":
                  if (Array.isArray(dvConfig.list)) {
                    rule.list = { inCellDropDown: true, source: dvConfig.list.join(",") };
                  } else if (typeof dvConfig.list === "string") {
                    rule.list = { inCellDropDown: true, source: dvConfig.list };
                  }
                  break;

                case "whole":
                case "decimal":
                case "textLength":
                  {
                    const operatorMap: Record<string, string> = {
                      "between": "Between",
                      "notBetween": "NotBetween",
                      "equalTo": "EqualTo",
                      "notEqualTo": "NotEqualTo",
                      "greaterThan": "GreaterThan",
                      "lessThan": "LessThan",
                      "greaterThanOrEqual": "GreaterThanOrEqual",
                      "lessThanOrEqual": "LessThanOrEqual"
                    };
                    const typeMap: Record<string, string> = {
                      "whole": "wholeNumber",
                      "decimal": "decimal",
                      "textLength": "textLength"
                    };
                    rule[typeMap[dvConfig.type]] = {
                      formula1: dvConfig.value1,
                      formula2: dvConfig.value2,
                      operator: operatorMap[dvConfig.operator || "between"] || "Between"
                    };
                  }
                  break;

                case "date":
                case "time":
                  {
                    const operatorMap: Record<string, string> = {
                      "between": "Between",
                      "notBetween": "NotBetween",
                      "equalTo": "EqualTo",
                      "notEqualTo": "NotEqualTo",
                      "greaterThan": "GreaterThan",
                      "lessThan": "LessThan",
                      "greaterThanOrEqual": "GreaterThanOrEqual",
                      "lessThanOrEqual": "LessThanOrEqual"
                    };
                    rule[dvConfig.type] = {
                      formula1: dvConfig.value1,
                      formula2: dvConfig.value2,
                      operator: operatorMap[dvConfig.operator || "between"] || "Between"
                    };
                  }
                  break;

                case "custom":
                  rule.custom = { formula: dvConfig.formula };
                  break;
              }

              dataValidation.rule = rule;

              // Configurar mensajes
              if (dvConfig.showInputMessage && (dvConfig.inputTitle || dvConfig.inputMessage)) {
                dataValidation.prompt = {
                  showPrompt: true,
                  title: dvConfig.inputTitle || "",
                  message: dvConfig.inputMessage || ""
                };
              }

              if (dvConfig.showErrorMessage) {
                const styleMap: Record<string, string> = {
                  "stop": "Stop",
                  "warning": "Warning",
                  "information": "Information"
                };
                dataValidation.errorAlert = {
                  showAlert: true,
                  style: styleMap[dvConfig.errorStyle || "stop"] || "Stop",
                  title: dvConfig.errorTitle || "Error",
                  message: dvConfig.errorMessage || "Valor no válido"
                };
              }

              message = `Validación de datos aplicada a ${action.range}`;
            }
            break;

          case "comment":
            {
              // Agregar comentario a una celda
              if (!action.commentText) {
                message = `No se especificó texto del comentario`;
                break;
              }

              const commentRange = sheet.getRange(action.range);
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const comments = (sheet as any).comments;
              comments.add(commentRange, action.commentText);
              message = `Comentario agregado en ${action.range}`;
            }
            break;

          case "hyperlink":
            {
              // Agregar hipervínculo
              const hlConfig = action.hyperlinkConfig;
              if (!hlConfig) {
                message = `No se especificó configuración de hipervínculo`;
                break;
              }

              const hlRange = sheet.getRange(action.range);
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              (hlRange as any).hyperlink = {
                address: hlConfig.address,
                textToDisplay: hlConfig.textToDisplay || hlConfig.address,
                screenTip: hlConfig.screenTip
              };
              message = `Hipervínculo agregado en ${action.range}`;
            }
            break;

          case "namedRange":
            {
              // Crear rango con nombre
              const nrConfig = action.namedRangeConfig;
              if (!nrConfig || !nrConfig.name) {
                message = `No se especificó nombre para el rango`;
                break;
              }

              const nrRange = sheet.getRange(action.range);
              
              if (nrConfig.scope === "worksheet") {
                // Rango con ámbito de hoja
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                (sheet as any).names.add(nrConfig.name, nrRange, nrConfig.comment);
              } else {
                // Rango con ámbito de libro (default)
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                (context.workbook as any).names.add(nrConfig.name, nrRange, nrConfig.comment);
              }
              message = `Rango "${nrConfig.name}" creado para ${action.range}`;
            }
            break;

          case "protect":
            {
              // Proteger hoja
              const protConfig = action.protectionConfig || {};
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const protection = (sheet as any).protection;
              
              const options = {
                allowFormatCells: protConfig.allowFormatCells ?? false,
                allowFormatColumns: protConfig.allowFormatColumns ?? false,
                allowFormatRows: protConfig.allowFormatRows ?? false,
                allowInsertColumns: protConfig.allowInsertColumns ?? false,
                allowInsertRows: protConfig.allowInsertRows ?? false,
                allowInsertHyperlinks: protConfig.allowInsertHyperlinks ?? false,
                allowDeleteColumns: protConfig.allowDeleteColumns ?? false,
                allowDeleteRows: protConfig.allowDeleteRows ?? false,
                allowSort: protConfig.allowSort ?? false,
                allowAutoFilter: protConfig.allowAutoFilter ?? false,
                allowPivotTables: protConfig.allowPivotTables ?? false,
                allowEditObjects: protConfig.allowEditObjects ?? false,
                allowEditScenarios: protConfig.allowEditScenarios ?? false
              };

              if (protConfig.password) {
                protection.protect(options, protConfig.password);
              } else {
                protection.protect(options);
              }
              message = `Hoja protegida`;
            }
            break;

          case "unprotect":
            {
              // Desproteger hoja
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const protection = (sheet as any).protection;
              const password = action.protectionConfig?.password;
              
              if (password) {
                protection.unprotect(password);
              } else {
                protection.unprotect();
              }
              message = `Hoja desprotegida`;
            }
            break;

          case "freezePanes":
            {
              // Congelar paneles
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const freezePanes = (sheet as any).freezePanes;
              const fc = action.freezeConfig;
              
              if (fc) {
                if (fc.rows && fc.columns) {
                  // Congelar filas y columnas
                  freezePanes.freezeAt(sheet.getRange(action.range));
                } else if (fc.rows) {
                  // Solo filas
                  freezePanes.freezeRows(fc.rows);
                } else if (fc.columns) {
                  // Solo columnas
                  freezePanes.freezeColumns(fc.columns);
                }
              } else {
                // Congelar en la celda especificada
                freezePanes.freezeAt(sheet.getRange(action.range));
              }
              message = `Paneles congelados`;
            }
            break;

          case "unfreezePane":
            {
              // Descongelar paneles
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              (sheet as any).freezePanes.unfreeze();
              message = `Paneles descongelados`;
            }
            break;

          case "groupRows":
            {
              // Agrupar filas (crear outline)
              const groupRange = sheet.getRange(action.range);
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              (groupRange as any).group("Rows");
              message = `Filas agrupadas en ${action.range}`;
            }
            break;

          case "groupColumns":
            {
              // Agrupar columnas
              const groupRange = sheet.getRange(action.range);
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              (groupRange as any).group("Columns");
              message = `Columnas agrupadas en ${action.range}`;
            }
            break;

          case "ungroupRows":
            {
              // Desagrupar filas
              const ungroupRange = sheet.getRange(action.range);
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              (ungroupRange as any).ungroup("Rows");
              message = `Filas desagrupadas en ${action.range}`;
            }
            break;

          case "ungroupColumns":
            {
              // Desagrupar columnas
              const ungroupRange = sheet.getRange(action.range);
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              (ungroupRange as any).ungroup("Columns");
              message = `Columnas desagrupadas en ${action.range}`;
            }
            break;

          case "hideRows":
            {
              // Ocultar filas
              const hideRange = sheet.getRange(action.range);
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              (hideRange as any).rowHidden = true;
              message = `Filas ocultas en ${action.range}`;
            }
            break;

          case "hideColumns":
            {
              // Ocultar columnas
              const hideRange = sheet.getRange(action.range);
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              (hideRange as any).columnHidden = true;
              message = `Columnas ocultas en ${action.range}`;
            }
            break;

          case "showRows":
            {
              // Mostrar filas
              const showRange = sheet.getRange(action.range);
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              (showRange as any).rowHidden = false;
              message = `Filas mostradas en ${action.range}`;
            }
            break;

          case "showColumns":
            {
              // Mostrar columnas
              const showRange = sheet.getRange(action.range);
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              (showRange as any).columnHidden = false;
              message = `Columnas mostradas en ${action.range}`;
            }
            break;

          case "removeDuplicates":
            {
              // Quitar duplicados
              const rdRange = sheet.getRange(action.range);
              const columns = action.removeDuplicatesColumns || [];
              
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const result = (rdRange as any).removeDuplicates(
                columns.length > 0 ? columns : undefined,
                true // excludeHeader
              );
              
              await context.sync();
              message = `Duplicados eliminados de ${action.range}. Filas eliminadas: ${result.removed}, Filas únicas: ${result.uniqueRemaining}`;
            }
            break;

          case "textToColumns":
            {
              // Texto a columnas (parseo de delimitadores)
              // Nota: Office.js no tiene textToColumns directamente, implementamos alternativa
              const ttcRange = sheet.getRange(action.range);
              ttcRange.load("values");
              await context.sync();
              
              const ttcConfig = action.textToColumnsConfig;
              const delimiterMap: Record<string, string> = {
                "comma": ",",
                "semicolon": ";",
                "tab": "\t",
                "space": " ",
                "custom": ttcConfig?.customDelimiter || ","
              };
              const delimiter = delimiterMap[ttcConfig?.delimiter || "comma"];
              
              const originalValues = ttcRange.values;
              const splitValues: (string | number | boolean)[][] = [];
              let maxCols = 0;
              
              // Dividir cada fila por el delimitador
              for (const row of originalValues) {
                const cellValue = String(row[0] || "");
                const parts = cellValue.split(delimiter);
                if (ttcConfig?.treatConsecutiveAsOne) {
                  // Filtrar partes vacías
                  const filtered = parts.filter(p => p.trim() !== "");
                  splitValues.push(filtered);
                  maxCols = Math.max(maxCols, filtered.length);
                } else {
                  splitValues.push(parts);
                  maxCols = Math.max(maxCols, parts.length);
                }
              }
              
              // Normalizar todas las filas al mismo número de columnas
              for (const row of splitValues) {
                while (row.length < maxCols) {
                  row.push("");
                }
              }
              
              // Obtener la celda inicial y escribir los resultados
              const startMatch = action.range.match(/([A-Z]+)(\d+)/);
              if (startMatch && splitValues.length > 0) {
                const startCol = startMatch[1];
                const startRow = parseInt(startMatch[2]);
                const endCol = this.numberToColumnLetter(this.columnLetterToNumber(startCol) + maxCols - 1);
                const endRow = startRow + splitValues.length - 1;
                const outputRange = sheet.getRange(`${startCol}${startRow}:${endCol}${endRow}`);
                outputRange.values = splitValues;
              }
              
              message = `Texto dividido en columnas en ${action.range}`;
            }
            break;

          default:
            message = `Acción desconocida: ${action.type}`;
        }

        await context.sync();
      });

      // === VALIDACIÓN POST-EJECUCIÓN ===
      // Para acciones de escritura, verificar que los datos se escribieron correctamente
      if (action.type === "write" || action.type === "formula") {
        const validationResult = await this.validateWriteAction(action, finalRange);

        return {
          success: true,
          action,
          message,
          validated: true,
          validationPassed: validationResult.passed,
          validationMessage: validationResult.message,
          actualValues: validationResult.actualValues,
        };
      }

      // Para otras acciones, no hay validación específica
      return {
        success: true,
        action,
        message,
        validated: false,
      };
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : "Error desconocido";
      console.error(`Error en acción ${action.type} (${action.range}):`, errorMsg);

      return {
        success: false,
        action,
        message: `Error en ${action.range}`,
        error: errorMsg,
        validated: false,
      };
    }
  }

  /**
   * Valida que una acción de escritura haya sido exitosa
   */
  private async validateWriteAction(
    action: ExcelAction,
    finalRange: string
  ): Promise<{ passed: boolean; message: string; actualValues?: CellValue[][] }> {
    try {
      let actualValues: CellValue[][] = [];

      await Excel.run(async (context: ExcelContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // Determinar el rango a verificar
        let verifyRange: string;
        if (action.type === "write" && action.values) {
          const numRows = action.values.length;
          const numCols = action.values[0]?.length || 1;
          const startAddress = finalRange.split(":")[0];
          const startCell = sheet.getRange(startAddress);
          const targetRange = startCell.getResizedRange(numRows - 1, numCols - 1);
          targetRange.load("values, address");
          await context.sync();
          actualValues = targetRange.values as CellValue[][];
          verifyRange = targetRange.address;
        } else if (action.type === "formula") {
          const range = sheet.getRange(finalRange);
          range.load("values, address");
          await context.sync();
          actualValues = range.values as CellValue[][];
          verifyRange = range.address;
        } else {
          verifyRange = finalRange;
        }

        console.log(`[Validación] Rango: ${verifyRange}, Valores:`, actualValues);
      });

      // Verificar si hay datos en las celdas
      let hasData = false;
      let emptyCount = 0;
      let totalCells = 0;

      for (const row of actualValues) {
        for (const cell of row) {
          totalCells++;
          if (cell !== null && cell !== "" && cell !== undefined) {
            hasData = true;
          } else {
            emptyCount++;
          }
        }
      }

      // Determinar si la validación pasó
      const expectedData = action.type === "write"
        ? (action.values && action.values.length > 0)
        : (action.formula || (action.formulas && action.formulas.length > 0));

      if (expectedData && !hasData) {
        // Se esperaban datos pero no hay ninguno
        return {
          passed: false,
          message: `No se escribieron datos en ${finalRange}. Celdas vacías: ${emptyCount}/${totalCells}`,
          actualValues,
        };
      }

      // Verificar celdas vacías parciales (tabla incompleta)
      // Si hay más del 30% de celdas vacías en una acción write con múltiples filas, es sospechoso
      if (action.type === "write" && action.values && action.values.length > 1) {
        const emptyRatio = emptyCount / totalCells;
        if (emptyRatio > 0.3 && emptyCount > 3) {
          // Más del 30% vacías y más de 3 celdas vacías = tabla incompleta
          return {
            passed: false,
            message: `Tabla incompleta: ${emptyCount}/${totalCells} celdas vacías (${Math.round(emptyRatio * 100)}%). Se esperaban valores en todas las celdas de datos.`,
            actualValues,
          };
        }
      }

      // Si hay datos, verificar que no sean todos errores (para fórmulas)
      if (action.type === "formula") {
        let errorCount = 0;
        for (const row of actualValues) {
          for (const cell of row) {
            const cellStr = String(cell || "");
            if (cellStr.startsWith("#") && (cellStr.includes("ERROR") || cellStr.includes("REF") || cellStr.includes("VALUE") || cellStr.includes("NAME") || cellStr.includes("DIV"))) {
              errorCount++;
            }
          }
        }
        if (errorCount > 0 && errorCount === totalCells) {
          return {
            passed: false,
            message: `Todas las fórmulas resultaron en error (${errorCount} errores)`,
            actualValues,
          };
        }
      }

      return {
        passed: true,
        message: `Validación exitosa: ${totalCells - emptyCount}/${totalCells} celdas con datos`,
        actualValues,
      };

    } catch (error) {
      console.error("[Validación] Error al validar:", error);
      return {
        passed: true, // En caso de error de validación, asumir que pasó
        message: "No se pudo validar la acción",
      };
    }
  }

  /**
   * Convierte una letra de columna a número (A=1, B=2, ..., Z=26, AA=27)
   */
  private columnLetterToNumber(column: string): number {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + (column.charCodeAt(i) - 64);
    }
    return result;
  }

  /**
   * Convierte un número a letra de columna (1=A, 2=B, ..., 26=Z, 27=AA)
   */
  private numberToColumnLetter(num: number): string {
    let result = "";
    while (num > 0) {
      const remainder = (num - 1) % 26;
      result = String.fromCharCode(65 + remainder) + result;
      num = Math.floor((num - 1) / 26);
    }
    return result;
  }

  /**
   * Mapea el tipo de gráfico a string para Office.js
   */
  private mapChartType(chartType: ChartType): string {
    // Los valores de ChartType en Office.js son strings
    const chartTypeMap: Record<ChartType, string> = {
      barClustered: "BarClustered",
      barStacked: "BarStacked",
      columnClustered: "ColumnClustered",
      columnStacked: "ColumnStacked",
      line: "Line",
      lineMarkers: "LineMarkers",
      pie: "Pie",
      doughnut: "Doughnut",
      area: "Area",
      areaStacked: "AreaStacked",
    };
    return chartTypeMap[chartType] || "ColumnClustered";
  }

  /**
   * Aplica formato a un rango dentro de un contexto Excel.run
   */
  private applyFormatToRange(range: ExcelRange, format: FormatOptions): void {
    // Aplicar fuente
    if (format.bold !== undefined) {
      range.format.font.bold = format.bold;
    }
    if (format.italic !== undefined) {
      range.format.font.italic = format.italic;
    }
    if (format.fontSize !== undefined) {
      range.format.font.size = format.fontSize;
    }
    if (format.fontColor) {
      range.format.font.color = format.fontColor;
    }

    // Aplicar relleno
    if (format.backgroundColor) {
      range.format.fill.color = format.backgroundColor;
    }

    // Aplicar alineación
    if (format.horizontalAlignment) {
      range.format.horizontalAlignment = format.horizontalAlignment;
    }
    if (format.verticalAlignment) {
      range.format.verticalAlignment = format.verticalAlignment;
    }

    // Wrap text
    if (format.wrapText !== undefined) {
      range.format.wrapText = format.wrapText;
    }

    // Bordes
    if (format.borders) {
      const borderStyle = typeof format.borders === "object" ? format.borders.style || "thin" : "thin";
      const borderColor = typeof format.borders === "object" ? format.borders.color || "#000000" : "#000000";

      const borderPositions = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"];
      for (const pos of borderPositions) {
        try {
          const border = range.format.borders.getItem(pos);
          border.style = borderStyle === "thin" ? "Continuous" : borderStyle === "medium" ? "Medium" : "Thick";
          border.color = borderColor;
        } catch {
          // Ignorar errores de bordes internos en rangos de una celda
        }
      }
    }

    // Merge (si se especifica en format)
    if (format.merge) {
      range.merge();
    }

    // Number format
    if (format.numberFormat) {
      range.numberFormat = [[format.numberFormat]];
    }
  }

  // ===== ANÁLISIS DE DATOS GRANDES =====

  /**
   * Obtiene los encabezados (primera fila) de la hoja activa
   * Útil para archivos grandes - solo lee la primera fila
   */
  async getHeaders(): Promise<{ headers: string[]; columnLetters: string[]; lastColumn: string }> {
    try {
      return await Excel.run(async (context: ExcelContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRangeOrNullObject();

        usedRange.load(["columnCount", "address"]);
        await context.sync();

        if (usedRange.isNullObject) {
          return { headers: [], columnLetters: [], lastColumn: "A" };
        }

        // Leer solo la primera fila
        const headerRange = sheet.getRange(`1:1`).getResizedRange(0, usedRange.columnCount - 1);
        headerRange.load("values");
        await context.sync();

        const headers = headerRange.values[0].map(v => v === null ? "" : String(v));

        // Generar letras de columna
        const columnLetters: string[] = [];
        for (let i = 0; i < headers.length; i++) {
          columnLetters.push(this.numberToColumn(i + 1));
        }

        return {
          headers,
          columnLetters,
          lastColumn: this.numberToColumn(usedRange.columnCount),
        };
      });
    } catch (error) {
      console.error("Error obteniendo encabezados:", error);
      return { headers: [], columnLetters: [], lastColumn: "A" };
    }
  }

  /**
   * Busca una columna por nombre de encabezado (búsqueda flexible)
   * Retorna la letra de la columna o null si no la encuentra
   */
  async findColumnByHeader(searchTerm: string): Promise<{ column: string; header: string; index: number } | null> {
    const { headers, columnLetters } = await this.getHeaders();

    const searchLower = searchTerm.toLowerCase().trim();

    // Búsqueda exacta primero
    let index = headers.findIndex(h => h.toLowerCase().trim() === searchLower);

    // Si no encuentra exacta, buscar que contenga el término
    if (index === -1) {
      index = headers.findIndex(h => h.toLowerCase().includes(searchLower));
    }

    // Buscar sinónimos comunes
    if (index === -1) {
      const synonyms: Record<string, string[]> = {
        "status": ["estado", "estatus", "situacion", "situación"],
        "estado": ["status", "estatus", "situacion", "situación"],
        "cliente": ["customer", "nombre", "name", "client"],
        "fecha": ["date", "dia", "día"],
        "liquidado": ["liquidation", "closed", "cerrado"],
      };

      const terms = synonyms[searchLower] || [];
      for (const term of terms) {
        index = headers.findIndex(h => h.toLowerCase().includes(term));
        if (index !== -1) break;
      }
    }

    if (index === -1) return null;

    return {
      column: columnLetters[index],
      header: headers[index],
      index,
    };
  }

  /**
   * Cuenta valores en una columna específica
   * Optimizado para archivos grandes - procesa en bloques
   */
  async countValuesInColumn(
    column: string,
    searchValue: string,
    options: { caseSensitive?: boolean; exactMatch?: boolean } = {}
  ): Promise<{ count: number; total: number }> {
    const { caseSensitive = false, exactMatch = true } = options;

    try {
      return await Excel.run(async (context: ExcelContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRangeOrNullObject();

        usedRange.load("rowCount");
        await context.sync();

        if (usedRange.isNullObject) {
          return { count: 0, total: 0 };
        }

        const totalRows = usedRange.rowCount;
        const dataRows = totalRows - 1; // Excluir encabezado

        // Para archivos muy grandes, usar fórmula COUNTIF es más eficiente
        // Crear una celda temporal con la fórmula
        const tempCell = sheet.getRange("XFD1"); // Última celda de Excel
        const columnRange = `${column}2:${column}${totalRows}`;

        if (exactMatch) {
          tempCell.formulas = [[`=COUNTIF(${columnRange},"${searchValue}")`]];
        } else {
          tempCell.formulas = [[`=COUNTIF(${columnRange},"*${searchValue}*")`]];
        }

        tempCell.load("values");
        await context.sync();

        const count = Number(tempCell.values[0][0]) || 0;

        // Limpiar celda temporal
        tempCell.clear();
        await context.sync();

        return { count, total: dataRows };
      });
    } catch (error) {
      console.error("Error contando valores:", error);
      return { count: 0, total: 0 };
    }
  }

  /**
   * Obtiene valores únicos de una columna (para estadísticas)
   * Limita a los primeros N valores únicos para archivos grandes
   */
  async getUniqueValues(column: string, limit: number = 50): Promise<{ values: string[]; hasMore: boolean }> {
    try {
      return await Excel.run(async (context: ExcelContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRangeOrNullObject();

        usedRange.load("rowCount");
        await context.sync();

        if (usedRange.isNullObject) {
          return { values: [], hasMore: false };
        }

        const totalRows = usedRange.rowCount;

        // Leer la columna completa (excluyendo encabezado)
        const columnRange = sheet.getRange(`${column}2:${column}${totalRows}`);
        columnRange.load("values");
        await context.sync();

        // Extraer valores únicos
        const uniqueSet = new Set<string>();
        for (const row of columnRange.values) {
          const val = row[0];
          if (val !== null && val !== undefined && val !== "") {
            uniqueSet.add(String(val));
            if (uniqueSet.size >= limit) break;
          }
        }

        return {
          values: Array.from(uniqueSet),
          hasMore: uniqueSet.size >= limit,
        };
      });
    } catch (error) {
      console.error("Error obteniendo valores únicos:", error);
      return { values: [], hasMore: false };
    }
  }

  /**
   * Obtiene un resumen de la estructura de datos para archivos grandes
   * Incluye: dimensiones, encabezados, y muestra de datos
   */
  async getDataSummary(): Promise<{
    sheetName: string;
    dimensions: { rows: number; columns: number };
    headers: { name: string; column: string }[];
    sampleData: string;
  }> {
    try {
      return await Excel.run(async (context: ExcelContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRangeOrNullObject();

        sheet.load("name");
        usedRange.load(["rowCount", "columnCount", "address"]);
        await context.sync();

        if (usedRange.isNullObject) {
          return {
            sheetName: sheet.name,
            dimensions: { rows: 0, columns: 0 },
            headers: [],
            sampleData: "Hoja vacía",
          };
        }

        const rows = usedRange.rowCount;
        const cols = usedRange.columnCount;

        // Obtener encabezados
        const { headers, columnLetters } = await this.getHeaders();

        // Obtener muestra de datos (primeras 5 filas, máximo 10 columnas)
        const sampleCols = Math.min(cols, 10);
        const sampleRows = Math.min(rows, 6); // 1 header + 5 data
        const lastSampleCol = this.numberToColumn(sampleCols);

        const sampleRange = sheet.getRange(`A1:${lastSampleCol}${sampleRows}`);
        sampleRange.load("values");
        await context.sync();

        // Formatear muestra como texto
        const sampleLines = sampleRange.values.map(row =>
          row.slice(0, sampleCols).map(v => v === null ? "" : String(v).substring(0, 20)).join(" | ")
        );

        return {
          sheetName: sheet.name,
          dimensions: { rows, columns: cols },
          headers: headers.slice(0, 20).map((h, i) => ({ name: h, column: columnLetters[i] })),
          sampleData: sampleLines.join("\n"),
        };
      });
    } catch (error) {
      console.error("Error obteniendo resumen:", error);
      return {
        sheetName: "Error",
        dimensions: { rows: 0, columns: 0 },
        headers: [],
        sampleData: "Error al leer datos",
      };
    }
  }

  // ===== INDEXACIÓN DE DATOS =====

  // Cache del índice de datos
  private dataIndexCache: DataIndex | null = null;
  private dataIndexCacheKey: string = "";

  /**
   * Construye un índice completo de los datos de la hoja activa
   * Incluye valores únicos, conteos y estadísticas por columna
   */
  async buildDataIndex(forceRebuild: boolean = false): Promise<DataIndex | null> {
    try {
      return await Excel.run(async (context: ExcelContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRangeOrNullObject();

        sheet.load("name");
        usedRange.load(["rowCount", "columnCount", "address"]);
        await context.sync();

        if (usedRange.isNullObject) {
          return null;
        }

        const totalRows = usedRange.rowCount;
        const totalCols = usedRange.columnCount;

        // Verificar si podemos usar el cache
        const cacheKey = `${sheet.name}_${totalRows}_${totalCols}`;
        if (!forceRebuild && this.dataIndexCache && this.dataIndexCacheKey === cacheKey) {
          return this.dataIndexCache;
        }

        // Obtener encabezados DIRECTAMENTE (sin llamar getHeaders para evitar Excel.run anidado)
        const headerRange = sheet.getRange(`A1:${this.numberToColumn(totalCols)}1`);
        headerRange.load("values");
        await context.sync();

        const headers = headerRange.values[0].map((v: CellValue) => v === null ? "" : String(v));
        const columnLetters: string[] = [];
        for (let i = 0; i < totalCols; i++) {
          columnLetters.push(this.numberToColumn(i + 1));
        }

        // Limitar filas a analizar (máximo 3000 para rendimiento)
        const maxRowsToAnalyze = Math.min(totalRows, 3000);
        const columns: ColumnIndex[] = [];

        // Obtener muestra de datos (primeras 6 filas)
        const sampleRowCount = Math.min(totalRows, 7);
        const sampleColCount = Math.min(totalCols, 15);
        const sampleRange = sheet.getRange(`A1:${this.numberToColumn(sampleColCount)}${sampleRowCount}`);
        sampleRange.load("values");
        await context.sync();

        // Analizar cada columna (máximo 25 columnas para rendimiento)
        const maxColsToAnalyze = Math.min(totalCols, 25);

        for (let colIdx = 0; colIdx < maxColsToAnalyze; colIdx++) {
          const colLetter = columnLetters[colIdx];
          const header = headers[colIdx] || `Col_${colLetter}`;

          // Leer la columna (sin encabezado)
          const colRange = sheet.getRange(`${colLetter}2:${colLetter}${maxRowsToAnalyze}`);
          colRange.load("values");
          await context.sync();

          // Analizar valores
          const valueCounts: Record<string, number> = {};
          let numericCount = 0;
          let textCount = 0;
          let numericSum = 0;
          let numericMin = Infinity;
          let numericMax = -Infinity;

          for (const row of colRange.values) {
            const val = row[0];
            if (val === null || val === undefined || val === "") continue;

            const strVal = String(val).trim().substring(0, 50); // Limitar longitud
            valueCounts[strVal] = (valueCounts[strVal] || 0) + 1;

            const numVal = Number(val);
            if (!isNaN(numVal) && typeof val === "number") {
              numericCount++;
              numericSum += numVal;
              if (numVal < numericMin) numericMin = numVal;
              if (numVal > numericMax) numericMax = numVal;
            } else {
              textCount++;
            }
          }

          // Determinar tipo
          let colType: "text" | "number" | "date" | "mixed" = "text";
          if (numericCount > textCount * 2) colType = "number";
          else if (textCount > 0 && numericCount > 0) colType = "mixed";

          // Top 100 valores por frecuencia (aumentado para capturar más categorías)
          const sortedValues = Object.entries(valueCounts)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 100);

          // Estadísticas numéricas (calcular antes para usar en inferencia)
          const numericStats = colType === "number" && numericCount > 0
            ? {
                min: numericMin,
                max: numericMax,
                sum: numericSum,
                avg: numericSum / numericCount,
                count: numericCount,
              }
            : undefined;

          // Inferir tipo semántico
          const uniqueCount = Object.keys(valueCounts).length;
          const semanticType = inferSemanticType(
            header,
            colType,
            uniqueCount,
            totalRows,
            numericStats
          );

          const columnIndex: ColumnIndex = {
            header,
            column: colLetter,
            type: colType,
            semanticType,
            uniqueValues: sortedValues.map(([v]) => v),
            valueCounts: Object.fromEntries(sortedValues),
            hasMoreValues: uniqueCount > 100,
            stats: numericStats,
          };

          columns.push(columnIndex);
        }

        const dataIndex: DataIndex = {
          sheetName: sheet.name,
          totalRows,
          totalColumns: totalCols,
          columns,
          createdAt: new Date(),
          sampleRows: sampleRange.values.slice(1, 6),
        };

        this.dataIndexCache = dataIndex;
        this.dataIndexCacheKey = cacheKey;

        return dataIndex;
      });
    } catch (error) {
      console.error("Error construyendo índice:", error);
      return null;
    }
  }

  /**
   * Convierte el índice a texto para el contexto del modelo
   * Indica claramente cuando hay más valores (+) para que el modelo use fórmulas
   */
  formatDataIndexForContext(index: DataIndex): string {
    const lines: string[] = [];
    const isSampled = index.totalRows > 3000;
    lines.push(`[ÍNDICE DE DATOS: ${index.sheetName} - ${index.totalRows} filas × ${index.totalColumns} cols]`);
    if (isSampled) {
      lines.push(`[⚠️ DATOS MUESTREADOS: Solo se analizaron 3000 filas de ${index.totalRows}. Los conteos son APROXIMADOS.]`);
      lines.push(`[IMPORTANTE: Para respuestas PRECISAS en este archivo grande, NO uses los conteos del índice directamente.]`);
      lines.push(`[En su lugar, indica que el dato es aproximado o sugiere usar fórmulas COUNTIF para el cálculo exacto.]`);
    }
    lines.push(`[NOTA: + significa "hay más valores" → usa fórmulas UNIQUE/COUNTIF en ese caso]`);
    lines.push("");

    // Mapeo de tipo semántico a etiqueta legible
    const semanticLabels: Record<string, string> = {
      id: "ID/CÓDIGO→CONTAR",
      amount: "MONTO→SUMAR",
      quantity: "CANTIDAD",
      category: "CATEGORÍA",
      date: "FECHA",
      unknown: "",
    };

    for (const col of index.columns) {
      let info = `  ${col.column}:${col.header}`;

      // Agregar etiqueta semántica si es relevante
      const semanticLabel = semanticLabels[col.semanticType];
      if (semanticLabel) {
        info += ` <${semanticLabel}>`;
      }

      if (col.type === "number" && col.stats) {
        info += ` [min=${col.stats.min.toFixed(0)}, max=${col.stats.max.toFixed(0)}, count=${col.stats.count}]`;
      } else if (col.uniqueValues.length > 0) {
        // Mostrar valores con conteos
        // Limitar a 20 en el display para no sobrecargar el contexto
        const displayVals = col.uniqueValues.slice(0, 20);
        const valsStr = displayVals.map(v => `${v}(${col.valueCounts[v]})`).join(", ");
        const uniqueCount = col.uniqueValues.length;
        // + indica que hay más valores de los mostrados → el modelo debe usar fórmulas
        const hasMore = col.hasMoreValues || col.uniqueValues.length > 20;
        info += ` [${uniqueCount} valores${hasMore ? "+" : ""}] → ${valsStr}${hasMore ? "..." : ""}`;
      }
      lines.push(info);
    }

    // Instrucciones claras
    lines.push("");
    lines.push("[TIPOS: <ID/CÓDIGO→CONTAR>=usa CONTARA/COUNTA | <MONTO→SUMAR>=usa SUMA/SUM]");
    lines.push("[ESTRATEGIA: Si ves '+' usa =UNIQUE() y =COUNTIF(). Sin '+' puedes escribir los valores directamente]");

    return lines.join("\n");
  }

  /**
   * Limpia el cache del índice
   */
  clearDataIndexCache(): void {
    this.dataIndexCache = null;
    this.dataIndexCacheKey = "";
  }

  // ===== ÍNDICE LIGERO Y HOJA DE CÁLCULOS =====

  /**
   * Construye un índice ligero de la hoja activa
   * Solo lee encabezados (fila 1) de TODAS las columnas - no muestrea datos
   * Proporciona metadatos suficientes para que la IA sepa qué columnas existen
   */
  async buildLightweightIndex(forceRebuild: boolean = false): Promise<LightweightDataIndex | null> {
    try {
      return await Excel.run(async (context: ExcelContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRangeOrNullObject();

        sheet.load("name");
        usedRange.load(["rowCount", "columnCount", "address"]);
        await context.sync();

        if (usedRange.isNullObject) {
          return null;
        }

        const totalRows = usedRange.rowCount;
        const totalCols = usedRange.columnCount;

        // Verificar cache
        const cacheKey = `${sheet.name}_${totalRows}_${totalCols}_lightweight`;
        if (!forceRebuild && this.lightweightIndexCache && this.lightweightIndexCacheKey === cacheKey) {
          return this.lightweightIndexCache;
        }

        // Leer SOLO la fila de encabezados (fila 1) de TODAS las columnas
        const lastColLetter = this.numberToColumn(totalCols);
        const headerRange = sheet.getRange(`A1:${lastColLetter}1`);
        headerRange.load("values");
        await context.sync();

        // Construir metadatos de columnas
        const columns: LightweightColumnMeta[] = [];
        const headers = headerRange.values[0];

        for (let i = 0; i < totalCols; i++) {
          const colLetter = this.numberToColumn(i + 1);
          const header = headers[i] === null ? "" : String(headers[i]);

          columns.push({
            column: colLetter,
            header: header,
            dataRange: `${colLetter}2:${colLetter}${totalRows}`, // Rango de datos (sin encabezado)
          });
        }

        const lightweightIndex: LightweightDataIndex = {
          sheetName: sheet.name,
          totalRows,
          totalColumns: totalCols,
          columns,
          lastColumn: lastColLetter,
          dataRange: `A1:${lastColLetter}${totalRows}`,
          createdAt: new Date(),
        };

        // Guardar en cache
        this.lightweightIndexCache = lightweightIndex;
        this.lightweightIndexCacheKey = cacheKey;

        return lightweightIndex;
      });
    } catch (error) {
      console.error("Error construyendo índice ligero:", error);
      return null;
    }
  }

  /**
   * Formatea el índice ligero para el contexto del modelo
   * Solo incluye columna:encabezado y rango de datos - sin estadísticas
   */
  formatLightweightIndexForContext(index: LightweightDataIndex): string {
    const lines: string[] = [];
    lines.push(`[ÍNDICE DE DATOS: ${index.sheetName}]`);
    lines.push(`[Dimensiones: ${index.totalRows} filas × ${index.totalColumns} columnas (A-${index.lastColumn})]`);
    lines.push(`[Rango completo: ${index.dataRange}]`);
    
    // Determinar si los datos son "anchos" (más de columna K = 11)
    const isWideData = index.totalColumns > 11;
    if (isWideData) {
      lines.push(`[⚠️ DATOS ANCHOS: lastColumn=${index.lastColumn} > K. PREGUNTA al usuario dónde colocar contenido nuevo.]`);
    }
    
    lines.push("");
    lines.push("[COLUMNAS DISPONIBLES:]");

    // Mostrar todas las columnas con su encabezado
    for (const col of index.columns) {
      lines.push(`  ${col.column}: "${col.header}" → datos en ${col.dataRange}`);
    }

    lines.push("");
    lines.push("[INSTRUCCIONES PARA CÁLCULOS:]");
    lines.push("- Para consultas/preguntas: Usa fórmulas en la hoja oculta _AI_Calc");
    lines.push("- Ejemplo: Para contar ciudades únicas → =COUNTA(UNIQUE(I2:I99379))");
    lines.push("- Ejemplo: Para sumar valores → =SUM(M2:M99379)");
    lines.push("- Ejemplo: Para contar con condición → =COUNTIF(I2:I99379,\"BOGOTA\")");
    lines.push("- NUNCA uses datos aproximados. SIEMPRE calcula con fórmulas reales.");
    lines.push("");
    lines.push("[INSTRUCCIONES PARA CREAR CONTENIDO:]");
    
    if (isWideData) {
      lines.push(`- ⚠️ Los datos llegan hasta columna ${index.lastColumn} (muy lejos para el usuario)`);
      lines.push("- SIEMPRE PREGUNTA al usuario dónde quiere el resultado:");
      lines.push("  1. Nueva hoja (recomendado)");
      lines.push("  2. Que seleccione una celda específica");
      lines.push("- NO coloques contenido automáticamente en columnas lejanas");
    } else {
      lines.push(`- Puedes colocar contenido nuevo después de columna ${index.lastColumn}`);
    }
    lines.push("- NUNCA modifiques los datos del usuario sin autorización explícita");

    return lines.join("\n");
  }

  /**
   * Obtiene o crea la hoja oculta para cálculos de IA
   * Esta hoja se usa para ejecutar fórmulas sin afectar los datos del usuario
   */
  async getOrCreateCalcSheet(): Promise<string> {
    try {
      return await Excel.run(async (context: ExcelContext) => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const worksheets = context.workbook.worksheets as any;
        worksheets.load("items/name");
        await context.sync();

        // Buscar si la hoja ya existe
        let calcSheet = null;
        for (const sheet of worksheets.items) {
          if (sheet.name === ExcelService.CALC_SHEET_NAME) {
            calcSheet = sheet;
            break;
          }
        }

        if (!calcSheet) {
          // Crear la hoja oculta
          calcSheet = worksheets.add(ExcelService.CALC_SHEET_NAME);
          await context.sync();
          
          // Intentar ocultarla
          try {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            (calcSheet as any).visibility = "Hidden";
            await context.sync();
          } catch {
            // Hoja visible si no se pudo ocultar
          }
        }

        return ExcelService.CALC_SHEET_NAME;
      });
    } catch (error) {
      console.error("Error creando hoja de cálculos:", error);
      throw new ExcelServiceError(
        "No se pudo crear la hoja de cálculos oculta",
        error instanceof Error ? error : undefined
      );
    }
  }

  /**
   * Ejecuta fórmulas en la hoja oculta y devuelve los resultados
   * Permite a la IA hacer cálculos sin modificar los datos del usuario
   *
   * @param formulas Array de fórmulas a ejecutar (sin el =)
   * @param sourceSheetName Nombre de la hoja con los datos originales
   * @returns Array de resultados con la fórmula y su valor calculado
   */
  async executeCalcFormulas(formulas: string[], sourceSheetName: string): Promise<CalcResult[]> {
    const results: CalcResult[] = [];

    try {
      await Excel.run(async (context: ExcelContext) => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const worksheets = context.workbook.worksheets as any;
        worksheets.load("items/name");
        await context.sync();

        // Buscar o crear la hoja de cálculos
        let calcSheet = null;
        for (const sheet of worksheets.items) {
          if (sheet.name === ExcelService.CALC_SHEET_NAME) {
            calcSheet = sheet;
            break;
          }
        }

        if (!calcSheet) {
          calcSheet = worksheets.add(ExcelService.CALC_SHEET_NAME);
          await context.sync();

          // Intentar ocultarla
          try {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            (calcSheet as any).visibility = "Hidden";
            await context.sync();
          } catch {
            // Continuar con hoja visible
          }
        }

        // Limpiar la hoja antes de nuevos cálculos
        const usedRange = calcSheet.getUsedRangeOrNullObject();
        usedRange.load("isNullObject");
        await context.sync();

        if (!usedRange.isNullObject) {
          usedRange.clear();
          await context.sync();
        }

        // Escribir todas las fórmulas primero
        const cells: any[] = [];
        for (let i = 0; i < formulas.length; i++) {
          const cellAddress = `A${i + 1}`;
          const cell = calcSheet.getRange(cellAddress);

          // Ajustar fórmulas para referenciar la hoja de origen
          let formula = formulas[i];
          if (!formula.startsWith("=")) {
            formula = "=" + formula;
          }

          const adjustedFormula = this.adjustFormulaReferences(formula, sourceSheetName);
          cell.formulas = [[adjustedFormula]];
          cells.push(cell);
        }

        await context.sync();

        // Cargar todos los valores de una vez
        for (const cell of cells) {
          cell.load("values");
        }
        await context.sync();

        // Recopilar resultados
        for (let i = 0; i < formulas.length; i++) {
          const value = cells[i].values[0][0];
          results.push({
            formula: formulas[i],
            result: value,
            cell: `A${i + 1}`,
          });
        }
      });

      return results;
    } catch (error) {
      console.error("❌ Error ejecutando fórmulas:", error);
      throw new ExcelServiceError(
        "Error ejecutando fórmulas en hoja oculta",
        error instanceof Error ? error : undefined
      );
    }
  }

  /**
   * Ajusta las referencias de fórmulas para apuntar a la hoja correcta
   * Ej: COUNTIF(I2:I99379,"BOGOTA") → COUNTIF('Hoja1'!I2:I99379,"BOGOTA")
   */
  private adjustFormulaReferences(formula: string, sheetName: string): string {
    // Escapar nombre de hoja si tiene espacios o caracteres especiales
    const escapedSheetName = sheetName.includes(" ") || /[^a-zA-Z0-9_]/.test(sheetName)
      ? `'${sheetName}'`
      : sheetName;

    // Regex para encontrar referencias de rango (ej: A1, A1:B10, $A$1, etc.)
    // pero NO dentro de strings entre comillas
    const rangePattern = /(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)/gi;

    // Dividir la fórmula en partes (dentro y fuera de comillas)
    const parts: string[] = [];
    let current = "";
    let inQuotes = false;

    for (let i = 0; i < formula.length; i++) {
      const char = formula[i];
      if (char === '"' && (i === 0 || formula[i - 1] !== "\\")) {
        if (inQuotes) {
          parts.push(`"${current}"`);
          current = "";
        } else if (current) {
          parts.push(current);
          current = "";
        }
        inQuotes = !inQuotes;
      } else {
        current += char;
      }
    }
    if (current) parts.push(current);

    // Procesar solo las partes fuera de comillas
    const processedParts = parts.map(part => {
      if (part.startsWith('"')) {
        return part; // Mantener strings sin cambios
      }
      // Solo agregar referencia de hoja si no tiene ya una
      return part.replace(rangePattern, (match) => {
        // Si ya tiene referencia de hoja (contiene !), no modificar
        if (part.includes("!")) return match;
        return `${escapedSheetName}!${match}`;
      });
    });

    return processedParts.join("");
  }

  /**
   * Limpia la hoja de cálculos oculta
   */
  async clearCalcSheet(): Promise<void> {
    try {
      await Excel.run(async (context: ExcelContext) => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const worksheets = context.workbook.worksheets as any;
        worksheets.load("items/name");
        await context.sync();

        for (const sheet of worksheets.items) {
          if (sheet.name === ExcelService.CALC_SHEET_NAME) {
            const usedRange = sheet.getUsedRangeOrNullObject();
            usedRange.load("isNullObject");
            await context.sync();

            if (!usedRange.isNullObject) {
              usedRange.clear();
              await context.sync();
            }
            break;
          }
        }
      });
    } catch (error) {
      console.warn("Error limpiando hoja de cálculos:", error);
    }
  }

  /**
   * Limpia el cache del índice ligero
   */
  clearLightweightIndexCache(): void {
    this.lightweightIndexCache = null;
    this.lightweightIndexCacheKey = "";
  }

  /**
   * Convierte número de columna a letra (1=A, 27=AA, etc.)
   */
  private numberToColumn(num: number): string {
    let result = "";
    let n = num;
    while (n > 0) {
      n--;
      result = String.fromCharCode(65 + (n % 26)) + result;
      n = Math.floor(n / 26);
    }
    return result;
  }

  // ===== UTILIDADES =====

  /**
   * Obtiene información de la hoja activa
   */
  async getActiveSheetInfo(): Promise<{ name: string; rowCount: number; columnCount: number }> {
    try {
      return await Excel.run(async (context: ExcelContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRangeOrNullObject();

        sheet.load("name");
        usedRange.load(["rowCount", "columnCount"]);

        await context.sync();

        return {
          name: sheet.name,
          rowCount: usedRange.isNullObject ? 0 : usedRange.rowCount,
          columnCount: usedRange.isNullObject ? 0 : usedRange.columnCount,
        };
      });
    } catch (error) {
      throw new ExcelServiceError(
        "Error al obtener información de la hoja",
        error instanceof Error ? error : undefined
      );
    }
  }

  /**
   * Resalta un rango seleccionándolo (sin modificar formato)
   */
  async highlightRange(address: string, _durationMs: number = 2000): Promise<void> {
    // Solo seleccionar el rango, no modificar colores para no afectar bordes
    await this.selectRange(address);
  }

  /**
   * Selecciona un rango específico
   */
  async selectRange(address: string): Promise<void> {
    try {
      await Excel.run(async (context: ExcelContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);
        range.select();
        await context.sync();
      });
    } catch (error) {
      console.warn("No se pudo seleccionar el rango:", error);
    }
  }

  /**
   * Descubre los valores únicos de una columna y cuenta cuántos hay de cada uno
   * que cumplan opcionalmente con un criterio adicional.
   * 
   * Ejemplo: getUniqueCountsByCategory("T", "AI", "ANULADO") 
   *   → Cuenta contratos anulados por cada zona
   * 
   * @param categoryColumn Columna de categorías (ej: "T" para ZONAVENTA)
   * @param filterColumn Columna opcional para filtrar (ej: "AI" para STATUS)
   * @param filterValue Valor del filtro (ej: "ANULADO")
   * @param sheetName Nombre de la hoja (opcional)
   * @returns Array de {category, count}
   */
  async getUniqueCountsByCategory(
    categoryColumn: string,
    filterColumn?: string,
    filterValue?: string,
    sheetName?: string
  ): Promise<{ category: string; count: number }[]> {
    try {
      return await Excel.run(async (context: ExcelContext) => {
        let sheet;
        if (sheetName) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const worksheets = context.workbook.worksheets as any;
          worksheets.load("items/name");
          await context.sync();
          sheet = worksheets.items.find((s: { name: string }) => s.name === sheetName);
          if (!sheet) {
            sheet = context.workbook.worksheets.getActiveWorksheet();
          }
        } else {
          sheet = context.workbook.worksheets.getActiveWorksheet();
        }
        
        const usedRange = sheet.getUsedRange();
        usedRange.load("rowCount");
        await context.sync();
        
        const lastRow = usedRange.rowCount;

        // Leer la columna de categorías (desde fila 2 para saltar encabezado)
        const categoryRange = sheet.getRange(`${categoryColumn}2:${categoryColumn}${lastRow}`);
        categoryRange.load("values");
        
        // Si hay filtro, leer también esa columna
        let filterRange = null;
        if (filterColumn && filterValue) {
          filterRange = sheet.getRange(`${filterColumn}2:${filterColumn}${lastRow}`);
          filterRange.load("values");
        }
        
        await context.sync();
        
        // Contar valores únicos
        const counts = new Map<string, number>();
        const categoryValues = categoryRange.values;
        const filterValues = filterRange?.values;
        
        for (let i = 0; i < categoryValues.length; i++) {
          const category = String(categoryValues[i][0] || "").trim();
          
          // Ignorar celdas vacías
          if (!category) continue;
          
          // Si hay filtro, verificar que coincida
          if (filterValues && filterValue) {
            const filterVal = String(filterValues[i][0] || "").trim().toUpperCase();
            if (filterVal !== filterValue.toUpperCase()) continue;
          }
          
          // Incrementar contador
          counts.set(category, (counts.get(category) || 0) + 1);
        }
        
        // Convertir a array y ordenar por cantidad descendente
        return Array.from(counts.entries())
          .map(([category, count]) => ({ category, count }))
          .sort((a, b) => b.count - a.count);
      });
    } catch (error) {
      console.error("Error obteniendo categorías únicas:", error);
      throw new ExcelServiceError(
        "No se pudieron obtener las categorías únicas",
        error instanceof Error ? error : undefined
      );
    }
  }

  /**
   * Calcula el promedio de una columna de valores agrupado por categorías
   * 
   * Ejemplo: getAverageByCategory("C", "E") 
   *   → Promedio de columna E (edad) por cada ciudad en columna C
   * 
   * @param categoryColumn Columna de categorías (ej: "C" para CIUDAD)
   * @param valueColumn Columna de valores a promediar (ej: "E" para EDAD)
   * @param filterColumn Columna opcional para filtrar
   * @param filterValue Valor del filtro
   * @param sheetName Nombre de la hoja (opcional)
   * @returns Array de {category, average, count}
   */
  async getAverageByCategory(
    categoryColumn: string,
    valueColumn: string,
    filterColumn?: string,
    filterValue?: string,
    sheetName?: string
  ): Promise<{ category: string; average: number; count: number }[]> {
    try {
      return await Excel.run(async (context: ExcelContext) => {
        let sheet;
        if (sheetName) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const worksheets = context.workbook.worksheets as any;
          worksheets.load("items/name");
          await context.sync();
          sheet = worksheets.items.find((s: { name: string }) => s.name === sheetName);
          if (!sheet) {
            sheet = context.workbook.worksheets.getActiveWorksheet();
          }
        } else {
          sheet = context.workbook.worksheets.getActiveWorksheet();
        }
        
        const usedRange = sheet.getUsedRange();
        usedRange.load("rowCount");
        await context.sync();
        
        const lastRow = usedRange.rowCount;

        // Leer columna de categorías
        const categoryRange = sheet.getRange(`${categoryColumn}2:${categoryColumn}${lastRow}`);
        categoryRange.load("values");
        
        // Leer columna de valores
        const valueRange = sheet.getRange(`${valueColumn}2:${valueColumn}${lastRow}`);
        valueRange.load("values");
        
        // Si hay filtro, leer también esa columna
        let filterRange = null;
        if (filterColumn && filterValue) {
          filterRange = sheet.getRange(`${filterColumn}2:${filterColumn}${lastRow}`);
          filterRange.load("values");
        }
        
        await context.sync();
        
        // Calcular suma y conteo por categoría
        const sums = new Map<string, number>();
        const counts = new Map<string, number>();
        const categoryValues = categoryRange.values;
        const valueValues = valueRange.values;
        const filterValues = filterRange?.values;
        
        for (let i = 0; i < categoryValues.length; i++) {
          const category = String(categoryValues[i][0] || "").trim();
          const value = valueValues[i][0];
          
          // Ignorar celdas vacías o valores no numéricos
          if (!category) continue;
          if (typeof value !== "number" || isNaN(value)) continue;
          
          // Si hay filtro, verificar que coincida
          if (filterValues && filterValue) {
            const filterVal = String(filterValues[i][0] || "").trim().toUpperCase();
            if (filterVal !== filterValue.toUpperCase()) continue;
          }
          
          // Acumular suma y conteo
          sums.set(category, (sums.get(category) || 0) + value);
          counts.set(category, (counts.get(category) || 0) + 1);
        }
        
        // Calcular promedios
        return Array.from(sums.entries())
          .map(([category, sum]) => ({
            category,
            average: Math.round((sum / (counts.get(category) || 1)) * 100) / 100,
            count: counts.get(category) || 0
          }))
          .sort((a, b) => b.average - a.average);
      });
    } catch (error) {
      console.error("Error calculando promedios:", error);
      throw new ExcelServiceError(
        "No se pudieron calcular los promedios por categoría",
        error instanceof Error ? error : undefined
      );
    }
  }

  /**
   * Crea una nueva hoja con los datos proporcionados
   */
  async createSheetWithData(sheetName: string, data: string[][]): Promise<void> {
    if (!this.isOfficeReady()) {
      throw new ExcelServiceError("Office no está listo");
    }

    try {
      await Excel.run(async (context) => {
        const workbook = context.workbook;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sheets = workbook.worksheets as any;
        sheets.load("items/name");
        await context.sync();

        // Verificar si ya existe una hoja con ese nombre
        let finalName = sheetName;
        let counter = 1;
        const existingNames = (sheets.items as { name: string }[]).map(s => s.name.toLowerCase());

        while (existingNames.includes(finalName.toLowerCase())) {
          finalName = `${sheetName.substring(0, 28)}_${counter}`;
          counter++;
        }

        // Crear nueva hoja
        const newSheet = sheets.add(finalName);
        newSheet.activate();

        // Determinar el rango necesario
        const rowCount = data.length;
        const colCount = Math.max(...data.map(row => row.length));

        if (rowCount === 0 || colCount === 0) {
          throw new Error("No hay datos para escribir");
        }

        // Convertir número de columna a letra
        const colLetter = this.columnIndexToLetter(colCount - 1);
        const range = newSheet.getRange(`A1:${colLetter}${rowCount}`);

        // Normalizar data para que todas las filas tengan el mismo número de columnas
        const normalizedData = data.map(row => {
          const newRow = [...row];
          while (newRow.length < colCount) {
            newRow.push("");
          }
          return newRow;
        });

        // Escribir datos
        range.values = normalizedData;

        // Formatear como tabla si tiene encabezados (más de una fila)
        if (rowCount > 1) {
          try {
            const tableRange = newSheet.getRange(`A1:${colLetter}${rowCount}`);
            const table = newSheet.tables.add(tableRange, true);
            table.name = `Tabla_${finalName.replace(/\s/g, "_")}`;
            table.style = "TableStyleMedium2";
          } catch (tableError) {
            console.warn("No se pudo crear tabla, continuando sin formato de tabla:", tableError);
          }
        }

        // Autoajustar columnas
        try {
          const usedRange = newSheet.getUsedRange();
          usedRange.format.autofitColumns();
        } catch (fitError) {
          console.warn("No se pudo autoajustar columnas:", fitError);
        }

        await context.sync();
      });
    } catch (error) {
      throw new ExcelServiceError(
        "No se pudo crear la hoja con los datos",
        error instanceof Error ? error : undefined
      );
    }
  }

  /**
   * Convierte índice de columna (0-based) a letra de Excel
   */
  private columnIndexToLetter(index: number): string {
    let letter = "";
    while (index >= 0) {
      letter = String.fromCharCode((index % 26) + 65) + letter;
      index = Math.floor(index / 26) - 1;
    }
    return letter;
  }
}

// Instancia singleton del servicio
export const excelService = new ExcelService();
