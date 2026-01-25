/**
 * Declaraciones de tipos globales para el proyecto
 */

// Office.js globals (complementan @types/office-js)
declare const Office: typeof import("@microsoft/office-js").Office;
declare const Excel: typeof import("@microsoft/office-js").Excel;

// Extensión de Window para Office
interface Window {
  Office?: typeof Office;
  Excel?: typeof Excel;
}
