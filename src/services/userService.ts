/**
 * User Service - Obtiene información del usuario de Office 365
 */

export interface UserInfo {
  name: string;
  email: string;
  firstName: string;
  isAuthenticated: boolean;
  source: 'sso' | 'fallback' | 'unknown';
}

// Cache del usuario
let cachedUser: UserInfo | null = null;

/**
 * Obtiene información del usuario actual de Office 365
 * Intenta SSO primero, luego fallback
 */
export async function getUserInfo(): Promise<UserInfo> {
  // Retornar cache si existe
  if (cachedUser) {
    return cachedUser;
  }

  // Intentar SSO primero
  try {
    const ssoUser = await tryGetUserFromSSO();
    if (ssoUser) {
      cachedUser = ssoUser;
      return ssoUser;
    }
  } catch {
    // SSO no disponible
  }

  // Fallback: usuario desconocido
  const fallbackUser: UserInfo = {
    name: 'Usuario',
    email: '',
    firstName: 'Usuario',
    isAuthenticated: false,
    source: 'unknown'
  };

  cachedUser = fallbackUser;
  return fallbackUser;
}

/**
 * Intenta obtener el usuario via SSO de Office
 */
async function tryGetUserFromSSO(): Promise<UserInfo | null> {
  if (typeof Office === 'undefined' || !Office.auth) {
    return null;
  }

  try {
    // Solicitar token de acceso
    const token = await Office.auth.getAccessToken({
      allowSignInPrompt: false, // No mostrar popup de login
      allowConsentPrompt: false, // No mostrar popup de consentimiento
      forMSGraphAccess: false // No necesitamos Graph, solo el token básico
    });

    if (!token) {
      return null;
    }

    // Decodificar el JWT para obtener claims del usuario
    const userInfo = decodeJwtPayload(token);

    if (userInfo) {
      return {
        name: userInfo.name || userInfo.preferred_username || 'Usuario',
        email: userInfo.preferred_username || userInfo.upn || userInfo.email || '',
        firstName: getFirstName(userInfo.name || userInfo.preferred_username || 'Usuario'),
        isAuthenticated: true,
        source: 'sso'
      };
    }
  } catch {
    // SSO error - continuará con fallback
  }

  return null;
}

/**
 * Decodifica el payload de un JWT (sin verificar firma)
 */
function decodeJwtPayload(token: string): any {
  try {
    const parts = token.split('.');
    if (parts.length !== 3) {
      return null;
    }

    // El payload es la segunda parte
    const payload = parts[1];

    // Decodificar base64url
    const base64 = payload.replace(/-/g, '+').replace(/_/g, '/');
    const jsonPayload = decodeURIComponent(
      atob(base64)
        .split('')
        .map(c => '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2))
        .join('')
    );

    return JSON.parse(jsonPayload);
  } catch {
    return null;
  }
}

/**
 * Extrae el primer nombre de un nombre completo
 */
function getFirstName(fullName: string): string {
  if (!fullName) return 'Usuario';

  const parts = fullName.trim().split(/\s+/);
  return parts[0] || 'Usuario';
}

/**
 * Limpia el cache del usuario (útil para logout)
 */
export function clearUserCache(): void {
  cachedUser = null;
}

/**
 * Obtiene un saludo personalizado basado en la hora
 */
export function getGreeting(firstName: string): string {
  const hour = new Date().getHours();

  if (hour >= 5 && hour < 12) {
    return `Buenos días, ${firstName}`;
  } else if (hour >= 12 && hour < 19) {
    return `Buenas tardes, ${firstName}`;
  } else {
    return `Buenas noches, ${firstName}`;
  }
}
