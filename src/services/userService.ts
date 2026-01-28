/**
 * User Service - Obtiene información del usuario de Office 365
 *
 * NOTA: SSO deshabilitado temporalmente porque Admin Center no soporta
 * la configuración WebApplicationInfo en el manifest.
 * El código SSO se mantiene comentado para habilitarlo en el futuro.
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

// Flag para habilitar/deshabilitar SSO (deshabilitado por ahora)
const SSO_ENABLED = false;

/**
 * Obtiene información del usuario actual de Office 365
 */
export async function getUserInfo(): Promise<UserInfo> {
  // Retornar cache si existe
  if (cachedUser) {
    return cachedUser;
  }

  // Intentar SSO si está habilitado
  if (SSO_ENABLED) {
    try {
      const ssoUser = await tryGetUserFromSSO();
      if (ssoUser) {
        cachedUser = ssoUser;
        return ssoUser;
      }
    } catch (error) {
      console.warn('[UserService] SSO no disponible');
    }
  }

  // Fallback: usuario genérico
  const fallbackUser: UserInfo = {
    name: 'Usuario',
    email: '',
    firstName: 'Usuario',
    isAuthenticated: false,
    source: 'fallback'
  };

  cachedUser = fallbackUser;
  return fallbackUser;
}

/**
 * Intenta obtener el usuario via SSO de Office
 * (Deshabilitado - requiere configuración WebApplicationInfo en manifest)
 */
async function tryGetUserFromSSO(): Promise<UserInfo | null> {
  if (typeof Office === 'undefined' || !Office.auth) {
    return null;
  }

  try {
    const token = await Office.auth.getAccessToken({
      allowSignInPrompt: false,
      allowConsentPrompt: false,
      forMSGraphAccess: false
    });

    if (!token) {
      return null;
    }

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
    // SSO no disponible
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

    const payload = parts[1];
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
