export const SESSION_COOKIE = "hsi_dashboard_session";
export const SESSION_MAX_AGE_SECONDS = 60 * 60 * 12;

export type UserRole = "PRQ_INTERNAL" | "VIEWER";
export type DashboardPermission = "DASHBOARD_VIEW" | "MASTER_ACCESS" | "MASTER_EDIT";
export type SessionUser = {
  role: UserRole;
  username: string;
};
export type DashboardPermissions = {
  canAccessMaster: boolean;
  canEditMaster: boolean;
  canViewDashboard: boolean;
};

const encoder = new TextEncoder();

function getSessionSecret() {
  return process.env.DASHBOARD_SESSION_SECRET || process.env.DASHBOARD_PASSWORD || "";
}

function getMasterPassword() {
  return process.env.MASTER_ACCESS_PASSWORD || "123456";
}

function base64UrlEncode(value: string | ArrayBuffer) {
  const bytes = typeof value === "string" ? encoder.encode(value) : new Uint8Array(value);
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function base64UrlDecode(value: string) {
  const padded = value.replace(/-/g, "+").replace(/_/g, "/").padEnd(Math.ceil(value.length / 4) * 4, "=");
  const binary = atob(padded);
  const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
  return new TextDecoder().decode(bytes);
}

function sameString(a: string, b: string) {
  let mismatch = a.length === b.length ? 0 : 1;
  const length = Math.max(a.length, b.length, 1);
  for (let index = 0; index < length; index += 1) {
    const left = a.charCodeAt(index % Math.max(a.length, 1)) || 0;
    const right = b.charCodeAt(index % Math.max(b.length, 1)) || 0;
    mismatch |= left ^ right;
  }
  return mismatch === 0;
}

async function sign(payload: string, secret: string) {
  const key = await crypto.subtle.importKey("raw", encoder.encode(secret), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  const signature = await crypto.subtle.sign("HMAC", key, encoder.encode(payload));
  return base64UrlEncode(signature);
}

export function credentialsConfigured() {
  return Boolean(process.env.DASHBOARD_USERNAME && process.env.DASHBOARD_PASSWORD);
}

export async function validateCredentials(username: string, password: string) {
  const configuredUsername = process.env.DASHBOARD_USERNAME ?? "";
  const configuredPassword = process.env.DASHBOARD_PASSWORD ?? "";
  if (!configuredUsername || !configuredPassword) return false;
  const [usernameHash, configuredUsernameHash, passwordHash, configuredPasswordHash] = await Promise.all([
    crypto.subtle.digest("SHA-256", encoder.encode(username)),
    crypto.subtle.digest("SHA-256", encoder.encode(configuredUsername)),
    crypto.subtle.digest("SHA-256", encoder.encode(password)),
    crypto.subtle.digest("SHA-256", encoder.encode(configuredPassword)),
  ]);
  return (
    sameString(base64UrlEncode(usernameHash), base64UrlEncode(configuredUsernameHash)) &&
    sameString(base64UrlEncode(passwordHash), base64UrlEncode(configuredPasswordHash))
  );
}

function configuredRole(): UserRole {
  return process.env.DASHBOARD_USER_ROLE === "VIEWER" ? "VIEWER" : "PRQ_INTERNAL";
}

export function permissionsForUser(user: SessionUser | null): DashboardPermissions {
  const isInternal = user?.role === "PRQ_INTERNAL";
  return {
    canAccessMaster: isInternal,
    canEditMaster: isInternal,
    canViewDashboard: Boolean(user),
  };
}

export function hasPermission(user: SessionUser | null, permission: DashboardPermission) {
  const permissions = permissionsForUser(user);
  if (permission === "DASHBOARD_VIEW") return permissions.canViewDashboard;
  if (permission === "MASTER_ACCESS") return permissions.canAccessMaster;
  return permissions.canEditMaster;
}

export async function validateMasterPassword(password: string) {
  const [passwordHash, configuredPasswordHash] = await Promise.all([
    crypto.subtle.digest("SHA-256", encoder.encode(password)),
    crypto.subtle.digest("SHA-256", encoder.encode(getMasterPassword())),
  ]);
  return sameString(base64UrlEncode(passwordHash), base64UrlEncode(configuredPasswordHash));
}

export async function createMasterUnlockToken(username: string) {
  const secret = getSessionSecret();
  if (!secret) throw new Error("Dashboard session secret is not configured.");
  const payload = base64UrlEncode(
    JSON.stringify({
      master: true,
      username,
      exp: Date.now() + SESSION_MAX_AGE_SECONDS * 1000,
    }),
  );
  return `${payload}.${await sign(payload, secret)}`;
}

export async function verifyMasterUnlockToken(token: string | undefined) {
  const secret = getSessionSecret();
  if (!secret || !token) return false;
  const [payload, signature] = token.split(".");
  if (!payload || !signature) return false;
  const expectedSignature = await sign(payload, secret);
  if (!sameString(signature, expectedSignature)) return false;
  try {
    const decoded = JSON.parse(base64UrlDecode(payload)) as { exp?: number; master?: boolean; username?: string };
    return Boolean(decoded.master && decoded.username && decoded.exp && decoded.exp > Date.now());
  } catch {
    return false;
  }
}

export async function createSessionToken(username: string) {
  const secret = getSessionSecret();
  if (!secret) throw new Error("Dashboard session secret is not configured.");
  const payload = base64UrlEncode(
    JSON.stringify({
      role: configuredRole(),
      username,
      exp: Date.now() + SESSION_MAX_AGE_SECONDS * 1000,
    }),
  );
  return `${payload}.${await sign(payload, secret)}`;
}

export async function getSessionUser(token: string | undefined): Promise<SessionUser | null> {
  const secret = getSessionSecret();
  if (!secret || !token) return null;
  const [payload, signature] = token.split(".");
  if (!payload || !signature) return null;
  const expectedSignature = await sign(payload, secret);
  if (!sameString(signature, expectedSignature)) return null;
  try {
    const decoded = JSON.parse(base64UrlDecode(payload)) as { exp?: number; role?: UserRole; username?: string };
    if (!decoded.username || !decoded.exp || decoded.exp <= Date.now()) return null;
    return {
      role: decoded.role === "VIEWER" ? "VIEWER" : "PRQ_INTERNAL",
      username: decoded.username,
    };
  } catch {
    return null;
  }
}

export async function verifySessionToken(token: string | undefined) {
  return Boolean(await getSessionUser(token));
}
