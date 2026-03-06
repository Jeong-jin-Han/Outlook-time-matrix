import { CLIENT_ID, TENANT_ID, OBJECT_ID } from "./config.js";

// Open side panel when the extension icon is clicked
chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true }).catch(console.error);

const DEFAULT_SETTINGS = {
  tenantId: TENANT_ID,
  clientId: CLIENT_ID,
  objectId: OBJECT_ID
};

// Only handles auth — Graph API calls are made directly from the side panel (popup.js)
chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  (async () => {
    try {
      if (message.type === "auth") {
        const token = await getAccessToken(message.forcePrompt ?? false);
        sendResponse({ ok: true, token });
        return;
      }
      sendResponse({ ok: false, error: "Unsupported message type" });
    } catch (error) {
      sendResponse({ ok: false, error: error.message || String(error) });
    }
  })();

  return true;
});

async function getAccessToken(forcePrompt) {
  const settings = await getSettings();
  if (!settings.clientId || !settings.tenantId) {
    throw new Error("Set Tenant ID and Client ID in extension options first.");
  }

  const tokenState = await chrome.storage.local.get(["accessToken", "accessTokenExpiresAt", "refreshToken"]);

  // Return cached access token if still valid
  if (!forcePrompt && tokenState.accessToken && tokenState.accessTokenExpiresAt && Date.now() < tokenState.accessTokenExpiresAt - 60_000) {
    return tokenState.accessToken;
  }

  // Try silent refresh before falling back to interactive auth
  if (!forcePrompt && tokenState.refreshToken) {
    try {
      return await refreshWithToken(tokenState.refreshToken, settings);
    } catch {
      await chrome.storage.local.remove(["accessToken", "accessTokenExpiresAt", "refreshToken"]);
    }
  }

  // If not forcing interactive auth, signal the caller to show sign-in UI
  if (!forcePrompt) {
    throw new Error("Not signed in");
  }

  const redirectUri = chrome.identity.getRedirectURL("microsoft");
  const verifier = generateRandomString(64);
  const challenge = await sha256base64url(verifier);
  const state = generateRandomString(32);
  const scope = "openid profile offline_access User.Read Calendars.ReadWrite";

  const authUrl = new URL(`https://login.microsoftonline.com/${encodeURIComponent(settings.tenantId)}/oauth2/v2.0/authorize`);
  authUrl.searchParams.set("client_id", settings.clientId);
  authUrl.searchParams.set("response_type", "code");
  authUrl.searchParams.set("redirect_uri", redirectUri);
  authUrl.searchParams.set("response_mode", "query");
  authUrl.searchParams.set("scope", scope);
  authUrl.searchParams.set("state", state);
  authUrl.searchParams.set("code_challenge", challenge);
  authUrl.searchParams.set("code_challenge_method", "S256");
  if (forcePrompt) {
    authUrl.searchParams.set("prompt", "select_account");
  }

  const callbackUrl = await chrome.identity.launchWebAuthFlow({
    url: authUrl.toString(),
    interactive: true
  });

  if (!callbackUrl) throw new Error("Authentication canceled.");

  const callback = new URL(callbackUrl);
  const returnedState = callback.searchParams.get("state");
  const code = callback.searchParams.get("code");
  const authError = callback.searchParams.get("error_description") || callback.searchParams.get("error");

  if (authError) throw new Error(authError);
  if (!code || returnedState !== state) throw new Error("Invalid OAuth callback response.");

  const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(settings.tenantId)}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: settings.clientId,
    grant_type: "authorization_code",
    code,
    redirect_uri: redirectUri,
    code_verifier: verifier,
    scope
  });

  const tokenResponse = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body
  });

  const tokenJson = await tokenResponse.json().catch(() => null);
  if (!tokenResponse.ok || !tokenJson?.access_token) {
    throw new Error(tokenJson?.error_description || "Token exchange failed.");
  }

  const expiresAt = Date.now() + (tokenJson.expires_in || 3600) * 1000;
  await chrome.storage.local.set({
    accessToken: tokenJson.access_token,
    accessTokenExpiresAt: expiresAt,
    ...(tokenJson.refresh_token ? { refreshToken: tokenJson.refresh_token } : {})
  });

  return tokenJson.access_token;
}

async function refreshWithToken(refreshToken, settings) {
  const scope = "openid profile offline_access User.Read Calendars.ReadWrite";
  const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(settings.tenantId)}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: settings.clientId,
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope
  });

  const tokenResponse = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body
  });

  const tokenJson = await tokenResponse.json().catch(() => null);
  if (!tokenResponse.ok || !tokenJson?.access_token) {
    throw new Error(tokenJson?.error_description || "Token refresh failed.");
  }

  const expiresAt = Date.now() + (tokenJson.expires_in || 3600) * 1000;
  await chrome.storage.local.set({
    accessToken: tokenJson.access_token,
    accessTokenExpiresAt: expiresAt,
    ...(tokenJson.refresh_token ? { refreshToken: tokenJson.refresh_token } : {})
  });

  return tokenJson.access_token;
}

async function getSettings() {
  const data = await chrome.storage.sync.get(["tenantId", "clientId", "objectId"]);
  return {
    tenantId: data.tenantId || DEFAULT_SETTINGS.tenantId,
    clientId: data.clientId || DEFAULT_SETTINGS.clientId,
    objectId: data.objectId || DEFAULT_SETTINGS.objectId
  };
}

function generateRandomString(length) {
  const charset = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~";
  const values = new Uint8Array(length);
  crypto.getRandomValues(values);
  let out = "";
  for (let i = 0; i < values.length; i += 1) {
    out += charset[values[i] % charset.length];
  }
  return out;
}

async function sha256base64url(value) {
  const bytes = new TextEncoder().encode(value);
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  const arr = Array.from(new Uint8Array(digest));
  const base64 = btoa(String.fromCharCode(...arr));
  return base64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}
