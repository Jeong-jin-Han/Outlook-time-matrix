import { CLIENT_ID, TENANT_ID, OBJECT_ID } from "./config.js";

const tenantIdEl = document.getElementById("tenantId");
const clientIdEl = document.getElementById("clientId");
const objectIdEl = document.getElementById("objectId");
const saveBtn = document.getElementById("saveBtn");
const statusEl = document.getElementById("status");
const redirectUriEl = document.getElementById("redirectUri");

const DEFAULT_SETTINGS = {
  tenantId: TENANT_ID,
  clientId: CLIENT_ID,
  objectId: OBJECT_ID
};

init();

async function init() {
  const settings = await chrome.storage.sync.get(["tenantId", "clientId", "objectId"]);
  tenantIdEl.value = settings.tenantId || DEFAULT_SETTINGS.tenantId;
  clientIdEl.value = settings.clientId || DEFAULT_SETTINGS.clientId;
  objectIdEl.value = settings.objectId || DEFAULT_SETTINGS.objectId;
  redirectUriEl.textContent = `Redirect URI: ${chrome.identity.getRedirectURL("microsoft")}`;
}

saveBtn.addEventListener("click", async () => {
  await chrome.storage.sync.set({
    tenantId: tenantIdEl.value.trim(),
    clientId: clientIdEl.value.trim(),
    objectId: objectIdEl.value.trim()
  });
  statusEl.textContent = "Saved.";
});
