const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

// --- DOM refs ---
const statusEl        = document.getElementById("status");
const signInBtn       = document.getElementById("signInBtn");
const calendarSelect  = document.getElementById("calendarSelect");
const loadBtn         = document.getElementById("loadBtn");
const refreshBtn      = document.getElementById("refreshBtn");
const outlookBtn      = document.getElementById("outlookBtn");
const viewToggleBtn   = document.getElementById("viewToggleBtn");
const settingsBtn     = document.getElementById("settingsBtn");
const settingsOverlay = document.getElementById("settingsOverlay");
const settingsClose   = document.getElementById("settingsClose");
const settingsSaveBtn = document.getElementById("settingsSaveBtn");
const categoryListEl  = document.getElementById("categoryList");
const helpBtn         = document.getElementById("helpBtn");
const helpOverlay     = document.getElementById("helpOverlay");
const helpClose       = document.getElementById("helpClose");
const cardTemplate    = document.getElementById("cardTemplate");
const commonRegion    = document.getElementById("commonRegion");
const commonToggle    = document.getElementById("commonToggle");
const quadrantsEl     = document.getElementById("quadrants");
const matrixWrap      = document.getElementById("matrixWrap");

const bucketIds = ["Q1", "Q2", "Q3", "Q4", "Common-ToDo", "Common-Finished"];
const matrixState = new Map();
let dragOverInfo = null; // { targetId, position: "before"|"after" }

// calendar id → Graph API color string
const calendarColorMap = new Map();

// category name → CSS hex (populated from masterCategories API)
const categoryColorMap = new Map();

// Microsoft user ID — used to namespace storage keys per account
let currentUserId = null;

// Name of the category that marks an event as finished (user-configured)
let finishedCategoryName = null;

// Outlook preset index → CSS hex color
const PRESET_COLOR = {
  preset0:  "#e74856", // red
  preset1:  "#f7630c", // orange
  preset2:  "#ca5010", // peach
  preset3:  "#f2c600", // yellow
  preset4:  "#16c60c", // green
  preset5:  "#00b7c3", // teal
  preset6:  "#847545", // olive
  preset7:  "#0078d4", // blue
  preset8:  "#8764b8", // purple
  preset9:  "#c239b3", // cranberry
  preset10: "#4a5568", // steel
  preset11: "#2d3748", // dark steel
  preset12: "#97989a", // gray
  preset13: "#515151", // dark gray
  preset14: "#1a1a1a", // black
  preset15: "#750b1c", // dark red
  preset16: "#e3008c", // hot pink
  preset17: "#001d4a", // navy
  preset18: "#4e2600", // cocoa
  preset19: "#3e2723", // umber
  preset20: "#00b294", // seafoam
  preset21: "#008272", // dark teal
  preset22: "#498205", // forest
  preset23: "#7719aa", // grape
  preset24: "#9a0089", // lavender
};

// --- Restore persisted UI state ---
chrome.storage.local.get(["sidebarCollapsed", "isListView"], ({ sidebarCollapsed, isListView }) => {
  if (sidebarCollapsed) {
    commonRegion.classList.add("collapsed");
    commonToggle.textContent = "\u25B6"; // ▶
  }
  if (isListView) {
    quadrantsEl.classList.add("list-view");
    matrixWrap.classList.add("list-mode");
    viewToggleBtn.innerHTML = "&#9776;";
  }
});

// --- Sidebar toggle ---
commonToggle.addEventListener("click", () => {
  const collapsed = commonRegion.classList.toggle("collapsed");
  commonToggle.textContent = collapsed ? "\u25B6" : "\u25C0"; // ▶ or ◀
  chrome.storage.local.set({ sidebarCollapsed: collapsed });
});

// --- Help overlay ---
helpBtn.addEventListener("click", () => helpOverlay.classList.toggle("hidden"));
helpClose.addEventListener("click", () => helpOverlay.classList.add("hidden"));
helpOverlay.addEventListener("click", (e) => {
  if (e.target === helpOverlay) helpOverlay.classList.add("hidden");
});

// --- Settings overlay ---
settingsBtn.addEventListener("click", () => showSettingsModal());
settingsClose.addEventListener("click", () => settingsOverlay.classList.add("hidden"));
settingsOverlay.addEventListener("click", (e) => {
  if (e.target === settingsOverlay) settingsOverlay.classList.add("hidden");
});

settingsSaveBtn.addEventListener("click", async () => {
  const selected = categoryListEl.querySelector(".category-item.selected");
  if (!selected) { setStatus("Select a category first."); return; }
  finishedCategoryName = selected.dataset.name;
  const calKey = `finishedCategory_${currentUserId}_${calendarSelect.value}`;
  await chrome.storage.local.set({ [calKey]: finishedCategoryName });
  settingsOverlay.classList.add("hidden");
  try {
    const token = await signIn(false);
    await loadEvents(token, true);
  } catch (e) {
    setStatus(e.message || String(e));
  }
});

function showSettingsModal() {
  categoryListEl.innerHTML = "";
  if (categoryColorMap.size === 0) {
    categoryListEl.innerHTML = "<p style='font-size:11px;color:var(--text-secondary)'>No categories found. Load calendar first.</p>";
  } else {
    for (const [name, css] of categoryColorMap) {
      const item = document.createElement("div");
      item.className = "category-item" + (name === finishedCategoryName ? " selected" : "");
      item.dataset.name = name;

      const swatch = document.createElement("span");
      swatch.className = "cat-swatch";
      swatch.style.background = css || "#888";

      const label = document.createElement("span");
      label.className = "cat-name";
      label.textContent = name;

      item.appendChild(swatch);
      item.appendChild(label);
      item.addEventListener("click", () => {
        categoryListEl.querySelectorAll(".category-item").forEach(el => el.classList.remove("selected"));
        item.classList.add("selected");
      });
      categoryListEl.appendChild(item);
    }
  }
  settingsOverlay.classList.remove("hidden");
}

// --- Quadrant zoom (click empty space to expand/collapse) ---
(function initQuadrantZoom() {
  const topSpans  = document.querySelectorAll(".axis.top span");
  const leftSpans = document.querySelectorAll(".axis.left span");

  // Q1=urgent+important, Q2=notUrgent+important, Q3=urgent+notImportant, Q4=notUrgent+notImportant
  const AXIS = { Q1: [0,0], Q2: [1,0], Q3: [0,1], Q4: [1,1] };

  function setHighlight(zone) {
    topSpans.forEach(s  => s.classList.remove("axis-zoom-highlight"));
    leftSpans.forEach(s => s.classList.remove("axis-zoom-highlight"));
    if (!zone) return;
    const [t, l] = AXIS[zone] ?? [];
    if (t !== undefined) topSpans[t]?.classList.add("axis-zoom-highlight");
    if (l !== undefined) leftSpans[l]?.classList.add("axis-zoom-highlight");
  }

  for (const quadrant of document.querySelectorAll(".quadrant")) {
    quadrant.addEventListener("click", (e) => {
      if (e.target.closest(".task")) return;
      if (e.target.closest(".task-open-btn")) return;
      if (quadrantsEl.classList.contains("zoomed")) {
        quadrantsEl.classList.remove("zoomed");
        quadrant.classList.remove("zoom-active");
        setHighlight(null);
      } else {
        quadrantsEl.classList.add("zoomed");
        quadrant.classList.add("zoom-active");
        setHighlight(quadrant.dataset.zone);
      }
    });
  }
})();

// --- View toggle (2×2 grid ↔ 4×1 list) ---
viewToggleBtn.addEventListener("click", () => {
  const isList = quadrantsEl.classList.toggle("list-view");
  matrixWrap.classList.toggle("list-mode", isList);
  viewToggleBtn.innerHTML = isList ? "&#9776;" : "&#8862;";
  // Exit zoom when switching views
  quadrantsEl.classList.remove("zoomed");
  document.querySelectorAll(".quadrant.zoom-active").forEach(q => q.classList.remove("zoom-active"));
  document.querySelectorAll(".axis-zoom-highlight").forEach(s => s.classList.remove("axis-zoom-highlight"));
  chrome.storage.local.set({ isListView: isList });
});

// --- Open Outlook Calendar ---
outlookBtn.addEventListener("click", () => {
  chrome.tabs.create({ url: "https://outlook.office.com/calendar" });
});

// --- Login / Logout toggle ---
signInBtn.addEventListener("click", async () => {
  if (signInBtn.textContent === "LOGOUT") {
    await chrome.storage.local.remove(["accessToken", "accessTokenExpiresAt", "refreshToken"]);
    signInBtn.textContent = "LOGIN";
    calendarSelect.innerHTML = "";
    matrixState.clear();
    renderMatrix();
    setStatus("Signed out.");
    return;
  }
  try {
    const token = await signIn(true);
    signInBtn.textContent = "LOGOUT";
    await fetchUserId(token);
    await loadCalendars(token);
    await loadMasterCategories(token);
    await loadEvents(token, true);
    if (!finishedCategoryName) showSettingsModal();
  } catch (error) {
    setStatus(error.message || String(error));
  }
});

// --- Load (smart) ---
loadBtn.addEventListener("click", async () => {
  try {
    const token = await signIn(false);
    await fetchUserId(token);
    const prevCalId = calendarSelect.value;
    await loadCalendars(token);
    // Restore previously selected calendar if it still exists
    if (prevCalId && [...calendarSelect.options].some(o => o.value === prevCalId)) {
      calendarSelect.value = prevCalId;
    }
    await loadEvents(token, true);
  } catch (error) {
    setStatus(error.message || String(error));
  }
});

// --- Refresh (fresh) — with confirmation ---
refreshBtn.addEventListener("click", async () => {
  const confirmed = window.confirm(
    "Refresh will reset ALL task placements.\n\n" +
    "Every task will be moved back to the ToDo list, " +
    "and your current matrix arrangement will be lost.\n\n" +
    "Continue?"
  );
  if (!confirmed) return;
  try {
    const token = await signIn(false);
    await loadEvents(token, false);
  } catch (error) {
    setStatus(error.message || String(error));
  }
});

// --- Bucket drag-drop listeners ---
for (const id of bucketIds) {
  const bucket = document.getElementById(id);
  if (!bucket) continue;
  bucket.addEventListener("dragover", (event) => {
    event.preventDefault();
    bucket.classList.add("drop-active");
  });
  bucket.addEventListener("dragleave", (event) => {
    if (!bucket.contains(event.relatedTarget)) {
      bucket.classList.remove("drop-active");
    }
  });
  bucket.addEventListener("drop", (event) => {
    event.preventDefault();
    bucket.classList.remove("drop-active");
    clearInsertIndicators();
    const eventId = event.dataTransfer.getData("text/plain");
    if (!eventId || !matrixState.has(eventId)) return;
    const data = matrixState.get(eventId);
    const zoneItems = [...matrixState.values()]
      .filter(item => item.zone === id && item.id !== eventId)
      .sort((a, b) => a.order - b.order);
    zoneItems.forEach((item, i) => { item.order = i; });
    data.zone = id;
    data.manuallyFinished = (id === "Common-Finished");
    data.order = zoneItems.length;
    dragOverInfo = null;
    renderMatrix();
    saveMatrixState();
  });
}

bootstrap();

// --- Bootstrap ---
async function bootstrap() {
  try {
    const token = await signIn(false);
    signInBtn.textContent = "LOGOUT";
    await fetchUserId(token);
    await loadCalendars(token);
    // Load per-account per-calendar finished category setting
    const calKey = `finishedCategory_${currentUserId}_${calendarSelect.value}`;
    const stored = await chrome.storage.local.get(calKey);
    if (stored[calKey]) finishedCategoryName = stored[calKey];
    await loadMasterCategories(token);
    await loadEvents(token, true);
    // First-time: show settings if no category selected yet
    if (!finishedCategoryName) showSettingsModal();
  } catch {
    setStatus("Please LOGIN to get started.");
  }
}

// --- Auth ---
async function signIn(forcePrompt) {
  // If not forcing a new login, try the cached token directly from storage
  // so we don't need to wake the dormant service worker at all.
  if (!forcePrompt) {
    const cached = await chrome.storage.local.get(["accessToken", "accessTokenExpiresAt"]);
    if (
      cached.accessToken &&
      cached.accessTokenExpiresAt &&
      Date.now() < cached.accessTokenExpiresAt - 60_000
    ) {
      setStatus("Signed in.");
      return cached.accessToken;
    }
  }

  // Need the service worker for interactive login or token refresh.
  // Retry once — the SW may just need a moment to restart after being dormant.
  let result;
  try {
    result = await chrome.runtime.sendMessage({ type: "auth", forcePrompt });
  } catch {
    await new Promise(r => setTimeout(r, 500));
    result = await chrome.runtime.sendMessage({ type: "auth", forcePrompt });
  }

  if (!result?.ok) throw new Error(result?.error || "Authentication failed");
  const stored = await chrome.storage.local.get("accessToken");
  if (!stored.accessToken) throw new Error("Token not found in storage after auth.");
  setStatus("Signed in.");
  return stored.accessToken;
}

// --- Fetch current Microsoft user ID ---
async function fetchUserId(token) {
  try {
    const res = await graphFetch(`${GRAPH_BASE}/me?$select=id`, token);
    currentUserId = res.id || null;
  } catch {
    currentUserId = null;
  }
}

// --- Calendar list ---
async function loadCalendars(token) {
  setStatus("Loading calendars...");
  const calendars = await listCalendarsFromGraph(token);
  if (!calendars.length) throw new Error("No calendar found for this account.");

  calendarSelect.innerHTML = "";
  const calIdKey = `selectedCalendarId_${currentUserId}`;
  const calNameKey = `selectedCalendarName_${currentUserId}`;
  const stored = await chrome.storage.sync.get([calIdKey, calNameKey]);

  for (const cal of calendars) {
    const option = document.createElement("option");
    option.value = cal.id;
    option.textContent = cal.groupName ? `${cal.groupName} → ${cal.name}` : cal.name;
    calendarSelect.appendChild(option);
    calendarColorMap.set(cal.id, cal.color || "auto");
  }

  let targetId = stored[calIdKey];
  if (!targetId && stored[calNameKey]) {
    targetId = calendars.find((c) => c.name === stored[calNameKey])?.id;
  }
  if (!targetId) {
    targetId = calendars.find((c) => c.name.toLowerCase() === "me")?.id || calendars[0].id;
  }
  calendarSelect.value = targetId;

  calendarSelect.onchange = async () => {
    await chrome.storage.sync.set({
      [calIdKey]: calendarSelect.value,
      [calNameKey]: calendarSelect.selectedOptions[0]?.textContent || ""
    });
  };

  await chrome.storage.sync.set({
    [calIdKey]: calendarSelect.value,
    [calNameKey]: calendarSelect.selectedOptions[0]?.textContent || ""
  });
}

// --- Master categories (event colors) ---
async function loadMasterCategories(token) {
  try {
    const res = await graphFetch(`${GRAPH_BASE}/me/outlook/masterCategories`, token);
    categoryColorMap.clear();
    for (const cat of res.value || []) {
      const css = PRESET_COLOR[cat.color] ?? null;
      categoryColorMap.set(cat.displayName, css);
      // DEBUG: open DevTools (right-click panel → Inspect) to see these
      console.log(`[Category] "${cat.displayName}" preset="${cat.color}" css="${css}"`);
    }
  } catch (e) {
    console.warn("[MasterCategories] failed:", e.message);
  }
}

// --- Load events ---
async function loadEvents(token, smart) {
  if (!calendarSelect.value) {
    setStatus("Select a calendar first.");
    return;
  }

  // Load per-account per-calendar finished category
  const calKey = `finishedCategory_${currentUserId}_${calendarSelect.value}`;
  const storedCat = await chrome.storage.local.get(calKey);
  finishedCategoryName = storedCat[calKey] || null;

  // Date range: Monday of current week → +14 days
  const now = new Date();
  const day = now.getDay(); // 0=Sun, 1=Mon, …
  const daysToMonday = day === 0 ? -6 : 1 - day;
  const monday = new Date(now);
  monday.setDate(now.getDate() + daysToMonday);
  monday.setHours(0, 0, 0, 0);
  const to = new Date(monday.getTime() + 14 * 24 * 60 * 60 * 1000);

  setStatus("Loading 2-week events...");
  // Always refresh categories so the color filter works even if bootstrap failed
  await loadMasterCategories(token);

  const events = await fetchCalendarEventsFromGraph(
    token, calendarSelect.value, monday.toISOString(), to.toISOString()
  );

  if (smart) {
    await smartBuildState(events, now);
  } else {
    freshBuildState(events);
  }

  renderMatrix();
  await saveMatrixState();
  setStatus(`Loaded ${events.length} events from ${calendarSelect.selectedOptions[0]?.textContent || "calendar"}.`);
}

// Smart: restore placements, auto-finish past events, clean stale entries
async function smartBuildState(events, now) {
  const placementsKey = `matrixPlacements_${currentUserId}_${calendarSelect.value}`;
  const stored = await chrome.storage.local.get(placementsKey);
  const placements = stored[placementsKey] || {};

  const currentIds = new Set(events.map(e => e.id));
  for (const id of Object.keys(placements)) {
    if (!currentIds.has(id)) delete placements[id];
  }

  matrixState.clear();
  events.forEach((event, index) => {
    const saved = placements[event.id];
    const endDate = new Date(event.end?.dateTime || event.end?.date || 0);
    const isPast = endDate < now;
    const isColorFinished = isEventFinished(event);

    let zone;
    if (isColorFinished || isPast) {
      // Always auto-finish — overrides any saved state
      zone = "Common-Finished";
    } else if (saved?.zone === "Common-Finished" && saved?.manual === true) {
      // User explicitly moved/checked this item as finished → respect that
      zone = "Common-Finished";
    } else if (saved?.zone && saved.zone !== "Common-Finished") {
      // Restore saved placement in matrix quadrants or ToDo
      zone = saved.zone;
    } else {
      // New event, or was previously auto-finished but no longer qualifies
      zone = "Common-ToDo";
    }

    matrixState.set(event.id, {
      id: event.id,
      subject: event.subject || "(No title)",
      start: event.start?.dateTime || event.start?.date,
      end: event.end?.dateTime || event.end?.date,
      webLink: event.webLink || null,
      zone,
      colorFinished: isColorFinished,
      manuallyFinished: zone === "Common-Finished" && !isColorFinished && !isPast,
      order: saved?.order ?? (1_000_000 + index)
    });
  });
}

// Fresh: ignore all saved placements, everything goes to Common-ToDo
function freshBuildState(events) {
  matrixState.clear();
  events.forEach((event, index) => {
    const isColorFinished = isEventFinished(event);
    matrixState.set(event.id, {
      id: event.id,
      subject: event.subject || "(No title)",
      start: event.start?.dateTime || event.start?.date,
      end: event.end?.dateTime || event.end?.date,
      webLink: event.webLink || null,
      zone: isColorFinished ? "Common-Finished" : "Common-ToDo",
      colorFinished: isColorFinished,
      order: index
    });
  });
}

// Returns true if the event has the user-configured finished category
function isEventFinished(event) {
  if (!finishedCategoryName) return false;
  return (event.categories || []).includes(finishedCategoryName);
}

// --- State persistence ---
async function saveMatrixState() {
  const placements = {};
  for (const [id, entry] of matrixState) {
    placements[id] = { zone: entry.zone, order: entry.order, manual: entry.manuallyFinished ?? false };
  }
  const placementsKey = `matrixPlacements_${currentUserId}_${calendarSelect.value}`;
  await chrome.storage.local.set({ [placementsKey]: placements });
}

// --- Graph API ---
async function graphFetch(url, token, init = {}) {
  const response = await fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(init.headers || {})
    }
  });
  if (!response.ok) {
    const body = await response.json().catch(() => null);
    throw new Error(body?.error?.message || `Graph API failed: ${response.status}`);
  }
  return response.json();
}

async function listCalendarsFromGraph(token) {
  const groupedUrl = `${GRAPH_BASE}/me/calendarGroups?$select=id,name&$expand=calendars($select=id,name,color,canEdit,isDefaultCalendar)`;
  try {
    const grouped = await graphFetch(groupedUrl, token);
    const result = [];
    for (const group of grouped.value || []) {
      for (const cal of group.calendars || []) {
        result.push({ ...cal, groupId: group.id, groupName: group.name || "" });
      }
    }
    if (result.length > 0) return result;
  } catch {
    // fall through
  }
  const fallback = await graphFetch(`${GRAPH_BASE}/me/calendars?$select=id,name,color,canEdit,isDefaultCalendar`, token);
  return (fallback.value || []).map((c) => ({ ...c, groupId: "", groupName: "" }));
}

async function fetchCalendarEventsFromGraph(token, calendarId, fromIso, toIso) {
  const params = new URLSearchParams({
    startDateTime: fromIso,
    endDateTime: toIso,
    $top: "200",
    $select: "id,subject,start,end,importance,categories,isAllDay,bodyPreview,showAs,webLink",
    $orderby: "start/dateTime"
  });
  const url = `${GRAPH_BASE}/me/calendars/${encodeURIComponent(calendarId)}/calendarView?${params.toString()}`;
  const response = await graphFetch(url, token, {
    headers: { Prefer: 'outlook.timezone="Asia/Seoul"' }
  });
  return response.value || [];
}

// --- Outlook color sync ---

// Sets or removes the finished category on an Outlook event, then reloads
async function syncEventColor(eventId, markFinished) {
  if (!finishedCategoryName) {
    setStatus("Select a finished category in Settings first.");
    return;
  }
  try {
    const token = await signIn(false);
    const ev = await graphFetch(
      `${GRAPH_BASE}/me/events/${encodeURIComponent(eventId)}?$select=categories`,
      token
    );
    let categories = ev.categories || [];
    if (markFinished) {
      // Put finished category first so it becomes the displayed color in Outlook
      categories = [finishedCategoryName, ...categories.filter(c => c !== finishedCategoryName)];
    } else {
      categories = categories.filter(c => c !== finishedCategoryName);
    }
    await graphFetch(`${GRAPH_BASE}/me/events/${encodeURIComponent(eventId)}`, token, {
      method: "PATCH",
      body: JSON.stringify({ categories })
    });
    await loadEvents(token, true);
  } catch (e) {
    setStatus("Color sync failed: " + (e.message || String(e)));
  }
}

// --- Rendering ---
function clearInsertIndicators() {
  document.querySelectorAll(".task.insert-before, .task.insert-after").forEach(el => {
    el.classList.remove("insert-before", "insert-after");
  });
}

function renderMatrix() {
  for (const bucketId of bucketIds) {
    const el = document.getElementById(bucketId);
    if (el) el.innerHTML = "";
  }

  const sorted = [...matrixState.values()].sort((a, b) => a.order - b.order);

  for (const entry of sorted) {
    const node = cardTemplate.content.firstElementChild.cloneNode(true);
    node.dataset.eventId = entry.id;
    node.querySelector(".subject").textContent = entry.subject;
    node.querySelector(".meta").textContent = fmtRange(entry.start, entry.end);

    // Finished checkbox
    const doneCheck = node.querySelector(".done-check");
    doneCheck.checked = entry.zone === "Common-Finished";

    doneCheck.addEventListener("mousedown", (e) => e.stopPropagation());
    doneCheck.addEventListener("change", (e) => {
      e.stopPropagation();
      if (doneCheck.checked) {
        const peers = [...matrixState.values()]
          .filter(i => i.zone === "Common-Finished" && i.id !== entry.id)
          .sort((a, b) => a.order - b.order);
        entry.zone = "Common-Finished";
        entry.manuallyFinished = true;
        entry.order = peers.length;
        syncEventColor(entry.id, true);
      } else {
        const peers = [...matrixState.values()]
          .filter(i => i.zone === "Common-ToDo" && i.id !== entry.id)
          .sort((a, b) => a.order - b.order);
        entry.zone = "Common-ToDo";
        entry.manuallyFinished = false;
        entry.order = peers.length;
        syncEventColor(entry.id, false);
      }
      renderMatrix();
      saveMatrixState();
    });

    // ↗ Open-in-Outlook button
    const openBtn = node.querySelector(".task-open-btn");
    if (openBtn) {
      openBtn.addEventListener("click", async (ev) => {
        ev.stopPropagation();
        ev.preventDefault();
        try {
          const allTabs = await chrome.tabs.query({});
          const outlookTab = allTabs.find(t => t.url && (
            t.url.startsWith("https://outlook.live.com/") ||
            t.url.startsWith("https://outlook.office.com/") ||
            t.url.startsWith("https://outlook.office365.com/")
          ));
          if (!outlookTab) {
            setStatus("\u{1F5D3} 버튼으로 Outlook을 먼저 열어주세요.");
            return;
          }
          if (entry.webLink) {
            await chrome.tabs.update(outlookTab.id, { url: entry.webLink, active: true });
          } else {
            await chrome.tabs.update(outlookTab.id, { active: true });
            setStatus("Outlook에서 해당 항목을 직접 찾아주세요.");
          }
        } catch (e) {
          setStatus(e.message || "Outlook 열기 실패");
        }
      });
    } else {
      console.warn("[renderMatrix] .task-open-btn not found — reload extension after HTML change");
    }

    // Drag start
    node.addEventListener("dragstart", (event) => {
      event.dataTransfer.setData("text/plain", entry.id);
      event.dataTransfer.effectAllowed = "move";
    });

    // Drag over (positional indicator)
    node.addEventListener("dragover", (event) => {
      event.preventDefault();
      event.stopPropagation();
      const rect = node.getBoundingClientRect();
      const position = event.clientY < rect.top + rect.height / 2 ? "before" : "after";
      if (dragOverInfo?.targetId !== entry.id || dragOverInfo?.position !== position) {
        clearInsertIndicators();
        dragOverInfo = { targetId: entry.id, position };
        node.classList.add(position === "before" ? "insert-before" : "insert-after");
      }
    });

    node.addEventListener("dragleave", (event) => {
      if (!node.contains(event.relatedTarget)) {
        node.classList.remove("insert-before", "insert-after");
        if (dragOverInfo?.targetId === entry.id) dragOverInfo = null;
      }
    });

    // Drop on card (insert before/after)
    node.addEventListener("drop", (event) => {
      event.preventDefault();
      event.stopPropagation();
      clearInsertIndicators();
      document.getElementById(entry.zone)?.classList.remove("drop-active");

      const sourceId = event.dataTransfer.getData("text/plain");
      if (!sourceId || !matrixState.has(sourceId) || sourceId === entry.id) {
        dragOverInfo = null;
        return;
      }

      const target = matrixState.get(entry.id);
      const source = matrixState.get(sourceId);
      const targetZone = target.zone;

      const zoneItems = [...matrixState.values()]
        .filter(item => item.zone === targetZone && item.id !== sourceId)
        .sort((a, b) => a.order - b.order);

      const targetIndex = zoneItems.findIndex(item => item.id === entry.id);
      const insertIndex = dragOverInfo?.position === "before" ? targetIndex : targetIndex + 1;
      zoneItems.splice(insertIndex, 0, source);

      source.zone = targetZone;
      source.manuallyFinished = (targetZone === "Common-Finished");
      zoneItems.forEach((item, i) => { item.order = i; });

      dragOverInfo = null;
      renderMatrix();
      saveMatrixState();
    });

    const bucket = document.getElementById(entry.zone);
    if (bucket) bucket.appendChild(node);
  }
}

function fmtDate(value) {
  if (!value) return "?";
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return "?";
  return d.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
}

function fmtRange(start, end) {
  const s = fmtDate(start);
  const e = fmtDate(end);
  return s === e ? s : `${s} → ${e}`;
}

let _statusTimer = null;
function setStatus(text) {
  if (!text) return;
  statusEl.textContent = text;
  statusEl.classList.add("visible");
  clearTimeout(_statusTimer);
  _statusTimer = setTimeout(() => statusEl.classList.remove("visible"), 3000);
}
