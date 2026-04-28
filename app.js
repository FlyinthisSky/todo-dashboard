// ===== CONFIGURATION =====
const CLIENT_ID = "2ca92e37-7ae8-4155-95ab-538a74cda14d";
const REDIRECT_URI = window.location.origin + window.location.pathname;
const REFRESH_INTERVAL = 15;
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

const msalConfig = {
    auth: { clientId: CLIENT_ID, authority: "https://login.microsoftonline.com/consumers", redirectUri: REDIRECT_URI },
    cache: { cacheLocation: "sessionStorage" }
};
const loginRequest = { scopes: ["Tasks.Read", "Tasks.ReadWrite"] };
let msalInstance = null;

// ===== AUTH =====
async function login() {
    if (!msalInstance) { showToast("Initialisation en cours..."); return; }
    try { await msalInstance.loginRedirect(loginRequest); }
    catch (err) { console.error("Login failed:", err); showToast("Erreur de connexion.", true); }
}
function logout() {
    if (refreshIntervalId) { clearInterval(refreshIntervalId); refreshIntervalId = null; }
    if (msalInstance) msalInstance.logoutRedirect({ postLogoutRedirectUri: REDIRECT_URI });
}
async function getAccessToken() {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) throw new Error("No account");
    try {
        return (await msalInstance.acquireTokenSilent({ ...loginRequest, account: accounts[0] })).accessToken;
    } catch (err) { await msalInstance.acquireTokenRedirect(loginRequest); }
}
async function onAuthenticated() {
    document.getElementById("login-screen").style.display = "none";
    document.getElementById("dashboard").style.display = "flex";
    showLoading(true);
    await loadAndRenderTasks();
    showLoading(false);
    updateNavLabel();
    startAutoRefresh();
}

// ===== UTILITY =====
function showLoading(v) { document.getElementById("loading").style.display = v ? "flex" : "none"; }

function showToast(message, isError = false, undoFn = null) {
    const existing = document.querySelector(".toast");
    if (existing) existing.remove();
    const toast = document.createElement("div");
    toast.className = "toast" + (isError ? " error" : "");
    toast.setAttribute("role", "alert");
    toast.setAttribute("aria-live", "assertive");
    const span = document.createElement("span");
    span.textContent = message;
    toast.appendChild(span);
    if (undoFn) {
        const btn = document.createElement("button");
        btn.className = "undo-btn";
        btn.textContent = "Annuler";
        btn.onclick = () => { undoFn(); toast.remove(); };
        toast.appendChild(btn);
    }
    document.body.appendChild(toast);
    setTimeout(() => toast.remove(), undoFn ? 6000 : 4000);
}

function escapeHtml(str) {
    const div = document.createElement("div");
    div.textContent = str;
    return div.innerHTML;
}

function setModalLoading(modal, loading) {
    modal.querySelectorAll(".modal-btn").forEach(btn => {
        btn.disabled = loading;
    });
}

function trapFocus(overlay, modal) {
    const focusableSelector = 'button:not(:disabled), input:not(:disabled), select:not(:disabled), textarea:not(:disabled), [tabindex]:not([tabindex="-1"])';
    const previouslyFocused = document.activeElement;

    function getFocusable() {
        return [...modal.querySelectorAll(focusableSelector)].filter(el => el.offsetParent !== null);
    }

    requestAnimationFrame(() => {
        const els = getFocusable();
        if (els.length > 0) els[0].focus();
    });

    function handleKeydown(e) {
        if (e.key === "Escape") {
            e.preventDefault();
            cleanup();
            overlay.remove();
            return;
        }
        if (e.key !== "Tab") return;
        const els = getFocusable();
        if (els.length === 0) return;
        const first = els[0], last = els[els.length - 1];
        if (e.shiftKey) {
            if (document.activeElement === first) { e.preventDefault(); last.focus(); }
        } else {
            if (document.activeElement === last) { e.preventDefault(); first.focus(); }
        }
    }

    function cleanup() {
        overlay.removeEventListener("keydown", handleKeydown);
        if (previouslyFocused && previouslyFocused.focus) previouslyFocused.focus();
    }

    overlay.addEventListener("keydown", handleKeydown);
    return cleanup;
}

// ===== GRAPH API =====
async function graphFetchWithRetry(url, options = {}, maxRetries = 3) {
    for (let attempt = 0; attempt <= maxRetries; attempt++) {
        const token = await getAccessToken();
        const r = await fetch(url, { ...options, headers: { Authorization: "Bearer " + token, ...options.headers } });
        if (r.status === 429) {
            const retryAfter = parseInt(r.headers.get("Retry-After") || "0", 10);
            const delay = Math.max(retryAfter, 1) * 1000 + attempt * 500;
            if (attempt < maxRetries) {
                console.warn("429 rate limited, retry " + (attempt + 1) + " in " + delay + "ms: " + url.split("?")[0]);
                await new Promise(resolve => setTimeout(resolve, delay));
                continue;
            }
        }
        return r;
    }
}
async function graphGet(url) {
    const r = await graphFetchWithRetry(url);
    if (!r.ok) throw new Error("Graph API error: " + r.status);
    return r.json();
}
async function graphPatch(url, body) {
    const r = await graphFetchWithRetry(url, { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    if (!r.ok) throw new Error("Graph PATCH error: " + r.status);
    return r.json();
}
async function graphGetAll(url) {
    let results = [];
    let nextUrl = url;
    while (nextUrl) {
        const data = await graphGet(nextUrl);
        results = results.concat(data.value);
        nextUrl = data["@odata.nextLink"] || null;
    }
    return results;
}
async function graphDelete(url) {
    const r = await graphFetchWithRetry(url, { method: "DELETE" });
    if (!r.ok) throw new Error("Graph DELETE error: " + r.status);
}
async function graphPost(url, body) {
    const r = await graphFetchWithRetry(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    if (!r.ok) throw new Error("Graph POST error: " + r.status);
    return r.json();
}
async function fetchLists() {
    let results = [];
    let nextUrl = GRAPH_BASE + "/me/todo/lists/delta";
    while (nextUrl) {
        const data = await graphGet(nextUrl);
        results = results.concat(data.value);
        nextUrl = data["@odata.nextLink"] || null;
        if (data["@odata.deltaLink"]) break;
    }
    return results.filter(l => !l["@removed"]);
}
async function fetchTasks(listId) {
    return graphGetAll(GRAPH_BASE + "/me/todo/lists/" + listId + "/tasks?$filter=status eq 'notStarted'");
}
async function fetchCompletedTasks(listId) {
    return graphGetAll(GRAPH_BASE + "/me/todo/lists/" + listId + "/tasks?$filter=status eq 'completed'");
}
async function deleteTask(listId, taskId) {
    return graphDelete(GRAPH_BASE + "/me/todo/lists/" + listId + "/tasks/" + taskId);
}
async function completeTask(listId, taskId) {
    return graphPatch(GRAPH_BASE + "/me/todo/lists/" + listId + "/tasks/" + taskId, { status: "completed" });
}
async function uncompleteTask(listId, taskId) {
    return graphPatch(GRAPH_BASE + "/me/todo/lists/" + listId + "/tasks/" + taskId, { status: "notStarted" });
}
async function updateTask(listId, taskId, updates) {
    return graphPatch(GRAPH_BASE + "/me/todo/lists/" + listId + "/tasks/" + taskId, updates);
}

// ===== COLOR MAPPING =====
const LIST_COLORS = {
    "ménage":       { gradient: "linear-gradient(135deg, #FDFD96, #f0e68c)", text: "#333" },
    "menage":       { gradient: "linear-gradient(135deg, #FDFD96, #f0e68c)", text: "#333" },
    "boulot":       { gradient: "linear-gradient(135deg, #FF6B6B, #ee5a5a)", text: "#fff" },
    "courses":      { gradient: "linear-gradient(135deg, #55efc4, #00b894)", text: "#1a1a2e" },
    "course":       { gradient: "linear-gradient(135deg, #55efc4, #00b894)", text: "#1a1a2e" },
    "projet perso": { gradient: "linear-gradient(135deg, #48dbfb, #0abde3)", text: "#fff" },
    "sortie":       { gradient: "linear-gradient(135deg, #a29bfe, #6c5ce7)", text: "#fff" },
};
const EXTRA_COLORS = [
    { gradient: "linear-gradient(135deg, #ffa502, #e67e22)", text: "#fff" },
    { gradient: "linear-gradient(135deg, #fd79a8, #e84393)", text: "#fff" },
    { gradient: "linear-gradient(135deg, #fdcb6e, #f39c12)", text: "#333" },
    { gradient: "linear-gradient(135deg, #74b9ff, #0984e3)", text: "#fff" },
    { gradient: "linear-gradient(135deg, #dfe6e9, #b2bec3)", text: "#333" },
];
let dynamicColorIndex = 0;
const dynamicColorMap = {};
function getListColor(listName) {
    const key = listName.toLowerCase().trim();
    if (LIST_COLORS[key]) return LIST_COLORS[key];
    if (dynamicColorMap[key]) return dynamicColorMap[key];
    dynamicColorMap[key] = EXTRA_COLORS[dynamicColorIndex % EXTRA_COLORS.length];
    dynamicColorIndex++;
    return dynamicColorMap[key];
}

async function loadColors() {
    try {
        const r = await fetch('colors.json');
        if (!r.ok) return;
        const data = await r.json();
        if (data.lists) {
            Object.keys(LIST_COLORS).forEach(k => delete LIST_COLORS[k]);
            Object.entries(data.lists).forEach(([key, val]) => {
                LIST_COLORS[key.toLowerCase().trim()] = val;
            });
        }
        if (data.extras) {
            EXTRA_COLORS.length = 0;
            data.extras.forEach(c => EXTRA_COLORS.push(c));
        }
        if (data.projects) {
            PROJECT_COLORS.length = 0;
            data.projects.forEach(c => PROJECT_COLORS.push(c));
        }
    } catch (e) {
        console.log("colors.json not found, using defaults.");
    }
}

// ===== PROJECT TAG COLORS =====
const PROJECT_COLORS = [
    "#e17055", "#00cec9", "#fdcb6e", "#6c5ce7", "#e84393",
    "#00b894", "#0984e3", "#d63031", "#ffeaa7", "#a29bfe",
];
const projectColorMap = {};
let projectColorIndex = 0;
let customTagColors = {};
let customTagTextColors = {};
let isRenamingTag = false;

function loadCustomTagColors() {
    try {
        customTagColors = JSON.parse(localStorage.getItem("customTagColors") || "{}");
        Object.entries(customTagColors).forEach(([tag, color]) => {
            projectColorMap[tag.toLowerCase()] = color;
        });
    } catch (e) { customTagColors = {}; }
}

function saveCustomTagColors() {
    localStorage.setItem("customTagColors", JSON.stringify(customTagColors));
}

function getProjectColor(tag) {
    const key = tag.toLowerCase();
    if (projectColorMap[key]) return projectColorMap[key];
    projectColorMap[key] = PROJECT_COLORS[projectColorIndex % PROJECT_COLORS.length];
    projectColorIndex++;
    return projectColorMap[key];
}

function setProjectColor(tag, color) {
    const key = tag.toLowerCase();
    projectColorMap[key] = color;
    customTagColors[key] = color;
    saveCustomTagColors();
}

function getTagTextColor(tag) {
    return customTagTextColors[tag.toLowerCase()] || "#fff";
}
function setTagTextColor(tag, color) {
    customTagTextColors[tag.toLowerCase()] = color;
    saveCustomTagTextColors();
}
function loadCustomTagTextColors() {
    try { customTagTextColors = JSON.parse(localStorage.getItem("customTagTextColors") || "{}"); }
    catch (e) { customTagTextColors = {}; }
}
function saveCustomTagTextColors() {
    localStorage.setItem("customTagTextColors", JSON.stringify(customTagTextColors));
}

function extractProjectTags(title) {
    const matches = title.match(/#\w+/g);
    return matches || [];
}

function getAllKnownTags() {
    const tags = new Set();
    allTasks.forEach(({ task }) => {
        extractProjectTags(task.title).forEach(t => tags.add(t.toLowerCase()));
    });
    Object.keys(customTagColors).forEach(t => tags.add(t));
    return [...tags].sort();
}

async function renameTag(oldTag, newTag) {
    if (isRenamingTag) return;
    isRenamingTag = true;
    try {
        oldTag = oldTag.toLowerCase();
        newTag = newTag.toLowerCase();
        if (!newTag.startsWith("#")) newTag = "#" + newTag;
        newTag = newTag.replace(/[^#\w]/g, "");
        if (newTag === "#" || newTag === oldTag) return;

        // Transfer color from old tag to new tag
        const color = getProjectColor(oldTag);
        setProjectColor(newTag, color);
        const txtColor = getTagTextColor(oldTag);
        if (txtColor !== "#fff") setTagTextColor(newTag, txtColor);
        delete projectColorMap[oldTag];
        delete customTagColors[oldTag];
        delete customTagTextColors[oldTag];
        saveCustomTagColors();
        saveCustomTagTextColors();

        // Update all tasks containing the old tag (active + completed)
        const allKnownTasks = allTasks;
        const tasksToUpdate = allKnownTasks.filter(({ task }) =>
            extractProjectTags(task.title).some(t => t.toLowerCase() === oldTag)
        );

        const promises = tasksToUpdate.map(({ task, listId }) => {
            const newTitle = task.title.replace(
                new RegExp(oldTag.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), "gi"),
                newTag
            );
            task.title = newTitle;
            return updateTask(listId, task.id, { title: newTitle });
        });

        await Promise.all(promises);
        if (hiddenTags.includes(oldTag)) {
            hiddenTags = hiddenTags.map(t => t === oldTag ? newTag : t);
            saveHiddenTags();
        }
        renderDashboard();
        renderTagPanel();
    } finally {
        isRenamingTag = false;
    }
}

// ===== DATE UTILITIES =====
const DAY_NAMES = ["Dim", "Lun", "Mar", "Mer", "Jeu", "Ven", "Sam"];
const MONTH_NAMES = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"];
let weekOffset = 0;
let viewMode = "week"; // "week" or "month"

function getMonday(date) {
    const d = new Date(date);
    const day = d.getDay();
    d.setDate(d.getDate() - day + (day === 0 ? -6 : 1));
    d.setHours(0, 0, 0, 0);
    return d;
}
function getWeekDays() {
    const center = new Date();
    center.setHours(0, 0, 0, 0);
    center.setDate(center.getDate() + weekOffset * 7 - 3);
    return Array.from({ length: 7 }, (_, i) => { const d = new Date(center); d.setDate(center.getDate() + i); return d; });
}
function getViewMonth() {
    const d = new Date();
    d.setDate(1);
    d.setMonth(d.getMonth() + weekOffset);
    d.setHours(0, 0, 0, 0);
    return d;
}
function getMonthDays(refDate) {
    const year = refDate.getFullYear(), month = refDate.getMonth();
    const first = new Date(year, month, 1);
    const last = new Date(year, month + 1, 0);
    // Start on Monday of the week containing the 1st
    const start = getMonday(first);
    // End on Sunday of the week containing the last day
    const end = new Date(last);
    const endDay = end.getDay();
    if (endDay !== 0) end.setDate(end.getDate() + (7 - endDay));
    end.setHours(0, 0, 0, 0);
    const days = [];
    const cur = new Date(start);
    while (cur <= end) { days.push(new Date(cur)); cur.setDate(cur.getDate() + 1); }
    return days;
}
function navigatePrev() {
    weekOffset--;
    renderDashboard();
    updateNavLabel();
}
function navigateNext() {
    weekOffset++;
    renderDashboard();
    updateNavLabel();
}
function navigateToday() {
    weekOffset = 0;
    renderDashboard();
    updateNavLabel();
}
function switchView(mode) {
    viewMode = mode;
    weekOffset = 0;
    document.querySelectorAll(".view-btn").forEach(b => b.classList.remove("active"));
    const btn = document.querySelector('.view-btn[data-view="' + mode + '"]');
    if (btn) btn.classList.add("active");
    renderDashboard();
    updateNavLabel();
}
function updateNavLabel() {
    const label = document.getElementById("nav-label");
    if (!label) return;
    if (viewMode === "week") {
        const days = getWeekDays();
        const start = days[0], end = days[6];
        const opts = { day: "numeric", month: "short" };
        if (start.getFullYear() !== end.getFullYear()) {
            label.textContent = start.toLocaleDateString("fr-FR", { ...opts, year: "numeric" }) + " — " + end.toLocaleDateString("fr-FR", { ...opts, year: "numeric" });
        } else if (start.getMonth() !== end.getMonth()) {
            label.textContent = start.toLocaleDateString("fr-FR", opts) + " — " + end.toLocaleDateString("fr-FR", opts) + " " + end.getFullYear();
        } else {
            label.textContent = start.getDate() + " — " + end.toLocaleDateString("fr-FR", opts) + " " + end.getFullYear();
        }
    } else {
        const ref = getViewMonth();
        label.textContent = MONTH_NAMES[ref.getMonth()] + " " + ref.getFullYear();
    }
    // Show/hide "Aujourd'hui" button
    const todayBtn = document.getElementById("nav-today-btn");
    if (todayBtn) todayBtn.style.display = weekOffset === 0 ? "none" : "";
}
function isSameDay(d1, d2) {
    return d1.getFullYear() === d2.getFullYear() && d1.getMonth() === d2.getMonth() && d1.getDate() === d2.getDate();
}
function isOverdue(dueDate) { const t = new Date(); t.setHours(0,0,0,0); return dueDate < t; }
function formatDateForGraph(date) {
    const y = date.getFullYear(), m = String(date.getMonth()+1).padStart(2,"0"), d = String(date.getDate()).padStart(2,"0");
    return y + "-" + m + "-" + d + "T00:00:00.0000000";
}
function formatDateForInput(date) {
    const y = date.getFullYear(), m = String(date.getMonth()+1).padStart(2,"0"), d = String(date.getDate()).padStart(2,"0");
    return y + "-" + m + "-" + d;
}

// ===== LIST FILTER =====
let hiddenLists = JSON.parse(localStorage.getItem("hiddenLists") || "[]");

function saveHiddenLists() { localStorage.setItem("hiddenLists", JSON.stringify(hiddenLists)); }

// ===== TAG FILTER =====
let hiddenTags = JSON.parse(localStorage.getItem("hiddenTags") || "[]");

function saveHiddenTags() { localStorage.setItem("hiddenTags", JSON.stringify(hiddenTags)); }

// ===== SEARCH =====
let searchQuery = "";

// ===== TASK ORDER (per-day reordering) =====
let taskOrder = JSON.parse(localStorage.getItem("taskOrder") || "{}");

function saveTaskOrder() {
    localStorage.setItem("taskOrder", JSON.stringify(taskOrder));
}

function getOrderedTasks(dayTasks, dateKey) {
    const order = taskOrder[dateKey];
    if (!order) return dayTasks;
    const orderMap = new Map(order.map((id, idx) => [id, idx]));
    const ordered = [...dayTasks];
    ordered.sort((a, b) => {
        const posA = orderMap.has(a.task.id) ? orderMap.get(a.task.id) : Infinity;
        const posB = orderMap.has(b.task.id) ? orderMap.get(b.task.id) : Infinity;
        return posA - posB;
    });
    return ordered;
}

function getDropIndex(container, y) {
    const cards = [...container.querySelectorAll(".task-card:not(.dragging)")];
    for (let i = 0; i < cards.length; i++) {
        const rect = cards[i].getBoundingClientRect();
        if (y < rect.top + rect.height / 2) return i;
    }
    return cards.length;
}

function toggleTagFilter() {
    const panel = document.getElementById("tag-filter-panel");
    const wasOpen = panel.classList.contains("open");
    panel.classList.toggle("open");
    document.getElementById("filter-panel").classList.remove("open");
    document.getElementById("tag-panel").classList.remove("open");
    if (panel.classList.contains("open")) renderTagFilterPanel();
    else if (wasOpen) renderDashboard();
}

function renderTagFilterPanel() {
    const panel = document.getElementById("tag-filter-panel");
    panel.innerHTML = '<h3>Tags affichés</h3>';

    const tags = getAllKnownTags();

    if (tags.length === 0) {
        const empty = document.createElement("div");
        empty.className = "tag-panel-empty";
        empty.textContent = "Aucun tag trouvé.";
        panel.appendChild(empty);
        return;
    }

    const actions = document.createElement("div");
    actions.className = "filter-actions";
    const selectAll = document.createElement("button");
    selectAll.textContent = "Tout cocher";
    selectAll.onclick = () => { hiddenTags = []; saveHiddenTags(); renderTagFilterPanel(); renderDashboard(); };
    const selectNone = document.createElement("button");
    selectNone.textContent = "Tout décocher";
    selectNone.onclick = () => { hiddenTags = [...tags, "__no_tag__"]; saveHiddenTags(); renderTagFilterPanel(); renderDashboard(); };
    actions.appendChild(selectAll);
    actions.appendChild(selectNone);
    panel.appendChild(actions);

    // Entry for tasks with no tag
    const noTagItem = document.createElement("label");
    noTagItem.className = "filter-item";
    const noTagCb = document.createElement("input");
    noTagCb.type = "checkbox";
    noTagCb.checked = !hiddenTags.includes("__no_tag__");
    noTagCb.addEventListener("change", () => {
        if (noTagCb.checked) hiddenTags = hiddenTags.filter(t => t !== "__no_tag__");
        else hiddenTags.push("__no_tag__");
        saveHiddenTags();
        renderDashboard();
    });
    const noTagVisual = document.createElement("span");
    noTagVisual.className = "cb";
    noTagItem.appendChild(noTagCb);
    noTagItem.appendChild(noTagVisual);
    noTagItem.appendChild(document.createTextNode("(Sans tag)"));
    panel.appendChild(noTagItem);

    // One entry per known tag
    tags.forEach(tag => {
        const item = document.createElement("label");
        item.className = "filter-item";
        const cb = document.createElement("input");
        cb.type = "checkbox";
        cb.checked = !hiddenTags.includes(tag);
        cb.addEventListener("change", () => {
            if (cb.checked) hiddenTags = hiddenTags.filter(t => t !== tag);
            else hiddenTags.push(tag);
            saveHiddenTags();
            renderDashboard();
        });
        const visual = document.createElement("span");
        visual.className = "cb";
        const swatch = document.createElement("span");
        swatch.style.cssText = "display:inline-block;width:10px;height:10px;border-radius:2px;flex-shrink:0;margin-right:4px;background:" + getProjectColor(tag);
        item.appendChild(cb);
        item.appendChild(visual);
        item.appendChild(swatch);
        item.appendChild(document.createTextNode(tag));
        panel.appendChild(item);
    });
}

function toggleFilter() {
    const panel = document.getElementById("filter-panel");
    const wasOpen = panel.classList.contains("open");
    panel.classList.toggle("open");
    document.getElementById("tag-panel").classList.remove("open");
    document.getElementById("tag-filter-panel").classList.remove("open");
    if (panel.classList.contains("open")) renderFilterPanel();
    else if (wasOpen) loadAndRenderTasks();
}

function renderFilterPanel() {
    const panel = document.getElementById("filter-panel");
    panel.innerHTML = '<h3>Listes affichées</h3>';

    const actions = document.createElement("div");
    actions.className = "filter-actions";
    const selectAll = document.createElement("button");
    selectAll.textContent = "Tout cocher";
    selectAll.onclick = () => { hiddenLists = []; saveHiddenLists(); renderFilterPanel(); renderDashboard(); };
    const selectNone = document.createElement("button");
    selectNone.textContent = "Tout décocher";
    selectNone.onclick = () => { hiddenLists = allLists.map(l => l.id); saveHiddenLists(); renderFilterPanel(); renderDashboard(); };
    actions.appendChild(selectAll);
    actions.appendChild(selectNone);
    panel.appendChild(actions);

    allLists.forEach(list => {
        const item = document.createElement("label");
        item.className = "filter-item";
        const cb = document.createElement("input");
        cb.type = "checkbox";
        cb.id = "filter-" + list.id;
        cb.checked = !hiddenLists.includes(list.id);
        cb.addEventListener("change", () => {
            if (cb.checked) {
                hiddenLists = hiddenLists.filter(id => id !== list.id);
            } else {
                hiddenLists.push(list.id);
            }
            saveHiddenLists();
            renderDashboard();
        });
        const visual = document.createElement("span");
        visual.className = "cb";
        item.appendChild(cb);
        item.appendChild(visual);
        item.appendChild(document.createTextNode(list.displayName));
        panel.appendChild(item);
    });
}

// ===== TAG PANEL =====
function toggleTagPanel() {
    const panel = document.getElementById("tag-panel");
    panel.classList.toggle("open");
    document.getElementById("filter-panel").classList.remove("open");
    document.getElementById("tag-filter-panel").classList.remove("open");
    if (panel.classList.contains("open")) renderTagPanel();
}

function renderTagPanel() {
    const panel = document.getElementById("tag-panel");
    panel.innerHTML = '<h3>Couleurs des tags</h3>';

    const tags = getAllKnownTags();
    if (tags.length === 0) {
        const empty = document.createElement("div");
        empty.className = "tag-panel-empty";
        empty.textContent = "Aucun tag utilisé. Ajoutez #tag dans le titre d'une tâche.";
        panel.appendChild(empty);
        return;
    }

    tags.forEach(tag => {
        const item = document.createElement("div");
        item.className = "tag-panel-item";

        const colorInput = document.createElement("input");
        colorInput.type = "color";
        colorInput.value = getProjectColor(tag);

        const hexInput = document.createElement("input");
        hexInput.type = "text";
        hexInput.className = "tag-hex-input";
        hexInput.value = getProjectColor(tag).toUpperCase();
        hexInput.maxLength = 7;
        hexInput.spellcheck = false;

        colorInput.addEventListener("input", () => {
            setProjectColor(tag, colorInput.value);
            hexInput.value = colorInput.value.toUpperCase();
            renderDashboard();
        });

        hexInput.addEventListener("input", () => {
            let v = hexInput.value.trim();
            if (!v.startsWith("#")) v = "#" + v;
            if (/^#[0-9A-Fa-f]{6}$/.test(v)) {
                colorInput.value = v;
                setProjectColor(tag, v);
                hexInput.classList.remove("invalid");
                renderDashboard();
            } else {
                hexInput.classList.add("invalid");
            }
        });

        hexInput.addEventListener("blur", () => {
            hexInput.value = getProjectColor(tag).toUpperCase();
            hexInput.classList.remove("invalid");
        });

        const nameInput = document.createElement("input");
        nameInput.type = "text";
        nameInput.className = "tag-name-input";
        nameInput.value = tag;
        nameInput.spellcheck = false;

        let renaming = false;
        nameInput.addEventListener("keydown", async (e) => {
            if (e.key === "Enter") {
                e.preventDefault();
                nameInput.blur();
            }
            if (e.key === "Escape") {
                nameInput.value = tag;
                nameInput.blur();
            }
        });
        nameInput.addEventListener("blur", async () => {
            const newVal = nameInput.value.trim().toLowerCase().replace(/[^#\w]/g, "");
            if (isRenamingTag || renaming || !newVal || newVal === tag) {
                nameInput.value = tag;
                return;
            }
            renaming = true;
            nameInput.disabled = true;
            nameInput.classList.add("renaming");
            try {
                await renameTag(tag, newVal);
            } catch (err) {
                nameInput.value = tag;
                nameInput.classList.remove("renaming");
                nameInput.disabled = false;
            }
            renaming = false;
        });

        const toggleBtn = document.createElement("button");
        toggleBtn.className = "tag-text-toggle";
        const textColor = getTagTextColor(tag);
        toggleBtn.textContent = "A";
        toggleBtn.style.color = textColor;
        toggleBtn.style.background = getProjectColor(tag);
        toggleBtn.title = textColor === "#fff" ? "Texte blanc (cliquer pour noir)" : "Texte noir (cliquer pour blanc)";
        toggleBtn.addEventListener("click", () => {
            const newColor = getTagTextColor(tag) === "#fff" ? "#000" : "#fff";
            setTagTextColor(tag, newColor);
            toggleBtn.style.color = newColor;
            toggleBtn.title = newColor === "#fff" ? "Texte blanc (cliquer pour noir)" : "Texte noir (cliquer pour blanc)";
            renderDashboard();
        });

        item.appendChild(colorInput);
        item.appendChild(hexInput);
        item.appendChild(nameInput);
        item.appendChild(toggleBtn);
        panel.appendChild(item);
    });
}

// ===== CLICK OUTSIDE TO CLOSE PANELS =====
document.addEventListener("mousedown", (e) => {
    const filterPanel = document.getElementById("filter-panel");
    const tagPanel = document.getElementById("tag-panel");
    const headerActions = document.querySelector(".header-actions");

    if (filterPanel.classList.contains("open") &&
        !filterPanel.contains(e.target) &&
        (!headerActions || !headerActions.contains(e.target))) {
        filterPanel.classList.remove("open");
        loadAndRenderTasks();
    }
    if (tagPanel.classList.contains("open") &&
        !tagPanel.contains(e.target) &&
        (!headerActions || !headerActions.contains(e.target))) {
        tagPanel.classList.remove("open");
    }
    const tagFilterPanel = document.getElementById("tag-filter-panel");
    if (tagFilterPanel.classList.contains("open") &&
        !tagFilterPanel.contains(e.target) &&
        (!headerActions || !headerActions.contains(e.target))) {
        tagFilterPanel.classList.remove("open");
        renderDashboard();
    }
});

// ===== COMPLETED TASKS =====

function openCompletedModal(item) {
    const { task, listId, listName } = item;
    const existing = document.querySelector(".modal-overlay");
    if (existing) existing.remove();

    const overlay = document.createElement("div");
    overlay.className = "modal-overlay";
    const modal = document.createElement("div");
    modal.className = "modal";
    modal.setAttribute("role", "dialog");
    modal.setAttribute("aria-modal", "true");
    modal.setAttribute("aria-label", "Modifier tâche complétée");

    const dueVal = task.dueDateTime ? formatDateForInput(new Date(task.dueDateTime.dateTime + "Z")) : "";
    modal.innerHTML = '<h2>Tâche complétée</h2>'
        + '<label>Titre</label>'
        + '<input type="text" id="modal-c-title" value="' + escapeHtml(task.title) + '">'
        + '<label>Date d\'échéance</label>'
        + '<input type="date" id="modal-c-date" value="' + dueVal + '">'
        + '<label>Importance</label>'
        + '<select id="modal-c-importance">'
        + '<option value="low"' + (task.importance === "low" ? " selected" : "") + '>Basse</option>'
        + '<option value="normal"' + (task.importance === "normal" ? " selected" : "") + '>Normale</option>'
        + '<option value="high"' + (task.importance === "high" ? " selected" : "") + '>Haute</option>'
        + '</select>'
        + '<label>Liste</label>'
        + '<select id="modal-c-list">' + allLists.map(l => '<option value="' + l.id + '"' + (l.id === listId ? " selected" : "") + '>' + escapeHtml(l.displayName) + '</option>').join("") + '</select>'
        + '<div class="modal-actions">'
        + '<button class="modal-btn danger" id="modal-c-delete">Supprimer</button>'
        + '<div style="display:flex;gap:8px;">'
        + '<button class="modal-btn secondary" id="modal-c-close">Fermer</button>'
        + '<button class="modal-btn primary" id="modal-c-save">Enregistrer</button>'
        + '<button class="modal-btn primary" id="modal-c-reopen">Réouvrir</button>'
        + '</div></div>';

    overlay.appendChild(modal);
    document.body.appendChild(overlay);
    const cleanupFocus = trapFocus(overlay, modal);
    function closeModal() { cleanupFocus(); overlay.remove(); }
    overlay.addEventListener("click", (e) => { if (e.target === overlay) closeModal(); });
    document.getElementById("modal-c-close").addEventListener("click", closeModal);

    function getModalValues() {
        const newTitle = document.getElementById("modal-c-title").value.trim();
        const newDate = document.getElementById("modal-c-date").value;
        const newImportance = document.getElementById("modal-c-importance").value;
        const newListId = document.getElementById("modal-c-list").value;
        const dueDateTime = newDate
            ? { dateTime: formatDateForGraph(new Date(newDate)), timeZone: "UTC" }
            : null;
        return { newTitle, newDate, newImportance, newListId, dueDateTime };
    }

    async function moveTaskToList(vals, status) {
        const taskData = { title: vals.newTitle, status, importance: vals.newImportance };
        if (vals.newDate) taskData.dueDateTime = vals.dueDateTime;
        if (task.body?.content) taskData.body = { content: task.body.content, contentType: "text" };
        await graphPost(GRAPH_BASE + "/me/todo/lists/" + vals.newListId + "/tasks", taskData);
        await deleteTask(listId, task.id);
    }

    // Save changes (keep completed, optionally move list)
    document.getElementById("modal-c-save").addEventListener("click", async () => {
        const vals = getModalValues();
        setModalLoading(modal, true);
        try {
            if (vals.newListId !== listId) {
                await moveTaskToList(vals, "completed");
            } else {
                await updateTask(listId, task.id, {
                    title: vals.newTitle, importance: vals.newImportance,
                    dueDateTime: vals.dueDateTime,
                });
            }
            closeModal();
            await loadAndRenderTasks();
            showToast("Tâche mise à jour !");
        } catch (err) {
            setModalLoading(modal, false);
            console.error("Save completed task failed:", err);
            showToast("Impossible de mettre à jour.", true);
        }
    });

    // Reopen task (optionally in a different list)
    document.getElementById("modal-c-reopen").addEventListener("click", async () => {
        const vals = getModalValues();
        setModalLoading(modal, true);
        try {
            if (vals.newListId !== listId) {
                await moveTaskToList(vals, "notStarted");
            } else {
                await updateTask(listId, task.id, {
                    status: "notStarted", title: vals.newTitle,
                    importance: vals.newImportance, dueDateTime: vals.dueDateTime,
                });
            }
            closeModal();
            await loadAndRenderTasks();
            showToast("Tâche réouverte !");
        } catch (err) {
            setModalLoading(modal, false);
            console.error("Reopen failed:", err);
            showToast("Impossible de réouvrir.", true);
        }
    });

    document.getElementById("modal-c-delete").addEventListener("click", () => {
        closeModal();
        const taskIndex = allTasks.findIndex(t => t.task.id === task.id && t.listId === listId);
        let removedItem = null;
        if (taskIndex !== -1) removedItem = allTasks.splice(taskIndex, 1)[0];
        renderDashboard();

        let cancelled = false;
        const deleteTimeout = setTimeout(async () => {
            if (cancelled) return;
            try {
                await deleteTask(listId, task.id);
            } catch (err) {
                console.error("Delete failed:", err);
                showToast("Impossible de supprimer.", true);
                if (removedItem) allTasks.push(removedItem);
                renderDashboard();
            }
        }, 5000);

        showToast("Tâche supprimée.", false, () => {
            cancelled = true;
            clearTimeout(deleteTimeout);
            if (removedItem) allTasks.push(removedItem);
            renderDashboard();
        });
    });
}

// ===== RENDERING =====
let allTasks = [];
let allLists = [];

async function loadAndRenderTasks() {
    try {
        allLists = await fetchLists();
        console.log("Loaded " + allLists.length + " lists:", allLists.map(l => l.displayName));
        allTasks = [];
        const failedLists = [];

        for (const list of allLists) {
            try {
                const tasks = await fetchTasks(list.id);
                const color = getListColor(list.displayName);
                for (const task of tasks) {
                    allTasks.push({ task, listId: list.id, listName: list.displayName, color });
                }
            } catch (err) {
                console.warn("Impossible de charger la liste '" + list.displayName + "':", err);
                failedLists.push(list.displayName);
            }
        }

        // Fetch completed tasks and merge them into allTasks
        for (const list of allLists) {
            try {
                const tasks = await fetchCompletedTasks(list.id);
                const color = getListColor(list.displayName);
                for (const task of tasks) {
                    if (task.completedDateTime) {
                        const completedDate = new Date(task.completedDateTime.dateTime + "Z");
                        allTasks.push({ task, listId: list.id, listName: list.displayName, color, completed: true, completedDate });
                    }
                }
            } catch (err) {
                console.warn("Impossible de charger les complétées de '" + list.displayName + "':", err);
            }
        }

        renderDashboard();
        updateLastRefresh();
        if (failedLists.length > 0) {
            showToast("Listes inaccessibles : " + failedLists.join(", "), true);
        }
    } catch (err) {
        console.error("Failed to load tasks:", err);
        showToast("Impossible de charger les tâches.", true);
    }
}

function renderDashboard() {
    const weekGrid = document.getElementById("week-grid");
    weekGrid.innerHTML = "";
    const hiddenSet = new Set(hiddenLists);
    const hiddenTagSet = new Set(hiddenTags);
    const visibleTasks = allTasks.filter(({ task, listId }) => {
        if (hiddenSet.has(listId) && !task.dueDateTime) return false;
        if (hiddenTagSet.size > 0) {
            const taskTags = extractProjectTags(task.title).map(t => t.toLowerCase());
            if (taskTags.length === 0) {
                if (hiddenTagSet.has("__no_tag__")) return false;
            } else {
                if (taskTags.every(t => hiddenTagSet.has(t))) return false;
            }
        }
        if (searchQuery && !task.title.toLowerCase().includes(searchQuery)) return false;
        return true;
    });
    const today = new Date(); today.setHours(0,0,0,0);

    if (viewMode === "month") {
        weekGrid.className = "week-grid month-view";
        renderMonthGrid(weekGrid, visibleTasks, today);
    } else {
        weekGrid.className = "week-grid";
        renderWeekGrid(weekGrid, visibleTasks, today);
    }

    // Inbox
    const inboxContainer = document.getElementById("inbox-tasks");
    inboxContainer.innerHTML = "";
    const inboxCol = document.getElementById("inbox-column");
    inboxCol.dataset.date = "inbox";
    setupDropZone(inboxCol);

    const inboxTasks = visibleTasks.filter(({ task, completed }) => !task.dueDateTime && !completed);
    if (inboxTasks.length === 0) {
        const empty = document.createElement("div");
        empty.className = "empty-msg";
        empty.textContent = "Rien en attente";
        inboxContainer.appendChild(empty);
    } else {
        getOrderedTasks(inboxTasks, "inbox").forEach(item => inboxContainer.appendChild(createTaskCard(item)));
    }

    updateNavLabel();
}

function renderWeekGrid(container, visibleTasks, today) {
    const weekDays = getWeekDays();
    weekDays.forEach((day) => {
        const col = document.createElement("div");
        col.className = "day-column" + (isSameDay(day, today) ? " today" : "");
        col.dataset.date = formatDateForInput(day);
        col.setAttribute("role", "region");
        col.setAttribute("aria-label", DAY_NAMES[day.getDay()] + " " + day.getDate());
        setupDropZone(col);

        const header = document.createElement("div");
        header.className = "column-header";
        header.textContent = DAY_NAMES[day.getDay()];
        const dateNum = document.createElement("div");
        dateNum.className = "column-date";
        dateNum.textContent = day.getDate();
        col.appendChild(header);
        col.appendChild(dateNum);

        const activeDayTasks = visibleTasks.filter(({ task, completed }) => {
            if (completed) return false;
            if (!task.dueDateTime) return false;
            const due = new Date(task.dueDateTime.dateTime + "Z");
            return isSameDay(due, day);
        });
        const completedDayTasks = visibleTasks.filter(({ completed, completedDate }) => {
            if (!completed) return false;
            return isSameDay(completedDate, day);
        });

        if (activeDayTasks.length === 0 && completedDayTasks.length === 0) {
            const empty = document.createElement("div");
            empty.className = "empty-msg";
            empty.textContent = "Rien de prévu";
            col.appendChild(empty);
        } else {
            getOrderedTasks(activeDayTasks, formatDateForInput(day)).forEach(item => col.appendChild(createTaskCard(item)));
            completedDayTasks.forEach(item => col.appendChild(createTaskCard(item)));
        }
        // Click on empty area to create task on that day
        col.addEventListener("click", (e) => {
            if (e.target === col || e.target === header || e.target === dateNum || e.target.classList.contains("empty-msg")) {
                openCreateModal(formatDateForInput(day));
            }
        });

        container.appendChild(col);
    });
}

function renderMonthGrid(container, visibleTasks, today) {
    const refDate = getViewMonth();
    const currentMonth = refDate.getMonth();
    const days = getMonthDays(refDate);

    // Day-of-week headers
    ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"].forEach(name => {
        const hdr = document.createElement("div");
        hdr.className = "month-day-header";
        hdr.textContent = name;
        container.appendChild(hdr);
    });

    days.forEach((day) => {
        const cell = document.createElement("div");
        const isCurrentMonth = day.getMonth() === currentMonth;
        cell.className = "month-cell" + (isSameDay(day, today) ? " today" : "") + (!isCurrentMonth ? " other-month" : "");
        cell.dataset.date = formatDateForInput(day);
        setupDropZone(cell);

        const dateLabel = document.createElement("div");
        dateLabel.className = "month-cell-date";
        dateLabel.textContent = day.getDate();
        cell.appendChild(dateLabel);

        const activeDayTasks = visibleTasks.filter(({ task, completed }) => {
            if (completed) return false;
            if (!task.dueDateTime) return false;
            const due = new Date(task.dueDateTime.dateTime + "Z");
            return isSameDay(due, day);
        });
        const completedDayTasks = visibleTasks.filter(({ completed, completedDate }) => {
            if (!completed) return false;
            return isSameDay(completedDate, day);
        });

        const orderedActiveTasks = getOrderedTasks(activeDayTasks, formatDateForInput(day));
        const allDayTasks = [...orderedActiveTasks, ...completedDayTasks];
        const maxPills = 3;
        allDayTasks.slice(0, maxPills).forEach(item => {
            const pill = document.createElement("div");
            pill.className = "month-task-pill" + (item.completed ? " completed" : "");
            const tags = extractProjectTags(item.task.title);
            if (tags.length > 0) {
                const projColor = getProjectColor(tags[0]);
                pill.style.background = projColor;
                pill.style.color = getTagTextColor(tags[0]);
            } else {
                pill.style.background = item.color.gradient;
                pill.style.color = item.color.text;
            }
            if (!item.completed && item.task.dueDateTime) {
                const due = new Date(item.task.dueDateTime.dateTime + "Z");
                if (isOverdue(due)) pill.classList.add("overdue");
            }
            pill.textContent = item.task.title.replace(/#\w+/g, "").trim();
            if (item.completed) {
                pill.draggable = false;
                pill.addEventListener("click", (e) => { e.stopPropagation(); openCompletedModal(item); });
            } else {
                pill.draggable = true;
                pill.addEventListener("click", (e) => { e.stopPropagation(); openEditModal(item); });
                pill.addEventListener("dragstart", (e) => {
                    pill.classList.add("dragging");
                    const sourceDate = pill.closest("[data-date]")?.dataset.date || "";
                    e.dataTransfer.setData("text/plain", JSON.stringify({ taskId: item.task.id, listId: item.listId, sourceDate: sourceDate }));
                    e.dataTransfer.effectAllowed = "move";
                });
                pill.addEventListener("dragend", () => pill.classList.remove("dragging"));
            }
            cell.appendChild(pill);
        });
        if (allDayTasks.length > maxPills) {
            const more = document.createElement("div");
            more.className = "month-more-indicator";
            more.textContent = "+" + (allDayTasks.length - maxPills) + " autre" + (allDayTasks.length - maxPills > 1 ? "s" : "");
            cell.appendChild(more);
        }

        // Click on empty area to create task on that day
        cell.addEventListener("click", (e) => {
            if (e.target === cell || e.target === dateLabel) {
                openCreateModal(formatDateForInput(day));
            }
        });

        container.appendChild(cell);
    });
}

function createTaskCard(item) {
    const { task, listId, listName, color, completed } = item;
    const card = document.createElement("div");
    card.className = "task-card" + (completed ? " completed" : "");
    card.draggable = !completed;
    card.tabIndex = 0;
    card.dataset.taskId = task.id;
    card.dataset.listId = listId;

    const tags = extractProjectTags(task.title);
    if (tags.length > 0) {
        const projColor = getProjectColor(tags[0]);
        card.style.background = "linear-gradient(135deg, " + projColor + ", " + projColor + "cc)";
        card.style.color = getTagTextColor(tags[0]);
    } else {
        card.style.background = color.gradient;
        card.style.color = color.text;
    }

    let overdueFlag = false;
    if (!completed && task.dueDateTime) {
        const due = new Date(task.dueDateTime.dateTime + "Z");
        if (isOverdue(due)) { card.classList.add("overdue"); overdueFlag = true; }
    }

    const label = document.createElement("div");
    label.className = "task-list-label";
    label.textContent = listName;

    const titleEl = document.createElement("div");
    titleEl.className = "task-title";
    titleEl.textContent = task.title.replace(/#\w+/g, "").trim();

    const importance = document.createElement("div");
    importance.className = "task-importance";
    importance.textContent = task.importance === "high" ? "\u26A1 Haute priorité" : task.importance === "low" ? "\u25CB Basse" : "\u25CF Normale";

    card.appendChild(label);
    card.appendChild(titleEl);
    if (!completed) card.appendChild(importance);

    tags.forEach(tag => {
        const tagEl = document.createElement("span");
        tagEl.className = "project-tag";
        tagEl.style.background = getProjectColor(tag) + "44";
        tagEl.style.color = "inherit";
        tagEl.textContent = tag;
        card.appendChild(tagEl);
    });

    if (overdueFlag) {
        const badge = document.createElement("div");
        badge.className = "overdue-badge";
        badge.textContent = "En retard";
        card.appendChild(badge);
    }

    if (completed) {
        card.addEventListener("click", () => openCompletedModal(item));
    } else {
        let clickTimer = null;
        card.addEventListener("click", (e) => {
            if (card.classList.contains("completing")) return;
            clearTimeout(clickTimer);
            clickTimer = setTimeout(() => openEditModal(item), 250);
        });
        card.addEventListener("dblclick", (e) => {
            e.preventDefault();
            clearTimeout(clickTimer);
            handleComplete(card, listId, task.id);
        });
    }
    if (!completed) {
        card.addEventListener("dragstart", (e) => {
            card.classList.add("dragging");
            const sourceDate = card.closest("[data-date]")?.dataset.date || "";
            e.dataTransfer.setData("text/plain", JSON.stringify({ taskId: task.id, listId: listId, sourceDate: sourceDate }));
            e.dataTransfer.effectAllowed = "move";
        });
        card.addEventListener("dragend", () => card.classList.remove("dragging"));
    }

    // Keyboard navigation for accessibility
    card.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") {
            e.preventDefault();
            if (completed) openCompletedModal(item);
            else openEditModal(item);
            return;
        }

        const column = card.closest("[data-date]");
        if (!column) return;
        const dateKey = column.dataset.date;

        // Alt+Up/Down = reorder within column
        if (e.altKey && (e.key === "ArrowUp" || e.key === "ArrowDown")) {
            e.preventDefault();
            const cards = [...column.querySelectorAll(".task-card")];
            const currentIndex = cards.indexOf(card);
            const newIndex = e.key === "ArrowUp" ? currentIndex - 1 : currentIndex + 1;
            if (newIndex < 0 || newIndex >= cards.length) return;
            const currentOrder = cards.map(c => c.dataset.taskId);
            const [moved] = currentOrder.splice(currentIndex, 1);
            currentOrder.splice(newIndex, 0, moved);
            taskOrder[dateKey] = currentOrder;
            saveTaskOrder();
            renderDashboard();
            requestAnimationFrame(() => {
                const newCard = document.querySelector('.task-card[data-task-id="' + task.id + '"]');
                if (newCard) newCard.focus();
            });
            return;
        }

        // Alt+Left/Right = move to adjacent column (week view only)
        if (e.altKey && (e.key === "ArrowLeft" || e.key === "ArrowRight")) {
            e.preventDefault();
            const columns = [...document.querySelectorAll("[data-date]")];
            const colIndex = columns.indexOf(column);
            const targetIndex = e.key === "ArrowLeft" ? colIndex - 1 : colIndex + 1;
            if (targetIndex < 0 || targetIndex >= columns.length) return;
            const targetDate = columns[targetIndex].dataset.date;
            let updates = {};
            if (targetDate === "inbox") {
                updates.dueDateTime = null;
            } else {
                updates.dueDateTime = { dateTime: formatDateForGraph(new Date(targetDate)), timeZone: "UTC" };
            }
            item.task.dueDateTime = updates.dueDateTime;
            if (taskOrder[dateKey]) taskOrder[dateKey] = taskOrder[dateKey].filter(id => id !== task.id);
            if (!taskOrder[targetDate]) taskOrder[targetDate] = [];
            taskOrder[targetDate].push(task.id);
            saveTaskOrder();
            renderDashboard();
            updateTask(listId, task.id, updates).catch(err => {
                console.error("Move failed:", err);
                showToast("Impossible de déplacer la tâche.", true);
                loadAndRenderTasks();
            });
            requestAnimationFrame(() => {
                const newCard = document.querySelector('.task-card[data-task-id="' + task.id + '"]');
                if (newCard) newCard.focus();
            });
            return;
        }
    });

    return card;
}

// ===== DRAG & DROP =====
function removeDropIndicator(container) {
    const existing = container.querySelector(".drop-indicator");
    if (existing) existing.remove();
}

function showDropIndicator(container, y) {
    const cards = [...container.querySelectorAll(".task-card:not(.dragging)")];
    if (cards.length === 0) { removeDropIndicator(container); return; }
    let indicator = container.querySelector(".drop-indicator");
    if (!indicator) {
        indicator = document.createElement("div");
        indicator.className = "drop-indicator";
    }
    const idx = getDropIndex(container, y);
    if (idx < cards.length) {
        cards[idx].before(indicator);
    } else {
        cards[cards.length - 1].after(indicator);
    }
}

function setupDropZone(el) {
    el.addEventListener("dragover", (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = "move";
        el.classList.add("drag-over");
        showDropIndicator(el, e.clientY);
    });
    el.addEventListener("dragleave", (e) => {
        if (!el.contains(e.relatedTarget)) {
            el.classList.remove("drag-over");
            removeDropIndicator(el);
        }
    });
    el.addEventListener("drop", async (e) => {
        e.preventDefault();
        el.classList.remove("drag-over");
        removeDropIndicator(el);
        try {
            const data = JSON.parse(e.dataTransfer.getData("text/plain"));
            const targetDate = el.dataset.date;
            const sourceDate = data.sourceDate || "";
            const dropIdx = getDropIndex(el, e.clientY);

            if (sourceDate === targetDate) {
                // Same column: reorder only (no API call)
                const dateKey = targetDate === "inbox" ? "inbox" : targetDate;
                // Build current order from visible cards
                const currentCards = [...el.querySelectorAll(".task-card:not(.dragging)")];
                const currentOrder = currentCards.map(c => c.dataset.taskId);
                // Insert dragged task at new position
                currentOrder.splice(dropIdx, 0, data.taskId);
                taskOrder[dateKey] = currentOrder;
                saveTaskOrder();
                renderDashboard();
            } else {
                // Different column: change date via API + update order
                let updates = {};
                if (targetDate === "inbox") {
                    updates.dueDateTime = null;
                } else {
                    updates.dueDateTime = { dateTime: formatDateForGraph(new Date(targetDate)), timeZone: "UTC" };
                }
                await updateTask(data.listId, data.taskId, updates);
                const item = allTasks.find(t => t.task.id === data.taskId && t.listId === data.listId);
                if (item) {
                    item.task.dueDateTime = updates.dueDateTime;
                }
                // Remove from source order
                const sourceKey = sourceDate === "inbox" ? "inbox" : sourceDate;
                if (taskOrder[sourceKey]) {
                    taskOrder[sourceKey] = taskOrder[sourceKey].filter(id => id !== data.taskId);
                }
                // Add to target order at drop position
                const targetKey = targetDate === "inbox" ? "inbox" : targetDate;
                if (!taskOrder[targetKey]) taskOrder[targetKey] = [];
                taskOrder[targetKey].splice(dropIdx, 0, data.taskId);
                saveTaskOrder();
                renderDashboard();
            }
        } catch (err) {
            console.error("Drop failed:", err);
            showToast("Impossible de déplacer la tâche.", true);
        }
    });
}

// ===== TAG CHIPS HELPER =====
function buildTagChips(container, selectedTags, onChangeCallback) {
    container.innerHTML = '';
    const knownTags = getAllKnownTags();
    const selected = new Set(selectedTags.map(t => t.toLowerCase()));

    const chipsDiv = document.createElement("div");
    chipsDiv.className = "tag-chips";

    knownTags.forEach(tag => {
        const chip = document.createElement("button");
        chip.type = "button";
        chip.className = "tag-chip" + (selected.has(tag) ? " active" : "");
        chip.style.background = getProjectColor(tag) + "44";
        chip.style.color = getProjectColor(tag);
        chip.textContent = tag;
        chip.addEventListener("click", () => {
            if (selected.has(tag)) {
                selected.delete(tag);
            } else {
                selected.add(tag);
            }
            onChangeCallback([...selected]);
            buildTagChips(container, [...selected], onChangeCallback);
        });
        chipsDiv.appendChild(chip);
    });
    container.appendChild(chipsDiv);

    const addRow = document.createElement("div");
    addRow.className = "tag-add-row";
    const input = document.createElement("input");
    input.type = "text";
    input.placeholder = "Nouveau #tag";
    const addBtn = document.createElement("button");
    addBtn.type = "button";
    addBtn.textContent = "+ Ajouter";
    addBtn.addEventListener("click", () => {
        let val = input.value.trim().toLowerCase();
        if (!val) return;
        if (!val.startsWith("#")) val = "#" + val;
        val = val.replace(/[^#\w]/g, "");
        if (val.length <= 1) return;
        selected.add(val);
        getProjectColor(val);
        onChangeCallback([...selected]);
        buildTagChips(container, [...selected], onChangeCallback);
    });
    input.addEventListener("keydown", (e) => {
        if (e.key === "Enter") { e.preventDefault(); addBtn.click(); }
    });
    addRow.appendChild(input);
    addRow.appendChild(addBtn);
    container.appendChild(addRow);
}

// ===== EDIT MODAL =====
function openEditModal(item) {
    const { task, listId, listName } = item;
    const existing = document.querySelector(".modal-overlay");
    if (existing) existing.remove();

    const currentTags = extractProjectTags(task.title);
    let selectedTags = currentTags.map(t => t.toLowerCase());
    const cleanTitle = task.title.replace(/#\w+/g, "").trim();

    const overlay = document.createElement("div");
    overlay.className = "modal-overlay";
    const modal = document.createElement("div");
    modal.className = "modal";
    modal.setAttribute("role", "dialog");
    modal.setAttribute("aria-modal", "true");
    modal.setAttribute("aria-label", "Modifier la tâche");
    modal.innerHTML = '<h2>Modifier la tâche</h2>'
        + '<label>Titre</label>'
        + '<input type="text" id="modal-title" value="' + escapeHtml(cleanTitle) + '">'
        + '<div class="tag-section"><label>Tags</label><div id="modal-tag-chips"></div></div>'
        + '<label>Description</label>'
        + '<textarea id="modal-body">' + escapeHtml(task.body?.content || "") + '</textarea>'
        + '<label>Date d\'échéance</label>'
        + '<input type="date" id="modal-date" value="' + (task.dueDateTime ? formatDateForInput(new Date(task.dueDateTime.dateTime + "Z")) : "") + '">'
        + '<label>Importance</label>'
        + '<select id="modal-importance">'
        + '<option value="low"' + (task.importance === "low" ? " selected" : "") + '>Basse</option>'
        + '<option value="normal"' + (task.importance === "normal" ? " selected" : "") + '>Normale</option>'
        + '<option value="high"' + (task.importance === "high" ? " selected" : "") + '>Haute</option>'
        + '</select>'
        + '<label>Liste</label>'
        + '<select id="modal-list">' + allLists.map(l => '<option value="' + l.id + '"' + (l.id === listId ? " selected" : "") + '>' + escapeHtml(l.displayName) + '</option>').join("") + '</select>'
        + '<div class="modal-actions">'
        + '<button class="modal-btn danger" id="modal-delete">Supprimer</button>'
        + '<div style="display:flex;gap:8px;">'
        + '<button class="modal-btn secondary" id="modal-cancel">Annuler</button>'
        + '<button class="modal-btn primary" id="modal-save">Enregistrer</button>'
        + '</div></div>';

    overlay.appendChild(modal);
    document.body.appendChild(overlay);
    const cleanupFocus = trapFocus(overlay, modal);
    function closeModal() { cleanupFocus(); overlay.remove(); }
    overlay.addEventListener("click", (e) => { if (e.target === overlay) closeModal(); });
    document.getElementById("modal-cancel").addEventListener("click", closeModal);

    buildTagChips(document.getElementById("modal-tag-chips"), selectedTags, (tags) => { selectedTags = tags; });

    document.getElementById("modal-save").addEventListener("click", async () => {
        let newTitle = document.getElementById("modal-title").value.trim();
        if (selectedTags.length > 0) newTitle += " " + selectedTags.join(" ");
        const newBody = document.getElementById("modal-body").value.trim();
        const newDate = document.getElementById("modal-date").value;
        const newImportance = document.getElementById("modal-importance").value;

        const updates = {
            title: newTitle,
            importance: newImportance,
            body: { content: newBody, contentType: "text" },
        };
        if (newDate) {
            updates.dueDateTime = { dateTime: formatDateForGraph(new Date(newDate)), timeZone: "UTC" };
        } else {
            updates.dueDateTime = null;
        }

        setModalLoading(modal, true);
        try {
            await updateTask(listId, task.id, updates);
            closeModal();
            await loadAndRenderTasks();
            showToast("Tâche mise à jour !");
        } catch (err) {
            setModalLoading(modal, false);
            console.error("Update failed:", err);
            showToast("Erreur lors de la mise à jour.", true);
        }
    });

    document.getElementById("modal-delete").addEventListener("click", () => {
        closeModal();
        const taskIndex = allTasks.findIndex(t => t.task.id === task.id && t.listId === listId);
        let removedItem = null;
        if (taskIndex !== -1) removedItem = allTasks.splice(taskIndex, 1)[0];
        renderDashboard();

        let cancelled = false;
        const deleteTimeout = setTimeout(async () => {
            if (cancelled) return;
            try {
                await deleteTask(listId, task.id);
            } catch (err) {
                console.error("Delete failed:", err);
                showToast("Impossible de supprimer la tâche.", true);
                if (removedItem) allTasks.push(removedItem);
                renderDashboard();
            }
        }, 5000);

        showToast("Tâche supprimée.", false, () => {
            cancelled = true;
            clearTimeout(deleteTimeout);
            if (removedItem) allTasks.push(removedItem);
            renderDashboard();
        });
    });
}

// ===== CREATE MODAL =====
function openCreateModal(defaultDate) {
    const existing = document.querySelector(".modal-overlay");
    if (existing) existing.remove();

    let selectedTags = [];

    const overlay = document.createElement("div");
    overlay.className = "modal-overlay";

    const modal = document.createElement("div");
    modal.className = "modal";
    modal.setAttribute("role", "dialog");
    modal.setAttribute("aria-modal", "true");
    modal.setAttribute("aria-label", "Créer une tâche");
    modal.innerHTML = '<h2>Nouvelle tâche</h2>'
        + '<label>Titre</label>'
        + '<input type="text" id="modal-create-title" placeholder="Titre de la tâche">'
        + '<div class="tag-section"><label>Tags</label><div id="modal-create-tag-chips"></div></div>'
        + '<label>Description</label>'
        + '<textarea id="modal-create-body" placeholder="Description (optionnel)"></textarea>'
        + '<label>Date d\'échéance</label>'
        + '<input type="date" id="modal-create-date" value="' + (defaultDate || '') + '">'
        + '<label>Importance</label>'
        + '<select id="modal-create-importance">'
        + '<option value="normal" selected>Normale</option>'
        + '<option value="low">Basse</option>'
        + '<option value="high">Haute</option>'
        + '</select>'
        + '<label>Liste</label>'
        + '<select id="modal-create-list">' + allLists.map(l => '<option value="' + l.id + '">' + escapeHtml(l.displayName) + '</option>').join("") + '</select>'
        + '<div class="modal-actions">'
        + '<div></div>'
        + '<div style="display:flex;gap:8px;">'
        + '<button class="modal-btn secondary" id="modal-create-cancel">Annuler</button>'
        + '<button class="modal-btn primary" id="modal-create-save">Créer</button>'
        + '</div></div>';

    overlay.appendChild(modal);
    document.body.appendChild(overlay);
    const cleanupFocus = trapFocus(overlay, modal);
    function closeModal() { cleanupFocus(); overlay.remove(); }
    overlay.addEventListener("click", (e) => { if (e.target === overlay) closeModal(); });
    document.getElementById("modal-create-cancel").addEventListener("click", closeModal);

    buildTagChips(document.getElementById("modal-create-tag-chips"), selectedTags, (tags) => { selectedTags = tags; });

    async function doCreate() {
        let title = document.getElementById("modal-create-title").value.trim();
        if (!title) { showToast("Le titre est obligatoire.", true); return; }
        if (selectedTags.length > 0) title += " " + selectedTags.join(" ");
        const body = document.getElementById("modal-create-body").value.trim();
        const date = document.getElementById("modal-create-date").value;
        const importance = document.getElementById("modal-create-importance").value;
        const listId = document.getElementById("modal-create-list").value;

        const taskData = { title, importance };
        if (body) taskData.body = { content: body, contentType: "text" };
        if (date) taskData.dueDateTime = { dateTime: formatDateForGraph(new Date(date)), timeZone: "UTC" };

        setModalLoading(modal, true);
        try {
            await graphPost(GRAPH_BASE + "/me/todo/lists/" + listId + "/tasks", taskData);
            closeModal();
            await loadAndRenderTasks();
            showToast("Tâche créée !");
        } catch (err) {
            setModalLoading(modal, false);
            console.error("Create failed:", err);
            showToast("Impossible de créer la tâche.", true);
        }
    }

    document.getElementById("modal-create-save").addEventListener("click", doCreate);
    document.getElementById("modal-create-title").addEventListener("keydown", (e) => {
        if (e.key === "Enter") { e.preventDefault(); doCreate(); }
    });
}

// ===== TASK COMPLETION =====
async function handleComplete(card, listId, taskId) {
    if (card.classList.contains("completing")) return;
    card.classList.add("completing");

    let cancelled = false;

    showToast("Tâche complétée !", false, () => {
        cancelled = true;
        card.classList.remove("completing");
        uncompleteTask(listId, taskId).catch(() => {});
    });

    try {
        await completeTask(listId, taskId);
        setTimeout(() => {
            if (cancelled) return;
            card.classList.add("fade-out");
            setTimeout(() => {
                card.remove();
                loadAndRenderTasks();
            }, 400);
        }, 3000);
    } catch (err) {
        card.classList.remove("completing");
        showToast("Impossible de compléter la tâche.", true);
    }
}

// ===== AUTO-REFRESH =====
let refreshIntervalId = null;
let refreshCountdown = REFRESH_INTERVAL * 60;

function startAutoRefresh() {
    if (refreshIntervalId) clearInterval(refreshIntervalId);
    refreshCountdown = REFRESH_INTERVAL * 60;
    refreshIntervalId = setInterval(() => {
        refreshCountdown--;
        updateRefreshTimer();
        if (refreshCountdown <= 0) { refreshTasks(); refreshCountdown = REFRESH_INTERVAL * 60; }
    }, 1000);
}
async function refreshTasks() {
    try { await loadAndRenderTasks(); }
    catch (err) { showToast("Impossible de rafraîchir.", true); }
    refreshCountdown = REFRESH_INTERVAL * 60;
}
function updateRefreshTimer() {
    const min = Math.floor(refreshCountdown / 60);
    const sec = refreshCountdown % 60;
    const el = document.getElementById("refresh-timer");
    if (el) el.textContent = "Auto-refresh dans " + min + ":" + (sec < 10 ? "0" : "") + sec;
}
function updateLastRefresh() {
    const now = new Date();
    const el = document.getElementById("last-update");
    if (el) el.textContent = "Dernière MàJ : " + now.getHours().toString().padStart(2,"0") + ":" + now.getMinutes().toString().padStart(2,"0");
}

// ===== SEARCH LISTENER =====
document.getElementById("search-input").addEventListener("input", (e) => {
    searchQuery = e.target.value.trim().toLowerCase();
    renderDashboard();
});

// ===== KEYBOARD SHORTCUTS =====
document.addEventListener("keydown", (e) => {
    const tag = e.target.tagName;
    const isInput = tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT";

    // Ctrl+K / Cmd+K = focus search (always active)
    if ((e.ctrlKey || e.metaKey) && e.key === "k") {
        e.preventDefault();
        document.getElementById("search-input").focus();
        return;
    }

    if (isInput) return;

    // N = new task
    if (e.key === "n" || e.key === "N") {
        e.preventDefault();
        openCreateModal();
        return;
    }
    // R = refresh
    if (e.key === "r" || e.key === "R") {
        e.preventDefault();
        refreshTasks();
        return;
    }
    // T = today
    if (e.key === "t" || e.key === "T") {
        e.preventDefault();
        navigateToday();
        return;
    }
    // Escape = close panels
    if (e.key === "Escape") {
        document.getElementById("filter-panel").classList.remove("open");
        document.getElementById("tag-panel").classList.remove("open");
        document.getElementById("tag-filter-panel").classList.remove("open");
        return;
    }
    // Alt+Left = navigate prev
    if (e.altKey && e.key === "ArrowLeft") {
        e.preventDefault();
        navigatePrev();
        return;
    }
    // Alt+Right = navigate next
    if (e.altKey && e.key === "ArrowRight") {
        e.preventDefault();
        navigateNext();
        return;
    }
});

// ===== STARTUP =====
async function init() {
    if (typeof msal === "undefined") { setTimeout(init, 200); return; }
    loadCustomTagColors();
    loadCustomTagTextColors();
    await loadColors();
    try {
        msalInstance = new msal.PublicClientApplication(msalConfig);
        await msalInstance.initialize();
        await msalInstance.handleRedirectPromise();
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) await onAuthenticated();
    } catch (err) {
        console.error("MSAL init failed:", err);
        showToast("Erreur d'initialisation. Rechargez la page.", true);
    }
}
init();
