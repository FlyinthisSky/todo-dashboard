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
function logout() { if (msalInstance) msalInstance.logoutRedirect({ postLogoutRedirectUri: REDIRECT_URI }); }
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

// ===== GRAPH API =====
async function graphGet(url) {
    const token = await getAccessToken();
    const r = await fetch(url, { headers: { Authorization: "Bearer " + token } });
    if (!r.ok) throw new Error("Graph API error: " + r.status);
    return r.json();
}
async function graphPatch(url, body) {
    const token = await getAccessToken();
    const r = await fetch(url, { method: "PATCH", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" }, body: JSON.stringify(body) });
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
    const token = await getAccessToken();
    const r = await fetch(url, { method: "DELETE", headers: { Authorization: "Bearer " + token } });
    if (!r.ok) throw new Error("Graph DELETE error: " + r.status);
}
async function graphPost(url, body) {
    const token = await getAccessToken();
    const r = await fetch(url, { method: "POST", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" }, body: JSON.stringify(body) });
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
    oldTag = oldTag.toLowerCase();
    newTag = newTag.toLowerCase();
    if (!newTag.startsWith("#")) newTag = "#" + newTag;
    newTag = newTag.replace(/[^#\w]/g, "");
    if (newTag === "#" || newTag === oldTag) return;

    // Transfer color from old tag to new tag
    const color = getProjectColor(oldTag);
    setProjectColor(newTag, color);
    delete projectColorMap[oldTag];
    delete customTagColors[oldTag];
    saveCustomTagColors();

    // Update all tasks containing the old tag (active + completed)
    const allKnownTasks = [...allTasks, ...completedTasks];
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
    renderDashboard();
    renderTagPanel();
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
    const monday = getMonday(new Date());
    monday.setDate(monday.getDate() + weekOffset * 7);
    return Array.from({ length: 7 }, (_, i) => { const d = new Date(monday); d.setDate(monday.getDate() + i); return d; });
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

function toggleFilter() {
    const panel = document.getElementById("filter-panel");
    const wasOpen = panel.classList.contains("open");
    panel.classList.toggle("open");
    document.getElementById("tag-panel").classList.remove("open");
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
            if (renaming || !newVal || newVal === tag) {
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

        item.appendChild(colorInput);
        item.appendChild(hexInput);
        item.appendChild(nameInput);
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
});

// ===== COMPLETED TASKS =====
let completedTasks = [];

function toggleCompleted() {
    const sidebar = document.getElementById("completed-sidebar");
    sidebar.classList.toggle("open");
    if (sidebar.classList.contains("open")) loadCompletedTasks();
}

async function loadCompletedTasks() {
    const container = document.getElementById("completed-tasks");
    container.innerHTML = '<div style="color:var(--text-muted);font-size:0.8rem;">Chargement...</div>';
    completedTasks = [];
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

    // Fetch all lists in parallel instead of sequentially
    const results = await Promise.allSettled(
        allLists.map(async (list) => {
            const tasks = await fetchCompletedTasks(list.id);
            return { tasks, listId: list.id, listName: list.displayName };
        })
    );

    for (const result of results) {
        if (result.status !== "fulfilled") continue;
        const { tasks, listId, listName } = result.value;
        for (const task of tasks) {
            if (task.completedDateTime) {
                const completedDate = new Date(task.completedDateTime.dateTime + "Z");
                if (completedDate >= thirtyDaysAgo) {
                    completedTasks.push({ task, listId, listName, completedDate });
                }
            }
        }
    }

    // Sort using cached date objects
    completedTasks.sort((a, b) => b.completedDate - a.completedDate);

    const hiddenSet = new Set(hiddenLists);
    const visibleCompleted = completedTasks.filter(({ listId }) => !hiddenSet.has(listId));
    container.innerHTML = "";
    if (visibleCompleted.length === 0) {
        container.innerHTML = '<div style="color:var(--text-muted);font-size:0.8rem;font-style:italic;">Aucune tâche complétée</div>';
        return;
    }

    // Batch DOM operations with DocumentFragment + event delegation
    const fragment = document.createDocumentFragment();
    visibleCompleted.forEach(({ task, listName, completedDate }, index) => {
        const card = document.createElement("div");
        card.className = "completed-card";
        card.style.cursor = "pointer";
        card.dataset.index = index;
        const titleSpan = document.createElement("div");
        titleSpan.className = "task-title";
        titleSpan.textContent = task.title;
        const meta = document.createElement("div");
        meta.className = "completed-meta";
        meta.innerHTML = '<span>' + escapeHtml(listName) + '</span>'
            + '<span>' + completedDate.getDate() + '/' + (completedDate.getMonth()+1) + '</span>';
        card.appendChild(titleSpan);
        card.appendChild(meta);
        fragment.appendChild(card);
    });
    container.appendChild(fragment);

    // Single delegated click listener instead of one per card
    container.onclick = (e) => {
        const card = e.target.closest(".completed-card");
        if (!card) return;
        const item = visibleCompleted[parseInt(card.dataset.index)];
        if (item) openCompletedModal(item);
    };
}

function openCompletedModal(item) {
    const { task, listId, listName } = item;
    const existing = document.querySelector(".modal-overlay");
    if (existing) existing.remove();

    const overlay = document.createElement("div");
    overlay.className = "modal-overlay";
    overlay.addEventListener("click", (e) => { if (e.target === overlay) overlay.remove(); });

    const modal = document.createElement("div");
    modal.className = "modal";

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
        + '<button class="modal-btn secondary" onclick="this.closest(\'.modal-overlay\').remove()">Fermer</button>'
        + '<button class="modal-btn primary" id="modal-c-save">Enregistrer</button>'
        + '<button class="modal-btn primary" id="modal-c-reopen">Réouvrir</button>'
        + '</div></div>';

    overlay.appendChild(modal);
    document.body.appendChild(overlay);

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
        try {
            if (vals.newListId !== listId) {
                await moveTaskToList(vals, "completed");
            } else {
                await updateTask(listId, task.id, {
                    title: vals.newTitle, importance: vals.newImportance,
                    dueDateTime: vals.dueDateTime,
                });
            }
            overlay.remove();
            loadCompletedTasks();
            showToast("Tâche mise à jour !");
        } catch (err) {
            console.error("Save completed task failed:", err);
            showToast("Impossible de mettre à jour.", true);
        }
    });

    // Reopen task (optionally in a different list)
    document.getElementById("modal-c-reopen").addEventListener("click", async () => {
        const vals = getModalValues();
        try {
            if (vals.newListId !== listId) {
                await moveTaskToList(vals, "notStarted");
            } else {
                await updateTask(listId, task.id, {
                    status: "notStarted", title: vals.newTitle,
                    importance: vals.newImportance, dueDateTime: vals.dueDateTime,
                });
            }
            overlay.remove();
            await loadAndRenderTasks();
            loadCompletedTasks();
            showToast("Tâche réouverte !");
        } catch (err) {
            console.error("Reopen failed:", err);
            showToast("Impossible de réouvrir.", true);
        }
    });

    document.getElementById("modal-c-delete").addEventListener("click", async () => {
        if (!confirm("Supprimer cette tâche définitivement ?")) return;
        try {
            await deleteTask(listId, task.id);
            overlay.remove();
            loadCompletedTasks();
            showToast("Tâche supprimée.");
        } catch (err) {
            console.error("Delete failed:", err);
            showToast("Impossible de supprimer.", true);
        }
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

        // Fetch all lists in parallel instead of sequentially
        const results = await Promise.allSettled(
            allLists.map(async (list) => {
                const tasks = await fetchTasks(list.id);
                return { tasks, listId: list.id, listName: list.displayName };
            })
        );

        results.forEach((result, i) => {
            if (result.status === "fulfilled") {
                const { tasks, listId, listName } = result.value;
                const color = getListColor(listName);
                for (const task of tasks) {
                    allTasks.push({ task, listId, listName, color });
                }
            } else {
                console.warn("Impossible de charger la liste '" + allLists[i].displayName + "':", result.reason);
                failedLists.push(allLists[i].displayName);
            }
        });

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
    const visibleTasks = allTasks.filter(({ listId }) => !hiddenSet.has(listId));
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

    const inboxTasks = visibleTasks.filter(({ task }) => !task.dueDateTime);
    if (inboxTasks.length === 0) {
        const empty = document.createElement("div");
        empty.className = "empty-msg";
        empty.textContent = "Rien en attente";
        inboxContainer.appendChild(empty);
    } else {
        inboxTasks.forEach(item => inboxContainer.appendChild(createTaskCard(item)));
    }

    updateNavLabel();
}

function renderWeekGrid(container, visibleTasks, today) {
    const weekDays = getWeekDays();
    weekDays.forEach((day) => {
        const col = document.createElement("div");
        col.className = "day-column" + (isSameDay(day, today) ? " today" : "");
        col.dataset.date = formatDateForInput(day);
        setupDropZone(col);

        const header = document.createElement("div");
        header.className = "column-header";
        header.textContent = DAY_NAMES[day.getDay()];
        const dateNum = document.createElement("div");
        dateNum.className = "column-date";
        dateNum.textContent = day.getDate();
        col.appendChild(header);
        col.appendChild(dateNum);

        const dayTasks = visibleTasks.filter(({ task }) => {
            if (!task.dueDateTime) return false;
            const due = new Date(task.dueDateTime.dateTime + "Z");
            return isSameDay(due, day);
        });

        if (dayTasks.length === 0) {
            const empty = document.createElement("div");
            empty.className = "empty-msg";
            empty.textContent = "Rien de prévu";
            col.appendChild(empty);
        } else {
            dayTasks.forEach(item => col.appendChild(createTaskCard(item)));
        }
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

        const dayTasks = visibleTasks.filter(({ task }) => {
            if (!task.dueDateTime) return false;
            const due = new Date(task.dueDateTime.dateTime + "Z");
            return isSameDay(due, day);
        });

        dayTasks.forEach(item => {
            const pill = document.createElement("div");
            pill.className = "month-task-pill";
            const tags = extractProjectTags(item.task.title);
            if (tags.length > 0) {
                pill.style.background = getProjectColor(tags[0]);
                pill.style.color = "#fff";
            } else {
                pill.style.background = item.color.gradient;
                pill.style.color = item.color.text;
            }
            if (item.task.dueDateTime) {
                const due = new Date(item.task.dueDateTime.dateTime + "Z");
                if (isOverdue(due)) pill.classList.add("overdue");
            }
            pill.textContent = item.task.title.replace(/#\w+/g, "").trim();
            pill.draggable = true;
            pill.addEventListener("click", (e) => { e.stopPropagation(); openEditModal(item); });
            pill.addEventListener("dragstart", (e) => {
                pill.classList.add("dragging");
                e.dataTransfer.setData("text/plain", JSON.stringify({ taskId: item.task.id, listId: item.listId }));
                e.dataTransfer.effectAllowed = "move";
            });
            pill.addEventListener("dragend", () => pill.classList.remove("dragging"));
            cell.appendChild(pill);
        });

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
    const { task, listId, listName, color } = item;
    const card = document.createElement("div");
    card.className = "task-card";
    card.draggable = true;
    card.dataset.taskId = task.id;
    card.dataset.listId = listId;

    const tags = extractProjectTags(task.title);
    if (tags.length > 0) {
        const projColor = getProjectColor(tags[0]);
        card.style.background = "linear-gradient(135deg, " + projColor + ", " + projColor + "cc)";
        card.style.color = "#fff";
    } else {
        card.style.background = color.gradient;
        card.style.color = color.text;
    }

    let overdueFlag = false;
    if (task.dueDateTime) {
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
    card.appendChild(importance);

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
    card.addEventListener("dragstart", (e) => {
        card.classList.add("dragging");
        e.dataTransfer.setData("text/plain", JSON.stringify({ taskId: task.id, listId: listId }));
        e.dataTransfer.effectAllowed = "move";
    });
    card.addEventListener("dragend", () => card.classList.remove("dragging"));

    return card;
}

// ===== DRAG & DROP =====
function setupDropZone(el) {
    el.addEventListener("dragover", (e) => { e.preventDefault(); e.dataTransfer.dropEffect = "move"; el.classList.add("drag-over"); });
    el.addEventListener("dragleave", (e) => {
        if (!el.contains(e.relatedTarget)) el.classList.remove("drag-over");
    });
    el.addEventListener("drop", async (e) => {
        e.preventDefault();
        el.classList.remove("drag-over");
        try {
            const data = JSON.parse(e.dataTransfer.getData("text/plain"));
            const targetDate = el.dataset.date;
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
            renderDashboard();
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
    overlay.addEventListener("click", (e) => { if (e.target === overlay) overlay.remove(); });

    const modal = document.createElement("div");
    modal.className = "modal";
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
        + '<button class="modal-btn secondary" onclick="this.closest(\'.modal-overlay\').remove()">Annuler</button>'
        + '<button class="modal-btn primary" id="modal-save">Enregistrer</button>'
        + '</div></div>';

    overlay.appendChild(modal);
    document.body.appendChild(overlay);

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

        try {
            await updateTask(listId, task.id, updates);
            overlay.remove();
            await loadAndRenderTasks();
            showToast("Tâche mise à jour !");
        } catch (err) {
            console.error("Update failed:", err);
            showToast("Erreur lors de la mise à jour.", true);
        }
    });

    document.getElementById("modal-delete").addEventListener("click", async () => {
        if (!confirm("Supprimer cette tâche définitivement ?")) return;
        try {
            await deleteTask(listId, task.id);
            overlay.remove();
            await loadAndRenderTasks();
            showToast("Tâche supprimée.");
        } catch (err) {
            console.error("Delete failed:", err);
            showToast("Impossible de supprimer la tâche.", true);
        }
    });
}

// ===== CREATE MODAL =====
function openCreateModal(defaultDate) {
    const existing = document.querySelector(".modal-overlay");
    if (existing) existing.remove();

    let selectedTags = [];

    const overlay = document.createElement("div");
    overlay.className = "modal-overlay";
    overlay.addEventListener("click", (e) => { if (e.target === overlay) overlay.remove(); });

    const modal = document.createElement("div");
    modal.className = "modal";
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
        + '<button class="modal-btn secondary" onclick="this.closest(\'.modal-overlay\').remove()">Annuler</button>'
        + '<button class="modal-btn primary" id="modal-create-save">Créer</button>'
        + '</div></div>';

    overlay.appendChild(modal);
    document.body.appendChild(overlay);

    buildTagChips(document.getElementById("modal-create-tag-chips"), selectedTags, (tags) => { selectedTags = tags; });
    setTimeout(() => document.getElementById("modal-create-title").focus(), 50);

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

        try {
            await graphPost(GRAPH_BASE + "/me/todo/lists/" + listId + "/tasks", taskData);
            overlay.remove();
            await loadAndRenderTasks();
            showToast("Tâche créée !");
        } catch (err) {
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
            setTimeout(() => card.remove(), 400);
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

// ===== STARTUP =====
async function init() {
    if (typeof msal === "undefined") { setTimeout(init, 200); return; }
    loadCustomTagColors();
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
