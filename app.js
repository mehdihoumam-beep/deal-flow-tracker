/* Deal Flow Tracker - Edition Cloud (OneDrive)
   Toutes les donnees sont lues/ecrites dans un fichier JSON stocke dans le
   dossier applicatif ("Apps/Deal Flow Tracker") du OneDrive personnel de
   l'utilisateur connecte, via Microsoft Graph. Rien n'est stocke sur un
   serveur : l'authentification (MSAL) donne un jeton d'acces directement
   utilise depuis le navigateur. */

let state = { config: { statuses: [], departements: [], secteurs: [], responsables: [], natures: [] }, tasks: [] };
let currentView = "kanban";
let filters = { search: "", departement: [], secteur: [], responsable: [], urgence: [], consultation: [] };
let collapsedGroups = {};
let expandedSubtasks = new Set();

const FILTER_DEFS = [
  { key: "departement", label: "Departement", getOptions: () => state.config.departements },
  { key: "secteur", label: "Secteur", getOptions: () => state.config.secteurs },
  { key: "responsable", label: "Responsable", getOptions: () => state.config.responsables },
  { key: "urgence", label: "Urgence", getOptions: () => ["OUI", "NON"] },
  { key: "consultation", label: "Consultation", getOptions: () => ["Oui", "Non"] },
];

const SUBTASK_STATUSES = ["A faire", "En cours", "Termine"];
const PRIORITIES = ["urgent", "high", "normal", "low"];

/* ---------- Etat par defaut (premiere utilisation, avant tout fichier OneDrive) ---------- */

function defaultState() {
  return {
    config: {
      statuses: [
        { key: "standby", label: "Standby", color: "#e93d82" },
        { key: "to_do", label: "To Do", color: "#87909e" },
        { key: "etude", label: "En cours d'etude", color: "#5f55ee" },
        { key: "attente_client", label: "Attente retour client", color: "#f8ae00" },
        { key: "risq", label: "En cours RISQ", color: "#656f7d" },
        { key: "mep_cours", label: "En cours de MEP", color: "#b660e0" },
        { key: "mep", label: "MEP", color: "#d33d44" },
        { key: "abandonne", label: "Abandonne", color: "#0f9d9f" },
        { key: "perdu", label: "Perdu", color: "#9d570f" },
        { key: "complete", label: "Complete", color: "#008844" },
      ],
      departements: ["FS", "FC", "AUTRES"],
      secteurs: ["Sante", "Hotellerie", "Transport", "Fonciere", "Aeronautique", "Automobile", "BTP", "Logistique", "Parking", "Energie", "Grande distribution", "Agriculture", "Holding", "Mines", "Industrie", "OPCI", "Education", "Services", "Telecommunication", "Port", "Recyclage", "Dessalement", "Restauration", "Minoterie", "Retail", "IT"],
      responsables: ["Khalil", "Inass", "Hiba", "Zaineb A.", "Hamza", "Khadija", "Mohamed", "Abdelhadi", "Zainab S.", "Wissal", "Mehdi"],
      natures: ["CMLT", "DP RAC"],
    },
    tasks: [],
  };
}

/* ---------- Authentification (MSAL) ---------- */

let msalInstance = null;
let currentAccount = null;
const GRAPH_SCOPES = ["Files.ReadWrite.AppFolder", "User.Read"];

async function initMsal() {
  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId: window.DEAL_FLOW_CONFIG.CLIENT_ID,
      authority: window.DEAL_FLOW_CONFIG.AUTHORITY,
      redirectUri: window.location.href.split("#")[0].split("?")[0],
    },
    cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false },
  });
  await msalInstance.initialize();
  await msalInstance.handleRedirectPromise();
}

function hasExistingSession() {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    currentAccount = accounts[0];
    return true;
  }
  return false;
}

async function login() {
  const loginStatus = document.getElementById("login-status");
  try {
    loginStatus.textContent = "Connexion en cours...";
    const result = await msalInstance.loginPopup({ scopes: GRAPH_SCOPES });
    currentAccount = result.account;
    return true;
  } catch (e) {
    console.error(e);
    loginStatus.textContent = "Echec de connexion : " + (e.errorMessage || e.message || String(e));
    return false;
  }
}

function logout() {
  msalInstance.logoutRedirect({ account: currentAccount });
}

async function getAccessToken() {
  const request = { scopes: GRAPH_SCOPES, account: currentAccount };
  try {
    const result = await msalInstance.acquireTokenSilent(request);
    return result.accessToken;
  } catch (e) {
    const result = await msalInstance.acquireTokenPopup(request);
    return result.accessToken;
  }
}

/* ---------- Microsoft Graph (OneDrive - dossier applicatif) ---------- */

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

function graphFileUrl() {
  const filename = encodeURIComponent(window.DEAL_FLOW_CONFIG.STATE_FILENAME);
  return `${GRAPH_BASE}/me/drive/special/approot:/${filename}:/content`;
}

async function graphFetchState() {
  const token = await getAccessToken();
  const res = await fetch(graphFileUrl(), { headers: { Authorization: `Bearer ${token}` } });
  if (res.status === 404) {
    const fresh = defaultState();
    await graphSaveStateRaw(fresh, token);
    return fresh;
  }
  if (!res.ok) throw new Error(`Lecture OneDrive impossible (${res.status})`);
  return await res.json();
}

async function graphSaveStateRaw(data, token) {
  const res = await fetch(graphFileUrl(), {
    method: "PUT",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(data),
  });
  if (!res.ok) throw new Error(`Ecriture OneDrive impossible (${res.status})`);
}

/* ---------- API (compatible avec le reste de l'app) ---------- */

async function fetchState() {
  state = await graphFetchState();
}

let saveTimer = null;
function persist() {
  const indicator = document.getElementById("save-indicator");
  indicator.textContent = "Synchronisation...";
  indicator.classList.add("saving");
  clearTimeout(saveTimer);
  saveTimer = setTimeout(async () => {
    try {
      const token = await getAccessToken();
      await graphSaveStateRaw(state, token);
      indicator.textContent = "✓ Synchronise avec OneDrive";
      indicator.classList.remove("saving");
    } catch (e) {
      console.error(e);
      indicator.textContent = "⚠ Erreur de synchronisation";
    }
  }, 400);
}

/* ---------- Helpers ---------- */

function uid(prefix) {
  return prefix + "_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}
function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str == null ? "" : String(str);
  return div.innerHTML;
}
function priorityLabel(p) { return p ? p.charAt(0).toUpperCase() + p.slice(1) : ""; }
function fmtDate(ms) {
  if (!ms) return "";
  return new Date(ms).toLocaleDateString("fr-FR", { day: "2-digit", month: "short", year: "2-digit" });
}
function dateInputVal(ms) { return ms ? new Date(ms).toISOString().slice(0, 10) : ""; }
function parseDateInput(str) { return str ? new Date(str + "T12:00:00").getTime() : null; }
function isOverdue(task) {
  if (!task.echeance) return false;
  if (["complete", "abandonne", "perdu"].includes(task.status)) return false;
  return task.echeance < Date.now();
}
function durationDays(start, end) {
  if (!start || !end) return null;
  return Math.round((end - start) / 86400000);
}
function statusInfo(key) {
  return state.config.statuses.find((s) => s.key === key) || { key, label: key, color: "#999" };
}
function subtaskProgress(task) {
  const subs = task.subtasks || [];
  if (!subs.length) return null;
  const done = subs.filter((s) => s.statut === "Termine").length;
  return { done, total: subs.length };
}

/* ---------- Filtering ---------- */

function matchesFilters(task) {
  if (filters.search && !task.name.toLowerCase().includes(filters.search.toLowerCase())) return false;
  if (filters.departement.length && !filters.departement.includes(task.departement)) return false;
  if (filters.secteur.length && !filters.secteur.includes(task.secteur)) return false;
  if (filters.urgence.length && !filters.urgence.includes(task.urgence)) return false;
  if (filters.consultation.length && !filters.consultation.includes(task.consultation)) return false;
  if (filters.responsable.length) {
    const resp = task.responsables || [];
    if (!resp.some((r) => filters.responsable.includes(r))) return false;
  }
  return true;
}

/* ---------- Header: filters bar ---------- */

function renderFiltersBar() {
  const bar = document.getElementById("filters-bar");
  bar.innerHTML = "";
  FILTER_DEFS.forEach((def) => {
    const options = def.getOptions();
    const selected = filters[def.key];
    const wrap = document.createElement("div");
    wrap.className = "multiselect";
    const btnLabel = selected.length ? `${def.label} (${selected.length})` : def.label;
    wrap.innerHTML = `
      <button class="multiselect-btn ${selected.length ? "active-filter" : ""}" type="button">${btnLabel}</button>
      <div class="multiselect-panel">
        ${options.map((opt) => `
          <label><input type="checkbox" value="${escapeHtml(opt)}" ${selected.includes(opt) ? "checked" : ""}> ${escapeHtml(opt)}</label>
        `).join("")}
      </div>
    `;
    const btn = wrap.querySelector(".multiselect-btn");
    const panel = wrap.querySelector(".multiselect-panel");
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      document.querySelectorAll(".multiselect-panel.open").forEach((p) => { if (p !== panel) p.classList.remove("open"); });
      panel.classList.toggle("open");
    });
    panel.querySelectorAll("input[type=checkbox]").forEach((cb) => {
      cb.addEventListener("change", () => {
        const val = cb.value;
        if (cb.checked) filters[def.key].push(val);
        else filters[def.key] = filters[def.key].filter((v) => v !== val);
        renderAll();
      });
    });
    bar.appendChild(wrap);
  });
}
document.addEventListener("click", () => {
  document.querySelectorAll(".multiselect-panel.open").forEach((p) => p.classList.remove("open"));
});

function renderStats() {
  const visible = state.tasks.filter(matchesFilters);
  document.getElementById("stat-total").textContent = `${visible.length} deal${visible.length !== 1 ? "s" : ""}`;
  const overdue = visible.filter(isOverdue).length;
  document.getElementById("stat-overdue").textContent = `${overdue} en retard`;
}

/* ---------- Kanban view ---------- */

function renderKanban() {
  const board = document.getElementById("board");
  board.innerHTML = "";
  const visible = state.tasks.filter(matchesFilters);

  state.config.statuses.forEach((st) => {
    const col = document.createElement("div");
    col.className = "column";
    const tasksInCol = visible.filter((t) => t.status === st.key);

    col.innerHTML = `
      <div class="column-header">
        <span class="dot" style="background:${st.color}"></span>
        <span class="name">${escapeHtml(st.label)}</span>
        <span class="count">${tasksInCol.length}</span>
      </div>
      <div class="card-list" data-status="${st.key}"></div>
      <div class="column-footer"><button class="add-card-btn">+ Ajouter</button></div>
    `;

    const list = col.querySelector(".card-list");
    tasksInCol.forEach((t) => list.appendChild(renderCard(t)));
    col.querySelector(".add-card-btn").addEventListener("click", () => openTaskModal(null, st.key));

    list.addEventListener("dragover", (e) => { e.preventDefault(); list.classList.add("drag-over"); });
    list.addEventListener("dragleave", () => list.classList.remove("drag-over"));
    list.addEventListener("drop", (e) => {
      e.preventDefault();
      list.classList.remove("drag-over");
      const id = e.dataTransfer.getData("text/plain");
      const task = state.tasks.find((x) => x.id === id);
      if (task) { task.status = st.key; persist(); renderAll(); }
    });

    board.appendChild(col);
  });
}

function renderCard(t) {
  const card = document.createElement("div");
  card.className = "card";
  card.draggable = true;
  const overdue = isOverdue(t);
  const subCount = (t.subtasks || []).length;
  const isExpanded = expandedSubtasks.has(t.id);

  let metaHtml = "";
  if (t.priority) metaHtml += `<span class="badge prio-${t.priority}">${priorityLabel(t.priority)}</span>`;
  if (t.echeance) metaHtml += `<span class="chip ${overdue ? "overdue" : ""}">${overdue ? "⚠ " : "📅 "}${fmtDate(t.echeance)}</span>`;
  if (t.urgence === "OUI") metaHtml += `<span class="chip overdue">Urgent</span>`;

  const avatars = (t.responsables || []).slice(0, 3).map((r) =>
    `<span class="avatar" title="${escapeHtml(r)}">${r.split(" ").map((p) => p[0]).join("").slice(0,2).toUpperCase()}</span>`
  ).join("");

  card.innerHTML = `
    <div class="title">
      <span class="title-text">${escapeHtml(t.name)}</span>
      ${subCount ? `<span class="subtask-toggle ${isExpanded ? "open" : ""}" data-id="${t.id}">(${subCount})</span>` : ""}
    </div>
    <div class="meta">${metaHtml}<span class="avatars">${avatars}</span></div>
    ${isExpanded ? renderInlineSubtasks(t) : ""}
  `;
  card.addEventListener("dragstart", (e) => { e.dataTransfer.setData("text/plain", t.id); setTimeout(() => card.classList.add("dragging"), 0); });
  card.addEventListener("dragend", () => card.classList.remove("dragging"));
  card.addEventListener("click", () => openTaskModal(t.id));

  const toggle = card.querySelector(".subtask-toggle");
  if (toggle) {
    toggle.addEventListener("click", (e) => {
      e.stopPropagation();
      if (expandedSubtasks.has(t.id)) expandedSubtasks.delete(t.id);
      else expandedSubtasks.add(t.id);
      renderAll();
    });
  }
  return card;
}

function renderInlineSubtasks(t) {
  const rows = (t.subtasks || []).map((s) => `
    <div class="inline-subtask-row">
      <span class="inline-subtask-name">${escapeHtml(s.name || "(sans nom)")}</span>
      ${s.responsable ? `<span class="chip">${escapeHtml(s.responsable)}</span>` : ""}
      <span class="chip st-status-${(s.statut || "").replace(/\s+/g, "_")}">${escapeHtml(s.statut || "")}</span>
    </div>
  `).join("");
  return `<div class="inline-subtasks" onclick="event.stopPropagation()">${rows}</div>`;
}

/* ---------- Table view ---------- */

function renderTable() {
  const container = document.getElementById("table-container");
  const visible = state.tasks.filter(matchesFilters);

  let rows = "";
  state.config.statuses.forEach((st) => {
    const tasksInGroup = visible.filter((t) => t.status === st.key);
    const collapsed = !!collapsedGroups[st.key];
    rows += `
      <tr class="group-header" data-status="${st.key}">
        <td colspan="11">
          <span class="chevron ${collapsed ? "collapsed" : ""}">▼</span>
          <span class="dot" style="background:${st.color}"></span>
          ${escapeHtml(st.label)} <span style="color:var(--text-dim); font-weight:400;">(${tasksInGroup.length})</span>
        </td>
      </tr>
    `;
    tasksInGroup.forEach((t) => {
      const dur = durationDays(t.date_debut, t.date_fin);
      const progress = subtaskProgress(t);
      const subCount = (t.subtasks || []).length;
      const isExpanded = expandedSubtasks.has(t.id);
      rows += `
        <tr class="deal-row ${collapsed ? "group-collapsed" : ""}" data-status="${st.key}" data-id="${t.id}">
          <td>${escapeHtml(t.name)}${subCount ? ` <span class="subtask-toggle ${isExpanded ? "open" : ""}" data-id="${t.id}">(${subCount})</span>` : ""}</td>
          <td>${t.priority ? `<span class="badge prio-${t.priority}">${priorityLabel(t.priority)}</span>` : ""}</td>
          <td>${escapeHtml(t.departement || "")}</td>
          <td>${escapeHtml(t.secteur || "")}</td>
          <td>${escapeHtml(t.consultation || "")}</td>
          <td>${escapeHtml(t.urgence || "")}</td>
          <td>${(t.responsables || []).join(", ")}</td>
          <td>${fmtDate(t.date_debut)}</td>
          <td>${isOverdue(t) ? `<span style="color:#d32f2f;font-weight:600;">${fmtDate(t.echeance)}</span>` : fmtDate(t.echeance)}</td>
          <td>${fmtDate(t.date_fin)}${dur !== null ? `<span class="duration-badge">${dur} j</span>` : ""}</td>
          <td>${progress ? `${progress.done}/${progress.total}` : ""}</td>
        </tr>
      `;
      if (isExpanded && subCount) {
        rows += `
          <tr class="subtasks-detail-row ${collapsed ? "group-collapsed" : ""}" data-status="${st.key}">
            <td colspan="11">${renderInlineSubtasks(t)}</td>
          </tr>
        `;
      }
    });
  });

  container.innerHTML = `
    <table class="deals-table">
      <thead>
        <tr>
          <th>Nom</th><th>Priorite</th><th>Departement</th><th>Secteur</th>
          <th>Consultation</th><th>Urgence</th><th>Responsables</th>
          <th>Debut</th><th>Echeance</th><th>Fin (duree)</th><th>Sous-taches</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;

  container.querySelectorAll("tr.group-header").forEach((row) => {
    row.addEventListener("click", () => {
      const st = row.dataset.status;
      collapsedGroups[st] = !collapsedGroups[st];
      renderTable();
    });
  });
  container.querySelectorAll("tr.deal-row").forEach((row) => {
    row.addEventListener("click", (e) => {
      const toggle = e.target.closest(".subtask-toggle");
      if (toggle) {
        e.stopPropagation();
        const id = toggle.dataset.id;
        if (expandedSubtasks.has(id)) expandedSubtasks.delete(id);
        else expandedSubtasks.add(id);
        renderAll();
        return;
      }
      openTaskModal(row.dataset.id);
    });
  });
}

/* ---------- Render orchestration ---------- */

function renderAll() {
  renderFiltersBar();
  renderStats();
  if (currentView === "kanban") renderKanban();
  else if (currentView === "table") renderTable();
  else renderDashboard();
}

document.getElementById("view-toggle").addEventListener("click", (e) => {
  const btn = e.target.closest(".view-btn");
  if (!btn) return;
  currentView = btn.dataset.view;
  document.querySelectorAll(".view-btn").forEach((b) => b.classList.toggle("active", b === btn));
  document.getElementById("kanban-view").classList.toggle("hidden", currentView !== "kanban");
  document.getElementById("table-view").classList.toggle("hidden", currentView !== "table");
  document.getElementById("dashboard-view").classList.toggle("hidden", currentView !== "dashboard");
  renderAll();
});

document.getElementById("search").addEventListener("input", (e) => { filters.search = e.target.value; renderAll(); });
document.getElementById("reset-filters").addEventListener("click", () => {
  filters = { search: "", departement: [], secteur: [], responsable: [], urgence: [], consultation: [] };
  document.getElementById("search").value = "";
  renderAll();
});
document.getElementById("new-deal-btn").addEventListener("click", () => openTaskModal(null));

/* ---------- Task modal ---------- */

const modalOverlay = document.getElementById("modal-overlay");
const modal = document.getElementById("modal");

function blankTask(presetStatus) {
  return {
    id: uid("task"), name: "", status: presetStatus || state.config.statuses[0].key,
    priority: null, departement: "", secteur: "", urgence: "", consultation: "", nature: "",
    responsables: [], date_debut: null, echeance: null, date_fin: null,
    notes: "", subtasks: [],
  };
}

function openTaskModal(taskId, presetStatus) {
  const isNew = !taskId;
  const task = isNew ? blankTask(presetStatus) : JSON.parse(JSON.stringify(state.tasks.find((t) => t.id === taskId)));

  modal.innerHTML = `
    <h2>${isNew ? "Nouveau deal" : "Modifier le deal"}</h2>
    <div class="field">
      <label>Nom du deal</label>
      <input type="text" id="f-name" value="${escapeHtml(task.name)}" placeholder="Ex: SAMTA">
    </div>
    <div class="field-row">
      <div class="field">
        <label>Statut</label>
        <select id="f-status">${state.config.statuses.map((s) => `<option value="${s.key}" ${s.key === task.status ? "selected" : ""}>${escapeHtml(s.label)}</option>`).join("")}</select>
      </div>
      <div class="field">
        <label>Priorite</label>
        <select id="f-priority">
          <option value="">-</option>
          ${PRIORITIES.map((p) => `<option value="${p}" ${p === task.priority ? "selected" : ""}>${priorityLabel(p)}</option>`).join("")}
        </select>
      </div>
    </div>
    <div class="field-row">
      <div class="field">
        <label>Departement</label>
        <select id="f-departement">
          <option value="">-</option>
          ${state.config.departements.map((d) => `<option value="${d}" ${d === task.departement ? "selected" : ""}>${escapeHtml(d)}</option>`).join("")}
        </select>
      </div>
      <div class="field">
        <label>Secteur</label>
        <select id="f-secteur">
          <option value="">-</option>
          ${state.config.secteurs.map((s) => `<option value="${s}" ${s === task.secteur ? "selected" : ""}>${escapeHtml(s)}</option>`).join("")}
        </select>
      </div>
    </div>
    <div class="field-row">
      <div class="field">
        <label>Urgence</label>
        <select id="f-urgence">
          <option value="">-</option>
          <option value="OUI" ${task.urgence === "OUI" ? "selected" : ""}>OUI</option>
          <option value="NON" ${task.urgence === "NON" ? "selected" : ""}>NON</option>
        </select>
      </div>
      <div class="field">
        <label>Consultation</label>
        <select id="f-consultation">
          <option value="">-</option>
          <option value="Oui" ${task.consultation === "Oui" ? "selected" : ""}>Oui</option>
          <option value="Non" ${task.consultation === "Non" ? "selected" : ""}>Non</option>
        </select>
      </div>
    </div>
    <div class="field">
      <label>Nature</label>
      <select id="f-nature">
        <option value="">-</option>
        ${state.config.natures.map((n) => `<option value="${escapeHtml(n)}" ${n === task.nature ? "selected" : ""}>${escapeHtml(n)}</option>`).join("")}
      </select>
    </div>
    <div class="field">
      <label>Responsables (plusieurs choix possibles)</label>
      <div class="chip-select" id="f-responsables">
        ${state.config.responsables.map((r) => `<span class="chip-option ${task.responsables.includes(r) ? "selected" : ""}" data-value="${escapeHtml(r)}">${escapeHtml(r)}</span>`).join("")}
      </div>
    </div>
    <div class="field-row">
      <div class="field">
        <label>Date de debut</label>
        <input type="date" id="f-debut" value="${dateInputVal(task.date_debut)}">
      </div>
      <div class="field">
        <label>Echeance</label>
        <input type="date" id="f-echeance" value="${dateInputVal(task.echeance)}">
      </div>
    </div>
    <div class="field">
      <label>Date de fin</label>
      <div class="duration-inline">
        <input type="date" id="f-fin" value="${dateInputVal(task.date_fin)}">
        <span class="duration-readout" id="duration-readout"></span>
      </div>
    </div>
    <div class="field">
      <label>Notes</label>
      <textarea id="f-notes" placeholder="Notes libres...">${escapeHtml(task.notes)}</textarea>
    </div>

    <div class="section-title">
      <span>Sous-taches</span>
      <button class="ghost" id="add-subtask-btn" type="button" style="padding:4px 9px;font-size:12px;">+ Ajouter</button>
    </div>
    <div class="subtasks-list" id="subtasks-list"></div>

    <div class="modal-actions">
      <div>${!isNew ? `<button class="ghost danger" id="delete-btn" type="button">Supprimer</button>` : ""}</div>
      <div class="right">
        <button class="ghost" id="cancel-btn" type="button">Annuler</button>
        <button class="primary" id="save-btn" type="button">Enregistrer</button>
      </div>
    </div>
  `;

  modalOverlay.classList.add("open");

  // responsables chip toggle
  modal.querySelectorAll("#f-responsables .chip-option").forEach((chip) => {
    chip.addEventListener("click", () => {
      chip.classList.toggle("selected");
      const val = chip.dataset.value;
      if (chip.classList.contains("selected")) {
        if (!task.responsables.includes(val)) task.responsables.push(val);
      } else {
        task.responsables = task.responsables.filter((r) => r !== val);
      }
    });
  });

  // duration readout
  function updateDuration() {
    const start = parseDateInput(document.getElementById("f-debut").value);
    const end = parseDateInput(document.getElementById("f-fin").value);
    const dur = durationDays(start, end);
    document.getElementById("duration-readout").textContent = dur !== null ? `${dur} jour${dur !== 1 ? "s" : ""}` : "";
  }
  document.getElementById("f-debut").addEventListener("input", updateDuration);
  document.getElementById("f-fin").addEventListener("input", updateDuration);
  updateDuration();

  // subtasks
  function renderSubtasks() {
    const list = document.getElementById("subtasks-list");
    if (!task.subtasks.length) {
      list.innerHTML = `<div style="font-size:12px;color:var(--text-dim);">Aucune sous-tache.</div>`;
      return;
    }
    list.innerHTML = task.subtasks.map((s, idx) => `
      <div class="subtask-row" data-idx="${idx}">
        <input type="text" class="st-name" placeholder="Nom de la sous-tache" value="${escapeHtml(s.name)}">
        <select class="st-resp">
          <option value="">Responsable</option>
          ${state.config.responsables.map((r) => `<option value="${r}" ${r === s.responsable ? "selected" : ""}>${escapeHtml(r)}</option>`).join("")}
        </select>
        <input type="date" class="st-due" value="${dateInputVal(s.echeance)}">
        <select class="st-statut">
          ${SUBTASK_STATUSES.map((st2) => `<option value="${st2}" ${st2 === s.statut ? "selected" : ""}>${st2}</option>`).join("")}
        </select>
        <button class="remove-subtask" type="button" title="Supprimer">&times;</button>
      </div>
    `).join("");

    list.querySelectorAll(".subtask-row").forEach((row) => {
      const idx = Number(row.dataset.idx);
      row.querySelector(".st-name").addEventListener("input", (e) => { task.subtasks[idx].name = e.target.value; });
      row.querySelector(".st-resp").addEventListener("change", (e) => { task.subtasks[idx].responsable = e.target.value; });
      row.querySelector(".st-due").addEventListener("change", (e) => { task.subtasks[idx].echeance = parseDateInput(e.target.value); });
      row.querySelector(".st-statut").addEventListener("change", (e) => { task.subtasks[idx].statut = e.target.value; });
      row.querySelector(".remove-subtask").addEventListener("click", () => {
        task.subtasks.splice(idx, 1);
        renderSubtasks();
      });
    });
  }
  renderSubtasks();
  document.getElementById("add-subtask-btn").addEventListener("click", () => {
    task.subtasks.push({ id: uid("sub"), name: "", responsable: "", echeance: null, statut: "A faire" });
    renderSubtasks();
  });

  document.getElementById("cancel-btn").addEventListener("click", closeTaskModal);
  modalOverlay.addEventListener("click", (e) => { if (e.target === modalOverlay) closeTaskModal(); });

  if (!isNew) {
    document.getElementById("delete-btn").addEventListener("click", () => {
      if (confirm("Supprimer ce deal ?")) {
        state.tasks = state.tasks.filter((t) => t.id !== taskId);
        persist();
        closeTaskModal();
        renderAll();
      }
    });
  }

  document.getElementById("save-btn").addEventListener("click", () => {
    const name = document.getElementById("f-name").value.trim();
    if (!name) { alert("Le nom du deal est requis."); return; }

    const updated = {
      ...task,
      name,
      status: document.getElementById("f-status").value,
      priority: document.getElementById("f-priority").value || null,
      departement: document.getElementById("f-departement").value,
      secteur: document.getElementById("f-secteur").value,
      urgence: document.getElementById("f-urgence").value,
      consultation: document.getElementById("f-consultation").value,
      nature: document.getElementById("f-nature").value,
      date_debut: parseDateInput(document.getElementById("f-debut").value),
      echeance: parseDateInput(document.getElementById("f-echeance").value),
      date_fin: parseDateInput(document.getElementById("f-fin").value),
      notes: document.getElementById("f-notes").value,
    };

    if (isNew) state.tasks.push(updated);
    else {
      const idx = state.tasks.findIndex((t) => t.id === taskId);
      state.tasks[idx] = updated;
    }
    persist();
    closeTaskModal();
    renderAll();
  });
}

function closeTaskModal() {
  modalOverlay.classList.remove("open");
  modal.innerHTML = "";
}

/* ---------- Settings modal ---------- */

const settingsOverlay = document.getElementById("settings-overlay");
const settingsModal = document.getElementById("settings-modal");

document.getElementById("settings-btn").addEventListener("click", openSettingsModal);

function openSettingsModal() {
  settingsModal.innerHTML = `
    <h2>&#9881;&#65039; Parametres des listes</h2>
    <div id="settings-body"></div>
    <div class="modal-actions">
      <div></div>
      <div class="right"><button class="primary" id="settings-close-btn" type="button">Fermer</button></div>
    </div>
  `;
  settingsOverlay.classList.add("open");
  renderSettingsBody();
  document.getElementById("settings-close-btn").addEventListener("click", closeSettingsModal);
  settingsOverlay.addEventListener("click", (e) => { if (e.target === settingsOverlay) closeSettingsModal(); });
}

function closeSettingsModal() {
  settingsOverlay.classList.remove("open");
  settingsModal.innerHTML = "";
  renderAll();
}

const SETTINGS_LISTS = [
  { key: "departements", label: "Departements", colored: false },
  { key: "secteurs", label: "Secteurs", colored: false },
  { key: "responsables", label: "Responsables", colored: false },
  { key: "natures", label: "Nature", colored: false },
];

function renderSettingsBody() {
  const body = document.getElementById("settings-body");
  let html = "";

  // Statuses (structured objects)
  html += `
    <div class="settings-section" data-list="statuses">
      <h3>Statuts (colonnes Kanban)</h3>
      <div class="settings-list">
        ${state.config.statuses.map((s, idx) => `
          <div class="settings-item">
            <span><span class="dot" style="background:${s.color};display:inline-block;margin-right:6px;"></span>${escapeHtml(s.label)}</span>
            <button class="remove-item" data-idx="${idx}" type="button">Supprimer</button>
          </div>
        `).join("")}
      </div>
      <div class="settings-add-row">
        <input type="color" id="new-status-color" value="#0071e3">
        <input type="text" id="new-status-label" placeholder="Nouveau statut...">
        <button class="ghost" id="add-status-btn" type="button">Ajouter</button>
      </div>
    </div>
  `;

  SETTINGS_LISTS.forEach((def) => {
    html += `
      <div class="settings-section" data-list="${def.key}">
        <h3>${def.label}</h3>
        <div class="settings-list">
          ${state.config[def.key].map((val, idx) => `
            <div class="settings-item">
              <span>${escapeHtml(val)}</span>
              <button class="remove-item" data-idx="${idx}" type="button">Supprimer</button>
            </div>
          `).join("")}
        </div>
        <div class="settings-add-row">
          <input type="text" class="new-item-input" placeholder="Ajouter un element...">
          <button class="ghost add-item-btn" type="button">Ajouter</button>
        </div>
      </div>
    `;
  });

  body.innerHTML = html;

  // status list handlers
  const statusSection = body.querySelector('[data-list="statuses"]');
  statusSection.querySelectorAll(".remove-item").forEach((btn) => {
    btn.addEventListener("click", () => {
      const idx = Number(btn.dataset.idx);
      const removed = state.config.statuses[idx];
      const inUse = state.tasks.some((t) => t.status === removed.key);
      if (inUse && !confirm(`"${removed.label}" est utilise par des deals existants. Supprimer quand meme ?`)) return;
      state.config.statuses.splice(idx, 1);
      persist();
      renderSettingsBody();
    });
  });
  statusSection.querySelector("#add-status-btn").addEventListener("click", () => {
    const label = document.getElementById("new-status-label").value.trim();
    const color = document.getElementById("new-status-color").value;
    if (!label) return;
    state.config.statuses.push({ key: uid("status"), label, color });
    persist();
    renderSettingsBody();
  });

  // simple list handlers
  SETTINGS_LISTS.forEach((def) => {
    const section = body.querySelector(`[data-list="${def.key}"]`);
    section.querySelectorAll(".remove-item").forEach((btn) => {
      btn.addEventListener("click", () => {
        const idx = Number(btn.dataset.idx);
        state.config[def.key].splice(idx, 1);
        persist();
        renderSettingsBody();
      });
    });
    section.querySelector(".add-item-btn").addEventListener("click", () => {
      const input = section.querySelector(".new-item-input");
      const val = input.value.trim();
      if (!val || state.config[def.key].includes(val)) return;
      state.config[def.key].push(val);
      persist();
      renderSettingsBody();
    });
  });
}

/* ---------- Dashboard ---------- */

let dashboardScope = "all"; // 'all' | 'FS' | 'FC'
let dashboardSelection = null; // { dim, value }

const CHART_COLORS = ["#0071e3", "#5f55ee", "#e93d82", "#f8ae00", "#0f9d9f", "#b660e0", "#d33d44", "#008844", "#87909e", "#ff8a3d", "#30a46c", "#e5484d", "#6647f0", "#12a594"];
function colorFor(i) { return CHART_COLORS[i % CHART_COLORS.length]; }
function scopeLabel() { return dashboardScope === "all" ? "" : ` — ${dashboardScope}`; }

function retardCategory(t) {
  if (!t.echeance) return "Sans echeance";
  if (["complete", "abandonne", "perdu"].includes(t.status)) return "Cloture";
  return t.echeance < Date.now() ? "En retard" : "A temps";
}

function countBy(data, keyFn) {
  const map = {};
  data.forEach((t) => {
    const v = keyFn(t);
    const values = Array.isArray(v) ? (v.length ? v : ["(non renseigne)"]) : [v || "(non renseigne)"];
    values.forEach((val) => { map[val] = (map[val] || 0) + 1; });
  });
  return map;
}

function dashboardDataset(base) {
  if (dashboardScope === "all") return base;
  return base.filter((t) => t.departement === dashboardScope);
}

function taskMatchesSelection(t, sel) {
  if (!sel) return true;
  switch (sel.dim) {
    case "departement": return (t.departement || "(non renseigne)") === sel.value;
    case "statut": return statusInfo(t.status).label === sel.value;
    case "urgence": return (t.urgence || "(non renseigne)") === sel.value;
    case "nature": return (t.nature || "(non renseigne)") === sel.value;
    case "secteur": return (t.secteur || "(non renseigne)") === sel.value;
    case "responsable":
      return (t.responsables || []).length ? t.responsables.includes(sel.value) : sel.value === "(non renseigne)";
    case "retard": return retardCategory(t) === sel.value;
    default: return true;
  }
}

function polarToCartesian(cx, cy, r, angleDeg) {
  const a = (angleDeg - 90) * Math.PI / 180;
  return { x: cx + r * Math.cos(a), y: cy + r * Math.sin(a) };
}
function donutSlicePath(cx, cy, R, r, startAngle, endAngle) {
  if (endAngle - startAngle >= 359.999) endAngle = startAngle + 359.999;
  const largeArc = (endAngle - startAngle) > 180 ? 1 : 0;
  const p1 = polarToCartesian(cx, cy, R, startAngle);
  const p2 = polarToCartesian(cx, cy, R, endAngle);
  const p3 = polarToCartesian(cx, cy, r, endAngle);
  const p4 = polarToCartesian(cx, cy, r, startAngle);
  return `M ${p1.x} ${p1.y} A ${R} ${R} 0 ${largeArc} 1 ${p2.x} ${p2.y} L ${p3.x} ${p3.y} A ${r} ${r} 0 ${largeArc} 0 ${p4.x} ${p4.y} Z`;
}

function buildDonutSVG(counts, dim, size) {
  size = size || 168;
  const entries = Object.entries(counts).filter(([, v]) => v > 0);
  const total = entries.reduce((s, [, v]) => s + v, 0);
  if (!total) return `<div class="chart-empty">Aucune donnee</div>`;
  const R = size / 2, r = R * 0.58, cx = R, cy = R;
  let angle = 0;
  let paths = "";
  entries.forEach(([key, val], i) => {
    const frac = val / total;
    const startAngle = angle;
    const endAngle = angle + frac * 360;
    angle = endAngle;
    const d = donutSlicePath(cx, cy, R, r, startAngle, endAngle);
    const selected = dashboardSelection && dashboardSelection.dim === dim && dashboardSelection.value === key;
    paths += `<path d="${d}" fill="${colorFor(i)}" class="donut-slice ${selected ? "selected" : ""}" data-dim="${dim}" data-value="${escapeHtml(key)}"><title>${escapeHtml(key)}: ${val}</title></path>`;
  });
  const legend = entries.map(([key, val], i) => `
    <div class="legend-item ${dashboardSelection && dashboardSelection.dim === dim && dashboardSelection.value === key ? "selected" : ""}" data-dim="${dim}" data-value="${escapeHtml(key)}">
      <span class="legend-dot" style="background:${colorFor(i)}"></span>
      <span class="legend-label">${escapeHtml(key)}</span>
      <span class="legend-value">${val}</span>
    </div>
  `).join("");
  return `<div class="donut-wrap"><svg viewBox="0 0 ${size} ${size}" width="${size}" height="${size}">${paths}</svg><div class="donut-legend">${legend}</div></div>`;
}

function buildBarChart(counts, dim) {
  const entries = Object.entries(counts).filter(([, v]) => v > 0).sort((a, b) => b[1] - a[1]);
  if (!entries.length) return `<div class="chart-empty">Aucune donnee</div>`;
  const max = Math.max(...entries.map(([, v]) => v));
  return `<div class="bar-chart">${entries.map(([key, val], i) => {
    const selected = dashboardSelection && dashboardSelection.dim === dim && dashboardSelection.value === key;
    const pct = Math.max(4, Math.round((val / max) * 100));
    return `
      <div class="bar-row ${selected ? "selected" : ""}" data-dim="${dim}" data-value="${escapeHtml(key)}">
        <span class="bar-label" title="${escapeHtml(key)}">${escapeHtml(key)}</span>
        <div class="bar-track"><div class="bar-fill" style="width:${pct}%; background:${colorFor(i)}"></div></div>
        <span class="bar-value">${val}</span>
      </div>`;
  }).join("")}</div>`;
}

function renderDashboard() {
  const container = document.getElementById("dashboard-container");
  const base = state.tasks.filter(matchesFilters);
  const data = dashboardDataset(base);

  const total = base.length;
  const fsCount = base.filter((t) => t.departement === "FS").length;
  const fcCount = base.filter((t) => t.departement === "FC").length;

  const byDept = countBy(base, (t) => t.departement);
  const byResp = countBy(data, (t) => t.responsables);
  const byStatus = countBy(data, (t) => statusInfo(t.status).label);
  const byUrgence = countBy(data, (t) => t.urgence);
  const byNature = countBy(data, (t) => t.nature);
  const bySecteur = countBy(data, (t) => t.secteur);
  const byRetard = countBy(data, (t) => retardCategory(t));

  container.innerHTML = `
    <div class="dashboard-toolbar">
      <div class="scope-counters" id="scope-counters">
        <button class="counter-pill ${dashboardScope === "all" ? "active" : ""}" data-scope="all"><strong>${total}</strong><span>Total</span></button>
        <button class="counter-pill ${dashboardScope === "FS" ? "active" : ""}" data-scope="FS"><strong>${fsCount}</strong><span>FS</span></button>
        <button class="counter-pill ${dashboardScope === "FC" ? "active" : ""}" data-scope="FC"><strong>${fcCount}</strong><span>FC</span></button>
      </div>
      <div class="spacer"></div>
      <button class="ghost" id="print-dashboard-btn">&#128424;&#65039; Imprimer</button>
    </div>

    <div class="charts-grid">
      <div class="chart-card"><h3>Repartition FC / FS</h3>${buildDonutSVG(byDept, "departement")}</div>
      <div class="chart-card"><h3>Par statut${scopeLabel()}</h3>${buildDonutSVG(byStatus, "statut")}</div>
      <div class="chart-card"><h3>Par urgence${scopeLabel()}</h3>${buildDonutSVG(byUrgence, "urgence")}</div>
      <div class="chart-card"><h3>Par nature${scopeLabel()}</h3>${buildDonutSVG(byNature, "nature")}</div>
      <div class="chart-card"><h3>Deals en retard${scopeLabel()}</h3>${buildDonutSVG(byRetard, "retard")}</div>
      <div class="chart-card wide"><h3>Par responsable${scopeLabel()}</h3>${buildBarChart(byResp, "responsable")}</div>
      <div class="chart-card wide"><h3>Par secteur${scopeLabel()}</h3>${buildBarChart(bySecteur, "secteur")}</div>
    </div>

    <div id="dashboard-detail"></div>
  `;

  container.querySelectorAll(".counter-pill").forEach((btn) => {
    btn.addEventListener("click", () => {
      dashboardScope = btn.dataset.scope;
      dashboardSelection = null;
      renderDashboard();
    });
  });

  document.getElementById("print-dashboard-btn").addEventListener("click", () => window.print());

  container.querySelectorAll("[data-dim]").forEach((el) => {
    el.addEventListener("click", () => {
      const dim = el.dataset.dim;
      const value = el.dataset.value;
      if (dashboardSelection && dashboardSelection.dim === dim && dashboardSelection.value === value) dashboardSelection = null;
      else dashboardSelection = { dim, value };
      renderDashboard();
    });
  });

  renderDashboardDetail(data, base);
}

const DIM_LABELS = { departement: "Departement", statut: "Statut", urgence: "Urgence", nature: "Nature", secteur: "Secteur", responsable: "Responsable", retard: "Delai" };

function renderDashboardDetail(data, base) {
  const detailContainer = document.getElementById("dashboard-detail");
  if (!detailContainer) return;
  if (!dashboardSelection) { detailContainer.innerHTML = ""; return; }

  const sourceSet = dashboardSelection.dim === "departement" ? base : data;
  const matches = sourceSet.filter((t) => taskMatchesSelection(t, dashboardSelection));

  detailContainer.innerHTML = `
    <div class="detail-header">
      <h3>Detail — ${DIM_LABELS[dashboardSelection.dim]} : ${escapeHtml(dashboardSelection.value)} (${matches.length})</h3>
      <button class="ghost" id="clear-selection-btn">&times; Effacer la selection</button>
    </div>
    <table class="deals-table detail-table">
      <thead><tr><th>Nom</th><th>Statut</th><th>Departement</th><th>Secteur</th><th>Responsables</th><th>Echeance</th></tr></thead>
      <tbody>
        ${matches.map((t) => `
          <tr class="deal-row" data-id="${t.id}">
            <td>${escapeHtml(t.name)}</td>
            <td>${escapeHtml(statusInfo(t.status).label)}</td>
            <td>${escapeHtml(t.departement || "")}</td>
            <td>${escapeHtml(t.secteur || "")}</td>
            <td>${(t.responsables || []).join(", ")}</td>
            <td>${fmtDate(t.echeance)}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
  document.getElementById("clear-selection-btn").addEventListener("click", () => { dashboardSelection = null; renderDashboard(); });
  detailContainer.querySelectorAll("tr.deal-row").forEach((row) => row.addEventListener("click", () => openTaskModal(row.dataset.id)));
}

/* ---------- Boot ---------- */

document.getElementById("logout-btn").addEventListener("click", () => logout());

async function startApp() {
  document.getElementById("login-screen").classList.add("hidden");
  document.getElementById("app-root").classList.remove("hidden");
  try {
    await fetchState();
  } catch (e) {
    console.error(e);
    alert("Impossible de charger les donnees depuis OneDrive : " + e.message);
    return;
  }
  renderAll();
}

(async function init() {
  if (!window.DEAL_FLOW_CONFIG || !window.DEAL_FLOW_CONFIG.CLIENT_ID || window.DEAL_FLOW_CONFIG.CLIENT_ID === "COLLE_TON_CLIENT_ID_ICI") {
    document.getElementById("login-status").textContent =
      "Configuration manquante : ouvre config.js et colle ton Client ID Azure (voir GUIDE-INSTALLATION.md).";
    return;
  }
  await initMsal();
  if (hasExistingSession()) {
    await startApp();
  } else {
    document.getElementById("login-btn").addEventListener("click", async () => {
      const ok = await login();
      if (ok) await startApp();
    });
  }
})();
