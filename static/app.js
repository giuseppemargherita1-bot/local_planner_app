const statusText = document.getElementById("statusText");
const undoBtn = document.getElementById("undoBtn");
const refreshBtn = document.getElementById("refreshBtn");
const importDemandsBtn = document.getElementById("importDemandsBtn");
const restoreBackupBtn = document.getElementById("restoreBackupBtn");
const activityLogBtn = document.getElementById("activityLogBtn");
const activityLogModal = document.getElementById("activityLogModal");
const activityLogClose = document.getElementById("activityLogClose");
const activityLogBody = document.getElementById("activityLogBody");
const importPreviewModal = document.getElementById("importPreviewModal");
const importPreviewText = document.getElementById("importPreviewText");
const importPreviewClose = document.getElementById("importPreviewClose");
const importPreviewCancel = document.getElementById("importPreviewCancel");
const importPreviewConfirm = document.getElementById("importPreviewConfirm");
const importUnknownRoles = document.getElementById("importUnknownRoles");
const importUnknownRolesList = document.getElementById("importUnknownRolesList");
if (importPreviewModal) importPreviewModal.hidden = true;
if (activityLogModal) activityLogModal.hidden = true;
const importPlanningBtn = document.getElementById("importPlanningBtn");
const openDemandSourceBtn = document.getElementById("openDemandSourceBtn");

const plannerHead = document.querySelector("#plannerMatrix thead");
const plannerBody = document.querySelector("#plannerMatrix tbody");
const resourcesTableBody = document.querySelector("#resourcesTable tbody");
const resourcesTable = document.getElementById("resourcesTable");
const resourcesTableHead = document.querySelector("#resourcesTable thead");
const savedReportViews = document.getElementById("savedReportViews");
const reportWeekFrom = document.getElementById("reportWeekFrom");
const reportWeekTo = document.getElementById("reportWeekTo");
const reportScale = document.getElementById("reportScale");
const reportChartType = document.getElementById("reportChartType");
const reportAggregateScope = document.getElementById("reportAggregateScope");
const reportProjectFilter = document.getElementById("reportProjectFilter");
const reportRoleFilter = document.getElementById("reportRoleFilter");
const reportModeRequestedBtn = document.getElementById("reportModeRequestedBtn");
const reportModeRealBtn = document.getElementById("reportModeRealBtn");
const reportPreviewBtn = document.getElementById("reportPreviewBtn");
const reportSaveViewBtn = document.getElementById("reportSaveViewBtn");
const reportPreview = document.getElementById("reportPreview");
const reportPreviewTitle = document.getElementById("reportPreviewTitle");
const reportPreviewSubtitle = document.getElementById("reportPreviewSubtitle");
const reportLegend = document.getElementById("reportLegend");
const reportSummaryCards = document.getElementById("reportSummaryCards");
const reportDetailArea = document.getElementById("reportDetailArea");
const reportCustomPanel = document.getElementById("reportCustomPanel");
const reportIndicatorAllocation = document.getElementById("reportIndicatorAllocation");
const reportIndicatorAllocationFill = document.getElementById("reportIndicatorAllocationFill");
const reportIndicatorAllocationValue = document.getElementById("reportIndicatorAllocationValue");
const reportIndicatorAllocationDetail = document.getElementById("reportIndicatorAllocationDetail");
const reportIndicatorCoverage = document.getElementById("reportIndicatorCoverage");
const reportIndicatorCoverageFill = document.getElementById("reportIndicatorCoverageFill");
const reportIndicatorCoverageValue = document.getElementById("reportIndicatorCoverageValue");
const reportIndicatorCoverageDetail = document.getElementById("reportIndicatorCoverageDetail");
const reportDemandOverviewCard = document.getElementById("reportDemandOverviewCard");
const reportDemandOverview = document.getElementById("reportDemandOverview");
const reportDemandOverviewLegend = document.getElementById("reportDemandOverviewLegend");
const reportDemandOverviewSubtitle = document.getElementById("reportDemandOverviewSubtitle");
const reportDemandRoleOverviewCard = document.getElementById("reportDemandRoleOverviewCard");
const reportDemandRoleOverview = document.getElementById("reportDemandRoleOverview");
const reportDemandRoleOverviewLegend = document.getElementById("reportDemandRoleOverviewLegend");
const reportDemandRoleOverviewSubtitle = document.getElementById("reportDemandRoleOverviewSubtitle");
const reportUtilizationOverviewCard = document.getElementById("reportUtilizationOverviewCard");
const reportUtilizationOverviewSubtitle = document.getElementById("reportUtilizationOverviewSubtitle");
const reportUtilizationOverviewLegend = document.getElementById("reportUtilizationOverviewLegend");
const reportUtilizationOverview = document.getElementById("reportUtilizationOverview");
const reportCustomInsightOpen = document.getElementById("reportCustomInsightOpen");
const roleList = document.getElementById("roleList");
const timelineEl = document.getElementById("timeline");
const sheetTabs = document.querySelector(".sheet-tabs");
const ganttSplit = document.getElementById("ganttSplit");
const ganttSplitHandle = document.getElementById("ganttSplitHandle");
const ganttGroupBy = document.getElementById("ganttGroupBy");
const ganttOrderWeek = document.getElementById("ganttOrderWeek");
const printPlannerBtn = document.getElementById("printPlannerBtn");
const printResourcesBtn = document.getElementById("printResourcesBtn");
const printGanttBtn = document.getElementById("printGanttBtn");
const plannerPrintModal = document.getElementById("plannerPrintModal");
const plannerPrintCloseBtn = document.getElementById("plannerPrintCloseBtn");
const plannerPrintWeekFrom = document.getElementById("plannerPrintWeekFrom");
const plannerPrintWeekTo = document.getElementById("plannerPrintWeekTo");
const plannerPrintWorkshopWrap = document.getElementById("plannerPrintWorkshopWrap");
const plannerPrintIncludeWorkshop = document.getElementById("plannerPrintIncludeWorkshop");
const plannerPrintRunBtn = document.getElementById("plannerPrintRunBtn");
const ganttDemandProjectFilter = document.getElementById("ganttDemandProjectFilter");
const ganttDemandRoleFilter = document.getElementById("ganttDemandRoleFilter");
const ganttDemandMatrixHead = document.querySelector("#ganttDemandMatrix thead");
const ganttDemandMatrixBody = document.querySelector("#ganttDemandMatrix tbody");
const plannerProjectFilter = document.getElementById("plannerProjectFilter");
const plannerRoleFilter = document.getElementById("plannerRoleFilter");
const showZeroDemandProjects = document.getElementById("showZeroDemandProjects");
const plannerHiddenNotice = document.getElementById("plannerHiddenNotice");
const resourceSearch = document.getElementById("resourceSearch");
const resourceTableSearch = document.getElementById("resourceTableSearch");
const resourceRoleFilter = document.getElementById("resourceRoleFilter");
const resourceStatusFilter = document.getElementById("resourceStatusFilter");
const resourceSortBy = document.getElementById("resourceSortBy");
const resourceColumnsBtn = document.getElementById("resourceColumnsBtn");
const resourceColumnsPanel = document.getElementById("resourceColumnsPanel");
const resourceSaveAllBtn = document.getElementById("resourceSaveAllBtn");
const resourcePool = document.getElementById("resourcePool");
const selectionInfo = document.getElementById("selectionInfo");
const selectionAllocations = document.getElementById("selectionAllocations");
const plannerSideCard = document.getElementById("plannerSideCard");
const plannerSideToggle = document.getElementById("plannerSideToggle");
const demandQty = document.getElementById("demandQty");
const demandWeekFrom = document.getElementById("demandWeekFrom");
const demandWeekTo = document.getElementById("demandWeekTo");
const demandSaveBtn = document.getElementById("demandSaveBtn");
const assignProject = document.getElementById("assignProject");
const assignRole = document.getElementById("assignRole");
const assignWeekFrom = document.getElementById("assignWeekFrom");
const assignWeekTo = document.getElementById("assignWeekTo");
const assignUseExternal = document.getElementById("assignUseExternal");
const assignExternalQty = document.getElementById("assignExternalQty");
const assignResource = document.getElementById("assignResource");
const assignSaveBtn = document.getElementById("assignSaveBtn");
const markUnavailableBtn = document.getElementById("markUnavailableBtn");
const assignClearBtn = document.getElementById("assignClearBtn");
const expandAssignBtn = document.getElementById("expandAssignBtn");
const expandGanttBtn = document.getElementById("expandGanttBtn");
const ganttShowExternalDetail = document.getElementById("ganttShowExternalDetail");
const assignCardDragHandle = document.getElementById("assignCardDragHandle");
const ganttCard = document.getElementById("ganttCard");
const ganttCardDragHandle = document.getElementById("ganttCardDragHandle");
const ganttActionModal = document.getElementById("ganttActionModal");
const ganttActionCard = document.getElementById("ganttActionCard");
const ganttActionDragHandle = document.getElementById("ganttActionDragHandle");
const ganttActionInfo = document.getElementById("ganttActionInfo");
const ganttProjectSelect = document.getElementById("ganttProjectSelect");
const ganttRoleSelect = document.getElementById("ganttRoleSelect");
const ganttWeekFromInput = document.getElementById("ganttWeekFromInput");
const ganttWeekToInput = document.getElementById("ganttWeekToInput");
const ganttConflictBox = document.getElementById("ganttConflictBox");
const ganttActionDeleteBtn = document.getElementById("ganttActionDeleteBtn");
const ganttActionUnassignBtn = document.getElementById("ganttActionUnassignBtn");
const ganttActionSaveBtn = document.getElementById("ganttActionSaveBtn");
const ganttActionCloseBtn = document.getElementById("ganttActionCloseBtn");
const summaryDemand = document.getElementById("summaryDemand");
const summaryAllocated = document.getElementById("summaryAllocated");
const summaryResidual = document.getElementById("summaryResidual");
const summaryResources = document.getElementById("summaryResources");
const demandSourcePath = document.getElementById("demandSourcePath");
const workshopBreakdownBox = document.getElementById("workshopBreakdownBox");
const workshopBreakdownSection = document.getElementById("workshopBreakdownSection");
const projectsAdminTableBody = document.querySelector("#projectsAdminTable tbody");
const projectAdminSearch = document.getElementById("projectAdminSearch");
const projectAdminNewBtn = document.getElementById("projectAdminNewBtn");
const projectAdminSaveBtn = document.getElementById("projectAdminSaveBtn");
const projectCodeInput = document.getElementById("projectCodeInput");
const projectActivityInput = document.getElementById("projectActivityInput");
const projectTypeInput = document.getElementById("projectTypeInput");
const projectClosedInput = document.getElementById("projectClosedInput");
const projectWorkshopRollupInput = document.getElementById("projectWorkshopRollupInput");
const projectDemandRoleInput = document.getElementById("projectDemandRoleInput");
const projectDemandAddRoleBtn = document.getElementById("projectDemandAddRoleBtn");
const projectDemandSaveBtn = document.getElementById("projectDemandSaveBtn");
const projectDemandMatrixHead = document.querySelector("#projectDemandMatrix thead");
const projectDemandMatrixBody = document.querySelector("#projectDemandMatrix tbody");

let state = {
  resources: [],
  projects: [],
  demands: [],
  allocations: [],
  overalls: {},
  unavailability: [],
  report_weekly: [],
  report_role_weekly: [],
  standard_roles: [],
  default_demand_path: "",
  current_demand_path: "",
  default_resources_path: "",
  default_plan_path: "",
  lastDemandDiffs: [],
  lastDemandDiffMap: {},
  lastDemandDiffMeta: {},
};

function isOverallLocked(projectCode) {
  return isOverallProjectCode(projectCode);
}

function allocationWeight(allocation) {
  const w = Number(allocation?.weight);
  return Number.isFinite(w) && w > 0 ? w : 1;
}

function sumAllocationWeights(allocations, week = null) {
  return (allocations || []).reduce((acc, allocation) => {
    if (week === null || week === undefined) return acc + allocationWeight(allocation);
    return acc + allocationEffectiveWeightForWeek(allocation, week);
  }, 0);
}

function formatQuantity(value) {
  const num = Number(value || 0);
  const rounded = Math.round(num * 10) / 10;
  return Number.isInteger(rounded)
    ? rounded.toLocaleString("it-IT")
    : rounded.toLocaleString("it-IT", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
}

function formatAllocatedDisplay(internalAllocated, externalAllocated) {
  const internal = Number(internalAllocated || 0);
  const external = Number(externalAllocated || 0);
  if (external > 0) {
    return `${formatQuantity(internal)}+${formatQuantity(external)}`;
  }
  return formatQuantity(internal);
}

async function logUserAction(action, detail = "") {
  try {
    await api("/api/log", {
      method: "POST",
      body: JSON.stringify({
        actor: "user",
        action,
        detail,
      }),
    });
  } catch (err) {
    // ignore log errors
  }
}

async function loadActivityLog() {
  const rows = await api("/api/logs?limit=200");
  if (!activityLogBody) return;
  activityLogBody.innerHTML = rows
    .map((row) => {
      const who = row.actor === "user" ? "Utente" : "Sistema";
      return `
        <tr>
          <td>${escapeHtml(row.ts || "")}</td>
          <td>${escapeHtml(who)}</td>
          <td>${escapeHtml(row.action || "")}</td>
          <td>${escapeHtml(row.detail || "")}</td>
        </tr>
      `;
    })
    .join("");
}

function openActivityLog() {
  if (!activityLogModal) return;
  activityLogModal.hidden = false;
  loadActivityLog().catch((err) => setStatus(`Errore log: ${err.message}`));
}

function closeActivityLog() {
  if (!activityLogModal) return;
  activityLogModal.hidden = true;
}

const GANTT_SPLIT_STORAGE_KEY = "manpower.ganttSplitHeight";
const GANTT_SHOW_EXT_DETAIL_STORAGE_KEY = "manpower.ganttShowExternalDetail";
let ganttSplitResizeState = null;

const REPORT_PRESETS = [
  {
    id: "demand_allocated",
    label: "Fabbisogno / Allocazioni",
    chartType: "stacked-demand",
    mode: "operativa",
    scale: "1W",
  },
  {
    id: "demand_allocated_role",
    label: "Fabbisogno / Allocazioni per Mansione",
    chartType: "stacked-demand",
    mode: "mansione",
    scale: "1W",
  },
  {
    id: "workforce_utilization",
    label: "Utilizzo Personale in Forza",
    chartType: "utilization",
    mode: "operativa",
    scale: "1W",
  },
  {
    id: "custom_insight",
    label: "Crea Insight Personalizzato",
    chartType: "stacked-demand",
    mode: "operativa",
    scale: "1W",
  },
];

let reportBuilderState = {
  presetId: "demand_allocated",
  weekFrom: 13,
  weekTo: 20,
  scale: "1W",
  chartType: "stacked-demand",
  aggregateScope: "",
  projectFilter: "",
  roleFilter: "",
  mode: "operativa",
};
let reportDetailOpen = false;
let reportDetailPresetId = "demand_allocated";

let selectedTarget = null;
let selectedAdminProjectId = null;
let projectDemandExtraRoles = {};
let adminProjectDraft = false;
let ganttActionState = null;
let ganttDragState = null;
let floatingDragState = null;
let timeline = null;
let isMoving = false;
let deletedResourceIds = new Set();
let selectedAllocationIds = new Set();
let resourceOriginalSnapshot = new Map();
let resourceColumnDrag = null;
let plannerPrintState = null;
let ganttFallbackSort = {
  mode: "resource",
  week: getCurrentWeek(),
  dir: "asc",
};

const DEFAULT_RESOURCE_COLUMN_WIDTHS = [
  48, 90, 220, 130, 130, 110, 110, 110, 130, 130,
  130, 120, 90, 95, 105, 105, 95, 90, 90, 100,
  110, 140, 110, 170, 120, 95, 160, 90, 180, 100,
];

function loadResourceColumnWidths() {
  try {
    const raw = window.localStorage.getItem("resourceColumnWidths");
    const parsed = raw ? JSON.parse(raw) : null;
    if (Array.isArray(parsed) && parsed.length === DEFAULT_RESOURCE_COLUMN_WIDTHS.length) {
      return parsed.map((v, i) => Math.max(48, parseIntOr(v, DEFAULT_RESOURCE_COLUMN_WIDTHS[i])));
    }
  } catch (err) {
    // ignore
  }
  return [...DEFAULT_RESOURCE_COLUMN_WIDTHS];
}

function saveResourceColumnWidths(widths) {
  try {
    window.localStorage.setItem("resourceColumnWidths", JSON.stringify(widths));
  } catch (err) {
    // ignore
  }
}

function applyResourceColumnWidths(widths = loadResourceColumnWidths()) {
  if (!resourcesTable) return;
  resourcesTable.querySelectorAll("colgroup col").forEach((col, idx) => {
    col.style.width = `${widths[idx] || DEFAULT_RESOURCE_COLUMN_WIDTHS[idx] || 100}px`;
  });
}

const RESOURCE_COLUMN_LABELS = [
  "ID", "Codice", "Risorsa", "Mansione 1", "Mansione 2", "Assunzione", "Fine", "Data nascita", "Telefono", "Comune res.",
  "Datore lavoro", "Livello ass.", "Tipo", "â‚¬/h straord", "â‚¬/G Pres. Off", "â‚¬/G Pres. Cant", "PB +FISSO", "â‚¬/h Off.", "â‚¬/h Cant.", "GLOB +FISSO",
  "Tipo DOC", "Numero DOC", "Scadenza DOC", "Email", "Sede", "Livello", "Certificazioni", "No trasferta", "Note operative", "Stato", "Azione",
];
const DEFAULT_VISIBLE_RESOURCE_COLUMNS = [0, 1, 2, 3, 4, 5, 6, 10, 11, 12, 20, 22, 24, 28, 29, 30];

function normalizeVisibleResourceColumns(cols) {
  const forced = new Set([2, 29, 30]);
  const set = new Set(
    (Array.isArray(cols) ? cols : DEFAULT_VISIBLE_RESOURCE_COLUMNS)
      .map((v) => parseIntOr(v, -1))
      .filter((v) => v >= 0)
  );
  forced.forEach((v) => set.add(v));
  return [...set].sort((a, b) => a - b);
}

function loadResourceVisibleColumns() {
  try {
    const raw = window.localStorage.getItem("resourceVisibleColumns");
    return normalizeVisibleResourceColumns(raw ? JSON.parse(raw) : DEFAULT_VISIBLE_RESOURCE_COLUMNS);
  } catch (err) {
    return normalizeVisibleResourceColumns(DEFAULT_VISIBLE_RESOURCE_COLUMNS);
  }
}

function saveResourceVisibleColumns(cols) {
  try {
    window.localStorage.setItem("resourceVisibleColumns", JSON.stringify(normalizeVisibleResourceColumns(cols)));
  } catch (err) {
    // ignore
  }
}

function applyResourceColumnVisibility() {
  if (!resourcesTable) return;
  const visible = new Set(loadResourceVisibleColumns());
  resourcesTable.querySelectorAll("colgroup col").forEach((col, idx) => {
    col.style.display = visible.has(idx) ? "" : "none";
  });
  resourcesTable.querySelectorAll("thead tr").forEach((tr) => {
    [...tr.children].forEach((cell, idx) => {
      cell.style.display = visible.has(idx) ? "" : "none";
    });
  });
  resourcesTable.querySelectorAll("tbody tr").forEach((tr) => {
    [...tr.children].forEach((cell, idx) => {
      cell.style.display = visible.has(idx) ? "" : "none";
    });
  });
}

function renderResourceColumnChooser() {
  if (!resourceColumnsPanel) return;
  const selected = new Set(loadResourceVisibleColumns());
  resourceColumnsPanel.innerHTML = `
    <div class="panel-actions">
      <button type="button" class="mini-btn" data-colset="base">Base</button>
      <button type="button" class="mini-btn" data-colset="all">Tutte</button>
    </div>
    ${RESOURCE_COLUMN_LABELS.map((label, idx) => `
      <label>
        <input type="checkbox" data-col-index="${idx}" ${selected.has(idx) ? "checked" : ""} ${[2, 29, 30].includes(idx) ? "disabled" : ""}/>
        <span>${label}</span>
      </label>
    `).join("")}
  `;
  resourceColumnsPanel.querySelectorAll("[data-colset]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const cols = btn.dataset.colset === "all"
        ? RESOURCE_COLUMN_LABELS.map((_, idx) => idx)
        : DEFAULT_VISIBLE_RESOURCE_COLUMNS;
      saveResourceVisibleColumns(cols);
      renderResourceColumnChooser();
      applyResourceColumnVisibility();
      bindResourceColumnResizers();
    });
  });
  resourceColumnsPanel.querySelectorAll("[data-col-index]").forEach((chk) => {
    chk.addEventListener("change", () => {
      const cols = [...resourceColumnsPanel.querySelectorAll("[data-col-index]:checked")].map((el) => parseIntOr(el.dataset.colIndex, -1));
      saveResourceVisibleColumns(cols);
      applyResourceColumnVisibility();
      bindResourceColumnResizers();
    });
  });
}

function bindResourceColumnResizers() {
  if (!resourcesTableHead || !resourcesTable) return;
  const widths = loadResourceColumnWidths();
  applyResourceColumnWidths(widths);
  resourcesTableHead.querySelectorAll("th").forEach((th, idx) => {
    th.querySelector(".col-resizer")?.remove();
    const grip = document.createElement("span");
    grip.className = "col-resizer";
    grip.dataset.colIndex = String(idx);
    th.appendChild(grip);
    grip.addEventListener("pointerdown", (ev) => {
      resourceColumnDrag = {
        pointerId: ev.pointerId,
        index: idx,
        startX: ev.clientX,
        startWidth: widths[idx] || th.getBoundingClientRect().width,
      };
      document.body.classList.add("is-col-resizing");
      ev.preventDefault();
      ev.stopPropagation();
    });
  });

  if (!window.__resourceColResizeBound) {
    window.addEventListener("pointermove", (ev) => {
      if (!resourceColumnDrag) return;
      const next = Math.max(48, Math.round(resourceColumnDrag.startWidth + (ev.clientX - resourceColumnDrag.startX)));
      widths[resourceColumnDrag.index] = next;
      applyResourceColumnWidths(widths);
    });
    const finish = () => {
      if (!resourceColumnDrag) return;
      saveResourceColumnWidths(widths);
      document.body.classList.remove("is-col-resizing");
      resourceColumnDrag = null;
    };
    window.addEventListener("pointerup", finish);
    window.addEventListener("pointercancel", finish);
    window.__resourceColResizeBound = true;
  }
}

function bindResourceTableKeyboardNavigation() {
  resourcesTableBody.querySelectorAll("input, select, textarea").forEach((el) => {
    el.addEventListener("keydown", (ev) => {
      const td = ev.target.closest("td");
      const tr = ev.target.closest("tr");
      if (!td || !tr) return;
      const rowIndex = [...resourcesTableBody.querySelectorAll("tr")].indexOf(tr);
      const cellIndex = [...tr.children].indexOf(td);
      let target = null;
      if (ev.key === "ArrowDown") {
        const nextRow = resourcesTableBody.querySelectorAll("tr")[rowIndex + 1];
        target = nextRow?.children?.[cellIndex]?.querySelector("input, select, textarea") || null;
      } else if (ev.key === "ArrowUp") {
        const prevRow = resourcesTableBody.querySelectorAll("tr")[rowIndex - 1];
        target = prevRow?.children?.[cellIndex]?.querySelector("input, select, textarea") || null;
      } else if (ev.key === "ArrowRight" && (ev.target.selectionStart ?? 0) === String(ev.target.value || "").length) {
        target = td.nextElementSibling?.querySelector("input, select, textarea") || null;
      } else if (ev.key === "ArrowLeft" && (ev.target.selectionStart ?? 0) === 0) {
        target = td.previousElementSibling?.querySelector("input, select, textarea") || null;
      }
      if (target) {
        ev.preventDefault();
        target.focus();
        if (typeof target.select === "function" && target.tagName === "INPUT") target.select();
      }
    });
  });
}
let undoStack = [];
let isUndoing = false;

function updateUndoButton() {
  if (!undoBtn) return;
  undoBtn.disabled = undoStack.length === 0;
  const label = undoStack.length > 0 ? undoStack[undoStack.length - 1].label : "";
  undoBtn.title = label ? `Annulla ultima modifica: ${label} (Ctrl+Z)` : "Annulla ultima modifica (Ctrl+Z)";
}

function pushUndo(entry) {
  if (isUndoing || !entry?.undo) return;
  undoStack.push(entry);
  if (undoStack.length > 50) undoStack.shift();
  updateUndoButton();
}

async function undoLastAction() {
  const entry = undoStack.pop();
  updateUndoButton();
  if (!entry) {
    setStatus("Nessuna modifica da annullare.");
    return;
  }
  try {
    isUndoing = true;
    await entry.undo();
    await loadAll(true);
    setStatus(entry.undoMessage || `Annullata: ${entry.label}.`);
  } catch (err) {
    console.error(err);
    setStatus(err.message || "Errore annullamento modifica");
  } finally {
    isUndoing = false;
    updateUndoButton();
  }
}

function captureViewState() {
  const plannerWrap = document.querySelector(".planner-wrap");
  const ganttWrap = timelineEl?.querySelector(".gantt-fallback-wrap");
  const activeSheet = document.querySelector(".sheet-btn.active")?.dataset.sheet || "planner";
  let timelineWindow = null;
  try {
    if (timeline && typeof timeline.getWindow === "function") {
      timelineWindow = timeline.getWindow();
    }
  } catch (err) {
    timelineWindow = null;
  }
  return {
    activeSheet,
    planner: plannerWrap ? { left: plannerWrap.scrollLeft, top: plannerWrap.scrollTop } : null,
    resourcePool: resourcePool ? { left: resourcePool.scrollLeft, top: resourcePool.scrollTop } : null,
    selectionAllocations: selectionAllocations ? { left: selectionAllocations.scrollLeft, top: selectionAllocations.scrollTop } : null,
    workshopBreakdown: workshopBreakdownBox ? { left: workshopBreakdownBox.scrollLeft, top: workshopBreakdownBox.scrollTop } : null,
    gantt: ganttWrap ? { left: ganttWrap.scrollLeft, top: ganttWrap.scrollTop } : null,
    timelineWindow,
  };
}

function restoreViewState(viewState) {
  if (!viewState) return;
  const activateBtn = document.querySelector(`.sheet-btn[data-sheet="${viewState.activeSheet}"]`);
  if (activateBtn && !activateBtn.classList.contains("active")) {
    activateBtn.click();
  }
  const plannerWrap = document.querySelector(".planner-wrap");
  const ganttWrap = timelineEl?.querySelector(".gantt-fallback-wrap");
  if (plannerWrap && viewState.planner) {
    plannerWrap.scrollLeft = viewState.planner.left;
    plannerWrap.scrollTop = viewState.planner.top;
  }
  if (resourcePool && viewState.resourcePool) {
    resourcePool.scrollLeft = viewState.resourcePool.left;
    resourcePool.scrollTop = viewState.resourcePool.top;
  }
  if (selectionAllocations && viewState.selectionAllocations) {
    selectionAllocations.scrollLeft = viewState.selectionAllocations.left;
    selectionAllocations.scrollTop = viewState.selectionAllocations.top;
  }
  if (workshopBreakdownBox && viewState.workshopBreakdown) {
    workshopBreakdownBox.scrollLeft = viewState.workshopBreakdown.left;
    workshopBreakdownBox.scrollTop = viewState.workshopBreakdown.top;
  }
  if (ganttWrap && viewState.gantt) {
    ganttWrap.scrollLeft = viewState.gantt.left;
    ganttWrap.scrollTop = viewState.gantt.top;
  }
  if (viewState.timelineWindow && timeline && typeof timeline.setWindow === "function") {
    try {
      timeline.setWindow(viewState.timelineWindow.start, viewState.timelineWindow.end, { animation: false });
    } catch (err) {
      // ignore restore issues
    }
  }
}

function getCurrentWeek() {
  return weekFromDate(new Date());
}

function visibleWeeks() {
  return Array.from({ length: 53 - getCurrentWeek() }, (_, i) => getCurrentWeek() + i);
}

function externalDetailEnabled() {
  return !!ganttShowExternalDetail?.checked;
}

function loadGanttExternalDetailFlag() {
  if (!ganttShowExternalDetail) return;
  try {
    const saved = window.localStorage.getItem(GANTT_SHOW_EXT_DETAIL_STORAGE_KEY);
    ganttShowExternalDetail.checked = saved === "1";
  } catch (err) {
    ganttShowExternalDetail.checked = false;
  }
}

function persistGanttExternalDetailFlag(enabled) {
  try {
    window.localStorage.setItem(GANTT_SHOW_EXT_DETAIL_STORAGE_KEY, enabled ? "1" : "0");
  } catch (err) {
    // ignore storage errors
  }
}

function extTooltipFromProjectMap(projectMap) {
  const rows = [...projectMap.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
  if (!rows.length) return "Nessun external in settimana.";
  return rows.map(([project, qty]) => `${project}: ${qty}`).join(" | ");
}

function buildExternalWeekMatrix(weeks) {
  const extResourceIds = new Set(
    (state.resources || []).filter((r) => isExternalResource(r)).map((r) => Number(r.id))
  );
  const extAllocations = (state.allocations || []).filter((a) => extResourceIds.has(Number(a.resource_id)));
  const aggregateByWeek = new Map();
  const byRole = new Map();
  weeks.forEach((week) => {
    const active = extAllocations.filter(
      (a) => Number(a.week_from) <= Number(week) && Number(a.week_to) >= Number(week)
    );
    const projectMap = new Map();
    active.forEach((a) => projectMap.set(a.project, (projectMap.get(a.project) || 0) + 1));
    aggregateByWeek.set(week, { qty: active.length, projects: projectMap });

    active.forEach((a) => {
      const role = (a.role || "").trim().toUpperCase() || "EXT";
      if (!byRole.has(role)) byRole.set(role, new Map());
      const roleWeek = byRole.get(role);
      if (!roleWeek.has(week)) roleWeek.set(week, { qty: 0, projects: new Map() });
      const cell = roleWeek.get(week);
      cell.qty += 1;
      cell.projects.set(a.project, (cell.projects.get(a.project) || 0) + 1);
    });
  });
  const roleRows = [...byRole.keys()].sort((a, b) => a.localeCompare(b));
  return { aggregateByWeek, byRole, roleRows };
}

function ensureGanttOrderWeekOptions() {
  if (!ganttOrderWeek) return;
  if (ganttOrderWeek.options.length > 0) return;
  for (let w = 1; w <= 52; w += 1) {
    const opt = document.createElement("option");
    opt.value = String(w);
    opt.textContent = `W${w}`;
    ganttOrderWeek.appendChild(opt);
  }
  ganttOrderWeek.value = String(getCurrentWeek());
}

function ganttOrderWeekValue() {
  if (!ganttOrderWeek) return parseWeek(ganttFallbackSort.week) || getCurrentWeek();
  const w = parseWeek(ganttOrderWeek.value);
  return w !== null ? w : getCurrentWeek();
}

function setGanttFallbackSortResource() {
  if (ganttFallbackSort.mode === "resource") {
    ganttFallbackSort.dir = ganttFallbackSort.dir === "asc" ? "desc" : "asc";
    return;
  }
  ganttFallbackSort.mode = "resource";
  ganttFallbackSort.dir = "asc";
}

function setGanttFallbackSortWeek(week) {
  const targetWeek = parseWeek(week);
  if (targetWeek === null) return;
  if (ganttFallbackSort.mode === "week" && Number(ganttFallbackSort.week) === Number(targetWeek)) {
    ganttFallbackSort.dir = ganttFallbackSort.dir === "desc" ? "asc" : "desc";
  } else {
    ganttFallbackSort.mode = "week";
    ganttFallbackSort.week = targetWeek;
    ganttFallbackSort.dir = "desc";
  }
}

function weekPopularityMap(resources, week) {
  const map = new Map();
  resources.forEach((r) => {
    const active = state.allocations.find(
      (a) =>
        Number(a.resource_id) === Number(r.id) &&
        Number(a.week_from) <= Number(week) &&
        Number(a.week_to) >= Number(week)
    );
    if (!active) return;
    const key = String(active.project || "");
    map.set(key, (map.get(key) || 0) + 1);
  });
  return map;
}

function sortResourcesForGantt(resources) {
  const list = resources.slice();
  if (ganttFallbackSort.mode === "resource") {
    const direction = ganttFallbackSort.dir === "desc" ? -1 : 1;
    return list.sort((a, b) => direction * a.name.localeCompare(b.name));
  }
  const week = parseWeek(ganttFallbackSort.week) || getCurrentWeek();
  const popularity = weekPopularityMap(list, week);
  const direction = ganttFallbackSort.dir === "asc" ? 1 : -1;
  return list.sort((a, b) => {
    const activeA = state.allocations.find(
      (x) =>
        Number(x.resource_id) === Number(a.id) &&
        Number(x.week_from) <= week &&
        Number(x.week_to) >= week
    );
    const activeB = state.allocations.find(
      (x) =>
        Number(x.resource_id) === Number(b.id) &&
        Number(x.week_from) <= week &&
        Number(x.week_to) >= week
    );
    const hasA = Boolean(activeA);
    const hasB = Boolean(activeB);
    if (hasA !== hasB) return hasA ? -1 : 1;
    if (!hasA && !hasB) return a.name.localeCompare(b.name);
    const popA = popularity.get(String(activeA.project || "")) || 0;
    const popB = popularity.get(String(activeB.project || "")) || 0;
    if (popA !== popB) return direction * (popB - popA);
    const projCmp = String(activeA.project || "").localeCompare(String(activeB.project || ""));
    if (projCmp !== 0) return projCmp;
    return a.name.localeCompare(b.name);
  });
}

function getDataSetCtor() {
  return window.vis?.DataSet || window.visData?.DataSet || null;
}

function getTimelineCtor() {
  return (
    window.vis?.Timeline ||
    window.vis?.timeline?.Timeline ||
    window.visTimeline?.Timeline ||
    window.Timeline ||
    null
  );
}

function setStatus(text) {
  statusText.textContent = text;
}

function safeOn(element, eventName, handler, label = eventName) {
  if (!element) {
    console.warn(`Listener non registrato: ${label}`);
    return;
  }
  element.addEventListener(eventName, (...args) => {
    try {
      return handler(...args);
    } catch (err) {
      console.error(`Errore listener ${label}:`, err);
      setStatus(`Errore ${label}: ${err.message}`);
    }
  });
}

function safeInit(label, fn) {
  try {
    return fn();
  } catch (err) {
    console.error(`Errore init ${label}:`, err);
    setStatus(`Errore init ${label}: ${err.message}`);
    return null;
  }
}

async function api(url, options = {}) {
  const res = await fetch(url, {
    headers: { "Content-Type": "application/json" },
    ...options,
  });
  if (!res.ok) {
    const msg = await res.text();
    throw new Error(msg || `Errore ${res.status}`);
  }
  return res.json();
}

function escapeHtml(text) {
  return String(text ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function parseWeek(value) {
  const raw = String(value ?? "").trim().toUpperCase();
  if (!raw) return null;
  const n = Number(raw.startsWith("W") ? raw.slice(1) : raw);
  if (!Number.isFinite(n) || n < 1 || n > 52) return null;
  return Math.trunc(n);
}

function parseWeekOrDate(value) {
  const direct = parseWeek(value);
  if (direct !== null) return direct;
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  if (["N.D.", "ND", "N/D"].includes(raw.toUpperCase())) return 0;
  const match = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  let date = null;
  if (match) {
    date = new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
  } else if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    date = new Date(raw);
  }
  if (!date || Number.isNaN(date.getTime())) return null;
  return weekFromDate(date);
}

function parseIntOr(value, fallback = 0) {
  const n = Number(value);
  return Number.isFinite(n) ? Math.trunc(n) : fallback;
}

function normalizeKey(value) {
  return String(value || "").trim().toUpperCase();
}

function weeksForRange(weekFrom = reportBuilderState.weekFrom, weekTo = reportBuilderState.weekTo) {
  const current = getCurrentWeek();
  const start = Math.max(current, parseIntOr(weekFrom, current));
  const end = Math.min(52, Math.max(start, parseIntOr(weekTo, start)));
  return Array.from({ length: end - start + 1 }, (_, i) => start + i);
}

function getSavedReportViews() {
  try {
    const raw = window.localStorage.getItem("savedReportViews");
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    return [];
  }
}

function saveSavedReportViews(views) {
  try {
    window.localStorage.setItem("savedReportViews", JSON.stringify(views));
  } catch (err) {
    // ignore
  }
}

function bindSheetTabs() {
  document.querySelectorAll(".sheet-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const target = btn.dataset.sheet;
      document.querySelectorAll(".sheet-btn").forEach((b) => b.classList.remove("active"));
      document.querySelectorAll(".sheet-panel").forEach((p) => p.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById(`sheet-${target}`)?.classList.add("active");
      window.requestAnimationFrame(updateWorkpaneHeight);
    });
  });
}

function updateWorkpaneHeight() {
  if (!sheetTabs) return;
  const tabsRect = sheetTabs.getBoundingClientRect();
  const available = Math.max(420, Math.floor(window.innerHeight - tabsRect.bottom - 12));
  document.documentElement.style.setProperty("--workpane-height", `${available}px`);
}

function bindWorkpaneLayout() {
  const schedule = () => window.requestAnimationFrame(updateWorkpaneHeight);
  schedule();
  window.addEventListener("resize", schedule);
  window.addEventListener("orientationchange", schedule);
}

function startOfWeek(week) {
  const jan4 = new Date(2026, 0, 4);
  const jan4Day = jan4.getDay() || 7;
  const mondayW1 = new Date(jan4);
  mondayW1.setDate(jan4.getDate() - (jan4Day - 1));
  const d = new Date(mondayW1);
  d.setDate(mondayW1.getDate() + (week - 1) * 7);
  return d;
}

function weekFromDate(date) {
  const target = new Date(date);
  const jan4 = new Date(2026, 0, 4);
  const jan4Day = jan4.getDay() || 7;
  const mondayW1 = new Date(jan4);
  mondayW1.setDate(jan4.getDate() - (jan4Day - 1));
  const diffMs = target - mondayW1;
  const week = Math.floor(diffMs / (7 * 24 * 3600 * 1000)) + 1;
  return Math.max(1, Math.min(52, week));
}

function addDays(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function formatShortDate(date) {
  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  return `${dd}/${mm}`;
}

function weekLabelTooltip(week) {
  const start = startOfWeek(week);
  const end = addDays(start, 6);
  return `W${week} = ${formatShortDate(start)}-${formatShortDate(end)}`;
}

function resourceVisibleInHorizon(resource) {
  const currentWeek = getCurrentWeek();
  const start = parseWeekOrDate(resource.hire_date) || 1;
  const end = parseWeekOrDate(resource.end_date);
  if (end !== null && end <= 0) return false;
  if (end !== null && end < currentWeek) return false;
  return start <= 52;
}

function resourceVisibleInGantt(resource) {
  const currentWeek = getCurrentWeek();
  const start = parseWeekOrDate(resource.hire_date) || 1;
  const end = parseWeekOrDate(resource.end_date);
  if (end !== null && end <= 0) return false;
  if (end !== null && end < 999 && end + 4 < currentWeek) return false;
  return start <= 52;
}

function isExternalResource(resource) {
  if (!resource) return false;
  const name = String(resource.name || "").trim().toUpperCase();
  const employer = String(resource.employer || "").trim().toUpperCase();
  return Boolean(resource.is_external) || employer === "EXT" || name.endsWith("-EXT");
}

function normalizeRoleKey(value) {
  return String(value || "")
    .toUpperCase()
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z0-9]+/g, " ")
    .trim();
}

function roleMatches(resource, role) {
  const target = normalizeRoleKey(role);
  if (!target) return false;
  return [resource.role1, resource.role2].some((r) => normalizeRoleKey(r) === target);
}

function resourceHasAnyRole(resource) {
  return [resource.role1, resource.role2].some((role) => String(role || "").trim());
}

function resourceActive(resource, week = null) {
  if (isExternalResource(resource)) return true;
  const start = parseWeekOrDate(resource.hire_date) || 1;
  const end = parseWeekOrDate(resource.end_date);
  if (resource.status === "CESSATO") return false;
  if (week === null) return true;
  if (week < start) return false;
  if (end !== null && end <= 0) return false;
  if (end !== null && end < 999 && week > end) return false;
  return true;
}

function projectById(id) {
  return state.projects.find((p) => Number(p.id) === Number(id));
}

function normalizeProjectCode(value) {
  return String(value || "").trim().toUpperCase();
}

function isOverallProjectCode(projectCode = "") {
  const code = normalizeProjectCode(projectCode);
  if (!code) return false;
  const overalls = Object.values(state.overalls || {});
  return overalls.some((o) => normalizeProjectCode(o?.name) === code || normalizeProjectCode(o?.id) === code);
}

function buildOverallOfficina() {
  const officinaCodes = [...new Set((state.projects || [])
    .map((p) => String(p.code || "").trim())
    .filter((code) => normalizeProjectCode(code).includes("OFFICINA")))];
  const prev = state.overalls || {};
  state.overalls = { ...prev };
  state.overalls.OVERALL_OFFICINA = {
    id: "OVERALL_OFFICINA",
    name: "OVERALL OFFICINA",
    project_codes: officinaCodes,
  };
}

function createCustomOverall(name, projectCodes) {
  state.overalls = state.overalls || {};
  const used = Object.values(state.overalls).flatMap((o) => (o.project_codes || []).map((p) => normalizeProjectCode(p)));
  const alreadyUsed = (projectCodes || []).filter((p) => used.includes(normalizeProjectCode(p)));
  if (alreadyUsed.length > 0) {
    throw new Error("Alcune commesse sono gia assegnate ad un altro overall");
  }
  const id = `OVERALL_${Date.now()}`;
  state.overalls[id] = {
    id,
    name,
    project_codes: (projectCodes || []).map((p) => String(p || "").trim()).filter(Boolean),
  };
  return state.overalls[id];
}

function getOwningOverall(projectCode) {
  const code = normalizeProjectCode(projectCode);
  if (!code) return null;
  const overalls = Array.isArray(state.overalls) ? state.overalls : Object.values(state.overalls || {});
  for (const overall of overalls) {
    const projects = Array.isArray(overall?.projects)
      ? overall.projects
      : Array.isArray(overall?.project_codes)
      ? overall.project_codes
      : [];
    const ownsProject = projects.some((p) => normalizeProjectCode(p) === code);
    if (!ownsProject) continue;
    const overallCode = normalizeProjectCode(overall?.code || overall?.id || overall?.name || "");
    return {
      ...overall,
      code: overallCode || String(overall?.code || overall?.id || overall?.name || "").trim(),
    };
  }
  return null;
}

function isAllocableProjectCode(projectCode) {
  const code = normalizeProjectCode(projectCode);
  if (!code) return false;
  if (isOverallProjectCode(code)) return true;
  const owning = getOwningOverall(code);
  if (!owning) return true;
  const owningCode = normalizeProjectCode(owning?.code || owning?.id || owning?.name || "");
  return owningCode === code;
}

function allocationEffectiveWeightForWeek(allocation, week) {
  const currentWeek = Number(week);
  if (!Number.isFinite(currentWeek)) return 0;
  if (Number(allocation?.week_from) > currentWeek || Number(allocation?.week_to) < currentWeek) return 0;
  const resource = state.resources.find((r) => Number(r.id) === Number(allocation.resource_id));
  if (!resource) return 0;
  if (isExternalResource(resource)) return 1;
  const activeInternal = state.allocations.filter((a) => {
    if (Number(a.resource_id) !== Number(allocation.resource_id)) return false;
    const otherRes = state.resources.find((r) => Number(r.id) === Number(a.resource_id));
    if (!otherRes || isExternalResource(otherRes)) return false;
    return Number(a.week_from) <= currentWeek && Number(a.week_to) >= currentWeek;
  });
  if (activeInternal.length > 2) {
    throw new Error("Una risorsa interna non può avere più di 2 allocazioni nella stessa settimana");
  }
  if (activeInternal.length === 2) return 0.5;
  return 1;
}

function validateInternalAllocationRange(resourceId, weekFrom, weekTo, ignoreAllocationId = null) {
  const resource = state.resources.find((r) => Number(r.id) === Number(resourceId));
  if (!resource || isExternalResource(resource)) return;
  const startWeek = Number(weekFrom);
  const endWeek = Number(weekTo);
  for (let week = startWeek; week <= endWeek; week += 1) {
    const overlaps = state.allocations.filter((a) => {
      if (Number(a.resource_id) !== Number(resourceId)) return false;
      if (ignoreAllocationId && Number(a.id) === Number(ignoreAllocationId)) return false;
      const otherRes = state.resources.find((r) => Number(r.id) === Number(a.resource_id));
      if (!otherRes || isExternalResource(otherRes)) return false;
      return Number(a.week_from) <= week && Number(a.week_to) >= week;
    });
    if (overlaps.length >= 2) {
      throw new Error("Una risorsa interna non può avere più di 2 allocazioni nella stessa settimana");
    }
  }
}

function isDoubleAllocation(resourceId, week) {
  const resource = state.resources.find((r) => Number(r.id) === Number(resourceId));
  if (!resource || isExternalResource(resource)) return false;
  const allocs = state.allocations.filter((a) =>
    Number(a.resource_id) === Number(resourceId) &&
    Number(a.week_from) <= Number(week) &&
    Number(a.week_to) >= Number(week) &&
    !isExternalResource(state.resources.find((r) => Number(r.id) === Number(a.resource_id)))
  );
  return allocs.length === 2;
}

function allocationHasDoubleAllocation(allocation) {
  const resource = state.resources.find((r) => Number(r.id) === Number(allocation.resource_id));
  if (!resource || isExternalResource(resource)) return false;
  for (let week = Number(allocation.week_from); week <= Number(allocation.week_to); week += 1) {
    if (isDoubleAllocation(allocation.resource_id, week)) return true;
  }
  return false;
}

function normalizeResourceAllocations(resourceId) {
  const resource = state.resources.find((r) => Number(r.id) === Number(resourceId));
  if (!resource || isExternalResource(resource)) return;
  state.allocations
    .filter((a) => Number(a.resource_id) === Number(resourceId))
    .forEach((a) => {
      if (!Number.isFinite(Number(a.weight)) || Number(a.weight) <= 0) a.weight = 1;
    });
}

async function persistResourceAllocationWeights(resourceId) {
  const resource = state.resources.find((r) => Number(r.id) === Number(resourceId));
  if (!resource || isExternalResource(resource)) return;
  const resourceAllocs = state.allocations.filter((a) => Number(a.resource_id) === Number(resourceId));
  for (const allocation of resourceAllocs) {
    await api(`/api/allocations/${allocation.id}`, {
      method: "PATCH",
      body: JSON.stringify({ weight: allocationWeight(allocation) }),
    });
  }
}

function plannerProjectMeta(projectId, fallbackProjectName = "") {
  const project = projectById(projectId);
  const projectCode = fallbackProjectName || project?.code || "";
  const owner = isOverallProjectCode(projectCode) ? null : getOwningOverall(projectCode);
  if (owner) {
    const ownerProject = (state.projects || []).find((p) =>
      normalizeProjectCode(p.code) === normalizeProjectCode(owner.name) ||
      normalizeProjectCode(p.code) === normalizeProjectCode(owner.id)
    );
    return {
      project_id: ownerProject ? Number(ownerProject.id) : Number(projectId),
      project: owner.name || owner.id,
    };
  }
  return {
    project_id: Number(projectId),
    project: projectCode,
  };
}

function isWorkshopChildProject(project) {
  return (
    project &&
    project.type === "WS" &&
    Number(project.workshop_rollup) === 1 &&
    project.code !== "OVERALL OFFICINA" &&
    project.code !== "WS OVERALL"
  );
}

function hasWorkshopChildDemand(role, week) {
  return state.demands.some((d) => {
    const project = projectById(d.project_id);
    return isWorkshopChildProject(project) && d.role === role && Number(d.week) === Number(week) && Number(d.qty || 0) > 0;
  });
}

function currentDemandQty(projectId, role, week) {
  const row = state.demands.find(
    (d) => Number(d.project_id) === Number(projectId) && d.role === role && Number(d.week) === Number(week)
  );
  return row ? Number(row.qty || 0) : 0;
}

function projectDemandBreakdown(projectId, role, week) {
  return state.demands
    .filter(
      (d) =>
        Number(d.project_id) === Number(projectId) &&
        d.role === role &&
        Number(d.week) === Number(week) &&
        Number(d.qty || 0) > 0
    )
    .sort((a, b) => Number(b.qty || 0) - Number(a.qty || 0) || a.project.localeCompare(b.project));
}

function currentWorkshopBreakdown(role, week) {
  const rows = state.demands
    .filter((d) => {
      const project = projectById(d.project_id);
      return (
        isWorkshopChildProject(project) &&
        d.role === role &&
        Number(d.week) === Number(week) &&
        Number(d.qty || 0) > 0
      );
    })
    .sort((a, b) => Number(b.qty || 0) - Number(a.qty || 0) || a.project.localeCompare(b.project));
  return rows;
}

function warningLabelsForWeek(allocation, week) {
  return (allocation.warning_segments || [])
    .filter((seg) => Number(seg.week_from) <= week && Number(seg.week_to) >= week)
    .map((seg) => {
      if (String(seg.type || "").toUpperCase() === "MANSIONE NON COERENTE") {
        return "FUORI MANSIONE";
      }
      const from = Math.max(Number(seg.week_from), week);
      const to = Math.min(Number(seg.week_to), week);
      if (from === Number(seg.week_from) && to === Number(seg.week_to)) return seg.type;
      return `${seg.type} W${from}${to > from ? `-W${to}` : ""}`;
    });
}

function compactWarningCode(type = "") {
  const raw = String(type || "").trim().toUpperCase();
  if (!raw) return "";
  if (raw.includes("MANSIONE") || raw.includes("FUORI")) return "NO-MANS";
  if (raw.includes("INDISP")) return "IND";
  if (raw.includes("NO FABB")) return "NO-FABB";
  if (raw.includes("SOVRAPP")) return "OVER";
  if (raw.includes("CONTRATTO") || raw.includes("CESSATO")) return "NA";
  return raw.replace(/\s+/g, "-").slice(0, 16);
}

function warningCodesForAllocationRange(allocation, weekFrom, weekTo) {
  const from = Number(weekFrom);
  const to = Number(weekTo);
  if (!Number.isFinite(from) || !Number.isFinite(to)) return [];
  const rangeFrom = Math.min(from, to);
  const rangeTo = Math.max(from, to);
  const set = new Set();
  (allocation?.warning_segments || []).forEach((seg) => {
    const segFrom = Number(seg.week_from);
    const segTo = Number(seg.week_to);
    if (segFrom > rangeTo || segTo < rangeFrom) return;
    const code = compactWarningCode(seg.type);
    if (code) set.add(code);
  });
  return [...set];
}

function demandQtyForPlannerProjectWeek(projectCode, role, week) {
  const targetProject = normalizeProjectCode(projectCode);
  const targetRole = normalizeRoleKey(role);
  return state.demands.reduce((acc, demand) => {
    const meta = plannerProjectMeta(demand.project_id, demand.project);
    if (normalizeProjectCode(meta.project) !== targetProject) return acc;
    if (normalizeRoleKey(demand.role) !== targetRole) return acc;
    if (Number(demand.week) !== Number(week)) return acc;
    return acc + Number(demand.qty || 0);
  }, 0);
}

function formatRoleStatus(resource, role) {
  if (!resource) return "NO-MANS";
  const targetRole = String(role || "").trim();
  if (!targetRole) return resourceHasAnyRole(resource) ? "OK" : "NO-MANS";
  return roleMatches(resource, targetRole) ? "OK" : "NO-MANS";
}

function formatShortPersonName(fullName) {
  const cleaned = String(fullName || "")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
  if (!cleaned) return "";
  const parts = cleaned.split(" ").filter(Boolean);
  if (parts.length < 2) return cleaned;
  const firstName = parts[parts.length - 1];
  const surname = parts.slice(0, -1).join(" ");
  return `${surname} ${firstName.slice(0, 3)}`;
}

function formatShortRole(role) {
  const raw = String(role || "").trim();
  if (!raw) return "";
  const key = normalizeRoleKey(raw);
  const map = {
    "SALDATORE TIG ELETTRODO": "SALD TIG",
    "SALDATORE FILO": "SALD FILO",
    "CAPO CANTIERE": "CAPO CANT",
    "CARPENTIERE": "CARP",
    "COIBENTATORE": "COIB",
    "TUBISTA": "TUB",
    "MONTATORE": "MONT",
    "MECCANICO SERVICE": "MECC SERV",
    "QUALITY CONTROL WELDING INSPECTOR": "QC/WI",
    "GENERICO": "GEN",
    "SOLLEVAMENTI": "SOLL",
  };
  if (map[key]) return map[key];
  const tokens = key.split(" ").filter(Boolean);
  if (!tokens.length) return raw.toUpperCase();
  if (tokens.length === 1) return tokens[0].slice(0, 5);
  return `${tokens[0].slice(0, 4)} ${tokens[1].slice(0, 4)}`.trim();
}

function formatAllocationWeightLabel(resource, allocation, weekFrom, weekTo) {
  if (!resource) return "";
  if (isExternalResource(resource)) return "EXT";
  const fromRaw = Number(weekFrom ?? allocation?.week_from);
  const toRaw = Number(weekTo ?? allocation?.week_to ?? fromRaw);
  if (!Number.isFinite(fromRaw) || !Number.isFinite(toRaw)) return "";
  const from = Math.min(fromRaw, toRaw);
  const to = Math.max(fromRaw, toRaw);
  if (!allocation) {
    const conflicts = allocationConflict(resource.id, from, to);
    if (conflicts.length === 0) return "";
    if (conflicts.length === 1) return "1:1";
    return "1/2";
  }
  let hasAny = false;
  for (let week = from; week <= to; week += 1) {
    const weight = allocationEffectiveWeightForWeek(allocation, week);
    if (weight > 0) hasAny = true;
    if (Math.abs(weight - 0.5) < 0.001) return "1/2";
  }
  return hasAny ? "1:1" : "";
}

function formatResourceCompactStatus(resource, role, weekFrom, weekTo) {
  if (!resource) return "";
  if (isExternalResource(resource)) return "EXT";
  const fromRaw = Number(weekFrom);
  const toRaw = Number(weekTo ?? weekFrom);
  const hasRange = Number.isFinite(fromRaw) && Number.isFinite(toRaw);
  const from = hasRange ? Math.min(fromRaw, toRaw) : null;
  const to = hasRange ? Math.max(fromRaw, toRaw) : null;
  if (!resourceActive(resource, hasRange ? from : null)) return "NA";
  if (hasRange) {
    const unavailable = unavailabilityConflict(resource.id, from, to);
    if (unavailable.length > 0) return "IND";
    const conflicts = allocationConflict(resource.id, from, to);
    if (conflicts.length === 0) return "";
    if (conflicts.length === 1) return "1:1";
    if (conflicts.length === 2) return "1/2";
    return "1/2";
  }
  return "";
}

function formatResourceDisplayLine(resource, options = {}) {
  const {
    role = "",
    weekFrom = null,
    weekTo = null,
    allocation = null,
    noFab = false,
  } = options;
  const name = formatShortPersonName(resource?.name || resource?.resource || "RISORSA");
  const rawRole = String(role || allocation?.role || resource?.role1 || resource?.role2 || "").trim();
  const shownRole = formatShortRole(rawRole);
  const status = formatResourceCompactStatus(resource, rawRole, weekFrom, weekTo);
  const weight = formatAllocationWeightLabel(resource, allocation, weekFrom, weekTo);
  if (resource && isExternalResource(resource)) {
    const extParts = [name];
    if (shownRole) extParts.push(shownRole);
    extParts.push("EXT");
    if (noFab) extParts.push("NO-FABB");
    return extParts.join(" | ");
  }
  const roleStatus = formatRoleStatus(resource, rawRole);
  const parts = [name];
  if (shownRole) parts.push(shownRole);
  if (roleStatus === "NO-MANS") {
    parts.push(roleStatus);
  }
  if (status && (status === "IND" || status === "NA")) {
    parts.push(status);
  }
  const allocState = weight || (status === "1:1" || status === "1/2" ? status : "");
  if (allocState && !parts.includes(allocState)) {
    parts.push(allocState);
  }
  if (noFab) {
    parts.push("NO-FABB");
  }
  return parts.join(" | ");
}

function compactLineTokens(text = "") {
  return String(text || "")
    .split("|")
    .map((part) => String(part || "").trim())
    .filter(Boolean);
}

function appendUniqueCompactCodes(baseLine, codes = []) {
  const tokens = compactLineTokens(baseLine);
  const used = new Set(tokens.map((t) => t.toUpperCase()));
  const extras = [];
  codes.forEach((code) => {
    const token = String(code || "").trim();
    if (!token) return;
    const key = token.toUpperCase();
    if (used.has(key)) return;
    used.add(key);
    extras.push(token);
  });
  return { line: tokens.join(" | "), extras };
}

function resourceWeekAllocationSegments(resource, week) {
  if (!resource) return [];
  if (!resourceActive(resource, week)) return [];
  return state.allocations
    .filter((allocation) =>
      Number(allocation.resource_id) === Number(resource.id) &&
      Number(allocation.week_from) <= Number(week) &&
      Number(allocation.week_to) >= Number(week)
    )
    .sort((a, b) => String(a.project || "").localeCompare(String(b.project || "")) || String(a.role || "").localeCompare(String(b.role || "")))
    .map((allocation) => {
      const roleStatus = formatRoleStatus(resource, allocation.role);
      const weight = formatAllocationWeightLabel(resource, allocation, week, week);
      const parts = [String(allocation.project || "").trim(), formatShortRole(allocation.role || ""), roleStatus];
      if (weight) parts.push(weight);
      return parts.filter(Boolean).join(" | ");
    });
}

function mansioneHoverText(resource, allocation, weekFrom, weekTo) {
  const from = Number(weekFrom || allocation?.week_from || 0);
  const to = Number(weekTo || allocation?.week_to || from || 0);
  const range = from && to ? (from === to ? `W${from}` : `W${from}-W${to}`) : "";
  const projectMeta = plannerProjectMeta(allocation?.project_id, allocation?.project);
  const targetProject = projectMeta.project || allocation?.project || "";
  const targetRole = allocation?.role || "";
  const segmentSet = new Set();
  const allowNoFabTag = from === to;
  let noFab = false;
  let hasActiveWeeks = false;
  for (let week = from; week <= to; week += 1) {
    if (resource && !resourceActive(resource, week)) continue;
    hasActiveWeeks = true;
    resourceWeekAllocationSegments(resource, week).forEach((segment) => segmentSet.add(segment));
    if (allowNoFabTag && demandQtyForPlannerProjectWeek(targetProject, targetRole, week) <= 0) {
      noFab = true;
    }
  }
  if (!hasActiveWeeks) return "";
  const segments = [...segmentSet];
  const name = formatShortPersonName(resource?.name || allocation?.resource || "RISORSA");
  const fallback = formatResourceDisplayLine(resource, {
    role: targetRole,
    weekFrom: from || null,
    weekTo: to || from || null,
    allocation,
  });
  const shortSegments = segments
    .map((segment) =>
      segment
        .split("|")
        .map((part) => String(part || "").trim())
        .filter(Boolean)
        .join("; ")
    )
    .filter(Boolean);
  const visibleSegments = shortSegments.slice(0, 2);
  if (shortSegments.length > 2) {
    visibleSegments.push(`... +${shortSegments.length - 2}`);
  }
  const detail = shortSegments.length ? `${name} | ${visibleSegments.join(" | ")}` : fallback;
  const withNoFab = appendUniqueCompactCodes(detail, noFab ? ["NO-FABB"] : []);
  return `${range ? `${range} | ` : ""}${[withNoFab.line, ...withNoFab.extras].filter(Boolean).join(" | ")}`;
}

function computeCellBadge({ indisp = false, noFab = false, fuoriMansione = false, external = false } = {}) {
  if (indisp) return { code: "IND", label: "INDISPONIBILE", cls: "badge-indisp" };
  if (noFab) return { code: "NO FAB", label: "NO FABBISOGNO", cls: "badge-no-fab" };
  if (fuoriMansione) return { code: "FM", label: "FUORI MANSIONE", cls: "badge-fm" };
  if (external) return { code: "EXT", label: "ESTERNO", cls: "badge-ext" };
  return null;
}

function cellBadgeHtml(badge) {
  if (!badge) return "";
  return `<span class="cell-badge ${badge.cls}" title="${escapeHtml(badge.label)}">!</span>`;
}

function plannerCellBadge(projectId, role, week, required, allocated, projectCode = "") {
  if (Number(allocated || 0) <= 0) return null;
  const targetProject = normalizeProjectCode(projectCode || projectById(projectId)?.code || "");
  const allocations = state.allocations.filter(
    (a) => {
      const meta = plannerProjectMeta(a.project_id, a.project);
      return normalizeProjectCode(meta.project) === targetProject &&
        normalizeRoleKey(a.role) === normalizeRoleKey(role) &&
        Number(a.week_from) <= Number(week) &&
        Number(a.week_to) >= Number(week);
    }
  );
  if (!allocations.length) return null;
  const hasDemandRole = state.demands.some((d) => {
    const meta = plannerProjectMeta(d.project_id, d.project);
    return normalizeProjectCode(meta.project) === targetProject &&
      normalizeRoleKey(d.role) === normalizeRoleKey(role) &&
      Number(d.week) === Number(week) &&
      Number(d.qty || 0) > 0;
  });
  const noFab =
    Number(required || 0) <= 0 ||
    Number(allocated || 0) > Number(required || 0) ||
    !hasDemandRole;
  let fuoriMansione = false;
  let external = false;
  allocations.forEach((allocation) => {
    const resource = state.resources.find((r) => Number(r.id) === Number(allocation.resource_id));
    if (!resource) return;
    if (!roleMatches(resource, role)) fuoriMansione = true;
    if (isExternalResource(resource)) external = true;
  });
  // Sul Planner mostriamo solo badge legati al fabbisogno (NO FAB) e all'origine esterna (EXT).
  return computeCellBadge({ noFab, fuoriMansione: false, external });
}

function plannerCellWarningText(projectId, role, week, projectCode = "") {
  const targetProject = normalizeProjectCode(projectCode || projectById(projectId)?.code || "");
  const allocations = state.allocations.filter(
    (a) => {
      const meta = plannerProjectMeta(a.project_id, a.project);
      const resource = state.resources.find((r) => Number(r.id) === Number(a.resource_id));
      if (!resource || !resourceActive(resource, week)) return false;
      return normalizeProjectCode(meta.project) === targetProject &&
        normalizeRoleKey(a.role) === normalizeRoleKey(role) &&
        Number(a.week_from) <= Number(week) &&
        Number(a.week_to) >= Number(week);
    }
  );
  if (!allocations.length) return "";
  const required = demandQtyForPlannerProjectWeek(targetProject, role, week);
  const byResource = new Map();
  allocations.forEach((allocation) => {
    const resource = state.resources.find((r) => Number(r.id) === Number(allocation.resource_id));
    const key = Number(allocation.resource_id) || String(allocation.resource || "");
    if (!byResource.has(key)) {
      byResource.set(key, {
        resource,
        name: resource?.name || allocation.resource || "RISORSA",
        contribution: 0,
      });
    }
    byResource.get(key).contribution += allocationEffectiveWeightForWeek(allocation, week);
  });
  const entries = [...byResource.values()].sort((a, b) => {
    const contributionDelta = Number(b.contribution || 0) - Number(a.contribution || 0);
    if (Math.abs(contributionDelta) > 0.0001) return contributionDelta;
    return a.name.localeCompare(b.name);
  });
  let covered = 0;
  return entries
    .map((entry) => {
      const before = covered;
      covered += Number(entry.contribution || 0);
      const alloc = allocations.find((a) => Number(a.resource_id) === Number(entry.resource?.id));
      const core = entry.resource
        ? formatResourceDisplayLine(entry.resource, {
            role,
            weekFrom: week,
            weekTo: week,
            allocation: alloc || null,
          })
        : `${formatShortPersonName(entry.name)} | ${formatShortRole(role)} | NO-MANS`;
      const noFab = Number(required || 0) <= 0 || before >= Number(required || 0);
      const packed = appendUniqueCompactCodes(core, noFab ? ["NO-FABB"] : []);
      return [packed.line, ...packed.extras].filter(Boolean).join(" | ");
    })
    .join(" || ");
}

function roleOptionsHtml(selected = "") {
  const roles = [...new Set((state.standard_roles || []).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  return roles
    .map((role) => `<option value="${escapeHtml(role)}" ${role === selected ? "selected" : ""}>${escapeHtml(role)}</option>`)
    .join("");
}

function demandRoleTotals(projectId, weekFrom, weekTo) {
  const totals = new Map();
  state.demands.forEach((d) => {
    if (Number(d.project_id) !== Number(projectId)) return;
    const week = Number(d.week);
    if (week < weekFrom || week > weekTo) return;
    const qty = Number(d.qty || 0);
    if (qty <= 0) return;
    totals.set(d.role, (totals.get(d.role) || 0) + qty);
  });
  return totals;
}

function bestDemandRole(projectId, weekFrom, weekTo) {
  const totals = demandRoleTotals(projectId, weekFrom, weekTo);
  let best = "";
  let bestQty = 0;
  totals.forEach((qty, role) => {
    if (qty > bestQty) {
      bestQty = qty;
      best = role;
    }
  });
  return best;
}

function demandHasRoleInRange(projectId, role, weekFrom, weekTo) {
  const targetRole = normalizeRoleKey(role);
  if (!targetRole) return false;
  const project = projectById(projectId);
  if (project?.code === "OVERALL OFFICINA") {
    for (let week = Number(weekFrom); week <= Number(weekTo); week += 1) {
      if (hasWorkshopChildDemand(targetRole, week)) return true;
    }
    return false;
  }
  return state.demands.some((d) => {
    if (Number(d.project_id) !== Number(projectId)) return false;
    const week = Number(d.week);
    if (week < weekFrom || week > weekTo) return false;
    const qty = Number(d.qty || 0);
    if (qty <= 0) return false;
    return normalizeRoleKey(d.role) === targetRole;
  });
}

function demandHasRoleInWeek(projectId, role, week) {
  return demandHasRoleInRange(projectId, role, week, week);
}

function ganttProjectOptionsHtml(selected = "") {
  const rows = state.projects
    .filter((p) => !p.closed && p.code !== "WS OVERALL")
    .sort((a, b) => projectSortWeek(a.id) - projectSortWeek(b.id) || a.code.localeCompare(b.code));
  const opts = [
    `<option value="INDISP" ${selected === "INDISP" ? "selected" : ""}>INDISP</option>`,
    `<option value="UNASSIGN" ${selected === "UNASSIGN" ? "selected" : ""}>SVINCOLA</option>`,
  ].concat(
    rows.map((p) => `<option value="${escapeHtml(p.code)}" ${p.code === selected ? "selected" : ""}>${escapeHtml(p.code)}</option>`)
  );
  return opts.join("");
}

function toggleGanttRoleInput() {
  if (!ganttProjectSelect || !ganttRoleSelect) return;
  const isIndisp = ganttProjectSelect.value === "INDISP";
  const isUnassign = ganttProjectSelect.value === "UNASSIGN";
  ganttRoleSelect.disabled = isIndisp || isUnassign;
  if (ganttActionDeleteBtn) {
    ganttActionDeleteBtn.hidden = !(ganttActionState?.mode === "unavailability");
  }
  renderGanttConflictBox();
}

function openGanttActionModal({ resource, week_from, week_to, project = "", role = "", mode = "create", unavailability_id = null, allocation_id = null }) {
  if (!ganttActionModal || !resource) return;
  ganttActionState = {
    resource_id: Number(resource.id),
    mode,
    unavailability_id: unavailability_id ? Number(unavailability_id) : null,
    allocation_id: allocation_id ? Number(allocation_id) : null,
  };
  ganttActionInfo.textContent = `${resource.name} | ${resource.role1 || ""}`;
  ganttProjectSelect.innerHTML = ganttProjectOptionsHtml(project || "");
  ganttRoleSelect.innerHTML = roleOptionsHtml(role || resource.role1 || "");
  ganttWeekFromInput.value = week_from;
  ganttWeekToInput.value = week_to;
  toggleGanttRoleInput();
  if (ganttActionCard) {
    ganttActionCard.style.left = "";
    ganttActionCard.style.top = "";
    ganttActionCard.style.margin = "";
  }
  ganttActionModal.hidden = false;
}

function closeGanttActionModal() {
  if (!ganttActionModal) return;
  ganttActionModal.hidden = true;
  ganttActionState = null;
}

function currentGanttConflictEntries() {
  if (!ganttActionState) return [];
  const resourceId = Number(ganttActionState.resource_id);
  const resource = state.resources.find((r) => Number(r.id) === Number(resourceId));
  if (isExternalResource(resource)) return [];
  let weekFrom = parseWeek(ganttWeekFromInput.value);
  let weekTo = parseWeek(ganttWeekToInput.value);
  if (weekFrom === null || weekTo === null) return [];
  if (weekTo < weekFrom) [weekFrom, weekTo] = [weekTo, weekFrom];
  const currentProject = ganttProjectSelect.value || "";
  const currentRole = (ganttRoleSelect?.value || "").trim().toUpperCase();
  const currentAllocationId = Number(ganttActionState.allocation_id || 0);

  const allocs = state.allocations
    .filter((a) => Number(a.resource_id) === resourceId)
    .filter((a) => Number(a.week_from) <= weekTo && Number(a.week_to) >= weekFrom)
    .filter((a) => {
      if (ganttActionState.mode !== "allocation") return true;
      if (Number(a.id) === currentAllocationId) return false;
      return true;
    })
    .map((a) => ({
      kind: "allocation",
      id: Number(a.id),
      project: a.project,
      detail: `${a.project} | ${a.role} | W${a.week_from}-W${a.week_to}`,
      keepThisLabel: "Tieni questa",
      keepOtherLabel: `Tieni ${a.project}`,
    }));

  const unavs = (state.unavailability || [])
    .filter((u) => Number(u.resource_id) === resourceId)
    .filter((u) => Number(u.week_from) <= weekTo && Number(u.week_to) >= weekFrom)
    .filter((u) => {
      if (ganttActionState.mode !== "unavailability") return true;
      return Number(u.id) !== Number(ganttActionState.unavailability_id || 0);
    })
    .map((u) => ({
      kind: "unavailability",
      id: Number(u.id),
      project: "INDISP",
      detail: `INDISP | W${u.week_from}-W${u.week_to}`,
      keepThisLabel: "Tieni questa",
      keepOtherLabel: "Tieni INDISP",
    }));

  if (currentProject === "INDISP") {
    return allocs;
  }
  return [...unavs, ...allocs];
}

async function resolveGanttConflict(action, item) {
  if (!ganttActionState || !item) return;
  let weekFrom = parseWeek(ganttWeekFromInput.value);
  let weekTo = parseWeek(ganttWeekToInput.value);
  if (weekFrom === null || weekTo === null) throw new Error("Compila week da e week a.");
  if (weekTo < weekFrom) [weekFrom, weekTo] = [weekTo, weekFrom];
  const selectedProject = ganttProjectSelect?.value || "";
  const selectedRole = (ganttRoleSelect?.value || "").trim();

  if (action === "remove-other") {
    if (item.kind === "allocation") {
      const before = state.allocations.find((a) => Number(a.id) === Number(item.id));
      if (before) {
        const beforeFrom = Number(before.week_from);
        const beforeTo = Number(before.week_to);
        const overlapFrom = Math.max(beforeFrom, weekFrom);
        const overlapTo = Math.min(beforeTo, weekTo);
        if (overlapTo >= overlapFrom) {
          if (overlapFrom <= beforeFrom && overlapTo >= beforeTo) {
            await api(`/api/allocations/${item.id}`, { method: "DELETE" });
          } else if (overlapFrom <= beforeFrom) {
            await api(`/api/allocations/${item.id}`, {
              method: "PATCH",
              body: JSON.stringify({ week_from: overlapTo + 1 }),
            });
          } else if (overlapTo >= beforeTo) {
            await api(`/api/allocations/${item.id}`, {
              method: "PATCH",
              body: JSON.stringify({ week_to: overlapFrom - 1 }),
            });
          } else {
            await api(`/api/allocations/${item.id}`, {
              method: "PATCH",
              body: JSON.stringify({ week_to: overlapFrom - 1 }),
            });
            await api("/api/allocations", {
              method: "POST",
              body: JSON.stringify({
                project_id: before.project_id,
                project: before.project,
                role: before.role,
                resource_id: before.resource_id,
                week_from: overlapTo + 1,
                week_to: beforeTo,
                note: before.note || "",
              }),
            });
          }
        }
      }
    } else {
      await deleteUnavailable(item.id, { refresh: false });
    }
    if (ganttActionState.mode === "allocation" && selectedProject && selectedProject !== "INDISP") {
      await loadAll(true);
      const resource = state.resources.find((r) => Number(r.id) === Number(ganttActionState.resource_id));
      const project = state.projects.find((p) => p.code === selectedProject);
      if (!resource) throw new Error("Risorsa non trovata.");
      if (!project) throw new Error("Commessa non trovata.");
      const role = selectedRole || resource.role1 || "";
      if (!role) throw new Error("Seleziona una mansione.");
      await saveAssignment(
        {
          project_id: project.id,
          project: project.code,
          role,
          resource_id: resource.id,
          week_from: weekFrom,
          week_to: weekTo,
        },
        { refresh: false }
      );
      await loadAll(true);
      closeGanttActionModal();
      setStatus(`Conflitto risolto: rimossa ${item.project} e salvata ${project.code}.`);
      return;
    }
    if (ganttActionState.mode === "unavailability" || selectedProject === "INDISP") {
      const resource = state.resources.find((r) => Number(r.id) === Number(ganttActionState.resource_id));
      if (!resource) throw new Error("Risorsa non trovata.");
      await saveUnavailable({ resource_id: resource.id, week_from: weekFrom, week_to: weekTo }, { refresh: false });
      await loadAll(true);
      closeGanttActionModal();
      setStatus(`Conflitto risolto: rimossa ${item.project} e salvata INDISP.`);
      return;
    }
    await loadAll(true);
    renderGanttConflictBox();
    setStatus(`Conflitto risolto: rimossa ${item.project}.`);
    return;
  }

  if (action === "remove-current") {
    if (ganttActionState.mode === "unavailability") {
      await deleteUnavailable(ganttActionState.unavailability_id, { refresh: false });
      await loadAll(true);
      closeGanttActionModal();
      setStatus(`Conservata ${item.project}, rimossa INDISP.`);
      return;
    }
    if (ganttActionState.mode === "allocation") {
      closeGanttActionModal();
      setStatus(`Conservata ${item.project}, nuova allocazione annullata.`);
    }
  }
}

function renderGanttConflictBox() {
  if (!ganttConflictBox) return;
  if (!ganttActionState) {
    ganttConflictBox.innerHTML = `<div class="empty-box">Nessuna sovrapposizione nel periodo selezionato.</div>`;
    return;
  }
  const items = currentGanttConflictEntries();
  let selfProjectOverlap = false;
  let currentAllocationForDecision = null;
  if (ganttActionState.mode === "allocation" && Number(ganttActionState.allocation_id || 0) > 0) {
    const currentAllocation = state.allocations.find((a) => Number(a.id) === Number(ganttActionState.allocation_id));
    currentAllocationForDecision = currentAllocation || null;
    let weekFrom = parseWeek(ganttWeekFromInput.value);
    let weekTo = parseWeek(ganttWeekToInput.value);
    if (weekFrom !== null && weekTo !== null) {
      if (weekTo < weekFrom) [weekFrom, weekTo] = [weekTo, weekFrom];
      const selectedProject = ganttProjectSelect?.value || "";
      const selectedRole = (ganttRoleSelect?.value || "").trim();
      const overlapsCurrent =
        !!currentAllocation &&
        Number(currentAllocation.week_from) <= Number(weekTo) &&
        Number(currentAllocation.week_to) >= Number(weekFrom);
      const changingProject =
        !!currentAllocation &&
        String(currentAllocation.project || "") !== String(selectedProject || "");
      const changingRole =
        !!currentAllocation &&
        selectedRole &&
        String(currentAllocation.role || "") !== String(selectedRole || "");
      selfProjectOverlap = overlapsCurrent && (changingProject || changingRole);
    }
  }
  if (!items.length) {
    if (!selfProjectOverlap) {
      ganttConflictBox.innerHTML = `<div class="empty-box">Nessuna sovrapposizione nel periodo selezionato.</div>`;
      return;
    }
    const selectedChoice = String(ganttActionState.overlap_choice || "");
    const chosenLabel =
      selectedChoice === "keep"
        ? "Scelta: lascia su OVERALL."
        : selectedChoice === "move"
        ? "Scelta: sposta su nuova commessa."
        : selectedChoice === "split"
        ? "Scelta: split 0,5 + 0,5 (solo UI)."
        : "";
    ganttConflictBox.innerHTML = `
      <div class="selection-sub" style="margin-bottom:8px;font-weight:700;">âš ï¸ Risorsa gia allocata nella settimana selezionata</div>
      <div class="inline-actions">
        <button class="cell-btn" data-gantt-overlap-choice="keep">1) Lascia</button>
        <button class="cell-btn" data-gantt-overlap-choice="move">2) Sposta</button>
        <button class="cell-btn" data-gantt-overlap-choice="split">3) Split 0,5 + 0,5</button>
      </div>
      ${chosenLabel ? `<div class="selection-sub" style="margin-top:8px;">${escapeHtml(chosenLabel)}</div>` : ""}
    `;
    ganttConflictBox.querySelectorAll("[data-gantt-overlap-choice]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const choice = btn.dataset.ganttOverlapChoice || "";
        ganttActionState.overlap_choice = choice;
        if (choice === "keep" && currentAllocationForDecision) {
          if (ganttProjectSelect) ganttProjectSelect.value = String(currentAllocationForDecision.project || "");
          if (ganttRoleSelect) ganttRoleSelect.value = String(currentAllocationForDecision.role || "");
        }
        renderGanttConflictBox();
      });
    });
    return;
  }
  ganttConflictBox.innerHTML = items
    .map(
      (item) => `
        <div class="selection-item">
          <div>
            <strong>Conflitto con ${escapeHtml(item.project)}</strong>
            <div class="selection-sub">${escapeHtml(item.detail)}</div>
          </div>
          <div class="inline-actions">
            <button class="cell-btn" data-gantt-conflict-remove-other="${item.kind}:${item.id}" data-gantt-conflict-project="${escapeHtml(item.project)}">${escapeHtml(item.keepThisLabel)}</button>
            <button class="cell-btn danger" data-gantt-conflict-remove-current="${item.kind}:${item.id}" data-gantt-conflict-project="${escapeHtml(item.project)}">${escapeHtml(item.keepOtherLabel)}</button>
          </div>
        </div>
      `
    )
    .join("");
  ganttConflictBox.querySelectorAll("[data-gantt-conflict-remove-other]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const [kind, id] = btn.dataset.ganttConflictRemoveOther.split(":");
      resolveGanttConflict("remove-other", { kind, id: Number(id), project: btn.dataset.ganttConflictProject || (kind === "unavailability" ? "INDISP" : "") }).catch((e) => setStatus(e.message));
    });
  });
  ganttConflictBox.querySelectorAll("[data-gantt-conflict-remove-current]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const [kind, id] = btn.dataset.ganttConflictRemoveCurrent.split(":");
      resolveGanttConflict("remove-current", { kind, id: Number(id), project: btn.dataset.ganttConflictProject || (kind === "unavailability" ? "INDISP" : "") }).catch((e) => setStatus(e.message));
    });
  });
}

async function saveGanttActionModal() {
  if (!ganttActionState) return;
  const resource = state.resources.find((r) => Number(r.id) === Number(ganttActionState.resource_id));
  if (!resource) throw new Error("Risorsa non trovata.");
  let week_from = parseWeek(ganttWeekFromInput.value);
  let week_to = parseWeek(ganttWeekToInput.value);
  if (week_from === null || week_to === null) throw new Error("Compila week da e week a.");
  if (week_to < week_from) [week_from, week_to] = [week_to, week_from];
  let splitUiChoice = false;
  if (ganttActionState.mode === "allocation" && Number(ganttActionState.allocation_id || 0) > 0) {
    const currentAllocation = state.allocations.find((a) => Number(a.id) === Number(ganttActionState.allocation_id));
    const selectedProject = ganttProjectSelect?.value || "";
    const overlapsCurrent =
      !!currentAllocation &&
      Number(currentAllocation.week_from) <= Number(week_to) &&
      Number(currentAllocation.week_to) >= Number(week_from);
    const changingProject =
      !!currentAllocation &&
      String(currentAllocation.project || "") !== String(selectedProject || "");
    if (overlapsCurrent && changingProject) {
      const choice = String(ganttActionState.overlap_choice || "");
      if (choice === "keep") {
        closeGanttActionModal();
        setStatus("Allocazione lasciata invariata.");
        return;
      }
      if (choice === "split") splitUiChoice = true;
    }
  }
  const pendingConflicts = currentGanttConflictEntries();
  if (pendingConflicts.length > 0) {
    throw new Error("Risolvi prima la sovrapposizione: tieni la nuova allocazione oppure conserva quella esistente.");
  }
  if (ganttProjectSelect.value === "INDISP") {
    await saveUnavailable({ resource_id: resource.id, week_from, week_to }, { refresh: false });
  } else if (ganttProjectSelect.value === "UNASSIGN") {
    await api("/api/allocations/unassign", {
      method: "POST",
      body: JSON.stringify({
        resource_id: resource.id,
        week_from,
        week_to,
      }),
    });
    logUserAction("Svincolo", `${resource.name} | W${week_from}-W${week_to}`);
    await loadAll(true);
    closeGanttActionModal();
    setStatus("Allocazione svincolata.");
    return;
  } else {
      const projectCode = ganttProjectSelect.value;
      const project = state.projects.find((p) => normalizeProjectCode(p.code) === normalizeProjectCode(projectCode));
      if (!project) throw new Error("Commessa non trovata.");
      const owner = getOwningOverall(project.code);
      if (!isAllocableProjectCode(project.code)) {
        const ownerCode = String(owner?.code || owner?.id || owner?.name || "OVERALL").trim();
        alert(`La commessa ${project.code} è gestita da ${ownerCode}; alloca su ${ownerCode}`);
        return;
      }
      let role = ganttRoleSelect.value.trim() || resource.role1 || "";
      if (!role) throw new Error("Seleziona una mansione.");
      if (ganttActionState.mode === "allocation" && ganttActionState.allocation_id) {
        validateInternalAllocationRange(resource.id, week_from, week_to, Number(ganttActionState.allocation_id));
        await api(`/api/allocations/${ganttActionState.allocation_id}`, {
          method: "PATCH",
          body: JSON.stringify({
            project_id: project.id,
            resource_id: resource.id,
            role,
            week_from,
            week_to,
            weight: 1,
          }),
        });
        logUserAction("Allocazione aggiornata", `${resource.name} -> ${project.code} | ${role} | W${week_from}-W${week_to}`);
    } else {
      validateInternalAllocationRange(resource.id, week_from, week_to);
      await saveAssignment(
        {
          project_id: project.id,
          project: project.code,
          role,
          resource_id: resource.id,
          week_from,
          week_to,
        },
        { refresh: false }
      );
    }
  }
  await loadAll(true);
  closeGanttActionModal();
  setStatus(
    splitUiChoice
      ? "Azione Gantt salvata. Split 0,5+0,5 registrato solo a livello UI."
      : "Azione Gantt salvata."
  );
}

async function deleteUnavailable(id, { refresh = true } = {}) {
  const before = (state.unavailability || []).find((u) => Number(u.id) === Number(id));
  await api(`/api/unavailability/${id}`, { method: "DELETE" });
  if (before) {
    pushUndo({
      label: `INDISP ${before.resource || ""}`.trim(),
      undoMessage: "Ultima INDISP annullata.",
      undo: async () => {
        await api("/api/unavailability", {
          method: "POST",
          body: JSON.stringify({
            resource_id: before.resource_id,
            week_from: before.week_from,
            week_to: before.week_to,
            reason: before.reason || "INDISP",
          }),
        });
      },
    });
  }
  if (refresh) await loadAll(true);
}

async function deleteGanttUnavailable() {
  const id = Number(ganttActionState?.unavailability_id || 0);
  if (!id) throw new Error("Periodo INDISP non trovato.");
  logUserAction("INDISP rimossa", `ID ${id}`);
  await deleteUnavailable(id, { refresh: false });
  await loadAll(true);
  closeGanttActionModal();
  setStatus("INDISP rimossa.");
}

function allocationConflict(resourceId, weekFrom, weekTo, ignoreAllocationId = null) {
  const resource = state.resources.find((r) => Number(r.id) === Number(resourceId));
  if (isExternalResource(resource)) return [];
  return state.allocations.filter((a) => {
    if (Number(a.resource_id) !== Number(resourceId)) return false;
    if (ignoreAllocationId && Number(a.id) === Number(ignoreAllocationId)) return false;
    return Number(a.week_from) <= Number(weekTo) && Number(a.week_to) >= Number(weekFrom);
  });
}

function unavailabilityConflict(resourceId, weekFrom, weekTo, ignoreId = null) {
  return (state.unavailability || []).filter((u) => {
    if (Number(u.resource_id) !== Number(resourceId)) return false;
    if (ignoreId && Number(u.id) === Number(ignoreId)) return false;
    return Number(u.week_from) <= Number(weekTo) && Number(u.week_to) >= Number(weekFrom);
  });
}

function fillRoleList() {
  const roles = new Set(state.standard_roles || []);
  state.resources.forEach((r) => {
    if (r.role1) roles.add(r.role1);
    if (r.role2) roles.add(r.role2);
  });
  state.demands.forEach((d) => {
    if (d.role) roles.add(d.role);
  });
  const options = [...roles].filter(Boolean).sort((a, b) => a.localeCompare(b));
  roleList.innerHTML = options.map((r) => `<option value="${escapeHtml(r)}"></option>`).join("");
  plannerRoleFilter.innerHTML = `<option value="">Tutte le mansioni</option>${options
    .map((r) => `<option value="${escapeHtml(r)}">${escapeHtml(r)}</option>`)
    .join("")}`;
}

function fillProjectFilter() {
  const projects = [...new Set([
    ...state.projects.map((p) => p.code),
    ...Object.values(state.overalls || {}).map((o) => o.name),
  ].filter(Boolean))].sort((a, b) => {
    const pa = projectSortWeekByCode(a);
    const pb = projectSortWeekByCode(b);
    return pa - pb || a.localeCompare(b);
  });
  plannerProjectFilter.innerHTML = `<option value="">Tutte le commesse</option>${projects
    .map((p) => `<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`)
    .join("")}`;
}

function projectSortWeek(projectId) {
  const current = getCurrentWeek();
  const weeks = [];
  state.demands.forEach((d) => {
    if (Number(d.project_id) === Number(projectId) && Number(d.qty || 0) > 0 && Number(d.week) >= current) {
      weeks.push(Number(d.week));
    }
  });
  state.allocations.forEach((a) => {
    if (Number(a.project_id) === Number(projectId) && Number(a.week_to) >= current) {
      weeks.push(Math.max(Number(a.week_from), current));
    }
  });
  return weeks.length ? Math.min(...weeks) : 999;
}

function projectSortWeekByCode(projectCode) {
  const code = normalizeProjectCode(projectCode);
  const project = state.projects.find((p) => normalizeProjectCode(p.code) === code);
  if (project) return projectSortWeek(project.id);
  const overall = Object.values(state.overalls || {}).find((o) =>
    normalizeProjectCode(o?.name) === code || normalizeProjectCode(o?.id) === code
  );
  if (!overall) return 999;
  const childIds = state.projects
    .filter((p) => (overall.project_codes || []).some((c) => normalizeProjectCode(c) === normalizeProjectCode(p.code)))
    .map((p) => Number(p.id))
    .filter((id) => Number.isFinite(id) && id > 0);
  if (!childIds.length) return 999;
  return Math.min(...childIds.map((id) => projectSortWeek(id)));
}

function renderSummary() {
  const rows = plannerRows();
  const diffMap = state.lastDemandDiffMap || {};
  const diffCurrentWeek = state.lastDemandDiffMeta?.current_week ? Number(state.lastDemandDiffMeta.current_week) : null;
  const weekly = plannerWeeklyTotals(rows);
  const weeksCount = Math.max(weekly.length, 1);
  const total = weekly.reduce(
    (acc, row) => {
      acc.required += Number(row.required || 0);
      acc.allocated += Number(row.allocated || 0);
      acc.residual += Number(row.residual || 0);
      return acc;
    },
    { required: 0, allocated: 0, residual: 0 }
  );
  const peakDemand = weekly.reduce((max, row) => Math.max(max, Number(row.required || 0)), 0);
  const peakResidual = weekly.reduce((max, row) => Math.max(max, Number(row.residual || 0)), 0);
  const avgDemand = total.required / weeksCount;
  const avgAllocated = total.allocated / weeksCount;
  const avgResidual = total.residual / weeksCount;
  const formatSummaryMetric = (value) => {
    const rounded = Math.round(Number(value || 0) * 10) / 10;
    return Number.isInteger(rounded)
      ? rounded.toLocaleString("it-IT")
      : rounded.toLocaleString("it-IT", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  };
  const setSummaryCard = (valueEl, label, value, title = "") => {
    if (!valueEl) return;
    const labelEl = valueEl.previousElementSibling;
    if (labelEl) labelEl.textContent = label;
    valueEl.textContent = formatSummaryMetric(value);
    const card = valueEl.closest(".summary-card");
    if (card) card.title = title;
  };
  setSummaryCard(
    summaryDemand,
    "Fabbisogno medio / sett.",
    avgDemand,
    `Vista corrente. Picco settimanale: ${peakDemand.toLocaleString("it-IT")}`
  );
  setSummaryCard(
    summaryAllocated,
    "Allocate medie / sett.",
    avgAllocated,
    `Vista corrente. Picco allocato: ${weekly.reduce((max, row) => Math.max(max, Number(row.allocated || 0)), 0).toLocaleString("it-IT")}`
  );
  setSummaryCard(
    summaryResidual,
    "Residuo medio / sett.",
    avgResidual,
    `Vista corrente. Picco residuo: ${peakResidual.toLocaleString("it-IT")}`
  );
  setSummaryCard(
    summaryResources,
    "Risorse attive",
    state.resources.filter((r) => r.status !== "CESSATO" && !isExternalResource(r)).length
  );
  demandSourcePath.textContent = currentDemandPathValue();
}

function computePlannerRows(projectFilter = "", roleFilter = "", includeZeroDemandProjects = false, updateHiddenNoticeFlag = false) {
  const demandMap = new Map();
  const allocationsByKey = new Map();
  const combos = new Map();
  const firstVisibleWeek = getCurrentWeek();
  const activeProjects = new Set();

  state.demands.forEach((d) => {
    const project = projectById(d.project_id);
    if (project?.code === "WS OVERALL") return;
    const meta = plannerProjectMeta(d.project_id, d.project);
    const key = `${meta.project}|${d.role}`;
    if (!combos.has(key)) {
      combos.set(key, { project_id: meta.project_id, project: meta.project, project_key: meta.project, role: d.role });
    }
    const qty = Number(d.qty || 0);
    demandMap.set(`${key}|${d.week}`, qty);
    if (Number(d.week) >= firstVisibleWeek && qty > 0) {
      activeProjects.add(meta.project);
    }
  });

  state.allocations.forEach((a) => {
    const project = projectById(a.project_id);
    if (project?.code === "WS OVERALL") return;
    const meta = plannerProjectMeta(a.project_id, a.project);
    const key = `${meta.project}|${a.role}`;
    if (!combos.has(key)) {
      combos.set(key, { project_id: meta.project_id, project: meta.project, project_key: meta.project, role: a.role });
    }
    for (let week = Number(a.week_from); week <= Number(a.week_to); week += 1) {
      const wk = `${key}|${week}`;
      if (!allocationsByKey.has(wk)) allocationsByKey.set(wk, []);
      allocationsByKey.get(wk).push(a);
      if (week >= firstVisibleWeek) activeProjects.add(meta.project);
    }
  });

  let rows = [...combos.values()].map((combo) => {
    const weeks = Array.from({ length: 52 }, (_, idx) => {
      const week = idx + 1;
      const req = Number(demandMap.get(`${combo.project_key}|${combo.role}|${week}`) || 0);
      const allocsRaw = allocationsByKey.get(`${combo.project_key}|${combo.role}|${week}`) || [];
      const allocs = allocsRaw.filter((a) => {
        const resource = state.resources.find((r) => Number(r.id) === Number(a.resource_id));
        return resource && resourceActive(resource, week);
      });
      const internalAllocated = allocs.reduce((acc, allocation) => {
        const resource = state.resources.find((r) => Number(r.id) === Number(allocation.resource_id));
        if (!resource || isExternalResource(resource)) return acc;
        return acc + allocationEffectiveWeightForWeek(allocation, week);
      }, 0);
      const externalAllocated = allocs.reduce((acc, allocation) => {
        const resource = state.resources.find((r) => Number(r.id) === Number(allocation.resource_id));
        if (!resource || !isExternalResource(resource)) return acc;
        return acc + 1;
      }, 0);
      const allocated = internalAllocated + externalAllocated;
      return {
        week,
        required: req,
        internalAllocated,
        externalAllocated,
        allocated,
        allocatedDisplay: formatAllocatedDisplay(internalAllocated, externalAllocated),
        residual: Math.max(req - allocated, 0),
        names: allocs.map((a) => {
          const resource = state.resources.find((r) => Number(r.id) === Number(a.resource_id));
          if (!resource) return `${formatShortPersonName(a.resource)} | ${formatShortRole(combo.role)} | 1:1`;
          return formatResourceDisplayLine(resource, {
            role: combo.role,
            weekFrom: week,
            weekTo: week,
            allocation: a,
          });
        }),
      };
    });
    return {
      ...combo,
      totalRequired: weeks.filter((w) => w.week >= firstVisibleWeek).reduce((n, w) => n + w.required, 0),
      totalInternalAllocated: weeks.filter((w) => w.week >= firstVisibleWeek).reduce((n, w) => n + Number(w.internalAllocated || 0), 0),
      totalExternalAllocated: weeks.filter((w) => w.week >= firstVisibleWeek).reduce((n, w) => n + Number(w.externalAllocated || 0), 0),
      totalAllocated: weeks.filter((w) => w.week >= firstVisibleWeek).reduce((n, w) => n + w.allocated, 0),
      totalResidual: weeks.filter((w) => w.week >= firstVisibleWeek).reduce((n, w) => n + w.residual, 0),
      weeks,
    };
  });

  const pf = (projectFilter || "").trim();
  const rf = (roleFilter || "").trim();
  rows = rows.filter((row) => (!pf || row.project === pf) && (!rf || row.role === rf));
  const hiddenProjects = [...new Set(rows.filter((row) => !activeProjects.has(row.project)).map((row) => row.project))];
  if (!includeZeroDemandProjects) {
    rows = rows.filter((row) => activeProjects.has(row.project));
  }
  if (updateHiddenNoticeFlag) updatePlannerHiddenNotice(hiddenProjects.length);
  rows.sort((a, b) => {
    const pa = projectSortWeekByCode(a.project);
    const pb = projectSortWeekByCode(b.project);
    return pa - pb || a.project.localeCompare(b.project) || a.role.localeCompare(b.role);
  });
  return rows;
}

function plannerRows() {
  return computePlannerRows(
    plannerProjectFilter.value.trim(),
    plannerRoleFilter.value.trim(),
    !!showZeroDemandProjects.checked,
    true
  );
}

function computeDemandOnlyRows(projectFilter = "", roleFilter = "", includeZeroDemandProjects = true) {
  const demandMap = new Map();
  const combos = new Map();
  const activeProjects = new Set();
  const firstVisibleWeek = getCurrentWeek();

  state.demands.forEach((d) => {
    const project = projectById(d.project_id);
    if (project?.code === "WS OVERALL") return;
    const meta = plannerProjectMeta(d.project_id, d.project);
    const key = `${meta.project}|${d.role}`;
    if (!combos.has(key)) combos.set(key, { project_id: meta.project_id, project: meta.project, project_key: meta.project, role: d.role });
    const qty = Number(d.qty || 0);
    demandMap.set(`${key}|${d.week}`, qty);
    if (Number(d.week) >= firstVisibleWeek && qty > 0) activeProjects.add(meta.project);
  });

  let rows = [...combos.values()].map((combo) => {
    const weeks = Array.from({ length: 52 }, (_, idx) => {
      const week = idx + 1;
      return { week, required: Number(demandMap.get(`${combo.project_key}|${combo.role}|${week}`) || 0) };
    });
    return {
      ...combo,
      weeks,
      totalRequired: weeks.filter((w) => w.week >= firstVisibleWeek).reduce((n, w) => n + w.required, 0),
    };
  });

  const pf = (projectFilter || "").trim();
  const rf = (roleFilter || "").trim();
  rows = rows.filter((row) => (!pf || row.project === pf) && (!rf || row.role === rf));
  if (!includeZeroDemandProjects) rows = rows.filter((row) => activeProjects.has(row.project));
  rows = rows.filter((row) => row.totalRequired > 0 || includeZeroDemandProjects);
  rows.sort((a, b) => projectSortWeekByCode(a.project) - projectSortWeekByCode(b.project) || a.project.localeCompare(b.project) || a.role.localeCompare(b.role));
  return rows;
}

function updatePlannerHiddenNotice(hiddenCount) {
  if (!plannerHiddenNotice) return;
  if (!hiddenCount) {
    plannerHiddenNotice.hidden = true;
    plannerHiddenNotice.textContent = "";
    return;
  }
  plannerHiddenNotice.hidden = false;
  plannerHiddenNotice.textContent = `${hiddenCount} commesse senza fabbisogno futuro sono nascoste. Spunta il filtro per vederle.`;
}

function plannerWeeklyTotals(rows) {
  const totals = new Map();
  rows.forEach((row) => {
    row.weeks
      .filter((w) => w.week >= getCurrentWeek())
      .forEach((w) => {
        if (!totals.has(w.week)) {
          totals.set(w.week, { week: w.week, required: 0, internalAllocated: 0, externalAllocated: 0, allocated: 0, residual: 0 });
        }
        const item = totals.get(w.week);
        item.required += Number(w.required || 0);
        item.internalAllocated += Number(w.internalAllocated || 0);
        item.externalAllocated += Number(w.externalAllocated || 0);
        item.allocated += Number(w.allocated || 0);
      });
  });
  return visibleWeeks().map((week) => {
    const item = totals.get(week) || { week, required: 0, internalAllocated: 0, externalAllocated: 0, allocated: 0, residual: 0 };
    return {
      ...item,
      allocatedDisplay: formatAllocatedDisplay(item.internalAllocated, item.externalAllocated),
      residual: Math.max(Number(item.required || 0) - Number(item.allocated || 0), 0),
    };
  });
}

function renderPlannerHeader(rows = plannerRows()) {
  const weeks = visibleWeeks();
  const totals = plannerWeeklyTotals(rows);
  plannerHead.innerHTML = `
    <tr class="planner-subtotals-row">
      <th class="sticky-col planner-subtotals-label">Subtotali</th>
      <th class="sticky-col-2 planner-subtotals-label">Vista filtrata</th>
      <th class="sticky-col-3 planner-subtotals-label">Settimana</th>
      ${totals
        .map(
          (w) => `
            <th class="week-head week-subtotal-head" title="${weekLabelTooltip(w.week)}">
              <span class="metric">R ${w.required}</span>
              <span class="metric">A ${w.allocatedDisplay || formatAllocatedDisplay(w.allocated, 0)}</span>
              <span class="metric">D ${w.residual}</span>
            </th>
          `
        )
        .join("")}
    </tr>
    <tr>
      <th class="sticky-col col-project">Commessa</th>
      <th class="sticky-col-2 col-role">Mansione</th>
      <th class="sticky-col-3 col-total">Scopertura</th>
      ${weeks.map((w) => `<th class="week-head" title="${weekLabelTooltip(w)}">W${w}</th>`).join("")}
    </tr>
  `;
}

function fillGanttDemandFilters() {
  if (!ganttDemandProjectFilter || !ganttDemandRoleFilter) return;
  const currentProject = ganttDemandProjectFilter.value;
  const currentRole = ganttDemandRoleFilter.value;
  const rows = computeDemandOnlyRows("", "", true);
  const projects = [...new Set(rows.map((r) => r.project))];
  const roles = [...new Set(rows.map((r) => r.role))];
  ganttDemandProjectFilter.innerHTML = `<option value="">Tutte le commesse</option>${projects
    .map((p) => `<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`)
    .join("")}`;
  ganttDemandRoleFilter.innerHTML = `<option value="">Tutte le mansioni</option>${roles
    .map((r) => `<option value="${escapeHtml(r)}">${escapeHtml(r)}</option>`)
    .join("")}`;
  ganttDemandProjectFilter.value = projects.includes(currentProject) ? currentProject : "";
  ganttDemandRoleFilter.value = roles.includes(currentRole) ? currentRole : "";
}

function renderGanttDemandMatrix() {
  if (!ganttDemandMatrixHead || !ganttDemandMatrixBody) return;
  const diffMap = state.lastDemandDiffMap || {};
  const diffCurrentWeek = state.lastDemandDiffMeta?.current_week ? Number(state.lastDemandDiffMeta.current_week) : null;
  const rows = computePlannerRows(
    ganttDemandProjectFilter?.value || "",
    ganttDemandRoleFilter?.value || "",
    false,
    false
  );
  const weeks = visibleWeeks();
  ganttDemandMatrixHead.innerHTML = `
    <tr class="planner-subtotals-row">
      <th class="gantt-demand-col-project planner-subtotals-label">Subtotali</th>
      <th class="gantt-demand-col-role planner-subtotals-label">Vista filtrata</th>
      ${plannerWeeklyTotals(rows)
        .map(
          (w) => `
            <th class="week-head week-subtotal-head" title="${weekLabelTooltip(w.week)}">
              <span class="metric">R ${w.required}</span>
              <span class="metric">A ${w.allocatedDisplay || formatAllocatedDisplay(w.allocated, 0)}</span>
              <span class="metric">D ${w.residual}</span>
            </th>
          `
        )
        .join("")}
    </tr>
    <tr>
      <th class="gantt-demand-col-project">Commessa</th>
      <th class="gantt-demand-col-role">Mansione</th>
      ${weeks.map((w) => `<th class="week-head" title="${weekLabelTooltip(w)}">W${w}</th>`).join("")}
    </tr>
  `;
  ganttDemandMatrixBody.innerHTML = rows.length
    ? rows
        .map(
          (row) => `
            <tr>
              <td class="gantt-demand-col-project project-cell">${escapeHtml(row.project)}</td>
              <td class="gantt-demand-col-role role-cell">${escapeHtml(row.role)}</td>
              ${weeks
                .map((week) => {
                  const cell = row.weeks.find((w) => Number(w.week) === Number(week));
                  const req = Number(cell?.required || 0);
                  const alloc = Number(cell?.allocated || 0);
                  const residual = Number(cell?.residual || 0);
                  const cls =
                    residual > 0
                      ? "cell-gap"
                      : req > 0 || alloc > 0
                      ? "cell-covered"
                      : "cell-empty";
                  const badge = plannerCellBadge(row.project_id, row.role, week, req, alloc, row.project);
                  const noFabClass = badge?.cls === "badge-no-fab" ? "cell-warning-nofab" : "";
                  const diffKey = `${row.project}||${row.role}||${week}`;
                  const diff = diffMap[diffKey];
                  const diffOk = diff && (!diffCurrentWeek || Number(week) >= diffCurrentWeek);
                  const diffClass = diffOk ? "cell-diff" : "";
                  const diffTitle = diffOk ? ` Delta fabb ${diff.prev_qty}->${diff.new_qty}` : "";
                  const warningText = plannerCellWarningText(row.project_id, row.role, week, row.project);
                  const nameText = cell?.names?.length ? cell.names.join(", ") : "";
                  const baseTitle = warningText || nameText;
                  const title = baseTitle
                    ? ` title="${escapeHtml(baseTitle)}${diffTitle}"`
                    : diffTitle
                    ? ` title="${escapeHtml(diffTitle.trim())}"`
                    : "";
                  return `
                    <td>
                      <button
                        class="planner-cell ${cls} ${diffClass} ${noFabClass}"
                        data-project-id="${row.project_id}"
                        data-project="${escapeHtml(row.project)}"
                        data-role="${escapeHtml(row.role)}"
                        data-week="${week}"
                        ${title}
                      >
                        ${cellBadgeHtml(badge)}
                        <span class="metric">R ${req}</span>
                        <span class="metric">A ${cell?.allocatedDisplay || formatAllocatedDisplay(alloc, 0)}</span>
                        <span class="metric">D ${residual}</span>
                      </button>
                    </td>`;
                })
                .join("")}
            </tr>
          `
        )
        .join("")
    : `<tr><td class="gantt-demand-col-project empty-box">Nessun fabbisogno trovato.</td><td class="gantt-demand-col-role"></td>${weeks.map(() => "<td></td>").join("")}</tr>`;

  ganttDemandMatrixBody.querySelectorAll(".planner-cell").forEach((btn) => {
    btn.addEventListener("click", () => {
      const target = {
        project_id: parseIntOr(btn.dataset.projectId, 0),
        project: btn.dataset.project,
        role: btn.dataset.role,
        week: parseIntOr(btn.dataset.week, 0),
        week_from: parseIntOr(btn.dataset.week, 0),
        week_to: parseIntOr(btn.dataset.week, 0),
      };
      selectTarget(target);
    });
  });
}

function openPlannerPrintModal() {
  if (!plannerPrintModal) return;
  plannerPrintWeekFrom.value = String(getCurrentWeek());
  plannerPrintWeekTo.value = "52";
  const isOverallWorkshop = (plannerProjectFilter?.value || "") === "OVERALL OFFICINA";
  if (plannerPrintWorkshopWrap) plannerPrintWorkshopWrap.hidden = !isOverallWorkshop;
  if (plannerPrintIncludeWorkshop) plannerPrintIncludeWorkshop.checked = false;
  plannerPrintModal.hidden = false;
}

function closePlannerPrintModal() {
  if (plannerPrintModal) plannerPrintModal.hidden = true;
}

function buildWorkshopPrintRows(weekFrom, weekTo) {
  const demandMap = new Map();
  const combos = new Map();
  state.demands.forEach((d) => {
    const project = projectById(d.project_id);
    if (!isWorkshopChildProject(project)) return;
    const key = `${project.code}|${d.role}|${d.week}`;
    demandMap.set(key, Number(d.qty || 0));
    const comboKey = `${project.code}|${d.role}`;
    if (!combos.has(comboKey)) combos.set(comboKey, { project: project.code, role: d.role });
  });
  return [...combos.values()]
    .map((combo) => {
      const weeks = [];
      let total = 0;
      for (let week = weekFrom; week <= weekTo; week += 1) {
        const required = Number(demandMap.get(`${combo.project}|${combo.role}|${week}`) || 0);
        weeks.push({ week, required });
        total += required;
      }
      return { ...combo, weeks, total };
    })
    .filter((row) => row.total > 0)
    .sort((a, b) => a.project.localeCompare(b.project) || a.role.localeCompare(b.role));
}

function runPlannerPrint() {
  let weekFrom = parseWeek(plannerPrintWeekFrom.value);
  let weekTo = parseWeek(plannerPrintWeekTo.value);
  if (weekFrom === null || weekTo === null) {
    setStatus("Compila week da e week a per la stampa.");
    return;
  }
  if (weekTo < weekFrom) [weekFrom, weekTo] = [weekTo, weekFrom];
  const weeks = visibleWeeks();
  const weekSet = new Set(weeks.filter((w) => w >= weekFrom && w <= weekTo));
  plannerPrintState = [];
  document.querySelectorAll("#plannerMatrix tr").forEach((tr) => {
    [...tr.children].forEach((cell, idx) => {
      if (idx < 3) return;
      const week = weeks[idx - 3];
      if (!weekSet.has(week)) {
        plannerPrintState.push({ el: cell, display: cell.style.display });
        cell.style.display = "none";
      }
    });
  });
  const includeWorkshop = !!plannerPrintIncludeWorkshop?.checked && (plannerProjectFilter?.value || "") === "OVERALL OFFICINA";
  if (includeWorkshop) {
    const host = document.createElement("div");
    host.id = "plannerPrintWorkshopHost";
    const rows = buildWorkshopPrintRows(weekFrom, weekTo);
    host.innerHTML = `
      <section class="card print-extra-card">
        <h2>Dettaglio Sottocommesse Officina</h2>
        <div class="table-wrap">
          <table class="print-workshop-table">
            <thead>
              <tr>
                <th>Commessa</th>
                <th>Mansione</th>
                ${Array.from({ length: weekTo - weekFrom + 1 }, (_, i) => `<th>W${weekFrom + i}</th>`).join("")}
              </tr>
            </thead>
            <tbody>
              ${rows
                .map(
                  (row) => `
                    <tr>
                      <td>${escapeHtml(row.project)}</td>
                      <td>${escapeHtml(row.role)}</td>
                      ${row.weeks.map((w) => `<td>${w.required}</td>`).join("")}
                    </tr>`
                )
                .join("")}
            </tbody>
          </table>
        </div>
      </section>`;
    document.getElementById("sheet-planner")?.appendChild(host);
    plannerPrintState.push({ extraHost: host });
  }
  closePlannerPrintModal();
  window.setTimeout(() => window.print(), 40);
}

window.addEventListener("afterprint", () => {
  if (!plannerPrintState) return;
  plannerPrintState.forEach((item) => {
    if (item.el) item.el.style.display = item.display;
    if (item.extraHost) item.extraHost.remove();
  });
  plannerPrintState = null;
});

function selectTarget(target, resourceId = null) {
  selectedTarget = { ...target };
  const plannerGrid = plannerSideCard?.closest(".planner-grid");
  if (plannerGrid?.classList.contains("side-collapsed")) {
    plannerGrid.classList.remove("side-collapsed");
    if (plannerSideToggle) {
      plannerSideToggle.title = "Nascondi pannello laterale";
      plannerSideToggle.setAttribute("aria-label", plannerSideToggle.title);
    }
  }
  plannerGrid?.classList.add("side-active");
  console.log("selectedTarget", selectedTarget);
  console.log("week", selectedTarget?.week);
  const baseWeek = target.week_from ?? target.week ?? "";
  const demandQtyForTarget = state.demands.reduce((acc, d) => {
    const meta = plannerProjectMeta(d.project_id, d.project);
    if (normalizeProjectCode(meta.project) !== normalizeProjectCode(target.project)) return acc;
    if (normalizeRoleKey(d.role) !== normalizeRoleKey(target.role)) return acc;
    if (Number(d.week) !== Number(baseWeek)) return acc;
    return acc + Number(d.qty || 0);
  }, 0);
  assignProject.value = target.project || "";
  assignRole.value = target.role || "";
  assignWeekFrom.value = baseWeek;
  assignWeekTo.value = target.week_to ?? target.week ?? "";
  demandWeekFrom.value = baseWeek;
  demandWeekTo.value = target.week_to ?? target.week ?? "";
  demandQty.value = String(demandQtyForTarget);
  selectionInfo.textContent = `${target.project} | ${target.role} | W${target.week_from ?? target.week} - W${target.week_to ?? target.week}`;
  try {
    renderAssignResourceOptions(resourceId);
    renderSelectionAllocations();
    renderResourcePool();
  } catch (err) {
    console.error("selectTarget panel render error", err);
    setStatus(`Errore pannello laterale: ${err.message || err}`);
  }
  window.requestAnimationFrame(() => {
    try {
      renderAssignResourceOptions(resourceId);
    } catch (err) {
      console.error("selectTarget delayed options error", err);
    }
  });
}

function renderPlannerMatrix() {
  plannerBody.innerHTML = "";
  const rows = plannerRows();
  const diffMap = state.lastDemandDiffMap || {};
  const diffCurrentWeek = state.lastDemandDiffMeta?.current_week ? Number(state.lastDemandDiffMeta.current_week) : null;
  renderPlannerHeader(rows);
  rows.forEach((row) => {
    const coveragePct = row.totalRequired > 0 ? Math.round((row.totalAllocated / row.totalRequired) * 100) : 100;
    const totalClass =
      row.totalRequired <= 0 || coveragePct >= 100
        ? "is-ok"
        : coveragePct >= 50
        ? "is-mid"
        : "is-gap";
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="sticky-col project-cell">${escapeHtml(row.project)}</td>
      <td class="sticky-col-2 role-cell">${escapeHtml(row.role)}</td>
      <td class="sticky-col-3 total-cell ${totalClass}">
        <div class="total-stack">
          <strong>${row.totalResidual}/${row.totalRequired}</strong>
          <span>${coveragePct}%</span>
        </div>
      </td>
      ${row.weeks
          .filter((w) => w.week >= getCurrentWeek())
          .map((w) => {
          const cls =
              w.residual > 0
                ? "cell-gap"
                : w.required > 0 || w.allocated > 0
                ? "cell-covered"
                : "cell-empty";
          const badge = plannerCellBadge(row.project_id, row.role, w.week, w.required, w.allocated, row.project);
          const noFabClass = badge?.cls === "badge-no-fab" ? "cell-warning-nofab" : "";
          const diffKey = `${row.project}||${row.role}||${w.week}`;
          const diff = diffMap[diffKey];
          const diffOk = diff && (!diffCurrentWeek || Number(w.week) >= diffCurrentWeek);
          const diffClass = diffOk ? "cell-diff" : "";
          const diffTitle = diffOk ? ` Delta fabb ${diff.prev_qty}->${diff.new_qty}` : "";
          const warningText = plannerCellWarningText(row.project_id, row.role, w.week, row.project);
          const nameText = w.names.length ? w.names.join(", ") : "";
          const baseTitle = warningText || nameText;
          const names = baseTitle
            ? ` title="${escapeHtml(baseTitle)}${diffTitle}"`
            : diffTitle
            ? ` title="${escapeHtml(diffTitle.trim())}"`
            : "";
          return `
            <td>
              <button
                class="planner-cell ${cls} ${diffClass} ${noFabClass}"
                data-project-id="${row.project_id}"
                data-project="${escapeHtml(row.project)}"
                data-role="${escapeHtml(row.role)}"
                data-week="${w.week}"
                ${names}
              >
                ${cellBadgeHtml(badge)}
                <span class="metric">R ${w.required}</span>
                <span class="metric">A ${w.allocatedDisplay || formatAllocatedDisplay(w.allocated, 0)}</span>
                <span class="metric">D ${w.residual}</span>
              </button>
            </td>
          `;
        })
        .join("")}
    `;
    plannerBody.appendChild(tr);
  });

  plannerBody.querySelectorAll(".planner-cell").forEach((btn) => {
    btn.addEventListener("click", () => {
      const target = {
        project_id: parseIntOr(btn.dataset.projectId, 0),
        project: btn.dataset.project,
        role: btn.dataset.role,
        week: parseIntOr(btn.dataset.week, 0),
        week_from: parseIntOr(btn.dataset.week, 0),
        week_to: parseIntOr(btn.dataset.week, 0),
      };
      selectTarget(target);
    });
    btn.addEventListener("dragover", (ev) => {
      ev.preventDefault();
      btn.classList.add("drag-over");
    });
    btn.addEventListener("dragleave", () => btn.classList.remove("drag-over"));
    btn.addEventListener("drop", async (ev) => {
      ev.preventDefault();
      btn.classList.remove("drag-over");
      const resourceId = parseIntOr(ev.dataTransfer.getData("text/plain"), 0);
      if (!resourceId) return;
      const target = {
        project_id: parseIntOr(btn.dataset.projectId, 0),
        project: btn.dataset.project,
        role: btn.dataset.role,
        week: parseIntOr(btn.dataset.week, 0),
      };
      try {
        await saveAssignment({
          project_id: target.project_id,
          project: target.project,
          role: target.role,
          resource_id: resourceId,
          week_from: target.week,
          week_to: target.week,
        });
        selectTarget({ ...target, week_from: target.week, week_to: target.week }, resourceId);
      } catch (err) {
        setStatus(err.message);
      }
    });
  });
}

function selectedAssignResourceIds() {
  return [...assignResource.selectedOptions]
    .map((opt) => parseIntOr(opt.value, 0))
    .filter((id) => id > 0);
}

function usingExternalAssignMode() {
  return !!assignUseExternal?.checked;
}

function extResourcesForRole(role = "") {
  return state.resources
    .filter((r) => isExternalResource(r))
    .filter((r) => !role || roleMatches(r, role))
    .sort((a, b) => a.name.localeCompare(b.name));
}

function updateAssignModeUi() {
  const useExt = usingExternalAssignMode();
  if (assignExternalQty) {
    assignExternalQty.disabled = !useExt;
    if (!assignExternalQty.value || Number(assignExternalQty.value) < 1) assignExternalQty.value = "1";
  }
  if (assignResource) {
    assignResource.multiple = !useExt;
    assignResource.size = useExt ? 12 : 15;
  }
  if (markUnavailableBtn) {
    markUnavailableBtn.disabled = useExt;
  }
}

function resourceCardClass(resource, selectedWeekFrom = null, selectedWeekTo = null, role = "") {
  if (isExternalResource(resource)) return "external";
  const weekFrom = selectedWeekFrom ?? null;
  const weekTo = selectedWeekTo ?? selectedWeekFrom ?? null;
  const hasRange = weekFrom !== null && weekTo !== null;
  const conflicts = hasRange ? allocationConflict(resource.id, weekFrom, weekTo) : [];
  const unavailable = hasRange ? unavailabilityConflict(resource.id, weekFrom, weekTo) : [];
  if (!resourceHasAnyRole(resource)) return "off-role";
  if (!resourceActive(resource, hasRange ? weekFrom : null)) return "inactive";
  if (unavailable.length > 0) return "unavailable";
  if (role && !roleMatches(resource, role)) return "off-role";
  if (conflicts.length === 0) return "free";
  if (conflicts.length === 1) return "busy";
  if (conflicts.length === 2) return "partial";
  return "busy";
}

function renderAssignResourceOptions(preselectedId = null) {
  if (!assignResource) return;
  updateAssignModeUi();
  let useExt = usingExternalAssignMode();
  const role = assignRole.value.trim();
  const wf = parseWeek(assignWeekFrom.value);
  const wt = parseWeek(assignWeekTo.value);
  const baseWeek = selectedTarget?.week_from ?? selectedTarget?.week ?? null;
  const weekFrom = wf ?? baseWeek;
  const weekTo = wt ?? weekFrom;
  let scopeResources = useExt
    ? extResourcesForRole(role)
    : state.resources.filter((r) => !isExternalResource(r));
  if (useExt && scopeResources.length === 0) {
    useExt = false;
    if (assignUseExternal) assignUseExternal.checked = false;
    updateAssignModeUi();
    scopeResources = state.resources.filter((r) => !isExternalResource(r));
  }
  const selectedSet = new Set(
    (preselectedId !== null && preselectedId !== undefined ? [preselectedId] : selectedAssignResourceIds()).map((id) =>
      Number(id)
    )
  );
  const options = scopeResources
    .map((r) => {
      const isExt = isExternalResource(r);
      const hasRange = weekFrom !== null && weekTo !== null;
      const conflict = hasRange && !isExt ? allocationConflict(r.id, weekFrom, weekTo) : [];
      const unavailable = hasRange ? unavailabilityConflict(r.id, weekFrom, weekTo) : [];
      const hasRole = resourceHasAnyRole(r);
      const roleOk = !role || roleMatches(r, role);
      const active = resourceActive(r, hasRange ? weekFrom : null);
      const sel = selectedSet.has(Number(r.id)) ? "selected" : "";
      let optionColor = "#111a2a";
      let optionWeight = "700";
      if (!active) {
        optionColor = "#8a94a8";
        optionWeight = "500";
      } else if (unavailable.length > 0 || (!isExt && (conflict.length === 1 || conflict.length > 2))) {
        optionColor = "#b42318";
        optionWeight = "700";
      } else if (!isExt && conflict.length === 2) {
        optionColor = "#b26b00";
        optionWeight = "700";
      } else if (isExt) {
        optionColor = "#1f7fd6";
        optionWeight = "700";
      }
      const displayLine = formatResourceDisplayLine(r, {
        role,
        weekFrom,
        weekTo,
      });
      const existingCodes = new Set(displayLine.split("|").map((v) => String(v || "").trim()).filter(Boolean));
      const extraCodes = [];
      if ((!hasRole || !roleOk) && !existingCodes.has("NO-MANS")) extraCodes.push("NO-MANS");
      if (!active && !existingCodes.has("NA")) extraCodes.push("NA");
      if (unavailable.length > 0 && !existingCodes.has("IND")) extraCodes.push("IND");
      const suffix = extraCodes.length ? ` | ${extraCodes.join(" | ")}` : "";
      return `<option value="${r.id}" ${sel} style="color:${optionColor};font-weight:${optionWeight};">${escapeHtml(displayLine)}${escapeHtml(suffix)}</option>`;
    })
    .join("");
  assignResource.innerHTML = options || `<option value="">Nessuna risorsa disponibile</option>`;
  if (useExt && assignResource.options.length > 0) {
    assignResource.selectedIndex = 0;
  }
}

function renderResourcePool() {
  const useExt = usingExternalAssignMode();
  const role = assignRole.value.trim();
  const weekFrom = parseWeek(assignWeekFrom.value) || selectedTarget?.week_from || selectedTarget?.week || null;
  const weekTo = parseWeek(assignWeekTo.value) || selectedTarget?.week_to || selectedTarget?.week || weekFrom;
  const search = resourceSearch.value.trim().toLowerCase();
  const rows = state.resources
    .filter((r) => (useExt ? isExternalResource(r) : !isExternalResource(r)))
    .filter((r) => (!useExt ? true : !role || roleMatches(r, role)))
    .filter((r) => {
      if (!search) return true;
      const hay = [r.name, r.employee_code, r.role1, r.role2, r.base_location, r.level].join(" ").toLowerCase();
      return hay.includes(search);
    })
    .sort((a, b) => a.name.localeCompare(b.name))
    .map((r) => {
      const cls = resourceCardClass(r, weekFrom, weekTo, role);
      const displayLine = formatResourceDisplayLine(r, {
        role,
        weekFrom,
        weekTo,
      });
      const note = displayLine;
      const lineParts = displayLine.split(" | ");
      const head = lineParts.shift() || r.name || "";
      const meta = lineParts.join(" | ");
      const statusClass = cls === "free" ? "available" : cls;
      return `
        <div class="resource-line ${statusClass}" draggable="true" data-resource-id="${r.id}" title="${escapeHtml(note)}">
          <span class="resource-line-name">${escapeHtml(head)}</span>
          <span class="resource-line-meta">${escapeHtml(meta)}</span>
        </div>
      `;
    })
    .join("");
  resourcePool.innerHTML = rows || `<div class="empty-box">Nessuna risorsa trovata.</div>`;
  resourcePool.querySelectorAll(".resource-line").forEach((rowEl) => {
    rowEl.addEventListener("dragstart", (ev) => {
      ev.dataTransfer.setData("text/plain", rowEl.dataset.resourceId);
    });
    rowEl.addEventListener("click", () => {
      renderAssignResourceOptions(parseIntOr(rowEl.dataset.resourceId, 0));
    });
  });
}

function renderSelectionAllocations() {
  if (!selectedTarget) {
    selectedAllocationIds.clear();
    selectionAllocations.innerHTML = `<div class="empty-box">Seleziona una cella per vedere le allocazioni.</div>`;
    if (workshopBreakdownBox) {
      workshopBreakdownBox.innerHTML = `<div class="empty-box">Seleziona OVERALL OFFICINA per vedere il dettaglio delle commesse WS.</div>`;
    }
    return;
  }
  const wf = parseWeek(assignWeekFrom.value) || selectedTarget.week_from || selectedTarget.week || 1;
  const wt = parseWeek(assignWeekTo.value) || selectedTarget.week_to || selectedTarget.week || wf;
  const rows = state.allocations
    .filter(
      (a) => {
        const meta = plannerProjectMeta(a.project_id, a.project);
        return normalizeProjectCode(meta.project) === normalizeProjectCode(selectedTarget.project) &&
          a.role === selectedTarget.role &&
          Number(a.week_from) <= wt &&
          Number(a.week_to) >= wf;
      }
    )
    .sort((a, b) => a.resource.localeCompare(b.resource));
  const visibleIds = new Set(rows.map((r) => Number(r.id)));
  selectedAllocationIds = new Set([...selectedAllocationIds].filter((id) => visibleIds.has(Number(id))));

  function contextualWarnings(allocation, viewFrom, viewTo) {
    return warningCodesForAllocationRange(allocation, viewFrom, viewTo);
  }

  selectionAllocations.innerHTML = rows.length
    ? rows
        .map(
          (a) => {
            const resource = state.resources.find((r) => Number(r.id) === Number(a.resource_id));
            const warnings = contextualWarnings(a, wf, wt);
            const line = resource
              ? formatResourceDisplayLine(resource, {
                  role: a.role,
                  weekFrom: wf,
                  weekTo: wt,
                  allocation: a,
                })
              : `${formatShortPersonName(a.resource)} | ${formatShortRole(a.role)}`;
            const packed = appendUniqueCompactCodes(line, warnings);
            const info = [
              packed.line,
              `W${a.week_from}${Number(a.week_to) !== Number(a.week_from) ? `-W${a.week_to}` : ""}`,
              ...packed.extras,
            ].filter(Boolean);
            return `
            <div class="selection-item ${selectedAllocationIds.has(Number(a.id)) ? "is-selected" : ""}" data-allocation-id="${a.id}">
              <span class="selection-line">${escapeHtml(info.join(" | "))}</span>
              <button class="cell-btn danger" data-del-allocation="${a.id}">X</button>
            </div>
          `
          }
        )
        .join("")
    : `<div class="empty-box">Nessuna allocazione nel periodo selezionato.</div>`;

  selectionAllocations.querySelectorAll("[data-del-allocation]").forEach((btn) => {
    btn.addEventListener("click", async () => {
      await api(`/api/allocations/${btn.dataset.delAllocation}`, { method: "DELETE" });
      await loadAll();
      selectedAllocationIds.delete(Number(btn.dataset.delAllocation));
      renderSelectionAllocations();
    });
  });
  selectionAllocations.querySelectorAll(".selection-item[data-allocation-id]").forEach((rowEl) => {
    rowEl.addEventListener("click", (ev) => {
      if (ev.target && ev.target.closest("[data-del-allocation]")) return;
      const id = Number(rowEl.dataset.allocationId);
      if (selectedAllocationIds.has(id)) selectedAllocationIds.delete(id);
      else selectedAllocationIds.add(id);
      rowEl.classList.toggle("is-selected", selectedAllocationIds.has(id));
      selectionAllocations.focus();
    });
  });
  renderWorkshopBreakdown().catch((err) => {
    console.error("Errore dettaglio officina:", err);
  });
}

async function deleteSelectedAllocations() {
  if (!selectedAllocationIds.size) return;
  const ids = [...selectedAllocationIds];
  for (const id of ids) {
    await api(`/api/allocations/${id}`, { method: "DELETE" });
  }
  selectedAllocationIds.clear();
  await loadAll(true);
  renderSelectionAllocations();
  setStatus(ids.length === 1 ? "Allocazione rimossa." : `${ids.length} allocazioni rimosse.`);
}

async function renderWorkshopBreakdown() {
  if (!workshopBreakdownBox) return;
  const overallSelected = (plannerProjectFilter?.value || "").trim().toUpperCase() === "OVERALL OFFICINA";
  if (!overallSelected) {
    if (workshopBreakdownSection) workshopBreakdownSection.hidden = true;
    workshopBreakdownBox.innerHTML = `<div class="empty-box">Seleziona una cella di OVERALL OFFICINA per vedere il dettaglio.</div>`;
    return;
  }
  if (!selectedTarget || selectedTarget.project !== "OVERALL OFFICINA") {
    if (workshopBreakdownSection) workshopBreakdownSection.hidden = false;
    workshopBreakdownBox.innerHTML = `<div class="empty-box">Seleziona una cella di OVERALL OFFICINA per vedere il dettaglio.</div>`;
    return;
  }
  if (workshopBreakdownSection) workshopBreakdownSection.hidden = false;
  const week = parseWeek(assignWeekFrom.value) || selectedTarget.week || selectedTarget.week_from || getCurrentWeek();
  let rows = [];
  try {
    rows = await api(
      `/api/workshop-breakdown?week=${encodeURIComponent(week)}&role=${encodeURIComponent(selectedTarget.role)}`
    );
  } catch (err) {
    rows = currentWorkshopBreakdown(selectedTarget.role, week);
  }
  if (!rows || !rows.length) {
    // Fallback locale: evita pannello vuoto se API non trova il dettaglio ma lo stato locale ce l'ha.
    rows = currentWorkshopBreakdown(selectedTarget.role, week);
  }
  workshopBreakdownBox.innerHTML = rows.length
    ? rows
        .map(
          (r) => `
            <div class="selection-item">
              <div>
                <strong>${escapeHtml(r.project)}</strong>
                <div class="selection-sub">${escapeHtml(r.role)} | W${r.week} | richiesto ${r.qty}</div>
              </div>
            </div>
          `
        )
        .join("")
    : `<div class="empty-box">Nessun dettaglio WS per ${escapeHtml(selectedTarget.role)} in W${week}.</div>`;
}

function resourcePayloadFromRow(tr) {
  return {
    employee_code: tr.querySelector('[data-f="employee_code"]').value.trim(),
    name: tr.querySelector('[data-f="name"]').value.trim(),
    role1: tr.querySelector('[data-f="role1"]').value.trim(),
    role2: tr.querySelector('[data-f="role2"]').value.trim(),
    hire_date: tr.querySelector('[data-f="hire_date"]').value.trim(),
    end_date: tr.querySelector('[data-f="end_date"]').value.trim(),
      birth_date: tr.querySelector('[data-f="birth_date"]').value.trim(),
      phone: tr.querySelector('[data-f="phone"]').value.trim(),
      residence_city: tr.querySelector('[data-f="residence_city"]').value.trim(),
      employer: tr.querySelector('[data-f="employer"]').value.trim() || "SPEC",
      hire_level: tr.querySelector('[data-f="hire_level"]').value.trim(),
      level: tr.querySelector('[data-f="level"]').value.trim(),
      contract_type: tr.querySelector('[data-f="contract_type"]').value.trim(),
    pb_overtime_hourly: parseFloatSafe(tr.querySelector('[data-f="pb_overtime_hourly"]').value),
    pb_day_office: parseFloatSafe(tr.querySelector('[data-f="pb_day_office"]').value),
    pb_day_site: parseFloatSafe(tr.querySelector('[data-f="pb_day_site"]').value),
    pb_fixed_plus: parseFloatSafe(tr.querySelector('[data-f="pb_fixed_plus"]').value),
    glob_hour_office: parseFloatSafe(tr.querySelector('[data-f="glob_hour_office"]').value),
    glob_hour_site: parseFloatSafe(tr.querySelector('[data-f="glob_hour_site"]').value),
    glob_fixed_plus: parseFloatSafe(tr.querySelector('[data-f="glob_fixed_plus"]').value),
    doc_type: tr.querySelector('[data-f="doc_type"]').value.trim(),
    doc_number: tr.querySelector('[data-f="doc_number"]').value.trim(),
    doc_expiry: tr.querySelector('[data-f="doc_expiry"]').value.trim(),
    email: tr.querySelector('[data-f="email"]').value.trim(),
    base_location: tr.querySelector('[data-f="base_location"]').value.trim(),
    certifications: tr.querySelector('[data-f="certifications"]').value.trim(),
    note: tr.querySelector('[data-f="note"]').value.trim(),
    no_travel: tr.querySelector('[data-f="no_travel"]').checked,
  };
}

function parseFloatSafe(value) {
  const normalized = String(value ?? "")
    .replace(",", ".")
    .trim();
  if (!normalized) return 0;
  const parsed = Number.parseFloat(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function sortValueEndDate(resource) {
  const endWeek = parseWeekOrDate(resource.end_date);
  if (endWeek === null) return 999;
  if (endWeek <= 0) return -1;
  return endWeek;
}

function populateResourceTableControls() {
  if (!resourceRoleFilter) return;
  const current = resourceRoleFilter.value;
  const roles = [...new Set(state.resources.flatMap((r) => [r.role1, r.role2]).filter(Boolean))].sort((a, b) =>
    a.localeCompare(b)
  );
  resourceRoleFilter.innerHTML = [
    `<option value="">Tutte le mansioni</option>`,
    ...roles.map((role) => `<option value="${escapeHtml(role)}">${escapeHtml(role)}</option>`),
  ].join("");
  resourceRoleFilter.value = roles.includes(current) ? current : "";
}

function filteredResourcesForTable() {
  const q = (resourceTableSearch?.value || "").trim().toLowerCase();
  const role = (resourceRoleFilter?.value || "").trim();
  const status = (resourceStatusFilter?.value || "").trim();
  const sortBy = resourceSortBy?.value || "name_asc";
  const rows = state.resources.filter((r) => {
    const hay = [
      r.name,
      r.employee_code,
      r.role1,
        r.role2,
        r.residence_city,
        r.employer,
        r.hire_level,
        r.level,
        r.contract_type,
        r.doc_number,
      r.note,
    ]
      .join(" ")
      .toLowerCase();
    if (q && !hay.includes(q)) return false;
    if (role && r.role1 !== role && r.role2 !== role) return false;
    if (status && (r.status || "") !== status) return false;
    return true;
  });
  rows.sort((a, b) => {
    if (sortBy === "name_desc") return b.name.localeCompare(a.name);
    if (sortBy === "role_asc") return (a.role1 || "").localeCompare(b.role1 || "") || a.name.localeCompare(b.name);
    if (sortBy === "end_asc") return sortValueEndDate(a) - sortValueEndDate(b) || a.name.localeCompare(b.name);
    if (sortBy === "end_desc") return sortValueEndDate(b) - sortValueEndDate(a) || a.name.localeCompare(b.name);
    return a.name.localeCompare(b.name);
  });
  return rows;
}

function renderResourcesTable() {
  populateResourceTableControls();
  resourcesTableBody.innerHTML = "";
  resourceOriginalSnapshot = new Map();
  const newRow = document.createElement("tr");
  newRow.className = "row-new";
  newRow.innerHTML = `
      <td>+</td>
      <td><input class="cell-input" data-f="employee_code" /></td>
      <td><input class="cell-input" data-f="name" placeholder="Nuova risorsa" /></td>
      <td><input class="cell-input" data-f="role1" list="roleList" /></td>
      <td><input class="cell-input" data-f="role2" list="roleList" /></td>
      <td><input class="cell-input" data-f="hire_date" placeholder="W12 o dd/mm/yyyy" /></td>
      <td><input class="cell-input" data-f="end_date" /></td>
        <td><input class="cell-input" data-f="birth_date" placeholder="dd/mm/yyyy" /></td>
        <td><input class="cell-input" data-f="phone" /></td>
        <td><input class="cell-input" data-f="residence_city" /></td>
        <td><input class="cell-input" data-f="employer" value="SPEC" /></td>
        <td><input class="cell-input" data-f="hire_level" /></td>
        <td>
          <select class="cell-input" data-f="contract_type">
            <option value=""></option>
            <option value="PB">PB</option>
            <option value="GLOB">GLOB</option>
        </select>
      </td>
      <td><input class="cell-input cell-input-number" data-f="pb_overtime_hourly" type="number" step="0.01" /></td>
      <td><input class="cell-input cell-input-number" data-f="pb_day_office" type="number" step="0.01" /></td>
      <td><input class="cell-input cell-input-number" data-f="pb_day_site" type="number" step="0.01" /></td>
      <td><input class="cell-input cell-input-number" data-f="pb_fixed_plus" type="number" step="0.01" /></td>
      <td><input class="cell-input cell-input-number" data-f="glob_hour_office" type="number" step="0.01" /></td>
      <td><input class="cell-input cell-input-number" data-f="glob_hour_site" type="number" step="0.01" /></td>
      <td><input class="cell-input cell-input-number" data-f="glob_fixed_plus" type="number" step="0.01" /></td>
      <td><input class="cell-input" data-f="doc_type" /></td>
        <td><input class="cell-input" data-f="doc_number" /></td>
        <td><input class="cell-input" data-f="doc_expiry" placeholder="dd/mm/yyyy" /></td>
        <td><input class="cell-input" data-f="email" /></td>
        <td><input class="cell-input" data-f="base_location" /></td>
        <td><input class="cell-input" data-f="level" /></td>
        <td><input class="cell-input" data-f="certifications" /></td>
        <td><input data-f="no_travel" type="checkbox" /></td>
        <td><input class="cell-input" data-f="note" /></td>
      <td></td>
      <td class="cell-actions"><span class="hint">Nuova riga</span></td>
  `;
  resourcesTableBody.appendChild(newRow);

  filteredResourcesForTable().forEach((r) => {
    resourceOriginalSnapshot.set(Number(r.id), JSON.stringify(resourcePayloadFromExistingResource(r)));
    const tr = document.createElement("tr");
    if (r.status === "CESSATO") tr.classList.add("row-ended");
    if (deletedResourceIds.has(Number(r.id))) tr.classList.add("row-pending-delete");
    tr.dataset.id = String(r.id);
      tr.innerHTML = `
        <td>${r.id}</td>
        <td><input class="cell-input" data-f="employee_code" value="${escapeHtml(r.employee_code || "")}" /></td>
        <td><input class="cell-input" data-f="name" value="${escapeHtml(r.name)}" /></td>
        <td><input class="cell-input" data-f="role1" list="roleList" value="${escapeHtml(r.role1 || "")}" /></td>
        <td><input class="cell-input" data-f="role2" list="roleList" value="${escapeHtml(r.role2 || "")}" /></td>
        <td><input class="cell-input" data-f="hire_date" value="${escapeHtml(r.hire_date || "")}" /></td>
        <td><input class="cell-input" data-f="end_date" value="${escapeHtml(r.end_date || "")}" /></td>
          <td><input class="cell-input" data-f="birth_date" value="${escapeHtml(r.birth_date || "")}" /></td>
          <td><input class="cell-input" data-f="phone" value="${escapeHtml(r.phone || "")}" /></td>
          <td><input class="cell-input" data-f="residence_city" value="${escapeHtml(r.residence_city || "")}" /></td>
          <td><input class="cell-input" data-f="employer" value="${escapeHtml(r.employer || "SPEC")}" /></td>
          <td><input class="cell-input" data-f="hire_level" value="${escapeHtml(r.hire_level || "")}" /></td>
          <td>
            <select class="cell-input" data-f="contract_type">
              <option value="" ${(r.contract_type || "") === "" ? "selected" : ""}></option>
              <option value="PB" ${r.contract_type === "PB" ? "selected" : ""}>PB</option>
              <option value="GLOB" ${r.contract_type === "GLOB" ? "selected" : ""}>GLOB</option>
          </select>
        </td>
        <td><input class="cell-input cell-input-number" data-f="pb_overtime_hourly" type="number" step="0.01" value="${escapeHtml(r.pb_overtime_hourly || "")}" /></td>
        <td><input class="cell-input cell-input-number" data-f="pb_day_office" type="number" step="0.01" value="${escapeHtml(r.pb_day_office || "")}" /></td>
        <td><input class="cell-input cell-input-number" data-f="pb_day_site" type="number" step="0.01" value="${escapeHtml(r.pb_day_site || "")}" /></td>
        <td><input class="cell-input cell-input-number" data-f="pb_fixed_plus" type="number" step="0.01" value="${escapeHtml(r.pb_fixed_plus || "")}" /></td>
        <td><input class="cell-input cell-input-number" data-f="glob_hour_office" type="number" step="0.01" value="${escapeHtml(r.glob_hour_office || "")}" /></td>
        <td><input class="cell-input cell-input-number" data-f="glob_hour_site" type="number" step="0.01" value="${escapeHtml(r.glob_hour_site || "")}" /></td>
        <td><input class="cell-input cell-input-number" data-f="glob_fixed_plus" type="number" step="0.01" value="${escapeHtml(r.glob_fixed_plus || "")}" /></td>
        <td><input class="cell-input" data-f="doc_type" value="${escapeHtml(r.doc_type || "")}" /></td>
          <td><input class="cell-input" data-f="doc_number" value="${escapeHtml(r.doc_number || "")}" /></td>
          <td><input class="cell-input" data-f="doc_expiry" value="${escapeHtml(r.doc_expiry || "")}" /></td>
          <td><input class="cell-input" data-f="email" value="${escapeHtml(r.email || "")}" /></td>
          <td><input class="cell-input" data-f="base_location" value="${escapeHtml(r.base_location || "")}" /></td>
          <td><input class="cell-input" data-f="level" value="${escapeHtml(r.level || "")}" /></td>
          <td><input class="cell-input" data-f="certifications" value="${escapeHtml(r.certifications || "")}" /></td>
          <td><input data-f="no_travel" type="checkbox" ${r.no_travel ? "checked" : ""} /></td>
          <td><input class="cell-input" data-f="note" value="${escapeHtml(r.note || "")}" /></td>
        <td>${escapeHtml(r.status || "")}</td>
        <td class="cell-actions"><button class="cell-btn danger-btn" data-act="delete-resource">${deletedResourceIds.has(Number(r.id)) ? "Annulla elimina" : "Elimina"}</button></td>
      `;
      resourcesTableBody.appendChild(tr);
    });
    resourcesTableBody.querySelectorAll("tr").forEach((tr) => {
      syncResourceContractFields(tr);
      const contractSelect = tr.querySelector('[data-f="contract_type"]');
      contractSelect?.addEventListener("change", () => syncResourceContractFields(tr));
    });
    resourcesTableBody.querySelectorAll('[data-act="delete-resource"]').forEach((btn) => {
      btn.addEventListener("click", (ev) => {
        const tr = ev.target.closest("tr");
      const id = parseIntOr(tr.dataset.id, 0);
      if (!id) return;
      if (deletedResourceIds.has(id)) {
        deletedResourceIds.delete(id);
      } else {
        deletedResourceIds.add(id);
      }
      renderResourcesTable();
      setStatus("Eliminazione messa in coda. Premi 'Salva modifiche' per confermare.");
    });
  });
  renderResourceColumnChooser();
  applyResourceColumnVisibility();
  bindResourceColumnResizers();
  bindResourceTableKeyboardNavigation();
}

function syncResourceContractFields(tr) {
  if (!tr) return;
  const type = (tr.querySelector('[data-f="contract_type"]')?.value || "").trim().toUpperCase();
  const pbFields = ["pb_overtime_hourly", "pb_day_office", "pb_day_site", "pb_fixed_plus"];
  const globFields = ["glob_hour_office", "glob_hour_site", "glob_fixed_plus"];
  const setDisabled = (field, disabled) => {
    const input = tr.querySelector(`[data-f="${field}"]`);
    if (!input) return;
    input.disabled = disabled;
    input.classList.toggle("is-disabled", disabled);
  };
  if (type === "PB") {
    pbFields.forEach((field) => setDisabled(field, false));
    globFields.forEach((field) => setDisabled(field, true));
    return;
  }
  if (type === "GLOB") {
    pbFields.forEach((field) => setDisabled(field, true));
    globFields.forEach((field) => setDisabled(field, false));
    return;
  }
  [...pbFields, ...globFields].forEach((field) => setDisabled(field, false));
}

function resourcePayloadFromExistingResource(resource) {
  return {
    employee_code: resource.employee_code || "",
    name: resource.name || "",
    role1: resource.role1 || "",
    role2: resource.role2 || "",
    hire_date: resource.hire_date || "",
    end_date: resource.end_date || "",
      birth_date: resource.birth_date || "",
      phone: resource.phone || "",
      residence_city: resource.residence_city || "",
      employer: resource.employer || "SPEC",
      hire_level: resource.hire_level || "",
      level: resource.level || "",
      contract_type: resource.contract_type || "",
    pb_overtime_hourly: resource.pb_overtime_hourly || 0,
    pb_day_office: resource.pb_day_office || 0,
    pb_day_site: resource.pb_day_site || 0,
    pb_fixed_plus: resource.pb_fixed_plus || 0,
    glob_hour_office: resource.glob_hour_office || 0,
    glob_hour_site: resource.glob_hour_site || 0,
    glob_fixed_plus: resource.glob_fixed_plus || 0,
    doc_type: resource.doc_type || "",
    doc_number: resource.doc_number || "",
    doc_expiry: resource.doc_expiry || "",
    email: resource.email || "",
    base_location: resource.base_location || "",
    certifications: resource.certifications || "",
    note: resource.note || "",
    no_travel: !!resource.no_travel,
  };
}

async function saveAllResources() {
  const rows = [...resourcesTableBody.querySelectorAll("tr")];
  let created = 0;
  let updated = 0;
  let deleted = 0;

  const newRow = rows.find((tr) => tr.classList.contains("row-new"));
  if (newRow) {
    const payload = resourcePayloadFromRow(newRow);
    const nameFilled = String(payload.name || "").trim() !== "";
    const roleFilled = String(payload.role1 || "").trim() !== "" || String(payload.role2 || "").trim() !== "";
    const hasMeaningful = nameFilled || roleFilled;
    if (hasMeaningful) {
      if (!nameFilled || !roleFilled) {
        setStatus("Per la nuova risorsa compila almeno nome e una mansione.");
        return;
      }
      await api("/api/resources", { method: "POST", body: JSON.stringify(payload) });
      created += 1;
    }
  }

  for (const tr of rows.filter((row) => row.dataset.id)) {
    const id = parseIntOr(tr.dataset.id, 0);
    if (!id) continue;
    if (deletedResourceIds.has(id)) {
      await api(`/api/resources/${id}`, { method: "DELETE" });
      deleted += 1;
      continue;
    }
    const payload = resourcePayloadFromRow(tr);
    if (!payload.name || (!payload.role1 && !payload.role2)) {
      setStatus("Ogni risorsa deve avere almeno nome e una mansione.");
      return;
    }
    const original = resourceOriginalSnapshot.get(id) || "";
    const current = JSON.stringify(payload);
    if (current !== original) {
      await api(`/api/resources/${id}`, { method: "PUT", body: JSON.stringify(payload) });
      updated += 1;
    }
  }

  deletedResourceIds = new Set();
  if (created || updated || deleted) {
    logUserAction("Risorse salvate", `Nuove ${created}, aggiornate ${updated}, eliminate ${deleted}`);
  }
  await loadAll(true);
  setStatus(`Risorse salvate. Nuove: ${created}, aggiornate: ${updated}, eliminate: ${deleted}.`);
}

function filteredProjectsAdmin() {
  const q = (projectAdminSearch?.value || "").trim().toLowerCase();
  return state.projects
    .filter((p) => {
      if (p.closed) return false;
      if (p.code === "WS OVERALL") return false;
      if (!q) return true;
      return [p.code, p.activity, p.type].join(" ").toLowerCase().includes(q);
    })
    .sort((a, b) => {
      if (a.code === "OVERALL OFFICINA") return -1;
      if (b.code === "OVERALL OFFICINA") return 1;
      return projectSortWeek(a.id) - projectSortWeek(b.id) || a.code.localeCompare(b.code);
    });
}

function ensureAdminProjectSelection() {
  if (adminProjectDraft) return;
  const rows = filteredProjectsAdmin();
  if (selectedAdminProjectId && rows.some((p) => Number(p.id) === Number(selectedAdminProjectId))) return;
  selectedAdminProjectId = rows[0] ? Number(rows[0].id) : null;
}

function selectedAdminProject() {
  return state.projects.find((p) => Number(p.id) === Number(selectedAdminProjectId)) || null;
}

function renderProjectsAdmin() {
  if (!projectsAdminTableBody) return;
  ensureAdminProjectSelection();
  projectsAdminTableBody.innerHTML = filteredProjectsAdmin()
    .map(
      (p) => `
        <tr data-project-admin-id="${p.id}" class="${Number(p.id) === Number(selectedAdminProjectId) ? "is-selected" : ""}">
          <td>${escapeHtml(p.code)}</td>
          <td>${escapeHtml(p.type || "")}</td>
          <td>${p.closed ? "Si" : ""}</td>
          <td>${p.workshop_rollup ? "Si" : ""}</td>
        </tr>
      `
    )
    .join("");
  projectsAdminTableBody.querySelectorAll("[data-project-admin-id]").forEach((tr) => {
    tr.addEventListener("click", () => {
      adminProjectDraft = false;
      selectedAdminProjectId = Number(tr.dataset.projectAdminId);
      projectDemandExtraRoles[selectedAdminProjectId] = projectDemandExtraRoles[selectedAdminProjectId] || [];
      fillProjectAdminForm();
      renderProjectsAdmin();
      renderProjectDemandMatrix();
    });
  });
  fillProjectAdminForm();
  renderProjectDemandMatrix();
}

function fillProjectAdminForm() {
  const project = selectedAdminProject();
  if (!project) {
    projectCodeInput.value = "";
    projectActivityInput.value = "";
    projectTypeInput.value = "SITE";
    projectClosedInput.checked = false;
    projectWorkshopRollupInput.checked = false;
    return;
  }
  projectCodeInput.value = project.code || "";
  projectActivityInput.value = project.activity || project.code || "";
  projectTypeInput.value = project.type || "SITE";
  projectClosedInput.checked = Boolean(project.closed);
  projectWorkshopRollupInput.checked = Boolean(project.workshop_rollup);
}

function demandRolesForProject(projectId) {
  const existing = state.demands
    .filter((d) => Number(d.project_id) === Number(projectId))
    .map((d) => d.role)
    .filter(Boolean);
  const extras = projectDemandExtraRoles[projectId] || [];
  return [...new Set([...existing, ...extras])].sort((a, b) => a.localeCompare(b));
}

function renderProjectDemandMatrix() {
  if (!projectDemandMatrixHead || !projectDemandMatrixBody) return;
  const project = selectedAdminProject();
  if (!project) {
    projectDemandMatrixHead.innerHTML = "";
    projectDemandMatrixBody.innerHTML = `<tr><td class="empty-box">Seleziona o crea una commessa.</td></tr>`;
    return;
  }
  if (project.code === "OVERALL OFFICINA") {
    const children = filteredProjectsAdmin().filter((p) => isWorkshopChildProject(p));
    projectDemandMatrixHead.innerHTML = "";
    projectDemandMatrixBody.innerHTML = children.length
      ? children
          .map(
            (child) => `
              <tr>
                <td class="sticky-role"><strong>${escapeHtml(child.code)}</strong></td>
                <td>
                  <button class="cell-btn" data-open-workshop-project="${child.id}">Apri sottocommessa</button>
                </td>
              </tr>
            `
          )
          .join("")
      : `<tr><td class="sticky-role empty-box">Nessuna sottocommessa WS presente. Crea una nuova commessa di tipo WS e attiva "In OVERALL OFFICINA".</td><td></td></tr>`;
    projectDemandMatrixBody.querySelectorAll("[data-open-workshop-project]").forEach((btn) => {
      btn.addEventListener("click", () => {
        adminProjectDraft = false;
        selectedAdminProjectId = Number(btn.dataset.openWorkshopProject);
        fillProjectAdminForm();
        renderProjectsAdmin();
        renderProjectDemandMatrix();
      });
    });
    return;
  }
  const weeks = visibleWeeks();
  const roles = demandRolesForProject(project.id);
  const totals = weeks.map((week) => {
    const required = roles.reduce((sum, role) => sum + currentDemandQty(project.id, role, week), 0);
    const allocated = state.allocations.filter(
      (a) =>
        Number(a.project_id) === Number(project.id) &&
        Number(a.week_from) <= week &&
        Number(a.week_to) >= week
    ).length;
    return {
      week,
      required,
      allocated,
      residual: Math.max(required - allocated, 0),
    };
  });
  projectDemandMatrixHead.innerHTML = `
    <tr class="planner-subtotals-row">
      <th class="sticky-role planner-subtotals-label">Subtotali</th>
      ${totals
        .map(
          (w) => `
            <th class="week-subtotal-head" title="${weekLabelTooltip(w.week)}">
              <span class="metric">R ${w.required}</span>
              <span class="metric">A ${w.allocated}</span>
              <span class="metric">D ${w.residual}</span>
            </th>
          `
        )
        .join("")}
    </tr>
    <tr>
      <th class="sticky-role">Mansione</th>
      ${weeks.map((w) => `<th title="${weekLabelTooltip(w)}">W${w}</th>`).join("")}
    </tr>
  `;
  if (roles.length === 0) {
    projectDemandMatrixBody.innerHTML = `<tr><td class="sticky-role empty-box">Nessuna mansione. Aggiungine una per iniziare.</td>${weeks
      .map(() => "<td></td>")
      .join("")}</tr>`;
    return;
  }
  projectDemandMatrixBody.innerHTML = roles
    .map(
      (role) => `
        <tr data-demand-role="${escapeHtml(role)}">
          <td class="sticky-role"><strong>${escapeHtml(role)}</strong></td>
          ${weeks
            .map((week) => {
              const qty = currentDemandQty(project.id, role, week);
              return `<td><input class="cell-input qty-input" data-demand-project="${project.id}" data-demand-role="${escapeHtml(
                role
              )}" data-demand-week="${week}" type="number" min="0" value="${qty}" /></td>`;
            })
            .join("")}
        </tr>
      `
    )
    .join("");
}

async function saveProjectAdmin() {
  const payload = {
    code: projectCodeInput.value.trim(),
    activity: projectActivityInput.value.trim(),
    type: projectTypeInput.value,
    closed: projectClosedInput.checked,
    workshop_rollup: projectWorkshopRollupInput.checked,
  };
  if (!payload.code) {
    setStatus("Inserisci il codice commessa.");
    return;
  }
  if (selectedAdminProjectId) {
    await api(`/api/projects/${selectedAdminProjectId}`, { method: "PUT", body: JSON.stringify(payload) });
  } else {
    const res = await api("/api/projects", { method: "POST", body: JSON.stringify(payload) });
    selectedAdminProjectId = Number(res.id);
  }
  adminProjectDraft = false;
  await loadAll(true);
  renderProjectsAdmin();
  setStatus("Commessa salvata.");
}

async function saveProjectDemandMatrix() {
  const project = selectedAdminProject();
  if (!project) {
    setStatus("Seleziona una commessa.");
    return;
  }
  const rows = [...projectDemandMatrixBody.querySelectorAll("[data-demand-project]")].map((input) => ({
    project_id: Number(input.dataset.demandProject),
    role: input.dataset.demandRole,
    week: Number(input.dataset.demandWeek),
    qty: Math.max(0, parseIntOr(input.value, 0)),
  }));
  await api("/api/demands/bulk", {
    method: "POST",
    body: JSON.stringify({ rows }),
  });
  await loadAll(true);
  renderProjectsAdmin();
  setStatus("Fabbisogno commessa aggiornato.");
}

function startNewProjectAdmin() {
  adminProjectDraft = true;
  selectedAdminProjectId = null;
  projectCodeInput.value = "";
  projectActivityInput.value = "";
  projectTypeInput.value = "SITE";
  projectClosedInput.checked = false;
  projectWorkshopRollupInput.checked = false;
  renderProjectsAdmin();
}

async function addDemandRoleToProject() {
  const project = selectedAdminProject();
  const role = projectDemandRoleInput.value.trim();
  if (!project) {
    setStatus("Salva o seleziona prima una commessa.");
    return;
  }
  if (!role) {
    setStatus("Inserisci una mansione da aggiungere.");
    return;
  }
  const normalized = role.toUpperCase();
  const weeks = visibleWeeks();
  await api("/api/demands/bulk", {
    method: "POST",
    body: JSON.stringify({
      rows: weeks.map((week) => ({
        project_id: Number(project.id),
        role: normalized,
        week,
        qty: currentDemandQty(project.id, normalized, week),
      })),
    }),
  });
  projectDemandExtraRoles[project.id] = [...new Set([...(projectDemandExtraRoles[project.id] || []), normalized])];
  projectDemandRoleInput.value = "";
  await loadAll(true);
  renderProjectsAdmin();
  renderProjectDemandMatrix();
  setStatus(`Mansione ${normalized} aggiunta alla commessa ${project.code}.`);
}

function defaultReportBuilderState() {
  const current = getCurrentWeek();
  return {
    presetId: "demand_allocated",
    weekFrom: current,
    weekTo: Math.min(52, current + 7),
    scale: "1W",
    chartType: "stacked-demand",
    aggregateScope: "",
    projectFilter: "",
    roleFilter: "",
    mode: "operativa",
  };
}

function loadReportBuilderState() {
  const base = defaultReportBuilderState();
  try {
    const raw = window.localStorage.getItem("report_builder_state_v1");
    const parsed = raw ? JSON.parse(raw) : {};
    reportBuilderState = { ...base, ...(parsed || {}) };
    // Always reopen the report page on the clean 4-quadrant dashboard.
    reportBuilderState.presetId = "demand_allocated";
    reportBuilderState.mode = "operativa";
  } catch (err) {
    reportBuilderState = base;
  }
  reportDetailPresetId = "demand_allocated";
  reportDetailOpen = false;
}

function saveReportBuilderState() {
  try {
    window.localStorage.setItem("report_builder_state_v1", JSON.stringify(reportBuilderState));
  } catch (err) {
    // ignore
  }
}

function syncReportBuilderControls() {
  if (reportWeekFrom) reportWeekFrom.value = reportBuilderState.weekFrom;
  if (reportWeekTo) reportWeekTo.value = reportBuilderState.weekTo;
  if (reportScale) reportScale.value = reportBuilderState.scale;
  if (reportChartType) reportChartType.value = reportBuilderState.chartType;
  if (reportAggregateScope) reportAggregateScope.value = reportBuilderState.aggregateScope;
  if (reportProjectFilter) reportProjectFilter.value = reportBuilderState.projectFilter;
  if (reportRoleFilter) reportRoleFilter.value = reportBuilderState.roleFilter;
  reportModeRequestedBtn?.classList.toggle("active", (reportBuilderState.mode || "operativa") === "mansione");
  reportModeRealBtn?.classList.toggle("active", (reportBuilderState.mode || "operativa") === "operativa");
}

function readReportBuilderControls() {
  reportBuilderState.weekFrom = parseIntOr(reportWeekFrom?.value, getCurrentWeek());
  reportBuilderState.weekTo = parseIntOr(reportWeekTo?.value, reportBuilderState.weekFrom);
  reportBuilderState.scale = reportScale?.value || "1W";
  reportBuilderState.chartType = reportChartType?.value || "stacked-demand";
  reportBuilderState.aggregateScope = reportAggregateScope?.value || "";
  reportBuilderState.projectFilter = reportProjectFilter?.value || "";
  reportBuilderState.roleFilter = reportRoleFilter?.value || "";
  saveReportBuilderState();
}

function reportEligibleProjects() {
  return (state.projects || [])
    .filter((p) => !p.closed)
    .filter((p) => {
      const code = String(p.code || "");
      if (code === "OVERALL OFFICINA") return true;
      if (String(p.type || "").toUpperCase() === "WS") return false;
      return true;
    })
    .sort((a, b) => {
      if (a.code === "OVERALL OFFICINA") return 1;
      if (b.code === "OVERALL OFFICINA") return -1;
      return projectSortWeek(a.id) - projectSortWeek(b.id) || String(a.code || "").localeCompare(String(b.code || ""));
    });
}

function fillReportBuilderFilters() {
  const prevProject = reportProjectFilter?.value || reportBuilderState.projectFilter || "";
  const prevRole = reportRoleFilter?.value || reportBuilderState.roleFilter || "";
  if (reportProjectFilter) {
    reportProjectFilter.innerHTML = '<option value="">Tutte le commesse</option>';
    reportEligibleProjects().forEach((project) => {
      const option = document.createElement("option");
      option.value = project.code;
      option.textContent = project.code;
      reportProjectFilter.appendChild(option);
    });
    reportProjectFilter.value = prevProject && [...reportProjectFilter.options].some((o) => o.value === prevProject) ? prevProject : "";
  }
  if (reportRoleFilter) {
    reportRoleFilter.innerHTML = '<option value="">Tutte le mansioni</option>';
    const roles = [...new Set((state.standard_roles || [])
      .concat((state.demands || []).map((d) => d.role))
      .filter(Boolean))]
      .sort((a, b) => String(a).localeCompare(String(b)));
    roles.forEach((role) => {
      const option = document.createElement("option");
      option.value = role;
      option.textContent = role;
      reportRoleFilter.appendChild(option);
    });
    reportRoleFilter.value = prevRole && [...reportRoleFilter.options].some((o) => o.value === prevRole) ? prevRole : "";
  }
}

function renderReportCardGrid() {
  [
    reportDemandOverviewCard,
    reportDemandRoleOverviewCard,
    reportUtilizationOverviewCard,
    reportCustomInsightOpen,
    reportIndicatorAllocation,
    reportIndicatorCoverage,
  ].filter(Boolean).forEach((card) => {
    const preset = card.dataset.reportCard;
    const active =
      preset === reportDetailPresetId ||
      (preset === "workforce_alloc_meter" && reportDetailPresetId === "workforce_utilization") ||
      (preset === "coverage_meter" && (reportDetailPresetId === "demand_allocated" || reportDetailPresetId === "demand_allocated_role"));
    card.classList.toggle("active", active);
  });
  if (reportCustomPanel) reportCustomPanel.hidden = reportDetailPresetId !== "custom_insight";
  if (reportDetailArea) reportDetailArea.hidden = !reportDetailOpen || reportDetailPresetId === "custom_insight";
}

function renderSavedReportViews() {
  if (!savedReportViews) return;
  const views = getSavedReportViews();
  if (!views.length) {
    savedReportViews.innerHTML = '<div class="empty-note">Nessuna vista salvata.</div>';
    return;
  }
  savedReportViews.innerHTML = views.map((view, idx) => `
    <div class="saved-report-row">
      <button class="saved-report-open" data-saved-report-open="${idx}">${escapeHtml(view.name || `Vista ${idx + 1}`)}</button>
      <button class="saved-report-delete secondary-btn" data-saved-report-delete="${idx}">Elimina</button>
    </div>
  `).join("");
  savedReportViews.querySelectorAll("[data-saved-report-open]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const view = views[parseIntOr(btn.dataset.savedReportOpen, -1)];
      if (!view?.config) return;
      reportBuilderState = { ...defaultReportBuilderState(), ...view.config };
      reportDetailPresetId = view.config.detailPresetId || view.config.presetId || "demand_allocated";
      reportDetailOpen = true;
      syncReportBuilderControls();
      saveReportBuilderState();
      renderReportCardGrid();
      renderReportBuilder();
    });
  });
  savedReportViews.querySelectorAll("[data-saved-report-delete]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const idx = parseIntOr(btn.dataset.savedReportDelete, -1);
      const next = views.filter((_, i) => i !== idx);
      saveSavedReportViews(next);
      renderSavedReportViews();
    });
  });
}

function activateReportPreset(presetId) {
  const preset = REPORT_PRESETS.find((item) => item.id === presetId);
  if (!preset) return;
  reportDetailPresetId = preset.id;
  renderReportBuilder();
}

function reportScopeProjects() {
  const resolveScopeCode = (code) => {
    const normalized = String(code || "").trim();
    if (!normalized) return "";
    if (isOverallProjectCode(normalized)) return normalized;
    const owner = getOwningOverall(normalized);
    return owner ? (owner.name || owner.id || normalized) : normalized;
  };
  if (reportBuilderState.aggregateScope === "OVERALL_CANTIERI") {
    return new Set(
      reportEligibleProjects()
        .filter((p) => p.code !== "OVERALL OFFICINA")
        .map((p) => resolveScopeCode(p.code))
        .filter(Boolean)
    );
  }
  if (reportBuilderState.aggregateScope === "OVERALL_GENERALE") {
    return new Set(reportEligibleProjects().map((p) => resolveScopeCode(p.code)).filter(Boolean));
  }
  if (reportBuilderState.projectFilter) return new Set([resolveScopeCode(reportBuilderState.projectFilter)].filter(Boolean));
  return new Set(reportEligibleProjects().map((p) => resolveScopeCode(p.code)).filter(Boolean));
}

function groupWeeks(weeks, scale) {
  if (scale === "2W") {
    const groups = [];
    for (let i = 0; i < weeks.length; i += 2) {
      const slice = weeks.slice(i, i + 2);
      groups.push({ label: slice.length > 1 ? `W${slice[0]}-W${slice[slice.length - 1]}` : `W${slice[0]}`, weeks: slice });
    }
    return groups;
  }
  if (scale === "1M" || scale === "3M") {
    const byMonth = new Map();
    weeks.forEach((week) => {
      const d = startOfWeek(week);
      const key = `${d.getFullYear()}-${d.getMonth() + 1}`;
      if (!byMonth.has(key)) {
        byMonth.set(key, { label: d.toLocaleDateString("it-IT", { month: "short" }).replace(".", ""), weeks: [] });
      }
      byMonth.get(key).weeks.push(week);
    });
    const months = [...byMonth.values()];
    if (scale === "1M") return months;
    const groups = [];
    for (let i = 0; i < months.length; i += 3) {
      const slice = months.slice(i, i + 3);
      groups.push({
        label: slice.map((m) => m.label).join("/"),
        weeks: slice.flatMap((m) => m.weeks),
      });
    }
    return groups;
  }
  return weeks.map((week) => ({ label: `W${week}`, weeks: [week] }));
}

function activeResourcesByWeek() {
  const map = new Map();
  weeksForRange(1, 52).forEach((week) => {
    const count = state.resources.filter((resource) => !isExternalResource(resource) && resourceActive(resource, week)).length;
    map.set(week, count);
  });
  return map;
}

function uniqueResourceKey(resourceId, resourceName = "") {
  return resourceId ? `id:${resourceId}` : `name:${String(resourceName || "").trim().toUpperCase()}`;
}

function countDistinctAllocationsForWeek(week, projectCodes = null) {
  const keys = new Set();
  state.allocations.forEach((allocation) => {
    const inWeek = Number(allocation.week_from) <= week && Number(allocation.week_to) >= week;
    if (!inWeek) return;
    const scopedProject = plannerProjectMeta(allocation.project_id, allocation.project).project;
    if (projectCodes && !projectCodes.has(scopedProject)) return;
    keys.add(uniqueResourceKey(allocation.resource_id, allocation.resource));
  });
  return keys.size;
}

function countUnavailableForWeek(week) {
  const keys = new Set();
  (state.unavailability || []).forEach((entry) => {
    if (Number(entry.week_from) <= week && Number(entry.week_to) >= week) {
      const resource = state.resources.find((r) => Number(r.id) === Number(entry.resource_id));
      if (!resource || !resourceActive(resource, week)) return;
      if (isExternalResource(resource)) return;
      keys.add(uniqueResourceKey(entry.resource_id, resource?.name || entry.resource));
    }
  });
  return keys.size;
}

function countDistinctResourcesForWeekByActualRole(week, projectCodes = null, role = "") {
  const keys = new Set();
  state.allocations.forEach((allocation) => {
    const inWeek = Number(allocation.week_from) <= Number(week) && Number(allocation.week_to) >= Number(week);
    if (!inWeek) return;
    const scopedProject = plannerProjectMeta(allocation.project_id, allocation.project).project;
    if (projectCodes && !projectCodes.has(scopedProject)) return;
    const resource = state.resources.find((item) => Number(item.id) === Number(allocation.resource_id));
    if (!resource) return;
    if (role && !roleMatches(resource, role)) return;
    keys.add(uniqueResourceKey(allocation.resource_id, allocation.resource));
  });
  return keys.size;
}

function countAllocationsForWeek(week, projectCodes = null) {
  return state.allocations.reduce((total, allocation) => {
    const inWeek = Number(allocation.week_from) <= Number(week) && Number(allocation.week_to) >= Number(week);
    if (!inWeek) return total;
    const scopedProject = plannerProjectMeta(allocation.project_id, allocation.project).project;
    if (projectCodes && !projectCodes.has(scopedProject)) return total;
    return total + allocationEffectiveWeightForWeek(allocation, week);
  }, 0);
}

function summarizeDemandCoverageForWeek(week, scopeProjects, roleFilter = "", modeOverride = "") {
  const mode = modeOverride || reportBuilderState.mode || "operativa";
  if (mode === "mansione") {
    const rows = computeDemandOnlyRows("", roleFilter, true).filter((row) => scopeProjects.has(row.project));
    let required = 0;
    let coherent = 0;
    let surplus = 0;
    let missing = 0;
    rows.forEach((row) => {
      const weekCell = row.weeks.find((item) => Number(item.week) === Number(week));
      const rowRequired = Number(weekCell?.required || 0);
      if (!rowRequired) return;
      required += rowRequired;
      const allocations = state.allocations.filter((allocation) =>
        plannerProjectMeta(allocation.project_id, allocation.project).project === row.project &&
        allocation.role === row.role &&
        Number(allocation.week_from) <= Number(week) &&
        Number(allocation.week_to) >= Number(week)
      );
      let coherentRow = 0;
      allocations.forEach((allocation) => {
        const resource = state.resources.find((item) => Number(item.id) === Number(allocation.resource_id));
        if (resource && roleMatches(resource, row.role)) {
          coherentRow += allocationEffectiveWeightForWeek(allocation, week);
        }
      });
      coherent += Math.min(coherentRow, rowRequired);
      surplus += Math.max(coherentRow - rowRequired, 0);
      missing += Math.max(rowRequired - coherentRow, 0);
    });
    return {
      required,
      coherent,
      offrole: 0,
      surplus,
      missing,
      allocated: coherent,
    };
  }

  const rows = computeDemandOnlyRows("", roleFilter, true).filter((row) => scopeProjects.has(row.project));
  let required = 0;
  let coherent = 0;
  let offrole = 0;

  rows.forEach((row) => {
    const weekCell = row.weeks.find((item) => Number(item.week) === Number(week));
    const rowRequired = Number(weekCell?.required || 0);
    required += rowRequired;

    const allocations = state.allocations.filter((allocation) =>
      plannerProjectMeta(allocation.project_id, allocation.project).project === row.project &&
      allocation.role === row.role &&
      Number(allocation.week_from) <= Number(week) &&
      Number(allocation.week_to) >= Number(week)
    );

    allocations.forEach((allocation) => {
      const resource = state.resources.find((item) => Number(item.id) === Number(allocation.resource_id));
      const weight = allocationEffectiveWeightForWeek(allocation, week);
      if (resource && roleMatches(resource, row.role)) {
        coherent += weight;
      } else {
        offrole += weight;
      }
    });
  });

  const coveredDemand = coherent + offrole;
  return {
    required,
    coherent,
    offrole,
    surplus: 0,
    missing: Math.max(required - coveredDemand, 0),
    allocated: coveredDemand,
  };
}

function computeDemandAllocatedSeries(modeOverride = "operativa") {
  const activeMap = activeResourcesByWeek();
  const weekGroups = groupWeeks(weeksForRange(), reportBuilderState.scale);
  const scopeProjects = reportScopeProjects();
  const roleFilter = reportBuilderState.roleFilter;
  const byGroup = weekGroups.map((group) => {
    const weekly = group.weeks.map((week) => summarizeDemandCoverageForWeek(week, scopeProjects, roleFilter, modeOverride));
    const activeWeekly = group.weeks.map((week) => activeMap.get(week) || 0);
    const unavailableWeekly = group.weeks.map((week) => countUnavailableForWeek(week));
    const divisor = Math.max(group.weeks.length, 1);
    const sums = weekly.reduce((acc, item) => {
      acc.required += item.required;
      acc.coherent += item.coherent;
      acc.offrole += item.offrole;
      acc.missing += item.missing;
      acc.surplus += item.surplus || 0;
      acc.allocated += item.allocated;
      return acc;
    }, { required: 0, coherent: 0, offrole: 0, missing: 0, surplus: 0, allocated: 0 });
    const active = Math.round(activeWeekly.reduce((a, b) => a + b, 0) / Math.max(activeWeekly.length, 1));
    const unavailable = Math.round(unavailableWeekly.reduce((a, b) => a + b, 0) / Math.max(unavailableWeekly.length, 1));
    const allocatedInternal = Math.min(Math.round(sums.allocated / divisor), Math.max(active - unavailable, 0));
    const notAllocatedInternal = Math.max(active - unavailable - allocatedInternal, 0);
    const toFind = Math.max(Math.round(sums.required / divisor) - allocatedInternal, 0);
    return {
      label: group.label,
      required: Math.round(sums.required / divisor),
      coherent: Math.round(sums.coherent / divisor),
      offrole: Math.round(sums.offrole / divisor),
      missing: Math.round(sums.missing / divisor),
      surplus: Math.round(sums.surplus / divisor),
      allocated: Math.round(sums.allocated / divisor),
      active,
      unavailable,
      notAllocatedInternal,
      toFind,
    };
  });
  return byGroup;
}

function computeUtilizationSeries() {
  const activeMap = activeResourcesByWeek();
  const weekGroups = groupWeeks(weeksForRange(), reportBuilderState.scale);
  const scopeProjects = reportScopeProjects();
  return weekGroups.map((group) => {
    const weekly = group.weeks.map((week) => {
      const totalInternalActive = Number(activeMap.get(week) || 0);
      const totalInternalUnavailable = Number(countUnavailableForWeek(week) || 0);
      const totalInternalAllocable = Math.max(totalInternalActive - totalInternalUnavailable, 0);
      const totalInternalAllocated = state.allocations.reduce((acc, allocation) => {
        const inWeek = Number(allocation.week_from) <= Number(week) && Number(allocation.week_to) >= Number(week);
        if (!inWeek) return acc;
        const scopedProject = plannerProjectMeta(allocation.project_id, allocation.project).project;
        if (scopeProjects && !scopeProjects.has(scopedProject)) return acc;
        const resource = state.resources.find((r) => Number(r.id) === Number(allocation.resource_id));
        if (!resource || isExternalResource(resource)) return acc;
        if (!resourceActive(resource, week)) return acc;
        return acc + allocationEffectiveWeightForWeek(allocation, week);
      }, 0);
      const totalExternalAllocated = state.allocations.reduce((acc, allocation) => {
        const inWeek = Number(allocation.week_from) <= Number(week) && Number(allocation.week_to) >= Number(week);
        if (!inWeek) return acc;
        const scopedProject = plannerProjectMeta(allocation.project_id, allocation.project).project;
        if (scopeProjects && !scopeProjects.has(scopedProject)) return acc;
        const resource = state.resources.find((r) => Number(r.id) === Number(allocation.resource_id));
        if (!resource || !isExternalResource(resource)) return acc;
        return acc + 1;
      }, 0);
      const totalInternalNotAllocated = Math.max(totalInternalAllocable - totalInternalAllocated, 0);
      const totalCoverage = totalInternalAllocated + totalExternalAllocated;
      return {
        totalInternalActive,
        totalInternalUnavailable,
        totalInternalAllocable,
        totalInternalAllocated,
        totalInternalNotAllocated,
        totalExternalAllocated,
        totalCoverage,
      };
    });
    const divisor = Math.max(weekly.length, 1);
    const totalInternalActive = Math.round(weekly.reduce((acc, item) => acc + item.totalInternalActive, 0) / divisor);
    const totalInternalUnavailable = Math.round(weekly.reduce((acc, item) => acc + item.totalInternalUnavailable, 0) / divisor);
    const totalInternalAllocable = Math.round(weekly.reduce((acc, item) => acc + item.totalInternalAllocable, 0) / divisor);
    const totalInternalAllocated = Math.round(weekly.reduce((acc, item) => acc + item.totalInternalAllocated, 0) / divisor);
    const totalInternalNotAllocated = Math.round(weekly.reduce((acc, item) => acc + item.totalInternalNotAllocated, 0) / divisor);
    const totalExternalAllocated = Math.round(weekly.reduce((acc, item) => acc + item.totalExternalAllocated, 0) / divisor);
    const totalCoverage = Math.round(weekly.reduce((acc, item) => acc + item.totalCoverage, 0) / divisor);
    return {
      label: group.label,
      active: totalInternalAllocable,
      allocated: totalInternalAllocated,
      unallocated: totalInternalNotAllocated,
      totalInternalActive,
      totalInternalUnavailable,
      totalInternalAllocable,
      totalInternalAllocated,
      totalInternalNotAllocated,
      totalExternalAllocated,
      totalCoverage,
    };
  });
}

function reportScopeLabel() {
  if (reportBuilderState.aggregateScope === "OVERALL_CANTIERI") return "OVERALL CANTIERI";
  if (reportBuilderState.aggregateScope === "OVERALL_GENERALE") return "OVERALL GENERALE AZIENDA";
  if (reportBuilderState.projectFilter) return reportBuilderState.projectFilter;
  return "Tutte le commesse";
}

function reportPresetTitle() {
  switch (reportDetailPresetId) {
    case "demand_allocated_role":
      return "Fabbisogno / Allocazioni per Mansione";
    case "workforce_utilization":
      return "Utilizzo Personale in Forza";
    case "custom_insight":
      return "Insight Personalizzato";
    default:
      return "Fabbisogno / Allocazioni";
  }
}

function renderReportIndicators(demandSeries, utilizationSeries) {
  const groups = Math.max(demandSeries.length, 1);
  const demandTotals = demandSeries.reduce((acc, item) => {
    acc.required += Number(item.required || 0);
    acc.missing += Number(item.missing || 0);
    return acc;
  }, { required: 0, missing: 0 });
  const utilTotals = utilizationSeries.reduce((acc, item) => {
    acc.active += Number(item.active || 0);
    acc.allocated += Number(item.allocated || 0);
    acc.unallocated += Number(item.unallocated || 0);
    return acc;
  }, { active: 0, allocated: 0, unallocated: 0 });
  const meanAllocated = Math.round(utilTotals.allocated / Math.max(utilizationSeries.length, 1));
  const meanActive = Math.round(utilTotals.active / Math.max(utilizationSeries.length, 1));
  const meanRequired = Math.round(demandTotals.required / groups);
  const meanMissing = Math.round(demandTotals.missing / groups);
  const allocPct = meanActive ? Math.round((meanAllocated / meanActive) * 100) : 0;
  const coveredDemand = Math.max(meanRequired - meanMissing, 0);
  const coveragePct = meanRequired ? Math.round((coveredDemand / meanRequired) * 100) : 100;
  const scaleLabel = reportBuilderState.scale === "1WEEK" ? "" : "Media nel periodo: ";
  const allocationDetail = reportBuilderState.scale === "1WEEK"
    ? `${meanAllocated} allocate su ${meanActive} in forza`
    : `${scaleLabel}${meanAllocated} allocate su ${meanActive} in forza`;
  const coverageDetail = reportBuilderState.scale === "1WEEK"
    ? `${meanMissing} mancanti su ${meanRequired} richieste`
    : `${scaleLabel}${meanMissing} mancanti su ${meanRequired} richieste`;
  const allocationTitle = `Indicatore allocazione personale.\n${allocationDetail}.\nNon allocate medie nel periodo: ${Math.round(utilTotals.unallocated / Math.max(utilizationSeries.length, 1))}.`;
  const coverageTitle = `Indicatore copertura mansioni.\n${coverageDetail}.\nCopertura media: ${coveredDemand} su ${meanRequired}.`;

  if (reportIndicatorAllocationFill) reportIndicatorAllocationFill.style.width = `${Math.max(0, Math.min(100, allocPct))}%`;
  if (reportIndicatorAllocationValue) reportIndicatorAllocationValue.textContent = `${allocPct}%`;
  if (reportIndicatorAllocationDetail) reportIndicatorAllocationDetail.textContent = allocationDetail;
  if (reportIndicatorAllocation) reportIndicatorAllocation.title = allocationTitle;

  if (reportIndicatorCoverageFill) reportIndicatorCoverageFill.style.width = `${Math.max(0, Math.min(100, coveragePct))}%`;
  if (reportIndicatorCoverageValue) reportIndicatorCoverageValue.textContent = `${coveragePct}%`;
  if (reportIndicatorCoverageDetail) reportIndicatorCoverageDetail.textContent = coverageDetail;
  if (reportIndicatorCoverage) reportIndicatorCoverage.title = coverageTitle;
}

function renderReportSummaryCards(cards) {
  if (!reportSummaryCards) return;
  reportSummaryCards.innerHTML = cards.map((card) => `
    <div class="report-summary-card">
      <div class="report-summary-label">${escapeHtml(card.label)}</div>
      <div class="report-summary-value">${escapeHtml(card.value)}</div>
    </div>
  `).join("");
}

function renderReportLegendItems(items) {
  if (!reportLegend) return;
  reportLegend.innerHTML = items.map((item) => `
    <span class="report-legend-item"><span class="swatch ${item.className}"></span>${escapeHtml(item.label)}</span>
  `).join("");
}

function renderLegendInto(target, items) {
  if (!target) return;
  target.innerHTML = items.map((item) => `
    <span class="report-legend-item"><span class="swatch ${item.className}"></span>${escapeHtml(item.label)}</span>
  `).join("");
}

function buildDemandAllocatedChart(series, options = {}) {
  const compact = !!options.compact;
  const mode = options.mode || reportBuilderState.mode || "operativa";
  const maxRequired = Math.max(
    1,
    ...series.map((item) => Math.max(item.required, item.coherent + item.offrole + item.missing + (item.surplus || 0)))
  );
  const width = compact ? Math.max(540, series.length * 76) : Math.max(860, series.length * 118);
  const height = compact ? 250 : 420;
  const left = compact ? 40 : 60;
  const top = compact ? 18 : 24;
  const chartHeight = compact ? 155 : 280;
  const plotWidth = width - left - (compact ? 16 : 24);
  const colWidth = compact
    ? Math.min(42, Math.max(18, plotWidth / Math.max(series.length, 1) - 12))
    : Math.min(66, Math.max(28, plotWidth / Math.max(series.length, 1) - 20));
  const gap = series.length > 1 ? (plotWidth - colWidth * series.length) / (series.length - 1) : 0;
  const y = (value) => top + chartHeight - (value / maxRequired) * chartHeight;
  const bars = series.map((item, idx) => {
    const x = left + idx * (colWidth + gap);
    const coherentH = (item.coherent / maxRequired) * chartHeight;
    const offroleH = (item.offrole / maxRequired) * chartHeight;
    const missingH = (item.missing / maxRequired) * chartHeight;
    const surplusH = ((item.surplus || 0) / maxRequired) * chartHeight;
    const demandY = y(item.required);
    const managedMode = mode === "operativa";
    return `
      <g>
        <rect x="${x}" y="${y(item.coherent)}" width="${colWidth}" height="${Math.max(coherentH, 0)}" rx="8" fill="#24a36a"></rect>
        ${managedMode && item.offrole > 0 ? `<rect x="${x}" y="${y(item.coherent + item.offrole)}" width="${colWidth}" height="${Math.max(offroleH, 0)}" rx="0" fill="#dd8a1e"></rect>` : ""}
        ${!managedMode && item.surplus > 0 ? `<rect x="${x}" y="${y(item.coherent + item.surplus)}" width="${colWidth}" height="${Math.max(surplusH, 0)}" rx="0" fill="#3f83f8"></rect>` : ""}
        ${item.missing > 0 ? `<rect x="${x}" y="${y(item.coherent + item.offrole + item.missing)}" width="${colWidth}" height="${Math.max(missingH, 0)}" rx="0" fill="#d84d43"></rect>` : ""}
        <line x1="${x - 6}" y1="${demandY}" x2="${x + colWidth + 6}" y2="${demandY}" stroke="#111827" stroke-width="2"></line>
        <circle cx="${x + colWidth / 2}" cy="${demandY}" r="4" fill="#111827"></circle>
        ${!compact && item.coherent > 0 ? `<text x="${x + colWidth / 2}" y="${y(item.coherent / 2) + 5}" text-anchor="middle" class="chart-value on-dark">${item.coherent}</text>` : ""}
        ${!compact && managedMode && item.offrole > 0 ? `<text x="${x + colWidth / 2}" y="${y(item.coherent + item.offrole / 2) + 5}" text-anchor="middle" class="chart-value on-dark">${item.offrole}</text>` : ""}
        ${!compact && !managedMode && item.surplus > 0 ? `<text x="${x + colWidth / 2}" y="${y(item.coherent + item.surplus / 2) + 5}" text-anchor="middle" class="chart-value on-dark">${item.surplus}</text>` : ""}
        ${!compact && item.missing > 0 ? `<text x="${x + colWidth / 2}" y="${y(item.coherent + item.offrole + item.missing / 2) + 5}" text-anchor="middle" class="chart-value on-dark">${item.missing}</text>` : ""}
        <text x="${x + colWidth / 2}" y="${top + chartHeight + (compact ? 20 : 28)}" text-anchor="middle" class="chart-axis-label${compact ? " compact" : ""}">${item.label}</text>
        ${!compact ? `<text x="${x + colWidth / 2}" y="${demandY - 10}" text-anchor="middle" class="chart-demand-label">${item.required}</text>` : ""}
      </g>
    `;
  }).join("");
  const grid = Array.from({ length: 5 }, (_, i) => {
    const value = Math.round((maxRequired / 4) * i);
    const yy = y(value);
    return `
      <line x1="${left}" y1="${yy}" x2="${width - 20}" y2="${yy}" stroke="#e6edf5" stroke-width="1"></line>
      <text x="${left - 10}" y="${yy + 4}" text-anchor="end" class="chart-tick">${value}</text>
    `;
  }).join("");
  const html = `
    <div class="report-chart-scroll">
      <svg viewBox="0 0 ${width} ${height}" class="report-svg" aria-label="Grafico fabbisogno vs allocato">
        ${grid}
        <line x1="${left}" y1="${top + chartHeight}" x2="${width - 20}" y2="${top + chartHeight}" stroke="#95a5bb"></line>
        ${bars}
      </svg>
    </div>
  `;
  const legendItems = mode === "operativa"
    ? [
        { label: "Allocati coerenti", className: "coherent" },
        { label: "Allocati fuori mansione", className: "offrole" },
        { label: "Mancanti", className: "missing" },
        { label: "Fabbisogno", className: "demandline" },
      ]
    : [
        { label: "Allocati coerenti", className: "coherent" },
        { label: "Surplus reale", className: "surplus" },
        { label: "Mancanti", className: "missing" },
        { label: "Fabbisogno", className: "demandline" },
      ];
  const totals = series.reduce((acc, item) => {
    acc.required += item.required;
    acc.allocated += item.allocated;
    acc.missing += item.missing;
    acc.coherent += item.coherent || 0;
    acc.offrole += item.offrole || 0;
    acc.surplus += item.surplus || 0;
    acc.active += item.active || 0;
    acc.unavailable += item.unavailable || 0;
    acc.notAllocatedInternal += item.notAllocatedInternal || 0;
    acc.toFind += item.toFind || 0;
    return acc;
  }, { required: 0, allocated: 0, missing: 0, coherent: 0, offrole: 0, surplus: 0, active: 0, unavailable: 0, notAllocatedInternal: 0, toFind: 0 });
  const cardSuffix = reportBuilderState.scale === "1M"
    ? " medio/mese"
    : reportBuilderState.scale === "3M"
      ? " medio/periodo"
      : reportBuilderState.scale === "2W"
        ? " medio/2 sett."
        : " medio/sett.";
  const count = Math.max(series.length, 1);
  if (mode === "operativa") {
    return {
      html,
      legendItems,
      summaryCards: [
      { label: `Fabbisogno${cardSuffix}`, value: String(Math.round(totals.required / count)) },
      { label: `Mie allocate${cardSuffix}`, value: String(Math.round(totals.allocated / count)) },
      { label: `Mie non allocate${cardSuffix}`, value: String(Math.round(totals.notAllocatedInternal / count)) },
      { label: `Mie indisponibili${cardSuffix}`, value: String(Math.round(totals.unavailable / count)) },
      { label: `Da trovare${cardSuffix}`, value: String(Math.round(totals.toFind / count)) },
      { label: `In forza${cardSuffix}`, value: String(Math.round(totals.active / count)) },
      ],
    };
  } else {
    return {
      html,
      legendItems,
      summaryCards: [
      { label: `Fabbisogno${cardSuffix}`, value: String(Math.round(totals.required / count)) },
      { label: `Coerenti reali${cardSuffix}`, value: String(Math.round(totals.coherent / count)) },
      { label: `Surplus reale${cardSuffix}`, value: String(Math.round(totals.surplus / count)) },
      { label: `Mancanti${cardSuffix}`, value: String(Math.round(totals.missing / count)) },
      ],
    };
  }
}

function renderDemandAllocatedChart(series, modeOverride = "") {
  if (!reportPreview) return;
  const built = buildDemandAllocatedChart(series, modeOverride ? { mode: modeOverride } : undefined);
  reportPreview.innerHTML = built.html;
  renderReportLegendItems(built.legendItems);
  renderReportSummaryCards(built.summaryCards);
}

function buildUtilizationChart(series, options = {}) {
  const compact = !!options.compact;
  const maxValue = Math.max(1, ...series.map((item) => Math.max(item.active, item.allocated, item.unallocated)));
  const width = compact ? Math.max(540, series.length * 76) : Math.max(860, series.length * 118);
  const height = compact ? 250 : 420;
  const left = compact ? 40 : 60;
  const top = compact ? 20 : 30;
  const chartHeight = compact ? 170 : 260;
  const zeroY = top + chartHeight / 2;
  const plotWidth = width - left - (compact ? 16 : 24);
  const colWidth = compact
    ? Math.min(40, Math.max(18, plotWidth / Math.max(series.length, 1) - 12))
    : Math.min(58, Math.max(26, plotWidth / Math.max(series.length, 1) - 20));
  const gap = series.length > 1 ? (plotWidth - colWidth * series.length) / (series.length - 1) : 0;
  const scale = (value) => (value / maxValue) * (chartHeight / 2 - 16);
  const bars = series.map((item, idx) => {
    const x = left + idx * (colWidth + gap);
    const allocatedH = scale(item.allocated);
    const unallocatedH = scale(item.unallocated);
    return `
      <g>
        <rect x="${x}" y="${zeroY - allocatedH}" width="${colWidth}" height="${allocatedH}" rx="8" fill="#1f7fd6"></rect>
        <rect x="${x}" y="${zeroY}" width="${colWidth}" height="${unallocatedH}" rx="8" fill="#d65353"></rect>
        ${!compact && item.allocated > 0 ? `<text x="${x + colWidth / 2}" y="${zeroY - allocatedH - 8}" text-anchor="middle" class="chart-value">${item.allocated}</text>` : ""}
        ${!compact && item.unallocated > 0 ? `<text x="${x + colWidth / 2}" y="${zeroY + unallocatedH + 16}" text-anchor="middle" class="chart-value">${item.unallocated}</text>` : ""}
        <text x="${x + colWidth / 2}" y="${height - (compact ? 22 : 28)}" text-anchor="middle" class="chart-axis-label${compact ? " compact" : ""}">${item.label}</text>
      </g>
    `;
  }).join("");
  const ticks = Array.from({ length: 5 }, (_, i) => {
    const value = Math.round((maxValue / 4) * i);
    return `
      <line x1="${left}" y1="${zeroY - scale(value)}" x2="${width - 20}" y2="${zeroY - scale(value)}" stroke="#eef3f8"></line>
      <line x1="${left}" y1="${zeroY + scale(value)}" x2="${width - 20}" y2="${zeroY + scale(value)}" stroke="#eef3f8"></line>
      <text x="${left - 10}" y="${zeroY - scale(value) + 4}" text-anchor="end" class="chart-tick">${value}</text>
      ${value > 0 ? `<text x="${left - 10}" y="${zeroY + scale(value) + 4}" text-anchor="end" class="chart-tick">-${value}</text>` : ""}
    `;
  }).join("");
  const html = `
    <div class="report-chart-scroll">
      <svg viewBox="0 0 ${width} ${height}" class="report-svg" aria-label="Grafico utilizzo personale in forza">
        ${ticks}
        <line x1="${left}" y1="${zeroY}" x2="${width - 20}" y2="${zeroY}" stroke="#7d91ab" stroke-width="2"></line>
        ${bars}
      </svg>
    </div>
  `;
  const legendItems = [
    { label: "Allocati", className: "allocated" },
    { label: "Non allocati", className: "unallocated" },
  ];
  const totals = series.reduce((acc, item) => {
    acc.active += item.active;
    acc.allocated += item.allocated;
    acc.unallocated += item.unallocated;
    acc.external += Number(item.totalExternalAllocated || 0);
    acc.coverage += Number(item.totalCoverage || 0);
    return acc;
  }, { active: 0, allocated: 0, unallocated: 0, external: 0, coverage: 0 });
  return {
    html,
    legendItems,
    summaryCards: [
      { label: "In forza allocabile media", value: String(Math.round(totals.active / Math.max(series.length, 1))) },
      { label: "Interni allocati medi", value: String(Math.round(totals.allocated / Math.max(series.length, 1))) },
      { label: "Interni non allocati medi", value: String(Math.round(totals.unallocated / Math.max(series.length, 1))) },
      { label: "Esterni allocati medi", value: String(Math.round(totals.external / Math.max(series.length, 1))) },
      { label: "Copertura totale media", value: String(Math.round(totals.coverage / Math.max(series.length, 1))) },
    ],
  };
}

function renderUtilizationChart(series) {
  if (!reportPreview) return;
  const built = buildUtilizationChart(series);
  reportPreview.innerHTML = built.html;
  renderReportLegendItems(built.legendItems);
  renderReportSummaryCards(built.summaryCards);
}

function renderOverviewDemandCards(demandSeriesOperativa, demandSeriesMansione) {
  if (reportDemandOverview) {
    const built = buildDemandAllocatedChart(demandSeriesOperativa, { compact: true, mode: "operativa" });
    reportDemandOverview.innerHTML = built.html;
    renderLegendInto(reportDemandOverviewLegend, built.legendItems);
    if (reportDemandOverviewSubtitle) {
      reportDemandOverviewSubtitle.textContent = `${reportScopeLabel()} | W${reportBuilderState.weekFrom}-W${reportBuilderState.weekTo} | ${reportBuilderState.scale}`;
    }
  }
  if (reportDemandRoleOverview) {
    const built = buildDemandAllocatedChart(demandSeriesMansione, { compact: true, mode: "mansione" });
    reportDemandRoleOverview.innerHTML = built.html;
    renderLegendInto(reportDemandRoleOverviewLegend, built.legendItems);
    if (reportDemandRoleOverviewSubtitle) {
      reportDemandRoleOverviewSubtitle.textContent = `${reportRoleFilter?.value || "Tutte le mansioni"} | profilo richiesto`;
    }
  }
}

function renderOverviewUtilizationCard(utilizationSeries) {
  if (!reportUtilizationOverview) return;
  const built = buildUtilizationChart(utilizationSeries, { compact: true });
  reportUtilizationOverview.innerHTML = built.html;
  renderLegendInto(reportUtilizationOverviewLegend, built.legendItems);
  if (reportUtilizationOverviewSubtitle) {
    reportUtilizationOverviewSubtitle.textContent = `${reportScopeLabel()} | utilizzo forza lavoro`;
  }
}

function renderReportBuilder() {
  if (!reportPreview) return;
  readReportBuilderControls();
  renderReportCardGrid();
  renderSavedReportViews();
  const scopeLabel = reportScopeLabel();
  const demandSeriesOperativa = computeDemandAllocatedSeries("operativa");
  const demandSeriesMansione = computeDemandAllocatedSeries("mansione");
  const utilizationSeries = computeUtilizationSeries();
  renderReportIndicators(demandSeriesOperativa, utilizationSeries);
  renderOverviewDemandCards(demandSeriesOperativa, demandSeriesMansione);
  renderOverviewUtilizationCard(utilizationSeries);
  if (!reportDetailOpen || reportDetailPresetId === "custom_insight") {
    if (reportPreview) reportPreview.innerHTML = "";
    if (reportLegend) reportLegend.innerHTML = "";
    if (reportSummaryCards) reportSummaryCards.innerHTML = "";
    if (reportDetailArea) reportDetailArea.hidden = true;
  } else {
    reportPreviewTitle.textContent = reportPresetTitle();
    reportPreviewSubtitle.textContent = `${scopeLabel} | W${reportBuilderState.weekFrom}-W${reportBuilderState.weekTo} | ${reportBuilderState.scale}${reportBuilderState.roleFilter ? ` | ${reportBuilderState.roleFilter}` : ""}`;
    if (reportDetailPresetId === "workforce_utilization") {
      renderUtilizationChart(utilizationSeries);
    } else {
      const detailSeries = reportDetailPresetId === "demand_allocated_role" ? demandSeriesMansione : demandSeriesOperativa;
      const detailMode = reportDetailPresetId === "demand_allocated_role" ? "mansione" : "operativa";
      renderDemandAllocatedChart(detailSeries, detailMode);
    }
    if (reportDetailArea) reportDetailArea.hidden = false;
  }
}

function formatPct(value) {
  if (value === null || value === undefined || Number.isNaN(value)) return "n/a";
  const sign = value > 0 ? "+" : "";
  return `${sign}${value}%`;
}

function summarizeImportAnalysis(analysis) {
  if (!analysis) return "Nessuna analisi disponibile.";
  const lines = [];
  if (analysis.current_week) {
    lines.push(`Periodo considerato: da W${analysis.current_week}.`);
    lines.push("");
  }
  const unknownRoles = (analysis.unknown_roles || []).slice();
  if (unknownRoles.length) {
    lines.push("Mansioni non riconosciute:");
    unknownRoles.slice(0, 8).forEach((role) => lines.push(`- ${role}`));
    if (unknownRoles.length > 8) lines.push(`... e altre ${unknownRoles.length - 8}`);
    lines.push("");
  }
  const projectSummaries = (analysis.project_summaries || []).slice();
  const adjusted = projectSummaries.filter((p) => p.change_type && p.change_type !== "invariata");
  if (adjusted.length) {
    lines.push("Variazioni principali:");
    adjusted.slice(0, 8).forEach((item) => {
      const pieces = [];
      if (item.change_type === "allungata" && item.prev_week_max && item.new_week_max) {
        pieces.push(`allungata fino a W${item.new_week_max} (prima W${item.prev_week_max})`);
      } else if (item.change_type === "accorciata" && item.prev_week_max && item.new_week_max) {
        pieces.push(`accorciata: ora finisce a W${item.new_week_max} (prima W${item.prev_week_max})`);
      }
      if (item.change_type === "aumentata") {
        pieces.push("fabbisogno aumentato");
      } else if (item.change_type === "ridotta") {
        pieces.push("fabbisogno ridotto");
      }
      const detail = pieces.length ? pieces.join(", ") : item.change_type;
      lines.push(`- ${item.project}: ${detail}`);
    });
  }

  const conflicts = (analysis.allocation_conflicts || [])
    .filter((item) => !analysis.current_week || Number(item.week) >= Number(analysis.current_week))
    .slice()
    .sort((a, b) => Math.abs(b.delta || 0) - Math.abs(a.delta || 0));
  if (conflicts.length) {
    lines.push("");
    lines.push("Conflitti allocazioni vs nuovo fabbisogno:");
    conflicts.slice(0, 5).forEach((item) => {
      const tag = item.delta > 0 ? "over" : "under";
      lines.push(
        `- ${item.project} | ${item.role} | W${item.week}: allocati ${item.allocated} vs fabb ${item.demand} (${tag} ${item.delta}${item.pct_delta !== null ? `, ${formatPct(item.pct_delta)}` : ""})`
      );
    });
  }

  return lines.join("\n");
}

function getDemandPathOverride() {
  try {
    return window.localStorage.getItem("demand_path_override") || "";
  } catch {
    return "";
  }
}

function setDemandPathOverride(path) {
  try {
    if (path) {
      window.localStorage.setItem("demand_path_override", path);
    } else {
      window.localStorage.removeItem("demand_path_override");
    }
  } catch {
    // ignore storage issues
  }
}

function currentDemandPathValue() {
  const override = getDemandPathOverride();
  return state.current_demand_path || override || state.default_demand_path || "";
}

function promptDemandPath(currentPath) {
  const input = window.prompt("Percorso del file fabbisogno:", currentPath || "");
  if (input === null) return null;
  const trimmed = String(input || "").trim();
  if (trimmed) setDemandPathOverride(trimmed);
  return trimmed || currentPath || "";
}

function renderImportRoleMapping(unknownRoles) {
  if (!importUnknownRoles || !importUnknownRolesList) return;
  if (!unknownRoles || !unknownRoles.length) {
    importUnknownRoles.hidden = true;
    importUnknownRolesList.innerHTML = "";
    return;
  }
  const roles = (state.standard_roles || []).slice().sort();
  importUnknownRolesList.innerHTML = "";
  unknownRoles.forEach((role) => {
    const row = document.createElement("div");
    row.className = "import-unknown-list-row";
    row.innerHTML = `
      <div class="import-unknown-role">${escapeHtml(role)}</div>
      <div>
        <select class="import-role-select" data-role="${escapeHtml(role)}">
          <option value="">Seleziona...</option>
          <option value="__ignore__">Ignora questa mansione (non importare)</option>
          <option value="__create__">Crea nuova mansione: ${escapeHtml(role)}</option>
          ${roles.map((r) => `<option value="${escapeHtml(r)}">${escapeHtml(r)}</option>`).join("")}
        </select>
      </div>
    `;
    importUnknownRolesList.appendChild(row);
  });
  importUnknownRoles.hidden = false;
}

function collectImportRoleMapping(unknownRoles) {
  if (!unknownRoles || !unknownRoles.length) {
    return { roleMapping: {}, createRoles: [] };
  }
  const selects = importUnknownRolesList?.querySelectorAll(".import-role-select") || [];
  const roleMapping = {};
  const createRoles = [];
  for (const sel of selects) {
    const sourceRole = sel.dataset.role || "";
    const value = sel.value || "";
    if (!sourceRole) continue;
    if (!value) {
      return { error: "Seleziona una mansione per tutti i ruoli non riconosciuti." };
    }
    if (value === "__ignore__") {
      roleMapping[sourceRole] = "__ignore__";
      continue;
    }
    if (value === "__create__") {
      createRoles.push(sourceRole);
    } else {
      roleMapping[sourceRole] = value;
    }
  }
  return { roleMapping, createRoles };
}

function showImportPreviewModal(text, unknownRoles = []) {
  if (!importPreviewModal || !importPreviewText) {
    const fallback = window.confirm(text);
    return Promise.resolve({ confirmed: fallback, roleMapping: {}, createRoles: [] });
  }
  importPreviewText.textContent = text || "";
  renderImportRoleMapping(unknownRoles || []);
  importPreviewModal.hidden = false;
  return new Promise((resolve) => {
    const cleanup = (result) => {
      importPreviewModal.hidden = true;
      importPreviewConfirm?.removeEventListener("click", onConfirm);
      importPreviewCancel?.removeEventListener("click", onCancel);
      importPreviewClose?.removeEventListener("click", onCancel);
      resolve(result);
    };
    const onConfirm = () => {
      const mapping = collectImportRoleMapping(unknownRoles || []);
      if (mapping.error) {
        window.alert(mapping.error);
        return;
      }
      cleanup({ confirmed: true, roleMapping: mapping.roleMapping, createRoles: mapping.createRoles });
    };
    const onCancel = () => cleanup({ confirmed: false, roleMapping: {}, createRoles: [] });
    importPreviewConfirm?.addEventListener("click", onConfirm);
    importPreviewCancel?.addEventListener("click", onCancel);
    importPreviewClose?.addEventListener("click", onCancel);
  });
}

function buildDemandDiffMap(diffs = []) {
  const map = {};
  diffs.forEach((item) => {
    const key = `${item.project}||${item.role}||${item.week}`;
    map[key] = item;
  });
  return map;
}

function foglio2EditKey(name, sourceRole) {
  return `${String(name || "").trim().toUpperCase()}|${String(sourceRole || "").trim().toUpperCase()}`;
}

function loadFoglio2EditsFromStorage() {
  try {
    const raw = window.localStorage.getItem("foglio2_edits_v1");
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function saveFoglio2EditsToStorage() {
  try {
    window.localStorage.setItem("foglio2_edits_v1", JSON.stringify(state.foglio2_edits || {}));
  } catch {
    // ignore storage issues
  }
}

function ensureFoglio2Edits() {
  const analysis = state.foglio2_analysis;
  if (!analysis) return;
  const next = { ...loadFoglio2EditsFromStorage(), ...(state.foglio2_edits || {}) };
  [...(analysis.new_resources || []), ...(analysis.unknown_roles || [])].forEach((row) => {
    const key = foglio2EditKey(row.name, row.source_role);
    if (!next[key]) {
      next[key] = {
        suggested_role: row.master_role1 || row.suggested_role || "",
        hire_date: row.hire_date || "",
        end_date: row.end_date || "",
      };
    } else if (!next[key].suggested_role && (row.master_role1 || row.suggested_role)) {
      next[key].suggested_role = row.master_role1 || row.suggested_role || "";
    }
  });
  state.foglio2_edits = next;
  saveFoglio2EditsToStorage();
}

function foglio2RoleOptions(selected = "") {
  const roles = [...new Set((state.standard_roles || []).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  return [`<option value="">Da confermare</option>`]
    .concat(
      roles.map(
        (role) => `<option value="${escapeHtml(role)}" ${role === selected ? "selected" : ""}>${escapeHtml(role)}</option>`
      )
    )
    .join("");
}

function bindFoglio2EditFields() {
  document.querySelectorAll("[data-foglio2-key]").forEach((el) => {
    el.addEventListener("change", (ev) => {
      const key = ev.target.dataset.foglio2Key;
      const field = ev.target.dataset.foglio2Field;
      if (!key || !field) return;
      if (!state.foglio2_edits[key]) state.foglio2_edits[key] = {};
      state.foglio2_edits[key][field] = ev.target.value;
      saveFoglio2EditsToStorage();
      setStatus("Conferme Foglio2 aggiornate localmente.");
    });
  });
}

function bindMiniExpandButtons() {
  document.querySelectorAll(".mini-expand-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const targetId = btn.dataset.expandTarget;
      const target = document.getElementById(targetId);
      if (!target) return;
      const card = btn.closest(".mini-card");
      card?.classList.toggle("expanded");
      btn.textContent = card?.classList.contains("expanded") ? "Riduci" : "Ingrandisci";
    });
  });
}

function renderFoglio2Analysis() {
  const analysis = state.foglio2_analysis;
  if (!analysis) {
    foglio2Summary.textContent = "Nessuna analisi eseguita.";
    foglio2NewResources.innerHTML = "";
    foglio2CeasedResources.innerHTML = "";
    foglio2UnknownRoles.innerHTML = "";
    foglio2Proposals.innerHTML = "";
    return;
  }
  ensureFoglio2Edits();
  const s = analysis.summary || {};
  foglio2Summary.innerHTML = `
    <div><strong>Fonte:</strong> ${escapeHtml(analysis.source_path || "")}</div>
    <div><strong>Trovati nel master:</strong> ${s.matched_resources || 0}</div>
    <div><strong>Nuovi nominativi:</strong> ${s.new_resources || 0}</div>
    <div><strong>Cessati esclusi:</strong> ${s.ceased_excluded || 0}</div>
    <div><strong>Mansioni da confermare:</strong> ${s.unknown_roles || 0}</div>
    <div><strong>Proposte settimanali:</strong> ${s.proposal_rows || 0}</div>
    <div><strong>Nota:</strong> qui puoi modificare solo la proposta. Il salvataggio definitivo lo facciamo dopo la tua conferma.</div>
  `;
  foglio2NewResources.innerHTML = (analysis.new_resources || [])
    .slice(0, 40)
    .map((r) => {
      const key = foglio2EditKey(r.name, r.source_role);
      const edit = state.foglio2_edits[key] || {};
      return `
        <div class="mini-row mini-row-edit">
          <div class="mini-main">
            <strong>${escapeHtml(r.name)}</strong>
            <span>${escapeHtml(r.source_role || "-")}</span>
          </div>
          <div class="mini-edit-grid">
            <label><small>Mansione</small><select data-foglio2-key="${escapeHtml(key)}" data-foglio2-field="suggested_role">${foglio2RoleOptions(edit.suggested_role || "")}</select></label>
            <label><small>Assunzione</small><input data-foglio2-key="${escapeHtml(key)}" data-foglio2-field="hire_date" value="${escapeHtml(edit.hire_date || "")}" /></label>
            <label><small>Cessazione</small><input data-foglio2-key="${escapeHtml(key)}" data-foglio2-field="end_date" value="${escapeHtml(edit.end_date || "")}" /></label>
          </div>
        </div>`;
    })
    .join("") || `<div class="empty-box">Nessun nuovo nominativo attivo.</div>`;
  foglio2CeasedResources.innerHTML = (analysis.ceased_resources || [])
    .slice(0, 40)
    .map(
      (r) =>
        `<div class="mini-row"><strong>${escapeHtml(r.name)}</strong><span>${escapeHtml(
          r.note || r.end_date || "Cessato"
        )}</span></div>`
    )
    .join("") || `<div class="empty-box">Nessun cessato escluso.</div>`;
  foglio2UnknownRoles.innerHTML = (analysis.unknown_roles || [])
    .slice(0, 40)
    .map((r) => {
      const key = foglio2EditKey(r.name, r.source_role);
      const edit = state.foglio2_edits[key] || {};
      return `
        <div class="mini-row mini-row-edit is-pending">
          <div class="mini-main">
            <strong>${escapeHtml(r.name)}</strong>
            <span>${escapeHtml(r.source_role || "-")}</span>
          </div>
          <div class="mini-edit-grid single">
            <label><small>Mansione standard</small><select data-foglio2-key="${escapeHtml(key)}" data-foglio2-field="suggested_role">${foglio2RoleOptions(edit.suggested_role || "")}</select></label>
          </div>
        </div>`;
    })
    .join("") || `<div class="empty-box">Nessuna mansione da confermare.</div>`;
  foglio2Proposals.innerHTML = (analysis.proposals || [])
    .filter((r) => Number(r.week_to) >= getCurrentWeek())
    .slice(0, 120)
    .map(
      (r) => {
        const shownWeekFrom = Math.max(Number(r.week_from), getCurrentWeek());
        return (
        `<div class="mini-row ${r.matched_master ? "" : "is-pending"}"><strong>${escapeHtml(
          r.name
        )}</strong><span>${escapeHtml(r.project)} | W${shownWeekFrom}-W${r.week_to} | ${escapeHtml(
          r.suggested_role || r.source_role || "-"
        )}${r.matched_master ? "" : " | nuovo nome"}</span></div>`
        );
      }
    )
    .join("") || `<div class="empty-box">Nessuna proposta ricavata da Foglio2.</div>`;
  bindFoglio2EditFields();
  bindMiniExpandButtons();
}

async function analyzeFoglio2() {
  setStatus("Analisi Foglio2 in corso...");
  const result = await api("/api/analyze-foglio2", {
    method: "POST",
    body: JSON.stringify({ plan_path: state.default_plan_path }),
  });
  state.foglio2_analysis = result.analysis || null;
  state.foglio2_edits = loadFoglio2EditsFromStorage();
  renderFoglio2Analysis();
  setStatus("Analisi Foglio2 completata.");
}

function colorFromText(text) {
  const palette = [
    { bg: "#d8b11e", border: "#8e7410" },
    { bg: "#2b7de9", border: "#1c549c" },
    { bg: "#18a27a", border: "#0f6a50" },
    { bg: "#e67e22", border: "#9d5312" },
    { bg: "#6b8e23", border: "#486016" },
    { bg: "#2f9fb7", border: "#1e6777" },
    { bg: "#8f6ccf", border: "#5d4690" },
    { bg: "#b38f00", border: "#765f00" },
    { bg: "#4e8d5d", border: "#345d3d" },
    { bg: "#3a7bd5", border: "#254f8a" },
    { bg: "#c78b2a", border: "#835b1a" },
    { bg: "#4f9d69", border: "#346746" },
  ];
  let hash = 0;
  for (let i = 0; i < text.length; i += 1) hash = text.charCodeAt(i) + ((hash << 5) - hash);
  return palette[Math.abs(hash) % palette.length];
}

async function patchAllocation(id, payload) {
  const before = state.allocations.find((a) => Number(a.id) === Number(id));
  const updated = await api(`/api/allocations/${id}`, {
    method: "PATCH",
    body: JSON.stringify(payload),
  });
  if (before) {
    pushUndo({
      label: `allocazione ${before.project} ${before.resource || ""}`.trim(),
      undoMessage: "Ultima modifica allocazione annullata.",
      undo: async () => {
        await api(`/api/allocations/${id}`, {
          method: "PATCH",
          body: JSON.stringify({
            project_id: before.project_id,
            resource_id: before.resource_id,
            role: before.role,
            week_from: before.week_from,
            week_to: before.week_to,
            weight: allocationWeight(before),
            note: before.note || "",
          }),
        });
      },
    });
  }
  return updated;
}

function renderTimeline() {
  const DataSetCtor = getDataSetCtor();
  const TimelineCtor = getTimelineCtor();
  if (!DataSetCtor || !TimelineCtor) {
    if (timeline && typeof timeline.destroy === "function") {
      timeline.destroy();
      timeline = null;
    }
    renderTimelineFallback();
    return;
  }

  ensureGanttOrderWeekOptions();
  const orderWeek = ganttOrderWeekValue();
  const ganttGroupMode = ganttGroupBy?.value || "resource";
  const visibleResourceIds = new Set(
    state.resources.filter((r) => resourceVisibleInGantt(r)).map((r) => Number(r.id))
  );

  let groups = [];
  let items = [];
  if (ganttGroupMode === "project") {
    const projectGroups = state.projects
      .filter((p) => p.code !== "WS OVERALL")
      .filter((p) => p.closed === 0 || p.code === "OVERALL OFFICINA")
      .map((p) => ({ id: `project-${p.id}`, content: `<b>${escapeHtml(p.code)}</b>`, code: p.code }));
    const allocCounts = new Map();
    state.allocations
      .filter((a) => Number(a.week_from) <= orderWeek && Number(a.week_to) >= orderWeek)
      .forEach((a) => {
        const key = `project-${a.project_id}`;
        allocCounts.set(key, (allocCounts.get(key) || 0) + allocationEffectiveWeightForWeek(a, orderWeek));
      });
    projectGroups.sort((a, b) => {
      const ca = allocCounts.get(a.id) || 0;
      const cb = allocCounts.get(b.id) || 0;
      if (cb !== ca) return cb - ca;
      return a.code.localeCompare(b.code);
    });
    const hasAllocations = new Set(
      state.allocations
        .map((a) => `project-${a.project_id}`)
    );
    groups = new DataSetCtor(projectGroups.filter((g) => hasAllocations.has(g.id)));
    items = new DataSetCtor(
      state.allocations
        .map((a) => {
          const colors = colorFromText(a.project || "");
          const warnings = (a.warning_segments || []).map((w) => w.type).join(" ").toUpperCase();
          let border = colors.border;
          let className = "bar-normal";
          if (warnings.includes("CESSATO")) {
            border = "#7f8795";
            className = "bar-ended";
          }
          if (allocationHasDoubleAllocation(a)) {
            className = `${className} double-allocation`;
          }
          const taskClass = allocationHasDoubleAllocation(a) ? "task-bar double-allocation" : "task-bar";
          const resource = state.resources.find((r) => Number(r.id) === Number(a.resource_id));
          const mansioneText = mansioneHoverText(resource, a, a.week_from, a.week_to);
          const warningCodes = mansioneText ? warningCodesForAllocationRange(a, a.week_from, a.week_to) : [];
          const compactTitle = [mansioneText, ...warningCodes].filter(Boolean).join(" | ");
          return {
            id: `alloc-${a.id}`,
            group: `project-${a.project_id}`,
            content: `<div class="${taskClass}" data-allocation-id="${a.id}" data-resource-id="${a.resource_id}" data-week="${a.week_from}" data-project="${escapeHtml(a.project || "")}">${escapeHtml(a.resource || "")}</div>`,
            start: startOfWeek(a.week_from),
            end: startOfWeek(Math.min(53, a.week_to + 1)),
            title: escapeHtml(compactTitle),
            className,
            style: `background:${colors.bg}; border:2px solid ${border}; color:#fff;`,
          };
        })
    );
  } else {
    groups = new DataSetCtor(
      state.resources
        .filter((r) => visibleResourceIds.has(Number(r.id)))
        .map((r) => ({
          id: r.id,
          content: `<b>${escapeHtml(r.name)}</b><div style="font-size:11px;color:#6e7788">${escapeHtml(r.role1 || "")}</div>`,
        }))
    );
    items = new DataSetCtor(
      state.allocations
        .filter((a) => visibleResourceIds.has(Number(a.resource_id)))
        .map((a) => {
          const colors = colorFromText(a.project || "");
          const warnings = (a.warning_segments || []).map((w) => w.type).join(" ").toUpperCase();
          let border = colors.border;
          let className = "bar-normal";
          if (warnings.includes("CESSATO")) {
            border = "#7f8795";
            className = "bar-ended";
          }
          if (allocationHasDoubleAllocation(a)) {
            className = `${className} double-allocation`;
          }
          const taskClass = allocationHasDoubleAllocation(a) ? "task-bar double-allocation" : "task-bar";
          const resource = state.resources.find((r) => Number(r.id) === Number(a.resource_id));
          const mansioneText = mansioneHoverText(resource, a, a.week_from, a.week_to);
          const warningCodes = mansioneText ? warningCodesForAllocationRange(a, a.week_from, a.week_to) : [];
          const compactTitle = [mansioneText, ...warningCodes].filter(Boolean).join(" | ");
          return {
            id: a.id,
            group: a.resource_id,
            content: `<div class="${taskClass}" data-allocation-id="${a.id}" data-resource-id="${a.resource_id}" data-week="${a.week_from}" data-project="${escapeHtml(a.project || "")}">${escapeHtml(a.project)}</div>`,
            start: startOfWeek(a.week_from),
            end: startOfWeek(Math.min(53, a.week_to + 1)),
            title: escapeHtml(compactTitle),
            className,
            style: `background:${colors.bg}; border:2px solid ${border}; color:#fff;`,
          };
        })
    );
  }

  const options = {
    stack: false,
    orientation: "top",
    editable: { updateTime: true, updateGroup: false, remove: false, add: false },
    margin: { item: 6 },
    zoomMin: 1000 * 60 * 60 * 24 * 7,
    zoomMax: 1000 * 60 * 60 * 24 * 7 * 60,
    min: startOfWeek(getCurrentWeek()),
    max: startOfWeek(53),
    onMove: async (item, callback) => {
      if (isMoving) return callback(null);
      isMoving = true;
      try {
        const wf = weekFromDate(item.start);
        const wt = weekFromDate(addDays(item.end, -1));
        const conflicts = allocationConflict(item.group, wf, wt, item.id);
        if (conflicts.length > 0) throw new Error("La risorsa e' gia allocata nel periodo selezionato.");
        await patchAllocation(item.id, { week_from: wf, week_to: wt });
        await loadAll();
        callback(item);
      } catch (err) {
        console.error(err);
        callback(null);
        setStatus(err.message || "Errore trascinamento barra");
      } finally {
        isMoving = false;
      }
    },
  };

  if (!timeline) {
    timeline = new TimelineCtor(timelineEl, items, groups, options);
  } else {
    timeline.setOptions(options);
    timeline.setGroups(groups);
    timeline.setItems(items);
  }
  bindGanttTimelineContextMenu();
}

function parseTimelineAllocationId(itemId) {
  if (itemId === null || itemId === undefined) return null;
  const raw = String(itemId);
  if (raw.startsWith("alloc-")) {
    const n = Number(raw.slice(6));
    return Number.isFinite(n) && n > 0 ? n : null;
  }
  const n = Number(raw);
  return Number.isFinite(n) && n > 0 ? n : null;
}

function bindGanttTimelineContextMenu() {
  if (!timeline || timeline.__ganttContextmenuBound) return;
  timeline.on("contextmenu", (props) => {
    const allocationId = parseTimelineAllocationId(props?.item);
    if (!allocationId) return;
    const ev = props?.event;
    if (ev?.preventDefault) ev.preventDefault();
    if (ev?.srcEvent?.preventDefault) ev.srcEvent.preventDefault();
    const allocation = state.allocations.find((a) => Number(a.id) === Number(allocationId));
    if (!allocation) return;
    const resource = state.resources.find((r) => Number(r.id) === Number(allocation.resource_id));
    if (!resource) return;
    openGanttActionModal({
      resource,
      week_from: Number(allocation.week_from),
      week_to: Number(allocation.week_to),
      project: allocation.project || "",
      role: allocation.role || resource.role1 || "",
      mode: "allocation",
      allocation_id: Number(allocation.id),
    });
  });
  timeline.__ganttContextmenuBound = true;
}

function renderTimelineFallback() {
  const activeWeek = parseWeek(ganttFallbackSort.week) || getCurrentWeek();
  const isResourceSort = ganttFallbackSort.mode === "resource";
  const resourceSortArrow = isResourceSort ? (ganttFallbackSort.dir === "desc" ? "â–¼" : "â–²") : "â‡…";
  const weeks = visibleWeeks();
  const weekHeads = weeks
    .map((w) => {
      const isWeekSort = ganttFallbackSort.mode === "week" && Number(activeWeek) === Number(w);
      const active = isWeekSort ? "is-order-week" : "";
      const arrow = isWeekSort ? (ganttFallbackSort.dir === "asc" ? "â–²" : "â–¼") : "â‡…";
      return `
        <th class="${active}" title="${weekLabelTooltip(w)}">
          <button class="gantt-header-sort-btn ${isWeekSort ? "is-active" : ""}" type="button" data-gantt-sort="week" data-week="${w}" title="Ordina per commessa piu popolosa in W${w}">
            W${w}<span class="caret">â–¼</span>
          </button>
        </th>
      `;
    })
    .join("");
  const baseResources = state.resources.filter((r) => resourceVisibleInGantt(r) && !isExternalResource(r));
  const extMatrix = buildExternalWeekMatrix(weeks);
  const rows = sortResourcesForGantt(baseResources)
    .map((r) => {
      const allocs = state.allocations.filter((a) => Number(a.resource_id) === Number(r.id));
      const unavs = (state.unavailability || []).filter((u) => Number(u.resource_id) === Number(r.id));
      const cells = weeks.map((week) => {
        const active = allocs.find((a) => Number(a.week_from) <= week && Number(a.week_to) >= week);
        const unav = unavs.find((u) => Number(u.week_from) <= week && Number(u.week_to) >= week);
        const endWeek = parseWeekOrDate(r.end_date);
        if (!active) {
          if (unav) {
            const badge = computeCellBadge({ indisp: true });
            return `<td class="gantt-fallback-indisp gantt-cell-action" data-kind="unavailability" data-unavailability-id="${unav.id}" data-resource-id="${r.id}" data-week="${week}" title="${escapeHtml(
              `${formatShortPersonName(r.name)} | ${formatShortRole(r.role1 || "")} | IND | W${unav.week_from}${Number(unav.week_to) > Number(unav.week_from) ? `-W${unav.week_to}` : ""}`
            )}">IND${cellBadgeHtml(badge)}</td>`;
          }
          if (resourceActive(r, week)) {
            const availableTitle = `${formatResourceDisplayLine(r, { role: r.role1 || "", weekFrom: week, weekTo: week })} | W${week}`;
            return `<td class="gantt-fallback-available gantt-cell-action" data-kind="available" data-resource-id="${r.id}" data-week="${week}" title="${escapeHtml(
              availableTitle
            )}"></td>`;
          }
          if (endWeek !== null && endWeek > 0 && endWeek < 999 && week > endWeek && week <= endWeek + 4) {
            return `<td class="gantt-fallback-ended" title="${escapeHtml(`${formatShortPersonName(r.name)} | ${formatShortRole(r.role1 || "")} | NA | W${week}`)}"></td>`;
          }
          return `<td class="gantt-fallback-empty" data-resource-id="${r.id}" data-week="${week}"></td>`;
        }
        const colors = colorFromText(active.project || "");
        const warningLabels = warningLabelsForWeek(active, week);
        const warnings = warningLabels.join(" ").toUpperCase();
        const surplus = !demandHasRoleInWeek(active.project_id, active.role, week);
        const badge = computeCellBadge({
          noFab: surplus,
          fuoriMansione: warnings.includes("FUORI MANSIONE"),
          external: isExternalResource(r),
        });
        const extraClass = warnings.includes("SOVRAPPOSIZIONE")
          ? "gantt-fallback-overlap"
          : warnings.includes("CONTRATTO")
          ? "gantt-fallback-contract"
          : warnings.includes("INDISPONIBILE")
          ? "gantt-fallback-overlap"
          : warnings.includes("CESSATO")
          ? "gantt-fallback-ended"
          : "";
        const doubleClass = isDoubleAllocation(r.id, week) ? "double-allocation" : "";
        const surplusClass = surplus ? "gantt-fallback-surplus" : "";
        const shortLabel = escapeHtml((active.project || "").slice(0, 10));
        const mansioneText = mansioneHoverText(r, active, active.week_from, active.week_to);
        const warningCodes = mansioneText ? warningLabels.map((w) => compactWarningCode(w)).filter(Boolean) : [];
        if (mansioneText && surplus) warningCodes.push("NO-FABB");
        const uniqueCodes = [...new Set(warningCodes)];
        const title = [mansioneText, ...uniqueCodes].filter(Boolean).join(" | ");
        return `<td class="gantt-fallback-fill gantt-cell-action ${extraClass} ${surplusClass} ${doubleClass}" draggable="true" data-kind="allocation" data-allocation-id="${active.id}" data-resource-id="${r.id}" data-week="${week}" style="--fill-bg:${colors.bg}; --fill-border:${colors.border}; background:${colors.bg}; border-color:${colors.border}" title="${escapeHtml(
          title
        )}">${shortLabel}${cellBadgeHtml(badge)}</td>`;
      }).join("");
      return `
        <tr>
          <td class="gantt-fallback-name">
            <strong>${escapeHtml(r.name)}</strong>
            <div>${escapeHtml(r.role1 || "")}</div>
          </td>
          ${cells}
        </tr>
      `;
    })
    .join("");

  const extAggregateRow = (() => {
    const cells = weeks
      .map((week) => {
        const data = extMatrix.aggregateByWeek.get(week) || { qty: 0, projects: new Map() };
        const text = data.qty > 0 ? `-EXT=${data.qty}` : "";
        const title = data.qty > 0
          ? `EXT aggregati W${week} | Totale: ${data.qty} | ${extTooltipFromProjectMap(data.projects)}`
          : `Nessun external in W${week}`;
        return `<td class="gantt-fallback-ext-agg" title="${escapeHtml(title)}">${escapeHtml(text)}</td>`;
      })
      .join("");
    return `
      <tr class="gantt-ext-row">
        <td class="gantt-fallback-name gantt-ext-name">
          <strong>-EXT</strong>
          <div>Totale external</div>
        </td>
        ${cells}
      </tr>
    `;
  })();

  const extDetailRows = externalDetailEnabled()
    ? extMatrix.roleRows
        .map((role) => {
          const perWeek = extMatrix.byRole.get(role) || new Map();
          const cells = weeks
            .map((week) => {
              const cell = perWeek.get(week) || { qty: 0, projects: new Map() };
              const text = cell.qty > 0 ? `-EXT ${cell.qty}` : "";
              const title = cell.qty > 0
                ? `${role}-EXT | W${week} | ${extTooltipFromProjectMap(cell.projects)}`
                : `${role}-EXT | W${week} | 0`;
              return `<td class="gantt-fallback-ext-detail" title="${escapeHtml(title)}">${escapeHtml(text)}</td>`;
            })
            .join("");
          return `
            <tr class="gantt-ext-detail-row">
              <td class="gantt-fallback-name gantt-ext-name">
                <strong>${escapeHtml(`${role}-EXT`)}</strong>
              </td>
              ${cells}
            </tr>
          `;
        })
        .join("")
    : "";

  timelineEl.innerHTML = `
    <div class="gantt-fallback-wrap">
      <table class="gantt-fallback-table">
        <thead>
          <tr>
            <th class="gantt-fallback-name">
              <button class="gantt-header-sort-btn ${isResourceSort ? "is-active" : ""}" type="button" data-gantt-sort="resource" title="Ordina per nome risorsa">
                <span>Risorsa</span><span class="caret">${resourceSortArrow}</span>
              </button>
            </th>
            ${weekHeads}
          </tr>
        </thead>
        <tbody>${rows}${extAggregateRow}${extDetailRows}</tbody>
      </table>
    </div>
  `;
  bindGanttFallbackInteractions();
  timelineEl.querySelectorAll('.gantt-header-sort-btn[data-gantt-sort="week"]').forEach((btn) => {
    const week = Number(btn.dataset.week);
    const selected = ganttFallbackSort.mode === "week" && Number(ganttFallbackSort.week) === week;
    const arrow = selected ? (ganttFallbackSort.dir === "asc" ? "\u25B2" : "\u25BC") : "\u21C5";
    btn.innerHTML = `<span>W${week}</span><span class="caret">${arrow}</span>`;
  });
  const resourceSortBtn = timelineEl.querySelector('.gantt-header-sort-btn[data-gantt-sort="resource"]');
  if (resourceSortBtn) {
    const arrow = isResourceSort ? (ganttFallbackSort.dir === "desc" ? "\u25BC" : "\u25B2") : "\u21C5";
    resourceSortBtn.innerHTML = `<span>Risorsa</span><span class="caret">${arrow}</span>`;
  }
  timelineEl.querySelectorAll(".gantt-header-sort-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const sortType = btn.dataset.ganttSort || "";
      if (sortType === "resource") {
        setGanttFallbackSortResource();
      } else if (sortType === "week") {
        setGanttFallbackSortWeek(btn.dataset.week);
      }
      renderTimelineFallback();
    });
  });
  setStatus("Gantt in modalita' compatibile.");
}

function bindGanttFallbackInteractions() {
  timelineEl.querySelectorAll(".gantt-cell-action").forEach((cell) => {
    cell.addEventListener("contextmenu", (ev) => {
      ev.preventDefault();
      const resource = state.resources.find((r) => Number(r.id) === Number(cell.dataset.resourceId));
      if (!resource) return;
      const allocation =
        cell.dataset.kind === "allocation"
          ? state.allocations.find((a) => Number(a.id) === Number(cell.dataset.allocationId))
          : null;
      openGanttActionModal({
        resource,
        week_from: allocation ? Number(allocation.week_from) : Number(cell.dataset.week),
        week_to: allocation ? Number(allocation.week_to) : Number(cell.dataset.week),
        project: allocation ? allocation.project : cell.dataset.kind === "unavailability" ? "INDISP" : "",
        role: allocation ? allocation.role : resource.role1 || "",
        mode: cell.dataset.kind,
        unavailability_id: cell.dataset.kind === "unavailability" ? Number(cell.dataset.unavailabilityId) : null,
        allocation_id: allocation ? Number(allocation.id) : null,
      });
    });
  });

  timelineEl.querySelectorAll('.gantt-fallback-fill[data-kind="allocation"]').forEach((cell) => {
    cell.addEventListener("dragstart", (ev) => {
      const dragWeek = Number(cell.dataset.week) || null;
      const allocation = state.allocations.find((a) => Number(a.id) === Number(cell.dataset.allocationId));
      let dragEdge = "end";
      if (allocation && dragWeek !== null) {
        const allocationWeekFrom = Number(allocation.week_from);
        const allocationWeekTo = Number(allocation.week_to);
        if (allocationWeekFrom === allocationWeekTo) {
          dragEdge = "end";
        } else {
          const distFromStart = Math.abs(dragWeek - allocationWeekFrom);
          const distFromEnd = Math.abs(dragWeek - allocationWeekTo);
          dragEdge = distFromStart < distFromEnd ? "start" : "end";
        }
      }
      ganttDragState = {
        kind: "allocation",
        allocation_id: Number(cell.dataset.allocationId),
        resource_id: Number(cell.dataset.resourceId),
        drag_week: dragWeek,
        edge: dragEdge,
      };
      ev.dataTransfer.setData("text/plain", String(cell.dataset.allocationId));
    });
    cell.addEventListener("dragend", () => {
      ganttDragState = null;
    });
  });

  timelineEl.querySelectorAll(".gantt-fallback-table td[data-week]").forEach((cell) => {
    cell.addEventListener("dragover", (ev) => {
      if (!ganttDragState) return;
      if (Number(cell.dataset.resourceId) !== Number(ganttDragState.resource_id)) return;
      ev.preventDefault();
    });
    cell.addEventListener("drop", async (ev) => {
      if (!ganttDragState || ganttDragState.kind !== "allocation") return;
      if (Number(cell.dataset.resourceId) !== Number(ganttDragState.resource_id)) return;
      ev.preventDefault();
      const allocation = state.allocations.find((a) => Number(a.id) === Number(ganttDragState.allocation_id));
      const targetWeek = Number(cell.dataset.week);
      if (!allocation || !targetWeek) return;
      let week_from = Number(allocation.week_from);
      let week_to = Number(allocation.week_to);
      if (ganttDragState.edge === "start") {
        week_from = Math.min(targetWeek, week_to);
      } else {
        week_to = Math.max(targetWeek, week_from);
      }
      try {
        await patchAllocation(allocation.id, { week_from, week_to });
        await loadAll(true);
        setStatus(`Allocazione ${allocation.project} aggiornata a W${week_from}-W${week_to}.`);
      } catch (err) {
        console.error(err);
        setStatus(err.message || "Errore trascinamento barra Gantt");
      } finally {
        ganttDragState = null;
      }
    });
  });
}

async function saveAssignment({ project_id, project, role, resource_id, week_from, week_to }, { refresh = true } = {}) {
  const resource = state.resources.find((r) => Number(r.id) === Number(resource_id));
  if (!resource) throw new Error("Risorsa non trovata.");
  const projectCode = String(project || "").trim();
  const owner = getOwningOverall(projectCode);
  if (!isAllocableProjectCode(projectCode)) {
    const ownerCode = String(owner?.code || owner?.id || owner?.name || "OVERALL").trim();
    const message = `La commessa ${projectCode} è gestita da ${ownerCode}; alloca su ${ownerCode}`;
    alert(message);
    throw new Error(message);
  }
  const projectRow = state.projects.find((p) => normalizeProjectCode(p.code) === normalizeProjectCode(projectCode));
  if (isOverallProjectCode(projectCode) && !projectRow) {
    throw new Error(`Overall ${projectCode} non presente tra le commesse; crea prima la commessa overall.`);
  }
  const resolvedProjectId = projectRow ? Number(projectRow.id) : Number(project_id);
  if (!Number.isFinite(resolvedProjectId) || resolvedProjectId <= 0) {
    throw new Error(`Commessa ${projectCode} non trovata.`);
  }
  const external = isExternalResource(resource);
  if (!resourceHasAnyRole(resource)) {
    throw new Error(`La risorsa ${resource.name} non ha una mansione assegnata. Compilala in Anagrafica Risorse prima di allocarla.`);
  }
  if (!resourceActive(resource, week_from)) throw new Error("La risorsa non e' attiva nella week selezionata.");

  const merge = external
    ? null
    : state.allocations.find(
        (a) =>
          Number(a.project_id) === Number(resolvedProjectId) &&
          Number(a.resource_id) === Number(resource_id) &&
          a.role === role &&
          Number(a.week_from) <= Number(week_to) + 1 &&
          Number(a.week_to) >= Number(week_from) - 1
      );

  validateInternalAllocationRange(resource_id, week_from, week_to, merge ? Number(merge.id) : null);

  if (merge) {
    await api(`/api/allocations/${merge.id}`, {
      method: "PATCH",
      body: JSON.stringify({
        week_from: Math.min(Number(merge.week_from), Number(week_from)),
        week_to: Math.max(Number(merge.week_to), Number(week_to)),
      }),
    });
  } else {
    await api("/api/allocations", {
      method: "POST",
      body: JSON.stringify({ project_id: resolvedProjectId, project: projectCode, role, resource_id, week_from, week_to, weight: 1 }),
    });
  }
  logUserAction(
    "Allocazione",
    `${resource.name} -> ${projectCode} | ${role} | W${week_from}-W${week_to}`
  );
  if (refresh) {
    await loadAll();
  }
  if (refresh && !roleMatches(resource, role)) {
    setStatus("Allocazione salvata con avviso: mansione non coerente.");
  }
}

async function saveUnavailable({ resource_id, week_from, week_to }, { refresh = true } = {}) {
  const resource = state.resources.find((r) => Number(r.id) === Number(resource_id));
  if (!resource) throw new Error("Risorsa non trovata.");
  if (!resourceActive(resource, week_from)) throw new Error("La risorsa non e' attiva nella week selezionata.");
  await api("/api/unavailability", {
    method: "POST",
    body: JSON.stringify({ resource_id, week_from, week_to, reason: "INDISP" }),
  });
  logUserAction("INDISP", `${resource.name} | W${week_from}-W${week_to}`);
  if (refresh) {
    await loadAll();
  }
}

async function handleAssignSave() {
  if (!selectedTarget) {
    setStatus("Seleziona una cella del planner.");
    return;
  }
  const useExt = usingExternalAssignMode();
  let resourceIds = selectedAssignResourceIds();
  const extQty = Math.max(1, parseIntOr(assignExternalQty?.value, 1));
  let week_from = parseWeek(assignWeekFrom.value);
  let week_to = parseWeek(assignWeekTo.value);
  if (week_from === null || week_to === null) {
    setStatus("Compila week da e week a.");
    return;
  }
  if (week_to < week_from) [week_from, week_to] = [week_to, week_from];
  if (useExt) {
    await api("/api/allocations/unassign-scope", {
      method: "POST",
      body: JSON.stringify({
        project_id: selectedTarget.project_id,
        role: selectedTarget.role,
        week_from,
        week_to,
        ext_only: true,
      }),
    });
    let extCandidates = extResourcesForRole(selectedTarget.role);
    if (extCandidates.length < extQty) {
      const missing = extQty - extCandidates.length;
      const baseExtName = `${selectedTarget.role}-EXT`;
      const usedNames = new Set(
        state.resources
          .filter((r) => isExternalResource(r))
          .map((r) => String(r.name || "").trim().toUpperCase())
          .filter(Boolean)
      );
      for (let i = 0; i < missing; i += 1) {
        let nameCandidate = baseExtName;
        let idx = 2;
        while (usedNames.has(nameCandidate.toUpperCase())) {
          nameCandidate = `${baseExtName} ${idx}`;
          idx += 1;
        }
        usedNames.add(nameCandidate.toUpperCase());
        await api("/api/resources", {
          method: "POST",
          body: JSON.stringify({
            name: nameCandidate,
            role1: selectedTarget.role,
            employer: "EXT",
            note: "Auto-creata per assegnazione EXT",
          }),
        });
      }
      await loadAll(true);
      extCandidates = extResourcesForRole(selectedTarget.role);
    }
    resourceIds = extCandidates.slice(0, extQty).map((r) => Number(r.id));
  }
  if (resourceIds.length === 0) {
    setStatus(useExt ? "Nessuna risorsa EXT disponibile per la mansione selezionata." : "Compila almeno una risorsa.");
    return;
  }
  const warnings = [];
  for (const resource_id of resourceIds) {
    await saveAssignment({
      project_id: selectedTarget.project_id,
      project: selectedTarget.project,
      role: selectedTarget.role,
      resource_id,
      week_from,
      week_to,
    }, { refresh: false });
    const resource = state.resources.find((r) => Number(r.id) === Number(resource_id));
    if (resource && !roleMatches(resource, selectedTarget.role)) {
      warnings.push(resource.name);
    }
  }
  await loadAll(true);
  renderAssignResourceOptions();
  selectTarget({ ...selectedTarget, week_from, week_to });
  if (useExt) {
    setStatus(`Assegnazione EXT aggiornata: ${resourceIds.length} risorse su W${week_from}-W${week_to} (sostituzione periodo).`);
    return;
  }
  if (warnings.length > 0) {
    setStatus(`Allocazione salvata con avviso mansione per: ${warnings.join(", ")}.`);
    return;
  }
  setStatus(resourceIds.length === 1 ? "Allocazione salvata." : `${resourceIds.length} allocazioni salvate.`);
}

async function handleMarkUnavailable() {
  const resourceIds = selectedAssignResourceIds();
  let week_from = parseWeek(assignWeekFrom.value);
  let week_to = parseWeek(assignWeekTo.value);
  if (resourceIds.length === 0 || week_from === null || week_to === null) {
    setStatus("Seleziona almeno una risorsa e compila week da/week a per INDISP.");
    return;
  }
  if (week_to < week_from) [week_from, week_to] = [week_to, week_from];
  for (const resource_id of resourceIds) {
    await saveUnavailable({ resource_id, week_from, week_to }, { refresh: false });
  }
  await loadAll(true);
  renderAssignResourceOptions();
  renderSelectionAllocations();
  setStatus(resourceIds.length === 1 ? "INDISP salvata." : `${resourceIds.length} periodi INDISP salvati.`);
}

async function handleDemandSave() {
  if (!selectedTarget) {
    setStatus("Seleziona una cella del planner.");
    return;
  }
  if (isOverallProjectCode(selectedTarget.project)) {
    alert("Il fabbisogno di OVERALL non è modificabile direttamente");
    return;
  }
  let weekFrom = parseWeek(demandWeekFrom.value);
  let weekTo = parseWeek(demandWeekTo.value);
  const qty = Math.max(0, parseIntOr(demandQty.value, 0));
  if (weekFrom === null || weekTo === null) {
    setStatus("Compila week da e week a del fabbisogno.");
    return;
  }
  if (weekTo < weekFrom) [weekFrom, weekTo] = [weekTo, weekFrom];
  const rows = [];
  for (let week = weekFrom; week <= weekTo; week += 1) {
    rows.push({
      project_id: selectedTarget.project_id,
      role: selectedTarget.role,
      week,
      qty,
    });
  }
  await api("/api/demands/bulk", {
    method: "POST",
    body: JSON.stringify({ rows }),
  });
  logUserAction("Fabbisogno", `${selectedTarget.project} | ${selectedTarget.role} | W${weekFrom}-W${weekTo} -> ${qty}`);
  await loadAll();
  selectTarget({ ...selectedTarget, week: weekFrom, week_from: weekFrom, week_to: weekTo });
  setStatus("Fabbisogno aggiornato.");
}

function clearSelection() {
  selectedTarget = null;
  selectedAllocationIds.clear();
  demandQty.value = "";
  demandWeekFrom.value = "";
  demandWeekTo.value = "";
  assignProject.value = "";
  assignRole.value = "";
  assignWeekFrom.value = "";
  assignWeekTo.value = "";
  if (assignUseExternal) assignUseExternal.checked = false;
  if (assignExternalQty) assignExternalQty.value = "1";
  assignResource.innerHTML = "";
  updateAssignModeUi();
  selectionInfo.textContent = "Seleziona una cella del planner.";
  selectionAllocations.innerHTML = `<div class="empty-box">Seleziona una cella per vedere le allocazioni.</div>`;
  if (workshopBreakdownBox) {
    workshopBreakdownBox.innerHTML = `<div class="empty-box">Seleziona OVERALL OFFICINA per vedere il dettaglio delle commesse WS.</div>`;
  }
  renderResourcePool();
}

function bindPanelExpanders() {
  expandAssignBtn?.addEventListener("click", () => {
    const card = expandAssignBtn.closest(".card");
    card?.classList.toggle("panel-expanded");
    if (card && !card.classList.contains("panel-expanded")) {
      card.style.left = "";
      card.style.top = "";
      card.style.inset = "";
      card.style.width = "";
      card.style.height = "";
      card.style.margin = "";
    }
    expandAssignBtn.textContent = card?.classList.contains("panel-expanded") ? "Riduci" : "Ingrandisci";
  });
  expandGanttBtn?.addEventListener("click", () => {
    const card = expandGanttBtn.closest(".card");
    card?.classList.toggle("panel-expanded");
    if (card && !card.classList.contains("panel-expanded")) {
      card.style.left = "";
      card.style.top = "";
      card.style.inset = "";
      card.style.width = "";
      card.style.height = "";
      card.style.margin = "";
    }
    expandGanttBtn.textContent = card?.classList.contains("panel-expanded") ? "Riduci" : "Ingrandisci";
  });
}

function bindPlannerSideToggle() {
  plannerSideToggle?.addEventListener("click", () => {
    const grid = plannerSideCard?.closest(".planner-grid");
    if (!grid) return;
    grid.classList.toggle("side-collapsed");
    const collapsed = grid.classList.contains("side-collapsed");
    if (collapsed) grid.classList.remove("side-active");
    plannerSideToggle.title = collapsed ? "Mostra pannello laterale" : "Nascondi pannello laterale";
    plannerSideToggle.setAttribute("aria-label", plannerSideToggle.title);
    setStatus(collapsed ? "Pannello laterale nascosto." : "Pannello laterale visibile.");
  });
}

function bindPlannerSideAutoExpand() {
  const grid = plannerSideCard?.closest(".planner-grid");
  if (!grid || !plannerSideCard) return;

  const targets = [
    plannerSideCard,
    resourcePool,
    selectionAllocations,
    assignProject,
    assignRole,
    assignWeekFrom,
    assignWeekTo,
    assignResource,
    assignSaveBtn,
    demandQty,
    demandWeekFrom,
    demandWeekTo,
    demandSaveBtn,
  ].filter(Boolean);

  const activate = () => {
    if (grid.classList.contains("side-collapsed")) return;
    grid.classList.add("side-active");
  };

  const deactivateIfOutside = () => {
    if (grid.classList.contains("side-collapsed")) {
      grid.classList.remove("side-active");
      return;
    }
    const activeEl = document.activeElement;
    if (activeEl && plannerSideCard.contains(activeEl)) return;
    grid.classList.remove("side-active");
  };

  targets.forEach((el) => {
    el.addEventListener("focusin", activate);
    el.addEventListener("mousedown", activate);
    el.addEventListener("click", activate);
  });

  plannerSideCard.addEventListener("focusout", () => {
    window.setTimeout(deactivateIfOutside, 0);
  });

  document.addEventListener("mousedown", (ev) => {
    const target = ev.target;
    if (plannerSideCard.contains(target) || plannerSideToggle?.contains(target)) {
      activate();
      return;
    }
    grid.classList.remove("side-active");
  });
}

function clampGanttSplitHeight(height) {
  const viewportCap = Math.max(window.innerHeight - 340, 320);
  return Math.min(Math.max(height, 260), viewportCap);
}

function applyGanttSplitHeight(height) {
  if (!ganttSplit) return;
  const clamped = clampGanttSplitHeight(height);
  ganttSplit.style.setProperty("--gantt-top-height", `${clamped}px`);
}

function loadGanttSplitHeight() {
  try {
    const saved = window.localStorage.getItem(GANTT_SPLIT_STORAGE_KEY);
    const parsed = parseInt(saved || "", 10);
    if (Number.isFinite(parsed)) {
      applyGanttSplitHeight(parsed);
      return;
    }
  } catch (err) {
    console.warn("gantt split storage", err);
  }
  applyGanttSplitHeight(560);
}

function persistGanttSplitHeight(height) {
  try {
    window.localStorage.setItem(GANTT_SPLIT_STORAGE_KEY, String(clampGanttSplitHeight(height)));
  } catch (err) {
    console.warn("gantt split save", err);
  }
}

function bindGanttSplitResize() {
  if (!ganttSplit || !ganttSplitHandle || !timelineEl) return;

  const pointerMove = (ev) => {
    if (!ganttSplitResizeState) return;
    const nextHeight = ganttSplitResizeState.startHeight + (ev.clientY - ganttSplitResizeState.startY);
    applyGanttSplitHeight(nextHeight);
  };

  const pointerEnd = () => {
    if (!ganttSplitResizeState) return;
    ganttSplit.classList.remove("is-resizing");
    const currentHeight = parseInt(getComputedStyle(ganttSplit).getPropertyValue("--gantt-top-height"), 10);
    if (Number.isFinite(currentHeight)) persistGanttSplitHeight(currentHeight);
    ganttSplitResizeState = null;
  };

  ganttSplitHandle.addEventListener("pointerdown", (ev) => {
    if (ganttCard?.classList.contains("panel-expanded")) return;
    const currentHeight = parseInt(getComputedStyle(ganttSplit).getPropertyValue("--gantt-top-height"), 10) || timelineEl.getBoundingClientRect().height || 560;
    ganttSplitResizeState = {
      startY: ev.clientY,
      startHeight: currentHeight,
    };
    ganttSplit.classList.add("is-resizing");
    ganttSplitHandle.setPointerCapture?.(ev.pointerId);
    ev.preventDefault();
  });

  window.addEventListener("pointermove", pointerMove);
  window.addEventListener("pointerup", pointerEnd);
  window.addEventListener("pointercancel", pointerEnd);
  window.addEventListener("resize", () => {
    const currentHeight = parseInt(getComputedStyle(ganttSplit).getPropertyValue("--gantt-top-height"), 10);
    if (Number.isFinite(currentHeight)) applyGanttSplitHeight(currentHeight);
  });

  ganttSplitHandle.addEventListener("keydown", (ev) => {
    if (ganttCard?.classList.contains("panel-expanded")) return;
    if (!["ArrowUp", "ArrowDown"].includes(ev.key)) return;
    ev.preventDefault();
    const currentHeight = parseInt(getComputedStyle(ganttSplit).getPropertyValue("--gantt-top-height"), 10) || 560;
    const delta = ev.key === "ArrowUp" ? -40 : 40;
    const nextHeight = clampGanttSplitHeight(currentHeight + delta);
    applyGanttSplitHeight(nextHeight);
    persistGanttSplitHeight(nextHeight);
  });

  loadGanttSplitHeight();
}

function bindGanttModal() {
  ganttActionCloseBtn?.addEventListener("click", closeGanttActionModal);
  ganttProjectSelect?.addEventListener("change", toggleGanttRoleInput);
  ganttRoleSelect?.addEventListener("change", renderGanttConflictBox);
  ganttWeekFromInput?.addEventListener("input", renderGanttConflictBox);
  ganttWeekToInput?.addEventListener("input", renderGanttConflictBox);
  ganttActionSaveBtn?.addEventListener("click", () => saveGanttActionModal().catch((e) => setStatus(e.message)));
  ganttActionDeleteBtn?.addEventListener("click", () => deleteGanttUnavailable().catch((e) => setStatus(e.message)));
  document.querySelectorAll("[data-modal-close='ganttActionModal']").forEach((el) => {
    el.addEventListener("click", closeGanttActionModal);
  });
  plannerPrintCloseBtn?.addEventListener("click", closePlannerPrintModal);
  plannerPrintRunBtn?.addEventListener("click", runPlannerPrint);
  document.querySelectorAll("[data-modal-close='plannerPrintModal']").forEach((el) => {
    el.addEventListener("click", closePlannerPrintModal);
  });
}

function bindDraggableModal(modalCard, dragHandle) {
  if (!modalCard || !dragHandle) return;
  dragHandle.addEventListener("pointerdown", (ev) => {
    const interactive = ev.target.closest("button, input, select, textarea, a");
    if (interactive) return;
    const rect = modalCard.getBoundingClientRect();
    modalCard.style.margin = "0";
    modalCard.style.left = `${rect.left}px`;
    modalCard.style.top = `${rect.top}px`;
    modalCard.setPointerCapture?.(ev.pointerId);
    modalCard.classList.add("is-dragging");
    floatingDragState = {
      pointerId: ev.pointerId,
      offsetX: ev.clientX - rect.left,
      offsetY: ev.clientY - rect.top,
      el: modalCard,
    };
    ev.preventDefault();
  });

  const move = (ev) => {
    if (!floatingDragState || ev.pointerId !== floatingDragState.pointerId || floatingDragState.el !== modalCard) return;
    const maxLeft = Math.max(window.innerWidth - modalCard.offsetWidth, 0);
    const maxTop = Math.max(window.innerHeight - modalCard.offsetHeight, 0);
    const left = Math.min(Math.max(ev.clientX - floatingDragState.offsetX, 0), maxLeft);
    const top = Math.min(Math.max(ev.clientY - floatingDragState.offsetY, 0), maxTop);
    modalCard.style.left = `${left}px`;
    modalCard.style.top = `${top}px`;
  };

  const end = (ev) => {
    if (!floatingDragState || ev.pointerId !== floatingDragState.pointerId || floatingDragState.el !== modalCard) return;
    modalCard.classList.remove("is-dragging");
    floatingDragState = null;
  };

  window.addEventListener("pointermove", move);
  window.addEventListener("pointerup", end);
  window.addEventListener("pointercancel", end);
}

function bindDraggableFloatingCard(card, dragHandle) {
  if (!card || !dragHandle) return;
  dragHandle.classList.add("modal-drag-handle");
  dragHandle.addEventListener("pointerdown", (ev) => {
    const interactive = ev.target.closest("button, input, select, textarea, a");
    if (interactive || !card.classList.contains("panel-expanded")) return;
    const rect = card.getBoundingClientRect();
    card.style.inset = "auto";
    card.style.margin = "0";
    card.style.left = `${rect.left}px`;
    card.style.top = `${rect.top}px`;
    card.style.width = `${rect.width}px`;
    card.style.height = `${rect.height}px`;
    card.classList.add("is-dragging");
    floatingDragState = {
      pointerId: ev.pointerId,
      offsetX: ev.clientX - rect.left,
      offsetY: ev.clientY - rect.top,
      el: card,
    };
    ev.preventDefault();
  });

  const move = (ev) => {
    if (!floatingDragState || ev.pointerId !== floatingDragState.pointerId || floatingDragState.el !== card) return;
    const maxLeft = Math.max(window.innerWidth - card.offsetWidth, 0);
    const maxTop = Math.max(window.innerHeight - card.offsetHeight, 0);
    const left = Math.min(Math.max(ev.clientX - floatingDragState.offsetX, 0), maxLeft);
    const top = Math.min(Math.max(ev.clientY - floatingDragState.offsetY, 0), maxTop);
    card.style.left = `${left}px`;
    card.style.top = `${top}px`;
  };

  const end = (ev) => {
    if (!floatingDragState || ev.pointerId !== floatingDragState.pointerId || floatingDragState.el !== card) return;
    card.classList.remove("is-dragging");
    floatingDragState = null;
  };

  window.addEventListener("pointermove", move);
  window.addEventListener("pointerup", end);
  window.addEventListener("pointercancel", end);
}

function bindKeyboardShortcuts() {
  const clickIfEnabled = (btn) => {
    if (!btn || btn.disabled) return;
    btn.click();
  };

  const onEnter = (elements, action) => {
    elements.filter(Boolean).forEach((el) => {
      el.addEventListener("keydown", (ev) => {
        if (ev.key !== "Enter") return;
        if (ev.shiftKey || ev.ctrlKey || ev.altKey || ev.metaKey) return;
        ev.preventDefault();
        action();
      });
    });
  };

  onEnter(
    [demandQty, demandWeekFrom, demandWeekTo],
    () => clickIfEnabled(demandSaveBtn)
  );

  onEnter(
    [assignWeekFrom, assignWeekTo, assignResource],
    () => clickIfEnabled(assignSaveBtn)
  );

  onEnter(
    [projectCodeInput, projectActivityInput, projectTypeInput, projectClosedInput, projectWorkshopRollupInput],
    () => clickIfEnabled(projectAdminSaveBtn)
  );

  onEnter(
    [projectDemandRoleInput],
    () => clickIfEnabled(projectDemandAddRoleBtn)
  );

  onEnter(
    [ganttProjectSelect, ganttRoleSelect, ganttWeekFromInput, ganttWeekToInput],
    () => clickIfEnabled(ganttActionSaveBtn)
  );

  document.addEventListener("keydown", (ev) => {
    if ((ev.ctrlKey || ev.metaKey) && !ev.shiftKey && !ev.altKey && ev.key.toLowerCase() === "z") {
      const tag = (document.activeElement?.tagName || "").toUpperCase();
      if (tag !== "INPUT" && tag !== "TEXTAREA" && tag !== "SELECT") {
        ev.preventDefault();
        undoLastAction().catch((e) => setStatus(e.message));
        return;
      }
    }
    if (ev.key !== "Escape") return;
    if (!ganttActionModal?.hidden) {
      ev.preventDefault();
      closeGanttActionModal();
    }
  });
}

async function loadAll(silent = false) {
  if (!silent) setStatus("Aggiornamento...");
  const currentProjectFilter = plannerProjectFilter ? plannerProjectFilter.value : "";
  const currentRoleFilter = plannerRoleFilter ? plannerRoleFilter.value : "";
  const currentGanttDemandProjectFilter = ganttDemandProjectFilter?.value || "";
  const currentGanttDemandRoleFilter = ganttDemandRoleFilter?.value || "";
  const data = await api("/api/all");
  state = {
    resources: data.resources || [],
    projects: data.projects || [],
    demands: data.demands || [],
    allocations: (data.allocations || []).map((a) => ({ ...a, weight: allocationWeight(a) })),
    overalls: state.overalls || {},
    unavailability: data.unavailability || [],
    report_weekly: data.report_weekly || [],
    report_role_weekly: data.report_role_weekly || [],
    standard_roles: data.standard_roles || [],
    default_demand_path: data.default_demand_path || "",
    default_resources_path: data.default_resources_path || "",
    default_plan_path: data.default_plan_path || "",
    current_demand_path: data.current_demand_path || getDemandPathOverride() || data.default_demand_path || "",
    lastDemandDiffs: state.lastDemandDiffs || [],
    lastDemandDiffMap: state.lastDemandDiffMap || {},
    lastDemandDiffMeta: state.lastDemandDiffMeta || {},
  };
  buildOverallOfficina();
  fillRoleList();
  fillProjectFilter();
  fillGanttDemandFilters();
  plannerProjectFilter.value = currentProjectFilter;
  plannerRoleFilter.value = currentRoleFilter;
  if (!parseWeek(ganttFallbackSort.week)) ganttFallbackSort.week = getCurrentWeek();
  if (ganttDemandProjectFilter) ganttDemandProjectFilter.value = currentGanttDemandProjectFilter;
  if (ganttDemandRoleFilter) ganttDemandRoleFilter.value = currentGanttDemandRoleFilter;
  renderSummary();
  renderPlannerHeader();
  renderPlannerMatrix();
  renderGanttDemandMatrix();
  renderProjectsAdmin();
  renderResourcesTable();
  renderResourcePool();
  updateAssignModeUi();
  renderSelectionAllocations();
  ensureGanttOrderWeekOptions();
  try {
    fillReportBuilderFilters();
    syncReportBuilderControls();
    renderReportBuilder();
  } catch (err) {
    console.error("Errore render report:", err);
    setStatus(`Errore report: ${err.message}`);
  }
  try {
    renderTimeline();
  } catch (err) {
    console.error("Errore render gantt:", err);
    setStatus(`Errore Gantt: ${err.message}`);
  }
  if (selectedTarget) {
    selectionInfo.textContent = `${selectedTarget.project} | ${selectedTarget.role} | W${assignWeekFrom.value || selectedTarget.week_from || selectedTarget.week} - W${assignWeekTo.value || selectedTarget.week_to || selectedTarget.week}`;
    renderAssignResourceOptions();
    renderSelectionAllocations();
  }
  if (!silent) setStatus(`Aggiornato ${new Date().toLocaleTimeString()}`);
}

safeOn(refreshBtn, "click", () => loadAll().catch((e) => setStatus(`Errore refresh: ${e.message}`)), "refresh");
safeOn(undoBtn, "click", () => undoLastAction().catch((e) => setStatus(e.message)), "undo");
safeOn(activityLogBtn, "click", openActivityLog, "activity log open");
safeOn(activityLogClose, "click", closeActivityLog, "activity log close");
safeOn(restoreBackupBtn, "click", async () => {
  try {
    const ok = window.confirm("Vuoi ripristinare l'ultimo salvataggio? Le modifiche recenti verranno perse.");
    if (!ok) return;
    setStatus("Ripristino ultimo salvataggio...");
    await api("/api/backup/restore-latest", { method: "POST", body: JSON.stringify({}) });
    await loadAll();
    logUserAction("Ripristino ultimo", "Ripristino da backup");
    setStatus("Ripristino completato.");
  } catch (err) {
    console.error(err);
    setStatus(`Errore ripristino: ${err.message}`);
  }
}, "restore latest backup");
safeOn(reportSaveViewBtn, "click", () => {
  const name = window.prompt("Nome della vista report:", "");
  if (!name) return;
  readReportBuilderControls();
  const views = getSavedReportViews();
  views.push({
    name: name.trim(),
    config: { ...reportBuilderState },
  });
  saveSavedReportViews(views);
  renderSavedReportViews();
  setStatus(`Vista report salvata: ${name.trim()}`);
}, "report save view");
safeOn(reportDemandOverviewCard, "click", () => {
  reportDetailOpen = true;
  activateReportPreset("demand_allocated");
}, "report demand overview");
safeOn(reportDemandRoleOverviewCard, "click", () => {
  reportDetailOpen = true;
  activateReportPreset("demand_allocated_role");
}, "report role overview");
safeOn(reportIndicatorAllocation, "click", () => {
  reportDetailOpen = true;
  activateReportPreset("workforce_utilization");
}, "report allocation indicator");
safeOn(reportIndicatorCoverage, "click", () => {
  reportDetailOpen = true;
  activateReportPreset("demand_allocated");
}, "report coverage indicator");
safeOn(reportUtilizationOverviewCard, "click", () => {
  reportDetailOpen = true;
  activateReportPreset("workforce_utilization");
}, "report utilization overview");
safeOn(reportCustomInsightOpen, "click", () => {
  reportDetailOpen = false;
  activateReportPreset("custom_insight");
}, "report custom insight open");
safeOn(reportPreviewBtn, "click", () => {
  reportDetailOpen = true;
  if (reportDetailPresetId === "custom_insight") {
    reportDetailPresetId = reportChartType?.value === "utilization" ? "workforce_utilization" : "demand_allocated";
  }
  renderReportBuilder();
}, "report preview");
safeOn(reportModeRequestedBtn, "click", () => {
  reportBuilderState.mode = "mansione";
  if (reportDetailPresetId === "workforce_utilization") reportDetailPresetId = "demand_allocated_role";
  syncReportBuilderControls();
  saveReportBuilderState();
  renderReportBuilder();
}, "report mode requested");
safeOn(reportModeRealBtn, "click", () => {
  reportBuilderState.mode = "operativa";
  if (reportDetailPresetId === "demand_allocated_role") reportDetailPresetId = "demand_allocated";
  syncReportBuilderControls();
  saveReportBuilderState();
  renderReportBuilder();
}, "report mode real");
safeOn(printPlannerBtn, "click", openPlannerPrintModal, "print planner");
safeOn(printResourcesBtn, "click", () => {
  const activeBtn = document.querySelector('.sheet-btn[data-sheet="risorse"]');
  if (activeBtn && !activeBtn.classList.contains("active")) activeBtn.click();
  window.setTimeout(() => window.print(), 30);
}, "print resources");
safeOn(printGanttBtn, "click", () => {
  const activeBtn = document.querySelector('.sheet-btn[data-sheet="gantt"]');
  if (activeBtn && !activeBtn.classList.contains("active")) activeBtn.click();
  window.setTimeout(() => window.print(), 30);
}, "print gantt");
safeOn(importDemandsBtn, "click", async () => {
  try {
    setStatus("Analisi nuovo fabbisogno...");
    const basePath = currentDemandPathValue();
    const demandPath = promptDemandPath(basePath);
    if (!demandPath) {
      setStatus("Import annullato.");
      return;
    }
    const preview = await api("/api/import-demands-preview", {
      method: "POST",
      body: JSON.stringify({
        demand_path: demandPath,
      }),
    });
    if (preview?.analysis?.diffs) {
      state.lastDemandDiffs = preview.analysis.diffs;
      state.lastDemandDiffMap = buildDemandDiffMap(preview.analysis.diffs);
      state.lastDemandDiffMeta = { current_week: preview.analysis.current_week || null };
    }
    const message = `Vuoi importare il nuovo fabbisogno?\n\n${summarizeImportAnalysis(preview.analysis)}\n\nVerrÃ  creato un salvataggio automatico.`;
    const importChoice = await showImportPreviewModal(
      message,
      preview?.analysis?.unknown_roles || []
    );
    if (!importChoice.confirmed) {
      setStatus("Import annullato.");
      return;
    }
    setStatus("Import fabbisogno in corso...");
    await api("/api/import-demands-only", {
      method: "POST",
      body: JSON.stringify({
        demand_path: demandPath,
        replace_existing_demands: true,
        role_mapping: importChoice.roleMapping || {},
        create_roles: importChoice.createRoles || [],
      }),
    });
    state.current_demand_path = demandPath;
    demandSourcePath.textContent = currentDemandPathValue();
    await loadAll();
    setStatus("Fabbisogno importato. Se serve, usa Ripristina ultimo.");
  } catch (err) {
    console.error(err);
    setStatus(`Errore import fabbisogno: ${err.message}`);
  }
}, "import demands");
safeOn(openDemandSourceBtn, "click", async () => {
  try {
    await api("/api/open-demand-source", {
      method: "POST",
      body: JSON.stringify({ demand_path: currentDemandPathValue() }),
    });
    setStatus("File sorgente aperto.");
  } catch (err) {
    console.error(err);
    setStatus(`Errore apertura file sorgente: ${err.message}`);
  }
}, "open demand source");
safeOn(importPlanningBtn, "click", async () => {
  try {
    setStatus("Import iniziale in corso...");
    await api("/api/import-resource-planning", {
      method: "POST",
      body: JSON.stringify({
        demand_path: currentDemandPathValue(),
        resources_path: state.default_resources_path,
        replace_all: true,
      }),
    });
    await loadAll();
    setStatus("Import iniziale completato.");
  } catch (err) {
    console.error(err);
    setStatus(`Errore import: ${err.message}`);
  }
}, "import planning");

safeOn(plannerProjectFilter, "change", () => {
  renderSummary();
  renderPlannerMatrix();
}, "planner project filter");
safeOn(plannerRoleFilter, "change", () => {
  renderSummary();
  renderPlannerMatrix();
  renderResourcePool();
}, "planner role filter");
safeOn(ganttDemandProjectFilter, "change", renderGanttDemandMatrix, "gantt demand project filter");
safeOn(ganttDemandRoleFilter, "change", renderGanttDemandMatrix, "gantt demand role filter");
safeOn(ganttShowExternalDetail, "change", () => {
  persistGanttExternalDetailFlag(externalDetailEnabled());
  renderTimeline();
}, "gantt external detail");
safeOn(showZeroDemandProjects, "change", () => {
  renderSummary();
  renderPlannerMatrix();
}, "show zero demand projects");
safeOn(resourceSearch, "input", renderResourcePool, "resource search");
safeOn(resourceTableSearch, "input", renderResourcesTable, "resource table search");
safeOn(resourceRoleFilter, "change", renderResourcesTable, "resource role filter");
safeOn(resourceStatusFilter, "change", renderResourcesTable, "resource status filter");
safeOn(resourceSortBy, "change", renderResourcesTable, "resource sort");
safeOn(resourceColumnsBtn, "click", (ev) => {
  ev.stopPropagation();
  resourceColumnsPanel?.classList.toggle("hidden");
  if (resourceColumnsPanel && !resourceColumnsPanel.classList.contains("hidden")) {
    renderResourceColumnChooser();
  }
}, "resource columns toggle");
safeOn(reportWeekFrom, "change", () => renderReportBuilder(), "report week from");
safeOn(reportWeekTo, "change", () => renderReportBuilder(), "report week to");
safeOn(reportScale, "change", () => renderReportBuilder(), "report scale");
safeOn(reportChartType, "change", () => renderReportBuilder(), "report chart type");
safeOn(reportAggregateScope, "change", () => {
  if (reportAggregateScope.value) reportProjectFilter.value = "";
  renderReportBuilder();
}, "report aggregate scope");
safeOn(reportProjectFilter, "change", () => {
  if (reportProjectFilter.value) reportAggregateScope.value = "";
  renderReportBuilder();
}, "report project filter");
safeOn(reportRoleFilter, "change", () => renderReportBuilder(), "report role filter");
safeOn(resourceSaveAllBtn, "click", () => saveAllResources().catch((e) => setStatus(e.message)), "save resources");
safeOn(projectAdminSearch, "input", renderProjectsAdmin, "project admin search");
safeOn(projectAdminNewBtn, "click", startNewProjectAdmin, "project admin new");
safeOn(projectAdminSaveBtn, "click", () => saveProjectAdmin().catch((e) => setStatus(e.message)), "project admin save");
safeOn(projectDemandAddRoleBtn, "click", () => addDemandRoleToProject().catch((e) => setStatus(e.message)), "project demand add role");
safeOn(projectDemandSaveBtn, "click", () => saveProjectDemandMatrix().catch((e) => setStatus(e.message)), "project demand save");
safeOn(assignWeekFrom, "change", () => {
  renderAssignResourceOptions();
  renderResourcePool();
  renderSelectionAllocations();
}, "assign week from");
safeOn(assignWeekTo, "change", () => {
  renderAssignResourceOptions();
  renderSelectionAllocations();
}, "assign week to");
safeOn(assignUseExternal, "change", () => {
  renderAssignResourceOptions();
  renderResourcePool();
  renderSelectionAllocations();
}, "assign use external");
safeOn(assignExternalQty, "change", () => {
  const qty = Math.max(1, parseIntOr(assignExternalQty.value, 1));
  assignExternalQty.value = String(qty);
}, "assign external qty");
safeOn(demandSaveBtn, "click", () => handleDemandSave().catch((e) => setStatus(e.message)), "demand save");
safeOn(assignSaveBtn, "click", () => handleAssignSave().catch((e) => setStatus(e.message)), "assign save");
safeOn(markUnavailableBtn, "click", () => handleMarkUnavailable().catch((e) => setStatus(e.message)), "mark unavailable");
safeOn(assignClearBtn, "click", clearSelection, "assign clear");
safeOn(selectionAllocations, "keydown", (ev) => {
  if (ev.key !== "Delete" && ev.key !== "Backspace") return;
  if (!selectedAllocationIds.size) return;
  ev.preventDefault();
  deleteSelectedAllocations().catch((e) => setStatus(e.message));
}, "selection allocations delete");

document.addEventListener("click", (ev) => {
  if (!resourceColumnsPanel || !resourceColumnsBtn) return;
  if (resourceColumnsPanel.contains(ev.target) || resourceColumnsBtn.contains(ev.target)) return;
  resourceColumnsPanel.classList.add("hidden");
});
 
safeInit("sheet tabs", bindSheetTabs);
safeInit("workpane layout", bindWorkpaneLayout);
safeInit("panel expanders", bindPanelExpanders);
safeInit("planner side toggle", bindPlannerSideToggle);
safeInit("planner side auto expand", bindPlannerSideAutoExpand);
safeInit("gantt modal", bindGanttModal);
safeInit("gantt split resize", bindGanttSplitResize);
safeInit("gantt modal drag", () => bindDraggableModal(ganttActionCard, ganttActionDragHandle));
safeInit("planner floating drag", () => bindDraggableFloatingCard(plannerSideCard, assignCardDragHandle));
safeInit("gantt floating drag", () => bindDraggableFloatingCard(ganttCard, ganttCardDragHandle));
safeInit("keyboard shortcuts", bindKeyboardShortcuts);
safeInit("report builder state", loadReportBuilderState);
safeInit("gantt external detail state", loadGanttExternalDetailFlag);
loadAll().catch((err) => {
  console.error(err);
  setStatus(`Errore caricamento: ${err.message || err}`);
});
