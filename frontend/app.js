const root = document.documentElement;
const verticalSplitter = document.getElementById("verticalSplitter");
const horizontalSplitter = document.getElementById("horizontalSplitter");
const sideInnerSplitter = document.getElementById("sideInnerSplitter");
const plannerGridWrap = document.getElementById("plannerGridWrap");
const plannerBody = document.getElementById("plannerBody");

const plannerHead = document.querySelector("thead");

const selectionBox = document.getElementById("selectionBox");
const detailProject = document.getElementById("detailProject");
const detailRole = document.getElementById("detailRole");
const detailWeekFrom = document.getElementById("detailWeekFrom");
const detailWeekTo = document.getElementById("detailWeekTo");
const detailRange = document.getElementById("detailRange");
const detailRequired = document.getElementById("detailRequired");
const detailAllocated = document.getElementById("detailAllocated");
const detailDiff = document.getElementById("detailDiff");

const modeDemandBtn = document.getElementById("modeDemandBtn");
const modeResourcesBtn = document.getElementById("modeResourcesBtn");
const demandPanel = document.getElementById("demandPanel");
const resourcesPanel = document.getElementById("resourcesPanel");
const saveDemandBtn = document.getElementById("saveDemandBtn");

const resourceSearchInput = document.getElementById("resourceSearchInput");
const showInactiveToggle = document.getElementById("showInactiveToggle");
const assignedResourceList = document.getElementById("assignedResourceList");
const availableResourceList = document.getElementById("availableResourceList");

const API_BASE = "http://127.0.0.1:8000";
const CURRENT_WEEK = 17;
const FIRST_WEEK = 1;
const LAST_WEEK = 43;
const WEEKS = Array.from({ length: LAST_WEEK - FIRST_WEEK + 1 }, (_, i) => FIRST_WEEK + i);

let verticalDrag = null;
let horizontalDrag = null;
let sideInnerDrag = null;

let resourcesData = [];
let projectsData = [];
let demandsData = [];
let allocationsData = [];
let allocationHistoryData = [];
let demandHistoryData = [];

let rowMetaMap = new Map();
let selectedCells = [];
let activeMode = "demand";
let extSequenceMap = new Map();

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function startVerticalDrag(event) {
  verticalDrag = {
    startX: event.clientX,
    startWidth: parseInt(getComputedStyle(root).getPropertyValue("--side-w"), 10) || 470,
  };
  if (verticalSplitter) verticalSplitter.classList.add("dragging");
  document.body.style.userSelect = "none";
}

function startHorizontalDrag(event) {
  horizontalDrag = {
    startY: event.clientY,
    startHeight: parseInt(getComputedStyle(root).getPropertyValue("--bottom-h"), 10) || 52,
  };
  if (horizontalSplitter) horizontalSplitter.classList.add("dragging");
  document.body.style.userSelect = "none";
}

function startSideInnerDrag(event) {
  sideInnerDrag = {
    startY: event.clientY,
    startHeight: parseInt(getComputedStyle(root).getPropertyValue("--side-top-h"), 10) || 640,
  };
  if (sideInnerSplitter) sideInnerSplitter.classList.add("dragging");
  document.body.style.userSelect = "none";
}

function onPointerMove(event) {
  if (verticalDrag) {
    const delta = verticalDrag.startX - event.clientX;
    const nextWidth = clamp(verticalDrag.startWidth + delta, 260, 760);
    root.style.setProperty("--side-w", `${nextWidth}px`);
  }

  if (horizontalDrag) {
    const delta = horizontalDrag.startY - event.clientY;
    const nextHeight = clamp(horizontalDrag.startHeight + delta, 36, 180);
    root.style.setProperty("--bottom-h", `${nextHeight}px`);
  }

  if (sideInnerDrag) {
    const delta = event.clientY - sideInnerDrag.startY;
    const nextHeight = clamp(sideInnerDrag.startHeight + delta, 150, 720);
    root.style.setProperty("--side-top-h", `${nextHeight}px`);
  }
}

function stopDrag() {
  verticalDrag = null;
  horizontalDrag = null;
  sideInnerDrag = null;
  if (verticalSplitter) verticalSplitter.classList.remove("dragging");
  if (horizontalSplitter) horizontalSplitter.classList.remove("dragging");
  if (sideInnerSplitter) sideInnerSplitter.classList.remove("dragging");
  document.body.style.userSelect = "";
}

if (verticalSplitter) verticalSplitter.addEventListener("pointerdown", startVerticalDrag);
if (horizontalSplitter) horizontalSplitter.addEventListener("pointerdown", startHorizontalDrag);
if (sideInnerSplitter) sideInnerSplitter.addEventListener("pointerdown", startSideInnerDrag);

window.addEventListener("pointermove", onPointerMove);
window.addEventListener("pointerup", stopDrag);

async function fetchJson(path, options = {}) {
  const response = await fetch(`${API_BASE}${path}`, options);
  if (!response.ok) {
    throw new Error(`Errore ${response.status} su ${path}`);
  }
  return response.json();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function getWeekStartDate(year, week) {
  const jan4 = new Date(year, 0, 4);
  const day = jan4.getDay() || 7;
  const mondayWeek1 = new Date(jan4);
  mondayWeek1.setDate(jan4.getDate() - day + 1);

  const result = new Date(mondayWeek1);
  result.setDate(mondayWeek1.getDate() + (week - 1) * 7);
  return result;
}

function formatDate(date) {
  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  return `${dd}/${mm}`;
}

function getWeekRangeLabel(weekFrom, weekTo = weekFrom) {
  const year = 2026;
  const start = getWeekStartDate(year, weekFrom);
  const end = new Date(getWeekStartDate(year, weekTo));
  end.setDate(end.getDate() + 6);
  return `${formatDate(start)} - ${formatDate(end)}`;
}

function applyWeekTooltips() {
  document.querySelectorAll(".week-head").forEach((cell) => {
    const text = (cell.textContent || "").trim().toUpperCase();
    const match = text.match(/^W(\d{1,2})$/);
    if (!match) return;
    const week = Number(match[1]);
    cell.title = getWeekRangeLabel(week);
  });
}

function scrollToCurrentWeek() {
  if (!plannerGridWrap) return;
  const current = document.querySelector(".second-line.current-week, .week-head.current-week");
  if (!current) return;

  const leftStickyWidth = 140 + 160 + 56;
  plannerGridWrap.scrollLeft = Math.max(0, current.offsetLeft - leftStickyWidth - 20);
}

function roundOneDecimal(value) {
  return Math.round(Number(value || 0) * 10) / 10;
}

function formatNumber(value) {
  const rounded = roundOneDecimal(Number(value || 0));
  if (Number.isInteger(rounded)) return String(rounded);
  return String(rounded).replace(".", ",");
}

function formatLoad(loadPercent) {
  const load = Number(loadPercent || 0);
  if (load >= 99.9) return "1:1";
  if (load > 0 && load <= 50.1) return "1/2";
  return `${formatNumber(load / 100)}`;
}

function formatAllocatedDisplay(internalAllocated, externalAllocated) {
  const internal = roundOneDecimal(Number(internalAllocated || 0));
  const external = roundOneDecimal(Number(externalAllocated || 0));

  if (external > 0) {
    return `${formatNumber(internal)}+${formatNumber(external)}`;
  }

  return formatNumber(internal);
}

function shortResourceName(fullName) {
  const clean = String(fullName || "").trim().replace(/\s+/g, " ");
  if (!clean) return "-";

  const parts = clean.split(" ");
  if (parts.length === 1) return parts[0].toUpperCase();

  const surname = parts.slice(0, -1).join(" ").toUpperCase();
  const firstName = parts[parts.length - 1].toUpperCase();
  const first3 = firstName.slice(0, 3);

  return `${surname} ${first3}`;
}

function normalizeRole(value) {
  return String(value || "").trim().toUpperCase();
}

function normalizeExtRoleLabel(role) {
  return normalizeRole(role)
    .replace(/\s+/g, " ")
    .replace(/\s+EXT$/g, "")
    .replace(/-EXT$/g, "")
    .trim();
}

function isExternalResource(resource) {
  if (!resource) return false;

  const name = normalizeRole(resource.name || resource.resource_name || "");
  const role = normalizeRole(resource.role || resource.resource_role || "");

  return (
    name.includes("-EXT") ||
    name.endsWith(" EXT") ||
    role.includes("-EXT") ||
    role.endsWith(" EXT")
  );
}

function getResourceById(resourceId) {
  return resourcesData.find((resource) => Number(resource.id) === Number(resourceId)) || null;
}

function getAllocationResource(allocation) {
  return getResourceById(allocation.resource_id) || {
    id: allocation.resource_id,
    name: allocation.resource_name || allocation.resourceName || "",
    role: allocation.resource_role || allocation.resourceRole || "",
  };
}

function mansioneLabel(rowRole, resourceRole) {
  const row = normalizeRole(rowRole);
  const res = normalizeRole(resourceRole);

  if (!res || res === row) return "ok mans";
  return `no mans ${res}`;
}

function getExtSequenceKey(role, week) {
  return `${normalizeExtRoleLabel(role)}__${Number(week)}`;
}

function getExtDisplayName(allocation, rowRole, week) {
  const allocationId = Number(allocation.id || allocation.history_id || allocation.allocation_id || 0);
  const role = normalizeExtRoleLabel(rowRole || allocation.role || allocation.resource_role || "EXT");
  const key = getExtSequenceKey(role, week);

  if (!extSequenceMap.has(key)) {
    extSequenceMap.set(key, {
      next: 1,
      byAllocationId: new Map(),
    });
  }

  const bucket = extSequenceMap.get(key);

  if (allocationId && bucket.byAllocationId.has(allocationId)) {
    return `${role}${bucket.byAllocationId.get(allocationId)}-EXT`;
  }

  const current = bucket.next;
  bucket.next += 1;

  if (allocationId) {
    bucket.byAllocationId.set(allocationId, current);
  }

  return `${role}${current}-EXT`;
}

function formatAssignmentLine(item, rowRole, week = null) {
  const resource = getAllocationResource(item);
  const isExt = isExternalResource(resource) || isExternalResource(item);

  if (isExt) {
    const extName = getExtDisplayName(item, rowRole, week || item.week);
    const load = formatLoad(item.load_percent || item.loadPercent || 100);
    return `${extName} | EXT | ${load}`;
  }

  const name = shortResourceName(resource.name || item.resource_name || item.resourceName);
  const realRole = resource.role || item.resource_role || item.resourceRole || "";
  const mans = mansioneLabel(rowRole, realRole);
  const load = formatLoad(item.load_percent || item.loadPercent);
  return `${name} | ${mans} | ${load}`;
}

function buildDemandMap(demands) {
  const map = new Map();
  for (const demand of demands) {
    const role = normalizeRole(demand.role);
    const key = `${Number(demand.project_id)}__${role}__${Number(demand.week)}`;
    map.set(key, Number(demand.quantity || 0));
  }
  return map;
}

function buildDemandRowSet(demands) {
  const set = new Set();
  for (const demand of demands) {
    const role = normalizeRole(demand.role);
    set.add(`${Number(demand.project_id)}__${role}`);
  }
  return set;
}

function parseHiringStartWeekFromNote(note) {
  const text = normalizeRole(note);
  const match = text.match(/ASSUNZIONE\s+DA\s+W?\s*(\d{1,2})/);
  if (!match) return null;
  return Number(match[1]);
}

function isExplicitlyUnavailable(resource) {
  const text = normalizeRole(resource.availability_note || "");
  return (
    text.includes("INDISP") ||
    text.includes("INDISPONIBILE") ||
    text.includes("NON DISPONIBILE")
  );
}

function isResourceAvailableForWeek(resource, week) {
  if (isExternalResource(resource)) return true;

  if (!resource.is_active) return false;
  if (isExplicitlyUnavailable(resource)) return false;

  const startWeek = parseHiringStartWeekFromNote(resource.availability_note);
  if (startWeek !== null && week < startWeek) return false;

  return true;
}

function buildAllocationMaps(allocations, resources) {
  const resourceById = new Map(resources.map((r) => [Number(r.id), r]));
  const demandRowSet = buildDemandRowSet(demandsData);
  const totalMap = new Map();
  const internalMap = new Map();
  const externalMap = new Map();

  for (const allocation of allocations) {
    const resource = resourceById.get(Number(allocation.resource_id));
    if (!resource) continue;

    if (!isResourceAvailableForWeek(resource, Number(allocation.week))) continue;

    const role = normalizeRole(allocation.role || resource.role || "");
    const rowKey = `${Number(allocation.project_id)}__${role}`;

    if (!demandRowSet.has(rowKey)) continue;

    const key = `${Number(allocation.project_id)}__${role}__${Number(allocation.week)}`;
    const value = Number(allocation.load_percent || 0) / 100;

    totalMap.set(key, (totalMap.get(key) || 0) + value);

    if (isExternalResource(resource)) {
      externalMap.set(key, (externalMap.get(key) || 0) + value);
    } else {
      internalMap.set(key, (internalMap.get(key) || 0) + value);
    }
  }

  return {
    totalMap,
    internalMap,
    externalMap,
  };
}

function buildHistoryMap(historyRows) {
  const map = new Map();

  for (const row of historyRows) {
    const role = normalizeRole(row.role);
    const key = `${Number(row.project_id)}__${role}__${Number(row.week)}`;
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(row);
  }

  return map;
}

function buildPlannerRows(projects, demands) {
  const projectsById = new Map(projects.map((p) => [Number(p.id), p]));
  const rowsMap = new Map();

  for (const demand of demands) {
    const project = projectsById.get(Number(demand.project_id));
    if (!project) continue;

    const role = normalizeRole(demand.role);
    const rowKey = `${Number(demand.project_id)}__${role}`;

    if (!rowsMap.has(rowKey)) {
      rowsMap.set(rowKey, {
        row_key: rowKey,
        project_id: Number(demand.project_id),
        project_name: project.name,
        role,
      });
    }
  }

  const rows = Array.from(rowsMap.values()).sort((a, b) => {
    if (a.project_name !== b.project_name) return a.project_name.localeCompare(b.project_name);
    return a.role.localeCompare(b.role);
  });

  rowMetaMap = new Map(rows.map((row, idx) => [idx, row]));
  return rows;
}

function getCoverageClass(required, allocated) {
  if (required === 0 && allocated === 0) return "coverage good";
  if (required === 0 && allocated > 0) return "coverage warn";
  if (allocated < required) return "coverage warn";
  if (allocated > required) return "coverage warn";
  return "coverage good";
}

function getCellClass(required, allocated, week, hasHistory, hasAllocationHistory) {
  let cls = "cell-empty planner-cell-clickable";

  if (required > 0 && allocated < required) {
    cls = "cell-demand planner-cell-clickable";
  } else if (required > 0 && allocated === required) {
    cls = "cell-ok planner-cell-clickable";
  } else if (required > 0 && allocated > required) {
    cls = "cell-surplus planner-cell-clickable";
  } else if (required === 0 && allocated > 0) {
    cls = "cell-surplus planner-cell-clickable";
  }

  if (week === CURRENT_WEEK) cls += " current-col";
  if (hasHistory || hasAllocationHistory) cls += " cell-history";

  return cls;
}

function setMode(mode) {
  activeMode = mode;

  modeDemandBtn.classList.toggle("mode-btn-active", mode === "demand");
  modeResourcesBtn.classList.toggle("mode-btn-active", mode === "resources");

  demandPanel.classList.toggle("panel-block-active", mode === "demand");
  demandPanel.classList.toggle("panel-block-collapsed", mode !== "demand");

  resourcesPanel.classList.toggle("panel-block-active", mode === "resources");
  resourcesPanel.classList.toggle("panel-block-collapsed", mode !== "resources");
}

function getSelectionSummary() {
  if (!selectedCells.length) return null;

  const first = selectedCells[0];
  const decoded = JSON.parse(decodeURIComponent(first.dataset.cell));
  const weeks = selectedCells.map((cell) => Number(cell.dataset.week)).sort((a, b) => a - b);

  const totalRequired = selectedCells.reduce((sum, cell) => {
    const data = JSON.parse(decodeURIComponent(cell.dataset.cell));
    return sum + Number(data.required || 0);
  }, 0);

  const totalAllocated = selectedCells.reduce((sum, cell) => {
    const data = JSON.parse(decodeURIComponent(cell.dataset.cell));
    return sum + Number(data.allocated || 0);
  }, 0);

  const totalInternalAllocated = selectedCells.reduce((sum, cell) => {
    const data = JSON.parse(decodeURIComponent(cell.dataset.cell));
    return sum + Number(data.internalAllocated || 0);
  }, 0);

  const totalExternalAllocated = selectedCells.reduce((sum, cell) => {
    const data = JSON.parse(decodeURIComponent(cell.dataset.cell));
    return sum + Number(data.externalAllocated || 0);
  }, 0);

  return {
    project_id: Number(decoded.project_id),
    project_name: decoded.project_name,
    role: decoded.role,
    row_key: `${Number(decoded.project_id)}__${decoded.role}`,
    week_from: weeks[0],
    week_to: weeks[weeks.length - 1],
    required: selectedCells.length === 1 ? decoded.required : totalRequired,
    allocated: totalAllocated,
    internalAllocated: totalInternalAllocated,
    externalAllocated: totalExternalAllocated,
    diff: totalRequired - totalAllocated,
  };
}

function updateSidePanelFromSelection() {
  const summary = getSelectionSummary();
  if (!summary) return;

  selectionBox.textContent =
    `${summary.project_name} | ${summary.role} | W${String(summary.week_from).padStart(2, "0")}` +
    `${summary.week_to !== summary.week_from ? ` - W${String(summary.week_to).padStart(2, "0")}` : ""}`;

  detailProject.value = summary.project_name;
  detailRole.value = summary.role;
  detailWeekFrom.value = summary.week_from;
  detailWeekTo.value = summary.week_to;
  detailRange.value = getWeekRangeLabel(summary.week_from, summary.week_to);
  detailRequired.value = formatNumber(summary.required);
  detailAllocated.value = formatAllocatedDisplay(summary.internalAllocated, summary.externalAllocated);
  detailDiff.value = formatNumber(summary.required - summary.allocated);

  renderResourceLists();
}

function clearCellSelectionVisuals() {
  plannerBody.querySelectorAll(".planner-cell-selected, .planner-cell-range").forEach((cell) => {
    cell.classList.remove("planner-cell-selected");
    cell.classList.remove("planner-cell-range");
  });
}

function applyCellSelectionVisuals() {
  clearCellSelectionVisuals();

  selectedCells.forEach((cell, index) => {
    if (index === 0) cell.classList.add("planner-cell-selected");
    else cell.classList.add("planner-cell-range");
  });
}

function selectSingleCell(cell) {
  selectedCells = [cell];
  applyCellSelectionVisuals();
  updateSidePanelFromSelection();
  cell.scrollIntoView({ block: "nearest", inline: "nearest" });
}

function extendSelection(direction) {
  if (!selectedCells.length) return;

  const anchor = selectedCells[0];
  const rowIndex = anchor.dataset.rowIndex;
  const weeks = selectedCells.map((cell) => Number(cell.dataset.week)).sort((a, b) => a - b);
  const minWeek = weeks[0];
  const maxWeek = weeks[weeks.length - 1];

  const targetWeek = direction === "right" ? maxWeek + 1 : minWeek - 1;
  if (targetWeek < FIRST_WEEK || targetWeek > LAST_WEEK) return;

  const targetCell = plannerBody.querySelector(
    `td[data-row-index="${rowIndex}"][data-week="${targetWeek}"]`
  );
  if (!targetCell) return;

  if (direction === "right") selectedCells.push(targetCell);
  else selectedCells.unshift(targetCell);

  applyCellSelectionVisuals();
  updateSidePanelFromSelection();
  targetCell.scrollIntoView({ block: "nearest", inline: "nearest" });
}

function moveSelection(deltaRow, deltaCol) {
  if (!selectedCells.length) return;

  const anchor = selectedCells[0];
  const currentRow = Number(anchor.dataset.rowIndex);
  const currentWeek = Number(anchor.dataset.week);

  const targetRow = currentRow + deltaRow;
  const targetWeek = currentWeek + deltaCol;

  const target = plannerBody.querySelector(
    `td[data-row-index="${targetRow}"][data-week="${targetWeek}"]`
  );

  if (!target) return;
  selectSingleCell(target);
}

function moveFocusToDemandPanel() {
  setMode("demand");
  if (detailRequired) {
    detailRequired.focus();
    detailRequired.select?.();
  }
}

function moveFocusToResourcesPanel() {
  setMode("resources");
  if (resourceSearchInput) {
    resourceSearchInput.focus();
    resourceSearchInput.select?.();
  }
}

function moveFocusToGrid() {
  if (plannerGridWrap) plannerGridWrap.focus();
}

function getSelectedWeeksSet() {
  const summary = getSelectionSummary();
  const selectedWeeks = new Set();

  if (!summary) return selectedWeeks;

  for (let w = summary.week_from; w <= summary.week_to; w += 1) {
    selectedWeeks.add(w);
  }

  return selectedWeeks;
}

function getResourceAllocationsForSelectedWeeks(resourceId) {
  const selectedWeeks = getSelectedWeeksSet();

  return allocationsData.filter((allocation) => {
    return (
      Number(allocation.resource_id) === Number(resourceId) &&
      selectedWeeks.has(Number(allocation.week))
    );
  });
}

function getResourceAllocationsForWeek(resourceId, week) {
  return allocationsData.filter((allocation) => {
    return Number(allocation.resource_id) === Number(resourceId) && Number(allocation.week) === Number(week);
  });
}

function getAssignedResourcesForSelection() {
  const summary = getSelectionSummary();
  if (!summary) return [];

  const selectedWeeks = getSelectedWeeksSet();

  return resourcesData.filter((resource) => {
    if (!isResourceAvailableForWeek(resource, summary.week_from)) return false;

    return allocationsData.some((allocation) => {
      const allocationRole = normalizeRole(allocation.role || "");

      return (
        Number(allocation.resource_id) === Number(resource.id) &&
        Number(allocation.project_id) === Number(summary.project_id) &&
        allocationRole === summary.role &&
        selectedWeeks.has(Number(allocation.week))
      );
    });
  });
}

function getResourceStatus(resource) {
  const summary = getSelectionSummary();
  if (!summary) return "free";

  if (isExternalResource(resource)) return "free";

  if (!resource.is_active) return "inactive";
  if (isExplicitlyUnavailable(resource)) return "unavailable";

  const startWeek = parseHiringStartWeekFromNote(resource.availability_note);
  if (startWeek !== null && summary.week_from < startWeek) return "future";

  const assigned = getAssignedResourcesForSelection().some((r) => Number(r.id) === Number(resource.id));
  if (assigned) return "allocated";

  const demandRowSet = buildDemandRowSet(demandsData);

  const allocations = getResourceAllocationsForSelectedWeeks(resource.id).filter((allocation) => {
    const allocationRole = normalizeRole(allocation.role || "");
    const rowKey = `${Number(allocation.project_id)}__${allocationRole}`;
    return demandRowSet.has(rowKey);
  });

  const maxPerWeek = new Map();

  allocations.forEach((allocation) => {
    const week = Number(allocation.week);
    maxPerWeek.set(week, (maxPerWeek.get(week) || 0) + 1);
  });

  const values = Array.from(maxPerWeek.values());

  if (values.some((count) => count >= 2)) return "saturated";
  if (values.some((count) => count === 1)) return "partial";

  const historical = getResourceAllocationsForSelectedWeeks(resource.id).filter((allocation) => {
    const allocationRole = normalizeRole(allocation.role || "");
    const rowKey = `${Number(allocation.project_id)}__${allocationRole}`;
    return !demandRowSet.has(rowKey);
  });

  if (historical.length > 0) return "history";

  return "free";
}

function allocationShortLabel(allocation) {
  const project = allocation.project_name || `Progetto ${allocation.project_id}`;
  const role = allocation.role || "-";
  const load = Number(allocation.load_percent || 0);
  return `${project} ${role} W${String(allocation.week).padStart(2, "0")} ${load}%`;
}

function getAllocationLocationText(resource, onlyHistorical = false) {
  const demandRowSet = buildDemandRowSet(demandsData);

  const allocations = getResourceAllocationsForSelectedWeeks(resource.id).filter((allocation) => {
    const allocationRole = normalizeRole(allocation.role || "");
    const rowKey = `${Number(allocation.project_id)}__${allocationRole}`;
    const isHistorical = !demandRowSet.has(rowKey);
    return onlyHistorical ? isHistorical : !isHistorical;
  });

  if (!allocations.length) return "";
  return allocations.map(allocationShortLabel).join(" / ");
}

function getResourceShortInfo(resource, status) {
  if (isExternalResource(resource)) return "EXT";
  if (status === "inactive") return "CESSATO";
  if (status === "unavailable") return "INDISP";
  if (status === "future") {
    const startWeek = parseHiringStartWeekFromNote(resource.availability_note);
    return `ASSUNZIONE W${String(startWeek).padStart(2, "0")}`;
  }

  if (status === "allocated") {
    const where = getAllocationLocationText(resource);
    return where ? `ASSEGNATO: ${where}` : "ASSEGNATO";
  }

  if (status === "partial") {
    const where = getAllocationLocationText(resource);
    return where ? `PARZIALE: ${where}` : "PARZIALE";
  }

  if (status === "saturated") {
    const where = getAllocationLocationText(resource);
    return where ? `SATURO: ${where}` : "SATURO";
  }

  if (status === "history") {
    const where = getAllocationLocationText(resource, true);
    return where ? `STORICO: ${where}` : "STORICO";
  }

  return "";
}

function getResourceTooltip(resource, status) {
  if (isExternalResource(resource)) return "Risorsa esterna virtuale";

  if (status === "future") {
    const startWeek = parseHiringStartWeekFromNote(resource.availability_note);
    return `Assunzione da W${String(startWeek).padStart(2, "0")}`;
  }

  if (status === "history") return getAllocationLocationText(resource, true);
  if (status !== "partial" && status !== "saturated" && status !== "allocated") return "";

  return getAllocationLocationText(resource);
}

function matchesResourceSearch(resource, search) {
  if (!search) return true;
  const haystack = `${resource.name} ${resource.role} ${resource.availability_note || ""}`.toLowerCase();
  return haystack.includes(search.toLowerCase());
}

function renderResourceItem(resource, status) {
  const classes = ["resource-item"];

  if (status === "allocated") classes.push("resource-allocated");
  if (status === "free") classes.push("resource-free");
  if (status === "unavailable") classes.push("resource-unavailable");
  if (status === "future") classes.push("resource-unavailable");
  if (status === "inactive") classes.push("resource-inactive");
  if (status === "partial") classes.push("resource-partial");
  if (status === "saturated") classes.push("resource-allocated");
  if (status === "history") classes.push("resource-history");
  if (isExternalResource(resource)) classes.push("resource-history");

  const tooltip = getResourceTooltip(resource, status);
  const extra = getResourceShortInfo(resource, status);

  return `
    <div
      class="${classes.join(" ")}"
      data-resource-id="${resource.id}"
      data-resource-status="${status}"
      title="${escapeHtml(tooltip)}"
    >
      ${escapeHtml(resource.name)} | ${escapeHtml(resource.role || "-")}${extra ? ` | ${escapeHtml(extra)}` : ""}
    </div>
  `;
}

function showMessageDialog(title, message) {
  showConflictDialog(
    title,
    `<p>${escapeHtml(message)}</p>`,
    [{ label: "OK", primary: true, handler: async () => {} }],
  );
}

function renderResourceLists() {
  if (!assignedResourceList || !availableResourceList) return;

  const summary = getSelectionSummary();
  if (!summary) {
    assignedResourceList.innerHTML = "";
    availableResourceList.innerHTML = "";
    return;
  }

  const assigned = getAssignedResourcesForSelection();
  const assignedIds = new Set(assigned.map((r) => Number(r.id)));

  const search = resourceSearchInput?.value?.trim() || "";
  const showInactive = !!showInactiveToggle?.checked;

  const visibleAssigned = assigned.filter((resource) => matchesResourceSearch(resource, search));

  const visibleAvailable = resourcesData.filter((resource) => {
    if (assignedIds.has(Number(resource.id))) return false;
    if (!showInactive && !resource.is_active && !isExternalResource(resource)) return false;
    return matchesResourceSearch(resource, search);
  });

  assignedResourceList.innerHTML = visibleAssigned.length
    ? visibleAssigned.map((resource) => renderResourceItem(resource, "allocated")).join("")
    : `<div class="resource-item resource-unavailable">Nessuna risorsa assegnata</div>`;

  availableResourceList.innerHTML = visibleAvailable.length
    ? visibleAvailable.map((resource) => renderResourceItem(resource, getResourceStatus(resource))).join("")
    : `<div class="resource-item resource-unavailable">Nessuna risorsa disponibile</div>`;

  assignedResourceList.querySelectorAll("[data-resource-id]").forEach((item) => {
    item.addEventListener("dblclick", async () => {
      const resourceId = Number(item.dataset.resourceId);
      await removeResourceFromSelection(resourceId);
    });
  });

  availableResourceList.querySelectorAll("[data-resource-id]").forEach((item) => {
    item.addEventListener("dblclick", async () => {
      const status = item.datasetResourceStatus || item.dataset.resourceStatus;
      if (status === "inactive" || status === "unavailable" || status === "future") return;

      const resourceId = Number(item.dataset.resourceId);
      await handleAvailableResourceDoubleClick(resourceId);
    });
  });
}

function showConflictDialog(title, bodyHtml, actions) {
  closeConflictDialog();

  const overlay = document.createElement("div");
  overlay.className = "conflict-overlay";
  overlay.id = "conflictOverlay";

  const dialog = document.createElement("div");
  dialog.className = "conflict-dialog";

  const actionButtons = actions.map((action, index) => {
    return `<button class="btn ${action.primary ? "btn-primary" : "btn-light"}" data-action-index="${index}" type="button">${escapeHtml(action.label)}</button>`;
  }).join("");

  dialog.innerHTML = `
    <div class="conflict-title">${escapeHtml(title)}</div>
    <div class="conflict-body">${bodyHtml}</div>
    <div class="conflict-actions">${actionButtons}</div>
  `;

  overlay.appendChild(dialog);
  document.body.appendChild(overlay);

  dialog.querySelectorAll("[data-action-index]").forEach((button) => {
    button.addEventListener("click", async () => {
      const index = Number(button.dataset.actionIndex);
      const action = actions[index];
      closeConflictDialog();
      await action.handler();
    });
  });
}

function closeConflictDialog() {
  const existing = document.getElementById("conflictOverlay");
  if (existing) existing.remove();
}

function allocationLabel(allocation) {
  return allocationShortLabel(allocation);
}

function selectedRowExistsInPlanner() {
  const summary = getSelectionSummary();
  if (!summary) return false;

  const demandRowSet = buildDemandRowSet(demandsData);
  return demandRowSet.has(summary.row_key);
}

async function handleAvailableResourceDoubleClick(resourceId) {
  const summary = getSelectionSummary();
  if (!summary) return;

  if (!selectedRowExistsInPlanner()) {
    showMessageDialog(
      "Riga non attiva",
      "Prima crea/attiva il fabbisogno per questa commessa e mansione, poi assegna la risorsa.",
    );
    return;
  }

  const resource = resourcesData.find((item) => Number(item.id) === Number(resourceId));
  if (!resource) return;

  if (isExternalResource(resource)) {
    await assignResourceToSelection(resourceId);
    return;
  }

  if (summary.week_from !== summary.week_to) {
    const demandRowSet = buildDemandRowSet(demandsData);
    const conflicts = getResourceAllocationsForSelectedWeeks(resourceId).filter((allocation) => {
      const allocationRole = normalizeRole(allocation.role || "");
      const rowKey = `${Number(allocation.project_id)}__${allocationRole}`;
      return demandRowSet.has(rowKey);
    });

    if (conflicts.length > 0) {
      showConflictDialog(
        "Conflitto su più settimane",
        "La risorsa è già allocata in almeno una delle settimane selezionate.<br>Gestisci una settimana alla volta.",
        [{ label: "OK", primary: true, handler: async () => {} }],
      );
      return;
    }

    await assignResourceToSelection(resourceId);
    return;
  }

  const week = summary.week_from;
  const demandRowSet = buildDemandRowSet(demandsData);
  const existing = getResourceAllocationsForWeek(resourceId, week).filter((allocation) => {
    const allocationRole = normalizeRole(allocation.role || "");
    const rowKey = `${Number(allocation.project_id)}__${allocationRole}`;
    return demandRowSet.has(rowKey);
  });

  if (existing.length === 0) {
    await resolveAllocationConflict(resourceId, "direct");
    return;
  }

  if (existing.length === 1) {
    const old = existing[0];

    showConflictDialog(
      "Risorsa già allocata",
      `
        <p><strong>${escapeHtml(resource.name)}</strong> è già allocato su:</p>
        <p>${escapeHtml(allocationLabel(old))}</p>
        <p>Cosa vuoi fare?</p>
      `,
      [
        { label: "Annulla", handler: async () => {} },
        {
          label: "Sostituisci",
          primary: true,
          handler: async () => {
            await resolveAllocationConflict(resourceId, "replace_all");
          },
        },
        {
          label: "Forza 50/50",
          handler: async () => {
            await resolveAllocationConflict(resourceId, "split_50");
          },
        },
      ],
    );

    return;
  }

  const body = `
    <p><strong>${escapeHtml(resource.name)}</strong> è già allocato su 2 commesse:</p>
    ${existing.map((item) => `<p>${escapeHtml(allocationLabel(item))}</p>`).join("")}
    <p>Scegli cosa tenere.</p>
  `;

  const actions = [
    { label: "Annulla", handler: async () => {} },
    ...existing.map((item) => ({
      label: `Togli ${item.project_name}`,
      handler: async () => {
        await resolveAllocationConflict(resourceId, "replace_one", item.id);
      },
    })),
    {
      label: "Togli entrambe e assegna 100%",
      primary: true,
      handler: async () => {
        await resolveAllocationConflict(resourceId, "replace_all");
      },
    },
  ];

  showConflictDialog("Risorsa satura", body, actions);
}

async function resolveAllocationConflict(resourceId, mode, removeAllocationId = null) {
  const summary = getSelectionSummary();
  if (!summary) return;

  const result = await fetchJson("/api/allocations/resolve-conflict", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      resource_id: Number(resourceId),
      project_id: Number(summary.project_id),
      role: summary.role,
      week: Number(summary.week_from),
      mode,
      remove_allocation_id: removeAllocationId,
      hours: 40,
      note: "Risoluzione conflitto da planner V2",
    }),
  });

  if (result && result.ok === false && result.reason === "demand_row_missing") {
    showMessageDialog(
      "Riga non attiva",
      result.message || "Prima crea/attiva il fabbisogno per questa commessa e mansione.",
    );
  }

  await loadAndRenderPlanner();
  setMode("resources");
}

function handleGridKeydown(event) {
  if (!selectedCells.length) return;

  switch (event.key) {
    case "ArrowRight":
      event.preventDefault();
      if (event.ctrlKey) extendSelection("right");
      else moveSelection(0, 1);
      break;
    case "ArrowLeft":
      event.preventDefault();
      if (event.ctrlKey) extendSelection("left");
      else moveSelection(0, -1);
      break;
    case "ArrowDown":
      event.preventDefault();
      moveSelection(1, 0);
      break;
    case "ArrowUp":
      event.preventDefault();
      moveSelection(-1, 0);
      break;
    case "Enter":
      event.preventDefault();
      if (activeMode === "resources") moveFocusToResourcesPanel();
      else moveFocusToDemandPanel();
      break;
    case "Tab":
      event.preventDefault();
      if (activeMode === "demand") setMode("resources");
      else setMode("demand");
      break;
    case "Delete":
    case "Backspace":
      event.preventDefault();
      break;
    default:
      break;
  }
}

function handleDetailKeydown(event) {
  if (event.key === "Escape") {
    event.preventDefault();
    moveFocusToGrid();
  }

  if (event.key === "Tab" && event.shiftKey && document.activeElement === detailRequired) {
    event.preventDefault();
    moveFocusToGrid();
  }
}

async function saveDemandRange() {
  const summary = getSelectionSummary();
  if (!summary) return;

  const payload = {
    project_id: Number(summary.project_id),
    role: summary.role,
    week_from: Number(detailWeekFrom.value),
    week_to: Number(detailWeekTo.value),
    quantity: Number(String(detailRequired.value).replace(",", ".")),
    note: "Salvataggio da planner V2",
  };

  await fetchJson("/api/demands/upsert-range", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  await loadAndRenderPlanner();
  moveFocusToGrid();
}

async function assignResourceToSelection(resourceId) {
  const summary = getSelectionSummary();
  if (!summary) return;

  const result = await fetchJson("/api/allocations/assign-range", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      resource_id: Number(resourceId),
      project_id: Number(summary.project_id),
      role: summary.role,
      week_from: Number(summary.week_from),
      week_to: Number(summary.week_to),
      hours: 40,
      load_percent: 100,
      note: "Assegnazione da planner V2",
    }),
  });

  if (result && result.message && Array.isArray(result.skipped) && result.skipped.includes("demand_row_missing")) {
    showMessageDialog("Riga non attiva", result.message);
  }

  await loadAndRenderPlanner();
  setMode("resources");
}

async function removeResourceFromSelection(resourceId) {
  const summary = getSelectionSummary();
  if (!summary) return;

  await fetchJson("/api/allocations/remove-range", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      resource_id: Number(resourceId),
      project_id: Number(summary.project_id),
      role: summary.role,
      week_from: Number(summary.week_from),
      week_to: Number(summary.week_to),
      hours: 40,
      load_percent: 100,
      note: "Rimozione da planner V2",
    }),
  });

  await loadAndRenderPlanner();
  setMode("resources");
}

function getCellAssignments(projectId, role, week) {
  const normalizedRole = normalizeRole(role);

  return allocationsData.filter((item) => {
    return (
      Number(item.project_id) === Number(projectId) &&
      normalizeRole(item.role || "") === normalizedRole &&
      Number(item.week) === Number(week)
    );
  });
}

function buildCellCalloutHtml(projectName, role, week, assignments, released, demandHistoryRows) {
  const parts = [];

  parts.push(`<div class="cell-callout-title">${escapeHtml(projectName)} | ${escapeHtml(role)} | W${String(week).padStart(2, "0")}</div>`);

  if (assignments.length) {
    parts.push(`<div class="cell-callout-section">Assegnati</div>`);
    assignments.forEach((item) => {
      parts.push(`<div class="cell-callout-line">${escapeHtml(formatAssignmentLine(item, role, week))}</div>`);
    });
  }

  if (released.length) {
    parts.push(`<div class="cell-callout-section">Non più conteggiati</div>`);
    released.forEach((item) => {
      const line = `${formatAssignmentLine(item, role, week)} | ${item.reason || "storico"}`;
      parts.push(`<div class="cell-callout-line cell-callout-released">${escapeHtml(line)}</div>`);
    });
  }

  if (demandHistoryRows.length) {
    const latest = demandHistoryRows[0];
    parts.push(`<div class="cell-callout-section">Storico fabbisogno</div>`);
    parts.push(
      `<div class="cell-callout-line">${escapeHtml(latest.old_quantity)} → ${escapeHtml(latest.new_quantity)} | ${escapeHtml(latest.created_at || "")}</div>`
    );
  }

  if (!assignments.length && !released.length && !demandHistoryRows.length) {
    parts.push(`<div class="cell-callout-line muted">Nessun dettaglio operativo.</div>`);
  }

  return parts.join("");
}

function showCellCallout(event, html) {
  hideCellCallout();

  const callout = document.createElement("div");
  callout.className = "cell-callout";
  callout.id = "cellCallout";
  callout.innerHTML = html;

  document.body.appendChild(callout);

  const margin = 12;
  const rect = callout.getBoundingClientRect();
  let left = event.clientX + margin;
  let top = event.clientY + margin;

  if (left + rect.width > window.innerWidth - 8) {
    left = event.clientX - rect.width - margin;
  }

  if (top + rect.height > window.innerHeight - 8) {
    top = event.clientY - rect.height - margin;
  }

  callout.style.left = `${Math.max(8, left)}px`;
  callout.style.top = `${Math.max(8, top)}px`;
}

function hideCellCallout() {
  const existing = document.getElementById("cellCallout");
  if (existing) existing.remove();
}

function computeWeekTotals(rows, demandMap, allocationMaps) {
  const totals = new Map();

  WEEKS.forEach((week) => {
    totals.set(week, {
      week,
      required: 0,
      internalAllocated: 0,
      externalAllocated: 0,
      allocated: 0,
      diff: 0,
    });
  });

  rows.forEach((row) => {
    WEEKS.forEach((week) => {
      const key = `${row.project_id}__${row.role}__${week}`;
      const required = demandMap.get(key) || 0;
      const internalAllocated = allocationMaps.internalMap.get(key) || 0;
      const externalAllocated = allocationMaps.externalMap.get(key) || 0;
      const allocated = internalAllocated + externalAllocated;

      const total = totals.get(week);
      total.required += required;
      total.internalAllocated += internalAllocated;
      total.externalAllocated += externalAllocated;
      total.allocated += allocated;
      total.diff = total.required - total.allocated;
    });
  });

  return totals;
}

function renderPlannerHeader(rows, demandMap, allocationMaps) {
  const head = plannerHead || document.querySelector("thead");
  if (!head) return;

  const totals = computeWeekTotals(rows, demandMap, allocationMaps);

  head.innerHTML = `
    <tr class="planner-subtotals-row">
      <th class="sticky-col planner-subtotals-label">Subtotali</th>
      <th class="sticky-col second planner-subtotals-label">Vista</th>
      <th class="sticky-col third planner-subtotals-label">Tot</th>
      ${WEEKS.map((week) => {
        const currentClass = week === CURRENT_WEEK ? " current-week" : "";

        if (week < CURRENT_WEEK) {
          return `<th class="week-head week-subtotal-head week-subtotal-past${currentClass}" title="${getWeekRangeLabel(week)}"></th>`;
        }

        const total = totals.get(week);
        const allocatedDisplay = formatAllocatedDisplay(total.internalAllocated, total.externalAllocated);

        return `
          <th class="week-head week-subtotal-head${currentClass}" title="${getWeekRangeLabel(week)}">
            <span class="metric">R${formatNumber(total.required)}</span>
            <span class="metric">A${allocatedDisplay}</span>
            <span class="metric">D${formatNumber(total.diff)}</span>
          </th>
        `;
      }).join("")}
    </tr>
    <tr>
      <th class="sticky-col">Commessa</th>
      <th class="sticky-col second">Mansione</th>
      <th class="sticky-col third">Cop.</th>
      ${WEEKS.map((week) => {
        const currentClass = week === CURRENT_WEEK ? " current-week" : "";
        return `<th class="week-head second-line${currentClass}" title="${getWeekRangeLabel(week)}">W${String(week).padStart(2, "0")}</th>`;
      }).join("")}
    </tr>
  `;
}

function renderPlanner() {
  if (!plannerBody) return;

  extSequenceMap = new Map();

  const demandMap = buildDemandMap(demandsData);
  const allocationMaps = buildAllocationMaps(allocationsData, resourcesData);
  const historyMap = buildHistoryMap(demandHistoryData);
  const allocationHistoryMap = buildHistoryMap(allocationHistoryData);
  const rows = buildPlannerRows(projectsData, demandsData);

  renderPlannerHeader(rows, demandMap, allocationMaps);

  if (rows.length === 0) {
    plannerBody.innerHTML = `
      <tr>
        <td class="sticky-col">Nessun dato planner</td>
        <td class="sticky-col second">-</td>
        <td class="sticky-col third coverage good">0/0<br><small>0%</small></td>
        ${WEEKS.map(() => `<td class="cell-empty"></td>`).join("")}
      </tr>
    `;
    return;
  }

  plannerBody.innerHTML = rows.map((row, rowIndex) => {
    let totalRequired = 0;
    let totalAllocated = 0;
    let totalInternalAllocated = 0;
    let totalExternalAllocated = 0;

    const weekCells = WEEKS.map((week) => {
      const key = `${row.project_id}__${row.role}__${week}`;
      const required = demandMap.get(key) || 0;
      const internalAllocated = allocationMaps.internalMap.get(key) || 0;
      const externalAllocated = allocationMaps.externalMap.get(key) || 0;
      const allocated = internalAllocated + externalAllocated;
      const diff = required - allocated;
      const allocatedDisplay = formatAllocatedDisplay(internalAllocated, externalAllocated);

      const historyRows = historyMap.get(key) || [];
      const releasedRows = allocationHistoryMap.get(key) || [];
      const hasHistory = historyRows.length > 0;
      const hasAllocationHistory = releasedRows.length > 0;
      const assignments = getCellAssignments(row.project_id, row.role, week);

      totalRequired += required;
      totalAllocated += allocated;
      totalInternalAllocated += internalAllocated;
      totalExternalAllocated += externalAllocated;

      const calloutHtml = encodeURIComponent(
        buildCellCalloutHtml(row.project_name, row.role, week, assignments, releasedRows, historyRows)
      );

      const payload = encodeURIComponent(JSON.stringify({
        project_id: row.project_id,
        project_name: row.project_name,
        role: row.role,
        week,
        required,
        allocated,
        internalAllocated,
        externalAllocated,
        diff,
      }));

      return `
        <td
          class="${getCellClass(required, allocated, week, hasHistory, hasAllocationHistory)}"
          data-cell='${payload}'
          data-callout='${calloutHtml}'
          data-row-index="${rowIndex}"
          data-week="${week}"
        >
          R${formatNumber(required)}<br>
          A${allocatedDisplay}<br>
          D${formatNumber(diff)}${hasHistory || hasAllocationHistory ? `<span class="history-marker">↺</span>` : ""}
        </td>
      `;
    }).join("");

    const coverageClass = getCoverageClass(totalRequired, totalAllocated);
    const percent = totalRequired > 0 ? Math.round((totalAllocated / totalRequired) * 100) : 100;
    const totalAllocatedDisplay = formatAllocatedDisplay(totalInternalAllocated, totalExternalAllocated);

    return `
      <tr>
        <td class="sticky-col">${escapeHtml(row.project_name)}</td>
        <td class="sticky-col second">${escapeHtml(row.role)}</td>
        <td class="sticky-col third ${coverageClass}">
          ${totalAllocatedDisplay}/${formatNumber(totalRequired)}<br>
          <small>${percent}%</small>
        </td>
        ${weekCells}
      </tr>
    `;
  }).join("");

  plannerBody.querySelectorAll("td[data-cell]").forEach((cell) => {
    cell.addEventListener("click", () => {
      selectSingleCell(cell);
      moveFocusToGrid();
    });

    cell.addEventListener("mouseenter", (event) => {
      const html = decodeURIComponent(cell.dataset.callout || "");
      showCellCallout(event, html);
    });

    cell.addEventListener("mousemove", (event) => {
      const html = decodeURIComponent(cell.dataset.callout || "");
      showCellCallout(event, html);
    });

    cell.addEventListener("mouseleave", () => {
      hideCellCallout();
    });
  });

  const firstCurrentCell = plannerBody.querySelector("td[data-cell].current-col");
  if (firstCurrentCell) {
    selectSingleCell(firstCurrentCell);
  }
}

async function loadAndRenderPlanner() {
  const previousSummary = getSelectionSummary();

  [projectsData, demandsData, allocationsData, resourcesData, demandHistoryData, allocationHistoryData] = await Promise.all([
    fetchJson("/api/projects"),
    fetchJson("/api/demands"),
    fetchJson("/api/allocations"),
    fetchJson("/api/resources"),
    fetchJson("/api/demand-history"),
    fetchJson("/api/allocation-history"),
  ]);

  renderPlanner();
  applyWeekTooltips();
  scrollToCurrentWeek();

  if (previousSummary) {
    const candidates = Array.from(plannerBody.querySelectorAll(`td[data-week="${previousSummary.week_from}"][data-cell]`));
    const targetCell = candidates.find((cell) => {
      const data = JSON.parse(decodeURIComponent(cell.dataset.cell));
      return Number(data.project_id) === Number(previousSummary.project_id) && data.role === previousSummary.role;
    });

    if (targetCell) {
      selectSingleCell(targetCell);
    }
  }
}

document.addEventListener("DOMContentLoaded", async () => {
  try {
    await loadAndRenderPlanner();

    if (plannerGridWrap) {
      plannerGridWrap.addEventListener("keydown", handleGridKeydown);
    }

    [
      detailRequired,
      detailAllocated,
      detailDiff,
      detailProject,
      detailRole,
      detailWeekFrom,
      detailWeekTo,
      detailRange,
    ].forEach((field) => {
      if (field) field.addEventListener("keydown", handleDetailKeydown);
    });

    if (modeDemandBtn) modeDemandBtn.addEventListener("click", () => setMode("demand"));
    if (modeResourcesBtn) modeResourcesBtn.addEventListener("click", () => setMode("resources"));
    if (resourceSearchInput) resourceSearchInput.addEventListener("input", renderResourceLists);
    if (showInactiveToggle) showInactiveToggle.addEventListener("change", renderResourceLists);

    if (saveDemandBtn) {
      saveDemandBtn.addEventListener("click", async () => {
        try {
          await saveDemandRange();
        } catch (error) {
          console.error("Errore salvataggio fabbisogno:", error);
        }
      });
    }
  } catch (error) {
    console.error("Errore caricamento planner:", error);
    if (plannerBody) {
      plannerBody.innerHTML = `
        <tr>
          <td class="sticky-col">Errore</td>
          <td class="sticky-col second">Planner</td>
          <td class="sticky-col third coverage warn">KO</td>
          ${WEEKS.map(() => `<td class="cell-empty"></td>`).join("")}
        </tr>
      `;
    }
  }
});