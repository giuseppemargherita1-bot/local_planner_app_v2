const root = document.documentElement;
const verticalSplitter = document.getElementById("verticalSplitter");
const horizontalSplitter = document.getElementById("horizontalSplitter");
const sideInnerSplitter = document.getElementById("sideInnerSplitter");
const plannerGridWrap = document.getElementById("plannerGridWrap");
const plannerBody = document.getElementById("plannerBody");
const plannerBottomDetail = document.getElementById("plannerBottomDetail");

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

const CURRENT_PERIOD_KEY = 2617;
const PERIOD_YEAR_SHORT = 26;
const FIRST_WEEK = 1;
const LAST_WEEK = 52;

const PERIODS = Array.from({ length: LAST_WEEK - FIRST_WEEK + 1 }, (_, i) => {
  const week = FIRST_WEEK + i;
  return {
    yearShort: PERIOD_YEAR_SHORT,
    week,
    periodKey: PERIOD_YEAR_SHORT * 100 + week,
    label: `W${String(week).padStart(2, "0")}`,
  };
});

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

function periodKeyFromWeek(week, yearShort = PERIOD_YEAR_SHORT) {
  const numericWeek = Number(week || 0);
  if (numericWeek >= 1000) return numericWeek;
  return Number(yearShort) * 100 + numericWeek;
}

function weekFromPeriodKey(periodKey) {
  const value = Number(periodKey || 0);
  if (value >= 1000) return value % 100;
  return value;
}

function normalizePeriodKey(row) {
  const explicit = Number(row?.period_key || row?.periodKey || 0);
  if (explicit > 0) return explicit;
  return periodKeyFromWeek(row?.week || 0);
}

function getPeriodByKey(periodKey) {
  return PERIODS.find((period) => Number(period.periodKey) === Number(periodKey)) || null;
}

function getWeekStartDate(yearFull, week) {
  const jan4 = new Date(yearFull, 0, 4);
  const day = jan4.getDay() || 7;
  const mondayWeek1 = new Date(jan4);
  mondayWeek1.setDate(jan4.getDate() - day + 1);

  const result = new Date(mondayWeek1);
  result.setDate(mondayWeek1.getDate() + (week - 1) * 7);
  return result;
}

function fullYearFromShort(yearShort) {
  return 2000 + Number(yearShort || PERIOD_YEAR_SHORT);
}

function formatDate(date) {
  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  return `${dd}/${mm}`;
}

function getPeriodRangeLabel(periodKeyFrom, periodKeyTo = periodKeyFrom) {
  const startPeriod = getPeriodByKey(periodKeyFrom) || {
    yearShort: Math.floor(Number(periodKeyFrom) / 100),
    week: weekFromPeriodKey(periodKeyFrom),
  };

  const endPeriod = getPeriodByKey(periodKeyTo) || {
    yearShort: Math.floor(Number(periodKeyTo) / 100),
    week: weekFromPeriodKey(periodKeyTo),
  };

  const start = getWeekStartDate(fullYearFromShort(startPeriod.yearShort), startPeriod.week);
  const end = new Date(getWeekStartDate(fullYearFromShort(endPeriod.yearShort), endPeriod.week));
  end.setDate(end.getDate() + 6);

  return `${formatDate(start)} - ${formatDate(end)}`;
}

function getWeekRangeLabel(weekFrom, weekTo = weekFrom) {
  return getPeriodRangeLabel(periodKeyFromWeek(weekFrom), periodKeyFromWeek(weekTo));
}

function applyWeekTooltips() {
  document.querySelectorAll(".week-head").forEach((cell) => {
    const periodKey = Number(cell.dataset.periodKey || 0);
    if (periodKey > 0) {
      cell.title = getPeriodRangeLabel(periodKey);
      return;
    }

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

  const leftStickyWidth = 170 + 130 + 58;
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


function getProjectById(projectId) {
  return projectsData.find((project) => Number(project.id) === Number(projectId)) || null;
}

function isOverallProject(project) {
  return Number(project?.is_overall || 0) === 1;
}

function isWorkshopChildProject(project) {
  return Number(project?.workshop_rollup || 0) === 1 && Number(project?.parent_overall_id || 0) > 0;
}

function getProjectFirstUsefulPeriod(projectId) {
  let first = Infinity;

  demandsData.forEach((demand) => {
    if (Number(demand.project_id) !== Number(projectId)) return;
    if (Number(demand.quantity || 0) <= 0) return;
    first = Math.min(first, normalizePeriodKey(demand));
  });

  allocationsData.forEach((allocation) => {
    if (Number(allocation.project_id) !== Number(projectId)) return;
    first = Math.min(first, normalizePeriodKey(allocation));
  });

  return first === Infinity ? 999999 : first;
}

function getProjectLastUsefulPeriod(projectId) {
  let last = 0;

  demandsData.forEach((demand) => {
    if (Number(demand.project_id) !== Number(projectId)) return;
    if (Number(demand.quantity || 0) <= 0) return;
    last = Math.max(last, normalizePeriodKey(demand));
  });

  allocationsData.forEach((allocation) => {
    if (Number(allocation.project_id) !== Number(projectId)) return;
    last = Math.max(last, normalizePeriodKey(allocation));
  });

  return last;
}

function projectHasUsefulRows(projectId) {
  const project = getProjectById(projectId);
  if (!project) return false;
  if (isWorkshopChildProject(project)) return false;

  const lastUseful = getProjectLastUsefulPeriod(projectId);

  // Mostra commesse future o passate da poco.
  // Nasconde commesse finite troppo indietro tipo 382/390.
  return lastUseful >= CURRENT_PERIOD_KEY - 4;
}

function hasExtInAssignments(assignments) {
  return assignments.some((item) => {
    const resource = getAllocationResource(item);
    return isExternalResource(resource) || isExternalResource(item);
  });
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

function getExtSequenceKey(role, periodKey) {
  return `${normalizeExtRoleLabel(role)}__${Number(periodKey)}`;
}

function getExtDisplayName(allocation, rowRole, periodKey) {
  const allocationId = Number(allocation.id || allocation.history_id || allocation.allocation_id || 0);
  const role = normalizeExtRoleLabel(rowRole || allocation.role || allocation.resource_role || "EXT");
  const key = getExtSequenceKey(role, periodKey);

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

function formatAssignmentLine(item, rowRole, periodKey = null) {
  const resource = getAllocationResource(item);
  const isExt = isExternalResource(resource) || isExternalResource(item);

  if (isExt) {
    const extName = getExtDisplayName(item, rowRole, periodKey || normalizePeriodKey(item));
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
    const periodKey = normalizePeriodKey(demand);
    const key = `${Number(demand.project_id)}__${role}__${periodKey}`;
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

function parseItalianDateFromText(value) {
  const text = String(value || "").trim();
  const match = text.match(/([0-9]{1,2})\/([0-9]{1,2})\/([0-9]{4})/);

  if (!match) {
    return null;
  }

  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3]);

  if (!day || !month || !year) {
    return null;
  }

  return new Date(year, month - 1, day);
}


function getResourceEndDate(resource) {
  const noteText = String(resource?.availability_note || "");
  const endMatch = noteText.match(/END=([^|]+)/i);

  if (!endMatch) {
    return null;
  }

  const rawEnd = String(endMatch[1] || "").trim().toUpperCase();

  if (!rawEnd) {
    return null;
  }

  if (
    rawEnd.includes("INDET") ||
    rawEnd.includes("INDETERMINATO") ||
    rawEnd.includes("31/12/2099") ||
    rawEnd.includes("2099")
  ) {
    return null;
  }

  return parseItalianDateFromText(rawEnd);
}


function isResourceContractEndedForPeriod(resource, periodKey) {
  if (!resource) {
    return true;
  }

  if (isExternalResource(resource)) {
    return false;
  }

  if (Number(resource.is_active) !== 1) {
    return true;
  }

  const noteText = normalizeRole(resource.availability_note || "");

  if (
    noteText.includes("FUORI_CONTRATTO") ||
    noteText.includes("FUORI CONTRATTO") ||
    noteText.includes("CESSATO") ||
    noteText.includes("LICENZIATO")
  ) {
    return true;
  }

  const endDate = getResourceEndDate(resource);

  if (!endDate) {
    return false;
  }

  const week = weekFromPeriodKey(periodKey);
  const selectedDate = getWeekStartDate(2026, week);

  return endDate < selectedDate;
}


function isResourceContractEndedForSelection(resource) {
  const summary = getSelectionSummary();

  if (!summary) {
    return isResourceContractEndedForPeriod(resource, CURRENT_PERIOD_KEY);
  }

  return isResourceContractEndedForPeriod(resource, summary.period_from);
}

function isResourceAvailableForPeriod(resource, periodKey) {
  if (!resource) {
    return false;
  }

  // Gli EXT sono validi per il conteggio A della matrice,
  // ma vengono nascosti dalla lista "Disponibili" in renderResourceLists().
  if (isExternalResource(resource)) {
    return true;
  }

  if (isResourceContractEndedForPeriod(resource, periodKey)) {
    return false;
  }

  if (isExplicitlyUnavailable(resource)) {
    return false;
  }

  const week = weekFromPeriodKey(periodKey);
  const startWeek = parseHiringStartWeekFromNote(resource.availability_note);

  if (startWeek !== null && week < startWeek) {
    return false;
  }

  return true;
}





function isResourceAvailableForWeek(resource, week) {
  return isResourceAvailableForPeriod(resource, periodKeyFromWeek(week));
}

function buildAllocationMaps(allocations, resources) {
  const resourceById = new Map(resources.map((resource) => [Number(resource.id), resource]));

  const totalMap = new Map();
  const internalMap = new Map();
  const externalMap = new Map();

  for (const allocation of allocations) {
    const resource = resourceById.get(Number(allocation.resource_id));

    if (!resource) {
      continue;
    }

    const periodKey = normalizePeriodKey(allocation);

    if (!isResourceAvailableForPeriod(resource, periodKey)) {
      continue;
    }

    const role = normalizeRole(allocation.role || resource.role || "");
    const key = `${Number(allocation.project_id)}__${role}__${periodKey}`;
    const value = Number(allocation.load_percent || 0) / 100;

    if (value <= 0) {
      continue;
    }

    // REGOLA CORRETTA DAL VECCHIO PLANNER:
    // ogni allocazione attiva/valida deve contare in A,
    // anche se R=0 o se non esiste fabbisogno sulla cella.
    // In quel caso la cella diventa R0/A1/D-1 = surplus.
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
    const periodKey = normalizePeriodKey(row);
    const key = `${Number(row.project_id)}__${role}__${periodKey}`;
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(row);
  }

  return map;
}

function rowHasUsefulFutureActivity(projectId, role) {
  const normalizedRole = normalizeRole(role);

  return PERIODS.some((period) => {
    if (Number(period.periodKey) < CURRENT_PERIOD_KEY) {
      return false;
    }

    const demandFound = demandsData.some((demand) => {
      return (
        Number(demand.project_id) === Number(projectId) &&
        normalizeRole(demand.role || "") === normalizedRole &&
        normalizePeriodKey(demand) === Number(period.periodKey) &&
        Number(demand.quantity || 0) > 0
      );
    });

    if (demandFound) {
      return true;
    }

    const allocationFound = allocationsData.some((allocation) => {
      if (
        Number(allocation.project_id) !== Number(projectId) ||
        normalizeRole(allocation.role || "") !== normalizedRole ||
        normalizePeriodKey(allocation) !== Number(period.periodKey)
      ) {
        return false;
      }

      const resource = resourcesData.find((item) => Number(item.id) === Number(allocation.resource_id));

      if (!resource) {
        return false;
      }

      return isResourceAvailableForPeriod(resource, period.periodKey);
    });

    return allocationFound;
  });
}


function rowHasRecentPastActivity(projectId, role) {
  const normalizedRole = normalizeRole(role);
  const recentPastStart = CURRENT_PERIOD_KEY - 2;

  return PERIODS.some((period) => {
    if (
      Number(period.periodKey) < recentPastStart ||
      Number(period.periodKey) >= CURRENT_PERIOD_KEY
    ) {
      return false;
    }

    const demandFound = demandsData.some((demand) => {
      return (
        Number(demand.project_id) === Number(projectId) &&
        normalizeRole(demand.role || "") === normalizedRole &&
        normalizePeriodKey(demand) === Number(period.periodKey) &&
        Number(demand.quantity || 0) > 0
      );
    });

    if (demandFound) {
      return true;
    }

    const allocationFound = allocationsData.some((allocation) => {
      if (
        Number(allocation.project_id) !== Number(projectId) ||
        normalizeRole(allocation.role || "") !== normalizedRole ||
        normalizePeriodKey(allocation) !== Number(period.periodKey)
      ) {
        return false;
      }

      const resource = resourcesData.find((item) => Number(item.id) === Number(allocation.resource_id));

      if (!resource) {
        return false;
      }

      return isResourceAvailableForPeriod(resource, period.periodKey);
    });

    return allocationFound;
  });
}


function shouldShowPlannerRow(projectId, role) {
  const showZero = document.querySelector(".inline-check input")?.checked;

  if (showZero) {
    return true;
  }

  const project = projectsData.find((item) => Number(item.id) === Number(projectId));
  const projectName = normalizeRole(project?.name || "");

  if (projectName.includes("OVERALL OFFICINA")) {
    return true;
  }

  // Vista operativa standard:
  // mostra solo righe con attivitÃ  corrente/futura oppure passato recente.
  // Nasconde commesse finite troppo tempo fa tipo 382/390.
  return (
    rowHasUsefulFutureActivity(projectId, role) ||
    rowHasRecentPastActivity(projectId, role)
  );
}

function getFirstUsefulPeriodForRow(projectId, role) {
  const normalizedRole = normalizeRole(role);

  const usefulPeriods = [];

  demandsData.forEach((demand) => {
    if (
      Number(demand.project_id) === Number(projectId) &&
      normalizeRole(demand.role || "") === normalizedRole &&
      Number(demand.quantity || 0) > 0
    ) {
      usefulPeriods.push(normalizePeriodKey(demand));
    }
  });

  allocationsData.forEach((allocation) => {
    if (
      Number(allocation.project_id) === Number(projectId) &&
      normalizeRole(allocation.role || "") === normalizedRole
    ) {
      usefulPeriods.push(normalizePeriodKey(allocation));
    }
  });

  const futureOrCurrent = usefulPeriods
    .filter((periodKey) => Number(periodKey) >= CURRENT_PERIOD_KEY)
    .sort((a, b) => a - b);

  if (futureOrCurrent.length) {
    return futureOrCurrent[0];
  }

  const past = usefulPeriods
    .filter((periodKey) => Number(periodKey) < CURRENT_PERIOD_KEY)
    .sort((a, b) => b - a);

  if (past.length) {
    return 9000 + (CURRENT_PERIOD_KEY - past[0]);
  }

  return 9999;
}

function buildPlannerRows(projects, demands) {
  const projectsById = new Map(projects.map((project) => [Number(project.id), project]));
  const rowsMap = new Map();

  for (const demand of demands) {
    const project = projectsById.get(Number(demand.project_id));

    if (!project) {
      continue;
    }

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

  // Manteniamo la correzione APPELLA:
  // aggiungiamo righe anche da allocazioni R0/A1,
  // ma solo se la riga Ã¨ utile nella vista operativa.
  for (const allocation of allocationsData) {
    const project = projectsById.get(Number(allocation.project_id));

    if (!project) {
      continue;
    }

    const resource = resourcesData.find((item) => Number(item.id) === Number(allocation.resource_id));

    if (!resource) {
      continue;
    }

    const periodKey = normalizePeriodKey(allocation);

    if (!isResourceAvailableForPeriod(resource, periodKey)) {
      continue;
    }

    const role = normalizeRole(allocation.role || resource.role || "");

    if (!role) {
      continue;
    }

    const rowKey = `${Number(allocation.project_id)}__${role}`;

    if (!rowsMap.has(rowKey)) {
      rowsMap.set(rowKey, {
        row_key: rowKey,
        project_id: Number(allocation.project_id),
        project_name: project.name,
        role,
      });
    }
  }

  let rows = Array.from(rowsMap.values());

  rows = rows.filter((row) => {
    return shouldShowPlannerRow(row.project_id, row.role);
  });

  rows.sort((a, b) => {
    const aFirst = getFirstUsefulPeriodForRow(a.project_id, a.role);
    const bFirst = getFirstUsefulPeriodForRow(b.project_id, b.role);

    if (aFirst !== bFirst) {
      return aFirst - bFirst;
    }

    if (a.project_name !== b.project_name) {
      return a.project_name.localeCompare(b.project_name);
    }

    return a.role.localeCompare(b.role);
  });

  rowMetaMap = new Map(rows.map((row, index) => [index, row]));

  return rows;
}





function getCoverageClass(required, allocated) {
  if (required === 0 && allocated === 0) return "coverage good";
  if (required === 0 && allocated > 0) return "coverage warn";
  if (allocated < required) return "coverage warn";
  if (allocated > required) return "coverage warn";
  return "coverage good";
}

function getCellClass(required, allocated, periodKey, hasHistory, hasAllocationHistory, hasExt) {
  let cls = "cell-empty planner-cell-clickable";

  if (required === 0 && allocated === 0) {
    cls = "cell-empty planner-cell-clickable";
  } else if (allocated < required) {
    cls = "cell-demand planner-cell-clickable";
  } else if (allocated >= required) {
    cls = "cell-ok planner-cell-clickable";
  }

  if (allocated > required) {
    cls += " cell-surplus-border";
  }

  if (periodKey === CURRENT_PERIOD_KEY) cls += " current-col";
  if (hasHistory || hasAllocationHistory) cls += " cell-history";
  if (hasExt) cls += " cell-has-ext";

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
  const periodKeys = selectedCells.map((cell) => Number(cell.dataset.periodKey)).sort((a, b) => a - b);

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

  const periodFrom = periodKeys[0];
  const periodTo = periodKeys[periodKeys.length - 1];

  return {
    project_id: Number(decoded.project_id),
    project_name: decoded.project_name,
    role: decoded.role,
    row_key: `${Number(decoded.project_id)}__${decoded.role}`,
    period_from: periodFrom,
    period_to: periodTo,
    week_from: weekFromPeriodKey(periodFrom),
    week_to: weekFromPeriodKey(periodTo),
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
    `${summary.period_to !== summary.period_from ? ` - W${String(summary.week_to).padStart(2, "0")}` : ""}`;

  detailProject.value = summary.project_name;
  detailRole.value = summary.role;
  detailWeekFrom.value = summary.week_from;
  detailWeekTo.value = summary.week_to;
  detailRange.value = getPeriodRangeLabel(summary.period_from, summary.period_to);
  detailRequired.value = formatNumber(summary.required);
  detailAllocated.value = formatAllocatedDisplay(summary.internalAllocated, summary.externalAllocated);
  detailDiff.value = formatNumber(summary.required - summary.allocated);

  renderResourceLists();
  renderBottomDetailFromSelection(summary);
}
function getSingleSelectedCellData() {
  if (!selectedCells.length) return null;

  try {
    return JSON.parse(decodeURIComponent(selectedCells[0].dataset.cell));
  } catch (error) {
    console.error("Errore lettura dati cella selezionata", error);
    return null;
  }
}

function resourceDisplayName(item) {
  const rawName =
    item.resource_name ||
    item.resourceName ||
    item.name ||
    "";

  return shortResourceName(rawName);
}

function resourceStatusLabel(item) {
  const status = normalizeRole(item.status || item.reason || "");

  if (item.external) return "EXT";
  if (status.includes("FUORI") || status.includes("CESS")) return "FUORI_CONTRATTO";
  if (status.includes("INDISP")) return "INDISP";
  if (status.includes("STORICO")) return "STORICO";

  return status || "ACTIVE";
}

function renderResourceDetailLines(resources, rowRole, periodKey, historical = false) {
  if (!Array.isArray(resources) || !resources.length) {
    return `<div class="bottom-muted">Nessuna risorsa ${historical ? "storica" : "attiva"}.</div>`;
  }

  return resources.map((item) => {
    const isExt = Boolean(item.external) || isExternalResource(item);
    const name = isExt
      ? getExtDisplayName(item, rowRole, periodKey)
      : resourceDisplayName(item);

    const realRole = item.resource_role || item.resourceRole || item.role || "";
    const mansione = isExt ? "EXT" : mansioneLabel(rowRole, realRole);
    const status = resourceStatusLabel(item);
    const load = historical ? "0:1" : (item.display_weight || formatLoad(item.load_percent || item.loadPercent || 100));

    const classes = [
      "bottom-resource-row",
      historical ? "bottom-resource-history" : "",
      isExt ? "bottom-resource-ext" : "",
    ].join(" ");

    return `
      <div class="${classes}">
        <div class="bottom-resource-name">${escapeHtml(name)}</div>
        <div class="bottom-resource-role">${escapeHtml(mansione)}</div>
        <div class="bottom-resource-load">${escapeHtml(load)}</div>
        <div class="bottom-resource-status">${escapeHtml(status)}</div>
      </div>
    `;
  }).join("");
}

async function renderBottomDetailFromSelection(summary) {
  if (!plannerBottomDetail || !summary) return;

  const cellData = getSingleSelectedCellData();

  if (!cellData) {
    plannerBottomDetail.innerHTML = `
      <div class="bottom-muted">Seleziona una cella per vedere dettaglio, storico e note operative.</div>
    `;
    return;
  }

  const isOverall = String(summary.project_name || "").toUpperCase().includes("OVERALL OFFICINA");

  plannerBottomDetail.innerHTML = `
    <div class="bottom-detail-header">
      <div>
        <strong>${escapeHtml(summary.project_name)}</strong>
        <span> / ${escapeHtml(summary.role)}</span>
        <span> / W${String(summary.week_from).padStart(2, "0")}</span>
      </div>
      <div>
        R${formatNumber(summary.required)}
        &nbsp; A${formatNumber(summary.allocated)}
        &nbsp; D${formatNumber(summary.required - summary.allocated)}
      </div>
    </div>
    <div class="bottom-muted">Caricamento dettaglio...</div>
  `;

  if (isOverall) {
    await renderOverallDetailInBottom(summary, cellData);
    return;
  }

  renderNormalCellDetailInBottom(summary, cellData);
}

function renderNormalCellDetailInBottom(summary, cellData) {
  const activeResources = Array.isArray(cellData.active_resources) ? cellData.active_resources : [];
  const historicalResources = Array.isArray(cellData.historical_resources) ? cellData.historical_resources : [];
  const badges = Array.isArray(cellData.badges) ? cellData.badges : [];

  const badgesHtml = badges.length
    ? `<div class="bottom-badges">${badges.map((badge) => `<span>${escapeHtml(badge)}</span>`).join("")}</div>`
    : "";

  plannerBottomDetail.innerHTML = `
    <div class="bottom-detail-header">
      <div>
        <strong>${escapeHtml(summary.project_name)}</strong>
        <span> / ${escapeHtml(summary.role)}</span>
        <span> / W${String(summary.week_from).padStart(2, "0")}</span>
      </div>
      <div>
        R${formatNumber(summary.required)}
        &nbsp; A${formatNumber(summary.allocated)}
        &nbsp; D${formatNumber(summary.required - summary.allocated)}
      </div>
    </div>

    ${badgesHtml}

    <div class="bottom-detail-grid">
      <div class="bottom-detail-block">
        <div class="bottom-detail-title">Risorse attive conteggiate</div>
        ${renderResourceDetailLines(activeResources, summary.role, summary.period_from, false)}
      </div>

      <div class="bottom-detail-block">
        <div class="bottom-detail-title">Storico / fuori contratto / non conteggiati</div>
        ${renderResourceDetailLines(historicalResources, summary.role, summary.period_from, true)}
      </div>
    </div>
  `;
}

async function renderOverallDetailInBottom(summary, cellData) {
  try {
    const params = new URLSearchParams({
      project_id: String(summary.project_id),
      role: summary.role,
      week: String(summary.week_from),
    });

    const data = await fetchJson(`/api/workshop-breakdown?${params.toString()}`);
    const sources = Array.isArray(data.sources) ? data.sources : [];

    const sourceRows = sources.length
      ? sources.map((source) => {
          return `
            <div class="bottom-overall-row">
              <div class="bottom-overall-project">${escapeHtml(source.project_name || "")}</div>
              <div class="bottom-overall-role">${escapeHtml(source.role || summary.role)}</div>
              <div class="bottom-overall-required">R${formatNumber(source.required || 0)}</div>
            </div>
          `;
        }).join("")
      : `<div class="bottom-muted">Nessuna sottocommessa trovata per questa cella.</div>`;

    const activeResources = Array.isArray(cellData.active_resources) ? cellData.active_resources : [];
    const historicalResources = Array.isArray(cellData.historical_resources) ? cellData.historical_resources : [];

    plannerBottomDetail.innerHTML = `
      <div class="bottom-detail-header">
        <div>
          <strong>${escapeHtml(summary.project_name)}</strong>
          <span> / ${escapeHtml(summary.role)}</span>
          <span> / W${String(summary.week_from).padStart(2, "0")}</span>
        </div>
        <div>
          R${formatNumber(summary.required)}
          &nbsp; A${formatNumber(summary.allocated)}
          &nbsp; D${formatNumber(summary.required - summary.allocated)}
        </div>
      </div>

      <div class="bottom-detail-grid bottom-detail-grid-wide">
        <div class="bottom-detail-block">
          <div class="bottom-detail-title">Dettaglio sottocommesse officina</div>
          ${sourceRows}
          <div class="bottom-overall-total">
            <div>Totale OVERALL</div>
            <div>R${formatNumber(data.total_required || 0)}</div>
          </div>
        </div>

        <div class="bottom-detail-block">
          <div class="bottom-detail-title">Allocazioni su OVERALL OFFICINA</div>
          ${renderResourceDetailLines(activeResources, summary.role, summary.period_from, false)}

          <div class="bottom-detail-title bottom-detail-title-secondary">
            Storico / fuori contratto
          </div>
          ${renderResourceDetailLines(historicalResources, summary.role, summary.period_from, true)}
        </div>
      </div>
    `;
  } catch (error) {
    console.error(error);
    plannerBottomDetail.innerHTML = `
      <div class="bottom-detail-header">
        <div>
          <strong>${escapeHtml(summary.project_name)}</strong>
          <span> / ${escapeHtml(summary.role)}</span>
          <span> / W${String(summary.week_from).padStart(2, "0")}</span>
        </div>
      </div>
      <div class="bottom-error">Errore caricamento dettaglio OVERALL OFFICINA.</div>
    `;
  }
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
  const periodKeys = selectedCells.map((cell) => Number(cell.dataset.periodKey)).sort((a, b) => a - b);
  const minPeriod = periodKeys[0];
  const maxPeriod = periodKeys[periodKeys.length - 1];

  const currentIndex = direction === "right"
    ? PERIODS.findIndex((p) => p.periodKey === maxPeriod)
    : PERIODS.findIndex((p) => p.periodKey === minPeriod);

  if (currentIndex < 0) return;

  const targetIndex = direction === "right" ? currentIndex + 1 : currentIndex - 1;
  if (targetIndex < 0 || targetIndex >= PERIODS.length) return;

  const targetPeriodKey = PERIODS[targetIndex].periodKey;

  const targetCell = plannerBody.querySelector(
    `td[data-row-index="${rowIndex}"][data-period-key="${targetPeriodKey}"]`
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
  const currentPeriodKey = Number(anchor.dataset.periodKey);
  const currentIndex = PERIODS.findIndex((p) => p.periodKey === currentPeriodKey);
  if (currentIndex < 0) return;

  const targetRow = currentRow + deltaRow;
  const targetIndex = currentIndex + deltaCol;
  if (targetIndex < 0 || targetIndex >= PERIODS.length) return;

  const targetPeriodKey = PERIODS[targetIndex].periodKey;

  const target = plannerBody.querySelector(
    `td[data-row-index="${targetRow}"][data-period-key="${targetPeriodKey}"]`
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

function getSelectedPeriodKeysSet() {
  const summary = getSelectionSummary();
  const selected = new Set();

  if (!summary) return selected;

  const startIndex = PERIODS.findIndex((p) => p.periodKey === summary.period_from);
  const endIndex = PERIODS.findIndex((p) => p.periodKey === summary.period_to);

  if (startIndex < 0 || endIndex < 0) return selected;

  for (let i = startIndex; i <= endIndex; i += 1) {
    selected.add(PERIODS[i].periodKey);
  }

  return selected;
}

function getSelectedWeeksSet() {
  const selectedPeriodKeys = getSelectedPeriodKeysSet();
  const weeks = new Set();

  selectedPeriodKeys.forEach((periodKey) => {
    weeks.add(weekFromPeriodKey(periodKey));
  });

  return weeks;
}

function getResourceAllocationsForSelectedPeriods(resourceId) {
  const selectedPeriodKeys = getSelectedPeriodKeysSet();

  return allocationsData.filter((allocation) => {
    return (
      Number(allocation.resource_id) === Number(resourceId) &&
      selectedPeriodKeys.has(normalizePeriodKey(allocation))
    );
  });
}

function getResourceAllocationsForSelectedWeeks(resourceId) {
  return getResourceAllocationsForSelectedPeriods(resourceId);
}

function getResourceAllocationsForPeriod(resourceId, periodKey) {
  return allocationsData.filter((allocation) => {
    return Number(allocation.resource_id) === Number(resourceId) && normalizePeriodKey(allocation) === Number(periodKey);
  });
}

function getResourceAllocationsForWeek(resourceId, week) {
  return getResourceAllocationsForPeriod(resourceId, periodKeyFromWeek(week));
}

function getAssignedResourcesForSelection() {
  const summary = getSelectionSummary();
  if (!summary) return [];

  const selectedPeriodKeys = getSelectedPeriodKeysSet();

  return resourcesData.filter((resource) => {
    if (!isResourceAvailableForPeriod(resource, summary.period_from)) return false;

    return allocationsData.some((allocation) => {
      const allocationRole = normalizeRole(allocation.role || "");

      return (
        Number(allocation.resource_id) === Number(resource.id) &&
        Number(allocation.project_id) === Number(summary.project_id) &&
        allocationRole === summary.role &&
        selectedPeriodKeys.has(normalizePeriodKey(allocation))
      );
    });
  });
}

function getResourceStatus(resource) {
  const summary = getSelectionSummary();

  if (!resource) {
    return "inactive";
  }

  if (isExternalResource(resource)) {
    return "external";
  }

  if (isResourceContractEndedForSelection(resource)) {
    return "inactive";
  }

  if (isExplicitlyUnavailable(resource)) {
    return "unavailable";
  }

  if (!summary) {
    return "free";
  }

  const startWeek = parseHiringStartWeekFromNote(resource.availability_note);

  if (startWeek !== null && summary.week_from < startWeek) {
    return "future";
  }

  const assigned = getAssignedResourcesForSelection().some((assignedResource) => {
    return Number(assignedResource.id) === Number(resource.id);
  });

  if (assigned) {
    return "allocated";
  }

  const demandRowSet = buildDemandRowSet(demandsData);

  const allocations = getResourceAllocationsForSelectedPeriods(resource.id).filter((allocation) => {
    const allocationRole = normalizeRole(allocation.role || "");
    const rowKey = `${Number(allocation.project_id)}__${allocationRole}`;
    return demandRowSet.has(rowKey);
  });

  const maxPerPeriod = new Map();

  allocations.forEach((allocation) => {
    const periodKey = normalizePeriodKey(allocation);
    maxPerPeriod.set(periodKey, (maxPerPeriod.get(periodKey) || 0) + 1);
  });

  const values = Array.from(maxPerPeriod.values());

  if (values.some((count) => count >= 2)) {
    return "saturated";
  }

  if (values.some((count) => count === 1)) {
    return "partial";
  }

  const historical = getResourceAllocationsForSelectedPeriods(resource.id).filter((allocation) => {
    const allocationRole = normalizeRole(allocation.role || "");
    const rowKey = `${Number(allocation.project_id)}__${allocationRole}`;
    return !demandRowSet.has(rowKey);
  });

  if (historical.length > 0) {
    return "history";
  }

  return "free";
}





function allocationShortLabel(allocation) {
  const project = allocation.project_name || `Progetto ${allocation.project_id}`;
  const role = allocation.role || "-";
  const periodKey = normalizePeriodKey(allocation);
  const week = weekFromPeriodKey(periodKey);
  const load = Number(allocation.load_percent || 0);
  return `${project} ${role} W${String(week).padStart(2, "0")} ${load}%`;
}

function getAllocationLocationText(resource, onlyHistorical = false) {
  const demandRowSet = buildDemandRowSet(demandsData);

  const allocations = getResourceAllocationsForSelectedPeriods(resource.id).filter((allocation) => {
    const allocationRole = normalizeRole(allocation.role || "");
    const rowKey = `${Number(allocation.project_id)}__${allocationRole}`;
    const isHistorical = !demandRowSet.has(rowKey);
    return onlyHistorical ? isHistorical : !isHistorical;
  });

  if (!allocations.length) return "";
  return allocations.map(allocationShortLabel).join(" / ");
}

function getAssignedResourceInfo(resource) {
  const summary = getSelectionSummary();

  if (!summary || !resource) {
    return "ASSEGNATA";
  }

  const selectedPeriodKeys = getSelectedPeriodKeysSet();

  const rows = allocationsData.filter((allocation) => {
    return (
      Number(allocation.resource_id) === Number(resource.id) &&
      Number(allocation.project_id) === Number(summary.project_id) &&
      normalizeRole(allocation.role || "") === summary.role &&
      selectedPeriodKeys.has(normalizePeriodKey(allocation))
    );
  });

  if (!rows.length) {
    return "ASSEGNATA";
  }

  const loads = [...new Set(rows.map((allocation) => formatLoad(allocation.load_percent || allocation.loadPercent || 100)))];
  const weeks = rows
    .map((allocation) => weekFromPeriodKey(normalizePeriodKey(allocation)))
    .sort((a, b) => a - b);

  const firstWeek = weeks[0];
  const lastWeek = weeks[weeks.length - 1];
  const weekText = firstWeek === lastWeek
    ? `W${String(firstWeek).padStart(2, "0")}`
    : `W${String(firstWeek).padStart(2, "0")}-W${String(lastWeek).padStart(2, "0")}`;

  return `${summary.role} | ${loads.join("+")} | ${weekText}`;
}


function getResourceBusyInfo(resource) {
  const summary = getSelectionSummary();

  if (!summary || !resource) {
    return "";
  }

  const selectedPeriodKeys = getSelectedPeriodKeysSet();

  const rows = allocationsData.filter((allocation) => {
    return (
      Number(allocation.resource_id) === Number(resource.id) &&
      selectedPeriodKeys.has(normalizePeriodKey(allocation))
    );
  });

  if (!rows.length) {
    return "";
  }

  const byProjectRole = new Map();

  rows.forEach((allocation) => {
    const projectName = allocation.project_name || `Progetto ${allocation.project_id}`;
    const role = allocation.role || "";
    const key = `${projectName}__${role}`;

    if (!byProjectRole.has(key)) {
      byProjectRole.set(key, {
        projectName,
        role,
        weeks: [],
        loads: [],
      });
    }

    const bucket = byProjectRole.get(key);
    bucket.weeks.push(weekFromPeriodKey(normalizePeriodKey(allocation)));
    bucket.loads.push(formatLoad(allocation.load_percent || allocation.loadPercent || 100));
  });

  return Array.from(byProjectRole.values()).map((bucket) => {
    const weeks = bucket.weeks.sort((a, b) => a - b);
    const firstWeek = weeks[0];
    const lastWeek = weeks[weeks.length - 1];
    const weekText = firstWeek === lastWeek
      ? `W${String(firstWeek).padStart(2, "0")}`
      : `W${String(firstWeek).padStart(2, "0")}-W${String(lastWeek).padStart(2, "0")}`;

    const loads = [...new Set(bucket.loads)];

    return `${bucket.projectName} | ${bucket.role} | ${loads.join("+")} | ${weekText}`;
  }).join(" / ");
}

function getResourceShortInfo(resource, status) {
  if (status === "external") {
    return "EXT";
  }

  if (status === "inactive") {
    const endDate = getResourceEndDate(resource);

    if (endDate) {
      const day = String(endDate.getDate()).padStart(2, "0");
      const month = String(endDate.getMonth() + 1).padStart(2, "0");
      const year = endDate.getFullYear();
      return `CESSATO DAL ${day}/${month}/${year}`;
    }

    return "CESSATO / FUORI CONTRATTO";
  }

  if (status === "unavailable") {
    return "INDISPONIBILE";
  }

  if (status === "future") {
    return "NON ANCORA DISPONIBILE";
  }

  if (status === "allocated") {
    return getAssignedResourceInfo(resource);
  }

  if (status === "saturated") {
    return getResourceBusyInfo(resource) || "SATURA";
  }

  if (status === "partial") {
    return getResourceBusyInfo(resource) || "PARZIALE";
  }

  if (status === "history") {
    return "SOLO STORICO";
  }

  return "DISPONIBILE";
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
  if (!assignedResourceList || !availableResourceList) {
    return;
  }

  const summary = getSelectionSummary();

  assignedResourceList.innerHTML = "";
  availableResourceList.innerHTML = "";

  if (!summary) {
    assignedResourceList.innerHTML = `<div class="resource-empty">Seleziona una cella.</div>`;
    availableResourceList.innerHTML = `<div class="resource-empty">Seleziona una cella.</div>`;
    return;
  }

  const searchText = normalizeRole(resourceSearchInput ? resourceSearchInput.value : "");
  const showInactive = Boolean(showInactiveToggle && showInactiveToggle.checked);
  const selectedPeriodKeys = getSelectedPeriodKeysSet();

  // Come nel vecchio planner:
  // le allocazioni EXT sono vere allocazioni e devono essere mostrate tra le assegnate,
  // ma non devono comparire tra le risorse disponibili.
  const assignedAllocations = allocationsData.filter((allocation) => {
    const allocationRole = normalizeRole(allocation.role || "");

    return (
      Number(allocation.project_id) === Number(summary.project_id) &&
      allocationRole === summary.role &&
      selectedPeriodKeys.has(normalizePeriodKey(allocation))
    );
  });

  if (!assignedAllocations.length) {
    assignedResourceList.innerHTML = `<div class="resource-empty">Nessuna risorsa assegnata</div>`;
  } else {
    assignedResourceList.innerHTML = assignedAllocations.map((allocation) => {
      const resource = getAllocationResource(allocation);
      const isExt = isExternalResource(resource) || isExternalResource(allocation);
      const status = isExt ? "external" : getResourceStatus(resource);
      const name = isExt
        ? getExtDisplayName(allocation, summary.role, normalizePeriodKey(allocation))
        : (resource.name || allocation.resource_name || "");

      const load = formatLoad(allocation.load_percent || allocation.loadPercent || 100);
      const week = weekFromPeriodKey(normalizePeriodKey(allocation));
      const role = resource.role || allocation.resource_role || allocation.role || "";

      return `
        <button
          class="resource-item resource-status-${escapeHtml(status)}"
          type="button"
          data-resource-id="${Number(resource.id || allocation.resource_id || 0)}"
        >
          <span class="resource-main">${escapeHtml(name)}</span>
          <span class="resource-sub">
            ${escapeHtml(role)} | ${escapeHtml(load)} | W${String(week).padStart(2, "0")}
          </span>
        </button>
      `;
    }).join("");
  }

  const assignedIds = new Set(
    assignedAllocations.map((allocation) => Number(allocation.resource_id))
  );

  const availableResources = resourcesData.filter((resource) => {
    if (!resource) {
      return false;
    }

    // EXT nascosti dalla lista disponibili.
    // Verranno aggiunti in seguito solo dal pulsante "Usa EXT".
    if (isExternalResource(resource)) {
      return false;
    }

    const status = getResourceStatus(resource);

    if (!showInactive && (status === "inactive" || status === "unavailable")) {
      return false;
    }

    if (assignedIds.has(Number(resource.id))) {
      return false;
    }

    if (searchText) {
      const haystack = normalizeRole(
        `${resource.name || ""} ${resource.role || ""} ${resource.availability_note || ""}`
      );

      if (!haystack.includes(searchText)) {
        return false;
      }
    }

    return true;
  });

  if (!availableResources.length) {
    availableResourceList.innerHTML = `<div class="resource-empty">Nessuna risorsa disponibile</div>`;
    return;
  }

  availableResourceList.innerHTML = availableResources.map((resource) => {
    const status = getResourceStatus(resource);
    const info = getResourceShortInfo(resource, status);
    const disabled = status === "inactive" || status === "unavailable";

    return `
      <button
        class="resource-item resource-status-${escapeHtml(status)}"
        type="button"
        data-resource-id="${Number(resource.id)}"
        ${disabled ? "disabled" : ""}
      >
        <span class="resource-main">${escapeHtml(resource.name || "")}</span>
        <span class="resource-sub">${escapeHtml(resource.role || "")} | ${escapeHtml(info)}</span>
      </button>
    `;
  }).join("");

  availableResourceList.querySelectorAll(".resource-item:not([disabled])").forEach((button) => {
    button.addEventListener("dblclick", async () => {
      const resourceId = Number(button.dataset.resourceId || 0);
      if (!resourceId || !summary) return;

      try {
        await fetchJson("/api/allocations/assign-range", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            resource_id: resourceId,
            project_id: summary.project_id,
            role: summary.role,
            week_from: summary.week_from,
            week_to: summary.week_to,
            hours: 40,
            load_percent: 100,
            note: "",
          }),
        });

        await loadAll();
      } catch (error) {
        console.error(error);
        alert("Errore assegnazione risorsa.");
      }
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

  if (summary.period_from !== summary.period_to) {
    const demandRowSet = buildDemandRowSet(demandsData);
    const conflicts = getResourceAllocationsForSelectedPeriods(resourceId).filter((allocation) => {
      const allocationRole = normalizeRole(allocation.role || "");
      const rowKey = `${Number(allocation.project_id)}__${allocationRole}`;
      return demandRowSet.has(rowKey);
    });

    if (conflicts.length > 0) {
      showConflictDialog(
        "Conflitto su piÃ¹ settimane",
        "La risorsa Ã¨ giÃ  allocata in almeno una delle settimane selezionate.<br>Gestisci una settimana alla volta.",
        [{ label: "OK", primary: true, handler: async () => {} }],
      );
      return;
    }

    await assignResourceToSelection(resourceId);
    return;
  }

  const periodKey = summary.period_from;
  const demandRowSet = buildDemandRowSet(demandsData);
  const existing = getResourceAllocationsForPeriod(resourceId, periodKey).filter((allocation) => {
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
      "Risorsa giÃ  allocata",
      `
        <p><strong>${escapeHtml(resource.name)}</strong> Ã¨ giÃ  allocato su:</p>
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
    <p><strong>${escapeHtml(resource.name)}</strong> Ã¨ giÃ  allocato su 2 commesse:</p>
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
    week_from: Number(summary.week_from),
    week_to: Number(summary.week_to),
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

function getCellAssignments(projectId, role, periodKey) {
  const normalizedRole = normalizeRole(role);

  return allocationsData.filter((item) => {
    return (
      Number(item.project_id) === Number(projectId) &&
      normalizeRole(item.role || "") === normalizedRole &&
      normalizePeriodKey(item) === Number(periodKey)
    );
  });
}

function buildCellCalloutHtml(projectName, role, periodKey, assignments, released, demandHistoryRows) {
  const week = weekFromPeriodKey(periodKey);
  const parts = [];

  parts.push(`<div class="cell-callout-title">${escapeHtml(projectName)} | ${escapeHtml(role)} | W${String(week).padStart(2, "0")}</div>`);

  if (assignments.length) {
    parts.push(`<div class="cell-callout-section">Assegnati</div>`);
    assignments.forEach((item) => {
      parts.push(`<div class="cell-callout-line">${escapeHtml(formatAssignmentLine(item, role, periodKey))}</div>`);
    });
  }

  if (released.length) {
    parts.push(`<div class="cell-callout-section">Non piÃ¹ conteggiati</div>`);
    released.forEach((item) => {
      const line = `${formatAssignmentLine(item, role, periodKey)} | ${item.reason || "storico"}`;
      parts.push(`<div class="cell-callout-line cell-callout-released">${escapeHtml(line)}</div>`);
    });
  }

  if (demandHistoryRows.length) {
    const latest = demandHistoryRows[0];
    parts.push(`<div class="cell-callout-section">Storico fabbisogno</div>`);
    parts.push(
      `<div class="cell-callout-line">${escapeHtml(latest.old_quantity)} â†’ ${escapeHtml(latest.new_quantity)} | ${escapeHtml(latest.created_at || "")}</div>`
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

  PERIODS.forEach((period) => {
    totals.set(period.periodKey, {
      periodKey: period.periodKey,
      required: 0,
      internalAllocated: 0,
      externalAllocated: 0,
      allocated: 0,
      diff: 0,
    });
  });

  rows.forEach((row) => {
    PERIODS.forEach((period) => {
      const key = `${row.project_id}__${row.role}__${period.periodKey}`;
      const required = demandMap.get(key) || 0;
      const internalAllocated = allocationMaps.internalMap.get(key) || 0;
      const externalAllocated = allocationMaps.externalMap.get(key) || 0;
      const allocated = internalAllocated + externalAllocated;

      const total = totals.get(period.periodKey);
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
      ${PERIODS.map((period) => {
        const currentClass = period.periodKey === CURRENT_PERIOD_KEY ? " current-week" : "";

        if (period.periodKey < CURRENT_PERIOD_KEY) {
          return `<th class="week-head week-subtotal-head week-subtotal-past${currentClass}" data-period-key="${period.periodKey}" title="${getPeriodRangeLabel(period.periodKey)}"></th>`;
        }

        const total = totals.get(period.periodKey);
        const allocatedDisplay = formatAllocatedDisplay(total.internalAllocated, total.externalAllocated);

        return `
          <th class="week-head week-subtotal-head${currentClass}" data-period-key="${period.periodKey}" title="${getPeriodRangeLabel(period.periodKey)}">
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
      ${PERIODS.map((period) => {
        const currentClass = period.periodKey === CURRENT_PERIOD_KEY ? " current-week" : "";
        return `<th class="week-head second-line${currentClass}" data-period-key="${period.periodKey}" title="${getPeriodRangeLabel(period.periodKey)}">${period.label}</th>`;
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
        ${PERIODS.map(() => `<td class="cell-empty"></td>`).join("")}
      </tr>
    `;
    return;
  }

  plannerBody.innerHTML = rows.map((row, rowIndex) => {
    let totalRequired = 0;
    let totalAllocated = 0;
    let totalInternalAllocated = 0;
    let totalExternalAllocated = 0;

    const periodCells = PERIODS.map((period) => {
      const key = `${row.project_id}__${row.role}__${period.periodKey}`;
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
      const assignments = getCellAssignments(row.project_id, row.role, period.periodKey);
      const hasExt = hasExtInAssignments(assignments);

      totalRequired += required;
      totalAllocated += allocated;
      totalInternalAllocated += internalAllocated;
      totalExternalAllocated += externalAllocated;

      const calloutHtml = encodeURIComponent(
        buildCellCalloutHtml(row.project_name, row.role, period.periodKey, assignments, releasedRows, historyRows)
      );

      const payload = encodeURIComponent(JSON.stringify({
        project_id: row.project_id,
        project_name: row.project_name,
        role: row.role,
        week: period.week,
        period_key: period.periodKey,
        required,
        allocated,
        internalAllocated,
        externalAllocated,
        diff,
      }));

      return `
        <td
          class="${getCellClass(required, allocated, period.periodKey, hasHistory, hasAllocationHistory, hasExt)}"
          data-cell='${payload}'
          data-callout='${calloutHtml}'
          data-row-index="${rowIndex}"
          data-week="${period.week}"
          data-period-key="${period.periodKey}"
        >
          ${hasExt ? `<span class="cell-ext-corner">!</span>` : ""}
          R${formatNumber(required)}<br>
          A${allocatedDisplay}<br>
          D${formatNumber(diff)}${hasHistory || hasAllocationHistory ? `<span class="history-marker">â†º</span>` : ""}
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
        ${periodCells}
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
    const candidates = Array.from(plannerBody.querySelectorAll(`td[data-period-key="${previousSummary.period_from}"][data-cell]`));
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
          ${PERIODS.map(() => `<td class="cell-empty"></td>`).join("")}
        </tr>
      `;
    }
  }
});

// === RISORSE SHEET PATCH START ===
(function enableResourcesSheetPatch() {
  function qs(id) {
    return document.getElementById(id);
  }

  function safeNormalize(value) {
    if (typeof normalizeRole === "function") {
      return normalizeRole(value);
    }
    return String(value || "").trim().toUpperCase();
  }

  function safeEscape(value) {
    if (typeof escapeHtml === "function") {
      return escapeHtml(value);
    }
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function getResourcesArray() {
    if (Array.isArray(window.resourcesData)) return window.resourcesData;
    try {
      if (Array.isArray(resourcesData)) return resourcesData;
    } catch (error) {
      return [];
    }
    return [];
  }

  async function getFreshResources() {
    try {
      const data = await fetchJson("/api/resources");
      if (Array.isArray(data)) {
        try {
          resourcesData = data;
        } catch (error) {
          window.resourcesData = data;
        }
        return data;
      }
    } catch (error) {
      console.error("Errore lettura /api/resources", error);
    }

    return getResourcesArray();
  }

  function isExtResourceLocal(resource) {
    if (typeof isExternalResource === "function") {
      return isExternalResource(resource);
    }

    const text = safeNormalize(`${resource?.name || ""} ${resource?.role || ""} ${resource?.availability_note || ""}`);
    return text.includes("-EXT") || text.includes("EMPLOYER=EXT") || text.includes(" EXT ");
  }

  function openResourcesSheet() {
    const sheet = qs("resourcesSheet");
    if (!sheet) {
      alert("Foglio risorse non trovato nel layout.");
      return;
    }

    sheet.hidden = false;
    renderResourcesSheet();

    setTimeout(() => {
      sheet.scrollIntoView({ behavior: "smooth", block: "start" });
    }, 50);
  }

  function closeResourcesSheet() {
    const sheet = qs("resourcesSheet");
    if (sheet) sheet.hidden = true;
  }

  async function reloadResourcesSheet() {
    await getFreshResources();
    renderResourcesSheet();
  }

  function renderResourcesSheet() {
    const body = qs("resourcesSheetBody");
    const search = qs("resourcesSheetSearch");

    if (!body) return;

    const resources = getResourcesArray();
    const searchText = safeNormalize(search ? search.value : "");

    const rows = resources
      .filter((resource) => {
        if (!searchText) return true;
        const haystack = safeNormalize(
          `${resource.id || ""} ${resource.name || ""} ${resource.role || ""} ${resource.availability_note || ""}`
        );
        return haystack.includes(searchText);
      })
      .sort((a, b) => {
        const aExt = isExtResourceLocal(a) ? 1 : 0;
        const bExt = isExtResourceLocal(b) ? 1 : 0;
        if (aExt !== bExt) return aExt - bExt;
        return String(a.name || "").localeCompare(String(b.name || ""));
      });

    if (!rows.length) {
      body.innerHTML = `
        <tr>
          <td colspan="5" class="resources-sheet-empty">Nessuna risorsa trovata.</td>
        </tr>
      `;
      return;
    }

    body.innerHTML = rows.map((resource) => {
      const id = Number(resource.id || 0);
      const active = Number(resource.is_active) === 1;
      const extClass = isExtResourceLocal(resource) ? " resources-sheet-ext-row" : "";

      return `
        <tr data-resource-id="${id}" class="${extClass}">
          <td class="resources-sheet-id">${id}</td>
          <td>
            <input class="resources-sheet-input" data-field="name" value="${safeEscape(resource.name || "")}" />
          </td>
          <td>
            <input class="resources-sheet-input" data-field="role" value="${safeEscape(resource.role || "")}" />
          </td>
          <td class="resources-sheet-active-cell">
            <select class="resources-sheet-input" data-field="is_active">
              <option value="1" ${active ? "selected" : ""}>Attiva</option>
              <option value="0" ${!active ? "selected" : ""}>Cessata/Fuori contratto</option>
            </select>
          </td>
          <td>
            <textarea class="resources-sheet-note" data-field="availability_note">${safeEscape(resource.availability_note || "")}</textarea>
          </td>
        </tr>
      `;
    }).join("");
  }

  function collectResourcesSheetChanges() {
    const body = qs("resourcesSheetBody");
    if (!body) return [];

    const resources = getResourcesArray();
    const changes = [];

    body.querySelectorAll("tr[data-resource-id]").forEach((row) => {
      const id = Number(row.dataset.resourceId || 0);
      const original = resources.find((resource) => Number(resource.id) === id);
      if (!id || !original) return;

      const getValue = (field) => {
        const input = row.querySelector(`[data-field="${field}"]`);
        return input ? input.value : "";
      };

      const next = {
        id,
        name: getValue("name").trim(),
        role: getValue("role").trim(),
        is_active: Number(getValue("is_active")),
        availability_note: getValue("availability_note").trim(),
      };

      const changed =
        String(original.name || "") !== next.name ||
        String(original.role || "") !== next.role ||
        Number(original.is_active || 0) !== next.is_active ||
        String(original.availability_note || "") !== next.availability_note;

      if (changed) changes.push(next);
    });

    return changes;
  }

  async function saveResourcesSheet() {
    const changes = collectResourcesSheetChanges();

    if (!changes.length) {
      alert("Nessuna modifica da salvare.");
      return;
    }

    if (!confirm(`Salvare ${changes.length} modifica/e risorsa?`)) {
      return;
    }

    try {
      for (const resource of changes) {
        await fetchJson(`/api/resources/${resource.id}`, {
          method: "PUT",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            name: resource.name,
            role: resource.role,
            is_active: resource.is_active,
            availability_note: resource.availability_note,
          }),
        });
      }

      if (typeof loadAll === "function") {
        await loadAll();
      } else {
        await getFreshResources();
      }

      renderResourcesSheet();
      alert("Risorse salvate.");
    } catch (error) {
      console.error(error);
      alert("Errore salvataggio risorse. Serve abilitare/controllare PUT /api/resources/{id} nel backend.");
    }
  }

  function bindResourcesSheetPatch() {
    const topbar = document.querySelector(".topbar-actions");
    const sheet = qs("resourcesSheet");

    if (topbar && !qs("topResourcesBtn")) {
      const buttons = Array.from(topbar.querySelectorAll("button"));
      const resourcesButton = buttons.find((button) => String(button.textContent || "").trim().toUpperCase() === "RISORSE");
      const refreshButton = buttons.find((button) => String(button.textContent || "").trim().toUpperCase() === "AGGIORNA");

      if (resourcesButton) resourcesButton.id = "topResourcesBtn";
      if (refreshButton) refreshButton.id = "topRefreshBtn";
    }

    const topResourcesBtn = qs("topResourcesBtn");
    const topRefreshBtn = qs("topRefreshBtn");
    const closeBtn = qs("resourcesSheetCloseBtn");
    const reloadBtn = qs("resourcesSheetReloadBtn");
    const saveBtn = qs("resourcesSheetSaveBtn");
    const search = qs("resourcesSheetSearch");

    if (topResourcesBtn && !topResourcesBtn.dataset.resourcesSheetBound) {
      topResourcesBtn.dataset.resourcesSheetBound = "1";
      topResourcesBtn.addEventListener("click", openResourcesSheet);
    }

    if (topRefreshBtn && !topRefreshBtn.dataset.resourcesSheetBound) {
      topRefreshBtn.dataset.resourcesSheetBound = "1";
      topRefreshBtn.addEventListener("click", async () => {
        if (typeof loadAll === "function") {
          await loadAll();
        } else {
          await getFreshResources();
        }
        if (sheet && !sheet.hidden) renderResourcesSheet();
      });
    }

    if (closeBtn && !closeBtn.dataset.resourcesSheetBound) {
      closeBtn.dataset.resourcesSheetBound = "1";
      closeBtn.addEventListener("click", closeResourcesSheet);
    }

    if (reloadBtn && !reloadBtn.dataset.resourcesSheetBound) {
      reloadBtn.dataset.resourcesSheetBound = "1";
      reloadBtn.addEventListener("click", reloadResourcesSheet);
    }

    if (saveBtn && !saveBtn.dataset.resourcesSheetBound) {
      saveBtn.dataset.resourcesSheetBound = "1";
      saveBtn.addEventListener("click", saveResourcesSheet);
    }

    if (search && !search.dataset.resourcesSheetBound) {
      search.dataset.resourcesSheetBound = "1";
      search.addEventListener("input", renderResourcesSheet);
    }
  }

  window.openResourcesSheet = openResourcesSheet;
  window.renderResourcesSheet = renderResourcesSheet;

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", bindResourcesSheetPatch);
  } else {
    bindResourcesSheetPatch();
  }

  setTimeout(bindResourcesSheetPatch, 500);
  setTimeout(bindResourcesSheetPatch, 1500);
})();
// === RISORSE SHEET PATCH END ===

// === SHOW ZERO TOGGLE RERENDER PATCH START ===
(function patchShowZeroToggleRerender() {
  function findShowZeroToggle() {
    const candidates = Array.from(document.querySelectorAll("input[type='checkbox']"));

    return candidates.find((input) => {
      const label = input.closest("label");
      const text = String(label ? label.textContent : "").toUpperCase();
      return (
        text.includes("MOSTRA") &&
        (
          text.includes("SENZA FABBISOGNO") ||
          text.includes("ZERO") ||
          text.includes("COMMESSE")
        )
      );
    }) || null;
  }

  async function rerenderPlannerAfterShowZeroChange() {
    try {
      if (typeof renderPlanner === "function") {
        renderPlanner();
        return;
      }

      if (typeof loadAll === "function") {
        await loadAll();
        return;
      }

      window.location.reload();
    } catch (error) {
      console.error("Errore refresh planner dopo toggle mostra senza fabbisogno", error);
      window.location.reload();
    }
  }

  function bindShowZeroToggle() {
    const toggle = findShowZeroToggle();

    if (!toggle) {
      return;
    }

    if (toggle.dataset.showZeroRerenderBound === "1") {
      return;
    }

    toggle.dataset.showZeroRerenderBound = "1";

    toggle.addEventListener("change", () => {
      rerenderPlannerAfterShowZeroChange();
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", bindShowZeroToggle);
  } else {
    bindShowZeroToggle();
  }

  setTimeout(bindShowZeroToggle, 500);
  setTimeout(bindShowZeroToggle, 1500);
})();
// === SHOW ZERO TOGGLE RERENDER PATCH END ===

// === BOTTOM DETAIL PANEL PATCH START ===
(function patchBottomDetailPanel() {
  function qs(id) {
    return document.getElementById(id);
  }

  function safeEscape(value) {
    if (typeof escapeHtml === "function") return escapeHtml(value);
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function safeNormalize(value) {
    if (typeof normalizeRole === "function") return normalizeRole(value);
    return String(value || "").trim().toUpperCase();
  }

  function safeFormatNumber(value) {
    if (typeof formatNumber === "function") return formatNumber(value);
    const n = Number(value || 0);
    return Number.isInteger(n) ? String(n) : n.toFixed(1);
  }

  function safeFormatLoad(value) {
    if (typeof formatLoad === "function") return formatLoad(value);
    const n = Number(value || 100);
    if (n === 100) return "1:1";
    if (n === 50) return "1/2";
    return `${n}%`;
  }

  function getBottomBox() {
    let box = qs("plannerBottomDetail");

    if (box) return box;

    const bottomContent = document.querySelector(".planner-bottom .bottom-content");

    if (bottomContent) {
      bottomContent.id = "plannerBottomDetail";
      return bottomContent;
    }

    return null;
  }

  function getSelectedCellFromDom() {
    const selected =
      document.querySelector(".planner-cell.selected") ||
      document.querySelector(".planner-cell.is-selected") ||
      document.querySelector(".planner-cell-clickable.selected") ||
      document.querySelector(".planner-cell-clickable.is-selected") ||
      document.querySelector("[data-cell].selected") ||
      document.querySelector("[data-cell].is-selected");

    if (!selected) return null;

    if (selected.dataset && selected.dataset.cell) {
      try {
        return JSON.parse(decodeURIComponent(selected.dataset.cell));
      } catch (error) {
        try {
          return JSON.parse(selected.dataset.cell);
        } catch (innerError) {
          console.error("Errore parsing data-cell", innerError);
        }
      }
    }

    return null;
  }

  function getSelectionSummarySafe() {
    if (typeof getSelectionSummary === "function") {
      try {
        return getSelectionSummary();
      } catch (error) {
        console.error("Errore getSelectionSummary", error);
      }
    }
    return null;
  }

  function getCellData() {
    const domCell = getSelectedCellFromDom();

    if (domCell) {
      return domCell;
    }

    const summary = getSelectionSummarySafe();

    if (!summary) return null;

    try {
      const row = Array.from(document.querySelectorAll("[data-cell]")).find((el) => {
        try {
          const parsed = JSON.parse(decodeURIComponent(el.dataset.cell));
          return (
            Number(parsed.project_id) === Number(summary.project_id) &&
            safeNormalize(parsed.role || "") === safeNormalize(summary.role || "") &&
            Number(parsed.period_key || parsed.periodKey || 0) === Number(summary.period_from || summary.periodKey || 0)
          );
        } catch (error) {
          return false;
        }
      });

      if (row) {
        return JSON.parse(decodeURIComponent(row.dataset.cell));
      }
    } catch (error) {
      return null;
    }

    return null;
  }

  function getResourceByIdLocal(id) {
    try {
      return resourcesData.find((resource) => Number(resource.id) === Number(id)) || null;
    } catch (error) {
      return null;
    }
  }

  function isExternalLocal(item, resource) {
    if (typeof isExternalResource === "function") {
      try {
        if (resource && isExternalResource(resource)) return true;
        if (item && isExternalResource(item)) return true;
      } catch (error) {
        // ignore
      }
    }

    const text = safeNormalize(
      `${item?.resource_name || ""} ${item?.resource_role || ""} ${item?.name || ""} ${item?.role || ""} ${item?.note || ""} ${resource?.name || ""} ${resource?.role || ""} ${resource?.availability_note || ""}`
    );

    return text.includes("-EXT") || text.includes("EMPLOYER=EXT") || text.includes(" EXT ");
  }

  function resourceNameLocal(item) {
    const resource = getResourceByIdLocal(item?.resource_id);
    const raw =
      item?.resource_name ||
      item?.resourceName ||
      item?.name ||
      resource?.name ||
      "";

    if (typeof shortResourceName === "function") {
      return shortResourceName(raw);
    }

    return raw;
  }

  function roleLabelLocal(item, rowRole) {
    const resource = getResourceByIdLocal(item?.resource_id);
    const realRole =
      item?.resource_role ||
      item?.resourceRole ||
      resource?.role ||
      item?.role ||
      "";

    if (isExternalLocal(item, resource)) return "EXT";

    if (typeof mansioneLabel === "function") {
      try {
        return mansioneLabel(rowRole, realRole);
      } catch (error) {
        return realRole || rowRole || "";
      }
    }

    return realRole || rowRole || "";
  }

  function statusLabelLocal(item, historical) {
    const resource = getResourceByIdLocal(item?.resource_id);

    if (isExternalLocal(item, resource)) return "EXT";

    const text = safeNormalize(`${item?.status || ""} ${item?.reason || ""} ${item?.note || ""} ${resource?.availability_note || ""}`);

    if (historical) {
      if (text.includes("FUORI") || text.includes("CESS")) return "FUORI_CONTRATTO";
      if (text.includes("INDISP")) return "INDISP";
      return "STORICO";
    }

    if (text.includes("FUORI") || text.includes("CESS")) return "FUORI_CONTRATTO";
    if (text.includes("INDISP")) return "INDISP";

    return "ACTIVE";
  }

  function loadLabelLocal(item, historical) {
    if (historical) return "0:1";

    if (item?.display_weight) return item.display_weight;

    const load = item?.load_percent ?? item?.loadPercent ?? 100;
    return safeFormatLoad(load);
  }

  function renderResourceRows(items, rowRole, historical) {
    if (!Array.isArray(items) || !items.length) {
      return `<div class="bottom-detail-muted">Nessuna risorsa ${historical ? "storica/non conteggiata" : "attiva conteggiata"}.</div>`;
    }

    return items.map((item) => {
      const resource = getResourceByIdLocal(item?.resource_id);
      const ext = isExternalLocal(item, resource);

      const classes = [
        "bottom-detail-resource-row",
        historical ? "bottom-detail-resource-history" : "",
        ext ? "bottom-detail-resource-ext" : "",
      ].join(" ");

      return `
        <div class="${classes}">
          <div class="bottom-detail-resource-name">${safeEscape(resourceNameLocal(item))}</div>
          <div class="bottom-detail-resource-role">${safeEscape(roleLabelLocal(item, rowRole))}</div>
          <div class="bottom-detail-resource-load">${safeEscape(loadLabelLocal(item, historical))}</div>
          <div class="bottom-detail-resource-status">${safeEscape(statusLabelLocal(item, historical))}</div>
        </div>
      `;
    }).join("");
  }

  function getActiveAndHistoricalFromCell(cell, summary) {
    const active = Array.isArray(cell?.active_resources) ? cell.active_resources : [];
    const historical = Array.isArray(cell?.historical_resources) ? cell.historical_resources : [];

    if (active.length || historical.length) {
      return { active, historical };
    }

    // fallback se la cella non contiene arrays backend
    try {
      const selectedPeriod = Number(cell?.period_key || summary?.period_from || 0);
      const projectId = Number(cell?.project_id || summary?.project_id || 0);
      const role = safeNormalize(cell?.role || summary?.role || "");

      const activeFallback = [];
      const historicalFallback = [];

      allocationsData.forEach((allocation) => {
        if (
          Number(allocation.project_id) !== projectId ||
          safeNormalize(allocation.role || "") !== role ||
          Number(normalizePeriodKey(allocation)) !== selectedPeriod
        ) {
          return;
        }

        const resource = getResourceByIdLocal(allocation.resource_id);

        if (!resource) return;

        if (typeof isResourceAvailableForPeriod === "function" && isResourceAvailableForPeriod(resource, selectedPeriod)) {
          activeFallback.push(allocation);
        } else {
          historicalFallback.push(allocation);
        }
      });

      return {
        active: activeFallback,
        historical: historicalFallback,
      };
    } catch (error) {
      return { active, historical };
    }
  }

  function renderNormalBottomDetail(box, cell, summary) {
    const required = Number(cell?.required ?? summary?.required ?? 0);
    const allocated = Number(cell?.allocated ?? summary?.allocated ?? 0);
    const internal = Number(cell?.internal_allocated ?? cell?.internalAllocated ?? allocated ?? 0);
    const external = Number(cell?.external_allocated ?? cell?.externalAllocated ?? 0);
    const diff = required - allocated;

    const projectName = cell?.project_name || summary?.project_name || "";
    const role = cell?.role || summary?.role || "";
    const week = Number(cell?.week || summary?.week_from || 0);
    const periodKey = Number(cell?.period_key || summary?.period_from || 0);

    const resources = getActiveAndHistoricalFromCell(cell, summary);
    const badges = Array.isArray(cell?.badges) ? cell.badges : [];

    const allocatedText = external > 0
      ? `${safeFormatNumber(internal)}+${safeFormatNumber(external)}`
      : safeFormatNumber(allocated);

    const badgesHtml = badges.length
      ? `<div class="bottom-detail-badges">${badges.map((badge) => `<span>${safeEscape(badge)}</span>`).join("")}</div>`
      : "";

    box.innerHTML = `
      <div class="bottom-detail-header">
        <div>
          <strong>${safeEscape(projectName)}</strong>
          <span>/ ${safeEscape(role)}</span>
          <span>/ W${String(week).padStart(2, "0")}</span>
        </div>
        <div class="bottom-detail-numbers">
          <span>R${safeFormatNumber(required)}</span>
          <span>A${safeEscape(allocatedText)}</span>
          <span>D${safeFormatNumber(diff)}</span>
        </div>
      </div>

      ${badgesHtml}

      <div class="bottom-detail-grid">
        <div class="bottom-detail-block">
          <div class="bottom-detail-title">Risorse attive conteggiate</div>
          ${renderResourceRows(resources.active, role, false)}
        </div>

        <div class="bottom-detail-block">
          <div class="bottom-detail-title">Storico / fuori contratto / non conteggiati</div>
          ${renderResourceRows(resources.historical, role, true)}
        </div>
      </div>
    `;
  }

  async function renderOverallBottomDetail(box, cell, summary) {
    const projectName = cell?.project_name || summary?.project_name || "";
    const role = cell?.role || summary?.role || "";
    const week = Number(cell?.week || summary?.week_from || 0);
    const projectId = Number(cell?.project_id || summary?.project_id || 0);

    const required = Number(cell?.required ?? summary?.required ?? 0);
    const allocated = Number(cell?.allocated ?? summary?.allocated ?? 0);
    const internal = Number(cell?.internal_allocated ?? cell?.internalAllocated ?? allocated ?? 0);
    const external = Number(cell?.external_allocated ?? cell?.externalAllocated ?? 0);
    const diff = required - allocated;

    const allocatedText = external > 0
      ? `${safeFormatNumber(internal)}+${safeFormatNumber(external)}`
      : safeFormatNumber(allocated);

    box.innerHTML = `
      <div class="bottom-detail-header">
        <div>
          <strong>${safeEscape(projectName)}</strong>
          <span>/ ${safeEscape(role)}</span>
          <span>/ W${String(week).padStart(2, "0")}</span>
        </div>
        <div class="bottom-detail-numbers">
          <span>R${safeFormatNumber(required)}</span>
          <span>A${safeEscape(allocatedText)}</span>
          <span>D${safeFormatNumber(diff)}</span>
        </div>
      </div>

      <div class="bottom-detail-muted">Caricamento dettaglio OVERALL OFFICINA...</div>
    `;

    try {
      const params = new URLSearchParams({
        project_id: String(projectId),
        role: role,
        week: String(week),
      });

      const data = await fetchJson(`/api/workshop-breakdown?${params.toString()}`);
      const sources = Array.isArray(data.sources) ? data.sources : [];
      const resources = getActiveAndHistoricalFromCell(cell, summary);

      const sourceRows = sources.length
        ? sources.map((source) => `
            <div class="bottom-overall-row">
              <div class="bottom-overall-project">${safeEscape(source.project_name || "")}</div>
              <div class="bottom-overall-role">${safeEscape(source.role || role)}</div>
              <div class="bottom-overall-required">R${safeFormatNumber(source.required || 0)}</div>
            </div>
          `).join("")
        : `<div class="bottom-detail-muted">Nessuna sottocommessa trovata per questa cella.</div>`;

      box.innerHTML = `
        <div class="bottom-detail-header">
          <div>
            <strong>${safeEscape(projectName)}</strong>
            <span>/ ${safeEscape(role)}</span>
            <span>/ W${String(week).padStart(2, "0")}</span>
          </div>
          <div class="bottom-detail-numbers">
            <span>R${safeFormatNumber(required)}</span>
            <span>A${safeEscape(allocatedText)}</span>
            <span>D${safeFormatNumber(diff)}</span>
          </div>
        </div>

        <div class="bottom-detail-grid bottom-detail-grid-wide">
          <div class="bottom-detail-block">
            <div class="bottom-detail-title">Dettaglio sottocommesse officina</div>
            ${sourceRows}
            <div class="bottom-overall-total">
              <div>Totale OVERALL</div>
              <div>R${safeFormatNumber(data.total_required || 0)}</div>
            </div>
          </div>

          <div class="bottom-detail-block">
            <div class="bottom-detail-title">Allocazioni su OVERALL OFFICINA</div>
            ${renderResourceRows(resources.active, role, false)}

            <div class="bottom-detail-title bottom-detail-title-secondary">Storico / fuori contratto</div>
            ${renderResourceRows(resources.historical, role, true)}
          </div>
        </div>
      `;
    } catch (error) {
      console.error("Errore dettaglio OVERALL", error);
      box.innerHTML += `
        <div class="bottom-detail-error">Errore caricamento dettaglio OVERALL OFFICINA.</div>
      `;
    }
  }

  async function renderBottomDetailPanel() {
    const box = getBottomBox();
    if (!box) return;

    const summary = getSelectionSummarySafe();
    const cell = getCellData();

    if (!summary && !cell) {
      box.innerHTML = `<div class="bottom-detail-muted">Seleziona una cella per vedere dettaglio, storico e note operative.</div>`;
      return;
    }

    const projectName = cell?.project_name || summary?.project_name || "";
    const isOverall = safeNormalize(projectName).includes("OVERALL OFFICINA");

    if (isOverall) {
      await renderOverallBottomDetail(box, cell || {}, summary || {});
    } else {
      renderNormalBottomDetail(box, cell || {}, summary || {});
    }
  }

  function bindBottomDetailClicks() {
    document.addEventListener("click", (event) => {
      const target = event.target.closest("[data-cell], .planner-cell, .planner-cell-clickable");

      if (!target) return;

      setTimeout(renderBottomDetailPanel, 60);
      setTimeout(renderBottomDetailPanel, 250);
    });
  }

  function patchUpdateSidePanel() {
    if (typeof updateSidePanelFromSelection !== "function") return;

    if (window.__bottomDetailUpdatePatched) return;
    window.__bottomDetailUpdatePatched = true;

    const original = updateSidePanelFromSelection;

    window.updateSidePanelFromSelection = function patchedUpdateSidePanelFromSelection() {
      const result = original.apply(this, arguments);
      setTimeout(renderBottomDetailPanel, 50);
      return result;
    };
  }

  window.renderBottomDetailPanel = renderBottomDetailPanel;

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => {
      bindBottomDetailClicks();
      patchUpdateSidePanel();
    });
  } else {
    bindBottomDetailClicks();
    patchUpdateSidePanel();
  }

  setTimeout(patchUpdateSidePanel, 500);
  setTimeout(patchUpdateSidePanel, 1500);
})();
// === BOTTOM DETAIL PANEL PATCH END ===

// === ALLOCATION HISTORY BOTTOM PATCH START ===
(function patchAllocationHistoryBottom() {
  window.allocationHistoryData = window.allocationHistoryData || [];

  function safeNormalize(value) {
    if (typeof normalizeRole === "function") return normalizeRole(value);
    return String(value || "").trim().toUpperCase();
  }

  function safeFormatLoad(value) {
    if (typeof formatLoad === "function") return formatLoad(value);
    const n = Number(value || 100);
    if (n === 100) return "1:1";
    if (n === 50) return "1/2";
    return `${n}%`;
  }

  async function loadAllocationHistoryData() {
    try {
      if (typeof fetchJson !== "function") return;

      const rows = await fetchJson("/api/allocation-history");

      if (Array.isArray(rows)) {
        window.allocationHistoryData = rows;
      }
    } catch (error) {
      console.error("Errore caricamento allocation_history", error);
      window.allocationHistoryData = [];
    }
  }

  function getSelectedContextForHistory() {
    let summary = null;

    if (typeof getSelectionSummary === "function") {
      try {
        summary = getSelectionSummary();
      } catch (error) {
        summary = null;
      }
    }

    let cell = null;
    const selected =
      document.querySelector("[data-cell].selected") ||
      document.querySelector("[data-cell].is-selected") ||
      document.querySelector(".planner-cell.selected[data-cell]") ||
      document.querySelector(".planner-cell.is-selected[data-cell]") ||
      document.querySelector(".planner-cell-clickable.selected[data-cell]") ||
      document.querySelector(".planner-cell-clickable.is-selected[data-cell]");

    if (selected && selected.dataset.cell) {
      try {
        cell = JSON.parse(decodeURIComponent(selected.dataset.cell));
      } catch (error) {
        try {
          cell = JSON.parse(selected.dataset.cell);
        } catch (innerError) {
          cell = null;
        }
      }
    }

    return {
      projectId: Number(cell?.project_id || summary?.project_id || 0),
      projectName: cell?.project_name || summary?.project_name || "",
      role: safeNormalize(cell?.role || summary?.role || ""),
      periodKey: Number(cell?.period_key || summary?.period_from || 0),
      week: Number(cell?.week || summary?.week_from || 0),
    };
  }

  function historyMatchesSelection(row, context) {
    if (!row || !context.projectId || !context.role || !context.periodKey) {
      return false;
    }

    return (
      Number(row.project_id) === Number(context.projectId) &&
      safeNormalize(row.role || "") === context.role &&
      Number(row.period_key || 0) === Number(context.periodKey)
    );
  }

  function historyRowKey(row) {
    return `${row.history_id || row.id || ""}__${row.resource_id || ""}__${row.resource_name || ""}`;
  }

  function appendHistoryRowsToBottom() {
    const box = document.getElementById("plannerBottomDetail") || document.querySelector(".planner-bottom .bottom-content");

    if (!box) return;

    const context = getSelectedContextForHistory();

    if (!context.projectId || !context.role || !context.periodKey) {
      return;
    }

    const rows = (window.allocationHistoryData || []).filter((row) => historyMatchesSelection(row, context));

    if (!rows.length) {
      return;
    }

    const historyBlock = Array.from(box.querySelectorAll(".bottom-detail-block")).find((block) => {
      return String(block.textContent || "").toUpperCase().includes("STORICO");
    });

    if (!historyBlock) {
      return;
    }

    const existingKeys = new Set(
      Array.from(historyBlock.querySelectorAll("[data-history-key]")).map((el) => el.dataset.historyKey)
    );

    const html = rows
      .filter((row) => !existingKeys.has(historyRowKey(row)))
      .map((row) => {
        const key = historyRowKey(row);
        const status = row.reason || "STORICO";
        const load = "0:1";

        return `
          <div class="bottom-detail-resource-row bottom-detail-resource-history" data-history-key="${key}">
            <div class="bottom-detail-resource-name">${row.resource_name || ""}</div>
            <div class="bottom-detail-resource-role">${row.resource_role || row.role || ""}</div>
            <div class="bottom-detail-resource-load">${load}</div>
            <div class="bottom-detail-resource-status">${status}</div>
          </div>
        `;
      })
      .join("");

    if (!html) {
      return;
    }

    const muted = historyBlock.querySelector(".bottom-detail-muted");

    if (muted) {
      muted.remove();
    }

    historyBlock.insertAdjacentHTML("beforeend", html);
  }

  async function refreshHistoryAndBottom() {
    await loadAllocationHistoryData();

    if (typeof window.renderBottomDetailPanel === "function") {
      await window.renderBottomDetailPanel();
    }

    setTimeout(appendHistoryRowsToBottom, 80);
    setTimeout(appendHistoryRowsToBottom, 250);
  }

  const originalLoadAll = typeof loadAll === "function" ? loadAll : null;

  if (originalLoadAll && !window.__allocationHistoryLoadAllPatched) {
    window.__allocationHistoryLoadAllPatched = true;

    window.loadAll = async function patchedLoadAllWithHistory() {
      const result = await originalLoadAll.apply(this, arguments);
      await loadAllocationHistoryData();
      return result;
    };
  }

  document.addEventListener("click", (event) => {
    const target = event.target.closest("[data-cell], .planner-cell, .planner-cell-clickable");

    if (!target) return;

    setTimeout(appendHistoryRowsToBottom, 120);
    setTimeout(appendHistoryRowsToBottom, 350);
  });

  loadAllocationHistoryData();
})();
// === ALLOCATION HISTORY BOTTOM PATCH END ===

// === FRONTEND OVERALL DEMANDS FROM ROLLUP PATCH START ===
(function patchFrontendOverallDemandsFromRollup() {
  window.workshopRequiredMapData = window.workshopRequiredMapData || [];

  function safeNormalize(value) {
    if (typeof normalizeRole === "function") return normalizeRole(value);
    return String(value || "").trim().toUpperCase();
  }

  function safeNumber(value) {
    const n = Number(value || 0);
    return Number.isFinite(n) ? n : 0;
  }

  function getOverallProjectId() {
    try {
      const project = (projectsData || []).find((item) => {
        return safeNormalize(`${item.name || ""} ${item.note || ""}`).includes("OVERALL OFFICINA");
      });

      return project ? Number(project.id) : 0;
    } catch (error) {
      return 0;
    }
  }

  async function loadWorkshopRequiredMapForFrontend() {
    try {
      if (typeof fetchJson !== "function") return [];

      const rows = await fetchJson("/api/workshop-required-map");

      if (Array.isArray(rows)) {
        window.workshopRequiredMapData = rows;
        return rows;
      }

      window.workshopRequiredMapData = [];
      return [];
    } catch (error) {
      console.error("Errore caricamento /api/workshop-required-map", error);
      window.workshopRequiredMapData = [];
      return [];
    }
  }

  function applyWorkshopRollupToDemandsData() {
    if (!Array.isArray(window.workshopRequiredMapData) || !window.workshopRequiredMapData.length) {
      return;
    }

    if (!Array.isArray(demandsData)) {
      return;
    }

    const overallProjectId = getOverallProjectId();

    if (!overallProjectId) {
      return;
    }

    // Togliamo i demands importati direttamente su OVERALL OFFICINA.
    // Nel vecchio planner OVERALL OFFICINA e' un rollup, quindi R deve arrivare dalle figlie officina.
    const withoutOverallDirectDemands = demandsData.filter((demand) => {
      return Number(demand.project_id) !== Number(overallProjectId);
    });

    const rollupDemands = window.workshopRequiredMapData
      .filter((row) => Number(row.overall_project_id || row.project_id) === Number(overallProjectId))
      .map((row, index) => {
        const periodKey = Number(row.period_key || 0);
        const week = Number(row.week || (periodKey % 100));

        return {
          id: `workshop-rollup-${index}`,
          project_id: overallProjectId,
          week: week,
          period_key: periodKey,
          role: safeNormalize(row.role || ""),
          quantity: safeNumber(row.required ?? row.quantity),
          note: "WORKSHOP_ROLLUP_SOURCES"
        };
      })
      .filter((row) => row.role && row.period_key && safeNumber(row.quantity) !== 0);

    demandsData = withoutOverallDirectDemands.concat(rollupDemands);

    try {
      window.demandsData = demandsData;
    } catch (error) {
      // ignore
    }
  }

  async function reloadOverallRollupAndRender() {
    await loadWorkshopRequiredMapForFrontend();
    applyWorkshopRollupToDemandsData();

    if (typeof renderPlanner === "function") {
      renderPlanner();
    }
  }

  const originalLoadAll = typeof loadAll === "function" ? loadAll : null;

  if (originalLoadAll && !window.__frontendOverallDemandsLoadAllPatched) {
    window.__frontendOverallDemandsLoadAllPatched = true;

    loadAll = async function patchedLoadAllFrontendOverallDemands() {
      const result = await originalLoadAll.apply(this, arguments);
      await loadWorkshopRequiredMapForFrontend();
      applyWorkshopRollupToDemandsData();

      if (typeof renderPlanner === "function") {
        renderPlanner();
      }

      return result;
    };

    try {
      window.loadAll = loadAll;
    } catch (error) {
      // ignore
    }
  }

  const originalRenderPlanner = typeof renderPlanner === "function" ? renderPlanner : null;

  if (originalRenderPlanner && !window.__frontendOverallDemandsRenderPatched) {
    window.__frontendOverallDemandsRenderPatched = true;

    renderPlanner = function patchedRenderPlannerFrontendOverallDemands() {
      applyWorkshopRollupToDemandsData();
      return originalRenderPlanner.apply(this, arguments);
    };

    try {
      window.renderPlanner = renderPlanner;
    } catch (error) {
      // ignore
    }
  }

  loadWorkshopRequiredMapForFrontend().then(() => {
    applyWorkshopRollupToDemandsData();

    if (typeof renderPlanner === "function") {
      renderPlanner();
    }
  });

  window.reloadOverallRollupAndRender = reloadOverallRollupAndRender;
})();
// === FRONTEND OVERALL DEMANDS FROM ROLLUP PATCH END ===

