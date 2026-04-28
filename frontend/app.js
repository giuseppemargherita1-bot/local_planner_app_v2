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

// === OLD WORKFLOW PLANNER HELPERS START ===
const OLD_WORKFLOW_ROLE_ORDER = [
  "CAPO CANTIERE",
  "QUALITY CONTROL / WELDING INSPECTOR",
  "ASPP",
  "CAPO SQUADRA",
  "TUBISTA",
  "CARPENTIERE",
  "SOLLEVAMENTI",
  "AUTISTA",
  "SALDATORE TIG-ELETTRODO",
  "SALDATORE FILO",
  "PWHT",
  "GENERICO",
  "MAGAZZINIERE",
  "MECCANICO",
  "MECCANICO SERVICE",
  "MONTATORE",
  "MANDRINATORE",
  "ELETTRICISTA",
  "PONTEGGIATORE",
  "COIBENTATORE",
  "VERNICIATORE",
];

function oldWorkflowProjectSortKey(projectName) {
  const text = String(projectName || "").trim().toUpperCase();

  if (text.includes("OVERALL OFFICINA")) {
    return "999999_OVERALL_OFFICINA";
  }

  const match = text.match(/^([0-9]+)[_-]?([0-9]+)?/);
  if (!match) {
    return text;
  }

  const main = String(match[1] || "").padStart(6, "0");
  const sub = String(match[2] || "").padStart(6, "0");

  return `${main}_${sub}_${text}`;
}

function oldWorkflowRoleSortKey(role) {
  const normalized = normalizeRole(role);
  const index = OLD_WORKFLOW_ROLE_ORDER.indexOf(normalized);

  if (index >= 0) {
    return `${String(index).padStart(4, "0")}_${normalized}`;
  }

  return `9999_${normalized}`;
}

function oldWorkflowComparePlannerRows(a, b) {
  const projectCompare = oldWorkflowProjectSortKey(a.project_name).localeCompare(
    oldWorkflowProjectSortKey(b.project_name),
    "it",
    { numeric: true, sensitivity: "base" }
  );

  if (projectCompare !== 0) {
    return projectCompare;
  }

  const roleCompare = oldWorkflowRoleSortKey(a.role).localeCompare(
    oldWorkflowRoleSortKey(b.role),
    "it",
    { numeric: true, sensitivity: "base" }
  );

  if (roleCompare !== 0) {
    return roleCompare;
  }

  return normalizeRole(a.role).localeCompare(normalizeRole(b.role), "it", {
    numeric: true,
    sensitivity: "base",
  });
}

function oldWorkflowFindPlannerProjectFilter() {
  return document.getElementById("plannerProjectFilter") ||
    Array.from(document.querySelectorAll("select")).find((select) => {
      const first = normalizeRole(select.options?.[0]?.textContent || "");
      return first.includes("TUTTE LE COMMESSE");
    }) ||
    null;
}

function oldWorkflowFindPlannerRoleFilter() {
  return document.getElementById("plannerRoleFilter") ||
    Array.from(document.querySelectorAll("select")).find((select) => {
      const first = normalizeRole(select.options?.[0]?.textContent || "");
      return first.includes("TUTTE LE MANSIONI");
    }) ||
    null;
}

function oldWorkflowProjectFilterValue() {
  const filter = oldWorkflowFindPlannerProjectFilter();
  return filter ? String(filter.value || "") : "";
}

function oldWorkflowRoleFilterValue() {
  const filter = oldWorkflowFindPlannerRoleFilter();
  return filter ? normalizeRole(filter.value || "") : "";
}

function oldWorkflowRowHasAnyValue(row, demandMap, allocationMaps) {
  for (const period of PERIODS) {
    const key = `${Number(row.project_id)}__${normalizeRole(row.role)}__${Number(period.periodKey)}`;

    if (Number(demandMap.get(key) || 0) > 0) {
      return true;
    }

    if (Number(allocationMaps.totalMap.get(key) || 0) > 0) {
      return true;
    }
  }

  return false;
}
// === OLD WORKFLOW PLANNER HELPERS END ===

function buildPlannerRows(projects, demands) {
  const projectsById = new Map(projects.map((project) => [Number(project.id), project]));
  const rowsMap = new Map();
  const demandMap = buildDemandMap(demands);
  const mapsForVisibility = buildAllocationMaps(allocationsData, resourcesData);

  const showZero = document.querySelector(".inline-check input")?.checked;
  const projectFilter = oldWorkflowProjectFilterValue();
  const roleFilter = oldWorkflowRoleFilterValue();

  function addRow(project, role, source) {
    if (!project || !role) {
      return;
    }

    const projectId = Number(project.id);
    const normalizedRole = normalizeRole(role);

    if (!projectId || !normalizedRole) {
      return;
    }

    const rowKey = `${projectId}__${normalizedRole}`;

    if (!rowsMap.has(rowKey)) {
      rowsMap.set(rowKey, {
        row_key: rowKey,
        project_id: projectId,
        project_name: project.name,
        role: normalizedRole,
        source,
      });
    }
  }

  // Fonte 1: fabbisogni. Questa e' la struttura commessa -> mansioni -> settimane.
  for (const demand of demands) {
    const project = projectsById.get(Number(demand.project_id));

    if (!project) {
      continue;
    }

    addRow(project, demand.role, "demand");
  }

  // Fonte 2: allocazioni. Serve per non perdere R0/A1, surplus, APPELLA/ZACCHEO.
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

    addRow(project, allocation.role || resource.role || "", "allocation");
  }

  let rows = Array.from(rowsMap.values());

  rows = rows.filter((row) => {
    if (projectFilter && String(row.project_id) !== String(projectFilter)) {
      return false;
    }

    if (roleFilter && normalizeRole(row.role) !== roleFilter) {
      return false;
    }

    if (showZero) {
      return true;
    }

    if (!shouldShowPlannerRow(row.project_id, row.role)) {
      return false;
    }

    return oldWorkflowRowHasAnyValue(row, demandMap, mapsForVisibility);
  });

  rows.sort(oldWorkflowComparePlannerRows);

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

// === PERIOD KEY FIELD SOURCE FIX START ===
function displayPeriodKey(value) {
  const numeric = Number(value || 0);
  if (!numeric) return "";
  if (numeric >= 1000) return String(numeric);
  return String(periodKeyFromWeek(numeric));
}

function writePeriodInput(inputOrId, value) {
  const input = typeof inputOrId === "string" ? document.getElementById(inputOrId) : inputOrId;
  if (!input) return;
  input.value = displayPeriodKey(value);
  input.min = "1000";
  input.max = "9999";
  input.step = "1";
  input.placeholder = "2617";
}

function readPeriodInput(inputOrId, fallback = 0) {
  const input = typeof inputOrId === "string" ? document.getElementById(inputOrId) : inputOrId;
  const raw = Number(input?.value || fallback || 0);
  if (!raw) return 0;
  return raw >= 1000 ? raw : periodKeyFromWeek(raw);
}
// === PERIOD KEY FIELD SOURCE FIX END ===

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
  detailWeekFrom.value = summary.period_from;
  detailWeekTo.value = summary.period_to;
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
      week: String(summary.period_from),
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
            week_from: summary.period_from,
            week_to: summary.period_to,
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
        "Conflitto su più settimane",
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
      week: Number(summary.period_from),
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
    week_from: Number(summary.period_from),
    week_to: Number(summary.period_to),
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
      week_from: Number(summary.period_from),
      week_to: Number(summary.period_to),
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
      week_from: Number(summary.period_from),
      week_to: Number(summary.period_to),
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
    parts.push(`<div class="cell-callout-section">Non più conteggiati</div>`);
    released.forEach((item) => {
      const line = `${formatAssignmentLine(item, role, periodKey)} | ${item.reason || "storico"}`;
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

// === PLANNER DEMAND ROLE PACKAGE START ===
(function plannerDemandRolePackage() {
  function qs(id) {
    return document.getElementById(id);
  }

  function norm(value) {
    if (typeof normalizeRole === "function") return normalizeRole(value);
    return String(value || "").trim().toUpperCase();
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

  function getSelectedCellElements() {
    return Array.from(document.querySelectorAll(
      "[data-cell].selected, [data-cell].is-selected, .planner-cell.selected[data-cell], .planner-cell.is-selected[data-cell], .planner-cell-clickable.selected[data-cell], .planner-cell-clickable.is-selected[data-cell]"
    ));
  }

  function parseCellFromElement(el) {
    if (!el || !el.dataset || !el.dataset.cell) return null;
    try {
      return JSON.parse(decodeURIComponent(el.dataset.cell));
    } catch (error) {
      try {
        return JSON.parse(el.dataset.cell);
      } catch (innerError) {
        return null;
      }
    }
  }

  function getSelectedCells() {
    return getSelectedCellElements()
      .map(parseCellFromElement)
      .filter(Boolean);
  }

  function getSelectionContext() {
    const cells = getSelectedCells();

    if (cells.length) {
      const first = cells[0];
      const periods = cells
        .map((cell) => Number(cell.period_key || 0))
        .filter(Boolean)
        .sort((a, b) => a - b);

      return {
        cells,
        project_id: Number(first.project_id || 0),
        project_name: first.project_name || "",
        role: norm(first.role || ""),
        week_from: periods.length ? periods[0] : Number(normalizePeriodKey(first) || 0),
        week_to: periods.length ? periods[periods.length - 1] : Number(normalizePeriodKey(first) || 0),
        period_from: periods.length ? periods[0] : Number(normalizePeriodKey(first) || 0),
        period_to: periods.length ? periods[periods.length - 1] : Number(normalizePeriodKey(first) || 0),
        required: Number(first.required || 0),
      };
    }

    if (typeof getSelectionSummary === "function") {
      try {
        const summary = getSelectionSummary();
        if (summary) {
          return {
            cells: [],
            project_id: Number(summary.project_id || 0),
            project_name: summary.project_name || "",
            role: norm(summary.role || ""),
            week_from: Number(summary.period_from || periodKeyFromWeek(summary.week_from || summary.week || 0)),
            week_to: Number(summary.period_to || periodKeyFromWeek(summary.week_to || summary.week_from || summary.week || 0)),
            period_from: Number(summary.period_from || periodKeyFromWeek(summary.week_from || summary.week || 0)),
            period_to: Number(summary.period_to || periodKeyFromWeek(summary.week_to || summary.week_from || summary.week || 0)),
            required: Number(summary.required || 0),
          };
        }
      } catch (error) {
        console.error("getSelectionSummary error", error);
      }
    }

    return null;
  }

  function findDemandInputs() {
    const inputs = Array.from(document.querySelectorAll("input, select, textarea"));

    function byIdOrName(words) {
      return inputs.find((el) => {
        const text = `${el.id || ""} ${el.name || ""} ${el.placeholder || ""}`.toUpperCase();
        return words.every((word) => text.includes(word));
      }) || null;
    }

    const quantity =
      qs("demandQuantity") ||
      qs("selectedDemandQuantity") ||
      qs("requiredInput") ||
      byIdOrName(["DEMAND"]) ||
      byIdOrName(["RICHIESTO"]) ||
      null;

    const weekFrom =
      qs("demandWeekFrom") ||
      qs("weekFromInput") ||
      qs("periodFromInput") ||
      byIdOrName(["WEEK", "FROM"]) ||
      byIdOrName(["DA"]) ||
      null;

    const weekTo =
      qs("demandWeekTo") ||
      qs("weekToInput") ||
      qs("periodToInput") ||
      byIdOrName(["WEEK", "TO"]) ||
      byIdOrName(["A"]) ||
      null;

    return { quantity, weekFrom, weekTo };
  }

  function unlockWeekToAndMultiSelectionDisplay() {
    const context = getSelectionContext();
    const { quantity, weekFrom, weekTo } = findDemandInputs();

    if (weekTo) {
      weekTo.removeAttribute("disabled");
      weekTo.removeAttribute("readonly");
      weekTo.type = "number";
      weekTo.min = "1";
      weekTo.max = "52";
      if (context?.period_to || context?.week_to) writePeriodInput(weekTo, context.period_to || context.week_to);
    }

    if (weekFrom) {
      weekFrom.removeAttribute("disabled");
      weekFrom.removeAttribute("readonly");
      weekFrom.type = "number";
      weekFrom.min = "1";
      weekFrom.max = "52";
      if (context?.period_from || context?.week_from) writePeriodInput(weekFrom, context.period_from || context.week_from);
    }

    if (quantity && context) {
      quantity.removeAttribute("disabled");
      quantity.removeAttribute("readonly");
      quantity.type = "number";
      quantity.step = "0.5";
      quantity.min = "0";

      // Regola corretta: in selezione multipla non mostrare la somma.
      // Mostra il valore della prima cella solo come default modificabile.
      if (context.cells.length > 1) {
        quantity.value = String(Number(context.cells[0]?.required || 0));
        quantity.title = "Selezione multipla: il valore inserito sarà applicato uguale a tutte le celle selezionate.";
      }
    }
  }

  async function saveDemandRangeFromPanel() {
    const context = getSelectionContext();
    const { quantity, weekFrom, weekTo } = findDemandInputs();

    if (!context || !context.project_id || !context.role) {
      alert("Seleziona prima una cella del planner.");
      return false;
    }

    if (!quantity) {
      alert("Campo Richiesto non trovato nel pannello.");
      return false;
    }

    const qty = Number(quantity.value || 0);
    const from = readPeriodInput(weekFrom, context.period_from || context.week_from || 0);
    const to = readPeriodInput(weekTo, context.period_to || context.week_to || from || 0);

    if (!from || !to) {
      alert("Settimana DA/A non valida.");
      return false;
    }

    const result = await fetchJson("/api/demands/upsert-range", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({
        project_id: context.project_id,
        role: context.role,
        quantity: qty,
        week_from: from,
        week_to: to,
      }),
    });

    if (!result.ok) {
      alert(result.error || "Errore salvataggio fabbisogno.");
      return false;
    }

    if (typeof loadAll === "function") {
      await loadAll();
    }

    return true;
  }

  function patchSaveDemandButton() {
    const buttons = Array.from(document.querySelectorAll("button"));
    const saveBtn = buttons.find((btn) => {
      const text = norm(btn.textContent || "");
      return text.includes("SALVA FABBISOGNO") || text === "SALVA" || text.includes("FABBISOGNO");
    });

    if (!saveBtn || saveBtn.dataset.v2DemandRangePatched === "1") return;

    saveBtn.dataset.v2DemandRangePatched = "1";

    saveBtn.addEventListener("click", async (event) => {
      const context = getSelectionContext();
      if (!context) return;

      const inputs = findDemandInputs();

      // Intercettiamo solo quando esistono campi fabbisogno.
      if (!inputs.quantity) return;

      event.preventDefault();
      event.stopPropagation();

      await saveDemandRangeFromPanel();
    }, true);
  }

  async function loadRolesForSelect(select) {
    let roles = [];

    try {
      roles = await fetchJson("/api/roles");
    } catch (error) {
      roles = [];
    }

    if (!Array.isArray(roles) || !roles.length) {
      try {
        const set = new Set();
        if (Array.isArray(resourcesData)) resourcesData.forEach((r) => r.role && set.add(norm(r.role)));
        if (Array.isArray(demandsData)) demandsData.forEach((d) => d.role && set.add(norm(d.role)));
        roles = Array.from(set).sort();
      } catch (error) {
        roles = [];
      }
    }

    select.innerHTML = roles
      .filter(Boolean)
      .map((role) => `<option value="${safeEscape(role)}">${safeEscape(role)}</option>`)
      .join("");
  }

  function ensureRoleRowModal() {
    let modal = qs("roleRowModal");
    if (modal) return modal;

    modal = document.createElement("div");
    modal.id = "roleRowModal";
    modal.className = "role-row-modal";
    modal.hidden = true;
    modal.innerHTML = `
      <div class="role-row-modal-card">
        <div class="role-row-modal-header">
          <strong>Nuova riga mansione</strong>
          <button type="button" id="roleRowCloseBtn" class="btn btn-light">Chiudi</button>
        </div>

        <div class="role-row-field">
          <label>Commessa</label>
          <input id="roleRowProject" type="text" disabled />
        </div>

        <div class="role-row-field">
          <label>Mansione</label>
          <select id="roleRowRole"></select>
        </div>

        <div class="role-row-grid">
          <div class="role-row-field">
            <label>Richiesto</label>
            <input id="roleRowQuantity" type="number" min="0" step="0.5" value="0" />
          </div>
          <div class="role-row-field">
            <label>Da W</label>
            <input id="roleRowWeekFrom" type="number" min="1" max="52" />
          </div>
          <div class="role-row-field">
            <label>A W</label>
            <input id="roleRowWeekTo" type="number" min="1" max="52" />
          </div>
        </div>

        <div class="role-row-help">
          Crea/attiva la riga mansione e imposta lo stesso fabbisogno su tutte le settimane del range.
        </div>

        <div class="role-row-actions">
          <button type="button" id="roleRowSaveBtn" class="btn btn-primary">Crea riga e salva fabbisogno</button>
        </div>
      </div>
    `;

    document.body.appendChild(modal);

    qs("roleRowCloseBtn").addEventListener("click", () => {
      modal.hidden = true;
    });

    qs("roleRowSaveBtn").addEventListener("click", saveRoleRowModal);

    return modal;
  }

  async function openRoleRowModal() {
    const context = getSelectionContext();
    if (!context || !context.project_id) {
      alert("Seleziona prima una cella della commessa.");
      return;
    }

    const modal = ensureRoleRowModal();
    const roleSelect = qs("roleRowRole");

    qs("roleRowProject").value = context.project_name || `ID ${context.project_id}`;
    qs("roleRowQuantity").value = "0";
    qs("roleRowWeekFrom").value = String(context.period_from || periodKeyFromWeek(context.week_from || 1));
    qs("roleRowWeekTo").value = String(context.period_to || periodKeyFromWeek(context.week_to || context.week_from || 1));

    await loadRolesForSelect(roleSelect);

    if (context.role && Array.from(roleSelect.options).some((opt) => opt.value === context.role)) {
      roleSelect.value = context.role;
    }

    modal.dataset.projectId = String(context.project_id);
    modal.hidden = false;
  }

  async function saveRoleRowModal() {
    const modal = qs("roleRowModal");
    const projectId = Number(modal?.dataset.projectId || 0);
    const role = norm(qs("roleRowRole")?.value || "");
    const quantity = Number(qs("roleRowQuantity")?.value || 0);
    const weekFrom = Number(qs("roleRowWeekFrom")?.value || 0);
    const weekTo = Number(qs("roleRowWeekTo")?.value || weekFrom || 0);

    if (!projectId || !role || !weekFrom || !weekTo) {
      alert("Compila mansione, richiesto, settimana da/a.");
      return;
    }

    if (quantity <= 0) {
      alert("Per creare una nuova riga mansione, il campo Richiesto deve essere maggiore di 0.");
      return;
    }

    const result = await fetchJson("/api/project-role-row", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({
        project_id: projectId,
        role,
        quantity,
        week_from: weekFrom,
        week_to: weekTo,
      }),
    });

    if (!result.ok) {
      alert(result.error || "Errore creazione riga mansione.");
      return;
    }

    modal.hidden = true;

    if (typeof loadAll === "function") {
      await loadAll();
    }
  }

  function ensureRoleRowButton() {
    if (qs("openRoleRowModalBtn")) return;

    const target =
      document.querySelector(".side-panel") ||
      document.querySelector(".right-panel") ||
      document.querySelector("aside") ||
      document.body;

    const btn = document.createElement("button");
    btn.id = "openRoleRowModalBtn";
    btn.type = "button";
    btn.className = "btn btn-light open-role-row-modal-btn";
    btn.textContent = "+ Riga mansione";
    btn.addEventListener("click", openRoleRowModal);

    target.prepend(btn);
  }

  function bindPlannerDemandRolePackage() {
    ensureRoleRowButton();
    patchSaveDemandButton();
    unlockWeekToAndMultiSelectionDisplay();
  }

  document.addEventListener("click", () => {
    setTimeout(bindPlannerDemandRolePackage, 60);
    setTimeout(unlockWeekToAndMultiSelectionDisplay, 120);
  });

  document.addEventListener("selectionchange", () => {
    setTimeout(unlockWeekToAndMultiSelectionDisplay, 80);
  });

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", bindPlannerDemandRolePackage);
  } else {
    bindPlannerDemandRolePackage();
  }

  setTimeout(bindPlannerDemandRolePackage, 500);
  setTimeout(bindPlannerDemandRolePackage, 1500);
})();
// === PLANNER DEMAND ROLE PACKAGE END ===

// === FIX ROLE ROW REFRESH SORT FILTERS START ===
(function fixRoleRowRefreshSortFilters() {
  function norm(value) {
    if (typeof normalizeRole === "function") return normalizeRole(value);
    return String(value || "").trim().toUpperCase();
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

  function getProjectCode(name) {
    const text = String(name || "").trim();
    const match = text.match(/^([0-9]+[_-]?[0-9]*)/);
    return match ? match[1] : text;
  }

  function getRoleOrder(role) {
    const order = [
      "ASPP",
      "CAPO CANTIERE",
      "CAPO SQUADRA",
      "QUALITY CONTROL / WELDING INSPECTOR",
      "TUBISTA",
      "SALDATORE TIG-ELETTRODO",
      "SALDATORE FILO",
      "CARPENTIERE",
      "MONTATORE",
      "MECCANICO",
      "MECCANICO SERVICE",
      "MANDRINATORE",
      "ELETTRICISTA",
      "COIBENTATORE",
      "PONTEGGIATORE",
      "SOLLEVAMENTI",
      "PWHT",
      "VERNICIATORE",
      "MAGAZZINIERE",
      "GENERICO",
      "AUTISTA",
    ];
    const idx = order.indexOf(norm(role));
    return idx >= 0 ? idx : 999;
  }

  function patchRowsSortInPlace() {
    try {
      if (!Array.isArray(window.plannerMatrixRows) && typeof plannerMatrixRows === "undefined") {
        return;
      }
    } catch (error) {
      return;
    }

    let rows;
    try {
      rows = Array.isArray(plannerMatrixRows) ? plannerMatrixRows : window.plannerMatrixRows;
    } catch (error) {
      rows = window.plannerMatrixRows;
    }

    if (!Array.isArray(rows)) return;

    rows.sort((a, b) => {
      const projectA = String(a.project_name || a.projectName || "");
      const projectB = String(b.project_name || b.projectName || "");

      const codeCmp = getProjectCode(projectA).localeCompare(getProjectCode(projectB), "it", {
        numeric: true,
        sensitivity: "base",
      });
      if (codeCmp !== 0) return codeCmp;

      const projectCmp = projectA.localeCompare(projectB, "it", {
        numeric: true,
        sensitivity: "base",
      });
      if (projectCmp !== 0) return projectCmp;

      const roleOrderCmp = getRoleOrder(a.role) - getRoleOrder(b.role);
      if (roleOrderCmp !== 0) return roleOrderCmp;

      return norm(a.role).localeCompare(norm(b.role), "it", {
        numeric: true,
        sensitivity: "base",
      });
    });
  }

  function getRowsFromDomOrData() {
    const rows = [];

    try {
      if (Array.isArray(plannerMatrixRows)) {
        plannerMatrixRows.forEach((row) => rows.push(row));
      }
    } catch (error) {}

    try {
      if (Array.isArray(window.plannerMatrixRows)) {
        window.plannerMatrixRows.forEach((row) => rows.push(row));
      }
    } catch (error) {}

    document.querySelectorAll("[data-cell]").forEach((el) => {
      try {
        const cell = JSON.parse(decodeURIComponent(el.dataset.cell));
        if (cell?.project_id && cell?.project_name) {
          rows.push({
            project_id: cell.project_id,
            project_name: cell.project_name,
            role: cell.role,
          });
        }
      } catch (error) {
        try {
          const cell = JSON.parse(el.dataset.cell);
          if (cell?.project_id && cell?.project_name) {
            rows.push({
              project_id: cell.project_id,
              project_name: cell.project_name,
              role: cell.role,
            });
          }
        } catch (innerError) {}
      }
    });

    try {
      if (Array.isArray(projectsData)) {
        projectsData.forEach((project) => {
          rows.push({
            project_id: project.id,
            project_name: project.name,
            role: "",
          });
        });
      }
    } catch (error) {}

    try {
      if (Array.isArray(demandsData)) {
        demandsData.forEach((demand) => {
          const project = Array.isArray(projectsData)
            ? projectsData.find((p) => Number(p.id) === Number(demand.project_id))
            : null;
          rows.push({
            project_id: demand.project_id,
            project_name: project?.name || "",
            role: demand.role,
          });
        });
      }
    } catch (error) {}

    return rows;
  }

  function findProjectFilter() {
    const selects = Array.from(document.querySelectorAll("select"));
    return selects.find((select) => {
      const text = `${select.id || ""} ${select.name || ""} ${select.closest("label")?.textContent || ""}`.toUpperCase();
      const firstOption = norm(select.options?.[0]?.textContent || "");
      return (
        text.includes("PROJECT") ||
        text.includes("COMMESS") ||
        firstOption.includes("TUTTE LE COMMESSE")
      );
    }) || null;
  }

  function findRoleFilter() {
    const selects = Array.from(document.querySelectorAll("select"));
    return selects.find((select) => {
      const text = `${select.id || ""} ${select.name || ""} ${select.closest("label")?.textContent || ""}`.toUpperCase();
      const firstOption = norm(select.options?.[0]?.textContent || "");
      return (
        text.includes("ROLE") ||
        text.includes("MANSION") ||
        firstOption.includes("TUTTE LE MANSIONI")
      );
    }) || null;
  }

  function preserveAndSetOptions(select, firstLabel, options, getValue, getLabel) {
    if (!select) return;

    const previous = select.value;

    const seen = new Set();
    const cleanOptions = [];

    options.forEach((item) => {
      const value = String(getValue(item) ?? "").trim();
      const label = String(getLabel(item) ?? "").trim();
      if (!value || seen.has(value)) return;
      seen.add(value);
      cleanOptions.push({ value, label });
    });

    cleanOptions.sort((a, b) => a.label.localeCompare(b.label, "it", {
      numeric: true,
      sensitivity: "base",
    }));

    select.innerHTML =
      `<option value="">${safeEscape(firstLabel)}</option>` +
      cleanOptions
        .map((item) => `<option value="${safeEscape(item.value)}">${safeEscape(item.label)}</option>`)
        .join("");

    if (previous && Array.from(select.options).some((opt) => opt.value === previous)) {
      select.value = previous;
    }
  }

  function populatePlannerFilters() {
    const rows = getRowsFromDomOrData();

    const projectMap = new Map();
    const roleSet = new Set();

    rows.forEach((row) => {
      const projectId = Number(row.project_id || 0);
      const projectName = String(row.project_name || "").trim();
      const role = norm(row.role || "");

      if (projectId && projectName) {
        projectMap.set(String(projectId), {
          id: String(projectId),
          name: projectName,
        });
      }

      if (role) roleSet.add(role);
    });

    const projectFilter = findProjectFilter();
    const roleFilter = findRoleFilter();

    preserveAndSetOptions(
      projectFilter,
      "Tutte le commesse",
      Array.from(projectMap.values()),
      (p) => p.id,
      (p) => p.name
    );

    preserveAndSetOptions(
      roleFilter,
      "Tutte le mansioni",
      Array.from(roleSet).map((role) => ({ role })),
      (r) => r.role,
      (r) => r.role
    );
  }

  async function hardPlannerRefresh() {
    if (typeof loadAll === "function") {
      await loadAll();
    }

    patchRowsSortInPlace();

    if (typeof renderPlanner === "function") {
      renderPlanner();
    }

    setTimeout(() => {
      patchRowsSortInPlace();
      populatePlannerFilters();
      if (typeof renderPlanner === "function") {
        renderPlanner();
      }
    }, 80);

    setTimeout(() => {
      populatePlannerFilters();
    }, 300);
  }

  function patchRenderPlannerSortAndFilters() {
    if (typeof renderPlanner !== "function") return;
    if (window.__renderPlannerSortFiltersPatched) return;

    window.__renderPlannerSortFiltersPatched = true;

    const originalRenderPlanner = renderPlanner;

    renderPlanner = function patchedRenderPlannerSortFilters() {
      patchRowsSortInPlace();
      const result = originalRenderPlanner.apply(this, arguments);
      setTimeout(populatePlannerFilters, 30);
      return result;
    };

    try {
      window.renderPlanner = renderPlanner;
    } catch (error) {}
  }

  function patchSaveRoleRowModalRefresh() {
    if (typeof window.saveRoleRowModal !== "function") {
      return;
    }

    if (window.__saveRoleRowModalRefreshPatched) {
      return;
    }

    window.__saveRoleRowModalRefreshPatched = true;

    const originalSaveRoleRowModal = window.saveRoleRowModal;

    window.saveRoleRowModal = async function patchedSaveRoleRowModalRefresh() {
      const result = await originalSaveRoleRowModal.apply(this, arguments);
      await hardPlannerRefresh();
      return result;
    };
  }

  function patchRoleRowSaveButtonFallback() {
    const btn = document.getElementById("roleRowSaveBtn");
    if (!btn || btn.dataset.refreshFallbackBound === "1") return;

    btn.dataset.refreshFallbackBound = "1";

    btn.addEventListener("click", () => {
      setTimeout(hardPlannerRefresh, 300);
      setTimeout(hardPlannerRefresh, 900);
    });
  }

  function bindFixes() {
    patchRenderPlannerSortAndFilters();
    patchSaveRoleRowModalRefresh();
    patchRoleRowSaveButtonFallback();
    patchRowsSortInPlace();
    populatePlannerFilters();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", bindFixes);
  } else {
    bindFixes();
  }

  document.addEventListener("click", () => {
    setTimeout(bindFixes, 80);
  });

  setTimeout(bindFixes, 500);
  setTimeout(bindFixes, 1500);

  window.refreshPlannerRowsAndFilters = hardPlannerRefresh;
})();
// === FIX ROLE ROW REFRESH SORT FILTERS END ===

// === OLD WORKFLOW PLANNER PATCH START ===
(function oldWorkflowPlannerPatch() {
  function optionEscape(value) {
    return escapeHtml(value);
  }

  function populatePlannerFiltersFromData() {
    const projectFilter = oldWorkflowFindPlannerProjectFilter();
    const roleFilter = oldWorkflowFindPlannerRoleFilter();

    if (!projectFilter && !roleFilter) {
      return;
    }

    const currentProject = projectFilter ? String(projectFilter.value || "") : "";
    const currentRole = roleFilter ? normalizeRole(roleFilter.value || "") : "";

    const projects = new Map();
    const roles = new Set();

    for (const demand of demandsData || []) {
      const project = projectsData.find((item) => Number(item.id) === Number(demand.project_id));
      if (!project) continue;

      projects.set(String(project.id), project.name);

      if (demand.role) {
        roles.add(normalizeRole(demand.role));
      }
    }

    for (const allocation of allocationsData || []) {
      const project = projectsData.find((item) => Number(item.id) === Number(allocation.project_id));
      if (!project) continue;

      const resource = resourcesData.find((item) => Number(item.id) === Number(allocation.resource_id));

      projects.set(String(project.id), project.name);

      if (allocation.role || resource?.role) {
        roles.add(normalizeRole(allocation.role || resource.role));
      }
    }

    if (projectFilter) {
      const orderedProjects = Array.from(projects.entries()).sort((a, b) => {
        return oldWorkflowProjectSortKey(a[1]).localeCompare(oldWorkflowProjectSortKey(b[1]), "it", {
          numeric: true,
          sensitivity: "base",
        });
      });

      projectFilter.innerHTML =
        `<option value="">Tutte le commesse</option>` +
        orderedProjects
          .map(([id, name]) => `<option value="${optionEscape(id)}">${optionEscape(name)}</option>`)
          .join("");

      if (currentProject && Array.from(projectFilter.options).some((option) => option.value === currentProject)) {
        projectFilter.value = currentProject;
      }
    }

    if (roleFilter) {
      const orderedRoles = Array.from(roles).sort((a, b) => {
        return oldWorkflowRoleSortKey(a).localeCompare(oldWorkflowRoleSortKey(b), "it", {
          numeric: true,
          sensitivity: "base",
        });
      });

      roleFilter.innerHTML =
        `<option value="">Tutte le mansioni</option>` +
        orderedRoles
          .map((role) => `<option value="${optionEscape(role)}">${optionEscape(role)}</option>`)
          .join("");

      if (currentRole && Array.from(roleFilter.options).some((option) => normalizeRole(option.value) === currentRole)) {
        roleFilter.value = currentRole;
      }
    }
  }

  function bindPlannerFilterEvents() {
    for (const filter of [oldWorkflowFindPlannerProjectFilter(), oldWorkflowFindPlannerRoleFilter()]) {
      if (!filter || filter.dataset.oldWorkflowFilterBound === "1") {
        continue;
      }

      filter.dataset.oldWorkflowFilterBound = "1";
      filter.addEventListener("change", () => {
        if (typeof renderPlanner === "function") {
          renderPlanner();
        }
      });
    }
  }

  async function refreshPlannerAfterRoleSave() {
    if (typeof loadAll === "function") {
      await loadAll();
    }

    if (typeof renderPlanner === "function") {
      renderPlanner();
    }

    setTimeout(populatePlannerFiltersFromData, 50);
    setTimeout(bindPlannerFilterEvents, 60);
  }

  function bindRoleRowSaveRefresh() {
    const btn = document.getElementById("roleRowSaveBtn");

    if (!btn || btn.dataset.oldWorkflowRoleSaveRefreshBound === "1") {
      return;
    }

    btn.dataset.oldWorkflowRoleSaveRefreshBound = "1";

    btn.addEventListener("click", () => {
      setTimeout(refreshPlannerAfterRoleSave, 350);
      setTimeout(refreshPlannerAfterRoleSave, 1000);
    });
  }

  if (typeof renderPlanner === "function" && !window.__oldWorkflowRenderPlannerPatch) {
    window.__oldWorkflowRenderPlannerPatch = true;

    const originalRenderPlanner = renderPlanner;

    renderPlanner = function patchedRenderPlannerOldWorkflow() {
      const result = originalRenderPlanner.apply(this, arguments);

      setTimeout(populatePlannerFiltersFromData, 30);
      setTimeout(bindPlannerFilterEvents, 40);
      setTimeout(bindRoleRowSaveRefresh, 50);

      return result;
    };

    try {
      window.renderPlanner = renderPlanner;
    } catch (error) {}
  }

  window.oldWorkflowRefreshPlanner = refreshPlannerAfterRoleSave;

  setTimeout(populatePlannerFiltersFromData, 500);
  setTimeout(bindPlannerFilterEvents, 550);
  setTimeout(bindRoleRowSaveRefresh, 600);
})();
// === OLD WORKFLOW PLANNER PATCH END ===

// === ROLE ROW AUTO REFRESH AFTER SAVE START ===
(function roleRowAutoRefreshAfterSave() {
  if (window.__roleRowAutoRefreshAfterSaveInstalled) {
    return;
  }

  window.__roleRowAutoRefreshAfterSaveInstalled = true;

  const originalFetch = window.fetch.bind(window);

  async function refreshPlannerAfterRoleRowSave() {
    try {
      if (typeof loadAll === "function") {
        await loadAll();
      }

      if (typeof renderPlanner === "function") {
        renderPlanner();
      }

      if (typeof populatePlannerFiltersFromData === "function") {
        populatePlannerFiltersFromData();
      }

      if (typeof oldWorkflowRefreshPlanner === "function") {
        // ulteriore sicurezza: usa il refresh installato dallo step 1
        setTimeout(oldWorkflowRefreshPlanner, 150);
      }
    } catch (error) {
      console.error("Errore refresh dopo nuova riga mansione", error);
    }
  }

  window.fetch = async function patchedFetch(input, init) {
    const url = typeof input === "string" ? input : String(input?.url || "");
    const method = String(init?.method || input?.method || "GET").toUpperCase();

    const isProjectRoleRowSave =
      method === "POST" &&
      url.includes("/api/project-role-row");

    const response = await originalFetch(input, init);

    if (isProjectRoleRowSave && response.ok) {
      // Il backend ha salvato. Ora ricarico i dati e ridisegno il planner.
      setTimeout(refreshPlannerAfterRoleRowSave, 150);
      setTimeout(refreshPlannerAfterRoleRowSave, 700);
      setTimeout(refreshPlannerAfterRoleRowSave, 1500);
    }

    return response;
  };
})();
// === ROLE ROW AUTO REFRESH AFTER SAVE END ===

// === ROLE ROW DIRECT SAVE AND REFRESH START ===
(function roleRowDirectSaveAndRefresh() {
  if (window.__roleRowDirectSaveAndRefreshInstalled) {
    return;
  }
  window.__roleRowDirectSaveAndRefreshInstalled = true;

  function n(value) {
    return Number(value || 0);
  }

  function norm(value) {
    if (typeof normalizeRole === "function") return normalizeRole(value);
    return String(value || "").trim().toUpperCase();
  }

  async function reloadPlannerCoreDataDirect() {
    const stamp = Date.now();

    try {
      resourcesData = await fetchJson(`/api/resources?_=${stamp}`);
    } catch (error) {
      console.warn("reload resources skipped", error);
    }

    try {
      projectsData = await fetchJson(`/api/projects?_=${stamp}`);
    } catch (error) {
      console.warn("reload projects skipped", error);
    }

    try {
      demandsData = await fetchJson(`/api/demands?_=${stamp}`);
    } catch (error) {
      console.warn("reload demands skipped", error);
    }

    try {
      allocationsData = await fetchJson(`/api/allocations?_=${stamp}`);
    } catch (error) {
      console.warn("reload allocations skipped", error);
    }

    try {
      allocationHistoryData = await fetchJson(`/api/allocation-history?_=${stamp}`);
    } catch (error) {
      // non bloccante
    }

    try {
      demandHistoryData = await fetchJson(`/api/demand-history?_=${stamp}`);
    } catch (error) {
      // non bloccante
    }

    if (typeof renderPlanner === "function") {
      renderPlanner();
    }

    if (typeof oldWorkflowRefreshPlanner === "function") {
      setTimeout(oldWorkflowRefreshPlanner, 100);
    }

    if (typeof scrollToCurrentWeek === "function") {
      setTimeout(scrollToCurrentWeek, 150);
    }
  }

  async function saveRoleRowDirect(event) {
    const btn = event.target?.closest?.("#roleRowSaveBtn");
    if (!btn) return;

    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();

    const modal = document.getElementById("roleRowModal");
    const projectId = n(modal?.dataset.projectId);
    const role = norm(document.getElementById("roleRowRole")?.value || "");
    const quantity = n(document.getElementById("roleRowQuantity")?.value);
    const weekFrom = n(document.getElementById("roleRowWeekFrom")?.value);
    const weekTo = n(document.getElementById("roleRowWeekTo")?.value || weekFrom);

    if (!projectId || !role || !weekFrom || !weekTo) {
      alert("Compila mansione, richiesto, settimana da/a.");
      return;
    }

    if (quantity <= 0) {
      alert("Per creare una nuova riga mansione, il campo Richiesto deve essere maggiore di 0.");
      return;
    }

    btn.disabled = true;
    const oldText = btn.textContent;
    btn.textContent = "Salvataggio...";

    try {
      const result = await fetchJson("/api/project-role-row", {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify({
          project_id: projectId,
          role,
          quantity,
          week_from: weekFrom,
          week_to: weekTo,
        }),
      });

      if (!result.ok) {
        alert(result.error || "Errore creazione riga mansione.");
        return;
      }

      if (modal) {
        modal.hidden = true;
      }

      await reloadPlannerCoreDataDirect();

      // Sicurezza: un secondo render dopo che il DOM ha chiuso la modale.
      setTimeout(reloadPlannerCoreDataDirect, 500);
    } catch (error) {
      console.error(error);
      alert(error.message || "Errore salvataggio nuova riga mansione.");
    } finally {
      btn.disabled = false;
      btn.textContent = oldText;
    }
  }

  document.addEventListener("click", saveRoleRowDirect, true);
})();
// === ROLE ROW DIRECT SAVE AND REFRESH END ===

// === OLD WORKFLOW COMMESSE/FABBISOGNI START ===
(function oldWorkflowCommesseFabbisogni() {
  if (window.__oldWorkflowCommesseInstalled) {
    return;
  }

  window.__oldWorkflowCommesseInstalled = true;

  const commesseState = {
    selectedProjectId: "",
    matrix: null,
    extraRoles: new Set(),
    dirty: false,
  };

  function norm(value) {
    return typeof normalizeRole === "function"
      ? normalizeRole(value)
      : String(value || "").trim().toUpperCase();
  }

  function html(value) {
    return typeof escapeHtml === "function"
      ? escapeHtml(value)
      : String(value ?? "")
          .replaceAll("&", "&amp;")
          .replaceAll("<", "&lt;")
          .replaceAll(">", "&gt;")
          .replaceAll('"', "&quot;")
          .replaceAll("'", "&#039;");
  }

  function roleSortKey(role) {
    if (typeof oldWorkflowRoleSortKey === "function") {
      return oldWorkflowRoleSortKey(role);
    }

    return norm(role);
  }

  function projectSortKey(name) {
    if (typeof oldWorkflowProjectSortKey === "function") {
      return oldWorkflowProjectSortKey(name);
    }

    return String(name || "").toUpperCase();
  }

  function numberValue(value) {
    const raw = String(value ?? "").replace(",", ".").trim();
    if (!raw) return 0;
    const parsed = Number(raw);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function formatQty(value) {
    const n = Number(value || 0);
    if (!n) return "";
    if (Number.isInteger(n)) return String(n);
    return String(Math.round(n * 10) / 10).replace(".", ",");
  }

  function ensureCommesseSheet() {
    let sheet = document.getElementById("oldWorkflowCommesseSheet");
    if (sheet) return sheet;

    sheet = document.createElement("section");
    sheet.id = "oldWorkflowCommesseSheet";
    sheet.className = "old-commesse-sheet card";
    sheet.hidden = true;

    sheet.innerHTML = `
      <div class="old-commesse-header">
        <div>
          <div class="side-title">Commesse / Fabbisogni</div>
          <div class="old-commesse-subtitle">
            Macro gestione fabbisogni: commessa → mansioni → settimane.
          </div>
        </div>

        <div class="old-commesse-actions">
          <input id="oldCommesseSearch" type="text" placeholder="Cerca commessa..." />
          <select id="oldCommesseProjectSelect"></select>
          <button class="btn btn-light" id="oldCommesseReloadBtn" type="button">Ricarica</button>
          <button class="btn btn-primary" id="oldCommesseSaveBtn" type="button">Salva modifiche</button>
          <button class="btn btn-light" id="oldCommesseCloseBtn" type="button">Chiudi</button>
        </div>
      </div>

      <div class="old-commesse-add-role">
        <div>
          <strong>Nuova mansione su commessa</strong>
          <span>Seleziona una mansione, valorizza le settimane, poi salva.</span>
        </div>
        <div class="old-commesse-add-role-controls">
          <select id="oldCommesseRoleSelect"></select>
          <button class="btn btn-light" id="oldCommesseAddRoleBtn" type="button">Aggiungi mansione</button>
        </div>
      </div>

      <div class="old-commesse-info" id="oldCommesseInfo">
        Seleziona una commessa.
      </div>

      <div class="old-commesse-table-wrap">
        <table class="old-commesse-table">
          <thead id="oldCommesseHead"></thead>
          <tbody id="oldCommesseBody"></tbody>
        </table>
      </div>
    `;

    document.querySelector(".planner-layout")?.appendChild(sheet);
    return sheet;
  }

  function showOnlyCommesse() {
    ensureCommesseSheet();

    const commesse = document.getElementById("oldWorkflowCommesseSheet");
    const plannerMain = document.querySelector(".planner-main");
    const plannerSide = document.querySelector(".planner-side");
    const plannerBottom = document.querySelector(".planner-bottom");
    const resourcesSheet = document.getElementById("resourcesSheet");
    const ganttSheet = document.getElementById("oldWorkflowGanttSheet");

    if (plannerMain) plannerMain.hidden = true;
    if (plannerSide) plannerSide.hidden = true;
    if (plannerBottom) plannerBottom.hidden = true;
    if (resourcesSheet) resourcesSheet.hidden = true;
    if (ganttSheet) ganttSheet.hidden = true;
    if (commesse) commesse.hidden = false;

    renderCommesseSheet();
  }

  function showPlannerFromCommesse() {
    const commesse = document.getElementById("oldWorkflowCommesseSheet");
    const plannerMain = document.querySelector(".planner-main");
    const plannerSide = document.querySelector(".planner-side");
    const plannerBottom = document.querySelector(".planner-bottom");

    if (commesse) commesse.hidden = true;
    if (plannerMain) plannerMain.hidden = false;
    if (plannerSide) plannerSide.hidden = false;
    if (plannerBottom) plannerBottom.hidden = false;

    if (typeof renderPlanner === "function") {
      renderPlanner();
    }
  }

  function bindTopButtons() {
    document.querySelectorAll(".topbar-actions button").forEach((button) => {
      const text = norm(button.textContent || "");

      if ((text === "PROGETTI" || text.includes("COMMESSE")) && button.dataset.oldCommesseBound !== "1") {
        button.dataset.oldCommesseBound = "1";
        button.addEventListener("click", showOnlyCommesse);
      }

      if (text === "PLANNER" && button.dataset.oldCommessePlannerBound !== "1") {
        button.dataset.oldCommessePlannerBound = "1";
        button.addEventListener("click", showPlannerFromCommesse);
      }
    });
  }

  function getVisibleProjects() {
    const search = norm(document.getElementById("oldCommesseSearch")?.value || "");

    return (projectsData || [])
      .filter((project) => {
        if (!project) return false;
        if (typeof isWorkshopChildProject === "function" && isWorkshopChildProject(project)) return false;

        const haystack = norm(`${project.name || ""} ${project.status || ""} ${project.note || ""}`);
        if (search && !haystack.includes(search)) return false;

        return true;
      })
      .sort((a, b) => projectSortKey(a.name).localeCompare(projectSortKey(b.name), "it", {
        numeric: true,
        sensitivity: "base",
      }));
  }

  function getAllRoles() {
    const roles = new Set();

    if (Array.isArray(window.OLD_WORKFLOW_ROLE_ORDER)) {
      window.OLD_WORKFLOW_ROLE_ORDER.forEach((role) => roles.add(norm(role)));
    }

    if (typeof OLD_WORKFLOW_ROLE_ORDER !== "undefined" && Array.isArray(OLD_WORKFLOW_ROLE_ORDER)) {
      OLD_WORKFLOW_ROLE_ORDER.forEach((role) => roles.add(norm(role)));
    }

    for (const demand of demandsData || []) {
      if (demand.role) roles.add(norm(demand.role));
    }

    for (const resource of resourcesData || []) {
      if (resource.role) roles.add(norm(resource.role));
    }

    return Array.from(roles).sort((a, b) => roleSortKey(a).localeCompare(roleSortKey(b), "it", {
      numeric: true,
      sensitivity: "base",
    }));
  }

  function populateProjectSelect() {
    const select = document.getElementById("oldCommesseProjectSelect");
    if (!select) return;

    const previous = commesseState.selectedProjectId || select.value || "";
    const projects = getVisibleProjects();

    select.innerHTML =
      `<option value="">Seleziona commessa</option>` +
      projects.map((project) => {
        return `<option value="${Number(project.id)}">${html(project.name || "")}</option>`;
      }).join("");

    if (previous && Array.from(select.options).some((option) => option.value === String(previous))) {
      select.value = String(previous);
      commesseState.selectedProjectId = String(previous);
    } else if (!commesseState.selectedProjectId && projects.length) {
      select.value = String(projects[0].id);
      commesseState.selectedProjectId = String(projects[0].id);
    }
  }

  function populateRoleSelect() {
    const select = document.getElementById("oldCommesseRoleSelect");
    if (!select) return;

    const existingRoles = new Set(getMatrixRoles().map((row) => norm(row.role)));

    const roles = getAllRoles().filter((role) => !existingRoles.has(norm(role)));

    select.innerHTML =
      `<option value="">Seleziona mansione</option>` +
      roles.map((role) => `<option value="${html(role)}">${html(role)}</option>`).join("");
  }

  function getMatrixRoles() {
    const roles = [];

    if (commesseState.matrix?.roles && Array.isArray(commesseState.matrix.roles)) {
      for (const row of commesseState.matrix.roles) {
        const role = norm(row.role);
        if (!role) continue;

        const total = Number(row.total || 0);
        const hasAnyWeek = Object.values(row.weeks || {}).some((cell) => Number(cell.quantity || 0) > 0);

        if (total > 0 || hasAnyWeek || commesseState.extraRoles.has(role)) {
          roles.push({
            role,
            weeks: row.weeks || {},
            total,
            isExtra: commesseState.extraRoles.has(role),
          });
        }
      }
    }

    for (const role of commesseState.extraRoles) {
      if (!roles.some((row) => norm(row.role) === norm(role))) {
        roles.push({
          role,
          weeks: {},
          total: 0,
          isExtra: true,
        });
      }
    }

    roles.sort((a, b) => roleSortKey(a.role).localeCompare(roleSortKey(b.role), "it", {
      numeric: true,
      sensitivity: "base",
    }));

    return roles;
  }

  async function loadProjectMatrix(projectId) {
    if (!projectId) {
      commesseState.matrix = null;
      return;
    }

    const result = await fetchJson(`/api/project-demand-matrix/${projectId}?_=${Date.now()}`);

    if (!result.ok) {
      alert(result.error || "Errore caricamento fabbisogni commessa.");
      commesseState.matrix = null;
      return;
    }

    commesseState.matrix = result;
    commesseState.extraRoles = new Set();
    commesseState.dirty = false;
  }

  function renderCommesseHead() {
    const head = document.getElementById("oldCommesseHead");
    if (!head) return;

    const weeks = PERIODS.map((period) => {
      const current = Number(period.periodKey) === Number(CURRENT_PERIOD_KEY) ? " current-week" : "";
      return `<th class="old-commesse-week${current}" data-period-key="${Number(period.periodKey)}">${html(period.label)}</th>`;
    }).join("");

    head.innerHTML = `
      <tr>
        <th class="old-commesse-role-head">Mansione</th>
        <th class="old-commesse-total-head">Totale</th>
        ${weeks}
      </tr>
    `;
  }

  function renderCommesseBody() {
    const body = document.getElementById("oldCommesseBody");
    const info = document.getElementById("oldCommesseInfo");
    if (!body) return;

    if (!commesseState.selectedProjectId || !commesseState.matrix) {
      body.innerHTML = `<tr><td class="old-commesse-empty" colspan="${PERIODS.length + 2}">Seleziona una commessa.</td></tr>`;
      if (info) info.textContent = "Seleziona una commessa.";
      return;
    }

    const project = commesseState.matrix.project;
    const rows = getMatrixRoles();

    if (info) {
      info.innerHTML = `
        <strong>${html(project?.name || "")}</strong>
        <span>ID ${html(project?.id || "")}</span>
        <span>${rows.length} mansioni valorizzate</span>
      `;
    }

    if (!rows.length) {
      body.innerHTML = `<tr><td class="old-commesse-empty" colspan="${PERIODS.length + 2}">Nessuna mansione valorizzata. Aggiungi una mansione dal menu sopra.</td></tr>`;
      return;
    }

    body.innerHTML = rows.map((row) => {
      let total = 0;

      const cells = PERIODS.map((period) => {
        const cell = row.weeks[String(period.periodKey)] || {};
        const qty = Number(cell.quantity || 0);
        total += qty;

        return `
          <td class="old-commesse-cell">
            <input
              type="number"
              min="0"
              step="0.5"
              value="${html(formatQty(qty))}"
              data-role="${html(row.role)}"
              data-period-key="${Number(period.periodKey)}"
            />
          </td>
        `;
      }).join("");

      return `
        <tr data-role-row="${html(row.role)}">
          <td class="old-commesse-role">${html(row.role)}${row.isExtra ? ' <span class="old-commesse-new-role">nuova</span>' : ""}</td>
          <td class="old-commesse-total" data-total-role="${html(row.role)}">${html(formatQty(total))}</td>
          ${cells}
        </tr>
      `;
    }).join("");

    body.querySelectorAll("input[data-role][data-period-key]").forEach((input) => {
      input.addEventListener("input", () => {
        commesseState.dirty = true;
        updateRoleTotal(input.dataset.role);
      });

      input.addEventListener("keydown", handleCommesseKeyboard);
    });
  }

  function updateRoleTotal(role) {
    const normalizedRole = norm(role);
    const inputs = Array.from(document.querySelectorAll(`input[data-role="${CSS.escape(normalizedRole)}"]`));
    const total = inputs.reduce((sum, input) => sum + numberValue(input.value), 0);
    const target = document.querySelector(`[data-total-role="${CSS.escape(normalizedRole)}"]`);
    if (target) target.textContent = formatQty(total);
  }

  function handleCommesseKeyboard(event) {
    const input = event.target;
    if (!input.matches("input[data-role][data-period-key]")) return;

    const td = input.closest("td");
    const tr = input.closest("tr");
    if (!td || !tr) return;

    const rowIndex = Array.from(tr.parentElement.children).indexOf(tr);
    const cellIndex = Array.from(tr.children).indexOf(td);

    let next = null;

    if (event.key === "ArrowRight") {
      next = tr.children[cellIndex + 1]?.querySelector("input");
    } else if (event.key === "ArrowLeft") {
      next = tr.children[cellIndex - 1]?.querySelector("input");
    } else if (event.key === "ArrowDown") {
      const nextRow = tr.parentElement.children[rowIndex + 1];
      next = nextRow?.children[cellIndex]?.querySelector("input");
    } else if (event.key === "ArrowUp") {
      const prevRow = tr.parentElement.children[rowIndex - 1];
      next = prevRow?.children[cellIndex]?.querySelector("input");
    }

    if (next) {
      event.preventDefault();
      next.focus();
      next.select();
    }
  }

  async function renderCommesseSheet() {
    ensureCommesseSheet();
    populateProjectSelect();
    populateRoleSelect();
    renderCommesseHead();

    if (commesseState.selectedProjectId) {
      await loadProjectMatrix(commesseState.selectedProjectId);
      populateRoleSelect();
    }

    renderCommesseBody();
    bindCommesseControls();
  }

  function bindCommesseControls() {
    const projectSelect = document.getElementById("oldCommesseProjectSelect");
    const search = document.getElementById("oldCommesseSearch");
    const reload = document.getElementById("oldCommesseReloadBtn");
    const close = document.getElementById("oldCommesseCloseBtn");
    const save = document.getElementById("oldCommesseSaveBtn");
    const addRole = document.getElementById("oldCommesseAddRoleBtn");

    if (projectSelect && projectSelect.dataset.bound !== "1") {
      projectSelect.dataset.bound = "1";
      projectSelect.addEventListener("change", async () => {
        if (commesseState.dirty && !confirm("Hai modifiche non salvate. Cambiare commessa?")) {
          projectSelect.value = commesseState.selectedProjectId;
          return;
        }

        commesseState.selectedProjectId = projectSelect.value || "";
        await loadProjectMatrix(commesseState.selectedProjectId);
        populateRoleSelect();
        renderCommesseBody();
      });
    }

    if (search && search.dataset.bound !== "1") {
      search.dataset.bound = "1";
      search.addEventListener("input", () => {
        populateProjectSelect();
      });
    }

    if (reload && reload.dataset.bound !== "1") {
      reload.dataset.bound = "1";
      reload.addEventListener("click", async () => {
        if (commesseState.selectedProjectId) {
          await loadProjectMatrix(commesseState.selectedProjectId);
        }
        populateRoleSelect();
        renderCommesseBody();
      });
    }

    if (close && close.dataset.bound !== "1") {
      close.dataset.bound = "1";
      close.addEventListener("click", showPlannerFromCommesse);
    }

    if (save && save.dataset.bound !== "1") {
      save.dataset.bound = "1";
      save.addEventListener("click", saveCommesseMatrix);
    }

    if (addRole && addRole.dataset.bound !== "1") {
      addRole.dataset.bound = "1";
      addRole.addEventListener("click", addCommesseRole);
    }
  }

  function addCommesseRole() {
    const select = document.getElementById("oldCommesseRoleSelect");
    const role = norm(select?.value || "");

    if (!commesseState.selectedProjectId) {
      alert("Seleziona prima una commessa.");
      return;
    }

    if (!role) {
      alert("Seleziona una mansione.");
      return;
    }

    commesseState.extraRoles.add(role);
    commesseState.dirty = true;
    populateRoleSelect();
    renderCommesseBody();

    const firstInput = document.querySelector(`input[data-role="${CSS.escape(role)}"]`);
    if (firstInput) {
      firstInput.focus();
      firstInput.select();
    }
  }

  function collectCommesseRows() {
    const rows = new Map();

    document.querySelectorAll("input[data-role][data-period-key]").forEach((input) => {
      const role = norm(input.dataset.role);
      const periodKey = String(input.dataset.periodKey || "");
      const quantity = numberValue(input.value);

      if (!role || !periodKey) return;

      if (!rows.has(role)) {
        rows.set(role, {});
      }

      rows.get(role)[periodKey] = quantity;
    });

    return rows;
  }

  async function saveCommesseMatrix() {
    const projectId = Number(commesseState.selectedProjectId || 0);

    if (!projectId) {
      alert("Seleziona una commessa.");
      return;
    }

    const saveBtn = document.getElementById("oldCommesseSaveBtn");
    const oldText = saveBtn?.textContent || "";

    if (saveBtn) {
      saveBtn.disabled = true;
      saveBtn.textContent = "Salvataggio...";
    }

    try {
      const rows = collectCommesseRows();
      let changedTotal = 0;

      for (const [role, quantities] of rows.entries()) {
        const hasAnyValue = Object.values(quantities).some((value) => Number(value || 0) > 0);
        const wasExisting = commesseState.matrix?.roles?.some((row) => norm(row.role) === role);

        // Se la riga e' nuova e tutta a zero, non salvo.
        if (!hasAnyValue && !wasExisting) {
          continue;
        }

        const result = await fetchJson(`/api/project-demand-matrix/${projectId}`, {
          method: "POST",
          headers: {"Content-Type": "application/json"},
          body: JSON.stringify({
            role,
            quantities,
          }),
        });

        if (!result.ok) {
          throw new Error(result.error || `Errore salvataggio ${role}`);
        }

        changedTotal += Number(result.changed || 0);
      }

      if (typeof loadAll === "function") {
        await loadAll();
      }

      await loadProjectMatrix(projectId);
      commesseState.extraRoles = new Set();
      commesseState.dirty = false;
      populateRoleSelect();
      renderCommesseBody();

      if (typeof renderPlanner === "function") {
        renderPlanner();
      }

      alert(`Salvataggio completato. Righe modificate: ${changedTotal}`);
    } catch (error) {
      console.error(error);
      alert(error.message || "Errore salvataggio fabbisogni commessa.");
    } finally {
      if (saveBtn) {
        saveBtn.disabled = false;
        saveBtn.textContent = oldText || "Salva modifiche";
      }
    }
  }

  window.renderOldWorkflowCommesse = renderCommesseSheet;

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", bindTopButtons);
  } else {
    bindTopButtons();
  }

  setTimeout(bindTopButtons, 500);
  setTimeout(bindTopButtons, 1500);
})();
// === OLD WORKFLOW COMMESSE/FABBISOGNI END ===

// === COMMESSE SHEET LAYOUT REFRESH FOCUS FIX START ===
(function commesseSheetLayoutRefreshFocusFix() {
  if (window.__commesseSheetLayoutRefreshFocusFixInstalled) {
    return;
  }
  window.__commesseSheetLayoutRefreshFocusFixInstalled = true;

  let lastCommessePeriodKey = Number(CURRENT_PERIOD_KEY || 2617);

  function norm(value) {
    if (typeof normalizeRole === "function") return normalizeRole(value);
    return String(value || "").trim().toUpperCase();
  }

  async function reloadPlannerDataAfterCommesseSave() {
    const stamp = Date.now();

    try { resourcesData = await fetchJson(`/api/resources?_=${stamp}`); } catch (error) {}
    try { projectsData = await fetchJson(`/api/projects?_=${stamp}`); } catch (error) {}
    try { demandsData = await fetchJson(`/api/demands?_=${stamp}`); } catch (error) {}
    try { allocationsData = await fetchJson(`/api/allocations?_=${stamp}`); } catch (error) {}
    try { allocationHistoryData = await fetchJson(`/api/allocation-history?_=${stamp}`); } catch (error) {}
    try { demandHistoryData = await fetchJson(`/api/demand-history?_=${stamp}`); } catch (error) {}

    if (typeof renderPlanner === "function") {
      renderPlanner();
    }

    if (typeof oldWorkflowRefreshPlanner === "function") {
      setTimeout(oldWorkflowRefreshPlanner, 100);
    }
  }

  function hidePlannerChromeForCommesse() {
    const commesse = document.getElementById("oldWorkflowCommesseSheet");
    if (!commesse || commesse.hidden) return;

    document.querySelectorAll(
      ".planner-main, .planner-side, .planner-bottom, #resourcesSheet, #oldWorkflowGanttSheet"
    ).forEach((node) => {
      if (node && node.id !== "oldWorkflowCommesseSheet") {
        node.hidden = true;
      }
    });

    commesse.hidden = false;
  }

  function showPlannerCleanFromCommesse() {
    const commesse = document.getElementById("oldWorkflowCommesseSheet");
    if (commesse) commesse.hidden = true;

    const plannerMain = document.querySelector(".planner-main");
    const plannerSide = document.querySelector(".planner-side");
    const plannerBottom = document.querySelector(".planner-bottom");

    if (plannerMain) plannerMain.hidden = false;
    if (plannerSide) plannerSide.hidden = false;
    if (plannerBottom) plannerBottom.hidden = false;

    reloadPlannerDataAfterCommesseSave();
  }

  function rememberFocusedCommesseWeek(event) {
    const input = event.target?.closest?.("input[data-period-key]");
    if (!input) return;

    const periodKey = Number(input.dataset.periodKey || 0);
    if (periodKey > 0) {
      lastCommessePeriodKey = periodKey;
    }
  }

  function focusRoleAtLastWeek(role) {
    const normalizedRole = norm(role);
    const selector = `input[data-role="${CSS.escape(normalizedRole)}"][data-period-key="${lastCommessePeriodKey}"]`;
    const input = document.querySelector(selector);

    if (input) {
      input.focus();
      input.select();

      const wrap = document.querySelector(".old-commesse-table-wrap");
      if (wrap) {
        const rect = input.getBoundingClientRect();
        const wrapRect = wrap.getBoundingClientRect();

        if (rect.left < wrapRect.left || rect.right > wrapRect.right) {
          input.scrollIntoView({ block: "nearest", inline: "center" });
        }
      }
    }
  }

  function patchAddRoleButton() {
    const btn = document.getElementById("oldCommesseAddRoleBtn");
    if (!btn || btn.dataset.focusFixBound === "1") return;

    btn.dataset.focusFixBound = "1";

    btn.addEventListener("click", () => {
      const role = norm(document.getElementById("oldCommesseRoleSelect")?.value || "");

      if (!role) return;

      setTimeout(() => focusRoleAtLastWeek(role), 80);
      setTimeout(() => focusRoleAtLastWeek(role), 220);
    }, true);
  }

  function patchCommesseSaveButton() {
    const btn = document.getElementById("oldCommesseSaveBtn");
    if (!btn || btn.dataset.refreshFixBound === "1") return;

    btn.dataset.refreshFixBound = "1";

    btn.addEventListener("click", () => {
      setTimeout(reloadPlannerDataAfterCommesseSave, 700);
      setTimeout(reloadPlannerDataAfterCommesseSave, 1600);
    }, true);
  }

  function patchCommesseCloseButton() {
    const btn = document.getElementById("oldCommesseCloseBtn");
    if (!btn || btn.dataset.closeFixBound === "1") return;

    btn.dataset.closeFixBound = "1";

    btn.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();
      showPlannerCleanFromCommesse();
    }, true);
  }

  function patchCommesseOpenButtons() {
    document.querySelectorAll(".topbar-actions button").forEach((button) => {
      const text = norm(button.textContent || "");

      if ((text === "PROGETTI" || text.includes("COMMESSE")) && button.dataset.commesseLayoutFixBound !== "1") {
        button.dataset.commesseLayoutFixBound = "1";

        button.addEventListener("click", () => {
          setTimeout(hidePlannerChromeForCommesse, 50);
          setTimeout(hidePlannerChromeForCommesse, 200);
        }, true);
      }

      if (text === "PLANNER" && button.dataset.commessePlannerRefreshFixBound !== "1") {
        button.dataset.commessePlannerRefreshFixBound = "1";

        button.addEventListener("click", () => {
          setTimeout(showPlannerCleanFromCommesse, 50);
        }, true);
      }
    });
  }

  function bindAll() {
    patchCommesseOpenButtons();
    patchAddRoleButton();
    patchCommesseSaveButton();
    patchCommesseCloseButton();
    document.addEventListener("focusin", rememberFocusedCommesseWeek, true);
    document.addEventListener("click", rememberFocusedCommesseWeek, true);
  }

  bindAll();
  setTimeout(bindAll, 500);
  setTimeout(bindAll, 1500);
  setTimeout(bindAll, 3000);

  window.reloadPlannerDataAfterCommesseSave = reloadPlannerDataAfterCommesseSave;
})();
// === COMMESSE SHEET LAYOUT REFRESH FOCUS FIX END ===

// === COMMESSE WEEK FOCUS AND HISTORY BADGE FIX START ===
(function commesseWeekFocusAndHistoryBadgeFix() {
  if (window.__commesseWeekFocusAndHistoryBadgeFixInstalled) {
    return;
  }
  window.__commesseWeekFocusAndHistoryBadgeFixInstalled = true;

  let lastCommessePeriodKeyFixed = Number(CURRENT_PERIOD_KEY || 2617);

  function norm(value) {
    if (typeof normalizeRole === "function") return normalizeRole(value);
    return String(value || "").trim().toUpperCase();
  }

  function getCommesseWrap() {
    return document.querySelector(".old-commesse-table-wrap");
  }

  function rememberPeriodFromInput(input) {
    if (!input || !input.matches?.("input[data-role][data-period-key]")) {
      return;
    }

    const periodKey = Number(input.dataset.periodKey || 0);
    if (periodKey > 0) {
      lastCommessePeriodKeyFixed = periodKey;
      window.__lastCommessePeriodKeyFixed = periodKey;
    }
  }

  function rememberPeriodFromViewport() {
    const wrap = getCommesseWrap();
    if (!wrap) return;

    const inputs = Array.from(wrap.querySelectorAll("input[data-role][data-period-key]"));
    if (!inputs.length) return;

    const wrapRect = wrap.getBoundingClientRect();
    const centerX = wrapRect.left + wrapRect.width / 2;

    let best = null;
    let bestDistance = Infinity;

    for (const input of inputs) {
      const rect = input.getBoundingClientRect();

      if (rect.right < wrapRect.left || rect.left > wrapRect.right) {
        continue;
      }

      const inputCenter = rect.left + rect.width / 2;
      const distance = Math.abs(inputCenter - centerX);

      if (distance < bestDistance) {
        bestDistance = distance;
        best = input;
      }
    }

    if (best) {
      rememberPeriodFromInput(best);
    }
  }

  function focusNewRoleAtRememberedWeek(role) {
    const normalizedRole = norm(role);
    const periodKey = Number(window.__lastCommessePeriodKeyFixed || lastCommessePeriodKeyFixed || CURRENT_PERIOD_KEY || 2617);

    if (!normalizedRole || !periodKey) return;

    const inputs = Array.from(document.querySelectorAll("input[data-role][data-period-key]"));

    const target =
      inputs.find((input) => norm(input.dataset.role) === normalizedRole && Number(input.dataset.periodKey) === periodKey) ||
      inputs.find((input) => norm(input.dataset.role) === normalizedRole && Number(input.dataset.periodKey) === Number(CURRENT_PERIOD_KEY)) ||
      inputs.find((input) => norm(input.dataset.role) === normalizedRole);

    if (!target) return;

    target.scrollIntoView({ block: "nearest", inline: "center" });
    target.focus();
    target.select();
  }

  document.addEventListener("focusin", (event) => {
    rememberPeriodFromInput(event.target);
  }, true);

  document.addEventListener("click", (event) => {
    rememberPeriodFromInput(event.target);
  }, true);

  document.addEventListener("input", (event) => {
    rememberPeriodFromInput(event.target);
  }, true);

  const wrapWatch = () => {
    const wrap = getCommesseWrap();
    if (!wrap || wrap.dataset.weekFocusScrollBound === "1") return;

    wrap.dataset.weekFocusScrollBound = "1";
    wrap.addEventListener("scroll", () => {
      window.clearTimeout(wrap.__weekFocusTimer);
      wrap.__weekFocusTimer = window.setTimeout(rememberPeriodFromViewport, 120);
    });
  };

  function patchAddRoleFocus() {
    const btn = document.getElementById("oldCommesseAddRoleBtn");
    if (!btn || btn.dataset.weekFocusHardPatchBound === "1") return;

    btn.dataset.weekFocusHardPatchBound = "1";

    btn.addEventListener("pointerdown", () => {
      // Prima del click salvo la settimana attualmente visibile.
      rememberPeriodFromViewport();

      const active = document.activeElement;
      if (active?.matches?.("input[data-role][data-period-key]")) {
        rememberPeriodFromInput(active);
      }
    }, true);

    btn.addEventListener("click", () => {
      const role = norm(document.getElementById("oldCommesseRoleSelect")?.value || "");
      if (!role) return;

      // Il vecchio handler mette il focus su W01. Noi lo correggiamo dopo il render.
      setTimeout(() => focusNewRoleAtRememberedWeek(role), 120);
      setTimeout(() => focusNewRoleAtRememberedWeek(role), 350);
      setTimeout(() => focusNewRoleAtRememberedWeek(role), 800);
    }, true);
  }

  function reloadPlannerAndHistoryAfterCommesseSave() {
    setTimeout(async () => {
      try {
        const stamp = Date.now();

        try { resourcesData = await fetchJson(`/api/resources?_=${stamp}`); } catch (error) {}
        try { projectsData = await fetchJson(`/api/projects?_=${stamp}`); } catch (error) {}
        try { demandsData = await fetchJson(`/api/demands?_=${stamp}`); } catch (error) {}
        try { allocationsData = await fetchJson(`/api/allocations?_=${stamp}`); } catch (error) {}
        try { allocationHistoryData = await fetchJson(`/api/allocation-history?_=${stamp}`); } catch (error) {}
        try { demandHistoryData = await fetchJson(`/api/demand-history?_=${stamp}`); } catch (error) {}

        if (typeof renderPlanner === "function") {
          renderPlanner();
        }

        if (typeof oldWorkflowRefreshPlanner === "function") {
          setTimeout(oldWorkflowRefreshPlanner, 150);
        }
      } catch (error) {
        console.error("Errore refresh storico dopo salvataggio commesse", error);
      }
    }, 500);
  }

  function patchSaveForHistoryRefresh() {
    const btn = document.getElementById("oldCommesseSaveBtn");
    if (!btn || btn.dataset.historyRefreshHardPatchBound === "1") return;

    btn.dataset.historyRefreshHardPatchBound = "1";
    btn.addEventListener("click", () => {
      reloadPlannerAndHistoryAfterCommesseSave();
      setTimeout(reloadPlannerAndHistoryAfterCommesseSave, 1000);
      setTimeout(reloadPlannerAndHistoryAfterCommesseSave, 2200);
    }, true);
  }

  function bindAll() {
    wrapWatch();
    patchAddRoleFocus();
    patchSaveForHistoryRefresh();
  }

  bindAll();
  setTimeout(bindAll, 500);
  setTimeout(bindAll, 1500);
  setTimeout(bindAll, 3000);
})();
// === COMMESSE WEEK FOCUS AND HISTORY BADGE FIX END ===

// === BASELINE CREATE PROJECT WORKFLOW START ===
(function baselineCreateProjectWorkflow() {
  if (window.__baselineCreateProjectWorkflowInstalled) {
    return;
  }
  window.__baselineCreateProjectWorkflowInstalled = true;

  const createState = {
    rows: new Map(),
  };

  function norm(value) {
    if (typeof normalizeRole === "function") return normalizeRole(value);
    return String(value || "").trim().toUpperCase();
  }

  function esc(value) {
    if (typeof escapeHtml === "function") return escapeHtml(value);
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function qty(value) {
    const parsed = Number(String(value ?? "").replace(",", "."));
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function fmt(value) {
    const n = Number(value || 0);
    if (!n) return "";
    if (Number.isInteger(n)) return String(n);
    return String(Math.round(n * 10) / 10).replace(".", ",");
  }

  function getRoleListForCreate() {
    const roles = new Set();

    if (typeof OLD_WORKFLOW_ROLE_ORDER !== "undefined" && Array.isArray(OLD_WORKFLOW_ROLE_ORDER)) {
      OLD_WORKFLOW_ROLE_ORDER.forEach((role) => roles.add(norm(role)));
    }

    for (const resource of resourcesData || []) {
      if (resource.role) roles.add(norm(resource.role));
    }

    for (const demand of demandsData || []) {
      if (demand.role) roles.add(norm(demand.role));
    }

    return Array.from(roles).sort((a, b) => {
      if (typeof oldWorkflowRoleSortKey === "function") {
        return oldWorkflowRoleSortKey(a).localeCompare(oldWorkflowRoleSortKey(b), "it", {
          numeric: true,
          sensitivity: "base",
        });
      }
      return a.localeCompare(b);
    });
  }

  function ensureCreateProjectModal() {
    let modal = document.getElementById("baselineCreateProjectModal");
    if (modal) return modal;

    modal = document.createElement("div");
    modal.id = "baselineCreateProjectModal";
    modal.className = "baseline-create-modal";
    modal.hidden = true;

    modal.innerHTML = `
      <div class="baseline-create-card">
        <div class="baseline-create-header">
          <div>
            <strong>Crea commessa</strong>
            <div>Le righe inserite qui diventano baseline iniziale: niente badge sulla creazione.</div>
          </div>
          <button class="btn btn-light" id="baselineCreateCloseBtn" type="button">Chiudi</button>
        </div>

        <div class="baseline-create-form">
          <label>
            <span>Codice / nome commessa</span>
            <input id="baselineProjectName" type="text" placeholder="Es. 500_26 (SITE)" />
          </label>

          <label>
            <span>Stato</span>
            <select id="baselineProjectStatus">
              <option value="ACTIVE">ACTIVE</option>
              <option value="PLANNED">PLANNED</option>
              <option value="CLOSED">CLOSED</option>
            </select>
          </label>

          <label class="baseline-create-note">
            <span>Note</span>
            <input id="baselineProjectNote" type="text" placeholder="Note iniziali..." />
          </label>
        </div>

        <div class="baseline-create-rolebar">
          <select id="baselineCreateRoleSelect"></select>
          <button class="btn btn-light" id="baselineCreateAddRoleBtn" type="button">Aggiungi mansione</button>
        </div>

        <div class="baseline-create-table-wrap">
          <table class="baseline-create-table">
            <thead id="baselineCreateHead"></thead>
            <tbody id="baselineCreateBody"></tbody>
          </table>
        </div>

        <div class="baseline-create-actions">
          <button class="btn btn-primary" id="baselineCreateSaveBtn" type="button">Crea commessa baseline</button>
        </div>
      </div>
    `;

    document.body.appendChild(modal);
    return modal;
  }

  function renderCreateHead() {
    const head = document.getElementById("baselineCreateHead");
    if (!head) return;

    head.innerHTML = `
      <tr>
        <th class="baseline-create-role-head">Mansione</th>
        ${PERIODS.map((period) => {
          const current = Number(period.periodKey) === Number(CURRENT_PERIOD_KEY) ? " current-week" : "";
          return `<th class="baseline-create-week${current}">${esc(period.label)}</th>`;
        }).join("")}
      </tr>
    `;
  }

  function renderCreateRoleSelect() {
    const select = document.getElementById("baselineCreateRoleSelect");
    if (!select) return;

    const existing = new Set(Array.from(createState.rows.keys()).map(norm));
    const roles = getRoleListForCreate().filter((role) => !existing.has(norm(role)));

    select.innerHTML =
      `<option value="">Seleziona mansione</option>` +
      roles.map((role) => `<option value="${esc(role)}">${esc(role)}</option>`).join("");
  }

  function renderCreateBody() {
    const body = document.getElementById("baselineCreateBody");
    if (!body) return;

    const rows = Array.from(createState.rows.keys()).sort((a, b) => {
      if (typeof oldWorkflowRoleSortKey === "function") {
        return oldWorkflowRoleSortKey(a).localeCompare(oldWorkflowRoleSortKey(b), "it", {
          numeric: true,
          sensitivity: "base",
        });
      }
      return a.localeCompare(b);
    });

    if (!rows.length) {
      body.innerHTML = `<tr><td class="baseline-create-empty" colspan="${PERIODS.length + 1}">Aggiungi una o più mansioni.</td></tr>`;
      return;
    }

    body.innerHTML = rows.map((role) => {
      const quantities = createState.rows.get(role) || {};

      const cells = PERIODS.map((period) => {
        const value = quantities[String(period.periodKey)] || 0;

        return `
          <td>
            <input
              type="number"
              min="0"
              step="0.5"
              value="${esc(fmt(value))}"
              data-create-role="${esc(role)}"
              data-create-period-key="${Number(period.periodKey)}"
            />
          </td>
        `;
      }).join("");

      return `
        <tr>
          <td class="baseline-create-role">${esc(role)}</td>
          ${cells}
        </tr>
      `;
    }).join("");

    body.querySelectorAll("input[data-create-role][data-create-period-key]").forEach((input) => {
      input.addEventListener("input", () => {
        const role = norm(input.dataset.createRole);
        const periodKey = String(input.dataset.createPeriodKey || "");
        if (!createState.rows.has(role)) createState.rows.set(role, {});
        createState.rows.get(role)[periodKey] = qty(input.value);
      });
    });
  }

  function openCreateProjectModal() {
    ensureCreateProjectModal();
    createState.rows = new Map();

    document.getElementById("baselineProjectName").value = "";
    document.getElementById("baselineProjectStatus").value = "ACTIVE";
    document.getElementById("baselineProjectNote").value = "";

    renderCreateHead();
    renderCreateRoleSelect();
    renderCreateBody();

    document.getElementById("baselineCreateProjectModal").hidden = false;
    setTimeout(() => document.getElementById("baselineProjectName")?.focus(), 50);
  }

  function closeCreateProjectModal() {
    const modal = document.getElementById("baselineCreateProjectModal");
    if (modal) modal.hidden = true;
  }

  function addCreateRole() {
    const role = norm(document.getElementById("baselineCreateRoleSelect")?.value || "");

    if (!role) {
      alert("Seleziona una mansione.");
      return;
    }

    createState.rows.set(role, {});
    renderCreateRoleSelect();
    renderCreateBody();

    const target =
      document.querySelector(`input[data-create-role="${CSS.escape(role)}"][data-create-period-key="${Number(CURRENT_PERIOD_KEY)}"]`) ||
      document.querySelector(`input[data-create-role="${CSS.escape(role)}"]`);

    if (target) {
      target.scrollIntoView({ block: "nearest", inline: "center" });
      target.focus();
      target.select();
    }
  }

  function collectCreateRows() {
    const rows = [];

    for (const [role, quantities] of createState.rows.entries()) {
      const cleaned = {};

      for (const period of PERIODS) {
        const value = qty(quantities[String(period.periodKey)] || 0);
        if (value > 0) {
          cleaned[String(period.periodKey)] = value;
        }
      }

      if (Object.keys(cleaned).length) {
        rows.push({
          role,
          quantities: cleaned,
        });
      }
    }

    return rows;
  }

  async function saveCreateProject() {
    const name = String(document.getElementById("baselineProjectName")?.value || "").trim();
    const status = String(document.getElementById("baselineProjectStatus")?.value || "ACTIVE").trim();
    const note = String(document.getElementById("baselineProjectNote")?.value || "").trim();

    if (!name) {
      alert("Inserisci codice/nome commessa.");
      return;
    }

    const rows = collectCreateRows();

    if (!rows.length) {
      alert("Inserisci almeno un fabbisogno iniziale.");
      return;
    }

    const btn = document.getElementById("baselineCreateSaveBtn");
    const oldText = btn?.textContent || "";

    if (btn) {
      btn.disabled = true;
      btn.textContent = "Creazione...";
    }

    try {
      const result = await fetchJson("/api/projects/create-baseline", {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify({
          name,
          status,
          note,
          rows,
        }),
      });

      if (!result.ok) {
        alert(result.error || "Errore creazione commessa.");
        return;
      }

      closeCreateProjectModal();

      if (typeof loadAll === "function") {
        await loadAll();
      }

      if (typeof renderPlanner === "function") {
        renderPlanner();
      }

      if (typeof renderOldWorkflowCommesse === "function") {
        setTimeout(renderOldWorkflowCommesse, 150);
      }

      alert(`Commessa creata. Baseline iniziale salvata. Righe fabbisogno: ${result.inserted || 0}`);
    } catch (error) {
      console.error(error);
      alert(error.message || "Errore creazione commessa.");
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.textContent = oldText || "Crea commessa baseline";
      }
    }
  }

  function injectCreateProjectButton() {
    const actions = document.querySelector(".old-commesse-actions");
    if (!actions || document.getElementById("baselineCreateProjectBtn")) return;

    const btn = document.createElement("button");
    btn.id = "baselineCreateProjectBtn";
    btn.className = "btn btn-primary";
    btn.type = "button";
    btn.textContent = "Crea commessa";

    actions.insertBefore(btn, actions.firstChild);
    btn.addEventListener("click", openCreateProjectModal);
  }

  function bindCreateModal() {
    ensureCreateProjectModal();

    const close = document.getElementById("baselineCreateCloseBtn");
    const add = document.getElementById("baselineCreateAddRoleBtn");
    const save = document.getElementById("baselineCreateSaveBtn");

    if (close && close.dataset.bound !== "1") {
      close.dataset.bound = "1";
      close.addEventListener("click", closeCreateProjectModal);
    }

    if (add && add.dataset.bound !== "1") {
      add.dataset.bound = "1";
      add.addEventListener("click", addCreateRole);
    }

    if (save && save.dataset.bound !== "1") {
      save.dataset.bound = "1";
      save.addEventListener("click", saveCreateProject);
    }
  }

  async function ensureBaselineOnExistingProjects() {
    try {
      await fetchJson("/api/projects/ensure-baseline", {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify({}),
      });
    } catch (error) {
      console.warn("Baseline projects non inizializzata", error);
    }
  }

  function bindAll() {
    injectCreateProjectButton();
    bindCreateModal();
  }

  ensureBaselineOnExistingProjects();

  bindAll();
  setTimeout(bindAll, 500);
  setTimeout(bindAll, 1500);
  setTimeout(bindAll, 3000);

  window.openCreateProjectModal = openCreateProjectModal;
})();
// === BASELINE CREATE PROJECT WORKFLOW END ===

// === PROJECTS SHEET V2 FRONTEND START ===
(function projectsSheetV2() {
  if (window.__projectsSheetV2Installed) {
    return;
  }
  window.__projectsSheetV2Installed = true;

  const stateV2 = {
    projects: [],
    selectedProjectId: "",
    matrix: null,
    extraRoles: new Set(),
    dirty: false,
    lastPeriodKey: Number(CURRENT_PERIOD_KEY || 2617),
    search: "",
  };

  function norm(value) {
    if (typeof normalizeRole === "function") return normalizeRole(value);
    return String(value || "").trim().toUpperCase();
  }

  function esc(value) {
    if (typeof escapeHtml === "function") return escapeHtml(value);
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function n(value) {
    const parsed = Number(String(value ?? "").replace(",", "."));
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function fmt(value) {
    const valueNumber = Number(value || 0);
    if (!valueNumber) return "";
    if (typeof formatNumber === "function") return formatNumber(valueNumber);
    if (Number.isInteger(valueNumber)) return String(valueNumber);
    return String(Math.round(valueNumber * 10) / 10).replace(".", ",");
  }

  function roleSortKeyV2(role) {
    if (typeof oldWorkflowRoleSortKey === "function") return oldWorkflowRoleSortKey(role);
    return norm(role);
  }

  function projectSortKeyV2(projectName) {
    if (typeof oldWorkflowProjectSortKey === "function") return oldWorkflowProjectSortKey(projectName);
    return norm(projectName);
  }

  function ensureSheet() {
    let sheet = document.getElementById("projectsSheetV2");
    if (sheet) return sheet;

    sheet = document.createElement("section");
    sheet.id = "projectsSheetV2";
    sheet.className = "projects-sheet-v2 card";
    sheet.hidden = true;

    sheet.innerHTML = `
      <div class="psv2-header">
        <div>
          <div class="side-title">Progetti / Commesse</div>
          <div class="psv2-subtitle">
            Foglio tabellare collegato a Planner, Gantt e Fabbisogni.
          </div>
        </div>

        <div class="psv2-actions">
          <button class="btn btn-primary" id="psv2NewProjectBtn" type="button">Nuova commessa</button>
          <button class="btn btn-light" id="psv2SaveProjectBtn" type="button">Salva commessa</button>
          <button class="btn btn-primary" id="psv2SaveDemandsBtn" type="button">Salva fabbisogni</button>
          <button class="btn btn-light" id="psv2RefreshBtn" type="button">Ricarica</button>
          <button class="btn btn-light" id="psv2CloseBtn" type="button">Chiudi</button>
        </div>
      </div>

      <div class="psv2-layout">
        <aside class="psv2-left">
          <div class="psv2-left-toolbar">
            <input id="psv2Search" type="text" placeholder="Cerca commessa..." />
          </div>
          <div class="psv2-project-table-wrap">
            <table class="psv2-project-table">
              <thead>
                <tr>
                  <th>Commessa</th>
                  <th>Tipo</th>
                </tr>
              </thead>
              <tbody id="psv2ProjectsBody"></tbody>
            </table>
          </div>
        </aside>

        <main class="psv2-main">
          <div class="psv2-form">
            <label>
              <span>Codice / nome</span>
              <input id="psv2Name" type="text" />
            </label>

            <label>
              <span>Cliente</span>
              <input id="psv2Client" type="text" />
            </label>

            <label>
              <span>Stato</span>
              <input id="psv2Status" type="text" />
            </label>

            <label>
              <span>Inizio</span>
              <input id="psv2StartDate" type="text" />
            </label>

            <label>
              <span>Fine</span>
              <input id="psv2EndDate" type="text" />
            </label>

            <label class="psv2-check">
              <input id="psv2WorkshopRollup" type="checkbox" />
              <span>Workshop / OVERALL rollup</span>
            </label>

            <label class="psv2-note">
              <span>Note</span>
              <input id="psv2Note" type="text" />
            </label>
          </div>

          <div class="psv2-info" id="psv2Info">
            Seleziona una commessa.
          </div>

          <div class="psv2-rolebar" id="psv2Rolebar">
            <div>
              <strong>Mansioni / Fabbisogni</strong>
              <span>Macro modifiche sulla commessa selezionata.</span>
            </div>
            <div>
              <select id="psv2RoleSelect"></select>
              <button class="btn btn-light" id="psv2AddRoleBtn" type="button">Aggiungi mansione</button>
            </div>
          </div>

          <div class="psv2-matrix-wrap">
            <table class="psv2-matrix">
              <thead id="psv2MatrixHead"></thead>
              <tbody id="psv2MatrixBody"></tbody>
            </table>
          </div>
        </main>
      </div>
    `;

    document.querySelector(".planner-layout")?.appendChild(sheet);
    return sheet;
  }

  function hideOperationalSheets() {
    const ids = [
      "oldWorkflowCommesseSheet",
      "oldWorkflowGanttSheet",
      "resourcesSheet",
    ];

    ids.forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.hidden = true;
    });

    document.querySelectorAll(".planner-main, .planner-side, .planner-bottom").forEach((el) => {
      el.hidden = true;
    });

    const sheet = ensureSheet();
    sheet.hidden = false;
  }

  function showPlanner() {
    const sheet = document.getElementById("projectsSheetV2");
    if (sheet) sheet.hidden = true;

    document.querySelectorAll(".planner-main, .planner-side, .planner-bottom").forEach((el) => {
      el.hidden = false;
    });

    reloadPlannerDataAfterProjectSheet();
  }

  async function reloadPlannerDataAfterProjectSheet() {
    const stamp = Date.now();

    try { resourcesData = await fetchJson(`/api/resources?_=${stamp}`); } catch (error) {}
    try { projectsData = await fetchJson(`/api/projects?_=${stamp}`); } catch (error) {}
    try { demandsData = await fetchJson(`/api/demands?_=${stamp}`); } catch (error) {}
    try { allocationsData = await fetchJson(`/api/allocations?_=${stamp}`); } catch (error) {}
    try { demandHistoryData = await fetchJson(`/api/demand-history?_=${stamp}`); } catch (error) {}
    try { allocationHistoryData = await fetchJson(`/api/allocation-history?_=${stamp}`); } catch (error) {}

    if (typeof renderPlanner === "function") renderPlanner();
    if (typeof oldWorkflowRefreshPlanner === "function") {
      setTimeout(oldWorkflowRefreshPlanner, 100);
    }
  }

  function isOverallProject(project) {
    if (!project) return false;
    if (project.is_overall) return true;
    const text = norm(`${project.name || ""} ${project.note || ""}`);
    return text.includes("OVERALL OFFICINA") || text.includes("WORKSHOP_ROLLUP");
  }

  async function loadProjects() {
    const result = await fetchJson(`/api/projects-sheet/projects?_=${Date.now()}`);
    if (!result.ok) throw new Error(result.error || "Errore caricamento commesse");

    stateV2.projects = result.projects || [];

    if (!stateV2.selectedProjectId && stateV2.projects.length) {
      stateV2.selectedProjectId = String(stateV2.projects[0].id);
    }
  }

  async function loadMatrix(projectId) {
    if (!projectId) {
      stateV2.matrix = null;
      return;
    }

    const result = await fetchJson(`/api/projects-sheet/matrix/${projectId}?_=${Date.now()}`);
    if (!result.ok) throw new Error(result.error || "Errore caricamento matrice commessa");

    stateV2.matrix = result;
    stateV2.extraRoles = new Set();
    stateV2.dirty = false;
  }

  function filteredProjects() {
    const search = norm(stateV2.search);

    return (stateV2.projects || [])
      .filter((project) => {
        if (!search) return true;
        return norm(`${project.name || ""} ${project.client || ""} ${project.status || ""} ${project.note || ""}`).includes(search);
      })
      .sort((a, b) => {
        return projectSortKeyV2(a.name).localeCompare(projectSortKeyV2(b.name), "it", {
          numeric: true,
          sensitivity: "base",
        });
      });
  }

  function selectedProject() {
    const id = Number(stateV2.selectedProjectId || 0);
    return (stateV2.projects || []).find((project) => Number(project.id) === id) || null;
  }

  function renderProjectsList() {
    const body = document.getElementById("psv2ProjectsBody");
    if (!body) return;

    const rows = filteredProjects();

    if (!rows.length) {
      body.innerHTML = `<tr><td colspan="2" class="psv2-empty">Nessuna commessa trovata.</td></tr>`;
      return;
    }

    body.innerHTML = rows.map((project) => {
      const selected = String(project.id) === String(stateV2.selectedProjectId);
      const type = isOverallProject(project) ? "OVERALL" : "INPUT";

      return `
        <tr class="${selected ? "selected" : ""}" data-project-id="${Number(project.id)}">
          <td>
            <strong>${esc(project.name || "")}</strong>
            <span>${esc(project.client || "")}</span>
          </td>
          <td>${esc(type)}</td>
        </tr>
      `;
    }).join("");

    body.querySelectorAll("[data-project-id]").forEach((row) => {
      row.addEventListener("click", async () => {
        if (stateV2.dirty && !confirm("Hai modifiche non salvate. Cambiare commessa?")) {
          return;
        }

        stateV2.selectedProjectId = String(row.dataset.projectId || "");
        await loadMatrix(stateV2.selectedProjectId);
        renderAll();
      });
    });
  }

  function fillProjectForm() {
    const project = stateV2.matrix?.project || selectedProject();

    document.getElementById("psv2Name").value = project?.name || "";
    document.getElementById("psv2Client").value = project?.client || "";
    document.getElementById("psv2Status").value = project?.status || "";
    document.getElementById("psv2StartDate").value = project?.start_date || "";
    document.getElementById("psv2EndDate").value = project?.end_date || "";
    document.getElementById("psv2Note").value = project?.note || "";
    document.getElementById("psv2WorkshopRollup").checked = isOverallProject(project);

    const isOverall = isOverallProject(project);
    document.getElementById("psv2WorkshopRollup").disabled = isOverall;
  }

  function allRoles() {
    const roles = new Set();

    if (typeof OLD_WORKFLOW_ROLE_ORDER !== "undefined" && Array.isArray(OLD_WORKFLOW_ROLE_ORDER)) {
      OLD_WORKFLOW_ROLE_ORDER.forEach((role) => roles.add(norm(role)));
    }

    for (const resource of resourcesData || []) {
      if (resource.role) roles.add(norm(resource.role));
    }

    for (const demand of demandsData || []) {
      if (demand.role) roles.add(norm(demand.role));
    }

    return Array.from(roles).sort((a, b) => {
      return roleSortKeyV2(a).localeCompare(roleSortKeyV2(b), "it", {
        numeric: true,
        sensitivity: "base",
      });
    });
  }

  function matrixRoles() {
    const rows = [];

    for (const row of stateV2.matrix?.roles || []) {
      const role = norm(row.role);
      const hasValue = Object.values(row.weeks || {}).some((cell) => Number(cell.quantity || 0) > 0);
      if (!role) continue;

      if (hasValue || Number(row.total || 0) > 0 || stateV2.extraRoles.has(role)) {
        rows.push({
          role,
          weeks: row.weeks || {},
          total: Number(row.total || 0),
          isExtra: stateV2.extraRoles.has(role),
        });
      }
    }

    for (const role of stateV2.extraRoles) {
      if (!rows.some((row) => norm(row.role) === role)) {
        rows.push({
          role,
          weeks: {},
          total: 0,
          isExtra: true,
        });
      }
    }

    return rows.sort((a, b) => {
      return roleSortKeyV2(a.role).localeCompare(roleSortKeyV2(b.role), "it", {
        numeric: true,
        sensitivity: "base",
      });
    });
  }

  function renderRoleSelect() {
    const select = document.getElementById("psv2RoleSelect");
    if (!select) return;

    const existing = new Set(matrixRoles().map((row) => norm(row.role)));
    const roles = allRoles().filter((role) => !existing.has(norm(role)));

    select.innerHTML =
      `<option value="">Seleziona mansione</option>` +
      roles.map((role) => `<option value="${esc(role)}">${esc(role)}</option>`).join("");
  }

  function renderHead() {
    const head = document.getElementById("psv2MatrixHead");
    if (!head) return;

    head.innerHTML = `
      <tr>
        <th class="psv2-role-head">Mansione / Commessa</th>
        <th class="psv2-total-head">Totale</th>
        ${PERIODS.map((period) => {
          const current = Number(period.periodKey) === Number(CURRENT_PERIOD_KEY) ? " current-week" : "";
          return `<th class="psv2-week${current}">${esc(period.label)}</th>`;
        }).join("")}
      </tr>
    `;
  }

  function renderOverallMatrix() {
    const body = document.getElementById("psv2MatrixBody");
    const info = document.getElementById("psv2Info");
    const rolebar = document.getElementById("psv2Rolebar");
    if (!body) return;

    if (rolebar) rolebar.hidden = true;

    const rows = stateV2.matrix?.workshop_rows || [];
    const byRole = new Map();

    for (const row of rows) {
      const role = norm(row.role || "");
      const projectName = String(row.project_name || row.source_project_name || "").trim();
      const periodKey = Number(row.period_key || 0);
      const required = Number(row.required || 0);

      if (!role || !projectName || !periodKey || required <= 0) continue;

      if (!byRole.has(role)) {
        byRole.set(role, {
          role,
          totals: new Map(),
          children: new Map(),
        });
      }

      const roleBucket = byRole.get(role);
      roleBucket.totals.set(periodKey, Number(roleBucket.totals.get(periodKey) || 0) + required);

      if (!roleBucket.children.has(projectName)) {
        roleBucket.children.set(projectName, {
          projectName,
          weeks: new Map(),
        });
      }

      const child = roleBucket.children.get(projectName);
      child.weeks.set(periodKey, Number(child.weeks.get(periodKey) || 0) + required);
    }

    const roleRows = Array.from(byRole.values()).sort((a, b) => {
      return roleSortKeyV2(a.role).localeCompare(roleSortKeyV2(b.role), "it", {
        numeric: true,
        sensitivity: "base",
      });
    });

    const childNames = new Set();
    roleRows.forEach((roleBucket) => {
      roleBucket.children.forEach((child) => childNames.add(child.projectName));
    });

    if (info) {
      info.innerHTML = `
        <strong>${esc(stateV2.matrix?.project?.name || "OVERALL OFFICINA")}</strong>
        <span>Vista rollup non editabile</span>
        <span>Sottocommesse: ${childNames.size}</span>
        <span>Fonte: workshop_rollup_sources</span>
      `;
    }

    if (!roleRows.length) {
      body.innerHTML = `
        <tr>
          <td class="psv2-empty" colspan="${PERIODS.length + 2}">
            Nessun dato trovato nel rollup officina.
          </td>
        </tr>
      `;
      return;
    }

    const html = [];

    for (const roleBucket of roleRows) {
      const roleTotal = PERIODS.reduce((sum, period) => {
        return sum + Number(roleBucket.totals.get(Number(period.periodKey)) || 0);
      }, 0);

      html.push(`
        <tr class="psv2-overall-role-row">
          <td class="psv2-role-cell"><strong>OVERALL - ${esc(roleBucket.role)}</strong></td>
          <td class="psv2-total-cell">${esc(fmt(roleTotal))}</td>
          ${PERIODS.map((period) => {
            const value = Number(roleBucket.totals.get(Number(period.periodKey)) || 0);
            return `<td class="psv2-overall-total-cell">${value > 0 ? esc(fmt(value)) : ""}</td>`;
          }).join("")}
        </tr>
      `);

      const children = Array.from(roleBucket.children.values()).sort((a, b) => {
        return projectSortKeyV2(a.projectName).localeCompare(projectSortKeyV2(b.projectName), "it", {
          numeric: true,
          sensitivity: "base",
        });
      });

      for (const child of children) {
        const childTotal = PERIODS.reduce((sum, period) => {
          return sum + Number(child.weeks.get(Number(period.periodKey)) || 0);
        }, 0);

        const matchingProject = (stateV2.projects || []).find((project) => {
          return norm(project.name) === norm(child.projectName);
        });

        html.push(`
          <tr class="psv2-overall-child-row" ${matchingProject ? `data-open-child-id="${Number(matchingProject.id)}"` : ""}>
            <td class="psv2-role-cell psv2-child-project">
              ${matchingProject ? `<button class="psv2-open-child" type="button" data-open-child-id="${Number(matchingProject.id)}">Apri</button>` : ""}
              ${esc(child.projectName)}
            </td>
            <td class="psv2-total-cell">${esc(fmt(childTotal))}</td>
            ${PERIODS.map((period) => {
              const value = Number(child.weeks.get(Number(period.periodKey)) || 0);
              return `<td class="psv2-overall-child-cell">${value > 0 ? esc(fmt(value)) : ""}</td>`;
            }).join("")}
          </tr>
        `);
      }
    }

    body.innerHTML = html.join("");

    body.querySelectorAll("[data-open-child-id]").forEach((node) => {
      node.addEventListener("click", async (event) => {
        event.stopPropagation();
        const id = Number(event.currentTarget.dataset.openChildId || 0);
        if (!id) return;

        stateV2.selectedProjectId = String(id);
        await loadMatrix(id);
        renderAll();
      });
    });
  }

  function renderInputMatrix() {
    const body = document.getElementById("psv2MatrixBody");
    const info = document.getElementById("psv2Info");
    const rolebar = document.getElementById("psv2Rolebar");
    if (!body) return;

    if (rolebar) rolebar.hidden = false;

    const project = stateV2.matrix?.project || selectedProject();
    const rows = matrixRoles();

    if (info) {
      const baseline = project?.baseline_at ? `Baseline: ${project.baseline_at}` : "Baseline non impostata";
      info.innerHTML = `
        <strong>${esc(project?.name || "Nuova commessa")}</strong>
        <span>${esc(baseline)}</span>
        <span>${rows.length} mansioni valorizzate</span>
      `;
    }

    if (!stateV2.selectedProjectId && !project) {
      body.innerHTML = `<tr><td class="psv2-empty" colspan="${PERIODS.length + 2}">Crea o seleziona una commessa.</td></tr>`;
      return;
    }

    if (!rows.length) {
      body.innerHTML = `<tr><td class="psv2-empty" colspan="${PERIODS.length + 2}">Aggiungi una mansione e valorizza le settimane.</td></tr>`;
      return;
    }

    body.innerHTML = rows.map((row) => {
      let total = 0;

      const cells = PERIODS.map((period) => {
        const cell = row.weeks[String(period.periodKey)] || {};
        const quantity = Number(cell.quantity || 0);
        total += quantity;

        return `
          <td class="psv2-input-cell">
            <input
              type="number"
              min="0"
              step="0.5"
              value="${esc(fmt(quantity))}"
              data-role="${esc(row.role)}"
              data-period-key="${Number(period.periodKey)}"
            />
          </td>
        `;
      }).join("");

      return `
        <tr data-role-row="${esc(row.role)}">
          <td class="psv2-role-cell">
            ${esc(row.role)}
            ${row.isExtra ? `<span class="psv2-new-role">nuova</span>` : ""}
          </td>
          <td class="psv2-total-cell" data-total-role="${esc(row.role)}">${esc(fmt(total))}</td>
          ${cells}
        </tr>
      `;
    }).join("");

    body.querySelectorAll("input[data-role][data-period-key]").forEach((input) => {
      input.addEventListener("focusin", () => rememberWeek(input));
      input.addEventListener("click", () => rememberWeek(input));
      input.addEventListener("input", () => {
        rememberWeek(input);
        stateV2.dirty = true;
        updateTotal(input.dataset.role);
      });
      input.addEventListener("keydown", handleKeyboard);
    });
  }

  function renderMatrix() {
    renderHead();

    if (stateV2.matrix?.is_overall || isOverallProject(stateV2.matrix?.project)) {
      renderOverallMatrix();
      return;
    }

    renderRoleSelect();
    renderInputMatrix();
  }

  function renderAll() {
    renderProjectsList();
    fillProjectForm();
    renderMatrix();
  }

  function updateTotal(role) {
    const normalized = norm(role);
    const inputs = Array.from(document.querySelectorAll(`input[data-role="${CSS.escape(normalized)}"][data-period-key]`));
    const total = inputs.reduce((sum, input) => sum + n(input.value), 0);
    const target = document.querySelector(`[data-total-role="${CSS.escape(normalized)}"]`);
    if (target) target.textContent = fmt(total);
  }

  function rememberWeek(input) {
    const periodKey = Number(input?.dataset?.periodKey || 0);
    if (periodKey > 0) {
      stateV2.lastPeriodKey = periodKey;
    }
  }

  function handleKeyboard(event) {
    const input = event.target;
    if (!input.matches("input[data-role][data-period-key]")) return;

    const td = input.closest("td");
    const tr = input.closest("tr");
    if (!td || !tr) return;

    const rowIndex = Array.from(tr.parentElement.children).indexOf(tr);
    const cellIndex = Array.from(tr.children).indexOf(td);

    let next = null;

    if (event.key === "ArrowRight") {
      next = tr.children[cellIndex + 1]?.querySelector("input");
    } else if (event.key === "ArrowLeft") {
      next = tr.children[cellIndex - 1]?.querySelector("input");
    } else if (event.key === "ArrowDown") {
      next = tr.parentElement.children[rowIndex + 1]?.children[cellIndex]?.querySelector("input");
    } else if (event.key === "ArrowUp") {
      next = tr.parentElement.children[rowIndex - 1]?.children[cellIndex]?.querySelector("input");
    }

    if (next) {
      event.preventDefault();
      next.focus();
      next.select();
    }
  }

  function addRole() {
    const role = norm(document.getElementById("psv2RoleSelect")?.value || "");
    if (!role) {
      alert("Seleziona una mansione.");
      return;
    }

    stateV2.extraRoles.add(role);
    stateV2.dirty = true;
    renderMatrix();

    const target =
      document.querySelector(`input[data-role="${CSS.escape(role)}"][data-period-key="${Number(stateV2.lastPeriodKey)}"]`) ||
      document.querySelector(`input[data-role="${CSS.escape(role)}"][data-period-key="${Number(CURRENT_PERIOD_KEY)}"]`) ||
      document.querySelector(`input[data-role="${CSS.escape(role)}"]`);

    if (target) {
      target.scrollIntoView({ block: "nearest", inline: "center" });
      target.focus();
      target.select();
    }
  }

  function collectRows() {
    const rows = new Map();

    document.querySelectorAll("#psv2MatrixBody input[data-role][data-period-key]").forEach((input) => {
      const role = norm(input.dataset.role || "");
      const periodKey = String(input.dataset.periodKey || "");
      if (!role || !periodKey) return;

      if (!rows.has(role)) rows.set(role, {});
      rows.get(role)[periodKey] = n(input.value);
    });

    return Array.from(rows.entries()).map(([role, quantities]) => ({ role, quantities }));
  }

  async function saveProjectOnly() {
    const current = stateV2.matrix?.project || selectedProject();

    const payload = {
      id: current?.id || 0,
      name: document.getElementById("psv2Name").value || "",
      client: document.getElementById("psv2Client").value || "",
      start_date: document.getElementById("psv2StartDate").value || "",
      end_date: document.getElementById("psv2EndDate").value || "",
      status: document.getElementById("psv2Status").value || "",
      note: document.getElementById("psv2Note").value || "",
      is_workshop_rollup: document.getElementById("psv2WorkshopRollup").checked,
    };

    const result = await fetchJson("/api/projects-sheet/save-project", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!result.ok) {
      throw new Error(result.error || "Errore salvataggio commessa");
    }

    stateV2.selectedProjectId = String(result.project.id);
    await loadProjects();
    await loadMatrix(stateV2.selectedProjectId);
    renderAll();

    return result.project;
  }

  async function saveProject() {
    const btn = document.getElementById("psv2SaveProjectBtn");
    const oldText = btn.textContent;
    btn.disabled = true;
    btn.textContent = "Salvataggio...";

    try {
      await saveProjectOnly();
      await reloadPlannerDataAfterProjectSheet();
      alert("Commessa salvata.");
    } catch (error) {
      console.error(error);
      alert(error.message || "Errore salvataggio commessa.");
    } finally {
      btn.disabled = false;
      btn.textContent = oldText;
    }
  }

  async function saveDemands() {
    const btn = document.getElementById("psv2SaveDemandsBtn");
    const oldText = btn.textContent;
    btn.disabled = true;
    btn.textContent = "Salvataggio...";

    try {
      let project = stateV2.matrix?.project || selectedProject();

      if (!project?.id) {
        project = await saveProjectOnly();
      }

      if (isOverallProject(project)) {
        alert("OVERALL è un rollup. Modifica le sottocommesse officina.");
        return;
      }

      const rows = collectRows();
      const baselineCreate = norm(project.note || "").includes("BASELINE_CREATE") && !(stateV2.matrix?.roles || []).length;

      const result = await fetchJson("/api/projects-sheet/save-demands", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          project_id: Number(project.id),
          rows,
          baseline_create: baselineCreate,
        }),
      });

      if (!result.ok) {
        throw new Error(result.error || "Errore salvataggio fabbisogni");
      }

      await loadProjects();
      await loadMatrix(project.id);
      renderAll();
      await reloadPlannerDataAfterProjectSheet();

      alert(`Fabbisogni salvati. Modifiche: ${result.changed || 0}, nuove celle: ${result.inserted || 0}.`);
    } catch (error) {
      console.error(error);
      alert(error.message || "Errore salvataggio fabbisogni.");
    } finally {
      btn.disabled = false;
      btn.textContent = oldText;
    }
  }

  function newProject() {
    stateV2.selectedProjectId = "";
    stateV2.matrix = {
      ok: true,
      is_overall: false,
      project: {
        id: 0,
        name: "",
        client: "",
        start_date: "",
        end_date: "",
        status: "attivo",
        note: "BASELINE_CREATE",
        baseline_at: "",
      },
      roles: [],
      workshop_rows: [],
    };
    stateV2.extraRoles = new Set();
    stateV2.dirty = false;
    renderAll();
    document.getElementById("psv2Name")?.focus();
  }

  async function openSheet() {
    ensureSheet();
    hideOperationalSheets();
    await loadProjects();

    if (stateV2.selectedProjectId) {
      await loadMatrix(stateV2.selectedProjectId);
    }

    renderAll();
    bindControls();
  }

  function bindControls() {
    const search = document.getElementById("psv2Search");
    const newBtn = document.getElementById("psv2NewProjectBtn");
    const saveProjectBtn = document.getElementById("psv2SaveProjectBtn");
    const saveDemandsBtn = document.getElementById("psv2SaveDemandsBtn");
    const refreshBtn = document.getElementById("psv2RefreshBtn");
    const closeBtn = document.getElementById("psv2CloseBtn");
    const addRoleBtn = document.getElementById("psv2AddRoleBtn");

    if (search && search.dataset.bound !== "1") {
      search.dataset.bound = "1";
      search.addEventListener("input", () => {
        stateV2.search = search.value || "";
        renderProjectsList();
      });
    }

    if (newBtn && newBtn.dataset.bound !== "1") {
      newBtn.dataset.bound = "1";
      newBtn.addEventListener("click", newProject);
    }

    if (saveProjectBtn && saveProjectBtn.dataset.bound !== "1") {
      saveProjectBtn.dataset.bound = "1";
      saveProjectBtn.addEventListener("click", saveProject);
    }

    if (saveDemandsBtn && saveDemandsBtn.dataset.bound !== "1") {
      saveDemandsBtn.dataset.bound = "1";
      saveDemandsBtn.addEventListener("click", saveDemands);
    }

    if (refreshBtn && refreshBtn.dataset.bound !== "1") {
      refreshBtn.dataset.bound = "1";
      refreshBtn.addEventListener("click", async () => {
        await loadProjects();
        if (stateV2.selectedProjectId) await loadMatrix(stateV2.selectedProjectId);
        renderAll();
      });
    }

    if (closeBtn && closeBtn.dataset.bound !== "1") {
      closeBtn.dataset.bound = "1";
      closeBtn.addEventListener("click", showPlanner);
    }

    if (addRoleBtn && addRoleBtn.dataset.bound !== "1") {
      addRoleBtn.dataset.bound = "1";
      addRoleBtn.addEventListener("click", addRole);
    }
  }

  function interceptTopbar() {
    document.querySelectorAll(".topbar-actions button").forEach((button) => {
      const text = norm(button.textContent || "");

      if ((text === "PROGETTI" || text.includes("COMMESSE")) && button.dataset.projectsSheetV2Bound !== "1") {
        button.dataset.projectsSheetV2Bound = "1";

        button.addEventListener("click", (event) => {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          openSheet().catch((error) => {
            console.error(error);
            alert(error.message || "Errore apertura foglio Progetti.");
          });
        }, true);
      }

      if (text === "PLANNER" && button.dataset.projectsSheetV2PlannerBound !== "1") {
        button.dataset.projectsSheetV2PlannerBound = "1";
        button.addEventListener("click", () => {
          const sheet = document.getElementById("projectsSheetV2");
          if (sheet) sheet.hidden = true;
        }, true);
      }
    });
  }

  window.openProjectsSheetV2 = openSheet;

  interceptTopbar();
  setTimeout(interceptTopbar, 500);
  setTimeout(interceptTopbar, 1500);
  setTimeout(interceptTopbar, 3000);
})();
// === PROJECTS SHEET V2 FRONTEND END ===

// === OLD WORKFLOW GANTT RESOURCES FRONTEND START ===
(function oldWorkflowGanttResourcesPort() {
  if (window.__oldWorkflowGanttResourcesPortInstalled) {
    return;
  }
  window.__oldWorkflowGanttResourcesPortInstalled = true;

  const state = {
    resources: [],
    projects: [],
    demands: [],
    allocations: [],
    allocationHistory: [],
    demandHistory: [],
    gantt: {
      mode: "resource",
      orderPeriodKey: typeof CURRENT_PERIOD_KEY !== "undefined" ? Number(CURRENT_PERIOD_KEY) : 2617,
      search: "",
      roleFilter: "",
      projectFilter: "",
      showExternalDetail: false,
      selected: null,
      drag: null,
      splitPercent: Number(localStorage.getItem("oldWorkflowGanttSplitPercent") || 66),
    },
    resourcesSheet: {
      search: "",
      roleFilter: "",
      statusFilter: "",
      sortBy: "name_asc",
      visibleColumns: loadVisibleResourceColumns(),
    },
  };

  const RESOURCE_COLUMNS = [
    { key: "id", label: "ID", readonly: true },
    { key: "code", label: "Codice", readonly: true },
    { key: "name", label: "Risorsa" },
    { key: "role", label: "Mansione 1" },
    { key: "role2", label: "Mansione 2", readonly: true },
    { key: "hire_date", label: "Assunzione", readonly: true },
    { key: "end_date", label: "Fine", readonly: true },
    { key: "birth_date", label: "Data nascita", readonly: true },
    { key: "phone", label: "Telefono", readonly: true },
    { key: "city", label: "Comune res.", readonly: true },
    { key: "employer", label: "Datore lavoro", readonly: true },
    { key: "level_hire", label: "Livello ass.", readonly: true },
    { key: "type", label: "Tipo", readonly: true },
    { key: "overtime", label: "€/h straord", readonly: true },
    { key: "day_off", label: "€/G Pres. Off", readonly: true },
    { key: "day_site", label: "€/G Pres. Cant", readonly: true },
    { key: "pb_fixed", label: "PB +FISSO", readonly: true },
    { key: "hour_off", label: "€/h Off.", readonly: true },
    { key: "hour_site", label: "€/h Cant.", readonly: true },
    { key: "glob_fixed", label: "GLOB +FISSO", readonly: true },
    { key: "doc_type", label: "Tipo DOC", readonly: true },
    { key: "doc_number", label: "Numero DOC", readonly: true },
    { key: "doc_expiry", label: "Scadenza DOC", readonly: true },
    { key: "email", label: "Email", readonly: true },
    { key: "site", label: "Sede", readonly: true },
    { key: "level", label: "Livello", readonly: true },
    { key: "certifications", label: "Certificazioni", readonly: true },
    { key: "no_travel", label: "No trasferta", readonly: true },
    { key: "availability_note", label: "Note operative" },
    { key: "is_active", label: "Stato" },
    { key: "action", label: "Azione", readonly: true },
  ];

  const DEFAULT_RESOURCE_VISIBLE_COLUMNS = [0, 1, 2, 3, 6, 10, 20, 22, 24, 28, 29, 30];

  function norm(value) {
    if (typeof normalizeRole === "function") return normalizeRole(value);
    return String(value || "").trim().toUpperCase();
  }

  function esc(value) {
    if (typeof escapeHtml === "function") return escapeHtml(value);
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function fmt(value) {
    if (typeof formatNumber === "function") return formatNumber(value);
    const n = Number(value || 0);
    if (!n) return "";
    if (Number.isInteger(n)) return String(n);
    return String(Math.round(n * 10) / 10).replace(".", ",");
  }

  function toPeriodKey(value) {
    const numeric = Number(value || 0);
    if (!numeric) return 0;
    if (numeric >= 1000) return numeric;
    if (typeof periodKeyFromWeek === "function") return periodKeyFromWeek(numeric);
    return Number(PERIOD_YEAR_SHORT || 26) * 100 + numeric;
  }

  function toWeek(value) {
    const numeric = Number(value || 0);
    if (!numeric) return 0;
    if (numeric >= 1000) return numeric % 100;
    return numeric;
  }

  function periodLabel(periodKey) {
    return String(Number(periodKey || 0));
  }

  function weekLabel(periodKey) {
    return `W${String(toWeek(periodKey)).padStart(2, "0")}`;
  }

  function normalizePeriod(row) {
    if (typeof normalizePeriodKey === "function") return normalizePeriodKey(row);
    return toPeriodKey(row?.period_key || row?.periodKey || row?.week || 0);
  }

  function projectById(projectId) {
    return state.projects.find((project) => Number(project.id) === Number(projectId)) || null;
  }

  function resourceById(resourceId) {
    return state.resources.find((resource) => Number(resource.id) === Number(resourceId)) || null;
  }

  function projectName(projectId) {
    return projectById(projectId)?.name || `Commessa ${projectId}`;
  }

  function resourceName(resourceId) {
    return resourceById(resourceId)?.name || `Risorsa ${resourceId}`;
  }

  function isExternalResourceLike(resource) {
    const text = norm(`${resource?.name || ""} ${resource?.role || ""} ${resource?.availability_note || ""}`);
    return text.includes("-EXT") || text.endsWith(" EXT") || text.includes(" ESTERNO");
  }

  function isInactiveResource(resource) {
    return Number(resource?.is_active) !== 1;
  }

  function isUnavailableResource(resource) {
    const text = norm(resource?.availability_note || "");
    return text.includes("INDISP") || text.includes("NON DISPONIBILE");
  }

  function resourceRole(resource) {
    return norm(resource?.role || resource?.resource_role || "");
  }

  function roleMatches(resource, role) {
    const r = resourceRole(resource);
    const target = norm(role);
    if (!target) return true;
    return r === target || r.includes(target) || target.includes(r);
  }

  function demandQuantity(projectId, role, periodKey) {
    const targetRole = norm(role);
    const targetPeriod = Number(periodKey);
    return (state.demands || []).reduce((sum, demand) => {
      if (Number(demand.project_id) !== Number(projectId)) return sum;
      if (norm(demand.role || "") !== targetRole) return sum;
      if (Number(normalizePeriod(demand)) !== targetPeriod) return sum;
      return sum + Number(demand.quantity || demand.qty || 0);
    }, 0);
  }

  function activeAllocations(projectId, role, periodKey) {
    const targetRole = norm(role);
    const targetPeriod = Number(periodKey);
    return (state.allocations || []).filter((allocation) => {
      return (
        Number(allocation.project_id) === Number(projectId) &&
        norm(allocation.role || "") === targetRole &&
        Number(normalizePeriod(allocation)) === targetPeriod
      );
    });
  }

  function resourceAllocations(resourceId, periodKey) {
    return (state.allocations || []).filter((allocation) => {
      return Number(allocation.resource_id) === Number(resourceId) && Number(normalizePeriod(allocation)) === Number(periodKey);
    });
  }

  function cellAnalysisForAllocation(allocation) {
    const periodKey = normalizePeriod(allocation);
    const required = demandQuantity(allocation.project_id, allocation.role, periodKey);
    const allocations = activeAllocations(allocation.project_id, allocation.role, periodKey);
    const allocated = allocations.reduce((sum, item) => sum + Number(item.load_percent || 100) / 100, 0);
    const diff = required - allocated;
    const resource = resourceById(allocation.resource_id);

    const noFab = allocated > 0 && (required <= 0 || allocated > required);
    const fuoriMansione = resource ? !roleMatches(resource, allocation.role) : false;
    const external = isExternalResourceLike(resource);
    const inactive = isInactiveResource(resource);
    const indisp = isUnavailableResource(resource);

    return {
      periodKey,
      required,
      allocated,
      diff,
      noFab,
      fuoriMansione,
      external,
      inactive,
      indisp,
      allocations,
    };
  }

  function projectColor(projectId) {
    const id = Number(projectId || 0);
    const hue = (id * 47) % 360;
    return `hsl(${hue}, 72%, 82%)`;
  }

  function projectBorder(projectId) {
    const id = Number(projectId || 0);
    const hue = (id * 47) % 360;
    return `hsl(${hue}, 66%, 38%)`;
  }

  function loadVisibleResourceColumns() {
    try {
      const raw = localStorage.getItem("oldWorkflowResourceVisibleColumns");
      const parsed = raw ? JSON.parse(raw) : null;
      if (Array.isArray(parsed)) return parsed;
    } catch (error) {}
    return DEFAULT_RESOURCE_VISIBLE_COLUMNS;
  }

  function saveVisibleResourceColumns(cols) {
    state.resourcesSheet.visibleColumns = cols;
    try {
      localStorage.setItem("oldWorkflowResourceVisibleColumns", JSON.stringify(cols));
    } catch (error) {}
  }

  function ensureShells() {
    ensureGantt();
    ensureResources();
    installOperationalIsolation();
  }

  function ensureGantt() {
    let sheet = document.getElementById("oldWorkflowGanttPort");
    if (sheet) return sheet;

    sheet = document.createElement("section");
    sheet.id = "oldWorkflowGanttPort";
    sheet.className = "old-workflow-sheet old-gantt-port card";
    sheet.hidden = true;

    sheet.innerHTML = `
      <div class="old-gantt-head">
        <div>
          <div class="side-title">Gantt Risorse</div>
          <div class="old-gantt-subtitle">Copia flusso vecchio: timeline risorse + Planner Fabbisogno sotto.</div>
        </div>
        <div class="old-gantt-actions">
          <input id="owGanttSearch" placeholder="Cerca risorsa..." />
          <select id="owGanttGroupBy">
            <option value="resource">Ordina: dipendente</option>
            <option value="project">Ordina: commessa su periodo</option>
          </select>
          <input id="owGanttOrderPeriod" type="number" min="1000" max="9999" step="1" />
          <select id="owGanttRoleFilter"><option value="">Tutte le mansioni</option></select>
          <select id="owGanttProjectFilter"><option value="">Tutte le commesse</option></select>
          <label class="old-gantt-check"><input id="owGanttShowExt" type="checkbox" /> <span>Mostra dettaglio external</span></label>
          <button class="btn btn-light" id="owGanttReloadBtn" type="button">Ricarica</button>
          <button class="btn btn-light" id="owGanttCloseBtn" type="button">Chiudi</button>
        </div>
      </div>

      <div class="old-gantt-legend">
        <span class="pill normal">Commessa</span>
        <span class="pill available">Disponibile</span>
        <span class="pill indisp">INDISP</span>
        <span class="pill overlap">Sovrapp.</span>
        <span class="pill contract">Contratto</span>
        <span class="pill ended">Cessato</span>
        <span class="pill nofab">NO FAB</span>
        <span class="pill fm">FM</span>
      </div>

      <div class="old-gantt-split" id="owGanttSplit">
        <div class="old-gantt-top" id="owGanttTop">
          <div class="old-gantt-grid-wrap">
            <table class="old-gantt-table">
              <thead id="owGanttHead"></thead>
              <tbody id="owGanttBody"></tbody>
            </table>
          </div>
        </div>

        <button class="old-gantt-split-handle" id="owGanttSplitHandle" type="button">
          <span></span>
        </button>

        <div class="old-gantt-bottom" id="owGanttBottom">
          <section class="old-gantt-demand-panel">
            <div class="mini-card-head">
              <h3>Planner Fabbisogno</h3>
              <div class="old-gantt-demand-filters">
                <select id="owGanttDemandProjectFilter"><option value="">Tutte le commesse</option></select>
                <select id="owGanttDemandRoleFilter"><option value="">Tutte le mansioni</option></select>
              </div>
            </div>
            <div class="old-gantt-demand-wrap">
              <table class="old-gantt-demand-table">
                <thead id="owGanttDemandHead"></thead>
                <tbody id="owGanttDemandBody"></tbody>
              </table>
            </div>
          </section>
        </div>
      </div>

      <div id="owGanttActionModal" class="old-modal-shell" hidden>
        <div class="old-modal-backdrop" data-close="1"></div>
        <section class="old-modal-card">
          <div class="old-modal-head">
            <h2>Azione Gantt</h2>
            <button class="btn btn-light" id="owGanttActionClose" type="button">Chiudi</button>
          </div>
          <div id="owGanttActionInfo" class="old-gantt-action-info"></div>
          <div class="old-gantt-action-grid">
            <label>
              <span>Commessa</span>
              <select id="owGanttActionProject"></select>
            </label>
            <label>
              <span>Mansione</span>
              <select id="owGanttActionRole"></select>
            </label>
            <label>
              <span>Periodo da</span>
              <input id="owGanttActionFrom" type="number" min="1000" max="9999" step="1" />
            </label>
            <label>
              <span>Periodo a</span>
              <input id="owGanttActionTo" type="number" min="1000" max="9999" step="1" />
            </label>
            <label>
              <span>Richiesto R</span>
              <input id="owGanttActionRequired" type="number" min="0" step="0.5" />
            </label>
            <label>
              <span>Ore</span>
              <input id="owGanttActionHours" type="number" min="0" step="1" value="40" />
            </label>
          </div>
          <div id="owGanttConflictBox" class="old-gantt-conflict-box"></div>
          <div class="old-gantt-action-buttons">
            <button class="btn btn-light danger-btn" id="owGanttActionDeleteIndisp" type="button" hidden>Elimina INDISP</button>
            <button class="btn btn-light danger-btn" id="owGanttActionUnassign" type="button">Svincola periodo</button>
            <button class="btn btn-light" id="owGanttActionDemandSave" type="button">Aggiorna fabbisogno</button>
            <button class="btn btn-primary" id="owGanttActionSave" type="button">Salva</button>
          </div>
        </section>
      </div>
    `;

    document.querySelector(".planner-layout")?.appendChild(sheet);
    return sheet;
  }

  function ensureResources() {
    let sheet = document.getElementById("oldWorkflowResourcesPort");
    if (sheet) return sheet;

    sheet = document.createElement("section");
    sheet.id = "oldWorkflowResourcesPort";
    sheet.className = "old-workflow-sheet old-resources-port card";
    sheet.hidden = true;

    sheet.innerHTML = `
      <div class="old-resources-head">
        <div>
          <div class="side-title">Anagrafica Risorse</div>
          <div class="old-resources-subtitle">Foglio stile vecchio: filtri, colonne, ordinamento, salvataggio.</div>
        </div>
        <div class="old-resources-actions">
          <input id="owResourcesSearch" placeholder="Cerca nome o codice" />
          <select id="owResourcesRoleFilter"><option value="">Tutte le mansioni</option></select>
          <select id="owResourcesStatusFilter">
            <option value="">Tutti gli stati</option>
            <option value="ATTIVO">Solo attivi</option>
            <option value="CESSATO">Solo cessati</option>
          </select>
          <select id="owResourcesSortBy">
            <option value="name_asc">Ordina: nome A-Z</option>
            <option value="name_desc">Ordina: nome Z-A</option>
            <option value="role_asc">Ordina: mansione A-Z</option>
            <option value="end_asc">Ordina: fine più vicina</option>
            <option value="end_desc">Ordina: fine più lontana</option>
          </select>
          <button class="btn btn-light" id="owResourcesColumnsBtn" type="button">Colonne</button>
          <button class="btn btn-primary" id="owResourcesSaveBtn" type="button">Salva modifiche</button>
          <button class="btn btn-light" id="owResourcesReloadBtn" type="button">Ricarica</button>
          <button class="btn btn-light" id="owResourcesCloseBtn" type="button">Chiudi</button>
        </div>
      </div>
      <div id="owResourcesColumnsPanel" class="old-resources-columns-panel" hidden></div>
      <div class="old-resources-wrap">
        <table class="old-resources-table">
          <colgroup id="owResourcesColgroup"></colgroup>
          <thead id="owResourcesHead"></thead>
          <tbody id="owResourcesBody"></tbody>
        </table>
      </div>
    `;

    document.querySelector(".planner-layout")?.appendChild(sheet);
    return sheet;
  }

  function isVisible(node) {
    if (!node || node.hidden) return false;
    const style = getComputedStyle(node);
    return style.display !== "none" && style.visibility !== "hidden";
  }

  function installOperationalIsolation() {
    if (window.__oldWorkflowOperationalIsolationInstalled) return;
    window.__oldWorkflowOperationalIsolationInstalled = true;

    function operationalOpen() {
      const ids = [
        "oldWorkflowGanttPort",
        "oldWorkflowResourcesPort",
        "projectsSheetV2",
        "oldWorkflowCommesseSheet",
        "ganttSheetV2",
        "resourcesSheet",
      ];
      return ids.some((id) => isVisible(document.getElementById(id)));
    }

    function update() {
      const active = operationalOpen();
      document.body.classList.toggle("operational-sheet-open", active);

      ["verticalSplitter", "horizontalSplitter", "sideInnerSplitter"].forEach((id) => {
        const node = document.getElementById(id);
        if (!node) return;
        node.style.display = active ? "none" : "";
        node.style.pointerEvents = active ? "none" : "";
      });

      document.querySelectorAll(".splitter, .splitter-vertical, .splitter-horizontal, .side-inner-splitter").forEach((node) => {
        node.style.display = active ? "none" : "";
        node.style.pointerEvents = active ? "none" : "";
      });

      if (active) {
        document.body.style.userSelect = "";
        document.body.style.cursor = "";
      }
    }

    document.addEventListener("pointerdown", (event) => {
      if (!operationalOpen()) return;
      if (
        event.target?.closest?.(".splitter") ||
        event.target?.closest?.(".splitter-vertical") ||
        event.target?.closest?.(".splitter-horizontal") ||
        event.target?.closest?.(".side-inner-splitter")
      ) {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
      }
    }, true);

    const observer = new MutationObserver(update);

    function observe() {
      ["oldWorkflowGanttPort", "oldWorkflowResourcesPort", "projectsSheetV2", "ganttSheetV2", "resourcesSheet"].forEach((id) => {
        const node = document.getElementById(id);
        if (node && node.dataset.isolationObserved !== "1") {
          node.dataset.isolationObserved = "1";
          observer.observe(node, { attributes: true, attributeFilter: ["hidden", "class", "style"] });
        }
      });
      update();
    }

    document.addEventListener("click", () => {
      setTimeout(observe, 20);
      setTimeout(update, 100);
    }, true);

    observe();
    setTimeout(observe, 500);
    setTimeout(observe, 1500);
  }

  function hideAllOperational() {
    ["oldWorkflowGanttPort", "oldWorkflowResourcesPort", "projectsSheetV2", "ganttSheetV2", "resourcesSheet", "oldWorkflowCommesseSheet"].forEach((id) => {
      const node = document.getElementById(id);
      if (node) node.hidden = true;
    });
  }

  function showPlanner() {
    hideAllOperational();
    document.querySelectorAll(".planner-main, .planner-side, .planner-bottom").forEach((node) => node.hidden = false);
    refreshV2Planner();
  }

  async function refreshV2Planner() {
    const stamp = Date.now();
    try { resourcesData = await fetchJson(`/api/resources?_=${stamp}`); } catch (e) {}
    try { projectsData = await fetchJson(`/api/projects?_=${stamp}`); } catch (e) {}
    try { demandsData = await fetchJson(`/api/demands?_=${stamp}`); } catch (e) {}
    try { allocationsData = await fetchJson(`/api/allocations?_=${stamp}`); } catch (e) {}
    try { allocationHistoryData = await fetchJson(`/api/allocation-history?_=${stamp}`); } catch (e) {}
    try { demandHistoryData = await fetchJson(`/api/demand-history?_=${stamp}`); } catch (e) {}
    if (typeof renderPlanner === "function") renderPlanner();
    if (typeof oldWorkflowRefreshPlanner === "function") setTimeout(oldWorkflowRefreshPlanner, 100);
  }

  async function loadState() {
    const result = await fetchJson(`/api/old-workflow/gantt-state?_=${Date.now()}`);
    if (!result.ok) throw new Error(result.error || "Errore caricamento Gantt");

    state.resources = result.resources || [];
    state.projects = result.projects || [];
    state.demands = result.demands || [];
    state.allocations = result.allocations || [];
    state.allocationHistory = result.allocation_history || [];
    state.demandHistory = result.demand_history || [];
  }

  function rolesList() {
    return Array.from(new Set([
      ...state.resources.map((r) => resourceRole(r)).filter(Boolean),
      ...state.demands.map((d) => norm(d.role)).filter(Boolean),
      ...state.allocations.map((a) => norm(a.role)).filter(Boolean),
    ])).sort((a, b) => a.localeCompare(b, "it", { numeric: true, sensitivity: "base" }));
  }

  function renderSelectOptions() {
    const roles = rolesList();
    const roleOptions = `<option value="">Tutte le mansioni</option>` + roles.map((role) => `<option value="${esc(role)}">${esc(role)}</option>`).join("");
    const roleOptionsAction = `<option value="">Seleziona mansione</option>` + roles.map((role) => `<option value="${esc(role)}">${esc(role)}</option>`).join("");

    ["owGanttRoleFilter", "owGanttDemandRoleFilter", "owResourcesRoleFilter"].forEach((id) => {
      const select = document.getElementById(id);
      if (!select) return;
      const old = select.value;
      select.innerHTML = roleOptions;
      select.value = old;
    });

    const actionRole = document.getElementById("owGanttActionRole");
    if (actionRole) {
      const old = actionRole.value;
      actionRole.innerHTML = roleOptionsAction;
      actionRole.value = old;
    }

    const projectOptions = `<option value="">Tutte le commesse</option>` + state.projects.map((project) => {
      return `<option value="${Number(project.id)}">${esc(project.name || "")}</option>`;
    }).join("");

    ["owGanttProjectFilter", "owGanttDemandProjectFilter"].forEach((id) => {
      const select = document.getElementById(id);
      if (!select) return;
      const old = select.value;
      select.innerHTML = projectOptions;
      select.value = old;
    });

    const actionProject = document.getElementById("owGanttActionProject");
    if (actionProject) {
      const old = actionProject.value;
      actionProject.innerHTML = `<option value="">Seleziona commessa</option>` + state.projects.map((project) => {
        return `<option value="${Number(project.id)}">${esc(project.name || "")}</option>`;
      }).join("");
      actionProject.value = old;
    }
  }

  function sortedGanttResources() {
    const search = norm(state.gantt.search);
    const roleFilter = norm(state.gantt.roleFilter);
    const projectFilter = Number(state.gantt.projectFilter || 0);

    let rows = state.resources.filter((resource) => {
      if (search && !norm(`${resource.name || ""} ${resource.role || ""} ${resource.availability_note || ""}`).includes(search)) return false;
      if (roleFilter && resourceRole(resource) !== roleFilter) return false;
      if (projectFilter) {
        const hasProject = state.allocations.some((a) => Number(a.resource_id) === Number(resource.id) && Number(a.project_id) === projectFilter);
        if (!hasProject) return false;
      }
      return true;
    });

    if (state.gantt.mode === "project") {
      const period = toPeriodKey(state.gantt.orderPeriodKey || CURRENT_PERIOD_KEY);
      const projectPopularity = new Map();

      for (const allocation of state.allocations) {
        if (Number(normalizePeriod(allocation)) !== Number(period)) continue;
        const key = Number(allocation.project_id);
        projectPopularity.set(key, (projectPopularity.get(key) || 0) + 1);
      }

      rows = rows.sort((a, b) => {
        const aAlloc = resourceAllocations(a.id, period)[0];
        const bAlloc = resourceAllocations(b.id, period)[0];
        const aHas = !!aAlloc;
        const bHas = !!bAlloc;

        if (aHas !== bHas) return aHas ? -1 : 1;

        if (!aHas && !bHas) {
          return String(a.name || "").localeCompare(String(b.name || ""), "it", { numeric: true, sensitivity: "base" });
        }

        const popA = projectPopularity.get(Number(aAlloc.project_id)) || 0;
        const popB = projectPopularity.get(Number(bAlloc.project_id)) || 0;
        if (popA !== popB) return popB - popA;

        const pCompare = projectName(aAlloc.project_id).localeCompare(projectName(bAlloc.project_id), "it", { numeric: true, sensitivity: "base" });
        if (pCompare !== 0) return pCompare;

        return String(a.name || "").localeCompare(String(b.name || ""), "it", { numeric: true, sensitivity: "base" });
      });
    } else {
      rows = rows.sort((a, b) => {
        return String(a.name || "").localeCompare(String(b.name || ""), "it", { numeric: true, sensitivity: "base" });
      });
    }

    return rows;
  }

  function renderGanttHead() {
    const head = document.getElementById("owGanttHead");
    if (!head) return;

    head.innerHTML = `
      <tr>
        <th class="ow-gantt-resource-head">Risorsa</th>
        <th class="ow-gantt-role-head">Mansione</th>
        ${PERIODS.map((period) => {
          const current = Number(period.periodKey) === Number(CURRENT_PERIOD_KEY) ? " current-week" : "";
          return `
            <th class="ow-gantt-week${current}" data-period-key="${Number(period.periodKey)}">
              <button type="button" data-sort-period="${Number(period.periodKey)}">
                <strong>${esc(periodLabel(period.periodKey))}</strong>
                <span>${esc(weekLabel(period.periodKey))}</span>
              </button>
            </th>
          `;
        }).join("")}
      </tr>
    `;

    head.querySelectorAll("[data-sort-period]").forEach((button) => {
      button.addEventListener("click", () => {
        state.gantt.mode = "project";
        state.gantt.orderPeriodKey = Number(button.dataset.sortPeriod || CURRENT_PERIOD_KEY);
        const mode = document.getElementById("owGanttGroupBy");
        const order = document.getElementById("owGanttOrderPeriod");
        if (mode) mode.value = "project";
        if (order) order.value = String(state.gantt.orderPeriodKey);
        renderGantt();
      });
    });
  }

  function allocationCellHtml(resource, period) {
    const periodKey = Number(period.periodKey);
    const allocations = resourceAllocations(resource.id, periodKey);
    const inactive = isInactiveResource(resource);
    const indisp = isUnavailableResource(resource);
    const external = isExternalResourceLike(resource);

    if (!allocations.length) {
      const cls = inactive ? " ended" : indisp ? " indisp" : " available";
      return `<td class="ow-gantt-cell${cls}" data-resource-id="${Number(resource.id)}" data-period-key="${periodKey}"></td>`;
    }

    const classes = ["ow-gantt-cell", "assigned"];
    if (allocations.length > 1) classes.push("overlap");
    if (inactive) classes.push("ended");
    if (indisp) classes.push("indisp");
    if (external) classes.push("external");

    const blocks = allocations.map((allocation) => {
      const analysis = cellAnalysisForAllocation(allocation);
      const flags = [];
      if (analysis.noFab) flags.push(`<span class="ow-badge nofab" title="NO FABBISOGNO">NO FAB</span>`);
      if (analysis.fuoriMansione) flags.push(`<span class="ow-badge fm" title="FUORI MANSIONE">FM</span>`);
      if (analysis.external) flags.push(`<span class="ow-badge ext" title="EXTERNAL">EXT</span>`);
      if (analysis.indisp) flags.push(`<span class="ow-badge ind" title="INDISPONIBILE">IND</span>`);

      const style = `--project-color:${projectColor(allocation.project_id)};--project-border:${projectBorder(allocation.project_id)}`;
      const project = projectName(allocation.project_id).replace(/\s*\((SITE|OFFICINA|CANTIERE SERVICE)\)\s*/gi, "").trim();

      return `
        <div class="ow-gantt-bar" style="${style}" data-allocation-id="${Number(allocation.id)}">
          <span>${esc(project)}</span>
          <small>${esc(norm(allocation.role || ""))}</small>
          ${flags.join("")}
        </div>
      `;
    }).join("");

    return `
      <td class="${classes.join(" ")}" data-resource-id="${Number(resource.id)}" data-period-key="${periodKey}">
        ${blocks}
      </td>
    `;
  }

  function renderGanttBody() {
    const body = document.getElementById("owGanttBody");
    if (!body) return;

    const rows = sortedGanttResources();

    if (!rows.length) {
      body.innerHTML = `<tr><td class="ow-empty" colspan="${PERIODS.length + 2}">Nessuna risorsa.</td></tr>`;
      return;
    }

    body.innerHTML = rows.map((resource) => {
      return `
        <tr data-resource-id="${Number(resource.id)}">
          <td class="ow-gantt-resource">${esc(resource.name || "")}</td>
          <td class="ow-gantt-role">${esc(resourceRole(resource))}</td>
          ${PERIODS.map((period) => allocationCellHtml(resource, period)).join("")}
        </tr>
      `;
    }).join("");

    bindGanttCells();
  }

  function bindGanttCells() {
    const body = document.getElementById("owGanttBody");
    if (!body) return;

    body.querySelectorAll(".ow-gantt-cell").forEach((cell) => {
      cell.addEventListener("pointerdown", (event) => {
        if (event.button !== 0) return;
        state.gantt.drag = {
          resourceId: Number(cell.dataset.resourceId || 0),
          startPeriod: Number(cell.dataset.periodKey || 0),
          endPeriod: Number(cell.dataset.periodKey || 0),
          active: true,
        };
        selectGanttCell(cell);
        markDragSelection();
      });

      cell.addEventListener("pointerenter", () => {
        if (!state.gantt.drag?.active) return;
        if (Number(cell.dataset.resourceId || 0) !== Number(state.gantt.drag.resourceId)) return;
        state.gantt.drag.endPeriod = Number(cell.dataset.periodKey || 0);
        markDragSelection();
      });

      cell.addEventListener("click", () => {
        selectGanttCell(cell);
      });
    });

    if (!window.__owGanttPointerUpBound) {
      window.__owGanttPointerUpBound = true;
      window.addEventListener("pointerup", () => {
        if (!state.gantt.drag?.active) return;
        const drag = state.gantt.drag;
        state.gantt.drag.active = false;

        const from = Math.min(drag.startPeriod, drag.endPeriod);
        const to = Math.max(drag.startPeriod, drag.endPeriod);

        if (state.gantt.selected && Number(state.gantt.selected.resourceId) === Number(drag.resourceId)) {
          state.gantt.selected.periodFrom = from;
          state.gantt.selected.periodTo = to;
          openGanttActionModal();
        }
      });
    }
  }

  function markDragSelection() {
    document.querySelectorAll(".ow-gantt-cell.drag-selected").forEach((node) => node.classList.remove("drag-selected"));

    const drag = state.gantt.drag;
    if (!drag) return;

    const from = Math.min(drag.startPeriod, drag.endPeriod);
    const to = Math.max(drag.startPeriod, drag.endPeriod);

    document.querySelectorAll(`.ow-gantt-cell[data-resource-id="${drag.resourceId}"]`).forEach((cell) => {
      const period = Number(cell.dataset.periodKey || 0);
      if (period >= from && period <= to) cell.classList.add("drag-selected");
    });
  }

  function selectGanttCell(cell) {
    const resourceId = Number(cell.dataset.resourceId || 0);
    const periodKey = Number(cell.dataset.periodKey || 0);
    const allocations = resourceAllocations(resourceId, periodKey);

    state.gantt.selected = {
      resourceId,
      periodKey,
      periodFrom: periodKey,
      periodTo: periodKey,
      allocations,
    };

    document.querySelectorAll(".ow-gantt-cell.selected").forEach((node) => node.classList.remove("selected"));
    cell.classList.add("selected");

    renderGanttDemandPanel();
    renderGanttActionInfoOnly();
  }

  function renderGanttActionInfoOnly() {
    const selected = state.gantt.selected;
    if (!selected) return;

    const detail = document.getElementById("owGanttConflictBox");
    if (!detail) return;

    const rows = selected.allocations.map((allocation) => {
      const analysis = cellAnalysisForAllocation(allocation);
      return `
        <div class="ow-gantt-detail-line">
          <strong>${esc(projectName(allocation.project_id))}</strong>
          <span>${esc(norm(allocation.role || ""))}</span>
          <span>R ${esc(fmt(analysis.required))}</span>
          <span>A ${esc(fmt(analysis.allocated))}</span>
          <span>D ${esc(fmt(analysis.diff))}</span>
          ${analysis.noFab ? `<b class="danger">NO FAB / SURPLUS</b>` : ""}
          ${analysis.fuoriMansione ? `<b class="danger">FUORI MANSIONE</b>` : ""}
        </div>
      `;
    });

    detail.innerHTML = rows.length ? rows.join("") : `<div class="ow-muted">Cella libera.</div>`;
  }

  function renderGanttDemandPanel() {
    renderGanttDemandFilters();

    const head = document.getElementById("owGanttDemandHead");
    const body = document.getElementById("owGanttDemandBody");
    if (!head || !body) return;

    const selected = state.gantt.selected;
    const selectedProjectFilter = Number(document.getElementById("owGanttDemandProjectFilter")?.value || 0);
    const selectedRoleFilter = norm(document.getElementById("owGanttDemandRoleFilter")?.value || "");

    const demandRows = new Map();

    for (const demand of state.demands) {
      const projectId = Number(demand.project_id);
      const role = norm(demand.role || "");
      const periodKey = normalizePeriod(demand);
      const quantity = Number(demand.quantity || demand.qty || 0);

      if (selectedProjectFilter && projectId !== selectedProjectFilter) continue;
      if (selectedRoleFilter && role !== selectedRoleFilter) continue;

      if (selected?.allocations?.length) {
        const relevant = selected.allocations.some((allocation) => {
          return Number(allocation.project_id) === projectId && norm(allocation.role || "") === role;
        });
        if (!relevant && !selectedProjectFilter && !selectedRoleFilter) continue;
      }

      const key = `${projectId}__${role}`;
      if (!demandRows.has(key)) {
        demandRows.set(key, {
          projectId,
          projectName: projectName(projectId),
          role,
          weeks: new Map(),
        });
      }

      demandRows.get(key).weeks.set(periodKey, quantity);
    }

    head.innerHTML = `
      <tr>
        <th class="ow-demand-project">Commessa</th>
        <th class="ow-demand-role">Mansione</th>
        ${PERIODS.map((period) => `<th>${esc(periodLabel(period.periodKey))}<span>${esc(weekLabel(period.periodKey))}</span></th>`).join("")}
      </tr>
    `;

    const rows = Array.from(demandRows.values()).sort((a, b) => {
      const p = a.projectName.localeCompare(b.projectName, "it", { numeric: true, sensitivity: "base" });
      if (p !== 0) return p;
      return a.role.localeCompare(b.role, "it", { numeric: true, sensitivity: "base" });
    });

    if (!rows.length) {
      body.innerHTML = `<tr><td class="ow-empty" colspan="${PERIODS.length + 2}">Nessun fabbisogno collegato alla selezione.</td></tr>`;
      return;
    }

    body.innerHTML = rows.map((row) => {
      return `
        <tr>
          <td class="ow-demand-project">${esc(row.projectName)}</td>
          <td class="ow-demand-role">${esc(row.role)}</td>
          ${PERIODS.map((period) => {
            const required = Number(row.weeks.get(Number(period.periodKey)) || 0);
            const allocated = activeAllocations(row.projectId, row.role, period.periodKey)
              .reduce((sum, allocation) => sum + Number(allocation.load_percent || 100) / 100, 0);
            const diff = required - allocated;
            const cls = required === 0 && allocated > 0 ? "surplus" : diff > 0 ? "missing" : allocated > required ? "surplus" : required || allocated ? "ok" : "";
            return `
              <td class="ow-demand-cell ${cls}" data-project-id="${row.projectId}" data-role="${esc(row.role)}" data-period-key="${Number(period.periodKey)}">
                ${required || allocated ? `R${fmt(required)} A${fmt(allocated)} D${fmt(diff)}` : ""}
              </td>
            `;
          }).join("")}
        </tr>
      `;
    }).join("");
  }

  function renderGanttDemandFilters() {
    const projectFilter = document.getElementById("owGanttDemandProjectFilter");
    const roleFilter = document.getElementById("owGanttDemandRoleFilter");
    if (!projectFilter || !roleFilter) return;

    const oldProject = projectFilter.value;
    const oldRole = roleFilter.value;

    projectFilter.innerHTML = `<option value="">Tutte le commesse</option>` + state.projects.map((project) => {
      return `<option value="${Number(project.id)}">${esc(project.name || "")}</option>`;
    }).join("");

    roleFilter.innerHTML = `<option value="">Tutte le mansioni</option>` + rolesList().map((role) => {
      return `<option value="${esc(role)}">${esc(role)}</option>`;
    }).join("");

    projectFilter.value = oldProject;
    roleFilter.value = oldRole;
  }

  function openGanttActionModal() {
    const selected = state.gantt.selected;
    if (!selected) return;

    const modal = document.getElementById("owGanttActionModal");
    const info = document.getElementById("owGanttActionInfo");
    const project = document.getElementById("owGanttActionProject");
    const role = document.getElementById("owGanttActionRole");
    const from = document.getElementById("owGanttActionFrom");
    const to = document.getElementById("owGanttActionTo");
    const required = document.getElementById("owGanttActionRequired");
    const hours = document.getElementById("owGanttActionHours");

    renderSelectOptions();

    const firstAlloc = selected.allocations[0];

    if (info) {
      info.innerHTML = `
        <strong>${esc(resourceName(selected.resourceId))}</strong>
        <span>${esc(String(selected.periodFrom || selected.periodKey))} - ${esc(String(selected.periodTo || selected.periodKey))}</span>
      `;
    }

    if (project) project.value = firstAlloc ? String(firstAlloc.project_id || "") : "";
    if (role) role.value = firstAlloc ? norm(firstAlloc.role || "") : resourceRole(resourceById(selected.resourceId));
    if (from) from.value = String(selected.periodFrom || selected.periodKey);
    if (to) to.value = String(selected.periodTo || selected.periodKey);
    if (hours) hours.value = firstAlloc ? String(firstAlloc.hours || 40) : "40";

    if (required) {
      const projectId = Number(project?.value || 0);
      const selectedRole = norm(role?.value || "");
      required.value = projectId && selectedRole ? String(demandQuantity(projectId, selectedRole, selected.periodKey) || 0) : "0";
    }

    renderGanttActionConflicts();

    if (modal) modal.hidden = false;
  }

  function renderGanttActionConflicts() {
    const selected = state.gantt.selected;
    const box = document.getElementById("owGanttConflictBox");
    if (!box || !selected) return;

    const lines = [];

    const periodFrom = toPeriodKey(document.getElementById("owGanttActionFrom")?.value || selected.periodFrom);
    const periodTo = toPeriodKey(document.getElementById("owGanttActionTo")?.value || selected.periodTo);
    const from = Math.min(periodFrom, periodTo);
    const to = Math.max(periodFrom, periodTo);

    for (const period of PERIODS) {
      const key = Number(period.periodKey);
      if (key < from || key > to) continue;

      const allocations = resourceAllocations(selected.resourceId, key);
      if (allocations.length > 1) {
        lines.push(`<div class="danger">SOVRAPP. ${key}: ${allocations.map((a) => esc(projectName(a.project_id))).join(", ")}</div>`);
      }

      allocations.forEach((allocation) => {
        const analysis = cellAnalysisForAllocation(allocation);
        if (analysis.noFab) lines.push(`<div class="danger">NO FAB ${key}: ${esc(projectName(allocation.project_id))} / ${esc(allocation.role)}</div>`);
        if (analysis.fuoriMansione) lines.push(`<div class="danger">FUORI MANSIONE ${key}: ${esc(resourceName(selected.resourceId))} su ${esc(allocation.role)}</div>`);
      });
    }

    box.innerHTML = lines.length ? lines.join("") : `<div class="ow-muted">Nessun conflitto rilevato.</div>`;
  }

  async function saveGanttAction() {
    const selected = state.gantt.selected;
    if (!selected) return;

    const payload = {
      resource_id: selected.resourceId,
      project_id: Number(document.getElementById("owGanttActionProject")?.value || 0),
      role: norm(document.getElementById("owGanttActionRole")?.value || ""),
      period_key_from: toPeriodKey(document.getElementById("owGanttActionFrom")?.value || selected.periodFrom),
      period_key_to: toPeriodKey(document.getElementById("owGanttActionTo")?.value || selected.periodTo),
      hours: Number(document.getElementById("owGanttActionHours")?.value || 40),
      note: "GANTT_OLD_WORKFLOW_ASSIGN",
    };

    const result = await fetchJson("/api/old-workflow/gantt-assign", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!result.ok) {
      alert(result.error || "Errore salvataggio Gantt.");
      return;
    }

    closeGanttActionModal();
    await loadState();
    renderGantt();
    await refreshV2Planner();
  }

  async function unassignGanttAction() {
    const selected = state.gantt.selected;
    if (!selected) return;

    const payload = {
      resource_id: selected.resourceId,
      project_id: Number(document.getElementById("owGanttActionProject")?.value || 0),
      role: norm(document.getElementById("owGanttActionRole")?.value || ""),
      period_key_from: toPeriodKey(document.getElementById("owGanttActionFrom")?.value || selected.periodFrom),
      period_key_to: toPeriodKey(document.getElementById("owGanttActionTo")?.value || selected.periodTo),
      reason: "GANTT_OLD_WORKFLOW_UNASSIGN",
      note: "Svincolo da Gantt old workflow",
    };

    const result = await fetchJson("/api/old-workflow/gantt-unassign", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!result.ok) {
      alert(result.error || "Errore svincolo.");
      return;
    }

    closeGanttActionModal();
    await loadState();
    renderGantt();
    await refreshV2Planner();
  }

  async function saveDemandFromGanttAction() {
    const projectId = Number(document.getElementById("owGanttActionProject")?.value || 0);
    const role = norm(document.getElementById("owGanttActionRole")?.value || "");
    const quantity = Number(document.getElementById("owGanttActionRequired")?.value || 0);
    const from = toPeriodKey(document.getElementById("owGanttActionFrom")?.value || 0);
    const to = toPeriodKey(document.getElementById("owGanttActionTo")?.value || from);

    const result = await fetchJson("/api/old-workflow/gantt-demand-upsert", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        project_id: projectId,
        role,
        quantity,
        period_key_from: from,
        period_key_to: to,
      }),
    });

    if (!result.ok) {
      alert(result.error || "Errore aggiornamento fabbisogno.");
      return;
    }

    await loadState();
    renderGantt();
    await refreshV2Planner();
    renderGanttActionConflicts();
  }

  function closeGanttActionModal() {
    const modal = document.getElementById("owGanttActionModal");
    if (modal) modal.hidden = true;
  }

  function renderGantt() {
    const order = document.getElementById("owGanttOrderPeriod");
    if (order && !order.value) order.value = String(state.gantt.orderPeriodKey || CURRENT_PERIOD_KEY);

    renderSelectOptions();
    renderGanttHead();
    renderGanttBody();
    renderGanttDemandPanel();
    bindGanttControls();
    applyGanttSplit();
  }

  function applyGanttSplit() {
    const split = document.getElementById("owGanttSplit");
    if (!split) return;
    const top = Math.max(35, Math.min(82, Number(state.gantt.splitPercent || 66)));
    split.style.setProperty("--ow-gantt-top", `${top}%`);
  }

  function bindGanttControls() {
    const map = [
      ["owGanttSearch", "input", (el) => { state.gantt.search = el.value || ""; renderGanttBody(); }],
      ["owGanttGroupBy", "change", (el) => { state.gantt.mode = el.value || "resource"; renderGanttBody(); }],
      ["owGanttOrderPeriod", "change", (el) => { state.gantt.orderPeriodKey = toPeriodKey(el.value || CURRENT_PERIOD_KEY); el.value = String(state.gantt.orderPeriodKey); renderGanttBody(); }],
      ["owGanttRoleFilter", "change", (el) => { state.gantt.roleFilter = el.value || ""; renderGanttBody(); }],
      ["owGanttProjectFilter", "change", (el) => { state.gantt.projectFilter = el.value || ""; renderGanttBody(); }],
      ["owGanttShowExt", "change", (el) => { state.gantt.showExternalDetail = el.checked; renderGanttBody(); }],
      ["owGanttDemandProjectFilter", "change", () => renderGanttDemandPanel()],
      ["owGanttDemandRoleFilter", "change", () => renderGanttDemandPanel()],
    ];

    map.forEach(([id, eventName, handler]) => {
      const el = document.getElementById(id);
      if (!el || el.dataset.bound === "1") return;
      el.dataset.bound = "1";
      el.addEventListener(eventName, () => handler(el));
    });

    const reload = document.getElementById("owGanttReloadBtn");
    if (reload && reload.dataset.bound !== "1") {
      reload.dataset.bound = "1";
      reload.addEventListener("click", async () => {
        await loadState();
        renderGantt();
      });
    }

    const close = document.getElementById("owGanttCloseBtn");
    if (close && close.dataset.bound !== "1") {
      close.dataset.bound = "1";
      close.addEventListener("click", showPlanner);
    }

    const modalClose = document.getElementById("owGanttActionClose");
    if (modalClose && modalClose.dataset.bound !== "1") {
      modalClose.dataset.bound = "1";
      modalClose.addEventListener("click", closeGanttActionModal);
    }

    const modal = document.getElementById("owGanttActionModal");
    if (modal && modal.dataset.bound !== "1") {
      modal.dataset.bound = "1";
      modal.addEventListener("click", (event) => {
        if (event.target?.dataset?.close === "1") closeGanttActionModal();
      });
    }

    const save = document.getElementById("owGanttActionSave");
    if (save && save.dataset.bound !== "1") {
      save.dataset.bound = "1";
      save.addEventListener("click", () => saveGanttAction().catch((err) => alert(err.message || "Errore Gantt")));
    }

    const unassign = document.getElementById("owGanttActionUnassign");
    if (unassign && unassign.dataset.bound !== "1") {
      unassign.dataset.bound = "1";
      unassign.addEventListener("click", () => unassignGanttAction().catch((err) => alert(err.message || "Errore svincolo")));
    }

    const demandSave = document.getElementById("owGanttActionDemandSave");
    if (demandSave && demandSave.dataset.bound !== "1") {
      demandSave.dataset.bound = "1";
      demandSave.addEventListener("click", () => saveDemandFromGanttAction().catch((err) => alert(err.message || "Errore fabbisogno")));
    }

    ["owGanttActionProject", "owGanttActionRole", "owGanttActionFrom", "owGanttActionTo"].forEach((id) => {
      const el = document.getElementById(id);
      if (!el || el.dataset.conflictBound === "1") return;
      el.dataset.conflictBound = "1";
      el.addEventListener("change", renderGanttActionConflicts);
      el.addEventListener("input", renderGanttActionConflicts);
    });

    bindGanttSplit();
  }

  function bindGanttSplit() {
    const handle = document.getElementById("owGanttSplitHandle");
    if (!handle || handle.dataset.bound === "1") return;
    handle.dataset.bound = "1";

    let dragging = null;

    handle.addEventListener("pointerdown", (event) => {
      const split = document.getElementById("owGanttSplit");
      if (!split) return;
      dragging = {
        top: split.getBoundingClientRect().top,
        height: split.getBoundingClientRect().height,
      };
      handle.setPointerCapture?.(event.pointerId);
      event.preventDefault();
    });

    handle.addEventListener("pointermove", (event) => {
      if (!dragging) return;
      const pct = ((event.clientY - dragging.top) / dragging.height) * 100;
      state.gantt.splitPercent = Math.max(35, Math.min(82, pct));
      localStorage.setItem("oldWorkflowGanttSplitPercent", String(state.gantt.splitPercent));
      applyGanttSplit();
    });

    const stop = () => { dragging = null; };
    handle.addEventListener("pointerup", stop);
    handle.addEventListener("pointercancel", stop);
  }

  async function openGantt() {
    ensureShells();
    hideAllOperational();
    document.querySelectorAll(".planner-main, .planner-side, .planner-bottom").forEach((node) => node.hidden = true);
    document.getElementById("oldWorkflowGanttPort").hidden = false;

    await loadState();
    renderGantt();

    setTimeout(() => {
      const current = document.querySelector(".ow-gantt-week.current-week");
      const wrap = document.querySelector(".old-gantt-grid-wrap");
      if (current && wrap) wrap.scrollLeft = Math.max(0, current.offsetLeft - 360);
    }, 80);
  }

  function renderResourcesSheet() {
    renderResourceFilters();
    renderResourceColumnsPanel();
    renderResourceTable();
    bindResourceControls();
  }

  function filteredResourcesSheetRows() {
    const search = norm(state.resourcesSheet.search);
    const roleFilter = norm(state.resourcesSheet.roleFilter);
    const status = state.resourcesSheet.statusFilter;
    const sortBy = state.resourcesSheet.sortBy;

    let rows = state.resources.filter((resource) => {
      if (search && !norm(`${resource.id || ""} ${resource.name || ""} ${resource.role || ""} ${resource.availability_note || ""}`).includes(search)) return false;
      if (roleFilter && resourceRole(resource) !== roleFilter) return false;
      if (status === "ATTIVO" && Number(resource.is_active) !== 1) return false;
      if (status === "CESSATO" && Number(resource.is_active) === 1) return false;
      return true;
    });

    rows = rows.sort((a, b) => {
      if (sortBy === "name_desc") return String(b.name || "").localeCompare(String(a.name || ""), "it", { numeric: true, sensitivity: "base" });
      if (sortBy === "role_asc") {
        const r = resourceRole(a).localeCompare(resourceRole(b), "it", { numeric: true, sensitivity: "base" });
        if (r !== 0) return r;
        return String(a.name || "").localeCompare(String(b.name || ""), "it", { numeric: true, sensitivity: "base" });
      }
      if (sortBy === "end_asc" || sortBy === "end_desc") {
        const ae = String(a.end_date || a.availability_note || "");
        const be = String(b.end_date || b.availability_note || "");
        const cmp = ae.localeCompare(be, "it", { numeric: true, sensitivity: "base" });
        return sortBy === "end_desc" ? -cmp : cmp;
      }
      return String(a.name || "").localeCompare(String(b.name || ""), "it", { numeric: true, sensitivity: "base" });
    });

    return rows;
  }

  function resourceValue(resource, col) {
    if (col.key === "code") return resource.code || resource.id || "";
    if (col.key === "role2") return resource.role2 || "";
    if (col.key === "is_active") return Number(resource.is_active) === 1 ? "ATTIVO" : "CESSATO";
    return resource[col.key] ?? "";
  }

  function renderResourceFilters() {
    const roleFilter = document.getElementById("owResourcesRoleFilter");
    if (roleFilter) {
      const old = roleFilter.value;
      roleFilter.innerHTML = `<option value="">Tutte le mansioni</option>` + rolesList().map((role) => `<option value="${esc(role)}">${esc(role)}</option>`).join("");
      roleFilter.value = old;
    }
  }

  function renderResourceColumnsPanel() {
    const panel = document.getElementById("owResourcesColumnsPanel");
    if (!panel) return;

    const visible = new Set(state.resourcesSheet.visibleColumns);

    panel.innerHTML = `
      <div class="ow-columns-actions">
        <button class="btn btn-light" type="button" data-columns="base">Base</button>
        <button class="btn btn-light" type="button" data-columns="all">Tutte</button>
      </div>
      ${RESOURCE_COLUMNS.map((col, idx) => `
        <label>
          <input type="checkbox" data-column-index="${idx}" ${visible.has(idx) ? "checked" : ""} ${idx === 2 || idx === 29 || idx === 30 ? "disabled" : ""}/>
          <span>${esc(col.label)}</span>
        </label>
      `).join("")}
    `;

    panel.querySelectorAll("[data-columns]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const cols = btn.dataset.columns === "all" ? RESOURCE_COLUMNS.map((_, idx) => idx) : DEFAULT_RESOURCE_VISIBLE_COLUMNS;
        saveVisibleResourceColumns(cols);
        renderResourcesSheet();
      });
    });

    panel.querySelectorAll("[data-column-index]").forEach((input) => {
      input.addEventListener("change", () => {
        const cols = Array.from(panel.querySelectorAll("[data-column-index]:checked")).map((node) => Number(node.dataset.columnIndex));
        if (!cols.includes(2)) cols.push(2);
        if (!cols.includes(29)) cols.push(29);
        if (!cols.includes(30)) cols.push(30);
        saveVisibleResourceColumns(cols.sort((a, b) => a - b));
        renderResourceTable();
      });
    });
  }

  function renderResourceTable() {
    const head = document.getElementById("owResourcesHead");
    const body = document.getElementById("owResourcesBody");
    const colgroup = document.getElementById("owResourcesColgroup");
    if (!head || !body || !colgroup) return;

    const visible = state.resourcesSheet.visibleColumns;
    const cols = visible.map((idx) => RESOURCE_COLUMNS[idx]).filter(Boolean);

    colgroup.innerHTML = cols.map(() => `<col />`).join("");
    head.innerHTML = `<tr>${cols.map((col) => `<th>${esc(col.label)}</th>`).join("")}</tr>`;

    const rows = filteredResourcesSheetRows();

    if (!rows.length) {
      body.innerHTML = `<tr><td class="ow-empty" colspan="${cols.length}">Nessuna risorsa.</td></tr>`;
      return;
    }

    body.innerHTML = rows.map((resource) => {
      return `
        <tr data-resource-id="${Number(resource.id)}">
          ${cols.map((col) => {
            const value = resourceValue(resource, col);
            if (col.key === "action") {
              return `<td><button class="btn btn-light ow-resource-row-save" type="button">Salva</button></td>`;
            }
            if (col.key === "is_active") {
              return `
                <td>
                  <select data-field="is_active">
                    <option value="1" ${Number(resource.is_active) === 1 ? "selected" : ""}>ATTIVO</option>
                    <option value="0" ${Number(resource.is_active) !== 1 ? "selected" : ""}>CESSATO</option>
                  </select>
                </td>
              `;
            }
            if (col.readonly) {
              return `<td>${esc(value)}</td>`;
            }
            return `<td><input data-field="${esc(col.key)}" value="${esc(value)}" /></td>`;
          }).join("")}
        </tr>
      `;
    }).join("");

    bindResourceTableInputs();
  }

  function bindResourceTableInputs() {
    const body = document.getElementById("owResourcesBody");
    if (!body) return;

    body.querySelectorAll("input, select").forEach((input) => {
      input.addEventListener("input", () => {
        input.closest("tr")?.classList.add("dirty");
      });
      input.addEventListener("change", () => {
        input.closest("tr")?.classList.add("dirty");
      });
    });

    body.querySelectorAll(".ow-resource-row-save").forEach((button) => {
      button.addEventListener("click", async () => {
        await saveResources(button.closest("tr"));
      });
    });

    body.querySelectorAll("input, select").forEach((el) => {
      el.addEventListener("keydown", (event) => {
        const td = event.target.closest("td");
        const tr = event.target.closest("tr");
        if (!td || !tr) return;

        const rows = Array.from(body.querySelectorAll("tr"));
        const rowIndex = rows.indexOf(tr);
        const cellIndex = Array.from(tr.children).indexOf(td);
        let target = null;

        if (event.key === "ArrowDown") target = rows[rowIndex + 1]?.children[cellIndex]?.querySelector("input, select");
        if (event.key === "ArrowUp") target = rows[rowIndex - 1]?.children[cellIndex]?.querySelector("input, select");
        if (event.key === "ArrowRight") target = td.nextElementSibling?.querySelector("input, select");
        if (event.key === "ArrowLeft") target = td.previousElementSibling?.querySelector("input, select");

        if (target) {
          event.preventDefault();
          target.focus();
          if (target.select) target.select();
        }
      });
    });
  }

  function collectResourceRows(scope = null) {
    const rows = [];
    const trs = scope ? [scope] : Array.from(document.querySelectorAll("#owResourcesBody tr.dirty"));

    trs.forEach((tr) => {
      const id = Number(tr.dataset.resourceId || 0);
      if (!id) return;

      const original = resourceById(id) || {};
      const row = {
        id,
        name: original.name || "",
        role: original.role || "",
        availability_note: original.availability_note || "",
        is_active: Number(original.is_active) === 1 ? 1 : 0,
      };

      tr.querySelectorAll("[data-field]").forEach((input) => {
        const field = input.dataset.field;
        if (field === "is_active") row.is_active = Number(input.value || 0);
        else row[field] = input.value || "";
      });

      rows.push(row);
    });

    return rows;
  }

  async function saveResources(scope = null) {
    const rows = collectResourceRows(scope);
    if (!rows.length) {
      alert("Nessuna modifica da salvare.");
      return;
    }

    const result = await fetchJson("/api/old-workflow/resources-save", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ rows }),
    });

    if (!result.ok) {
      alert(result.error || "Errore salvataggio risorse.");
      return;
    }

    await loadState();
    renderResourcesSheet();
    await refreshV2Planner();

    alert(`Risorse salvate: ${result.changed || 0}`);
  }

  function bindResourceControls() {
    const bindings = [
      ["owResourcesSearch", "input", (el) => { state.resourcesSheet.search = el.value || ""; renderResourceTable(); }],
      ["owResourcesRoleFilter", "change", (el) => { state.resourcesSheet.roleFilter = el.value || ""; renderResourceTable(); }],
      ["owResourcesStatusFilter", "change", (el) => { state.resourcesSheet.statusFilter = el.value || ""; renderResourceTable(); }],
      ["owResourcesSortBy", "change", (el) => { state.resourcesSheet.sortBy = el.value || "name_asc"; renderResourceTable(); }],
    ];

    bindings.forEach(([id, eventName, handler]) => {
      const el = document.getElementById(id);
      if (!el || el.dataset.bound === "1") return;
      el.dataset.bound = "1";
      el.addEventListener(eventName, () => handler(el));
    });

    const columnsBtn = document.getElementById("owResourcesColumnsBtn");
    const panel = document.getElementById("owResourcesColumnsPanel");
    if (columnsBtn && columnsBtn.dataset.bound !== "1") {
      columnsBtn.dataset.bound = "1";
      columnsBtn.addEventListener("click", () => {
        if (panel) panel.hidden = !panel.hidden;
      });
    }

    const save = document.getElementById("owResourcesSaveBtn");
    if (save && save.dataset.bound !== "1") {
      save.dataset.bound = "1";
      save.addEventListener("click", () => saveResources().catch((err) => alert(err.message || "Errore salvataggio risorse")));
    }

    const reload = document.getElementById("owResourcesReloadBtn");
    if (reload && reload.dataset.bound !== "1") {
      reload.dataset.bound = "1";
      reload.addEventListener("click", async () => {
        await loadState();
        renderResourcesSheet();
      });
    }

    const close = document.getElementById("owResourcesCloseBtn");
    if (close && close.dataset.bound !== "1") {
      close.dataset.bound = "1";
      close.addEventListener("click", showPlanner);
    }
  }

  async function openResources() {
    ensureShells();
    hideAllOperational();
    document.querySelectorAll(".planner-main, .planner-side, .planner-bottom").forEach((node) => node.hidden = true);
    document.getElementById("oldWorkflowResourcesPort").hidden = false;
    await loadState();
    renderResourcesSheet();
  }

  function interceptTopbar() {
    document.querySelectorAll(".topbar-actions button").forEach((button) => {
      const text = norm(button.textContent || "");

      if (text === "GANTT" && button.dataset.oldWorkflowGanttPortBound !== "1") {
        button.dataset.oldWorkflowGanttPortBound = "1";
        button.addEventListener("click", (event) => {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          openGantt().catch((error) => {
            console.error(error);
            alert(error.message || "Errore apertura Gantt.");
          });
        }, true);
      }

      if (text === "RISORSE" && button.dataset.oldWorkflowResourcesPortBound !== "1") {
        button.dataset.oldWorkflowResourcesPortBound = "1";
        button.addEventListener("click", (event) => {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          openResources().catch((error) => {
            console.error(error);
            alert(error.message || "Errore apertura Risorse.");
          });
        }, true);
      }

      if (text === "PLANNER" && button.dataset.oldWorkflowPlannerReturnBound !== "1") {
        button.dataset.oldWorkflowPlannerReturnBound = "1";
        button.addEventListener("click", () => showPlanner(), true);
      }
    });
  }

  function disableProvisionalGantt() {
    const old = document.getElementById("ganttSheetV2");
    if (old) old.hidden = true;
  }

  window.openOldWorkflowGantt = openGantt;
  window.openOldWorkflowResources = openResources;

  ensureShells();
  disableProvisionalGantt();
  interceptTopbar();

  setTimeout(interceptTopbar, 500);
  setTimeout(interceptTopbar, 1500);
  setTimeout(interceptTopbar, 3000);
})();
// === OLD WORKFLOW GANTT RESOURCES FRONTEND END ===
