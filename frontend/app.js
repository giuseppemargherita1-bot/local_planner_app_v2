const root = document.documentElement;
const verticalSplitter = document.getElementById("verticalSplitter");
const horizontalSplitter = document.getElementById("horizontalSplitter");
const sideInnerSplitter = document.getElementById("sideInnerSplitter");
const plannerGridWrap = document.getElementById("plannerGridWrap");
const plannerBody = document.getElementById("plannerBody");

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

let rowMetaMap = new Map();
let selectedCells = [];
let activeMode = "demand";

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function startVerticalDrag(event) {
  verticalDrag = {
    startX: event.clientX,
    startWidth: parseInt(getComputedStyle(root).getPropertyValue("--side-w"), 10) || 430,
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
    startHeight: parseInt(getComputedStyle(root).getPropertyValue("--side-top-h"), 10) || 420,
  };
  if (sideInnerSplitter) sideInnerSplitter.classList.add("dragging");
  document.body.style.userSelect = "none";
}

function onPointerMove(event) {
  if (verticalDrag) {
    const delta = verticalDrag.startX - event.clientX;
    const nextWidth = clamp(verticalDrag.startWidth + delta, 260, 700);
    root.style.setProperty("--side-w", `${nextWidth}px`);
  }

  if (horizontalDrag) {
    const delta = horizontalDrag.startY - event.clientY;
    const nextHeight = clamp(horizontalDrag.startHeight + delta, 36, 180);
    root.style.setProperty("--bottom-h", `${nextHeight}px`);
  }

  if (sideInnerDrag) {
    const delta = event.clientY - sideInnerDrag.startY;
    const nextHeight = clamp(sideInnerDrag.startHeight + delta, 150, 620);
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
  const current = document.querySelector(".second-line.current-week");
  if (!current) return;

  const leftStickyWidth = 140 + 160 + 56;
  plannerGridWrap.scrollLeft = Math.max(0, current.offsetLeft - leftStickyWidth - 20);
}

function buildDemandMap(demands) {
  const map = new Map();
  for (const demand of demands) {
    const role = String(demand.role || "").trim().toUpperCase();
    const key = `${demand.project_id}__${role}__${demand.week}`;
    map.set(key, Number(demand.quantity || 0));
  }
  return map;
}

function buildAllocationMap(allocations, resources) {
  const resourceById = new Map(resources.map((r) => [r.id, r]));
  const map = new Map();

  for (const allocation of allocations) {
    const resource = resourceById.get(allocation.resource_id);
    if (!resource || !resource.is_active) continue;

    const note = String(resource.availability_note || "").toLowerCase();
    if (note.includes("indisp") || note.includes("indispon")) continue;

    const role = String(resource.role || "").trim().toUpperCase();
    const key = `${allocation.project_id}__${role}__${allocation.week}`;
    map.set(key, (map.get(key) || 0) + 1);
  }

  return map;
}

function buildPlannerRows(projects, demands) {
  const projectsById = new Map(projects.map((p) => [p.id, p]));
  const rowsMap = new Map();

  for (const demand of demands) {
    const project = projectsById.get(demand.project_id);
    if (!project) continue;

    const role = String(demand.role || "").trim().toUpperCase();
    const rowKey = `${demand.project_id}__${role}`;

    if (!rowsMap.has(rowKey)) {
      rowsMap.set(rowKey, {
        row_key: rowKey,
        project_id: demand.project_id,
        project_name: project.name,
        role,
      });
    }
  }

  const rows = Array.from(rowsMap.values()).sort((a, b) => {
    if (a.project_name !== b.project_name) {
      return a.project_name.localeCompare(b.project_name);
    }
    return a.role.localeCompare(b.role);
  });

  rowMetaMap = new Map(rows.map((row, idx) => [idx, row]));
  return rows;
}

function getCoverageClass(required, allocated) {
  if (required === 0) return "coverage good";
  if (allocated < required) return "coverage warn";
  if (allocated > required) return "coverage warn";
  return "coverage good";
}

function getCellClass(required, allocated, week) {
  let cls = "cell-empty planner-cell-clickable";

  if (required > 0 && allocated >= required) {
    cls = "cell-ok planner-cell-clickable";
  } else if (required > 0 && allocated < required) {
    cls = "cell-demand planner-cell-clickable";
  }

  if (week === CURRENT_WEEK) {
    cls += " current-col";
  }

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

  return {
    project_id: decoded.project_id,
    project_name: decoded.project_name,
    role: decoded.role,
    week_from: weeks[0],
    week_to: weeks[weeks.length - 1],
    required: selectedCells.length === 1 ? decoded.required : totalRequired,
    allocated: totalAllocated,
    diff: totalRequired - totalAllocated,
  };
}

function updateSidePanelFromSelection() {
  const summary = getSelectionSummary();
  if (!summary) return;

  selectionBox.textContent = `${summary.project_name} | ${summary.role} | W${String(summary.week_from).padStart(2, "0")}${summary.week_to !== summary.week_from ? ` - W${String(summary.week_to).padStart(2, "0")}` : ""}`;
  detailProject.value = summary.project_name;
  detailRole.value = summary.role;
  detailWeekFrom.value = summary.week_from;
  detailWeekTo.value = summary.week_to;
  detailRange.value = getWeekRangeLabel(summary.week_from, summary.week_to);
  detailRequired.value = summary.required;
  detailAllocated.value = summary.allocated;
  detailDiff.value = summary.required - summary.allocated;

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
    if (index === 0) {
      cell.classList.add("planner-cell-selected");
    } else {
      cell.classList.add("planner-cell-range");
    }
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

  if (direction === "right") {
    selectedCells.push(targetCell);
  } else {
    selectedCells.unshift(targetCell);
  }

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

function getAssignedResourcesForSelection() {
  const summary = getSelectionSummary();
  if (!summary) return [];

  const selectedWeeks = new Set();
  for (let w = summary.week_from; w <= summary.week_to; w += 1) {
    selectedWeeks.add(w);
  }

  return resourcesData.filter((resource) => {
    if (!resource.is_active) return false;

    const note = String(resource.availability_note || "").toLowerCase();
    if (note.includes("indisp") || note.includes("indispon")) return false;

    return allocationsData.some((allocation) => {
      return (
        allocation.resource_id === resource.id &&
        allocation.project_id === summary.project_id &&
        selectedWeeks.has(Number(allocation.week))
      );
    });
  });
}

function getResourceStatus(resource) {
  if (!resource.is_active) {
    return "inactive";
  }

  const note = String(resource.availability_note || "").toLowerCase();
  if (note.includes("indisp") || note.includes("indispon")) {
    return "unavailable";
  }

  const assigned = getAssignedResourcesForSelection().some((r) => r.id === resource.id);
  if (assigned) {
    return "allocated";
  }

  return "free";
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
  if (status === "inactive") classes.push("resource-inactive");

  return `
    <div class="${classes.join(" ")}" data-resource-id="${resource.id}">
      ${resource.name} | ${resource.role || "-"}${resource.availability_note ? ` | ${resource.availability_note}` : ""}
    </div>
  `;
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
  const assignedIds = new Set(assigned.map((r) => r.id));

  const search = resourceSearchInput?.value?.trim() || "";
  const showInactive = !!showInactiveToggle?.checked;

  const visibleAssigned = assigned
    .filter((resource) => matchesResourceSearch(resource, search));

  const visibleAvailable = resourcesData
    .filter((resource) => {
      if (assignedIds.has(resource.id)) return false;
      if (!showInactive && !resource.is_active) return false;
      return matchesResourceSearch(resource, search);
    });

  assignedResourceList.innerHTML = visibleAssigned.length
    ? visibleAssigned.map((resource) => renderResourceItem(resource, "allocated")).join("")
    : `<div class="resource-item resource-unavailable">Nessuna risorsa assegnata</div>`;

  availableResourceList.innerHTML = visibleAvailable.length
    ? visibleAvailable
        .map((resource) => renderResourceItem(resource, getResourceStatus(resource)))
        .join("")
    : `<div class="resource-item resource-unavailable">Nessuna risorsa disponibile</div>`;
}

function handleGridKeydown(event) {
  if (!selectedCells.length) return;

  switch (event.key) {
    case "ArrowRight":
      event.preventDefault();
      if (event.ctrlKey) {
        extendSelection("right");
      } else {
        moveSelection(0, 1);
      }
      break;
    case "ArrowLeft":
      event.preventDefault();
      if (event.ctrlKey) {
        extendSelection("left");
      } else {
        moveSelection(0, -1);
      }
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
      if (activeMode === "resources") {
        moveFocusToResourcesPanel();
      } else {
        moveFocusToDemandPanel();
      }
      break;
    case "Tab":
      event.preventDefault();
      if (activeMode === "demand") {
        setMode("resources");
      } else {
        setMode("demand");
      }
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
    project_id: summary.project_id,
    role: summary.role,
    week_from: Number(detailWeekFrom.value),
    week_to: Number(detailWeekTo.value),
    quantity: Number(detailRequired.value),
    note: "",
  };

  await fetchJson("/api/demands/upsert-range", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });

  await loadAndRenderPlanner();
  moveFocusToGrid();
}

function renderPlanner() {
  if (!plannerBody) return;

  const demandMap = buildDemandMap(demandsData);
  const allocationMap = buildAllocationMap(allocationsData, resourcesData);
  const rows = buildPlannerRows(projectsData, demandsData);

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

    const weekCells = WEEKS.map((week) => {
      const key = `${row.project_id}__${row.role}__${week}`;
      const required = demandMap.get(key) || 0;
      const allocated = allocationMap.get(key) || 0;
      const diff = required - allocated;

      totalRequired += required;
      totalAllocated += allocated;

      const payload = encodeURIComponent(JSON.stringify({
        project_id: row.project_id,
        project_name: row.project_name,
        role: row.role,
        week,
        required,
        allocated,
        diff,
      }));

      return `
        <td
          class="${getCellClass(required, allocated, week)}"
          data-cell='${payload}'
          data-row-index="${rowIndex}"
          data-week="${week}"
        >
          R${required}<br>
          A${allocated}<br>
          D${diff}
        </td>
      `;
    }).join("");

    const coverageClass = getCoverageClass(totalRequired, totalAllocated);
    const percent = totalRequired > 0
      ? Math.round((totalAllocated / totalRequired) * 100)
      : 100;

    return `
      <tr>
        <td class="sticky-col">${row.project_name}</td>
        <td class="sticky-col second">${row.role}</td>
        <td class="sticky-col third ${coverageClass}">
          ${totalAllocated}/${totalRequired}<br>
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
  });

  const firstCurrentCell = plannerBody.querySelector("td[data-cell].current-col");
  if (firstCurrentCell) {
    selectSingleCell(firstCurrentCell);
  }
}

async function loadAndRenderPlanner() {
  [projectsData, demandsData, allocationsData, resourcesData] = await Promise.all([
    fetchJson("/api/projects"),
    fetchJson("/api/demands"),
    fetchJson("/api/allocations"),
    fetchJson("/api/resources"),
  ]);

  renderPlanner();
  applyWeekTooltips();
  scrollToCurrentWeek();
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
      if (field) {
        field.addEventListener("keydown", handleDetailKeydown);
      }
    });

    if (modeDemandBtn) {
      modeDemandBtn.addEventListener("click", () => setMode("demand"));
    }

    if (modeResourcesBtn) {
      modeResourcesBtn.addEventListener("click", () => setMode("resources"));
    }

    if (resourceSearchInput) {
      resourceSearchInput.addEventListener("input", renderResourceLists);
    }

    if (showInactiveToggle) {
      showInactiveToggle.addEventListener("change", renderResourceLists);
    }

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