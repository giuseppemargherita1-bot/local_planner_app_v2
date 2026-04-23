const root = document.documentElement;
const verticalSplitter = document.getElementById("verticalSplitter");
const horizontalSplitter = document.getElementById("horizontalSplitter");
const sideInnerSplitter = document.getElementById("sideInnerSplitter");
const plannerGridWrap = document.getElementById("plannerGridWrap");

let verticalDrag = null;
let horizontalDrag = null;
let sideInnerDrag = null;

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function startVerticalDrag(event) {
  verticalDrag = {
    startX: event.clientX,
    startWidth: parseInt(getComputedStyle(root).getPropertyValue("--side-w"), 10) || 250,
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
    startHeight: parseInt(getComputedStyle(root).getPropertyValue("--side-top-h"), 10) || 210,
  };
  if (sideInnerSplitter) sideInnerSplitter.classList.add("dragging");
  document.body.style.userSelect = "none";
}

function onPointerMove(event) {
  if (verticalDrag) {
    const delta = verticalDrag.startX - event.clientX;
    const nextWidth = clamp(verticalDrag.startWidth + delta, 200, 420);
    root.style.setProperty("--side-w", `${nextWidth}px`);
  }

  if (horizontalDrag) {
    const delta = horizontalDrag.startY - event.clientY;
    const nextHeight = clamp(horizontalDrag.startHeight + delta, 36, 180);
    root.style.setProperty("--bottom-h", `${nextHeight}px`);
  }

  if (sideInnerDrag) {
    const delta = event.clientY - sideInnerDrag.startY;
    const nextHeight = clamp(sideInnerDrag.startHeight + delta, 80, 420);
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

function scrollToCurrentWeek() {
  if (!plannerGridWrap) return;
  const current = document.querySelector(".second-line.current-week");
  if (!current) return;

  const leftStickyWidth = 140 + 160 + 56;
  plannerGridWrap.scrollLeft = Math.max(0, current.offsetLeft - leftStickyWidth - 20);
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

function applyWeekTooltips() {
  const year = 2026;
  document.querySelectorAll(".week-head").forEach((cell) => {
    const text = (cell.textContent || "").trim().toUpperCase();
    const match = text.match(/^W(\d{1,2})$/);
    if (!match) return;

    const week = Number(match[1]);
    const start = getWeekStartDate(year, week);
    const end = new Date(start);
    end.setDate(start.getDate() + 6);

    cell.title = `${formatDate(start)} - ${formatDate(end)}`;
  });
}

document.addEventListener("DOMContentLoaded", () => {
  scrollToCurrentWeek();
  applyWeekTooltips();
});