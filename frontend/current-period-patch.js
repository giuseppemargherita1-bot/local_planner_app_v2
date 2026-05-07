(function patchCurrentPlannerPeriod() {
  const FIRST_WEEK = 1;
  const LAST_WEEK = 52;

  function getIsoWeekInfo(date = new Date()) {
    const utcDate = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const day = utcDate.getUTCDay() || 7;
    utcDate.setUTCDate(utcDate.getUTCDate() + 4 - day);

    const isoYear = utcDate.getUTCFullYear();
    const yearStart = new Date(Date.UTC(isoYear, 0, 1));
    const week = Math.ceil((((utcDate - yearStart) / 86400000) + 1) / 7);

    return {
      yearFull: isoYear,
      yearShort: isoYear % 100,
      week,
      periodKey: (isoYear % 100) * 100 + week,
    };
  }

  function getCurrentPlannerPeriod() {
    return window.__plannerCurrentPeriod || getIsoWeekInfo();
  }

  function getCurrentPeriodKey() {
    return Number(getCurrentPlannerPeriod().periodKey || 0);
  }

  function getCurrentYearShort() {
    return Number(getCurrentPlannerPeriod().yearShort || 0);
  }

  function getCurrentWeek() {
    return Number(getCurrentPlannerPeriod().week || 0);
  }

  function normalizeText(value) {
    if (typeof normalizeRole === "function") {
      return normalizeRole(value);
    }

    return String(value || "").trim().toUpperCase();
  }

  function parseCellPayload(raw) {
    if (!raw) return null;

    try {
      return JSON.parse(decodeURIComponent(raw));
    } catch (error) {
      try {
        return JSON.parse(raw);
      } catch (innerError) {
        return null;
      }
    }
  }

  function getElementPeriodKey(element) {
    if (!element) return 0;

    const direct = Number(element.dataset.periodKey || 0);
    if (direct > 0) return direct;

    const weekValue = Number(element.dataset.week || 0);
    if (weekValue > 0) {
      return weekValue >= 1000 ? weekValue : periodKeyFromWeek(weekValue, getCurrentYearShort());
    }

    const cellData = parseCellPayload(element.dataset.cell || "");
    if (!cellData) return 0;

    const cellPeriodKey = Number(cellData.period_key || 0);
    if (cellPeriodKey > 0) return cellPeriodKey;

    const cellWeek = Number(cellData.week || 0);
    if (cellWeek > 0) {
      return periodKeyFromWeek(cellWeek, getCurrentYearShort());
    }

    return 0;
  }

  function syncCurrentWeekMarkers() {
    const currentWeek = getCurrentWeek();
    const currentPeriodKey = getCurrentPeriodKey();

    document.querySelectorAll(".week-head.current-week").forEach((element) => {
      element.classList.remove("current-week");
    });

    document.querySelectorAll(".week-head.second-line").forEach((element) => {
      const match = String(element.textContent || "").trim().toUpperCase().match(/^W(\d{1,2})$/);
      if (!match) return;
      if (Number(match[1]) === currentWeek) {
        element.classList.add("current-week");
      }
    });

    document.querySelectorAll("[data-period-key], [data-week], [data-cell]").forEach((element) => {
      const periodKey = getElementPeriodKey(element);
      if (!periodKey) return;

      element.classList.toggle("current-week", Number(periodKey) === currentPeriodKey);
    });
  }

  function patchGlobalFunction(name, factory) {
    if (typeof window[name] !== "function") return;

    const original = window[name];
    const patched = factory(original);

    window[name] = patched;

    try {
      globalThis[name] = patched;
    } catch (error) {
      // ignore global rebinding failures
    }
  }

  window.__plannerCurrentPeriod = getIsoWeekInfo();

  patchGlobalFunction("periodKeyFromWeek", () => function patchedPeriodKeyFromWeek(week, yearShort) {
    const numericWeek = Number(week || 0);
    if (numericWeek >= 1000) return numericWeek;

    const resolvedYearShort = Number(yearShort || getCurrentYearShort());
    return resolvedYearShort * 100 + numericWeek;
  });

  patchGlobalFunction("fullYearFromShort", () => function patchedFullYearFromShort(yearShort) {
    return 2000 + Number(yearShort || getCurrentYearShort());
  });

  patchGlobalFunction("isResourceContractEndedForPeriod", () => function patchedIsResourceContractEndedForPeriod(resource, periodKey) {
    if (!resource) {
      return true;
    }

    if (typeof isExternalResource === "function" && isExternalResource(resource)) {
      return false;
    }

    if (Number(resource.is_active) !== 1) {
      return true;
    }

    const noteText = normalizeText(resource.availability_note || "");
    if (
      noteText.includes("FUORI_CONTRATTO") ||
      noteText.includes("FUORI CONTRATTO") ||
      noteText.includes("CESSATO") ||
      noteText.includes("LICENZIATO")
    ) {
      return true;
    }

    const endDate = typeof getResourceEndDate === "function" ? getResourceEndDate(resource) : null;
    if (!endDate) {
      return false;
    }

    const numericPeriodKey = Number(periodKey || 0);
    const selectedWeek = typeof weekFromPeriodKey === "function"
      ? weekFromPeriodKey(numericPeriodKey)
      : (numericPeriodKey % 100);
    const selectedYearShort = numericPeriodKey >= 1000
      ? Math.floor(numericPeriodKey / 100)
      : getCurrentYearShort();
    const selectedDate = typeof getWeekStartDate === "function"
      ? getWeekStartDate(2000 + selectedYearShort, selectedWeek)
      : new Date(2000 + selectedYearShort, 0, 1);

    return endDate < selectedDate;
  });

  patchGlobalFunction("isResourceContractEndedForSelection", () => function patchedIsResourceContractEndedForSelection(resource) {
    const summary = typeof getSelectionSummary === "function" ? getSelectionSummary() : null;
    const selectedPeriodKey = Number(summary?.period_from || getCurrentPeriodKey());
    return isResourceContractEndedForPeriod(resource, selectedPeriodKey);
  });

  patchGlobalFunction("projectHasUsefulRows", () => function patchedProjectHasUsefulRows(projectId) {
    const project = typeof getProjectById === "function" ? getProjectById(projectId) : null;
    if (!project) return false;
    if (typeof isWorkshopChildProject === "function" && isWorkshopChildProject(project)) return false;

    const lastUseful = typeof getProjectLastUsefulPeriod === "function"
      ? getProjectLastUsefulPeriod(projectId)
      : 0;

    return lastUseful >= getCurrentPeriodKey() - 4;
  });

  patchGlobalFunction("rowHasUsefulFutureActivity", () => function patchedRowHasUsefulFutureActivity(projectId, role) {
    const normalizedRole = normalizeText(role);
    const periods = Array.isArray(PERIODS) ? PERIODS : [];
    const currentPeriodKey = getCurrentPeriodKey();

    return periods.some((period) => {
      if (Number(period.periodKey) < currentPeriodKey) {
        return false;
      }

      const demandFound = (demandsData || []).some((demand) => {
        return (
          Number(demand.project_id) === Number(projectId) &&
          normalizeText(demand.role || "") === normalizedRole &&
          normalizePeriodKey(demand) === Number(period.periodKey) &&
          Number(demand.quantity || 0) > 0
        );
      });

      if (demandFound) {
        return true;
      }

      return (allocationsData || []).some((allocation) => {
        if (
          Number(allocation.project_id) !== Number(projectId) ||
          normalizeText(allocation.role || "") !== normalizedRole ||
          normalizePeriodKey(allocation) !== Number(period.periodKey)
        ) {
          return false;
        }

        const resource = (resourcesData || []).find((item) => Number(item.id) === Number(allocation.resource_id));
        if (!resource) {
          return false;
        }

        return isResourceAvailableForPeriod(resource, period.periodKey);
      });
    });
  });

  patchGlobalFunction("rowHasRecentPastActivity", () => function patchedRowHasRecentPastActivity(projectId, role) {
    const normalizedRole = normalizeText(role);
    const currentPeriodKey = getCurrentPeriodKey();
    const recentPastStart = currentPeriodKey - 2;
    const periods = Array.isArray(PERIODS) ? PERIODS : [];

    return periods.some((period) => {
      if (Number(period.periodKey) < recentPastStart || Number(period.periodKey) >= currentPeriodKey) {
        return false;
      }

      const demandFound = (demandsData || []).some((demand) => {
        return (
          Number(demand.project_id) === Number(projectId) &&
          normalizeText(demand.role || "") === normalizedRole &&
          normalizePeriodKey(demand) === Number(period.periodKey) &&
          Number(demand.quantity || 0) > 0
        );
      });

      if (demandFound) {
        return true;
      }

      return (allocationsData || []).some((allocation) => {
        if (
          Number(allocation.project_id) !== Number(projectId) ||
          normalizeText(allocation.role || "") !== normalizedRole ||
          normalizePeriodKey(allocation) !== Number(period.periodKey)
        ) {
          return false;
        }

        const resource = (resourcesData || []).find((item) => Number(item.id) === Number(allocation.resource_id));
        if (!resource) {
          return false;
        }

        return isResourceAvailableForPeriod(resource, period.periodKey);
      });
    });
  });

  patchGlobalFunction("renderPlanner", (original) => function patchedRenderPlanner() {
    const result = original.apply(this, arguments);
    requestAnimationFrame(syncCurrentWeekMarkers);
    return result;
  });

  patchGlobalFunction("scrollToCurrentWeek", (original) => function patchedScrollToCurrentWeek() {
    syncCurrentWeekMarkers();
    return original.apply(this, arguments);
  });

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", syncCurrentWeekMarkers, { once: true });
  } else {
    syncCurrentWeekMarkers();
  }

  setTimeout(syncCurrentWeekMarkers, 150);
  setTimeout(syncCurrentWeekMarkers, 600);
})();
