(function patchCurrentPlannerPeriod() {
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

  function periodKeyToParts(periodKey) {
    const numericPeriodKey = Number(periodKey || 0);
    if (!numericPeriodKey) {
      return {
        yearShort: getCurrentYearShort(),
        week: 0,
      };
    }

    if (numericPeriodKey >= 1000) {
      return {
        yearShort: Math.floor(numericPeriodKey / 100),
        week: numericPeriodKey % 100,
      };
    }

    return {
      yearShort: getCurrentYearShort(),
      week: numericPeriodKey,
    };
  }

  function getPeriodStartDate(periodKey) {
    const parts = periodKeyToParts(periodKey);
    if (!parts.week) {
      return null;
    }

    if (typeof getWeekStartDate === "function") {
      return getWeekStartDate(2000 + parts.yearShort, parts.week);
    }

    return new Date(2000 + parts.yearShort, 0, 1 + ((parts.week - 1) * 7));
  }

  function comparePeriods(leftPeriodKey, rightPeriodKey) {
    const leftDate = getPeriodStartDate(leftPeriodKey);
    const rightDate = getPeriodStartDate(rightPeriodKey);

    if (!leftDate || !rightDate) {
      return Number(leftPeriodKey || 0) - Number(rightPeriodKey || 0);
    }

    return leftDate.getTime() - rightDate.getTime();
  }

  function diffWeeksFromCurrent(periodKey) {
    const targetDate = getPeriodStartDate(periodKey);
    const currentDate = getPeriodStartDate(getCurrentPeriodKey());

    if (!targetDate || !currentDate) {
      return 0;
    }

    return Math.round((targetDate.getTime() - currentDate.getTime()) / 604800000);
  }

  function collectDataPeriodKeys() {
    const keys = new Set([getCurrentPeriodKey()]);

    [window.demandsData, window.allocationsData, window.allocationHistoryData, window.demandHistoryData].forEach((rows) => {
      if (!Array.isArray(rows)) return;

      rows.forEach((row) => {
        try {
          const periodKey = Number(typeof normalizePeriodKey === "function" ? normalizePeriodKey(row) : (row?.period_key || row?.week || 0));
          if (periodKey > 0) {
            keys.add(periodKey);
          }
        } catch (error) {
          // ignore malformed rows
        }
      });
    });

    return Array.from(keys).sort(comparePeriods);
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

  function findResource(resourceId) {
    return (window.resourcesData || []).find((item) => Number(item.id) === Number(resourceId)) || null;
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

    const selectedDate = getPeriodStartDate(periodKey);
    if (!selectedDate) {
      return false;
    }

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

    const lastUseful = typeof getProjectLastUsefulPeriod === "function" ? getProjectLastUsefulPeriod(projectId) : 0;
    return diffWeeksFromCurrent(lastUseful) >= -4;
  });

  patchGlobalFunction("rowHasUsefulFutureActivity", () => function patchedRowHasUsefulFutureActivity(projectId, role) {
    const normalizedRole = normalizeText(role);

    return collectDataPeriodKeys().some((periodKey) => {
      if (diffWeeksFromCurrent(periodKey) < 0) {
        return false;
      }

      const demandFound = (window.demandsData || []).some((demand) => {
        return (
          Number(demand.project_id) === Number(projectId) &&
          normalizeText(demand.role || "") === normalizedRole &&
          Number(normalizePeriodKey(demand)) === Number(periodKey) &&
          Number(demand.quantity || 0) > 0
        );
      });

      if (demandFound) {
        return true;
      }

      return (window.allocationsData || []).some((allocation) => {
        if (
          Number(allocation.project_id) !== Number(projectId) ||
          normalizeText(allocation.role || "") !== normalizedRole ||
          Number(normalizePeriodKey(allocation)) !== Number(periodKey)
        ) {
          return false;
        }

        const resource = findResource(allocation.resource_id);
        if (!resource) {
          return false;
        }

        return typeof isResourceAvailableForPeriod === "function"
          ? isResourceAvailableForPeriod(resource, periodKey)
          : true;
      });
    });
  });

  patchGlobalFunction("rowHasRecentPastActivity", () => function patchedRowHasRecentPastActivity(projectId, role) {
    const normalizedRole = normalizeText(role);

    return collectDataPeriodKeys().some((periodKey) => {
      const distance = diffWeeksFromCurrent(periodKey);
      if (distance < -2 || distance >= 0) {
        return false;
      }

      const demandFound = (window.demandsData || []).some((demand) => {
        return (
          Number(demand.project_id) === Number(projectId) &&
          normalizeText(demand.role || "") === normalizedRole &&
          Number(normalizePeriodKey(demand)) === Number(periodKey) &&
          Number(demand.quantity || 0) > 0
        );
      });

      if (demandFound) {
        return true;
      }

      return (window.allocationsData || []).some((allocation) => {
        if (
          Number(allocation.project_id) !== Number(projectId) ||
          normalizeText(allocation.role || "") !== normalizedRole ||
          Number(normalizePeriodKey(allocation)) !== Number(periodKey)
        ) {
          return false;
        }

        const resource = findResource(allocation.resource_id);
        if (!resource) {
          return false;
        }

        return typeof isResourceAvailableForPeriod === "function"
          ? isResourceAvailableForPeriod(resource, periodKey)
          : true;
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
