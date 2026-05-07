(function patchFastPlannerEdits() {
  function patchGlobalFunction(name, factory) {
    if (typeof window[name] !== "function") return;

    const original = window[name];
    const patched = factory(original);
    window[name] = patched;

    try {
      globalThis[name] = patched;
    } catch (error) {
      // ignore
    }
  }

  function hasPrefix(value, prefixes) {
    return prefixes.some((prefix) => String(value || "").startsWith(prefix));
  }

  async function refreshPlannerDataLight() {
    if (typeof fetchJson !== "function") return;

    const tasks = [];

    if (typeof window.demandsData !== "undefined") {
      tasks.push(
        fetchJson("/api/demands").then((rows) => {
          window.demandsData = Array.isArray(rows) ? rows : [];
        })
      );
    }

    if (typeof window.allocationsData !== "undefined") {
      tasks.push(
        fetchJson("/api/allocations").then((rows) => {
          window.allocationsData = Array.isArray(rows) ? rows : [];
        })
      );
    }

    await Promise.all(tasks);

    if (typeof window.reloadOverallRollupAndRender === "function") {
      await window.reloadOverallRollupAndRender();
      return;
    }

    if (typeof window.renderPlanner === "function") {
      window.renderPlanner();
    }
  }

  patchGlobalFunction("fetchJson", (originalFetchJson) => async function patchedFetchJson(path, options = {}) {
    const method = String(options?.method || "GET").toUpperCase();
    const normalizedPath = String(path || "");

    const result = await originalFetchJson.call(this, path, options);

    if (
      method !== "GET" &&
      hasPrefix(normalizedPath, [
        "/api/demands",
        "/api/allocations",
      ])
    ) {
      Promise.resolve().then(() => refreshPlannerDataLight()).catch((error) => {
        console.error("Errore refresh planner leggero", error);
      });
    }

    return result;
  });

  patchGlobalFunction("loadAll", (originalLoadAll) => {
    let skipHeavyReloadOnce = false;

    const wrappedRefresh = async () => {
      skipHeavyReloadOnce = true;
      await refreshPlannerDataLight();
    };

    window.refreshPlannerDataLight = wrappedRefresh;

    return async function patchedLoadAll() {
      if (skipHeavyReloadOnce) {
        skipHeavyReloadOnce = false;
        return;
      }

      return originalLoadAll.apply(this, arguments);
    };
  });
})();
