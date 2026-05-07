(function patchPlannerPerformance() {
  const CACHEABLE_GET_PREFIXES = [
    "/api/resources",
    "/api/projects",
    "/api/workshop-required-map",
    "/api/allocation-history",
    "/api/demand-history",
  ];

  const INVALIDATION_RULES = [
    {
      prefixes: [
        "/api/demands",
      ],
      invalidate: [
        "/api/demands",
        "/api/demand-history",
      ],
    },
    {
      prefixes: [
        "/api/allocations",
      ],
      invalidate: [
        "/api/allocations",
        "/api/allocation-history",
      ],
    },
    {
      prefixes: [
        "/api/resources",
      ],
      invalidate: [
        "/api/resources",
        "/api/allocations",
        "/api/allocation-history",
      ],
    },
    {
      prefixes: [
        "/api/projects",
      ],
      invalidate: [
        "/api/projects",
        "/api/demands",
        "/api/allocations",
        "/api/workshop-required-map",
      ],
    },
  ];

  const responseCache = new Map();
  const inFlightRequests = new Map();

  function deepClone(value) {
    if (typeof structuredClone === "function") {
      return structuredClone(value);
    }

    return JSON.parse(JSON.stringify(value));
  }

  function normalizePath(path) {
    return String(path || "").trim();
  }

  function isCacheableGet(path) {
    return CACHEABLE_GET_PREFIXES.some((prefix) => normalizePath(path).startsWith(prefix));
  }

  function invalidateByPrefixes(prefixes) {
    if (!Array.isArray(prefixes) || !prefixes.length) return;

    Array.from(responseCache.keys()).forEach((key) => {
      if (prefixes.some((prefix) => key.startsWith(prefix))) {
        responseCache.delete(key);
      }
    });
  }

  function invalidateForMutation(path) {
    const normalizedPath = normalizePath(path);

    INVALIDATION_RULES.forEach((rule) => {
      if (rule.prefixes.some((prefix) => normalizedPath.startsWith(prefix))) {
        invalidateByPrefixes(rule.invalidate);
      }
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
      // ignore
    }
  }

  patchGlobalFunction("fetchJson", (originalFetchJson) => async function patchedFetchJson(path, options = {}) {
    const normalizedPath = normalizePath(path);
    const method = String(options?.method || "GET").toUpperCase();

    if (method === "GET" && isCacheableGet(normalizedPath)) {
      if (responseCache.has(normalizedPath)) {
        return deepClone(responseCache.get(normalizedPath));
      }

      if (inFlightRequests.has(normalizedPath)) {
        const sharedResult = await inFlightRequests.get(normalizedPath);
        return deepClone(sharedResult);
      }

      const pending = Promise.resolve(originalFetchJson.call(this, path, options))
        .then((result) => {
          responseCache.set(normalizedPath, deepClone(result));
          return result;
        })
        .finally(() => {
          inFlightRequests.delete(normalizedPath);
        });

      inFlightRequests.set(normalizedPath, pending);
      const result = await pending;
      return deepClone(result);
    }

    const result = await originalFetchJson.call(this, path, options);

    if (method !== "GET") {
      invalidateForMutation(normalizedPath);
    }

    return result;
  });

  patchGlobalFunction("loadAll", (originalLoadAll) => {
    let inFlightLoad = null;
    let rerunRequested = false;

    return async function patchedLoadAll() {
      if (inFlightLoad) {
        rerunRequested = true;
        return inFlightLoad;
      }

      const run = async () => {
        try {
          return await originalLoadAll.apply(this, arguments);
        } finally {
          inFlightLoad = null;
          if (rerunRequested) {
            rerunRequested = false;
            void patchedLoadAll.apply(this, arguments);
          }
        }
      };

      inFlightLoad = run();
      return inFlightLoad;
    };
  });

  window.__plannerPerformanceCache = {
    clear() {
      responseCache.clear();
      inFlightRequests.clear();
    },
  };
})();
