(function patchPlannerRenderCoalescing() {
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

  patchGlobalFunction("renderPlanner", (originalRenderPlanner) => {
    let scheduled = false;
    let lastArgs = [];
    let lastThis = null;
    let lastResult;

    return function patchedRenderPlanner() {
      lastArgs = Array.from(arguments);
      lastThis = this;

      if (scheduled) {
        return lastResult;
      }

      scheduled = true;
      lastResult = Promise.resolve().then(() => new Promise((resolve) => {
        requestAnimationFrame(() => {
          scheduled = false;
          resolve(originalRenderPlanner.apply(lastThis, lastArgs));
        });
      }));

      return lastResult;
    };
  });
})();
