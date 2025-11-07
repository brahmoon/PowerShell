const STORAGE_KEY = 'psBrowserPilot.graph.autosave';
const MAX_STORAGE_SIZE = 5 * 1024 * 1024; // 5MB safety limit

const safeStringify = (value) => {
  try {
    return JSON.stringify(value);
  } catch (error) {
    console.warn('Failed to stringify graph for autosave', error);
    return null;
  }
};

export function saveLocalGraph(graph) {
  if (typeof window === 'undefined' || !window.localStorage) {
    return false;
  }
  const payload = {
    version: 1,
    updatedAt: Date.now(),
    graph,
  };
  const serialized = safeStringify(payload);
  if (!serialized) {
    return false;
  }
  if (serialized.length > MAX_STORAGE_SIZE) {
    console.warn('Autosave payload is too large and was skipped.');
    return false;
  }
  try {
    window.localStorage.setItem(STORAGE_KEY, serialized);
    return true;
  } catch (error) {
    console.warn('Failed to persist autosaved graph', error);
    return false;
  }
}

export function loadLocalGraph() {
  if (typeof window === 'undefined' || !window.localStorage) {
    return null;
  }
  try {
    const stored = window.localStorage.getItem(STORAGE_KEY);
    if (!stored) {
      return null;
    }
    const payload = JSON.parse(stored);
    if (payload && typeof payload === 'object' && payload.graph) {
      return payload.graph;
    }
    return null;
  } catch (error) {
    console.warn('Failed to restore autosaved graph', error);
    return null;
  }
}

export function clearLocalGraph() {
  if (typeof window === 'undefined' || !window.localStorage) {
    return;
  }
  try {
    window.localStorage.removeItem(STORAGE_KEY);
  } catch (error) {
    console.warn('Failed to clear autosaved graph', error);
  }
}
