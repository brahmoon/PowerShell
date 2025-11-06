const hasOwn = (target, key) => Object.prototype.hasOwnProperty.call(target, key);

export const createStorageHelper = (context = {}) => {
  const { node, updateConfig } = context || {};

  const saveNodeConfig = (key, value, options = {}) => {
    if (typeof updateConfig === 'function') {
      updateConfig(key, value, options);
    } else if (node?.config) {
      node.config[key] = value;
    }
  };

  const loadNodeConfig = (key, defaultValue = '') => {
    if (!node?.config) {
      return defaultValue;
    }
    if (hasOwn(node.config, key)) {
      return node.config[key];
    }
    return defaultValue;
  };

  const saveJsonConfig = (key, data, options = {}) => {
    if (data === undefined || data === null) {
      saveNodeConfig(key, '', options);
      return;
    }
    try {
      const serialized = typeof data === 'string' ? data : JSON.stringify(data);
      saveNodeConfig(key, serialized, options);
    } catch (error) {
      console.warn('Failed to serialize node config JSON', error);
    }
  };

  const loadJsonConfig = (key, fallback = null) => {
    const raw = loadNodeConfig(key, '');
    if (typeof raw !== 'string' || !raw.trim()) {
      return fallback;
    }
    try {
      return JSON.parse(raw);
    } catch (error) {
      console.warn('Failed to parse node config JSON', error);
      return fallback;
    }
  };

  return {
    saveNodeConfig,
    loadNodeConfig,
    saveJsonConfig,
    loadJsonConfig,
  };
};

export const StorageHelper = { create: createStorageHelper };
