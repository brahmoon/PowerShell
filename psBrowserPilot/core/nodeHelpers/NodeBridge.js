export const DEFAULT_SERVER_URL = 'http://127.0.0.1:8787';
export const RUN_SCRIPT_PATH = '/runScript';
const STORAGE_KEY = 'nodeflow.psServerUrl';

export const normalizeServerUrl = (value) => {
  if (!value) return '';
  let url = String(value).trim();
  if (!url) return '';
  if (!/^https?:\/\//i.test(url)) {
    url = `http://${url}`;
  }
  return url.replace(/\/+$/, '');
};

export const loadServerUrl = () => {
  try {
    const stored = window?.localStorage?.getItem(STORAGE_KEY);
    if (stored) {
      return normalizeServerUrl(stored) || DEFAULT_SERVER_URL;
    }
  } catch (error) {
    console.warn('Failed to read stored PowerShell server URL', error);
  }
  return DEFAULT_SERVER_URL;
};

export const saveServerUrl = (url) => {
  try {
    if (!url) {
      window?.localStorage?.removeItem(STORAGE_KEY);
      return;
    }
    window?.localStorage?.setItem(STORAGE_KEY, normalizeServerUrl(url));
  } catch (error) {
    console.warn('Failed to persist PowerShell server URL', error);
  }
};

const toTemplateString = (strings = [], values = []) => {
  let result = '';
  for (let i = 0; i < strings.length; i += 1) {
    result += strings[i];
    if (i < values.length) {
      result += values[i];
    }
  }
  return result;
};

export const coerceScriptDefinition = (definition, blockContext = {}) => {
  if (definition === undefined || definition === null) {
    return '';
  }

  if (typeof definition === 'function') {
    const result = definition(blockContext) ?? '';
    return coerceScriptDefinition(result, blockContext);
  }

  if (Array.isArray(definition)) {
    return definition
      .filter((line) => line !== undefined && line !== null)
      .map((line) => String(line))
      .join('\n');
  }

  if (typeof definition === 'object') {
    if (Array.isArray(definition.strings) && Array.isArray(definition.values)) {
      return toTemplateString(definition.strings, definition.values);
    }
    if (typeof definition.toString === 'function') {
      return definition.toString();
    }
  }

  return String(definition);
};

const resolveEndpoint = (serverUrl) => {
  const base = normalizeServerUrl(serverUrl) || normalizeServerUrl(loadServerUrl()) || DEFAULT_SERVER_URL;
  return `${base}${RUN_SCRIPT_PATH}`;
};

const withTimeout = async (promiseFactory, timeout) => {
  if (!timeout || typeof AbortController !== 'function') {
    return promiseFactory();
  }
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeout);
  try {
    return await promiseFactory(controller.signal);
  } finally {
    clearTimeout(timer);
  }
};

export const createNodeBridge = (context = {}, { logger } = {}) => {
  const literal = context?.toPowerShellLiteral || ((value) => value);
  const nodeLogger = logger || {
    debug: () => {},
    info: () => {},
    warn: console.warn.bind(console),
    error: console.error.bind(console),
  };

  const buildBlockContext = (args = {}) => ({
    literal,
    node: context?.node || null,
    config: context?.node?.config || {},
    updateConfig: context?.updateConfig,
    resolveInput: context?.resolveInput,
    ...args,
  });

  const run = async (definition, { server, timeout, args } = {}) => {
    const script = coerceScriptDefinition(definition, buildBlockContext(args));
    if (!script || !String(script).trim()) {
      throw new Error('PowerShell script is empty.');
    }

    const endpoint = resolveEndpoint(server);

    const response = await withTimeout(
      (signal) =>
        fetch(endpoint, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json; charset=utf-8' },
          body: JSON.stringify({ script }),
          signal,
        }),
      timeout
    );

    let payload;
    try {
      payload = await response.json();
    } catch (error) {
      nodeLogger.error('Failed to parse PowerShell response JSON', error);
      throw new Error('PowerShell サーバーから無効な応答を受信しました。');
    }

    if (!response.ok) {
      const message = payload?.error || `HTTP ${response.status}`;
      throw new Error(message);
    }

    return payload;
  };

  const requestJson = async (definition, options = {}) => {
    const payload = await run(definition, options);
    const output = typeof payload?.output === 'string' ? payload.output.trim() : '';
    if (!output) {
      return null;
    }
    try {
      return JSON.parse(output);
    } catch (error) {
      nodeLogger.error('Failed to parse JSON from PowerShell output', error, output);
      throw new Error('PowerShell の応答を解析できませんでした。');
    }
  };

  return {
    run,
    requestJson,
    coerceScript: (definition, args) => coerceScriptDefinition(definition, buildBlockContext(args)),
  };
};

export const NodeBridge = { create: createNodeBridge };

export const script = (strings, ...values) => ({ strings, values });
