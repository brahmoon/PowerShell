import { DEFAULT_SERVER_URL, normalizeServerUrl } from '../export/scriptRunner.js';

const NODE_ENDPOINTS = {
  list: '/nodes/list',
  save: '/nodes/save',
  delete: '/nodes/delete',
};

const DEFAULT_CONSTANT_PLACEHOLDER = '# TODO: set value';
const DEFAULT_UI_MARKUP = [
  '<div class="custom-ui-node">',
  '  <!-- Build your custom node UI here -->',
  '</div>',
].join('\n');
const DEFAULT_UI_SCRIPT = [
  "// The context argument provides helpers like updateConfig, resolveInput, and toPowerShellLiteral.",
  "// Update outputs by calling updateConfig('Result', toPowerShellLiteral('value')).",
  'const { controls, updateConfig, toPowerShellLiteral } = context;',
  'if (controls) {',
  "  const button = controls.querySelector('button');",
  '  const status = controls.querySelector("[data-status]");',
  '  if (button && status) {',
  "    let active = context.node?.config?.Result__raw === 'true';",
  '    const apply = (value) => {',
  '      active = value;',
  "      status.textContent = value ? 'ON' : 'OFF';",
  "      button.textContent = value ? 'Turn OFF' : 'Turn ON';",
  "      updateConfig('Result__raw', value ? 'true' : 'false', { silent: true });",
  "      updateConfig('Result', toPowerShellLiteral(value ? 'true' : 'false'));",
  '    };',
  '    apply(active);',
  "    button.addEventListener('click', () => apply(!active));",
  '    return () => {',
  "      button.replaceWith(button.cloneNode(true));",
  '    };',
  '  }',
  '}',
].join('\n');
const DEFAULT_UI_STYLE = [
  '/* Styles applied to the custom UI node sample */',
  '.custom-ui-node {',
  '  display: inline-flex;',
  '  align-items: center;',
  '  gap: 0.75rem;',
  '}',
  '.custom-ui-node button {',
  '  padding: 0.25rem 0.75rem;',
  '}',
  '.custom-ui-node [data-status] {',
  '  font-weight: 600;',
  '}',
].join('\n');
const CONTROL_TYPES = new Map(
  [
    ['text', 'TextBox'],
    ['textbox', 'TextBox'],
    ['text-box', 'TextBox'],
    ['reference', 'Reference'],
    ['file', 'Reference'],
    ['checkbox', 'CheckBox'],
    ['check-box', 'CheckBox'],
    ['radiobutton', 'RadioButton'],
    ['radio', 'RadioButton'],
    ['select', 'SelectBox'],
    ['selectbox', 'SelectBox'],
    ['select-box', 'SelectBox'],
  ].map(([key, value]) => [key, value])
);

const normalizeControlType = (value) => {
  const normalized = String(value ?? '').trim().toLowerCase();
  return CONTROL_TYPES.get(normalized) || 'TextBox';
};

const normalizeBooleanConstant = (value) =>
  /^(true|1|yes|on)$/i.test(String(value ?? '').trim()) ? 'True' : 'False';

const PLACEHOLDER_PATTERN = /\{\{\s*(input|output|config)\.([A-Za-z0-9_]+)\s*\}\}/g;
const CONFIG_INPUT_ASSIGN_PATTERN =
  /\{\{\s*config\.([A-Za-z0-9_]+)\s*\}\}\s*=\s*\{\{\s*input\.([A-Za-z0-9_]+)\s*\}\}/g;

const sanitizeId = (value) =>
  String(value ?? '')
    .trim()
    .replace(/\s+/g, '_')
    .replace(/[^A-Za-z0-9_]/g, '_');

const toBindingKey = (value) => sanitizeId(value).toLowerCase();

const normalizeExecutionMode = (value) =>
  String(value ?? '')
    .trim()
    .toLowerCase() === 'ui'
    ? 'ui'
    : 'powershell';

const normalizeUiSpec = (ui) => {
  if (!ui || typeof ui !== 'object') {
    return { markup: '', script: '', style: '' };
  }
  const normalizeText = (value) =>
    typeof value === 'string' ? value.replace(/\r\n/g, '\n') : '';
  return {
    markup: normalizeText(ui.markup),
    script: normalizeText(ui.script),
    style: normalizeText(ui.style),
  };
};

const normalizeList = (value) => {
  if (Array.isArray(value)) {
    return value
      .map((item) => String(item ?? '').trim())
      .filter((item, index, self) => item && self.indexOf(item) === index);
  }
  if (typeof value === 'string') {
    return value
      .split(/\r?\n/)
      .map((item) => String(item ?? '').trim())
      .filter((item, index, self) => item && self.indexOf(item) === index);
  }
  return [];
};

const normalizeConstants = (constants) => {
  if (!Array.isArray(constants)) {
    return [];
  }
  const byKey = new Map();
  constants.forEach((constant) => {
    if (!constant) return;
    const rawKey = typeof constant.key === 'string' ? constant.key : constant.id;
    const key = sanitizeId(rawKey);
    if (!key) return;
    const type = normalizeControlType(constant.type);
    const rawSource = Object.prototype.hasOwnProperty.call(constant, 'value')
      ? constant.value
      : constant.default;
    const toText = (value) => String(value ?? '').trim();
    const normalizedValue = toText(rawSource);
    if (type === 'SelectBox') {
      const candidates = [];
      if (Array.isArray(constant.options)) {
        constant.options.forEach((option) => {
          if (option && typeof option === 'object') {
            candidates.push(option.value);
          } else {
            candidates.push(option);
          }
        });
      }
      if (normalizedValue) {
        candidates.push(normalizedValue);
      }
      const options = candidates
        .map((option) => toText(option))
        .filter((option, index, self) => option && self.indexOf(option) === index);
      if (!options.length) {
        return;
      }
      const preferred = options.includes(normalizedValue) ? normalizedValue : options[0];
      if (byKey.has(key)) {
        const existing = byKey.get(key);
        if (existing.type !== 'SelectBox') {
          return;
        }
        options.forEach((option) => {
          if (!existing.options.includes(option)) {
            existing.options.push(option);
          }
        });
        if (!existing.value || !existing.options.includes(existing.value)) {
          existing.value = preferred;
        }
      } else {
        byKey.set(key, {
          key,
          type,
          value: preferred,
          options,
        });
      }
      return;
    }
    if (byKey.has(key)) {
      return;
    }
    let value;
    if (type === 'CheckBox' || type === 'RadioButton') {
      value = normalizeBooleanConstant(normalizedValue);
    } else {
      value = normalizedValue || DEFAULT_CONSTANT_PLACEHOLDER;
    }
    byKey.set(key, {
      key,
      type,
      value,
    });
  });
  return Array.from(byKey.values());
};

const normalizeSpec = (spec, previous = null) => {
  const baseId = sanitizeId(spec?.id || spec?.identifier || '');
  const label = String(spec?.label ?? '').trim();
  const category = String(spec?.category ?? '').trim();
  const now = new Date().toISOString();
  const normalized = {
    id: baseId || (label ? sanitizeId(label) : ''),
    label: label || 'Untitled node',
    category: category || 'Custom',
    execution: normalizeExecutionMode(spec?.execution),
    inputs: normalizeList(spec?.inputs),
    outputs: normalizeList(spec?.outputs),
    constants: normalizeConstants(spec?.constants),
    script: typeof spec?.script === 'string' ? spec.script.replace(/\r\n/g, '\n') : '',
    ui: normalizeUiSpec(spec?.ui),
    description: typeof spec?.description === 'string' ? spec.description : '',
    createdAt: previous?.createdAt || spec?.createdAt || now,
    updatedAt: spec?.updatedAt || previous?.updatedAt || now,
  };
  if (!normalized.id) {
    normalized.id = `custom_node_${Date.now()}`;
  }
  return normalized;
};

const deriveConfigInputBindings = (spec) => {
  const script = typeof spec?.script === 'string' ? spec.script : '';
  if (!script.trim()) {
    return new Map();
  }
  const bindings = new Map();
  const normalizedScript = script.replace(/\r\n/g, '\n');
  CONFIG_INPUT_ASSIGN_PATTERN.lastIndex = 0;
  let match;
  // Capture direct assignments of config placeholders from input placeholders so that
  // downstream TextBox controls can mirror connected input values inside the editor UI.
  while ((match = CONFIG_INPUT_ASSIGN_PATTERN.exec(normalizedScript))) {
    const configKey = toBindingKey(match[1]);
    const inputKey = toBindingKey(match[2]);
    if (!configKey || !inputKey || bindings.has(configKey)) {
      continue;
    }
    bindings.set(configKey, inputKey);
  }
  return bindings;
};

const ensureServerUrl = (value) => normalizeServerUrl(value) || DEFAULT_SERVER_URL;

const requestNodeServer = async (path, { method = 'GET', body, serverUrl } = {}) => {
  const base = ensureServerUrl(serverUrl);
  const url = `${base}${path}`;
  const init = { method: method || 'GET' };
  if (body !== undefined) {
    init.headers = { 'Content-Type': 'application/json; charset=utf-8' };
    init.body = JSON.stringify(body);
  }
  let response;
  try {
    response = await fetch(url, init);
  } catch (error) {
    throw new Error(`PowerShell サーバーへの接続に失敗しました: ${error.message}`);
  }

  let data = null;
  const text = await response.text();
  if (text) {
    try {
      data = JSON.parse(text);
    } catch (error) {
      throw new Error('サーバーから無効なレスポンスを受信しました。');
    }
  }

  if (!response.ok) {
    const message = data?.error || `HTTP ${response.status}`;
    throw new Error(message);
  }
  return data;
};

const createScriptFunction = (template) => {
  const normalized = typeof template === 'string' ? template.replace(/\r\n/g, '\n') : '';
  if (!normalized.trim()) {
    return () => '';
  }
  return ({ inputs = {}, outputs = {}, config = {} }) =>
    normalized.replace(PLACEHOLDER_PATTERN, (match, scope, key) => {
      const source = scope === 'input' ? inputs : scope === 'output' ? outputs : config;
      return Object.prototype.hasOwnProperty.call(source, key) ? source[key] : match;
    });
};

const createUiRenderer = (spec) => {
  const ui = spec?.ui || {};
  const markup = typeof ui.markup === 'string' ? ui.markup : '';
  const script = typeof ui.script === 'string' ? ui.script : '';
  const style = typeof ui.style === 'string' ? ui.style : '';

  const normalizedMarkup = markup.replace(/\r\n/g, '\n');
  const normalizedScript = script.replace(/\r\n/g, '\n');
  const normalizedStyle = style.replace(/\r\n/g, '\n');

  let compiled = null;
  if (normalizedScript.trim()) {
    try {
      compiled = new Function(
        'context',
        `'use strict';\n${normalizedScript}\n`
      );
    } catch (error) {
      console.error('Failed to compile custom UI node script', error);
    }
  }

  let styleElement = null;
  const ensureStylesheet = () => {
    if (!normalizedStyle.trim()) {
      return null;
    }
    if (styleElement && styleElement.isConnected) {
      return styleElement;
    }
    const doc = typeof document !== 'undefined' ? document : null;
    if (!doc) {
      return null;
    }
    styleElement = doc.createElement('style');
    styleElement.type = 'text/css';
    if (spec?.id) {
      styleElement.dataset.customNode = spec.id;
    }
    styleElement.textContent = normalizedStyle;
    doc.head?.appendChild(styleElement);
    return styleElement;
  };

  return (context) => {
    const doc = context?.controls?.ownerDocument || (typeof document !== 'undefined' ? document : null);
    if (context?.controls) {
      context.controls.innerHTML = normalizedMarkup;
    }
    const activeStyle = ensureStylesheet();
    if (!compiled) {
      return () => {};
    }
    let teardown = null;
    try {
      const result = compiled({
        ...context,
        markup: normalizedMarkup,
        styleElement: activeStyle,
        document: doc,
      });
      if (typeof result === 'function') {
        teardown = result;
      }
    } catch (error) {
      console.error('Failed to execute custom UI node script', error);
      if (context?.controls && doc) {
        const errorEl = doc.createElement('div');
        errorEl.className = 'custom-node-error';
        errorEl.textContent = `UI script error: ${error.message}`;
        context.controls.innerHTML = '';
        context.controls.appendChild(errorEl);
      }
    }
    return () => {
      if (typeof teardown === 'function') {
        try {
          teardown();
        } catch (disposeError) {
          console.warn('Failed to dispose custom UI node', disposeError);
        }
      }
    };
  };
};

export const SAMPLE_NODE_TEMPLATES = [
  {
    id: 'sample_log_message',
    label: 'Sample: Log Message',
    category: 'Samples',
    execution: 'powershell',
    description:
      'Demonstrates how to emit a message using a constant input and return the same value as an output.',
    inputs: [],
    outputs: ['LoggedMessage'],
    constants: [
      { type: 'TextBox', key: 'message', value: '"Hello from custom node"' },
    ],
    script: [
      '# This script writes a message and keeps a reference to it.',
      'Write-Host {{config.message}}',
      '{{output.LoggedMessage}} = {{config.message}}',
    ].join('\n'),
  },
  {
    id: 'sample_math_add',
    label: 'Sample: Sum Inputs',
    category: 'Samples',
    execution: 'powershell',
    description: 'Shows how to combine two incoming values and expose a calculated result.',
    inputs: ['FirstValue', 'SecondValue'],
    outputs: ['Total'],
    constants: [
      {
        type: 'RadioButton',
        key: 'castAsInt',
        value: 'False',
      },
    ],
    script: [
      '# Sample math operation',
      "if ({{config.castAsInt}} -eq '$true') {",
      '  $first = [int]({{input.FirstValue}})',
      '  $second = [int]({{input.SecondValue}})',
      '} else {',
      '  $first = {{input.FirstValue}}',
      '  $second = {{input.SecondValue}}',
      '}',
      '{{output.Total}} = $first + $second',
    ].join('\n'),
  },
  {
    id: 'sample_invoke_command',
    label: 'Sample: Invoke ScriptBlock',
    category: 'Samples',
    execution: 'powershell',
    description:
      'Executes a custom script block with parameters taken from inputs and constants, then exposes the result.',
    inputs: ['ScriptInput'],
    outputs: ['Result'],
    constants: [
      {
        type: 'TextBox',
        key: 'scriptBlock',
        value: '[ScriptBlock]::Create("param($value) $value")',
      },
    ],
    script: [
      '# Invoke a script block with one input argument',
      '$__sb = {{config.scriptBlock}}',
      '{{output.Result}} = $__sb.Invoke({{input.ScriptInput}})',
    ].join('\n'),
  },
  {
    id: 'sample_ui_toggle_output',
    label: 'Sample UI: Toggle Output',
    category: 'Samples',
    execution: 'ui',
    description: 'Example GUI node that toggles a boolean output from the designer.',
    inputs: [],
    outputs: ['IsEnabled'],
    constants: [],
    ui: {
      markup: [
        '<div class="custom-ui-node sample-toggle-node">',
        '  <button type="button" data-role="toggle">有効にする</button>',
        '  <span data-status>OFF</span>',
        '</div>',
      ].join('\n'),
      script: [
        'const { node, controls, updateConfig, toPowerShellLiteral } = context;',
        'if (!controls) {',
        '  return;',
        '}',
        'const button = controls.querySelector("[data-role=\"toggle\"]");',
        'const status = controls.querySelector("[data-status]");',
        'if (!button || !status) {',
        '  return;',
        '}',
        "let current = (node?.config?.IsEnabled__raw || '').toLowerCase() === 'true';",
        'const renderState = (value) => {',
        "  status.textContent = value ? 'ON' : 'OFF';",
        "  button.textContent = value ? '無効にする' : '有効にする';",
        "  updateConfig('IsEnabled__raw', value ? 'true' : 'false', { silent: true });",
        "  updateConfig('IsEnabled', toPowerShellLiteral(value ? 'true' : 'false'));",
        '};',
        'renderState(current);',
        'const handleClick = () => {',
        '  current = !current;',
        '  renderState(current);',
        '};',
        'button.addEventListener(\'click\', handleClick);',
        'return () => button.removeEventListener(\'click\', handleClick);',
      ].join('\n'),
      style: [
        '.sample-toggle-node {',
        '  display: inline-flex;',
        '  align-items: center;',
        '  gap: 0.75rem;',
        '  background: rgba(59, 130, 246, 0.12);',
        '  padding: 0.5rem 0.75rem;',
        '  border-radius: 0.75rem;',
        '}',
        '.sample-toggle-node button {',
        '  padding: 0.3rem 0.9rem;',
        '  border-radius: 999px;',
        '  border: none;',
        '  background: #2563eb;',
        '  color: #fff;',
        '  cursor: pointer;',
        '}',
        '.sample-toggle-node button:hover {',
        '  background: #1d4ed8;',
        '}',
        '.sample-toggle-node [data-status] {',
        '  font-weight: 600;',
        '  letter-spacing: 0.05em;',
        '}',
      ].join('\n'),
    },
  },
  {
    id: 'sample_ui_excel_workbook_selector',
    label: 'Sample UI: Excel Workbook Selector',
    category: 'Samples',
    execution: 'ui',
    description:
      'Demonstrates how to reproduce the built-in Excel workbook selector, including identical ports and PowerShell integration.',
    inputs: [],
    outputs: ['WorkbookName'],
    constants: [],
    initialConfig: {
      WorkbookName: '',
      WorkbookName__raw: '',
      selectedWorkbookLabel: '',
    },
    ui: {
      markup: [
        '<div class="excel-ui-node workbook-selector-node" data-role="root">',
        '  <div class="excel-ui-row">',
        '    <select class="excel-ui-select" data-role="workbook"></select>',
        '    <button type="button" class="excel-ui-button" data-role="refresh">更新</button>',
        '  </div>',
        '  <div class="excel-ui-status" data-role="status"></div>',
        '</div>',
      ].join('\n'),
      script: [
        "const { node, controls, updateConfig, toPowerShellLiteral } = context;",
        "if (!controls) return;",
        "const select = controls.querySelector('[data-role=\"workbook\"]');",
        "const refreshBtn = controls.querySelector('[data-role=\"refresh\"]');",
        "const status = controls.querySelector('[data-role=\"status\"]');",
        "const DEFAULT_SERVER_URL = 'http://127.0.0.1:8787';",
        "const RUN_SCRIPT_PATH = '/runScript';",

        "// --- Utility ---",
        "const normalizeServerUrl = (value) => {",
        "  if (!value) return '';",
        "  let url = String(value).trim();",
        "  if (!/^https?:\\/\\//i.test(url)) url = `http://${url}`;",
        "  return url.replace(/\\/+$/, '');",
        "};",

        "const loadServerUrl = () => {",
        "  try {",
        "    const stored = window?.localStorage?.getItem('nodeflow.psServerUrl');",
        "    if (stored) return normalizeServerUrl(stored) || DEFAULT_SERVER_URL;",
        "  } catch (err) {",
        "    console.warn('Failed to read stored server URL', err);",
        "  }",
        "  return DEFAULT_SERVER_URL;",
        "};",

        "const getRunScriptEndpoint = () => {",
        "  const base = normalizeServerUrl(loadServerUrl()) || DEFAULT_SERVER_URL;",
        "  return `${base}${RUN_SCRIPT_PATH}`;",
        "};",

        "// --- Fetch wrapper ---",
        "const invokePowerShell = async (script) => {",
        "  const endpoint = getRunScriptEndpoint();",
        "  const response = await fetch(endpoint, {",
        "    method: 'POST',",
        "    headers: { 'Content-Type': 'application/json; charset=utf-8' },",
        "    body: JSON.stringify({ script }),",
        "  });",
        "  let payload;",
        "  try {",
        "    payload = await response.json();",
        "  } catch {",
        "    throw new Error('PowerShell サーバーから無効な応答を受信しました。');",
        "  }",
        "  if (!response.ok) {",
        "    const msg = payload?.error || `HTTP ${response.status}`;",
        "    throw new Error(msg);",
        "  }",
        "  if (payload?.ok === false) {",
        "    const msg = Array.isArray(payload?.errors) && payload.errors.length",
        "      ? payload.errors.join('\\n')",
        "      : payload?.error || 'PowerShell がエラーを返しました。';",
        "    throw new Error(msg);",
        "  }",
        "  if (typeof payload?.output === 'string') return payload.output.trim();",
        "  if (typeof payload === 'string') return payload.trim();",
        "  return '';",
        "};",

        "// --- JSON helper (robust parser) ---",
        "const requestExcelJson = async (script) => {",
        "  const output = await invokePowerShell(script);",
        "  if (!output) return null;",
        "  const clean = output.replace(/^[\\uFEFF\\s\\r\\n]+/, '').replace(/[\\s\\r\\n]+$/, '');",
        "  const start = clean.indexOf('{');",
        "  const end = clean.lastIndexOf('}');",
        "  if (start === -1 || end === -1 || end <= start) {",
        "    console.warn('JSON not found in PowerShell output:', clean);",
        "    throw new Error('PowerShell の応答を解析できませんでした。');",
        "  }",
        "  const jsonText = clean.slice(start, end + 1);",
        "  try {",
        "    return JSON.parse(jsonText);",
        "  } catch (error) {",
        "    console.error('JSON parse failed:', error, jsonText);",
        "    throw new Error('PowerShell の応答を解析できませんでした。');",
        "  }",
        "};",

        "// --- PowerShell Script to list Excel workbooks ---",
        "const LIST_WORKBOOKS_SCRIPT = [",
        "  \"$ErrorActionPreference = 'Stop'\",",
        "  '$excel = $null',",
        "  '$workbook = $null',",
        "  '$sheet = $null',",
        "  '$workbooks = $null',",
        "  '$result = $null',",
        "  '',",
        "  'try {',",
        "  '  try { $excel = [Runtime.Interopservices.Marshal]::GetActiveObject(\"Excel.Application\") } catch { $excel = $null }',",
        "  '  if (-not $excel) {',",
        "  '    $result = [pscustomobject]@{ ok = $false; error = \"Excel が見つかりません。\"; workbooks = @() }',",
        "  '  } else {',",
        "  '    $names = @()',",
        "  '    foreach ($wb in @($excel.Workbooks)) { if ($null -ne $wb) { $names += [string]$wb.Name } }',",
        "  '    $result = [pscustomobject]@{ ok = $true; workbooks = $names }',",
        "  '  }',",
        "  '} catch {',",
        "  '  $result = [pscustomobject]@{ ok = $false; error = $_.Exception.Message; workbooks = @() }',",
        "  '} finally {',",
        "  '  if ($workbooks -ne $null) {',",
        "  '    foreach ($wb in @($workbooks)) { if ($null -ne $wb) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null } }',",
        "  '    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbooks) | Out-Null',",
        "  '  }',",
        "  '  if ($sheet -ne $null) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null }',",
        "  '  if ($workbook -ne $null) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null }',",
        "  '  if ($excel -ne $null) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }',",
        "  '  [GC]::Collect(); [GC]::WaitForPendingFinalizers()',",
        "  '}',",
        "  '$result | ConvertTo-Json -Compress'",
        "].join('\\n');",

        "// --- UI binding ---",
        "if (!select || !refreshBtn || !status) return;",
        "let disposed = false;",

        "const setStatus = (msg, state = '') => {",
        "  if (disposed) return;",
        "  status.textContent = msg || '';",
        "  status.dataset.state = state;",
        "};",

        "const applySelection = (workbookName) => {",
        "  const raw = workbookName || '';",
        "  const literal = raw ? toPowerShellLiteral(raw) : '';",
        "  updateConfig('WorkbookName__raw', raw, { silent: true });",
        "  updateConfig('WorkbookName', literal, { silent: true });",
        "  updateConfig('selectedWorkbookLabel', raw, { silent: false });",
        "};",

        "const populateOptions = (items, { preserveSelection = true } = {}) => {",
        "  select.innerHTML = '';",
        "  const placeholder = document.createElement('option');",
        "  placeholder.value = '';",
        "  placeholder.textContent = items.length ? 'ブックを選択' : 'ブックが見つかりません';",
        "  placeholder.disabled = true;",
        "  placeholder.selected = true;",
        "  select.appendChild(placeholder);",
        "  items.forEach((item) => {",
        "    const opt = document.createElement('option');",
        "    opt.value = item;",
        "    opt.textContent = item;",
        "    select.appendChild(opt);",
        "  });",
        "  if (preserveSelection) {",
        "    const stored = node?.config?.WorkbookName__raw || node?.config?.selectedWorkbookLabel;",
        "    if (stored && items.includes(stored)) select.value = stored;",
        "  }",
        "  select.disabled = !items.length;",
        "};",

        "const fetchWorkbooks = async () => {",
        "  setStatus('Excelのブックを取得中…', 'pending');",
        "  refreshBtn.disabled = true;",
        "  try {",
        "    const data = await requestExcelJson(LIST_WORKBOOKS_SCRIPT);",
        "    if (disposed) return;",
        "    if (!data?.ok) throw new Error(data?.error || 'ブック情報を取得できませんでした');",
        "    const workbooks = Array.isArray(data.workbooks) ? data.workbooks : [];",
        "    populateOptions(workbooks);",
        "    if (workbooks.length) setStatus(`${workbooks.length} 件のブックを検出しました。`, 'success');",
        "    else setStatus('開いているブックが見つかりません。', 'info');",
        "  } catch (err) {",
        "    console.error('Workbook fetch failed:', err);",
        "    populateOptions([], { preserveSelection: false });",
        "    setStatus(`取得に失敗しました: ${err.message}`, 'error');",
        "  } finally {",
        "    if (!disposed) refreshBtn.disabled = false;",
        "  }",
        "};",

        "const handleChange = (e) => {",
        "  const value = e.target.value || '';",
        "  applySelection(value);",
        "  setStatus(value ? `${value} を選択しました。` : '', value ? 'success' : 'info');",
        "};",

        "const handleRefresh = (e) => {",
        "  e.preventDefault();",
        "  fetchWorkbooks();",
        "};",

        "select.addEventListener('change', handleChange);",
        "refreshBtn.addEventListener('click', handleRefresh);",
        "populateOptions([], { preserveSelection: true });",
        "const stored = node?.config?.selectedWorkbookLabel || node?.config?.WorkbookName__raw || '';",
        "if (stored) setStatus(`${stored} を再選択できます。`, 'info');",
        "fetchWorkbooks();",

        "return () => {",
        "  disposed = true;",
        "  select.removeEventListener('change', handleChange);",
        "  refreshBtn.removeEventListener('click', handleRefresh);",
        "};"
      ].join('\n'),
      style: [
        '.excel-ui-node {',
        '  display: grid;',
        '  gap: 0.6rem;',
        '  font-size: 0.9rem;',
        '}',
        '.excel-ui-row {',
        '  display: flex;',
        '  align-items: center;',
        '  gap: 0.5rem;',
        '}',
        '.excel-ui-select {',
        '  flex: 1;',
        '  border: 1px solid var(--border);',
        '  border-radius: 8px;',
        '  padding: 0.45rem 0.6rem;',
        '  background: var(--panel-subtle);',
        '  font-size: 0.9rem;',
        '  color: var(--text);',
        '}',
        '.excel-ui-button {',
        '  border: none;',
        '  border-radius: 999px;',
        '  padding: 0.45rem 0.9rem;',
        '  font-weight: 600;',
        '  cursor: pointer;',
        '  transition: background 0.2s ease, transform 0.2s ease, box-shadow 0.2s ease;',
        '  background: rgba(37, 99, 235, 0.18);',
        '  color: #1d4ed8;',
        '}',
        '.excel-ui-button:hover:not(:disabled),',
        '.excel-ui-button:focus-visible:not(:disabled) {',
        '  transform: translateY(-1px);',
        '  box-shadow: 0 6px 18px rgba(37, 99, 235, 0.28);',
        '}',
        '.excel-ui-button:disabled {',
        '  opacity: 0.55;',
        '  cursor: not-allowed;',
        '  transform: none;',
        '  box-shadow: none;',
        '}',
        '.excel-ui-status {',
        '  min-height: 1.1em;',
        '  font-size: 0.8rem;',
        '  color: var(--muted);',
        '}',
        ".excel-ui-status[data-state='success'] {",
        '  color: #1f9d72;',
        '}',
        ".excel-ui-status[data-state='error'] {",
        '  color: #dc2626;',
        '}',
        ".excel-ui-status[data-state='pending'] {",
        '  color: #2563eb;',
        '}',
      ].join('\n'),
    },
  },
  {
    id: 'sample_ui_excel_selection_inspector',
    label: 'Sample UI: Excel Selection Inspector',
    category: 'Samples',
    execution: 'ui',
    description:
      'Companion UI node that mirrors the built-in Excel selection inspector with five independent outputs.',
    inputs: ['WorkbookName'],
    outputs: ['SelectionAddress', 'Sheet', 'SelectionText', 'FontColorHex', 'InteriorColorHex'],
    constants: [],
    initialConfig: {
      WorkbookName__raw: '',
      SelectionAddress: '',
      SelectionAddress__raw: '',
      Sheet: '',
      Sheet__raw: '',
      SelectionText: '',
      SelectionText__raw: '',
      FontColorHex: '',
      FontColorHex__raw: '',
      InteriorColorHex: '',
      InteriorColorHex__raw: '',
      lastWorkbook: '',
      lastSheet: '',
    },
    ui: {
      markup: [
        '<div class="excel-ui-node selection-inspector-node" data-role="root">',
        '  <button type="button" class="excel-ui-button primary" data-role="fetch">選択セル情報を取得</button>',
        '  <div class="excel-ui-status" data-role="status"></div>',
        '  <dl class="excel-selection-result">',
        '    <dt>セル座標</dt><dd data-role="address">-</dd>',
        '    <dt>シート名</dt><dd data-role="sheet">-</dd>',
        '    <dt>セルのテキスト</dt><dd data-role="text">-</dd>',
        '    <dt>Font.Color</dt><dd data-role="font">-</dd>',
        '    <dt>Interior.Color</dt><dd data-role="interior">-</dd>',
        '  </dl>',
        '</div>',
      ].join('\n'),
      script: [
        "const { node, controls, updateConfig, resolveInput, toPowerShellLiteral } = context;",
        'if (!controls) {',
        '  return;',
        '}',
        "const button = controls.querySelector('[data-role=\"fetch\"]');",
        "const status = controls.querySelector('[data-role=\"status\"]');",
        "const addressEl = controls.querySelector('[data-role=\"address\"]');",
        "const sheetEl = controls.querySelector('[data-role=\"sheet\"]');",
        "const textEl = controls.querySelector('[data-role=\"text\"]');",
        "const fontEl = controls.querySelector('[data-role=\"font\"]');",
        "const interiorEl = controls.querySelector('[data-role=\"interior\"]');",
        'const DEFAULT_SERVER_URL = \"http://127.0.0.1:8787\";',
        "const RUN_SCRIPT_PATH = '/runScript';",
        'const normalizeServerUrl = (value) => {',
        '  if (!value) return \"\";',
        '  let url = String(value).trim();',
        '  if (!url) return \"\";',
        "  if (!/^https?:\\/\\//i.test(url)) {",
        "    url = `http://${url}`;",
        '  }',
        "  return url.replace(/\\/+$/, '');",
        '};',
        'const loadServerUrl = () => {',
        '  try {',
        "    const stored = window?.localStorage?.getItem('nodeflow.psServerUrl');",
        '    if (stored) {',
        '      const normalized = normalizeServerUrl(stored);',
        '      return normalized || DEFAULT_SERVER_URL;',
        '    }',
        '  } catch (error) {',
        "    console.warn('Failed to read stored PowerShell server URL', error);",
        '  }',
        '  return DEFAULT_SERVER_URL;',
        '};',
        'const getRunScriptEndpoint = () => {',
        '  const base = normalizeServerUrl(loadServerUrl()) || DEFAULT_SERVER_URL;',
        "  return `${base.replace(/\\/+$/, '')}${RUN_SCRIPT_PATH}`;",
        '};',
        'const invokePowerShell = async (script) => {',
        '  const endpoint = getRunScriptEndpoint();',
        '  const response = await fetch(endpoint, {',
        "    method: 'POST',",
        "    headers: { 'Content-Type': 'application/json; charset=utf-8' },",
        '    body: JSON.stringify({ script }),',
        '  });',
        '  let payload;',
        '  try {',
        '    payload = await response.json();',
        '  } catch (error) {',
        "    throw new Error('PowerShell サーバーから無効な応答を受信しました。');",
        '  }',
        '  if (!response.ok) {',
        '    const message = payload?.error || `HTTP ${response.status}`;',
        '    throw new Error(message);',
        '  }',
        '  if (payload?.ok === false) {',
        '    const message = Array.isArray(payload?.errors) && payload.errors.length',
        '      ? payload.errors.join("\\n")',
        "      : 'PowerShell がエラーを返しました。';",
        '    throw new Error(message);',
        '  }',
        "  return typeof payload?.output === 'string' ? payload.output.trim() : '';",
        '};',
        'const requestExcelJson = async (script) => {',
        '  const output = await invokePowerShell(script);',
        '  if (!output) {',
        '    return null;',
        '  }',
        '  try {',
        '    return JSON.parse(output);',
        '  } catch (error) {',
        "    throw new Error('PowerShell の応答を解析できませんでした。');",
        '  }',
        '};',
        'const COLOR_HELPERS = [',
        "  'function Convert-OleToHex {',",
        "  '  param([Parameter()][object]$Ole)',",
        "  '  if ($null -eq $Ole) { return $null }',",
        "  '  try {',",
        "  '    $color = [System.Drawing.ColorTranslator]::FromOle([int]$Ole)',",
        "  \"    return ('#{0:X2}{1:X2}{2:X2}' -f $color.R, $color.G, $color.B)\",",
        "  '  } catch {',",
        "  '    return $null',",
        "  '  }',",
        "  '}',",
        "].join('\\n');",
        'const RELEASE_SNIPPET = [',
        "  'if ($selection -ne $null) {',",
        "  '  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($selection) | Out-Null',",
        "  '}',",
        "  'if ($sheet -ne $null) {',",
        "  '  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null',",
        "  '}',",
        "  'if ($workbooks -ne $null) {',",
        "  '  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbooks) | Out-Null',",
        "  '}',",
        "  'if ($workbook -ne $null) {',",
        "  '  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null',",
        "  '}',",
        "  'if ($excel -ne $null) {',",
        "  '  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null',",
        "  '}',",
        "  '[GC]::Collect()',",
        "  '[GC]::WaitForPendingFinalizers()',",
        "].join('\\n');",
        'const buildSelectionInspectorScript = (workbookLiteral) => {',
        '  const targetLine = `    $targetName = ${workbookLiteral || "\'\'"}`;',
        '  return [',
        "    \"$ErrorActionPreference = 'Stop'\",",
        "    'Add-Type -AssemblyName System.Drawing',",
        "    '',",
        "    COLOR_HELPERS,",
        "    '',",
        "    '$excel = $null',",
        "    '$workbooks = $null',",
        "    '$workbook = $null',",
        "    '$sheet = $null',",
        "    '$selection = $null',",
        "    '$result = $null',",
        "    '',",
        "    'try {',",
        "    '  try {',",
        "    \"    $excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')\",",
        "    '  } catch {',",
        "    '    $excel = $null',",
        "    '  }',",
        "    '',",
        "    '  if (-not $excel) {',",
        "    '    $result = [pscustomobject]@{',",
        "    '      ok = $false',",
        "    \"      error = 'Excel が見つかりません。'\",",
        "    '    }',",
        "    '  } else {',",
        "    targetLine,",
        "    '    $workbooks = $excel.Workbooks',",
        "    '',",
        "    '    if ([string]::IsNullOrWhiteSpace($targetName)) {',",
        "    '      $workbook = $excel.ActiveWorkbook',",
        "    '    } else {',",
        "    '      foreach ($candidate in @($workbooks)) {',",
        "    '        if ($null -eq $candidate) { continue }',",
        "    '        if ($candidate.Name -eq $targetName) {',",
        "    '          $workbook = $candidate',",
        "    '          break',",
        "    '        }',",
        "    '        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($candidate) | Out-Null',",
        "    '      }',",
        "    '    }',",
        "    '',",
        "    '    if (-not $workbook) {',",
        "    '      if ([string]::IsNullOrWhiteSpace($targetName)) {',",
        "    '        $result = [pscustomobject]@{',",
        "    '          ok = $false',",
        "    \"          error = 'アクティブなブックが見つかりません。'\",",
        "    '        }',",
        "    '      } else {',",
        "    '        $result = [pscustomobject]@{',",
        "    '          ok = $false',",
        "    '          error = \"Workbook not found: $targetName\"',",
        "    '        }',",
        "    '      }',",
        "    '    } else {',",
        "    '      try {',",
        "    '        $selection = $workbook.Application.Selection',",
        "    '        if ($null -eq $selection) {',",
        "    '          $result = [pscustomobject]@{',",
        "    '            ok = $false',",
        "    \"            error = '選択範囲が見つかりません。'\",",
        "    '            workbook = $workbook.Name',",
        "    '          }',",
        "    '        } else {',",
        "    '          $sheet = $selection.Worksheet',",
        "    '          $result = [pscustomobject]@{',",
        "    '            ok = $true',",
        "    '            workbook = $workbook.Name',",
        "    '            sheet = $sheet.Name',",
        "    '            address = $selection.Address()',",
        "    '            value = $selection.Text',",
        "    '            font_color = @{',",
        "    '              ole = $selection.Font.Color',",
        "    '              hex = Convert-OleToHex $selection.Font.Color',",
        "    '            }',",
        "    '            interior_color = @{',",
        "    '              ole = $selection.Interior.Color',",
        "    '              hex = Convert-OleToHex $selection.Interior.Color',",
        "    '            }',",
        "    '          }',",
        "    '        }',",
        "    '      } catch {',",
        "    '        $result = [pscustomobject]@{',",
        "    '          ok = $false',",
        "    '          error = $_.Exception.Message',",
        "    '          workbook = $workbook.Name',",
        "    '        }',",
        "    '      }',",
        "    '    }',",
        "    '  }',",
        "    '} catch {',",
        "    '  $result = [pscustomobject]@{',",
        "    '    ok = $false',",
        "    '    error = $_.Exception.Message',",
        "    '  }',",
        "    '} finally {',",
        "    RELEASE_SNIPPET,",
        "    '}',",
        "    '',",
        "    'if ($result -eq $null) {',",
        "    '  $result = [pscustomobject]@{',",
        "    '    ok = $false',",
        "    \"    error = '結果を取得できませんでした。'\",",
        "    '  }',",
        "    '}',",
        "    '',",
        "    '$result | ConvertTo-Json -Compress',",
        '  ].join("\\n");',
        '};',
        'if (!button || !status || !addressEl || !textEl || !fontEl || !interiorEl) {',
        '  return;',
        '}',
        'let disposed = false;',
        'const setStatus = (message, state = \"\") => {',
        '  if (disposed) return;',
        '  status.textContent = message || \"\";',
        '  status.dataset.state = state;',
        '};',
        'const setOutput = (key, rawValue) => {',
        '  const raw = rawValue ?? \"\";',
        "  updateConfig(`${key}__raw`, raw, { silent: true });",
        "  updateConfig(key, raw ? toPowerShellLiteral(raw) : '', { silent: false });",
        '};',
        'const updateDisplay = (data) => {',
        "  addressEl.textContent = data.address || '-';",
        "  sheetEl.textContent = data.sheet || '-';",
        "  textEl.textContent = data.value || '-';",
        "  fontEl.textContent = data.fontHex || '-';",
        "  interiorEl.textContent = data.interiorHex || '-';",
        '};',
        'const applyStoredResult = () => {',
        '  updateDisplay({',
        "    address: node?.config?.SelectionAddress__raw,",
        "    sheet: node?.config?.Sheet__raw,",
        "    value: node?.config?.SelectionText__raw,",
        "    fontHex: node?.config?.FontColorHex__raw,",
        "    interiorHex: node?.config?.InteriorColorHex__raw,",
        '  });',
        '};',
        'const fetchSelection = async () => {',
        '  const workbook = (resolveInput ? resolveInput(\'WorkbookName\', { preferRaw: true }) : \"\") || node?.config?.WorkbookName__raw;',
        '  button.disabled = true;',
        "  setStatus('選択セル情報を取得中…', 'pending');",
        '  try {',
        '    const workbookLiteral = workbook ? toPowerShellLiteral(workbook) : null;',
        '    const script = buildSelectionInspectorScript(workbookLiteral);',
        '    const data = await requestExcelJson(script);',
        '    if (disposed) return;',
        '    if (!data?.ok) {',
        "      throw new Error(data?.error || '情報を取得できませんでした');",
        '    }',
        '    const address = data.address || \"\";',
        '    const value = data.value ?? \"\";',
        '    const fontHex = data.font_color?.hex || \"\";',
        '    const interiorHex = data.interior_color?.hex || \"\";',
        '    const resolvedWorkbook = data.workbook || workbook || \"\";',
        "    setOutput('SelectionAddress', address);",
        "    setOutput('Sheet', data.sheet || '');",
        "    setOutput('SelectionText', value);",
        "    setOutput('FontColorHex', fontHex);",
        "    setOutput('InteriorColorHex', interiorHex);",
        "    updateConfig('WorkbookName__raw', resolvedWorkbook, { silent: true });",
        "    updateConfig('WorkbookName', resolvedWorkbook ? toPowerShellLiteral(resolvedWorkbook) : '', { silent: true });",
        "    updateConfig('lastWorkbook', resolvedWorkbook, { silent: true });",
        "    updateConfig('lastSheet', data.sheet || '', { silent: true });",
        '    updateDisplay({ address, sheet: data.sheet || "", value, fontHex, interiorHex });',
        '    const workbookLabel = data.workbook || workbook;',
        '    const sheetLabel = data.sheet ? `${data.sheet}` : \"\";',
        '    const suffix = sheetLabel ? `${sheetLabel} - ${address}` : address;',
        '    if (workbookLabel && suffix) {',
        '      setStatus(`${workbookLabel}: ${suffix}`, \"success\");',
        '    } else if (suffix) {',
        "      setStatus(suffix, 'success');",
        '    } else {',
        "      setStatus('選択情報を更新しました。', 'success');",
        '    }',
        '  } catch (error) {',
        '    if (!disposed) {',
        "      setStatus(`取得に失敗しました: ${error.message}`, 'error');",
        '    }',
        '  } finally {',
        '    if (!disposed) {',
        '      button.disabled = false;',
        '    }',
        '  }',
        '};',
        'const handleClick = (event) => {',
        '  event.preventDefault();',
        '  fetchSelection();',
        '};',
        'button.addEventListener(\'click\', handleClick);',
        'applyStoredResult();',
        'if (node?.config?.lastWorkbook || node?.config?.lastSheet) {',
        "  const workbookLabel = node?.config?.lastWorkbook || '';",
        "  const sheetLabel = node?.config?.lastSheet || '';",
        "  const address = node?.config?.SelectionAddress__raw || '';",
        '  const summary = [workbookLabel, sheetLabel, address].filter(Boolean).join(" / ");',
        '  if (summary) {',
        "    setStatus(`前回の結果: ${summary}`, 'info');",
        '  }',
        '} else {',
        "  setStatus('Excelの選択範囲情報を取得できます。', 'info');",
        '}',
        'return () => {',
        '  disposed = true;',
        '  button.removeEventListener(\'click\', handleClick);',
        '};',
      ].join('\n'),
      style: [
        '.excel-ui-node {',
        '  display: grid;',
        '  gap: 0.6rem;',
        '  font-size: 0.9rem;',
        '}',
        '.excel-ui-button {',
        '  border: none;',
        '  border-radius: 999px;',
        '  padding: 0.45rem 0.9rem;',
        '  font-weight: 600;',
        '  cursor: pointer;',
        '  transition: background 0.2s ease, transform 0.2s ease, box-shadow 0.2s ease;',
        '  background: rgba(37, 99, 235, 0.18);',
        '  color: #1d4ed8;',
        '}',
        '.excel-ui-button.primary {',
        '  background: linear-gradient(135deg, #2563eb, #3b82f6);',
        '  color: #ffffff;',
        '}',
        '.excel-ui-button:hover:not(:disabled),',
        '.excel-ui-button:focus-visible:not(:disabled) {',
        '  transform: translateY(-1px);',
        '  box-shadow: 0 6px 18px rgba(37, 99, 235, 0.28);',
        '}',
        '.excel-ui-button:disabled {',
        '  opacity: 0.55;',
        '  cursor: not-allowed;',
        '  transform: none;',
        '  box-shadow: none;',
        '}',
        '.excel-ui-status {',
        '  min-height: 1.1em;',
        '  font-size: 0.8rem;',
        '  color: var(--muted);',
        '}',
        ".excel-ui-status[data-state='success'] {",
        '  color: #1f9d72;',
        '}',
        ".excel-ui-status[data-state='error'] {",
        '  color: #dc2626;',
        '}',
        ".excel-ui-status[data-state='pending'] {",
        '  color: #2563eb;',
        '}',
        '.excel-selection-result {',
        '  display: grid;',
        '  grid-template-columns: auto 1fr;',
        '  gap: 0.35rem 0.5rem;',
        '  font-size: 0.85rem;',
        '}',
        '.excel-selection-result dt {',
        '  margin: 0;',
        '  color: var(--muted);',
        '}',
        '.excel-selection-result dd {',
        '  margin: 0;',
        '  font-weight: 600;',
        '  color: var(--text);',
        '  word-break: break-word;',
        '}',
      ].join('\n'),
    },
  },
  {
    id: 'sample_excel_list_workbooks',
    label: 'Sample: Excel List Workbooks',
    category: 'Samples',
    execution: 'powershell',
    description:
      'PowerShell snippet used by the built-in Excel workbook selector to enumerate active workbooks.',
    inputs: [],
    outputs: ['ExcelResponseJson'],
    constants: [],
    script: [
      "$ErrorActionPreference = 'Stop'",
      '$excel = $null',
      '$workbooks = $null',
      '$result = $null',
      '',
      'try {',
      '  try {',
      "    $excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')",
      '  } catch {',
      '    $excel = $null',
      '  }',
      '',
      '  if (-not $excel) {',
      '    $result = [pscustomobject]@{',
      '      ok = $false',
      "      error = 'Excel が見つかりません。'",
      '      workbooks = @()',
      '    }',
      '  } else {',
      '    $workbooks = $excel.Workbooks',
      '    $names = @()',
      '    foreach ($wb in @($workbooks)) {',
      '      if ($null -ne $wb) {',
      '        $names += [string]$wb.Name',
      '      }',
      '    }',
      '    $result = [pscustomobject]@{',
      '      ok = $true',
      '      workbooks = $names',
      '    }',
      '  }',
      '} catch {',
      '  $result = [pscustomobject]@{',
      '    ok = $false',
      '    error = $_.Exception.Message',
      '    workbooks = @()',
      '  }',
      '} finally {',
      '  if ($workbooks -ne $null) {',
      '    foreach ($wb in @($workbooks)) {',
      '      if ($null -ne $wb) {',
      '        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null',
      '      }',
      '    }',
      '    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbooks) | Out-Null',
      '  }',
      '  if ($excel -ne $null) {',
      '    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null',
      '  }',
      '  [GC]::Collect()',
      '  [GC]::WaitForPendingFinalizers()',
      '}',
      '',
      'if ($result -eq $null) {',
      '  $result = [pscustomobject]@{',
      '    ok = $false',
      "    error = '結果を取得できませんでした。'",
      '    workbooks = @()',
      '  }',
      '}',
      '',
      '$json = $result | ConvertTo-Json -Compress',
      '{{output.ExcelResponseJson}} = $json',
    ].join('\n'),
  },
  {
    id: 'sample_excel_selection_info',
    label: 'Sample: Excel Selection Info',
    category: 'Samples',
    execution: 'powershell',
    description:
      'PowerShell script invoked by the Excel selection inspector, including the Add-Type call required for color parsing.',
    inputs: ['WorkbookName'],
    outputs: ['ResultJson'],
    constants: [],
    script: [
      "$ErrorActionPreference = 'Stop'",
      'Add-Type -AssemblyName System.Drawing',
      '',
      'function Convert-OleToHex {',
      '  param([Parameter()][object]$Ole)',
      '  if ($null -eq $Ole) { return $null }',
      '  try {',
      '    $color = [System.Drawing.ColorTranslator]::FromOle([int]$Ole)',
      "    return ('#{0:X2}{1:X2}{2:X2}' -f $color.R, $color.G, $color.B)",
      '  } catch {',
      '    return $null',
      '  }',
      '}',
      '',
      '$excel = $null',
      '$workbooks = $null',
      '$workbook = $null',
      '$sheet = $null',
      '$selection = $null',
      '$result = $null',
      '',
      'try {',
      '  try {',
      "    $excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')",
      '  } catch {',
      '    $excel = $null',
      '  }',
      '',
      '  if (-not $excel) {',
      '    $result = [pscustomobject]@{',
      '      ok = $false',
      "      error = 'Excel が見つかりません。'",
      '    }',
      '  } else {',
      '    $targetName = {{input.WorkbookName}}',
      '    $workbooks = $excel.Workbooks',
      '',
      '    if ([string]::IsNullOrWhiteSpace($targetName)) {',
      '      $workbook = $excel.ActiveWorkbook',
      '    } else {',
      '      foreach ($candidate in @($workbooks)) {',
      '        if ($null -eq $candidate) { continue }',
        '        if ($candidate.Name -eq $targetName) {',
      '          $workbook = $candidate',
      '          break',
      '        }',
      '        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($candidate) | Out-Null',
      '      }',
      '    }',
      '',
      '    if (-not $workbook) {',
      '      if ([string]::IsNullOrWhiteSpace($targetName)) {',
      '        $result = [pscustomobject]@{',
      '          ok = $false',
      "          error = 'アクティブなブックが見つかりません。'",
      '        }',
      '      } else {',
      '        $result = [pscustomobject]@{',
      '          ok = $false',
      '          error = "Workbook not found: $targetName"',
      '        }',
      '      }',
      '    } else {',
      '      try {',
      '        $selection = $workbook.Application.Selection',
      '        if ($null -eq $selection) {',
      '          $result = [pscustomobject]@{',
      '            ok = $false',
      "            error = '選択範囲が見つかりません。'",
      '            workbook = $workbook.Name',
      '          }',
      '        } else {',
      '          $sheet = $selection.Worksheet',
      '          $result = [pscustomobject]@{',
      '            ok = $true',
      '            workbook = $workbook.Name',
      '            sheet = $sheet.Name',
      '            address = $selection.Address()',
      '            value = $selection.Text',
      '            font_color = @{',
      '              ole = $selection.Font.Color',
      '              hex = Convert-OleToHex $selection.Font.Color',
      '            }',
      '            interior_color = @{',
      '              ole = $selection.Interior.Color',
      '              hex = Convert-OleToHex $selection.Interior.Color',
      '            }',
      '          }',
      '        }',
      '      } catch {',
      '        $result = [pscustomobject]@{',
      '          ok = $false',
      '          error = $_.Exception.Message',
      '          workbook = $workbook.Name',
      '        }',
      '      }',
      '    }',
      '  }',
      '} catch {',
      '  $result = [pscustomobject]@{',
      '    ok = $false',
      '    error = $_.Exception.Message',
      '  }',
      '} finally {',
      '  if ($selection -ne $null) {',
      '    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($selection) | Out-Null',
      '  }',
      '  if ($sheet -ne $null) {',
      '    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null',
      '  }',
      '  if ($workbooks -ne $null) {',
      '    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbooks) | Out-Null',
      '  }',
      '  if ($workbook -ne $null) {',
      '    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null',
      '  }',
      '  if ($excel -ne $null) {',
      '    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null',
      '  }',
      '  [GC]::Collect()',
      '  [GC]::WaitForPendingFinalizers()',
      '}',
      '',
      'if ($result -eq $null) {',
      '  $result = [pscustomobject]@{',
      '    ok = $false',
      "    error = '結果を取得できませんでした。'",
      '  }',
      '}',
      '',
      '$json = $result | ConvertTo-Json -Compress',
      '{{output.ResultJson}} = $json',
    ].join('\n'),
  },
];

export const createEmptySpec = (execution = 'powershell') => ({
  id: '',
  label: '',
  category: 'Custom',
  execution: execution === 'ui' ? 'ui' : 'powershell',
  inputs: [],
  outputs: [],
  constants: [
    { type: 'TextBox', key: 'note', value: '# TODO: describe behavior' },
  ],
  script: [
    '# Use {{input.Name}} to reference incoming values,',
    '# {{config.key}} for constant fields, and {{output.Result}} for outputs.',
    '# Remove these lines and write your PowerShell snippet here.',
  ].join('\n'),
  ui: {
    markup: DEFAULT_UI_MARKUP,
    script: DEFAULT_UI_SCRIPT,
    style: DEFAULT_UI_STYLE,
  },
});

export const listCustomNodeSpecs = async ({ serverUrl } = {}) => {
  const data = await requestNodeServer(NODE_ENDPOINTS.list, { method: 'GET', serverUrl });
  const specs = Array.isArray(data?.nodes) ? data.nodes : [];
  return specs
    .map((spec) => normalizeSpec(spec, spec))
    .filter((spec, index, self) => spec.id && self.findIndex((item) => item.id === spec.id) === index)
    .sort((a, b) => (a.label || '').localeCompare(b.label || ''));
};

export const saveCustomNodeSpec = async (spec, { serverUrl, previous } = {}) => {
  if (!spec || typeof spec !== 'object') {
    throw new Error('保存するノードの情報が無効です。');
  }
  const normalized = normalizeSpec(spec, previous || spec);
  normalized.updatedAt = new Date().toISOString();
  const data = await requestNodeServer(NODE_ENDPOINTS.save, {
    method: 'POST',
    serverUrl,
    body: { spec: normalized },
  });
  if (data?.spec) {
    return normalizeSpec(data.spec, normalized);
  }
  return normalized;
};

export const deleteCustomNodeSpec = async (id, { serverUrl } = {}) => {
  if (!id) {
    return [];
  }
  const data = await requestNodeServer(NODE_ENDPOINTS.delete, {
    method: 'POST',
    serverUrl,
    body: { id },
  });
  const specs = Array.isArray(data?.nodes) ? data.nodes : [];
  return specs
    .map((spec) => normalizeSpec(spec, spec))
    .filter((spec, index, self) => spec.id && self.findIndex((item) => item.id === spec.id) === index)
    .sort((a, b) => (a.label || '').localeCompare(b.label || ''));
};

export const specsToDefinitions = (specs) =>
  (specs || [])
    .map((spec) => {
      const normalized = normalizeSpec(spec, spec);
      const sanitizedInputMap = new Map();
      normalized.inputs.forEach((input) => {
        const bindingKey = toBindingKey(input);
        if (!bindingKey || sanitizedInputMap.has(bindingKey)) {
          return;
        }
        sanitizedInputMap.set(bindingKey, input);
      });
      const derivedBindings = deriveConfigInputBindings(normalized);
      const initialConfig = {};

      const definition = {
        id: normalized.id,
        label: normalized.label,
        category: normalized.category || 'Custom',
        execution: normalized.execution,
        inputs: normalized.inputs,
        outputs: normalized.outputs,
        controls: normalized.constants.map((constant) => {
          const controlKind = constant.type || 'TextBox';
          const optionValues =
            controlKind === 'SelectBox'
              ? Array.from(
                  new Set(
                    [
                      ...(Array.isArray(constant.options) ? constant.options : []),
                      ...(constant.value ? [constant.value] : []),
                    ].map((option) => String(option ?? '').trim())
                  )
                ).filter(Boolean)
                : [];
          let defaultValue;
          if (controlKind === 'CheckBox' || controlKind === 'RadioButton') {
            defaultValue = normalizeBooleanConstant(constant.value);
          } else if (controlKind === 'SelectBox') {
            const preferred = String(constant.value ?? '').trim();
            defaultValue = optionValues.includes(preferred) ? preferred : optionValues[0] || '';
          } else {
            defaultValue = constant.value;
          }
          let boundInputName;
          if (controlKind === 'TextBox') {
            const constantBindingKey = toBindingKey(constant.key);
            if (sanitizedInputMap.has(constantBindingKey)) {
              boundInputName = sanitizedInputMap.get(constantBindingKey);
            } else {
              const derived = derivedBindings.get(constantBindingKey);
              if (derived) {
                boundInputName = sanitizedInputMap.get(derived) || derived;
              }
            }
          }
          if (boundInputName) {
            initialConfig[`${constant.key}__raw`] = '';
          }
          return {
            key: constant.key,
            displayKey: constant.key,
            controlKind,
            type:
              controlKind === 'TextBox' || controlKind === 'Reference' ? 'text' : controlKind,
            placeholder:
              controlKind === 'TextBox' || controlKind === 'Reference'
                ? constant.value || DEFAULT_CONSTANT_PLACEHOLDER
                : undefined,
            default: defaultValue,
            options:
              controlKind === 'SelectBox'
                ? optionValues.map((value) => ({ value, label: value }))
                : undefined,
            bindsToInput: boundInputName,
          };
        }),
        script: createScriptFunction(normalized.script),
        specId: normalized.id,
        sourceSpec: normalized,
      };

      if (normalized.execution === 'ui') {
        definition.render = createUiRenderer(normalized);
      }

      if (Object.keys(initialConfig).length) {
        definition.initialConfig = initialConfig;
      }

      return definition;
    })
    .filter((definition, index, self) =>
      definition.id && self.findIndex((item) => item.id === definition.id) === index
    );

export const importSampleSpec = (sampleId) =>
  SAMPLE_NODE_TEMPLATES.find((template) => template.id === sampleId) || null;
