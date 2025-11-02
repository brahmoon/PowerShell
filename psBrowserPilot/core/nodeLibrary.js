const STORAGE_KEY = 'nodeflow.customNodes.v1';
const DEFAULT_CONSTANT_PLACEHOLDER = '# TODO: set value';
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

const sanitizeId = (value) =>
  String(value ?? '')
    .trim()
    .replace(/\s+/g, '_')
    .replace(/[^A-Za-z0-9_]/g, '_');

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
    inputs: normalizeList(spec?.inputs),
    outputs: normalizeList(spec?.outputs),
    constants: normalizeConstants(spec?.constants),
    script: typeof spec?.script === 'string' ? spec.script.replace(/\r\n/g, '\n') : '',
    description: typeof spec?.description === 'string' ? spec.description : '',
    createdAt: previous?.createdAt || spec?.createdAt || now,
    updatedAt: spec?.updatedAt || previous?.updatedAt || now,
  };
  if (!normalized.id) {
    normalized.id = `custom_node_${Date.now()}`;
  }
  return normalized;
};

const readSpecs = () => {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed
      .map((spec) => normalizeSpec(spec, spec))
      .filter((spec, index, self) => spec.id && self.findIndex((item) => item.id === spec.id) === index)
      .sort((a, b) => (a.label || '').localeCompare(b.label || ''));
  } catch (error) {
    console.error('Failed to read custom node specs', error);
    return [];
  }
};

const writeSpecs = (specs) => {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(specs));
  } catch (error) {
    alert('Failed to store custom node: ' + error.message);
  }
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

export const SAMPLE_NODE_TEMPLATES = [
  {
    id: 'sample_log_message',
    label: 'Sample: Log Message',
    category: 'Samples',
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
];

export const createEmptySpec = () => ({
  id: '',
  label: '',
  category: 'Custom',
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
});

export const listCustomNodeSpecs = () => readSpecs();

export const saveCustomNodeSpec = (spec) => {
  const specs = readSpecs();
  const previous = specs.find((item) => item.id === spec?.id);
  const normalized = normalizeSpec(spec, previous);
  normalized.updatedAt = new Date().toISOString();
  if (previous) {
    const index = specs.findIndex((item) => item.id === previous.id);
    specs[index] = { ...normalized, createdAt: previous.createdAt };
  } else {
    specs.push(normalized);
  }
  writeSpecs(specs);
  return normalized;
};

export const deleteCustomNodeSpec = (id) => {
  const specs = readSpecs().filter((item) => item.id !== id);
  writeSpecs(specs);
  return specs;
};

export const specsToDefinitions = (specs) =>
  (specs || [])
    .map((spec) => {
      const normalized = normalizeSpec(spec, spec);
      return {
        id: normalized.id,
        label: normalized.label,
        category: normalized.category || 'Custom',
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
          };
        }),
        script: createScriptFunction(normalized.script),
        specId: normalized.id,
        sourceSpec: normalized,
      };
    })
    .filter((definition, index, self) =>
      definition.id && self.findIndex((item) => item.id === definition.id) === index
    );

export const importSampleSpec = (sampleId) =>
  SAMPLE_NODE_TEMPLATES.find((template) => template.id === sampleId) || null;
