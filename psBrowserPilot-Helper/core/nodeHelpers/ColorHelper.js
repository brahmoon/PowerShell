const clampChannel = (value) => {
  const num = Number(value);
  if (!Number.isFinite(num)) {
    throw new Error('RGB の値が数値ではありません。');
  }
  return Math.max(0, Math.min(255, Math.round(num)));
};

export const rgbToHex = (r, g, b) =>
  `#${[r, g, b]
    .map((channel) => clampChannel(channel).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase()}`;

export const rgbToOle = (r, g, b) => {
  const rr = clampChannel(r);
  const gg = clampChannel(g);
  const bb = clampChannel(b);
  return rr + gg * 256 + bb * 65536;
};

export const hexToRgb = (hex) => {
  const value = String(hex || '').trim().replace(/^#/, '');
  if (!/^[0-9a-f]{3}([0-9a-f]{3})?$/i.test(value)) {
    throw new Error('HEX の形式が不正です。');
  }
  const full = value.length === 3 ? value.split('').map((c) => c + c).join('') : value;
  const int = parseInt(full, 16);
  return {
    r: (int >> 16) & 255,
    g: (int >> 8) & 255,
    b: int & 255,
  };
};

export const oleToColor = (oleValue) => {
  const numeric = Number(oleValue);
  if (!Number.isFinite(numeric)) {
    throw new Error('OLE 色の値が不正です。');
  }
  const ole = numeric >>> 0;
  const r = ole & 255;
  const g = (ole >>> 8) & 255;
  const b = (ole >>> 16) & 255;
  return { hex: rgbToHex(r, g, b), ole };
};

const tupleToColor = (tuple) => {
  if (!Array.isArray(tuple) || tuple.length !== 3) {
    throw new Error('RGB の配列は 3 要素で指定してください。');
  }
  const [r, g, b] = tuple.map(clampChannel);
  return { hex: rgbToHex(r, g, b), ole: rgbToOle(r, g, b) };
};

const colorFromHexString = (value) => {
  const str = String(value || '').trim();
  if (!str) {
    throw new Error('HEX の形式が不正です。');
  }
  let normalized = str;
  if (/^0x[0-9a-f]+$/i.test(normalized)) {
    normalized = `#${normalized.slice(2)}`;
  } else if (!normalized.startsWith('#')) {
    normalized = `#${normalized}`;
  }
  const { r, g, b } = hexToRgb(normalized);
  return { hex: rgbToHex(r, g, b), ole: rgbToOle(r, g, b) };
};

const tryParseStructured = (value) => {
  if (typeof value !== 'string') {
    return value;
  }
  const trimmed = value.trim();
  if (!trimmed) {
    return '';
  }
  if (
    (trimmed.startsWith('{') && trimmed.endsWith('}')) ||
    (trimmed.startsWith('[') && trimmed.endsWith(']')) ||
    /^(true|false|null|-?\d+(?:\.\d+)?)$/i.test(trimmed)
  ) {
    try {
      return JSON.parse(trimmed);
    } catch (error) {
      // ignore parse failure and return original string
    }
  }
  return value;
};

const isPlainObject = (value) => Object.prototype.toString.call(value) === '[object Object]';

const flattenColorInput = (input) => {
  const value = tryParseStructured(input);
  if (value === undefined || value === null || value === '') {
    return [];
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    return [oleToColor(value)];
  }

  if (typeof value === 'string') {
    const str = value.trim();
    if (!str) {
      return [];
    }
    if (/^#/.test(str) || /^0x[0-9a-f]+$/i.test(str) || /^[0-9a-f]{3,6}$/i.test(str)) {
      try {
        return [colorFromHexString(str)];
      } catch (error) {
        return [];
      }
    }
    const rgbMatch = str.match(/^rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/i);
    if (rgbMatch) {
      return [tupleToColor([rgbMatch[1], rgbMatch[2], rgbMatch[3]])];
    }
    const parts = str.split(/[\s,;]+/).filter(Boolean);
    if (parts.length === 3 && parts.every((p) => /^-?\d+(?:\.\d+)?$/.test(p))) {
      try {
        return [tupleToColor(parts)];
      } catch (error) {
        return [];
      }
    }
    if (parts.length > 1 && parts.every((p) => /^#?[0-9a-f]{3,6}$/i.test(p))) {
      return parts.flatMap((part) => flattenColorInput(part));
    }
    return [];
  }

  if (Array.isArray(value)) {
    if (value.length === 3 && value.every((item) => Number.isFinite(Number(item)))) {
      return [tupleToColor(value)];
    }
    return value.flatMap((item) => flattenColorInput(item));
  }

  if (isPlainObject(value)) {
    const results = [];
    const directKeys = ['hex', 'Hex', 'color', 'Color', 'colour', 'Colour', 'value', 'Value'];
    directKeys.forEach((key) => {
      if (Object.prototype.hasOwnProperty.call(value, key)) {
        results.push(...flattenColorInput(value[key]));
      }
    });
    if ('ole' in value) {
      results.push(...flattenColorInput(value.ole));
    }
    if ('rgb' in value) {
      results.push(...flattenColorInput(value.rgb));
    }
    if ('RGB' in value) {
      results.push(...flattenColorInput(value.RGB));
    }
    if (['r', 'g', 'b'].every((key) => Object.prototype.hasOwnProperty.call(value, key))) {
      results.push(tupleToColor([value.r, value.g, value.b]));
    }
    const nestedKeys = ['colors', 'Colours', 'swatches', 'items', 'data', 'list'];
    nestedKeys.forEach((key) => {
      if (Object.prototype.hasOwnProperty.call(value, key)) {
        results.push(...flattenColorInput(value[key]));
      }
    });
    return results;
  }

  return [];
};

export const parse = (value) => {
  const colors = flattenColorInput(value);
  if (!colors.length) {
    throw new Error('Color の解釈に失敗しました（HEX / rgb(r,g,b) / [r,g,b] / OLE 対応）。');
  }
  const seen = new Set();
  return colors.filter((color) => {
    const key = String(color?.ole ?? '');
    if (!key || seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
};

export const COLOR_HELPERS_SCRIPT = `
function Convert-OleToHex {
  param([Parameter()][object]$Ole)
  if ($null -eq $Ole) { return $null }
  try {
    $color = [System.Drawing.ColorTranslator]::FromOle([int]$Ole)
    return ('#{0:X2}{1:X2}{2:X2}' -f $color.R, $color.G, $color.B)
  } catch {
    return $null
  }
}
`;

export const ColorHelper = {
  parse,
  rgbToHex,
  rgbToOle,
  hexToRgb,
  oleToColor,
  COLOR_HELPERS_SCRIPT,
};
