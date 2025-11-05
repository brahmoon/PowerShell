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
      // fallthrough to string value
    }
  }
  return value;
};

const isPlainObject = (value) => Object.prototype.toString.call(value) === '[object Object]';

export const parseRangeInfo = (input) => {
  const info = { address: '', workbook: '', sheet: '' };
  const visited = new Set();

  const assignString = (key, candidate) => {
    if (info[key]) return;
    if (typeof candidate === 'string') {
      const trimmed = candidate.trim();
      if (trimmed) {
        info[key] = trimmed;
      }
    }
  };

  const inspect = (value) => {
    if (info.address && info.workbook && info.sheet) {
      return;
    }
    const parsed = tryParseStructured(value);
    if (parsed === undefined || parsed === null || parsed === '') {
      return;
    }
    if (typeof parsed === 'string') {
      assignString('address', parsed);
      return;
    }
    if (typeof parsed === 'number' || typeof parsed === 'boolean') {
      return;
    }
    if (Array.isArray(parsed)) {
      parsed.forEach((item) => inspect(item));
      return;
    }
    if (!isPlainObject(parsed) || visited.has(parsed)) {
      return;
    }
    visited.add(parsed);

    assignString('address', parsed.SelectionAddress__raw);
    assignString('address', parsed.SelectionAddress);
    assignString('address', parsed.address);
    assignString('address', parsed.Address);
    assignString('address', parsed.range);
    assignString('address', parsed.Range);
    assignString('address', parsed.targetAddress);

    assignString('workbook', parsed.workbook);
    assignString('workbook', parsed.Workbook);
    assignString('workbook', parsed.workbookName);
    assignString('workbook', parsed.WorkbookName);
    assignString('workbook', parsed.book);
    assignString('workbook', parsed.Book);

    assignString('sheet', parsed.sheet);
    assignString('sheet', parsed.Sheet);
    assignString('sheet', parsed.sheetName);
    assignString('sheet', parsed.SheetName);
    assignString('sheet', parsed.worksheet);
    assignString('sheet', parsed.Worksheet);

    const nestedKeys = [
      'selection',
      'Selection',
      'target',
      'Target',
      'value',
      'Value',
      'data',
      'Data',
      'range',
      'Range',
    ];
    nestedKeys.forEach((key) => {
      if (Object.prototype.hasOwnProperty.call(parsed, key)) {
        inspect(parsed[key]);
      }
    });
  };

  inspect(input);
  return info;
};

export const parseSheetInfo = (input) => {
  const info = { sheet: '', workbook: '' };
  const visited = new Set();

  const assignSheet = (value) => {
    if (!info.sheet && typeof value === 'string') {
      const trimmed = value.trim();
      if (trimmed) {
        info.sheet = trimmed;
      }
    }
  };

  const assignWorkbook = (value) => {
    if (!info.workbook && typeof value === 'string') {
      const trimmed = value.trim();
      if (trimmed) {
        info.workbook = trimmed;
      }
    }
  };

  const inspect = (value) => {
    if (info.sheet && info.workbook) {
      return;
    }
    const parsed = tryParseStructured(value);
    if (parsed === undefined || parsed === null || parsed === '') {
      return;
    }
    if (typeof parsed === 'string') {
      assignSheet(parsed);
      return;
    }
    if (typeof parsed === 'number' || typeof parsed === 'boolean') {
      return;
    }
    if (Array.isArray(parsed)) {
      parsed.forEach((item) => inspect(item));
      return;
    }
    if (!isPlainObject(parsed) || visited.has(parsed)) {
      return;
    }
    visited.add(parsed);

    assignSheet(parsed.sheet);
    assignSheet(parsed.Sheet);
    assignSheet(parsed.sheetName);
    assignSheet(parsed.SheetName);
    assignSheet(parsed.worksheet);
    assignSheet(parsed.Worksheet);

    assignWorkbook(parsed.workbook);
    assignWorkbook(parsed.Workbook);
    assignWorkbook(parsed.workbookName);
    assignWorkbook(parsed.WorkbookName);
    assignWorkbook(parsed.book);
    assignWorkbook(parsed.Book);

    const nested = [
      'selection',
      'Selection',
      'target',
      'Target',
      'source',
      'Source',
      'context',
      'Context',
      'sheet',
      'Sheet',
      'worksheet',
      'Worksheet',
    ];
    nested.forEach((key) => {
      if (Object.prototype.hasOwnProperty.call(parsed, key)) {
        inspect(parsed[key]);
      }
    });
  };

  inspect(input);
  return info;
};

export const createRangeSpec = (info, literal) => {
  const literalFn = typeof literal === 'function' ? literal : (value) => value;
  const addressLiteral = info?.address ? literalFn(info.address) : null;
  const workbookLiteral = info?.workbook ? literalFn(info.workbook) : null;
  const sheetLiteral = info?.sheet ? literalFn(info.sheet) : null;
  return {
    addressLiteral,
    workbookLiteral,
    sheetLiteral,
  };
};

export const ExcelHelper = {
  parseRangeInfo,
  parseSheetInfo,
  createRangeSpec,
};
