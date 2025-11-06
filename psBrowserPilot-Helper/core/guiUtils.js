export const toPowerShellLiteral = (value) => {
  if (value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'boolean') {
    return value ? '$true' : '$false';
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? String(value) : '';
  }
  if (typeof value === 'object') {
    try {
      return toPowerShellLiteral(JSON.stringify(value));
    } catch (error) {
      return "''";
    }
  }
  const text = String(value);
  const trimmed = text.trim();
  if (!trimmed) {
    return "''";
  }
  if (
    (trimmed.startsWith("'") && trimmed.endsWith("'")) ||
    (trimmed.startsWith('"') && trimmed.endsWith('"'))
  ) {
    return trimmed;
  }
  if (/^\$[A-Za-z0-9_]+$/.test(trimmed)) {
    return trimmed;
  }
  if (/^(?:(?:0x)[0-9A-Fa-f]+|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?)$/.test(trimmed)) {
    return trimmed;
  }
  return `'${text.replace(/'/g, "''")}'`;
};
