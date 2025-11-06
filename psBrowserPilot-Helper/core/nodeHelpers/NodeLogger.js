const formatTag = (context = {}) => {
  const id = context?.node?.id ? `#${context.node.id}` : '';
  const label = context?.node?.definition?.label || context?.node?.label;
  if (label && id) {
    return `[${label}${id}]`;
  }
  if (label) {
    return `[${label}]`;
  }
  if (id) {
    return `[Node${id}]`;
  }
  return '[Node]';
};

export const createNodeLogger = (context = {}) => {
  const tag = formatTag(context);
  const bind = (method) => (...args) => {
    try {
      console[method]?.(tag, ...args);
    } catch {
      // ignore console access errors
    }
  };
  return {
    debug: bind('debug'),
    info: bind('info'),
    warn: bind('warn'),
    error: bind('error'),
    tag,
  };
};

export const NodeLogger = { create: createNodeLogger };
