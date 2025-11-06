export const createButton = (label, { className = '', type = 'button' } = {}) => {
  const button = document.createElement('button');
  button.type = type;
  if (className) {
    button.className = className;
  }
  if (label !== undefined) {
    button.textContent = label;
  }
  return button;
};

export const replaceWithClone = (element) => {
  if (!element?.parentNode) {
    return element;
  }
  const clone = element.cloneNode(true);
  element.parentNode.replaceChild(clone, element);
  return clone;
};

export const removeChildren = (element) => {
  if (!element) {
    return;
  }
  while (element.firstChild) {
    element.removeChild(element.firstChild);
  }
};

export const bindEvent = (element, type, handler, options) => {
  element?.addEventListener(type, handler, options);
  return () => element?.removeEventListener(type, handler, options);
};

export const UIHelper = {
  createButton,
  replaceWithClone,
  removeChildren,
  bindEvent,
};
