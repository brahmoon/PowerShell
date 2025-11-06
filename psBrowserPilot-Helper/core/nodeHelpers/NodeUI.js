const STATUS_CLASSNAMES = ['info', 'success', 'error', 'pending', 'warning'];

const normalizeState = (state) => {
  if (typeof state !== 'string') {
    return '';
  }
  const trimmed = state.trim();
  return STATUS_CLASSNAMES.includes(trimmed) ? trimmed : trimmed || '';
};

const applyState = (element, state) => {
  if (!element) return;
  STATUS_CLASSNAMES.forEach((name) => element.classList.remove(`is-${name}`));
  if (state) {
    element.classList.add(`is-${state}`);
    element.setAttribute('data-state', state);
  } else {
    element.removeAttribute('data-state');
  }
};

export const setStatus = (element, message = '', state = '') => {
  if (!element) return;
  element.textContent = typeof message === 'string' ? message : '';
  const normalized = normalizeState(state);
  applyState(element, normalized);
};

export const clearStatus = (element) => {
  if (!element) return;
  element.textContent = '';
  applyState(element, '');
};

export const bindStatus = (element, { initialMessage = '', initialState = '' } = {}) => {
  if (!element) {
    return {
      set: () => {},
      clear: () => {},
      dispose: () => {},
    };
  }

  const controller = {
    element,
    set(message, state = '') {
      setStatus(element, message, state);
    },
    clear() {
      clearStatus(element);
    },
    dispose() {
      this.element = null;
    },
  };

  if (initialMessage || initialState) {
    controller.set(initialMessage, initialState);
  }

  return controller;
};

export const createStatusController = bindStatus;

export { STATUS_CLASSNAMES };
