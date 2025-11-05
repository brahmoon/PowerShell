const createElement = (tag, className) => {
  const element = document.createElement(tag);
  if (className) {
    element.className = className;
  }
  return element;
};

export const createPaletteController = ({
  grid,
  addButton,
  root = grid?.parentElement || null,
  storage,
  storageKey = 'palette',
  onApply,
}) => {
  if (!grid || !addButton) {
    return {
      load: () => [],
      addColors: () => ({ added: 0, total: 0 }),
      clear: () => {},
      remove: () => null,
      persist: () => {},
      dispose: () => {},
      get size() {
        return 0;
      },
      entries() {
        return [];
      },
    };
  }

  const paletteMap = new Map();
  let contextMenu = null;
  let contextTargetKey = null;
  const documentHandlers = new Map();

  const persist = ({ silent = false } = {}) => {
    if (!storage) {
      return;
    }
    const payload = Array.from(paletteMap.values()).map((entry) => ({
      hex: entry.hex,
      ole: entry.ole,
    }));
    storage.saveJsonConfig?.(storageKey, payload, { silent });
  };

  const hideContextMenu = () => {
    if (!contextMenu) {
      return;
    }
    contextMenu.classList.add('is-hidden');
    contextTargetKey = null;
  };

  const ensureContextMenu = () => {
    if (contextMenu || !root) {
      return contextMenu;
    }
    const menu = createElement('div', 'palette-swatch-menu is-hidden');
    menu.tabIndex = -1;
    menu.addEventListener('contextmenu', (event) => event.preventDefault());
    menu.addEventListener('pointerdown', (event) => event.stopPropagation());

    const deleteButton = createElement('button', 'palette-swatch-menu__item');
    deleteButton.type = 'button';
    deleteButton.textContent = '削除';
    deleteButton.addEventListener('click', () => {
      if (!contextTargetKey) {
        return;
      }
      remove(contextTargetKey, { persist: true });
      hideContextMenu();
    });

    menu.appendChild(deleteButton);
    root.appendChild(menu);
    contextMenu = menu;

    const handlePointerDown = (event) => {
      if (!contextMenu || contextMenu.classList.contains('is-hidden')) {
        return;
      }
      const target = event.target;
      if (!(target instanceof Node) || !root.contains(target) || !contextMenu.contains(target)) {
        hideContextMenu();
      }
    };
    const handleKeyDown = (event) => {
      if (event.key === 'Escape') {
        hideContextMenu();
      }
    };
    const handleScroll = () => hideContextMenu();

    document.addEventListener('pointerdown', handlePointerDown);
    document.addEventListener('keydown', handleKeyDown);
    document.addEventListener('scroll', handleScroll, true);

    documentHandlers.set('pointerdown', handlePointerDown);
    documentHandlers.set('keydown', handleKeyDown);
    documentHandlers.set('scroll', handleScroll);

    return menu;
  };

  const openContextMenu = (event, key) => {
    const menu = ensureContextMenu();
    if (!menu) {
      return;
    }
    contextTargetKey = key;
    menu.classList.remove('is-hidden');
    const rect = root?.getBoundingClientRect();
    if (!rect) {
      hideContextMenu();
      return;
    }
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    menu.style.setProperty('--menu-left', `${Math.max(0, Math.round(x))}px`);
    menu.style.setProperty('--menu-top', `${Math.max(0, Math.round(y))}px`);
    menu.style.left = `${Math.max(0, Math.round(x))}px`;
    menu.style.top = `${Math.max(0, Math.round(y))}px`;
  };

  const createSwatchButton = (hex, ole) => {
    const button = createElement('button', 'palette-swatch');
    button.type = 'button';
    button.dataset.key = String(ole);
    button.style.setProperty('--swatch-color', hex);
    button.textContent = hex;
    button.addEventListener('click', async (event) => {
      if (typeof onApply !== 'function') {
        return;
      }
      const result = onApply({ hex, ole, event, button });
      if (result && typeof result.then === 'function') {
        button.disabled = true;
        try {
          await result;
        } finally {
          button.disabled = false;
        }
      }
    });
    button.addEventListener('contextmenu', (event) => {
      event.preventDefault();
      event.stopPropagation();
      openContextMenu(event, button.dataset.key);
    });
    return button;
  };

  const add = (hex, ole, { persist: shouldPersist = true } = {}) => {
    if (!hex || ole === undefined || ole === null) {
      return false;
    }
    const key = String(ole);
    if (paletteMap.has(key)) {
      return false;
    }
    const entry = { hex, ole };
    const button = createSwatchButton(hex, ole);
    paletteMap.set(key, { ...entry, button });
    grid.insertBefore(button, addButton);
    if (shouldPersist) {
      persist();
    }
    return true;
  };

  const addColors = (colors, { replace = false, persist: shouldPersist = true } = {}) => {
    const list = Array.isArray(colors) ? colors : [];
    if (replace) {
      clear({ persist: false });
    }
    let added = 0;
    list.forEach((item) => {
      if (!item) return;
      const hex = typeof item.hex === 'string' ? item.hex.trim() : '';
      const ole = item.ole;
      if (!hex || ole === undefined || ole === null) {
        return;
      }
      if (add(hex, ole, { persist: false })) {
        added += 1;
      }
    });
    if (shouldPersist && (replace || added)) {
      persist();
    }
    return { added, total: list.length };
  };

  const clear = ({ persist: shouldPersist = false } = {}) => {
    paletteMap.forEach((entry) => {
      entry.button?.remove();
    });
    paletteMap.clear();
    hideContextMenu();
    if (shouldPersist) {
      persist();
    }
  };

  const remove = (key, { persist: shouldPersist = false } = {}) => {
    if (key === undefined || key === null) {
      return null;
    }
    const stringKey = String(key);
    const entry = paletteMap.get(stringKey);
    if (!entry) {
      return null;
    }
    entry.button?.remove();
    paletteMap.delete(stringKey);
    if (shouldPersist) {
      persist();
    }
    return { hex: entry.hex, ole: entry.ole };
  };

  const load = ({ replace = true } = {}) => {
    if (!storage) {
      return { added: 0, total: 0 };
    }
    const stored = storage.loadJsonConfig?.(storageKey, []);
    if (!Array.isArray(stored) || !stored.length) {
      return { added: 0, total: 0 };
    }
    return addColors(stored, { replace, persist: false });
  };

  const dispose = () => {
    clear({ persist: false });
    if (contextMenu) {
      contextMenu.remove();
      contextMenu = null;
    }
    documentHandlers.forEach((handler, type) => {
      if (type === 'scroll') {
        document.removeEventListener(type, handler, true);
      } else {
        document.removeEventListener(type, handler);
      }
    });
    documentHandlers.clear();
  };

  return {
    addColors,
    add,
    clear,
    remove,
    load,
    persist,
    hideContextMenu,
    dispose,
    get size() {
      return paletteMap.size;
    },
    entries() {
      return Array.from(paletteMap.values()).map(({ hex, ole }) => ({ hex, ole }));
    },
  };
};

export const PaletteUI = { create: createPaletteController };
