const STORAGE_KEY = 'nodeflow.psServerUrl';
export const DEFAULT_SERVER_URL = 'http://127.0.0.1:8787';

export function normalizeServerUrl(value) {
  if (!value) return '';
  let url = String(value).trim();
  if (!url) return '';
  if (!/^https?:\/\//i.test(url)) {
    url = `http://${url}`;
  }
  return url.replace(/\/+$/, '');
}

export function loadServerUrl() {
  try {
    const stored = window?.localStorage?.getItem(STORAGE_KEY);
    if (stored) {
      return normalizeServerUrl(stored) || DEFAULT_SERVER_URL;
    }
  } catch (error) {
    console.warn('Failed to read stored PowerShell server URL', error);
  }
  return DEFAULT_SERVER_URL;
}

export function saveServerUrl(url) {
  try {
    window?.localStorage?.setItem(STORAGE_KEY, url);
  } catch (error) {
    console.warn('Failed to persist PowerShell server URL', error);
  }
}

function setStatus(statusEl, message, state = '') {
  if (!statusEl) return;
  statusEl.textContent = message;
  const base = 'run-console__status';
  statusEl.className = state ? `${base} is-${state}` : base;
}

function formatEndpoint(baseUrl) {
  const normalized = normalizeServerUrl(baseUrl) || DEFAULT_SERVER_URL;
  return `${normalized.replace(/\/+$/, '')}/runScript`;
}

const DEFAULT_PANEL_HEIGHT = 280;
const MIN_PANEL_HEIGHT = 160;
const COLLAPSED_HEIGHT = 26;

const panelState = {
  container: null,
  elements: {},
  height: DEFAULT_PANEL_HEIGHT,
  lastHeight: DEFAULT_PANEL_HEIGHT,
  isCollapsed: true,
  isFullscreen: false,
  currentScript: '',
  currentEndpoint: '',
  execute: null,
  renderOutput: null,
  resizeHandler: null,
};

const updateBodyOffset = (height, collapsed) => {
  if (typeof document === 'undefined') {
    return;
  }
  const offset = collapsed ? COLLAPSED_HEIGHT : height;
  document.body.classList.add('has-run-console');
  document.body.classList.toggle('run-console-collapsed', collapsed);
  document.body.style.setProperty('--run-console-height', `${Math.round(height)}px`);
  document.body.style.setProperty('--run-console-offset', `${Math.round(offset)}px`);
};

const applyHeight = (height) => {
  const limit = panelState.isFullscreen ? window.innerHeight - 24 : window.innerHeight - 120;
  const maxHeight = Math.max(MIN_PANEL_HEIGHT, Math.min(height, limit));
  panelState.height = maxHeight;
  if (!panelState.isFullscreen) {
    panelState.lastHeight = maxHeight;
  }
  if (panelState.container && !panelState.isCollapsed) {
    panelState.container.style.height = `${maxHeight}px`;
    updateBodyOffset(maxHeight, false);
  }
};

const updateCollapseButton = () => {
  const button = panelState.elements.collapse;
  if (!button) {
    return;
  }
  if (panelState.isCollapsed) {
    button.textContent = '⛶';
    button.setAttribute('aria-label', 'デバッグコンソールを全画面表示');
  } else {
    button.textContent = '✕';
    button.setAttribute('aria-label', 'デバッグコンソールを最小化');
  }
};

const setFullscreen = (fullscreen) => {
  panelState.isFullscreen = fullscreen;
  const { container } = panelState;
  if (!container) {
    updateCollapseButton();
    return;
  }
  const previousHeight = panelState.height;
  container.classList.toggle('is-fullscreen', fullscreen);
  if (fullscreen) {
    if (!panelState.isCollapsed) {
      panelState.lastHeight =
        previousHeight || panelState.lastHeight || DEFAULT_PANEL_HEIGHT;
    }
    const height = Math.max(MIN_PANEL_HEIGHT, window.innerHeight - 24);
    panelState.height = height;
    container.style.height = `${height}px`;
    updateBodyOffset(height, false);
  } else {
    panelState.height = panelState.lastHeight || DEFAULT_PANEL_HEIGHT;
    if (!panelState.isCollapsed) {
      container.style.height = `${panelState.height}px`;
      updateBodyOffset(panelState.height, false);
    }
  }
  updateCollapseButton();
};

const setCollapsed = (collapsed) => {
  panelState.isCollapsed = collapsed;
  const { container, elements } = panelState;
  if (!container) {
    return;
  }
  container.classList.toggle('is-collapsed', collapsed);
  if (collapsed) {
    setFullscreen(false);
    container.style.height = '';
  } else {
    container.style.height = `${panelState.height}px`;
  }
  if (elements.toggle) {
    elements.toggle.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
    elements.toggle.textContent = collapsed ? '▴' : '▾';
  }
  if (elements.collapse) {
    elements.collapse.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
  }
  updateBodyOffset(panelState.height, collapsed);
  updateCollapseButton();
};

const ensureConsole = () => {
  if (panelState.container) {
    return panelState;
  }

  const container = document.createElement('section');
  container.className = 'run-console is-collapsed';
  container.setAttribute('role', 'region');
  container.setAttribute('aria-label', 'PowerShell 実行コンソール');
  container.innerHTML = `
    <div class="run-console__resize" data-role="resize" title="ドラッグでサイズ変更"></div>
    <header class="run-console__header">
      <button type="button" class="run-console__toggle" data-role="toggle" aria-expanded="false" aria-controls="run-console-body">▴</button>
      <div class="run-console__titles">
        <span class="run-console__context" data-role="context">全ノードを実行</span>
      </div>
      <span class="run-console__status" data-role="status">待機中</span>
      <span class="run-console__server" data-role="server"></span>
      <div class="run-console__actions">
        <button type="button" class="secondary" data-action="copy-output">出力をコピー</button>
        <button type="button" class="secondary" data-action="copy-script">スクリプトをコピー</button>
        <button type="button" class="primary" data-action="rerun">再実行</button>
      </div>
      <button type="button" class="run-console__collapse" data-role="collapse" aria-label="デバッグコンソールを最小化">✕</button>
    </header>
    <div class="run-console__body" id="run-console-body">
      <div class="run-console__section">
        <h3>出力</h3>
        <pre class="run-console__output" data-role="output"><span class="run-console__output-text is-empty">(出力はありません)</span></pre>
      </div>
      <details class="run-console__details" open>
        <summary>生成されたスクリプトを表示</summary>
        <textarea readonly class="run-console__script" data-role="script"></textarea>
      </details>
    </div>
  `;

  document.body.appendChild(container);
  document.body.classList.add('has-run-console');

  const elements = {
    toggle: container.querySelector('[data-role="toggle"]'),
    resize: container.querySelector('[data-role="resize"]'),
    status: container.querySelector('[data-role="status"]'),
    server: container.querySelector('[data-role="server"]'),
    context: container.querySelector('[data-role="context"]'),
    output: container.querySelector('[data-role="output"]'),
    script: container.querySelector('[data-role="script"]'),
    rerun: container.querySelector('[data-action="rerun"]'),
    copyOutput: container.querySelector('[data-action="copy-output"]'),
    copyScript: container.querySelector('[data-action="copy-script"]'),
    collapse: container.querySelector('[data-role="collapse"]'),
  };

  elements.toggle?.addEventListener('click', () => {
    setCollapsed(!panelState.isCollapsed);
    if (!panelState.isCollapsed) {
      applyHeight(panelState.height);
    }
  });

  elements.resize?.addEventListener('pointerdown', (event) => {
    event.preventDefault();
    if (panelState.isCollapsed) {
      setCollapsed(false);
    }
    if (panelState.isFullscreen) {
      setFullscreen(false);
    }
    const startHeight = panelState.height;
    const startY = event.clientY;
    const pointerId = event.pointerId;
    const handleEl = event.currentTarget || event.target;
    try {
      handleEl?.setPointerCapture(pointerId);
    } catch (error) {
      // Ignore if pointer capture isn't supported.
    }
    const handleMove = (moveEvent) => {
      const delta = moveEvent.clientY - startY;
      applyHeight(startHeight - delta);
    };
    const handleUp = (upEvent) => {
      if (upEvent.pointerId !== pointerId) {
        return;
      }
      window.removeEventListener('pointermove', handleMove);
      window.removeEventListener('pointerup', handleUp);
      window.removeEventListener('pointercancel', handleUp);
      try {
        handleEl?.releasePointerCapture(pointerId);
      } catch (error) {
        // Ignore release errors.
      }
    };
    window.addEventListener('pointermove', handleMove);
    window.addEventListener('pointerup', handleUp);
    window.addEventListener('pointercancel', handleUp);
  });

  elements.rerun?.addEventListener('click', () => {
    panelState.execute?.();
  });

  elements.copyScript?.addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(panelState.currentScript || '');
      setStatus(panelState.elements.status, 'スクリプトをコピーしました。', 'info');
    } catch (error) {
      setStatus(panelState.elements.status, `コピーに失敗しました: ${error.message}`, 'error');
    }
  });

  elements.copyOutput?.addEventListener('click', async () => {
    const text = panelState.elements.output?.textContent || '';
    if (!text || text === '(出力はありません)') {
      setStatus(panelState.elements.status, 'コピーできる出力がありません。', 'info');
      return;
    }
    try {
      await navigator.clipboard.writeText(text);
      setStatus(panelState.elements.status, '出力をコピーしました。', 'info');
    } catch (error) {
      setStatus(panelState.elements.status, `コピーに失敗しました: ${error.message}`, 'error');
    }
  });

  elements.collapse?.addEventListener('click', () => {
    if (panelState.isCollapsed) {
      const previousHeight =
        panelState.lastHeight || panelState.height || DEFAULT_PANEL_HEIGHT;
      panelState.height = previousHeight;
      setCollapsed(false);
      setFullscreen(true);
    } else {
      setCollapsed(true);
    }
  });

  panelState.container = container;
  panelState.elements = elements;
  panelState.height = DEFAULT_PANEL_HEIGHT;
  updateCollapseButton();
  panelState.renderOutput = (outputText, errors = []) => {
    const outputEl = panelState.elements.output;
    if (!outputEl) {
      return;
    }
    outputEl.innerHTML = '';
    const fragment = document.createDocumentFragment();
    const textValue =
      outputText === null || outputText === undefined ? '' : String(outputText);
    const hasOutput = textValue.trim().length > 0;
    if (hasOutput) {
      const span = document.createElement('span');
      span.className = 'run-console__output-text';
      span.textContent = textValue;
      fragment.appendChild(span);
    }
    const normalizedErrors = Array.isArray(errors)
      ? errors.map((line) => String(line || '').trim()).filter(Boolean)
      : [];
    if (normalizedErrors.length) {
      const span = document.createElement('span');
      span.className = 'run-console__output-error';
      span.textContent = normalizedErrors.join('\n');
      fragment.appendChild(span);
      outputEl.classList.add('has-error');
    } else {
      outputEl.classList.remove('has-error');
    }
    if (!fragment.childNodes.length) {
      const span = document.createElement('span');
      span.className = 'run-console__output-text is-empty';
      span.textContent = '(出力はありません)';
      fragment.appendChild(span);
    }
    outputEl.appendChild(fragment);
  };
  panelState.renderOutput('', []);
  setCollapsed(true);
  updateBodyOffset(panelState.height, true);

  if (!panelState.resizeHandler) {
    panelState.resizeHandler = () => {
      if (!panelState.container) {
        return;
      }
      if (panelState.isCollapsed) {
        updateBodyOffset(panelState.height, true);
        return;
      }
      applyHeight(panelState.height);
      if (panelState.container && !panelState.isCollapsed) {
        panelState.container.style.height = `${panelState.height}px`;
      }
    };
    window.addEventListener('resize', panelState.resizeHandler);
  }

  return panelState;
};

export function runScriptWithDialog(
  script,
  { serverUrl, contextLabel, targetNodeId, nodeIds } = {}
) {
  const endpoint = formatEndpoint(serverUrl);
  const state = ensureConsole();
  const { elements } = state;

  state.currentScript = script;
  state.currentEndpoint = endpoint;
  state.targetNodeId = targetNodeId || null;
  state.nodeIds = Array.isArray(nodeIds) ? [...nodeIds] : null;

  if (elements.script) {
    elements.script.value = script;
  }
  if (elements.server) {
    elements.server.textContent = `エンドポイント: ${endpoint}`;
  }
  if (elements.context) {
    const label = contextLabel || targetNodeId || '';
    elements.context.textContent = label ? `ノード: ${label}` : '全ノードを実行';
  }
  state.renderOutput?.('', []);

  setCollapsed(false);
  applyHeight(state.height);

  const execute = async () => {
    state.renderOutput?.('', []);
    if (elements.rerun) {
      elements.rerun.disabled = true;
    }
    setStatus(elements.status, 'PowerShell にスクリプトを送信しています…', 'pending');

    try {
      const response = await fetch(state.currentEndpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json; charset=utf-8' },
        body: JSON.stringify({ script: state.currentScript }),
      });

      let json;
      try {
        json = await response.json();
      } catch (error) {
        throw new Error('サーバーから無効なレスポンスを受信しました。');
      }

      if (!response.ok) {
        const message = json?.error || `HTTP ${response.status}`;
        throw new Error(message);
      }

      const hasErrors = json?.ok === false || (Array.isArray(json?.errors) && json.errors.length);
      const outputText = json?.output && String(json.output).trim() ? String(json.output) : '';
      const errorLines = Array.isArray(json?.errors) ? json.errors : [];
      state.renderOutput?.(outputText, errorLines);

      if (hasErrors) {
        setStatus(elements.status, 'PowerShell がエラーを返しました。', 'error');
      } else {
        setStatus(elements.status, 'PowerShell での実行が完了しました。', 'success');
      }
    } catch (error) {
      state.renderOutput?.('', [error.message]);
      setStatus(elements.status, 'サーバーへの接続に失敗しました。', 'error');
    } finally {
      if (elements.rerun) {
        elements.rerun.disabled = false;
      }
    }
  };

  state.execute = execute;
  return state.execute();
}
