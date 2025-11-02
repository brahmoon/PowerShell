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
  const base = 'run-dialog__status';
  statusEl.className = state ? `${base} is-${state}` : base;
}

function formatEndpoint(baseUrl) {
  const normalized = normalizeServerUrl(baseUrl) || DEFAULT_SERVER_URL;
  return `${normalized.replace(/\/+$/, '')}/runScript`;
}

export function runScriptWithDialog(script, { serverUrl } = {}) {
  const endpoint = formatEndpoint(serverUrl);
  const dialog = document.createElement('dialog');
  dialog.className = 'run-dialog';
  dialog.innerHTML = `
    <form method="dialog" class="run-dialog__form">
      <header class="run-dialog__header">
        <h2>PowerShell Script Runner</h2>
        <p class="run-dialog__server">エンドポイント: <code data-role="server"></code></p>
        <p class="run-dialog__status" data-role="status">PowerShell にスクリプトを送信しています…</p>
      </header>
      <section class="run-dialog__body">
        <pre class="run-dialog__output" data-role="output">(出力はありません)</pre>
        <div class="run-dialog__errors" data-role="errors"></div>
        <details class="run-dialog__details">
          <summary>生成されたスクリプトを表示</summary>
          <textarea readonly class="run-dialog__script" data-role="script"></textarea>
        </details>
      </section>
      <menu class="run-dialog__actions">
        <button value="close">閉じる</button>
        <button type="button" class="secondary" data-action="copy-output">出力をコピー</button>
        <button type="button" class="secondary" data-action="copy-script">スクリプトをコピー</button>
        <button type="button" class="primary" data-action="rerun">再実行</button>
      </menu>
    </form>
  `;

  document.body.appendChild(dialog);

  const statusEl = dialog.querySelector('[data-role="status"]');
  const serverEl = dialog.querySelector('[data-role="server"]');
  const outputEl = dialog.querySelector('[data-role="output"]');
  const errorsEl = dialog.querySelector('[data-role="errors"]');
  const scriptEl = dialog.querySelector('[data-role="script"]');
  const rerunButton = dialog.querySelector('[data-action="rerun"]');
  const copyOutputButton = dialog.querySelector('[data-action="copy-output"]');
  const copyScriptButton = dialog.querySelector('[data-action="copy-script"]');

  if (scriptEl) {
    scriptEl.value = script;
  }
  if (serverEl) {
    serverEl.textContent = endpoint;
  }

  const execute = async () => {
    if (outputEl) {
      outputEl.textContent = '(出力はありません)';
    }
    if (errorsEl) {
      errorsEl.textContent = '';
    }
    if (rerunButton) {
      rerunButton.disabled = true;
    }
    setStatus(statusEl, 'PowerShell にスクリプトを送信しています…', 'pending');

    try {
      const response = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json; charset=utf-8' },
        body: JSON.stringify({ script }),
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
      if (outputEl) {
        const text = json?.output;
        outputEl.textContent = text && String(text).trim() ? String(text) : '(出力はありません)';
      }
      if (errorsEl) {
        if (Array.isArray(json?.errors) && json.errors.length) {
          errorsEl.textContent = json.errors.join('\n');
        } else {
          errorsEl.textContent = '';
        }
      }

      if (hasErrors) {
        setStatus(statusEl, 'PowerShell がエラーを返しました。', 'error');
      } else {
        setStatus(statusEl, 'PowerShell での実行が完了しました。', 'success');
      }
    } catch (error) {
      if (errorsEl) {
        errorsEl.textContent = error.message;
      }
      setStatus(statusEl, 'サーバーへの接続に失敗しました。', 'error');
    } finally {
      if (rerunButton) {
        rerunButton.disabled = false;
      }
    }
  };

  rerunButton?.addEventListener('click', () => {
    execute();
  });

  copyScriptButton?.addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(script);
      setStatus(statusEl, 'スクリプトをコピーしました。', 'info');
    } catch (error) {
      setStatus(statusEl, `コピーに失敗しました: ${error.message}`, 'error');
    }
  });

  copyOutputButton?.addEventListener('click', async () => {
    if (!outputEl) return;
    const text = outputEl.textContent || '';
    if (!text || text === '(出力はありません)') {
      setStatus(statusEl, 'コピーできる出力がありません。', 'info');
      return;
    }
    try {
      await navigator.clipboard.writeText(text);
      setStatus(statusEl, '出力をコピーしました。', 'info');
    } catch (error) {
      setStatus(statusEl, `コピーに失敗しました: ${error.message}`, 'error');
    }
  });

  dialog.addEventListener('close', () => {
    dialog.remove();
  });

  dialog.showModal();
  execute();
}
