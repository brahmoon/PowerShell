const EXCEL_SERVER_BASE = 'http://127.0.0.1:8080';

const requestJson = async (path, options = {}) => {
  const url = `${EXCEL_SERVER_BASE}${path}`;
  const response = await fetch(url, {
    method: 'GET',
    headers: { 'Content-Type': 'application/json' },
    ...options,
    body:
      options.body === undefined
        ? undefined
        : typeof options.body === 'string'
        ? options.body
        : JSON.stringify(options.body),
  });
  let data = null;
  try {
    data = await response.json();
  } catch (error) {
    data = null;
  }
  if (!response.ok) {
    const message = data?.error || `HTTP ${response.status}`;
    throw new Error(message);
  }
  return data;
};

const createExcelWorkbookSelector = () => ({
  id: 'ui_excel_workbook_selector',
  label: 'Excel Workbook Selector',
  category: 'Excel UI',
  execution: 'ui',
  inputs: [],
  outputs: ['WorkbookName'],
  initialConfig: {
    WorkbookName: '',
    WorkbookName__raw: '',
    selectedWorkbookLabel: '',
  },
  script: () => '',
  render: ({ node, controls, updateConfig, toPowerShellLiteral: literal }) => {
    if (!controls) {
      return null;
    }
    controls.innerHTML = '';
    const wrapper = document.createElement('div');
    wrapper.className = 'excel-ui-node workbook-selector-node';

    const header = document.createElement('div');
    header.className = 'excel-ui-row';
    const select = document.createElement('select');
    select.className = 'excel-ui-select';
    const refreshBtn = document.createElement('button');
    refreshBtn.type = 'button';
    refreshBtn.className = 'excel-ui-button';
    refreshBtn.textContent = '更新';
    header.append(select, refreshBtn);

    const status = document.createElement('div');
    status.className = 'excel-ui-status';

    wrapper.append(header, status);
    controls.appendChild(wrapper);

    let disposed = false;

    const setStatus = (message, type = '') => {
      if (disposed) return;
      status.textContent = message || '';
      status.dataset.state = type;
    };

    const applySelection = (workbookName) => {
      const raw = workbookName || '';
      const scriptValue = raw ? literal(raw) : '';
      updateConfig('WorkbookName__raw', raw, { silent: true });
      updateConfig('WorkbookName', scriptValue, { silent: true });
      updateConfig('selectedWorkbookLabel', raw, { silent: false });
    };

    const populateOptions = (items, { preserveSelection = true } = {}) => {
      select.innerHTML = '';
      const placeholder = document.createElement('option');
      placeholder.value = '';
      placeholder.textContent = items.length ? 'ブックを選択' : 'ブックが見つかりません';
      placeholder.disabled = true;
      placeholder.selected = true;
      select.appendChild(placeholder);
      items.forEach((item) => {
        const option = document.createElement('option');
        option.value = item;
        option.textContent = item;
        select.appendChild(option);
      });
      if (preserveSelection) {
        const stored = node.config.WorkbookName__raw || node.config.selectedWorkbookLabel;
        if (stored && items.includes(stored)) {
          select.value = stored;
        }
      }
      select.disabled = !items.length;
    };

    const fetchWorkbooks = async () => {
      setStatus('Excelのブックを取得中…', 'pending');
      refreshBtn.disabled = true;
      try {
        const data = await requestJson('/listWorkbooks');
        if (disposed) return;
        const workbooks = Array.isArray(data?.workbooks) ? data.workbooks : [];
        populateOptions(workbooks);
        if (workbooks.length) {
          setStatus(`${workbooks.length} 件のブックを検出しました。`, 'success');
        } else {
          setStatus('開いているブックが見つかりません。', 'info');
        }
      } catch (error) {
        if (!disposed) {
          populateOptions([], { preserveSelection: false });
          setStatus(`取得に失敗しました: ${error.message}`, 'error');
        }
      } finally {
        if (!disposed) {
          refreshBtn.disabled = false;
        }
      }
    };

    select.addEventListener('change', (event) => {
      const value = event.currentTarget.value;
      applySelection(value);
      if (value) {
        setStatus(`選択中: ${value}`, 'success');
      } else {
        setStatus('ブックを選択してください。', 'info');
      }
    });

    refreshBtn.addEventListener('click', (event) => {
      event.preventDefault();
      fetchWorkbooks();
    });

    const stored = node.config.selectedWorkbookLabel;
    populateOptions(stored ? [stored] : [], { preserveSelection: true });
    if (stored) {
      const scriptValue = literal(stored);
      updateConfig('WorkbookName__raw', stored, { silent: true });
      updateConfig('WorkbookName', scriptValue, { silent: true });
      setStatus(`選択中: ${stored}`, 'success');
    } else {
      setStatus('ブックを選択してください。', 'info');
    }

    fetchWorkbooks();

    return () => {
      disposed = true;
    };
  },
});

const createExcelSelectionInspector = () => ({
  id: 'ui_excel_selection_inspector',
  label: 'Excel Selection Inspector',
  category: 'Excel UI',
  execution: 'ui',
  inputs: ['WorkbookName'],
  outputs: ['SelectionAddress', 'SelectionText', 'FontColorHex', 'InteriorColorHex'],
  initialConfig: {
    WorkbookName__raw: '',
    SelectionAddress: '',
    SelectionAddress__raw: '',
    SelectionText: '',
    SelectionText__raw: '',
    FontColorHex: '',
    FontColorHex__raw: '',
    InteriorColorHex: '',
    InteriorColorHex__raw: '',
    lastWorkbook: '',
    lastSheet: '',
  },
  script: () => '',
  render: ({
    node,
    controls,
    updateConfig,
    resolveInput,
    toPowerShellLiteral: literal,
  }) => {
    if (!controls) {
      return null;
    }
    controls.innerHTML = '';
    const wrapper = document.createElement('div');
    wrapper.className = 'excel-ui-node selection-inspector-node';

    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'excel-ui-button primary';
    button.textContent = '選択セル情報を取得';

    const status = document.createElement('div');
    status.className = 'excel-ui-status';

    const result = document.createElement('dl');
    result.className = 'excel-selection-result';

    const createItem = (label) => {
      const dt = document.createElement('dt');
      dt.textContent = label;
      const dd = document.createElement('dd');
      dd.textContent = '';
      result.append(dt, dd);
      return dd;
    };

    const addressEl = createItem('セル座標');
    const textEl = createItem('セルのテキスト');
    const fontEl = createItem('Font.Color');
    const interiorEl = createItem('Interior.Color');

    wrapper.append(button, status, result);
    controls.appendChild(wrapper);

    let disposed = false;

    const setStatus = (message, type = '') => {
      if (disposed) return;
      status.textContent = message || '';
      status.dataset.state = type;
    };

    const setOutput = (key, rawValue) => {
      const raw = rawValue ?? '';
      updateConfig(`${key}__raw`, raw, { silent: true });
      updateConfig(key, raw ? literal(raw) : '', { silent: false });
    };

    const updateDisplay = (data) => {
      addressEl.textContent = data.address || '-';
      textEl.textContent = data.value || '-';
      fontEl.textContent = data.fontHex || '-';
      interiorEl.textContent = data.interiorHex || '-';
    };

    const applyStoredResult = () => {
      updateDisplay({
        address: node.config.SelectionAddress__raw,
        value: node.config.SelectionText__raw,
        fontHex: node.config.FontColorHex__raw,
        interiorHex: node.config.InteriorColorHex__raw,
      });
    };

    const ensureWorkbook = async (workbookName) => {
      if (!workbookName) {
        return;
      }
      await requestJson('/setWorkbook', {
        method: 'POST',
        body: { name: workbookName },
      });
      updateConfig('WorkbookName__raw', workbookName, { silent: true });
    };

    const fetchSelection = async () => {
      const workbook =
        resolveInput('WorkbookName', { preferRaw: true }) || node.config.WorkbookName__raw;
      button.disabled = true;
      setStatus('選択セル情報を取得中…', 'pending');
      try {
        await ensureWorkbook(workbook);
        const data = await requestJson('/getSelection');
        if (disposed) return;
        if (!data?.ok) {
          throw new Error(data?.error || '情報を取得できませんでした');
        }
        const address = data.address || '';
        const value = data.value ?? '';
        const fontHex = data.font_color?.hex || '';
        const interiorHex = data.interior_color?.hex || '';

        setOutput('SelectionAddress', address);
        setOutput('SelectionText', value);
        setOutput('FontColorHex', fontHex);
        setOutput('InteriorColorHex', interiorHex);
        updateConfig('lastWorkbook', data.workbook || workbook || '', { silent: true });
        updateConfig('lastSheet', data.sheet || '', { silent: true });
        applyStoredResult();
        const workbookLabel = data.workbook || workbook;
        const sheetLabel = data.sheet ? `${data.sheet}` : '';
        const suffix = sheetLabel ? `${sheetLabel} - ${address}` : address;
        if (workbookLabel && suffix) {
          setStatus(`${workbookLabel}: ${suffix}`, 'success');
        } else if (suffix) {
          setStatus(suffix, 'success');
        } else {
          setStatus('選択情報を更新しました。', 'success');
        }
      } catch (error) {
        if (!disposed) {
          setStatus(`取得に失敗しました: ${error.message}`, 'error');
        }
      } finally {
        if (!disposed) {
          button.disabled = false;
        }
      }
    };

    button.addEventListener('click', (event) => {
      event.preventDefault();
      fetchSelection();
    });

    applyStoredResult();
    if (node.config.lastWorkbook || node.config.lastSheet) {
      const workbookLabel = node.config.lastWorkbook || '';
      const sheetLabel = node.config.lastSheet || '';
      const address = node.config.SelectionAddress__raw || '';
      const summary = [workbookLabel, sheetLabel, address].filter(Boolean).join(' / ');
      if (summary) {
        setStatus(`前回の結果: ${summary}`, 'info');
      }
    } else {
      setStatus('Excelの選択範囲情報を取得できます。', 'info');
    }

    return () => {
      disposed = true;
    };
  },
});

export const BUILTIN_NODES = [
  createExcelWorkbookSelector(),
  createExcelSelectionInspector(),
];
