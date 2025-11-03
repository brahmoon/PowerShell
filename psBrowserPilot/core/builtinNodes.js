import {
  DEFAULT_SERVER_URL,
  loadServerUrl,
  normalizeServerUrl,
} from '../export/scriptRunner.js';

const RUN_SCRIPT_PATH = '/runScript';

const RELEASE_COM_OBJECTS_SNIPPET = `
if ($sheet -ne $null) {
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
}
if ($workbook -ne $null) {
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
}
if ($excel -ne $null) {
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
`;

const COLOR_HELPERS = `
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

const LIST_WORKBOOKS_SCRIPT = `
$ErrorActionPreference = 'Stop'
$excel = $null
$workbook = $null
$sheet = $null
$workbooks = $null
$result = $null

try {
  try {
    $excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
  } catch {
    $excel = $null
  }

  if (-not $excel) {
    $result = [pscustomobject]@{
      ok = $false
      error = 'Excel が見つかりません。'
      workbooks = @()
    }
  } else {
    $workbooks = $excel.Workbooks
    $names = @()
    foreach ($wb in @($workbooks)) {
      if ($null -ne $wb) {
        $names += [string]$wb.Name
      }
    }
    $result = [pscustomobject]@{
      ok = $true
      workbooks = $names
    }
  }
} catch {
  $result = [pscustomobject]@{
    ok = $false
    error = $_.Exception.Message
    workbooks = @()
  }
} finally {
  if ($workbooks -ne $null) {
    foreach ($wb in @($workbooks)) {
      if ($null -ne $wb) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
      }
    }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbooks) | Out-Null
  }
  ${RELEASE_COM_OBJECTS_SNIPPET}
}

if ($result -eq $null) {
  $result = [pscustomobject]@{
    ok = $false
    error = '結果を取得できませんでした。'
    workbooks = @()
  }
}

$result | ConvertTo-Json -Compress
`;

const buildSelectionInspectorScript = (workbookLiteral) => `
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

${COLOR_HELPERS}

$excel = $null
$workbooks = $null
$workbook = $null
$sheet = $null
$selection = $null
$result = $null

try {
  try {
    $excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
  } catch {
    $excel = $null
  }

  if (-not $excel) {
    $result = [pscustomobject]@{
      ok = $false
      error = 'Excel が見つかりません。'
    }
  } else {
    $targetName = ${workbookLiteral || "''"}
    $workbooks = $excel.Workbooks

    if ([string]::IsNullOrWhiteSpace($targetName)) {
      $workbook = $excel.ActiveWorkbook
    } else {
      foreach ($candidate in @($workbooks)) {
        if ($null -eq $candidate) { continue }
        if ($candidate.Name -eq $targetName) {
          $workbook = $candidate
          break
        }
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($candidate) | Out-Null
      }
    }

    if (-not $workbook) {
      if ([string]::IsNullOrWhiteSpace($targetName)) {
        $result = [pscustomobject]@{
          ok = $false
          error = 'アクティブなブックが見つかりません。'
        }
      } else {
        $result = [pscustomobject]@{
          ok = $false
          error = "Workbook not found: $targetName"
        }
      }
    } else {
      try {
        $selection = $workbook.Application.Selection
        if ($null -eq $selection) {
          $result = [pscustomobject]@{
            ok = $false
            error = '選択範囲が見つかりません。'
            workbook = $workbook.Name
          }
        } else {
          $sheet = $selection.Worksheet
          $result = [pscustomobject]@{
            ok = $true
            workbook = $workbook.Name
            sheet = $sheet.Name
            address = $selection.Address()
            value = $selection.Text
            font_color = @{
              ole = $selection.Font.Color
              hex = Convert-OleToHex $selection.Font.Color
            }
            interior_color = @{
              ole = $selection.Interior.Color
              hex = Convert-OleToHex $selection.Interior.Color
            }
          }
        }
      } catch {
        $result = [pscustomobject]@{
          ok = $false
          error = $_.Exception.Message
          workbook = $workbook.Name
        }
      }
    }
  }
} catch {
  $result = [pscustomobject]@{
    ok = $false
    error = $_.Exception.Message
  }
} finally {
  if ($selection -ne $null) {
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($selection) | Out-Null
  }
  if ($workbooks -ne $null) {
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbooks) | Out-Null
  }
  ${RELEASE_COM_OBJECTS_SNIPPET}
}

if ($result -eq $null) {
  $result = [pscustomobject]@{
    ok = $false
    error = '結果を取得できませんでした。'
  }
}

$result | ConvertTo-Json -Compress
`;

const getRunScriptEndpoint = () => {
  const base = normalizeServerUrl(loadServerUrl()) || DEFAULT_SERVER_URL;
  return `${base.replace(/\/+$/, '')}${RUN_SCRIPT_PATH}`;
};

const invokePowerShell = async (script) => {
  const endpoint = getRunScriptEndpoint();
  const response = await fetch(endpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
    body: JSON.stringify({ script }),
  });

  let payload;
  try {
    payload = await response.json();
  } catch (error) {
    throw new Error('PowerShell サーバーから無効な応答を受信しました。');
  }

  if (!response.ok) {
    const message = payload?.error || `HTTP ${response.status}`;
    throw new Error(message);
  }

  if (payload?.ok === false) {
    const message = Array.isArray(payload?.errors) && payload.errors.length
      ? payload.errors.join('\n')
      : 'PowerShell がエラーを返しました。';
    throw new Error(message);
  }

  return typeof payload?.output === 'string' ? payload.output.trim() : '';
};

const requestExcelJson = async (script) => {
  const output = await invokePowerShell(script);
  if (!output) {
    return null;
  }
  try {
    return JSON.parse(output);
  } catch (error) {
    throw new Error('PowerShell の応答を解析できませんでした。');
  }
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
        const data = await requestExcelJson(LIST_WORKBOOKS_SCRIPT);
        if (disposed) return;
        if (!data?.ok) {
          throw new Error(data?.error || 'ブック情報を取得できませんでした');
        }
        const workbooks = Array.isArray(data.workbooks) ? data.workbooks : [];
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

    const fetchSelection = async () => {
      const workbook =
        resolveInput('WorkbookName', { preferRaw: true }) || node.config.WorkbookName__raw;
      button.disabled = true;
      setStatus('選択セル情報を取得中…', 'pending');
      try {
        const script = buildSelectionInspectorScript(workbook ? literal(workbook) : null);
        const data = await requestExcelJson(script);
        if (disposed) return;
        if (!data?.ok) {
          throw new Error(data?.error || '情報を取得できませんでした');
        }
        const address = data.address || '';
        const value = data.value ?? '';
        const fontHex = data.font_color?.hex || '';
        const interiorHex = data.interior_color?.hex || '';
        const resolvedWorkbook = data.workbook || workbook || '';

        setOutput('SelectionAddress', address);
        setOutput('SelectionText', value);
        setOutput('FontColorHex', fontHex);
        setOutput('InteriorColorHex', interiorHex);
        updateConfig('WorkbookName__raw', resolvedWorkbook, { silent: true });
        updateConfig('WorkbookName', resolvedWorkbook ? literal(resolvedWorkbook) : '', {
          silent: true,
        });
        updateConfig('lastWorkbook', resolvedWorkbook, { silent: true });
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
