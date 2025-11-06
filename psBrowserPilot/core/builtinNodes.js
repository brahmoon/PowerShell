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

const buildActiveCellScript = (workbookLiteral) => `
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

${COLOR_HELPERS}

$excel = $null
$workbooks = $null
$workbook = $null
$sheet = $null
$cell = $null
$mergeArea = $null
$result = $null
$mergeRows = $null
$mergeColumns = $null
$mergeCount = $null

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
        $cell = $workbook.Application.ActiveCell
        if ($null -eq $cell) {
          $result = [pscustomobject]@{
            ok = $false
            error = 'アクティブなセルが見つかりません。'
            workbook = $workbook.Name
          }
        } else {
          $sheet = $cell.Worksheet
          try { $mergeArea = $cell.MergeArea } catch { $mergeArea = $null }
          $address = $cell.Address()
          $addressLocal = $cell.AddressLocal()
          $mergeAddress = $null
          $mergeAddressLocal = $null
          if ($mergeArea -ne $null) {
            try { $mergeAddress = $mergeArea.Address() } catch { $mergeAddress = $null }
            try { $mergeAddressLocal = $mergeArea.AddressLocal() } catch { $mergeAddressLocal = $null }
            try { $mergeRows = $mergeArea.Rows.Count } catch { $mergeRows = $null }
            try { $mergeColumns = $mergeArea.Columns.Count } catch { $mergeColumns = $null }
            try { $mergeCount = $mergeArea.Count } catch { $mergeCount = $null }
          }
          $result = [pscustomobject]@{
            ok = $true
            workbook = $workbook.Name
            sheet = $sheet?.Name
            address = if ($mergeAddress) { $mergeAddress } else { $address }
            address_local = if ($mergeAddressLocal) { $mergeAddressLocal } else { $addressLocal }
            selection_address = $address
            selection_address_local = $addressLocal
            text = $cell.Text
            font_color = @{
              ole = $cell.Font.Color
              hex = Convert-OleToHex $cell.Font.Color
            }
            interior_color = @{
              ole = $cell.Interior.Color
              hex = Convert-OleToHex $cell.Interior.Color
            }
            merge_area = if ($mergeArea -ne $null) {
              [pscustomobject]@{
                address = $mergeAddress
                address_local = $mergeAddressLocal
                rows = $mergeRows
                columns = $mergeColumns
                count = $mergeCount
              }
            } else { $null }
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
  if ($mergeArea -ne $null) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($mergeArea) | Out-Null }
  if ($cell -ne $null) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($cell) | Out-Null }
  if ($sheet -ne $null) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null }
  if ($workbooks -ne $null) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbooks) | Out-Null }
  if ($workbook -ne $null) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null }
  if ($excel -ne $null) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
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
  outputs: ['WorkbookName', 'Workbook'],
  initialConfig: {
    WorkbookName: '',
    WorkbookName__raw: '',
    Workbook: '',
    Workbook__raw: '',
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
      updateConfig('Workbook__raw', raw, { silent: true });
      updateConfig('Workbook', scriptValue, { silent: true });
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
      updateConfig('Workbook__raw', stored, { silent: true });
      updateConfig('Workbook', scriptValue, { silent: true });
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
  outputs: ['SelectionAddress', 'Sheet', 'SelectionText', 'FontColorHex', 'InteriorColorHex'],
  chainExecution: true,
  initialConfig: {
    WorkbookName__raw: '',
    SelectionAddress: '',
    SelectionAddress__raw: '',
    Sheet: '',
    Sheet__raw: '',
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
    runAuto,
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
      if (type) {
        status.dataset.state = type;
      } else if (status.dataset.state) {
        delete status.dataset.state;
      }
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

    const syncStatusFromConfig = () => {
      const workbookLabel = node.config.lastWorkbook || '';
      const sheetLabel = node.config.lastSheet || '';
      const address = node.config.SelectionAddress__raw || '';
      const summary = [workbookLabel, sheetLabel, address].filter(Boolean).join(' / ');
      if (summary) {
        setStatus(`前回の結果: ${summary}`, 'info');
      } else {
        setStatus('Excelの選択範囲情報を取得できます。', 'info');
      }
    };

    const runtime = node.runtime || (node.runtime = {});
    const runtimeState =
      (runtime.selectionInspector = {
        beginAuto: (message) => {
          if (disposed) return;
          button.disabled = true;
          setStatus(message || '選択セル情報を取得中…', 'pending');
        },
        finishAuto: () => {
          if (disposed) return;
          button.disabled = false;
        },
        setStatus: (message, state = '') => setStatus(message, state),
        applyResult: (data) => {
          if (disposed) return;
          updateDisplay(data || {});
        },
        syncStatus: syncStatusFromConfig,
      });

    button.addEventListener('click', async (event) => {
      event.preventDefault();
      runtimeState.beginAuto?.('選択セル情報を取得中…');
      try {
        if (typeof runAuto === 'function') {
          await runAuto({ includeUpstream: true });
        }
      } catch (error) {
        runtimeState.setStatus?.(`取得に失敗しました: ${error.message}`, 'error');
      } finally {
        runtimeState.finishAuto?.();
      }
    });

    applyStoredResult();
    runtimeState.syncStatus?.();

    return () => {
      disposed = true;
      if (node.runtime?.selectionInspector) {
        delete node.runtime.selectionInspector;
      }
    };
  },
  autoExecute: async ({ node, updateConfig, resolveInput, toPowerShellLiteral }) => {
    const runtime = node.runtime?.selectionInspector;
    runtime?.beginAuto?.('選択セル情報を取得中…');

    const setOutput = (key, rawValue) => {
      const raw = rawValue ?? '';
      updateConfig(`${key}__raw`, raw, { silent: true });
      updateConfig(key, raw ? toPowerShellLiteral(raw) : '', { silent: false });
    };

    let workbookRaw = '';
    if (typeof resolveInput === 'function') {
      try {
        workbookRaw =
          resolveInput('WorkbookName', { preferRaw: true }) ||
          resolveInput('WorkbookName', { preferRaw: false });
      } catch (error) {
        workbookRaw = '';
      }
    }
    if (!workbookRaw && typeof node.config.WorkbookName__raw === 'string') {
      workbookRaw = node.config.WorkbookName__raw;
    }
    if (typeof workbookRaw !== 'string') {
      workbookRaw = '';
    }

    const trimmedWorkbook = workbookRaw.trim();
    const workbookLiteral = trimmedWorkbook ? toPowerShellLiteral(trimmedWorkbook) : null;

    let data;
    try {
      const script = buildSelectionInspectorScript(workbookLiteral);
      data = await requestExcelJson(script);
    } catch (error) {
      runtime?.setStatus?.(`取得に失敗しました: ${error.message}`, 'error');
      runtime?.finishAuto?.();
      throw error;
    }

    if (!data?.ok) {
      const message = data?.error || '情報を取得できませんでした。';
      runtime?.setStatus?.(`取得に失敗しました: ${message}`, 'error');
      runtime?.finishAuto?.();
      throw new Error(message);
    }

    const address = data.address || '';
    const value = data.value ?? '';
    const fontHex = data.font_color?.hex || '';
    const interiorHex = data.interior_color?.hex || '';
    const sheetName = data.sheet || '';
    const resolvedWorkbook = data.workbook || trimmedWorkbook || '';

    setOutput('SelectionAddress', address);
    setOutput('Sheet', sheetName);
    setOutput('SelectionText', value);
    setOutput('FontColorHex', fontHex);
    setOutput('InteriorColorHex', interiorHex);

    updateConfig('WorkbookName__raw', resolvedWorkbook, { silent: true });
    updateConfig('WorkbookName', resolvedWorkbook ? toPowerShellLiteral(resolvedWorkbook) : '', {
      silent: true,
    });
    updateConfig('Workbook__raw', resolvedWorkbook, { silent: true });
    updateConfig('Workbook', resolvedWorkbook ? toPowerShellLiteral(resolvedWorkbook) : '', {
      silent: true,
    });
    updateConfig('lastWorkbook', resolvedWorkbook, { silent: true });
    updateConfig('lastSheet', sheetName, { silent: true });

    runtime?.applyResult?.({
      address,
      value,
      fontHex,
      interiorHex,
    });

    const workbookLabel = resolvedWorkbook;
    const sheetLabel = sheetName || '';
    const suffix = sheetLabel ? `${sheetLabel} - ${address}` : address;
    if (workbookLabel && suffix) {
      runtime?.setStatus?.(`${workbookLabel}: ${suffix}`, 'success');
    } else if (suffix) {
      runtime?.setStatus?.(suffix, 'success');
    } else {
      runtime?.setStatus?.('選択情報を更新しました。', 'success');
    }

    runtime?.finishAuto?.();
  },
});

const createExcelGetActiveCellNode = () => ({
  id: 'ui_excel_get_active_cell',
  label: 'Excel Get Active Cell',
  category: 'Excel UI',
  execution: 'ui',
  inputs: ['Workbook'],
  outputs: ['Sheet', 'Range'],
  initialConfig: {
    Workbook__raw: '',
    Workbook: '',
    Sheet__raw: '',
    Sheet: '',
    Range__raw: '',
    Range: '',
    StatusMessage: '',
    StatusState: '',
    lastWorkbook: '',
    lastSheet: '',
    lastAddress: '',
    lastUpdated: '',
  },
  script: () => '',
  render: ({ node, controls, runAuto }) => {
    if (!controls) {
      return null;
    }
    controls.innerHTML = '';

    const wrapper = document.createElement('div');
    wrapper.className = 'excel-ui-node active-cell-node';

    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'excel-ui-button primary';
    button.textContent = 'アクティブセルを取得';

    const status = document.createElement('div');
    status.className = 'excel-ui-status';

    const result = document.createElement('dl');
    result.className = 'excel-selection-result';

    const createItem = (label) => {
      const dt = document.createElement('dt');
      dt.textContent = label;
      const dd = document.createElement('dd');
      dd.textContent = '-';
      result.append(dt, dd);
      return dd;
    };

    const workbookEl = createItem('Workbook');
    const sheetEl = createItem('Sheet');
    const addressEl = createItem('適用範囲');
    const selectionEl = createItem('セル');
    const mergeEl = createItem('結合範囲');
    const valueEl = createItem('値');

    wrapper.append(button, status, result);
    controls.appendChild(wrapper);

    let disposed = false;

    const setStatus = (message, type = '') => {
      if (disposed) return;
      status.textContent = message || '';
      status.dataset.state = type;
    };

    const formatTextValue = (value) => {
      if (value === null || value === undefined) {
        return '';
      }
      const text = String(value);
      return text.trim() ? text : '';
    };

    const applyResult = (data = {}) => {
      if (disposed) return;
      workbookEl.textContent = data.workbook || '-';
      sheetEl.textContent = data.sheet || '-';
      addressEl.textContent = data.address || '-';
      selectionEl.textContent = data.selection || '-';
      const mergeLabel =
        (data.mergeArea && (data.mergeArea.address || data.mergeArea.Address)) ||
        (typeof data.mergeArea === 'string' ? data.mergeArea : '');
      mergeEl.textContent = mergeLabel || '-';
      const textValue = formatTextValue(data.text);
      valueEl.textContent = textValue || '-';
    };

    const parseStoredRange = () => {
      const raw = node.config.Range__raw;
      if (!raw || typeof raw !== 'string') {
        return {};
      }
      try {
        const parsed = JSON.parse(raw);
        return parsed && typeof parsed === 'object' ? parsed : {};
      } catch (error) {
        return {};
      }
    };

    const syncFromConfig = () => {
      const stored = parseStoredRange();
      applyResult({
        workbook: node.config.lastWorkbook || stored.Workbook || '',
        sheet: node.config.lastSheet || stored.Sheet || '',
        address: stored.Address || node.config.lastAddress || '',
        selection: stored.SelectionAddress || '',
        mergeArea: stored.MergeArea || '',
        text: stored.Text ?? stored.Value ?? '',
      });
      if (node.config.StatusMessage) {
        setStatus(node.config.StatusMessage, node.config.StatusState || 'info');
      } else if (node.config.lastSheet || node.config.lastAddress) {
        const summary = [
          node.config.lastWorkbook,
          node.config.lastSheet,
          node.config.lastAddress,
        ]
          .filter(Boolean)
          .join(' / ');
        setStatus(summary || 'アクティブセルを取得しました。', 'info');
      } else {
        setStatus('Excelのアクティブセル情報を取得します。', 'info');
      }
    };

    const runtime = node.runtime || (node.runtime = {});
    runtime.activeCell = {
      setStatus,
      applyResult: (data) => {
        applyResult({
          workbook: data?.workbook,
          sheet: data?.sheet,
          address: data?.address,
          selection: data?.selection,
          mergeArea: data?.mergeArea,
          text: data?.text,
        });
      },
    };

    button.addEventListener('click', async (event) => {
      event.preventDefault();
      button.disabled = true;
      setStatus('アクティブセルを取得中…', 'pending');
      try {
        if (typeof runAuto === 'function') {
          await runAuto({ includeUpstream: true });
        }
      } catch (error) {
        setStatus(`取得に失敗しました: ${error.message}`, 'error');
      } finally {
        button.disabled = false;
      }
    });

    syncFromConfig();

    return () => {
      disposed = true;
      if (node.runtime?.activeCell) {
        delete node.runtime.activeCell;
      }
    };
  },
  autoExecute: async ({ node, updateConfig, resolveInput, toPowerShellLiteral }) => {
    const runtime = node.runtime?.activeCell;
    runtime?.setStatus?.('アクティブセルを取得中…', 'pending');

    let workbookRaw = '';
    if (typeof resolveInput === 'function') {
      try {
        workbookRaw =
          resolveInput('Workbook', { preferRaw: true }) ||
          resolveInput('Workbook', { preferRaw: false });
      } catch (error) {
        workbookRaw = '';
      }
    }
    if (!workbookRaw && typeof node.config.Workbook__raw === 'string') {
      workbookRaw = node.config.Workbook__raw;
    }
    if (typeof workbookRaw !== 'string') {
      workbookRaw = '';
    }
    const trimmedWorkbook = workbookRaw.trim();
    const workbookLiteral = trimmedWorkbook ? toPowerShellLiteral(trimmedWorkbook) : null;

    let data;
    try {
      data = await requestExcelJson(buildActiveCellScript(workbookLiteral));
    } catch (error) {
      const message = error?.message || 'アクティブセル情報を取得できませんでした。';
      updateConfig('StatusMessage', message, { silent: true });
      updateConfig('StatusState', 'error', { silent: true });
      runtime?.setStatus?.(`取得に失敗しました: ${message}`, 'error');
      throw error;
    }

    if (!data?.ok) {
      const message = data?.error || 'アクティブセル情報を取得できませんでした。';
      updateConfig('StatusMessage', message, { silent: true });
      updateConfig('StatusState', 'error', { silent: true });
      runtime?.setStatus?.(`取得に失敗しました: ${message}`, 'error');
      throw new Error(message);
    }

    const workbookName = data.workbook || trimmedWorkbook || '';
    const sheetName = data.sheet || '';
    const selectionAddress =
      data.selection_address || data.selectionAddress || data.address || '';
    const address = data.address || selectionAddress || '';
    const mergeArea = data.merge_area || null;
    const textValue = data.text ?? data.value ?? '';

    const payload = {
      Workbook: workbookName,
      Sheet: sheetName,
      Address: address,
      AddressLocal: data.address_local || data.addressLocal || address,
      SelectionAddress: selectionAddress,
      SelectionAddressLocal:
        data.selection_address_local ||
        data.selectionAddressLocal ||
        data.address_local ||
        data.addressLocal ||
        selectionAddress,
      MergeArea: mergeArea,
      Text: textValue,
      Value: textValue,
    };
    if (data.font_color) {
      payload.FontColor = data.font_color;
    }
    if (data.interior_color) {
      payload.InteriorColor = data.interior_color;
    }

    const rawRange = JSON.stringify(payload);

    updateConfig('Workbook__raw', workbookName, { silent: true });
    updateConfig('Workbook', workbookName ? toPowerShellLiteral(workbookName) : '', {
      silent: true,
    });
    updateConfig('Sheet__raw', sheetName, { silent: true });
    updateConfig('Sheet', sheetName ? toPowerShellLiteral(sheetName) : '', {
      silent: false,
    });
    updateConfig('Range__raw', rawRange, { silent: true });
    updateConfig('Range', rawRange ? toPowerShellLiteral(rawRange) : '', {
      silent: false,
    });
    updateConfig('lastWorkbook', workbookName, { silent: true });
    updateConfig('lastSheet', sheetName, { silent: true });
    updateConfig('lastAddress', address, { silent: true });
    updateConfig('lastUpdated', new Date().toISOString(), { silent: true });

    const summaryParts = [workbookName, sheetName, address].filter(Boolean);
    const message = summaryParts.length
      ? `${summaryParts.join(' / ')} を取得しました。`
      : 'アクティブセルを取得しました。';
    updateConfig('StatusMessage', message, { silent: true });
    updateConfig('StatusState', 'success', { silent: true });

    runtime?.applyResult?.({
      workbook: workbookName,
      sheet: sheetName,
      address,
      selection: selectionAddress,
      mergeArea,
      text: textValue,
    });
    runtime?.setStatus?.(message, 'success');
  },
});

export const BUILTIN_NODES = [
  createExcelWorkbookSelector(),
  createExcelSelectionInspector(),
  createExcelGetActiveCellNode(),
];
