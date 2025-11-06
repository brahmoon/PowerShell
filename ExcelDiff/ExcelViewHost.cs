using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelDiff
{
    internal sealed class ExcelViewHost : IDisposable
    {
        private readonly Panel _hostPanel;
        private Excel.Application? _excelApp;
        private Excel.Workbook? _activeWorkbook;
        private string? _currentWorkbookPath;
        private IntPtr _excelWindowHandle;
        private bool _disposed;

        public ExcelViewHost(Panel hostPanel)
        {
            _hostPanel = hostPanel ?? throw new ArgumentNullException(nameof(hostPanel));
            _hostPanel.HandleCreated += HostPanelOnHandleCreated;
            _hostPanel.HandleDestroyed += HostPanelOnHandleDestroyed;
            _hostPanel.Resize += HostPanelOnResize;
        }

        public void LoadWorkbook(string workbookPath)
        {
            if (string.IsNullOrWhiteSpace(workbookPath))
            {
                throw new ArgumentException("Workbook path is required.", nameof(workbookPath));
            }

            EnsureExcelApplication();

            if (!string.Equals(_currentWorkbookPath, workbookPath, StringComparison.OrdinalIgnoreCase))
            {
                CloseWorkbook();
                _activeWorkbook = _excelApp!.Workbooks.Open(workbookPath, ReadOnly: true);
                _activeWorkbook.Activate();
                _currentWorkbookPath = workbookPath;
            }

            EmbedExcelWindow();
        }

        public void NavigateTo(string sheetName, int row, int column)
        {
            if (_activeWorkbook == null || _excelApp == null)
            {
                return;
            }

            Excel.Worksheet? worksheet = null;
            Excel.Range? target = null;
            try
            {
                worksheet = string.IsNullOrWhiteSpace(sheetName)
                    ? _activeWorkbook.ActiveSheet as Excel.Worksheet
                    : _activeWorkbook.Worksheets[sheetName] as Excel.Worksheet;

                worksheet?.Activate();

                if (row > 0 && column > 0)
                {
                    target = worksheet?.Cells[row, column] as Excel.Range;
                    target?.Select();

                    var window = _excelApp.ActiveWindow;
                    if (window != null)
                    {
                        window.ScrollRow = row;
                        window.ScrollColumn = column;
                    }
                }
            }
            finally
            {
                ReleaseCom(target);
                ReleaseCom(worksheet);
            }
        }

        public void ResizeEmbeddedWindow()
        {
            if (_excelWindowHandle != IntPtr.Zero && _hostPanel.IsHandleCreated)
            {
                NativeMethods.MoveWindow(_excelWindowHandle, 0, 0, _hostPanel.ClientSize.Width, _hostPanel.ClientSize.Height, true);
            }
        }

        private void EnsureExcelApplication()
        {
            if (_excelApp != null)
            {
                return;
            }

            _excelApp = new Excel.Application
            {
                Visible = true,
                DisplayAlerts = false,
                ScreenUpdating = true
            };

            _excelWindowHandle = new IntPtr(_excelApp.Hwnd);
            NativeMethods.ShowWindow(_excelWindowHandle, NativeMethods.SW_SHOW);
            EmbedExcelWindow();
        }

        private void EmbedExcelWindow()
        {
            if (_excelApp == null || _excelWindowHandle == IntPtr.Zero || !_hostPanel.IsHandleCreated)
            {
                return;
            }

            NativeMethods.SetParent(_excelWindowHandle, _hostPanel.Handle);
            NativeMethods.ApplyChildWindowStyle(_excelWindowHandle);
            NativeMethods.ShowWindow(_excelWindowHandle, NativeMethods.SW_SHOW);
            ResizeEmbeddedWindow();
            NativeMethods.SetFocus(_excelWindowHandle);

            _excelApp.Visible = true;
            var window = _excelApp.ActiveWindow;
            if (window != null)
            {
                window.WindowState = Excel.XlWindowState.xlNormal;
                window.DisplayWorkbookTabs = true;
                window.DisplayGridlines = true;
            }
        }

        private void HostPanelOnResize(object? sender, EventArgs e) => ResizeEmbeddedWindow();

        private void HostPanelOnHandleCreated(object? sender, EventArgs e) => EmbedExcelWindow();

        private void HostPanelOnHandleDestroyed(object? sender, EventArgs e)
        {
            if (_excelWindowHandle != IntPtr.Zero)
            {
                NativeMethods.SetParent(_excelWindowHandle, IntPtr.Zero);
            }
        }

        private void CloseWorkbook()
        {
            if (_activeWorkbook == null)
            {
                return;
            }

            try
            {
                _activeWorkbook.Close(false);
            }
            catch
            {
                // Ignore failures when closing workbooks.
            }
            finally
            {
                ReleaseCom(_activeWorkbook);
                _activeWorkbook = null;
                _currentWorkbookPath = null;
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;

            _hostPanel.HandleCreated -= HostPanelOnHandleCreated;
            _hostPanel.HandleDestroyed -= HostPanelOnHandleDestroyed;
            _hostPanel.Resize -= HostPanelOnResize;

            CloseWorkbook();

            if (_excelApp != null)
            {
                try
                {
                    _excelApp.Quit();
                }
                catch
                {
                    // Ignore cleanup failures during shutdown.
                }
                finally
                {
                    ReleaseCom(_excelApp);
                    _excelApp = null;
                }
            }
        }

        private static void ReleaseCom(object? comObject)
        {
            if (comObject is null)
            {
                return;
            }

            try
            {
                Marshal.FinalReleaseComObject(comObject);
            }
            catch
            {
                // Suppress any release errors to avoid crashing on shutdown.
            }
        }
    }
}
