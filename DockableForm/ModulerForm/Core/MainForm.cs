using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace DockableModularForm
{
    public class MainForm : Form
    {
        private readonly string _modulesPath;
        private readonly TabControl _tabControl;
        private IList<ModuleDescriptor> _modules = new List<ModuleDescriptor>();
        private readonly Timer _visibilityTimer;
        private readonly int _collapsedWidth = 12;
        private readonly int _expandedWidth = 420;
        private bool _isExpanded;

        public MainForm(string modulesPath)
        {
            _modulesPath = modulesPath;
            Text = "Modular Dockable Form";
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            TopMost = true;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            DoubleBuffered = true;

            _tabControl = new TabControl
            {
                Dock = DockStyle.Fill
            };
            Controls.Add(_tabControl);

            _visibilityTimer = new Timer
            {
                Interval = 250
            };
            _visibilityTimer.Tick += VisibilityTimer_Tick;

            Load += MainForm_Load;
            FormClosing += MainForm_FormClosing;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            Bounds = CalculateInitialBounds();
            Collapse();

            _modules = ModuleLoader.LoadModules(_modulesPath);
            if (_modules.Any())
            {
                foreach (var descriptor in _modules)
                {
                    try
                    {
                        var control = descriptor.Instance.GetControl() ?? new UserControl();
                        control.Dock = DockStyle.Fill;

                        var title = !string.IsNullOrWhiteSpace(descriptor.Instance.Name)
                            ? descriptor.Instance.Name
                            : descriptor.Manifest?.Name ?? descriptor.ModuleType.Name;

                        var tab = new TabPage(title);
                        tab.Controls.Add(control);
                        _tabControl.TabPages.Add(tab);

                        TryInvoke(() => descriptor.Instance.OnLoad());
                    }
                    catch (Exception ex)
                    {
                        var tab = new TabPage(descriptor.ModuleType?.Name ?? "Module");
                        tab.Controls.Add(new Label
                        {
                            Dock = DockStyle.Fill,
                            TextAlign = ContentAlignment.MiddleCenter,
                            Text = $"Failed to load module UI:\n{ex.Message}"
                        });
                        _tabControl.TabPages.Add(tab);
                    }
                }
            }
            else
            {
                var infoTab = new TabPage("Getting Started");
                infoTab.Controls.Add(new Label
                {
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Text = "Modules folder is empty.\nAdd DLL modules into the Modules directory."
                });
                _tabControl.TabPages.Add(infoTab);
            }

            _visibilityTimer.Start();
        }

        private Rectangle CalculateInitialBounds()
        {
            var workingArea = Screen.PrimaryScreen.WorkingArea;
            return new Rectangle(
                workingArea.Right - _expandedWidth,
                workingArea.Top,
                _expandedWidth,
                workingArea.Height);
        }

        private void VisibilityTimer_Tick(object sender, EventArgs e)
        {
            var cursorPosition = Cursor.Position;
            bool containsCursor = Bounds.Contains(cursorPosition);

            if (containsCursor && !_isExpanded)
            {
                Expand();
            }
            else if (!containsCursor && _isExpanded)
            {
                Collapse();
            }
        }

        private void Expand()
        {
            if (_isExpanded)
            {
                return;
            }

            var workingArea = Screen.PrimaryScreen.WorkingArea;
            Left = workingArea.Right - Width;
            Top = workingArea.Top;
            Height = workingArea.Height;
            _isExpanded = true;
        }

        private void Collapse()
        {
            if (!_isExpanded)
            {
                var workingArea = Screen.PrimaryScreen.WorkingArea;
                Width = _expandedWidth;
                Left = workingArea.Right - _collapsedWidth;
                Top = workingArea.Top;
                Height = workingArea.Height;
                return;
            }

            var targetLeft = Screen.PrimaryScreen.WorkingArea.Right - _collapsedWidth;
            Width = _expandedWidth;
            Left = targetLeft;
            _isExpanded = false;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (var module in _modules)
            {
                TryInvoke(() => module.Instance?.OnUnload());
            }
        }

        private static void TryInvoke(Action action)
        {
            try
            {
                action?.Invoke();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[MainForm] Module callback failed: {ex.Message}");
            }
        }
    }
}
