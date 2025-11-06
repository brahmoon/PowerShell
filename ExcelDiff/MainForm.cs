using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ExcelDiff
{
    public partial class MainForm : Form
    {
        private readonly ExcelViewHost _excelViewHost;
        private readonly OpenFileDialog _openFileDialog;
        private IReadOnlyList<DiffResult> _diffResults = Array.Empty<DiffResult>();
        private string? _leftWorkbookPath;
        private string? _rightWorkbookPath;

        public MainForm()
        {
            InitializeComponent();

            _excelViewHost = new ExcelViewHost(excelHostPanel);
            _openFileDialog = new OpenFileDialog
            {
                Filter = "Excel ブック (*.xlsx;*.xlsm;*.xlsb;*.xls)|*.xlsx;*.xlsm;*.xlsb;*.xls|すべてのファイル (*.*)|*.*",
                Multiselect = false,
                Title = "Excel ブックを選択"
            };

            compareButton.Click += CompareButton_Click;
            leftBrowseButton.Click += LeftBrowseButton_Click;
            rightBrowseButton.Click += RightBrowseButton_Click;
            resultsListView.SelectedIndexChanged += ResultsListView_SelectedIndexChanged;
            resultsListView.Resize += ResultsListView_Resize;
            FormClosing += MainForm_FormClosing;

            UpdateListViewColumnWidths();
        }

        private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
        {
            _excelViewHost.Dispose();
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            _excelViewHost.ResizeEmbeddedWindow();
        }

        private void LeftBrowseButton_Click(object? sender, EventArgs e)
        {
            if (TrySelectWorkbook(out var selectedPath))
            {
                leftPathTextBox.Text = selectedPath;
            }
        }

        private void RightBrowseButton_Click(object? sender, EventArgs e)
        {
            if (TrySelectWorkbook(out var selectedPath))
            {
                rightPathTextBox.Text = selectedPath;
            }
        }

        private bool TrySelectWorkbook(out string selectedPath)
        {
            selectedPath = string.Empty;
            if (_openFileDialog.ShowDialog(this) == DialogResult.OK && File.Exists(_openFileDialog.FileName))
            {
                selectedPath = _openFileDialog.FileName;
                return true;
            }

            return false;
        }

        private void CompareButton_Click(object? sender, EventArgs e)
        {
            var leftPath = leftPathTextBox.Text.Trim();
            var rightPath = rightPathTextBox.Text.Trim();

            if (!ValidateWorkbookPath(leftPath, "比較元") || !ValidateWorkbookPath(rightPath, "比較先"))
            {
                return;
            }

            if (string.Equals(leftPath, rightPath, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(this, "同じファイルを比較することはできません。別々のブックを選択してください。", "ExcelDiff", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                UseWaitCursor = true;
                compareButton.Enabled = false;
                placeholderLabel.Text = "差分を解析しています...";
                placeholderLabel.Visible = true;
                resultsListView.Items.Clear();

                _diffResults = ExcelDiffEngine.Compare(leftPath, rightPath);
                PopulateResultList(_diffResults);
                _leftWorkbookPath = leftPath;
                _rightWorkbookPath = rightPath;

                if (_diffResults.Count == 0)
                {
                    placeholderLabel.Text = "差分は検出されませんでした。";
                    placeholderLabel.Visible = true;
                }
                else
                {
                    placeholderLabel.Visible = false;
                    _excelViewHost.LoadWorkbook(leftPath);
                    resultsListView.Items[0].Selected = true;
                    resultsListView.Select();
                }
            }
            catch (Exception ex)
            {
                placeholderLabel.Text = "差分を選択するとExcelがここに表示されます";
                placeholderLabel.Visible = true;
                MessageBox.Show(this, FormattableString.Invariant($"比較中にエラーが発生しました:\n{ex.Message}"), "ExcelDiff", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                compareButton.Enabled = true;
                UseWaitCursor = false;
            }
        }

        private void PopulateResultList(IEnumerable<DiffResult> results)
        {
            resultsListView.BeginUpdate();
            resultsListView.Items.Clear();

            foreach (var result in results
                         .OrderBy(r => r.SheetName, StringComparer.OrdinalIgnoreCase)
                         .ThenBy(r => r.Row)
                         .ThenBy(r => r.Column))
            {
                var item = new ListViewItem(result.SheetName)
                {
                    Tag = result
                };
                item.SubItems.Add(result.Address);
                item.SubItems.Add(result.LeftValue);
                item.SubItems.Add(result.RightValue);
                resultsListView.Items.Add(item);
            }

            resultsListView.EndUpdate();
        }

        private void ResultsListView_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (resultsListView.SelectedItems.Count == 0)
            {
                return;
            }

            if (resultsListView.SelectedItems[0].Tag is not DiffResult selectedDiff)
            {
                return;
            }

            try
            {
                if (!string.IsNullOrEmpty(_leftWorkbookPath))
                {
                    _excelViewHost.LoadWorkbook(_leftWorkbookPath);
                    _excelViewHost.NavigateTo(selectedDiff.SheetName, selectedDiff.Row, selectedDiff.Column);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, FormattableString.Invariant($"Excel の表示中にエラーが発生しました:\n{ex.Message}"), "ExcelDiff", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ResultsListView_Resize(object? sender, EventArgs e)
        {
            UpdateListViewColumnWidths();
        }

        private void UpdateListViewColumnWidths()
        {
            if (resultsListView.Columns.Count < 4)
            {
                return;
            }

            var totalWidth = resultsListView.ClientSize.Width;
            if (totalWidth <= 0)
            {
                return;
            }

            var sheetWidth = Math.Max(120, totalWidth / 5);
            var addressWidth = Math.Max(80, totalWidth / 8);
            var remaining = totalWidth - sheetWidth - addressWidth;
            if (remaining < 200)
            {
                remaining = Math.Max(200, totalWidth - sheetWidth - addressWidth);
            }

            var leftWidth = Math.Max(150, remaining / 2);
            var rightWidth = Math.Max(150, remaining - leftWidth);

            var totalAssigned = sheetWidth + addressWidth + leftWidth + rightWidth;
            if (totalAssigned > totalWidth)
            {
                var overflow = totalAssigned - totalWidth;
                var adjustLeft = Math.Min(overflow / 2, leftWidth - 120);
                var adjustRight = overflow - adjustLeft;
                leftWidth -= adjustLeft;
                rightWidth -= adjustRight;
            }

            leftWidth = Math.Max(120, leftWidth);
            rightWidth = Math.Max(120, rightWidth);

            resultsListView.BeginUpdate();
            columnSheet.Width = sheetWidth;
            columnAddress.Width = addressWidth;
            columnLeft.Width = leftWidth;
            columnRight.Width = rightWidth;
            resultsListView.EndUpdate();
        }

        private static bool ValidateWorkbookPath(string path, string role)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                MessageBox.Show(FormattableString.Invariant($"{role}のブックを指定してください。"), "ExcelDiff", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            if (!File.Exists(path))
            {
                MessageBox.Show(FormattableString.Invariant($"{role}のブックが見つかりません:\n{path}"), "ExcelDiff", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }
    }
}
