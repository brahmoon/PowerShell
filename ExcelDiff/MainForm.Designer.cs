using System.Drawing;
using System.Windows.Forms;

namespace ExcelDiff
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components;
        private TableLayoutPanel headerLayout;
        private Label leftLabel;
        private Label rightLabel;
        private TextBox leftPathTextBox;
        private Button leftBrowseButton;
        private TextBox rightPathTextBox;
        private Button rightBrowseButton;
        private Button compareButton;
        private SplitContainer mainSplitContainer;
        private ListView resultsListView;
        private ColumnHeader columnSheet;
        private ColumnHeader columnAddress;
        private ColumnHeader columnLeft;
        private ColumnHeader columnRight;
        private Panel excelHostPanel;
        private Label placeholderLabel;

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                components?.Dispose();
            }

            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            headerLayout = new TableLayoutPanel();
            leftLabel = new Label();
            rightLabel = new Label();
            leftPathTextBox = new TextBox();
            leftBrowseButton = new Button();
            rightPathTextBox = new TextBox();
            rightBrowseButton = new Button();
            compareButton = new Button();
            mainSplitContainer = new SplitContainer();
            resultsListView = new ListView();
            columnSheet = new ColumnHeader();
            columnAddress = new ColumnHeader();
            columnLeft = new ColumnHeader();
            columnRight = new ColumnHeader();
            excelHostPanel = new Panel();
            placeholderLabel = new Label();
            ((System.ComponentModel.ISupportInitialize)mainSplitContainer).BeginInit();
            mainSplitContainer.Panel1.SuspendLayout();
            mainSplitContainer.Panel2.SuspendLayout();
            mainSplitContainer.SuspendLayout();
            SuspendLayout();
            // 
            // headerLayout
            // 
            headerLayout.AutoSize = true;
            headerLayout.ColumnCount = 5;
            headerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45F));
            headerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40F));
            headerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45F));
            headerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40F));
            headerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));
            headerLayout.Controls.Add(leftLabel, 0, 0);
            headerLayout.Controls.Add(rightLabel, 2, 0);
            headerLayout.Controls.Add(leftPathTextBox, 0, 1);
            headerLayout.Controls.Add(leftBrowseButton, 1, 1);
            headerLayout.Controls.Add(rightPathTextBox, 2, 1);
            headerLayout.Controls.Add(rightBrowseButton, 3, 1);
            headerLayout.Controls.Add(compareButton, 4, 1);
            headerLayout.Dock = DockStyle.Top;
            headerLayout.Location = new Point(0, 0);
            headerLayout.Margin = new Padding(12);
            headerLayout.Name = "headerLayout";
            headerLayout.Padding = new Padding(12, 12, 12, 0);
            headerLayout.RowCount = 2;
            headerLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            headerLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            headerLayout.Size = new Size(1164, 88);
            headerLayout.TabIndex = 0;
            // 
            // leftLabel
            // 
            leftLabel.AutoSize = true;
            leftLabel.Dock = DockStyle.Fill;
            leftLabel.Location = new Point(12, 12);
            leftLabel.Margin = new Padding(0, 0, 8, 4);
            leftLabel.Name = "leftLabel";
            leftLabel.Size = new Size(454, 20);
            leftLabel.TabIndex = 0;
            leftLabel.Text = "比較元ブック";
            // 
            // rightLabel
            // 
            rightLabel.AutoSize = true;
            rightLabel.Dock = DockStyle.Fill;
            rightLabel.Location = new Point(514, 12);
            rightLabel.Margin = new Padding(0, 0, 8, 4);
            rightLabel.Name = "rightLabel";
            rightLabel.Size = new Size(454, 20);
            rightLabel.TabIndex = 1;
            rightLabel.Text = "比較先ブック";
            // 
            // leftPathTextBox
            // 
            leftPathTextBox.Dock = DockStyle.Fill;
            leftPathTextBox.Margin = new Padding(0, 0, 8, 8);
            leftPathTextBox.Name = "leftPathTextBox";
            leftPathTextBox.PlaceholderText = "例: C:\\Data\\Left.xlsx";
            leftPathTextBox.Size = new Size(454, 27);
            leftPathTextBox.TabIndex = 2;
            // 
            // leftBrowseButton
            // 
            leftBrowseButton.AutoSize = true;
            leftBrowseButton.Margin = new Padding(0, 0, 12, 8);
            leftBrowseButton.Name = "leftBrowseButton";
            leftBrowseButton.Size = new Size(28, 29);
            leftBrowseButton.TabIndex = 3;
            leftBrowseButton.Text = "…";
            leftBrowseButton.UseVisualStyleBackColor = true;
            // 
            // rightPathTextBox
            // 
            rightPathTextBox.Dock = DockStyle.Fill;
            rightPathTextBox.Margin = new Padding(0, 0, 8, 8);
            rightPathTextBox.Name = "rightPathTextBox";
            rightPathTextBox.PlaceholderText = "例: C:\\Data\\Right.xlsx";
            rightPathTextBox.Size = new Size(454, 27);
            rightPathTextBox.TabIndex = 4;
            // 
            // rightBrowseButton
            // 
            rightBrowseButton.AutoSize = true;
            rightBrowseButton.Margin = new Padding(0, 0, 12, 8);
            rightBrowseButton.Name = "rightBrowseButton";
            rightBrowseButton.Size = new Size(28, 29);
            rightBrowseButton.TabIndex = 5;
            rightBrowseButton.Text = "…";
            rightBrowseButton.UseVisualStyleBackColor = true;
            // 
            // compareButton
            // 
            compareButton.Anchor = AnchorStyles.Right;
            compareButton.AutoSize = true;
            compareButton.Margin = new Padding(0, 0, 0, 8);
            compareButton.Name = "compareButton";
            compareButton.Padding = new Padding(12, 6, 12, 6);
            compareButton.Size = new Size(108, 34);
            compareButton.TabIndex = 6;
            compareButton.Text = "比較実行";
            compareButton.UseVisualStyleBackColor = true;
            // 
            // mainSplitContainer
            // 
            mainSplitContainer.Dock = DockStyle.Fill;
            mainSplitContainer.Location = new Point(0, 88);
            mainSplitContainer.Name = "mainSplitContainer";
            // 
            // mainSplitContainer.Panel1
            // 
            mainSplitContainer.Panel1.Controls.Add(resultsListView);
            mainSplitContainer.Panel1MinSize = 280;
            // 
            // mainSplitContainer.Panel2
            // 
            mainSplitContainer.Panel2.Controls.Add(excelHostPanel);
            mainSplitContainer.Size = new Size(1164, 612);
            mainSplitContainer.SplitterDistance = 380;
            mainSplitContainer.TabIndex = 1;
            // 
            // resultsListView
            // 
            resultsListView.Columns.AddRange(new[] { columnSheet, columnAddress, columnLeft, columnRight });
            resultsListView.Dock = DockStyle.Fill;
            resultsListView.FullRowSelect = true;
            resultsListView.GridLines = true;
            resultsListView.HideSelection = false;
            resultsListView.MultiSelect = false;
            resultsListView.Name = "resultsListView";
            resultsListView.Size = new Size(380, 612);
            resultsListView.TabIndex = 0;
            resultsListView.UseCompatibleStateImageBehavior = false;
            resultsListView.View = View.Details;
            // 
            // columnSheet
            // 
            columnSheet.Text = "シート";
            columnSheet.Width = 120;
            // 
            // columnAddress
            // 
            columnAddress.Text = "セル";
            columnAddress.Width = 80;
            // 
            // columnLeft
            // 
            columnLeft.Text = "比較元";
            columnLeft.Width = 120;
            // 
            // columnRight
            // 
            columnRight.Text = "比較先";
            columnRight.Width = 120;
            // 
            // excelHostPanel
            // 
            excelHostPanel.Controls.Add(placeholderLabel);
            excelHostPanel.Dock = DockStyle.Fill;
            excelHostPanel.Location = new Point(0, 0);
            excelHostPanel.Name = "excelHostPanel";
            excelHostPanel.Size = new Size(780, 612);
            excelHostPanel.TabIndex = 0;
            // 
            // placeholderLabel
            // 
            placeholderLabel.Dock = DockStyle.Fill;
            placeholderLabel.ForeColor = Color.DimGray;
            placeholderLabel.Location = new Point(0, 0);
            placeholderLabel.Name = "placeholderLabel";
            placeholderLabel.Padding = new Padding(32);
            placeholderLabel.Size = new Size(780, 612);
            placeholderLabel.TabIndex = 0;
            placeholderLabel.Text = "差分を選択するとExcelがここに表示されます";
            placeholderLabel.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1164, 700);
            Controls.Add(mainSplitContainer);
            Controls.Add(headerLayout);
            MinimumSize = new Size(920, 540);
            Name = "MainForm";
            Text = "ExcelDiff";
            mainSplitContainer.Panel1.ResumeLayout(false);
            mainSplitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)mainSplitContainer).EndInit();
            mainSplitContainer.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }
    }
}
