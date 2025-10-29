#requires -Version 5.1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$iconAnimationCode = @"
using System;
using System.Drawing;
using System.Windows.Forms;
using System.Drawing.Drawing2D;

public class IconDockableForm : Form
{
    private readonly Timer _animationTimer;
    private readonly Timer _hoverMonitor;
    private readonly double _animationDuration = 420.0; // milliseconds
    private DateTime _animationStart;
    private bool _isAnimating = false;
    private bool _isExpanding = false;
    private float _animationValue = 0f; // 0 collapsed, 1 expanded

    private readonly int _collapsedDiameter = 24;
    private readonly int _headerHeight = 40;
    private readonly Size _expandedSize = new Size(360, 420);
    private readonly Point _expandedLocation = new Point(120, 120);
    private readonly Color _iconColor = Color.FromArgb(200, 128, 0, 168);
    private readonly Color _headerColor = Color.FromArgb(230, 90, 0, 140);
    private readonly Color _contentColor = Color.FromArgb(245, 28, 28, 42);

    private Panel _headerPanel;
    private Panel _contentPanel;
    private Label _titleLabel;
    private Label _contentLabel;

    private Point _collapsedCenter;

    public IconDockableForm()
    {
        this.SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true);
        this.FormBorderStyle = FormBorderStyle.None;
        this.StartPosition = FormStartPosition.Manual;
        this.ShowInTaskbar = false;
        this.TopMost = true;
        this.BackColor = Color.FromArgb(10, 0, 0, 0);
        this.Opacity = 0.88;

        _collapsedCenter = new Point(_expandedLocation.X + _expandedSize.Width / 2, _expandedLocation.Y + _headerHeight / 2);
        Point collapsedLocation = new Point(_collapsedCenter.X - _collapsedDiameter / 2, _collapsedCenter.Y - _collapsedDiameter / 2);
        this.Location = collapsedLocation;
        this.ClientSize = new Size(_collapsedDiameter, _collapsedDiameter);

        CreateLayout();
        ApplyGeometry(0f);

        _animationTimer = new Timer();
        _animationTimer.Interval = 16;
        _animationTimer.Tick += AnimationTimerOnTick;

        _hoverMonitor = new Timer();
        _hoverMonitor.Interval = 60;
        _hoverMonitor.Tick += HoverMonitorOnTick;
        _hoverMonitor.Start();

        this.FormClosed += (s, e) =>
        {
            _hoverMonitor.Stop();
            _hoverMonitor.Dispose();
            _animationTimer.Stop();
            _animationTimer.Dispose();
        };
    }

    private void CreateLayout()
    {
        _headerPanel = new Panel();
        _headerPanel.BackColor = _headerColor;
        _headerPanel.ForeColor = Color.White;
        _headerPanel.Cursor = Cursors.Hand;

        _titleLabel = new Label();
        _titleLabel.Text = "Dockable Form";
        _titleLabel.Font = new Font("Segoe UI", 10f, FontStyle.Bold);
        _titleLabel.ForeColor = Color.White;
        _titleLabel.AutoSize = false;
        _titleLabel.TextAlign = ContentAlignment.MiddleLeft;
        _titleLabel.Padding = new Padding(12, 0, 0, 0);
        _titleLabel.Dock = DockStyle.Fill;
        _headerPanel.Controls.Add(_titleLabel);

        _contentPanel = new Panel();
        _contentPanel.BackColor = _contentColor;

        _contentLabel = new Label();
        _contentLabel.Text = "Hover over the circular icon to open the dockable form.\nThe form expands from the icon and draws downward.";
        _contentLabel.Font = new Font("Segoe UI", 9f);
        _contentLabel.ForeColor = Color.WhiteSmoke;
        _contentLabel.AutoSize = false;
        _contentLabel.TextAlign = ContentAlignment.MiddleCenter;
        _contentLabel.Dock = DockStyle.Fill;

        _contentPanel.Controls.Add(_contentLabel);

        this.Controls.Add(_contentPanel);
        this.Controls.Add(_headerPanel);
    }

    private void HoverMonitorOnTick(object sender, EventArgs e)
    {
        Rectangle bounds = this.Bounds;
        Point cursor = Cursor.Position;
        bool inside = bounds.Contains(cursor);

        if (inside && !_isAnimating && _animationValue < 1f)
        {
            StartExpand();
        }
        else if (!inside && !_isAnimating && _animationValue > 0f)
        {
            StartCollapse();
        }
    }

    private void StartExpand()
    {
        _isAnimating = true;
        _isExpanding = true;
        _animationStart = DateTime.Now;
        _animationTimer.Start();
    }

    private void StartCollapse()
    {
        _isAnimating = true;
        _isExpanding = false;
        _animationStart = DateTime.Now;
        _animationTimer.Start();
    }

    private void AnimationTimerOnTick(object sender, EventArgs e)
    {
        double elapsed = (DateTime.Now - _animationStart).TotalMilliseconds;
        double rawProgress = Math.Max(0.0, Math.Min(1.0, elapsed / _animationDuration));
        float eased = EaseOutCubic((float)rawProgress);

        if (_isExpanding)
        {
            _animationValue = eased;
        }
        else
        {
            _animationValue = 1f - eased;
        }

        ApplyGeometry(_animationValue);

        if (rawProgress >= 1.0)
        {
            _animationTimer.Stop();
            _isAnimating = false;
            _animationValue = _isExpanding ? 1f : 0f;
            ApplyGeometry(_animationValue);
        }
    }

    private static float EaseOutCubic(float t)
    {
        float inv = 1f - t;
        return 1f - inv * inv * inv;
    }

    private void ApplyGeometry(float state)
    {
        // state 0 collapsed, 1 expanded
        float widthProgress = Math.Min(1f, state / 0.5f);
        float heightStage = state < 0.5f ? state / 0.5f : (state - 0.5f) / 0.5f;

        float easedWidth = EaseOutCubic(widthProgress);
        float width = _collapsedDiameter + (_expandedSize.Width - _collapsedDiameter) * easedWidth;

        float height;
        if (state < 0.5f)
        {
            float easedHeight = EaseOutCubic(Math.Max(0f, Math.Min(1f, state / 0.5f)));
            height = _collapsedDiameter + (_headerHeight - _collapsedDiameter) * easedHeight;
        }
        else
        {
            float easedHeight = EaseOutCubic(Math.Max(0f, Math.Min(1f, heightStage)));
            height = _headerHeight + (_expandedSize.Height - _headerHeight) * easedHeight;
        }

        if (state <= 0.01f)
        {
            width = _collapsedDiameter;
            height = _collapsedDiameter;
        }
        else if (state >= 0.99f)
        {
            width = _expandedSize.Width;
            height = _expandedSize.Height;
        }

        int clientWidth = (int)Math.Round(width);
        int clientHeight = (int)Math.Round(height);

        if (state < 0.5f)
        {
            int x = _collapsedCenter.X - clientWidth / 2;
            int y = _collapsedCenter.Y - clientHeight / 2;
            this.Location = new Point(x, y);
        }
        else
        {
            int x = _collapsedCenter.X - clientWidth / 2;
            int y = _expandedLocation.Y;
            this.Location = new Point(x, y);
        }

        this.SuspendLayout();
        this.ClientSize = new Size(clientWidth, clientHeight);
        UpdateRegion(state, width, height);
        UpdateLayout(state, clientWidth, clientHeight);
        this.ResumeLayout();
        this.Invalidate();
    }

    private void UpdateLayout(float state, int width, int height)
    {
        int headerVisibleHeight = Math.Min(height, _headerHeight);
        _headerPanel.Visible = state > 0.05f;
        _headerPanel.Bounds = new Rectangle(0, 0, width, headerVisibleHeight);

        int contentHeight = Math.Max(0, height - _headerHeight);
        _contentPanel.Visible = contentHeight > 0;
        _contentPanel.Bounds = new Rectangle(0, _headerHeight, width, contentHeight);

        if (_contentPanel.Visible)
        {
            int padding = 16;
            _contentLabel.Padding = new Padding(padding);
        }
    }

    private void UpdateRegion(float state, float width, float height)
    {
        Region newRegion;
        if (state <= 0.05f)
        {
            using (GraphicsPath ellipsePath = new GraphicsPath())
            {
                ellipsePath.AddEllipse(0, 0, width, height);
                newRegion = new Region(ellipsePath);
            }
        }
        else if (height <= _headerHeight + 1f)
        {
            using (GraphicsPath capsulePath = new GraphicsPath())
            {
                float radius = height / 2f;
                capsulePath.AddArc(0, 0, height, height, 90, 180);
                capsulePath.AddArc(width - height, 0, height, height, 270, 180);
                capsulePath.CloseFigure();
                newRegion = new Region(capsulePath);
            }
        }
        else
        {
            using (GraphicsPath roundedRect = new GraphicsPath())
            {
                int radius = 12;
                roundedRect.StartFigure();
                roundedRect.AddArc(0, 0, radius * 2, radius * 2, 180, 90);
                roundedRect.AddArc(width - radius * 2, 0, radius * 2, radius * 2, 270, 90);
                roundedRect.AddArc(width - radius * 2, height - radius * 2, radius * 2, radius * 2, 0, 90);
                roundedRect.AddArc(0, height - radius * 2, radius * 2, radius * 2, 90, 90);
                roundedRect.CloseFigure();
                newRegion = new Region(roundedRect);
            }
        }

        if (this.Region != null)
        {
            this.Region.Dispose();
        }
        this.Region = newRegion;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

        using (SolidBrush brush = new SolidBrush(_iconColor))
        {
            e.Graphics.FillRegion(brush, this.Region);
        }

        if (_animationValue > 0.05f)
        {
            Rectangle headerRect = new Rectangle(0, 0, this.ClientSize.Width, Math.Min(this.ClientSize.Height, _headerHeight));
            using (SolidBrush headerBrush = new SolidBrush(Color.FromArgb(220, 100, 0, 160)))
            {
                e.Graphics.FillRectangle(headerBrush, headerRect);
            }
        }
    }

    [STAThread]
    public static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new IconDockableForm());
    }
}
"@;

Add-Type -TypeDefinition $iconAnimationCode -ReferencedAssemblies System.Windows.Forms, System.Drawing

[IconDockableForm]::Main()
