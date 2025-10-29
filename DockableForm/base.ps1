# C# code definition
$csCode = @"
using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Collections.Generic;
using System.Reflection;

// イージング関数を提供するクラス
public static class EasingFunctions
{
    // 線形の変化（デフォルト）
    public static float Linear(float t) { return t; }
    
    // イーズイン（徐々に加速）
    public static float EaseInQuad(float t) { return t * t; }
    public static float EaseInCubic(float t) { return t * t * t; }
    
    // イーズアウト（徐々に減速）
    public static float EaseOutQuad(float t) { return 1 - (1 - t) * (1 - t); }
    public static float EaseOutCubic(float t) { return 1 - (float)Math.Pow(1 - t, 3); }
    
    // イーズインアウト（加速して減速）
    public static float EaseInOutQuad(float t) { 
        if (t < 0.5f)
            return 2 * t * t; 
        else
            return 1 - (float)Math.Pow(-2 * t + 2, 2) / 2;
    }
    
    // バウンス効果（弾む）
    public static float EaseOutBounce(float t)
    {
        const float n1 = 7.5625f;
        const float d1 = 2.75f;
        
        if (t < 1 / d1)
            return n1 * t * t;
        else if (t < 2 / d1)
            return n1 * (t -= 1.5f / d1) * t + 0.75f;
        else if (t < 2.5 / d1)
            return n1 * (t -= 2.25f / d1) * t + 0.9375f;
        else
            return n1 * (t -= 2.625f / d1) * t + 0.984375f;
    }
    
    // 弾性（ゴム）効果
    public static float EaseOutElastic(float t)
    {
        const float c4 = (float)(2 * Math.PI) / 3;
        
        if (t == 0)
            return 0;
        if (t == 1)
            return 1;
        return (float)(Math.Pow(2, -10 * t) * Math.Sin((t * 10 - 0.75) * c4) + 1);
    }
    
    // イージング関数の種類の列挙型
    public enum EasingType
    {
        None,           // イージングなし（線形）
        Linear,         // 線形（同じく変化なし）
        EaseInQuad,     // 二次関数で加速
        EaseOutQuad,    // 二次関数で減速
        EaseInOutQuad,  // 二次関数で加速後減速
        EaseInCubic,    // 三次関数で加速
        EaseOutCubic,   // 三次関数で減速
        EaseOutBounce,  // バウンス効果
        EaseOutElastic  // 弾性効果
    }
    
    // 指定されたイージングタイプに基づいて値を計算
    public static float ApplyEasing(float t, EasingType easingType)
    {
        switch (easingType)
        {
            case EasingType.None:
            case EasingType.Linear: 
                return Linear(t);
            case EasingType.EaseInQuad: 
                return EaseInQuad(t);
            case EasingType.EaseOutQuad: 
                return EaseOutQuad(t);
            case EasingType.EaseInOutQuad: 
                return EaseInOutQuad(t);
            case EasingType.EaseInCubic: 
                return EaseInCubic(t);
            case EasingType.EaseOutCubic: 
                return EaseOutCubic(t);
            case EasingType.EaseOutBounce: 
                return EaseOutBounce(t);
            case EasingType.EaseOutElastic: 
                return EaseOutElastic(t);
            default: 
                return Linear(t);
        }
    }
}

// ダブルバッファリング対応のPanel
public class DoubleBufferedPanel : Panel
{
    public DoubleBufferedPanel()
    {
        this.DoubleBuffered = true;
        this.SetStyle(ControlStyles.AllPaintingInWmPaint, true);
        this.SetStyle(ControlStyles.OptimizedDoubleBuffer, true);
        this.UpdateStyles();
    }
}

// アニメーションエンジンの基本クラス（C# 5.0対応版 - 軽量化）
public class AnimationEngine
{
    private Timer _timer;
    private int _currentStep = 0;
    private int _totalSteps;
    private EasingFunctions.EasingType _easingType;
    private bool _isCompleted = false;
    private Action<float> _updateAction;
    private Action _completedAction;
    
    public bool IsRunning { get; private set; }
    
    // アニメーション速度の最適化（FPSを30に削減）
    public AnimationEngine(int durationMs, int fps = 30, EasingFunctions.EasingType easingType = EasingFunctions.EasingType.None)
    {
        _totalSteps = (int)(durationMs / (1000.0 / fps));
        _easingType = easingType;
        _timer = new Timer();
        _timer.Interval = (int)(1000.0 / fps);
        _timer.Tick += Timer_Tick;
    }
    
    private void Timer_Tick(object sender, EventArgs e)
    {
        if (_isCompleted)
        {
            Stop();
            if (_completedAction != null)
            {
                _completedAction();
            }
            return;
        }
        
        _currentStep++;
        if (_currentStep >= _totalSteps)
        {
            _currentStep = _totalSteps;
            _isCompleted = true;
        }
        
        float progress = (float)_currentStep / _totalSteps;
        float easedProgress = EasingFunctions.ApplyEasing(progress, _easingType);
        
        if (_updateAction != null)
        {
            _updateAction(easedProgress);
        }
    }
    
    public void Start(Action<float> updateAction, Action completedAction = null)
    {
        if (IsRunning) return;
        
        _currentStep = 0;
        _isCompleted = false;
        _updateAction = updateAction;
        _completedAction = completedAction;
        
        IsRunning = true;
        _timer.Start();
    }
    
    public void Stop()
    {
        if (!IsRunning) return;
        
        _timer.Stop();
        IsRunning = false;
    }
    
    // 現在のプログレス値を返す (0.0 〜 1.0)
    public float GetCurrentProgress()
    {
        if (_totalSteps == 0) return 0;
        return (float)_currentStep / _totalSteps;
    }
    
    // アニメーションの方向を反転する
    public void ReverseAnimation(Action<float> updateAction, Action completedAction = null)
    {
        if (!IsRunning) return;
        
        // 現在の進行度から逆方向に進む
        int newStartStep = _totalSteps - _currentStep;
        
        Stop();
        
        _currentStep = newStartStep;
        _isCompleted = false;
        _updateAction = updateAction;
        _completedAction = completedAction;
        
        IsRunning = true;
        _timer.Start();
    }
}

// ドックポジションの列挙型
public enum DockPosition
{
    Top,
    Left,
    Right,
    None
}

// ピン留めモードの列挙型
public enum PinMode
{
    None,           // ピン留めなし
    Pinned,         // ピン留め
    PinnedTopMost   // ピン留め＆最前面表示
}

// アニメーション付きDockableFormクラス（C# 5.0対応版 - 改良版）
public class AnimatedDockableForm : Form
{
    // Win32 API for mouse tracking
    [DllImport("user32.dll")]
    static extern bool GetCursorPos(out POINT lpPoint);

    // ちらつき防止のためのWin32 API
    [DllImport("user32.dll")]
    public static extern int SendMessage(IntPtr hWnd, Int32 wMsg, bool wParam, Int32 lParam);
    private const int WM_SETREDRAW = 11;

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int X;
        public int Y;
    }

    // Properties
    private PinMode _pinMode = PinMode.None;
    private bool _isFullScreen = false;
    private bool _showStandardWindowButtons = false;
    private Timer _mouseTrackTimer = new Timer();
    private Timer _hideDelayTimer = new Timer();
    private Timer _showDelayTimer = new Timer();
    private bool _mouseLeavePending = false;
    private bool _showDelayPending = false;
    private int _hideDelayMilliseconds = 1000; // 1秒の遅延
    private int _showDelayMilliseconds = 1000; // 1秒の遅延
    private NotifyIcon _notifyIcon = new NotifyIcon();
    private Point _dragStartPoint;
    private bool _isDragging = false;
    private DockPosition _dockPosition = DockPosition.Top;
    private Rectangle _triggerArea;
    private int _triggerWidth = 500;
    private int _triggerHeight = 1;
    private int _triggerX = 0;
    private int _triggerY = 0;
    private int _formLeftPosition;
    private int _originalHeight;
    private int _originalWidth;
    private bool _isInitialLoad = true;
    private bool _isClosingAnimation = false;
    private bool _shouldHideWhenAnimationComplete = false;
    private AnimationEngine _animationEngine;
    private Size _dockSize;
    private Size _fullScreenSize;
    private Point _dockLocation;
    private bool _suspendLayout = false;
    private bool _isResizing = false;

    
    // リサイズ方向の列挙型
    private enum ResizeDirection
    {
        None,
        Left,
        Right,
        Bottom,
        BottomLeft,
        BottomRight,
        Top,
        TopLeft,
        TopRight
    }

    // リサイズ関連のフィールド
    private ResizeDirection _currentResizeDirection = ResizeDirection.None;
    private Point _resizeStartPoint;
    private Size _originalResizeSize;
    private bool _allowTopResize = false; // 上部リサイズを許可するかどうかのフラグ
    private const int RESIZE_BORDER_SIZE = 6; // リサイズ可能なボーダーの幅

    // CreateParams のオーバーライド
    protected override CreateParams CreateParams
    {
        get
        {
            CreateParams cp = base.CreateParams;
            cp.ExStyle |= 0x02000000; // WS_EX_COMPOSITED フラグを追加
            return cp;
        }
    }

    // UI Controls
    private Button _pinButton;
    private Button _fullscreenButton;
    private DoubleBufferedPanel _contentPanel;
    private Label _statusLabel;
    private Panel _headerPanel;

    // Public properties
    public PinMode PinMode
    {
        get { return _pinMode; }
        set { 
            _pinMode = value;
            UpdatePinModeSettings();
        }
    }
    
    public bool IsFullScreen
    {
        get { return _isFullScreen; }
        set { 
            if (_isFullScreen != value)
            {
                _isFullScreen = value;
                UpdateFullScreenButtonAppearance();
                ToggleFullScreenMode();
            }
        }
    }

    public bool ShowStandardWindowButtons
    {
        get { return _showStandardWindowButtons; }
        set {
            _showStandardWindowButtons = value;
            UpdateWindowButtonsVisibility();
        }
    }

    public DockPosition DockPosition
    {
        get { return _dockPosition; }
        set {
            _dockPosition = value;
            UpdateDockPosition();
        }
    }

    public int TriggerWidth
    {
        get { return _triggerWidth; }
        set {
            _triggerWidth = value;
            UpdateTriggerArea();
        }
    }
    
    public int TriggerHeight
    {
        get { return _triggerHeight; }
        set {
            _triggerHeight = value;
            UpdateTriggerArea();
        }
    }
    
    public Panel ContentPanel
    {
        get { return _contentPanel; }
    }
    
    public Label StatusLabel
    {
        get { return _statusLabel; }
    }
    
    // 上部リサイズを許可するかどうかのプロパティ
    public bool AllowTopResize
    {
        get { return _allowTopResize; }
        set { _allowTopResize = value; }
    }

    // アニメーション時間の短縮プロパティ（ms）
    // デフォルト値は300（元の値から短縮）
    private int _animationDuration = 300; // フィールドとして定義し初期化
    public int AnimationDuration
    {
        get { return _animationDuration; }
        set { _animationDuration = value; }
    }
    
    // 遅延時間のカスタマイズ用プロパティ
    public int HideDelayMilliseconds
    {
        get { return _hideDelayMilliseconds; }
        set
        {
            _hideDelayMilliseconds = value;
            _hideDelayTimer.Interval = _hideDelayMilliseconds;
        }
    }

    public void SetTriggerArea(int x, int y, int width, int height)
    {
        _triggerX = x;
        _triggerY = y;
        _triggerWidth = width;
        _triggerHeight = height;
        UpdateTriggerArea();
    }
    
    public void CenterTriggerAreaOnScreen()
    {
        _triggerX = (Screen.PrimaryScreen.WorkingArea.Width - this.Width) / 2;
        _triggerY = 0;
        UpdateTriggerArea();
    }

    public AnimatedDockableForm()
    {
        // 初期化処理を最適化された順序で行う
        InitializeForm();
        InitializeComponents();
        InitializeHideDelay();
        InitializeShowDelay();
        
        // リサイズ機能とマウストラッキングの初期化
        InitializeResizeCapability();
        SetupMouseTracking();
        SetupTrayIcon();
        
        // アニメーションエンジンを最適化して初期化
        _animationEngine = new AnimationEngine(_animationDuration, 30, EasingFunctions.EasingType.EaseOutCubic);
        
        // フォームロード時のイベント追加
        this.Load += AnimatedDockableForm_Load;
        this.Shown += AnimatedDockableForm_Shown;
        
        // サイズ設定の保存
        _dockSize = new Size(500, 300);
        _fullScreenSize = new Size(Screen.PrimaryScreen.WorkingArea.Width, Screen.PrimaryScreen.WorkingArea.Height);
        
        // 上部リサイズを無効化（デフォルト設定）
        AllowTopResize = false;
        
        // リサイズ中の更新抑制
        this.ResizeBegin += (s, e) => {
            if (!_isClosingAnimation && !_isInitialLoad)
            {
                _isResizing = true;
                // リサイズ中の描画を一時停止
                SendMessage(this.Handle, WM_SETREDRAW, false, 0);
            }
        };
        
        this.ResizeEnd += (s, e) => {
            if (_isResizing)
            {
                _isResizing = false;
                // 描画再開
                SendMessage(this.Handle, WM_SETREDRAW, true, 0);
                this.Refresh();
            }
        };
    }
    
    private void AnimatedDockableForm_Load(object sender, EventArgs e)
    {
        // 元のサイズを保存
        _originalHeight = this.Height;
        _originalWidth = this.Width;
        _dockLocation = this.Location;

        // トリガーエリアの幅をフォームの幅に合わせる
        _triggerWidth = this.Width;
        
        // 初回表示時は高さ0から開始
        if (_isInitialLoad)
        {
            this.Height = 0;
        }
    }
    
    private void AnimatedDockableForm_Shown(object sender, EventArgs e)
    {
        // 初回表示時のみアニメーション
        if (_isInitialLoad)
        {
            // 高さアニメーション開始
            StartHeightAnimation(true);
            _isInitialLoad = false;
        }
    }
    
    private void InitializeHideDelay()
    {
        // 遅延タイマーの設定
        _hideDelayTimer.Interval = _hideDelayMilliseconds;
        _hideDelayTimer.Tick += HideDelayTimer_Tick;
        _hideDelayTimer.Enabled = false;
    }

    private void InitializeShowDelay()
    {
        _showDelayTimer.Interval = _showDelayMilliseconds;
        _showDelayTimer.Tick += ShowDelayTimer_Tick;
        _showDelayTimer.Enabled = false;
    }

    // 遅延タイマーのTickイベントハンドラ
    private void HideDelayTimer_Tick(object sender, EventArgs e)
    {
        // タイマーを停止
        _hideDelayTimer.Stop();
        
        // 遅延後もまだマウスがフォーム外にあるか確認
        if (_mouseLeavePending && !IsMouseInForm() && this.Visible && !_animationEngine.IsRunning)
        {
            if (_showDelayTimer.Enabled || _showDelayPending)
            {
                _showDelayTimer.Stop();
                _showDelayPending = false;
            }

            // アニメーションで閉じる
            _shouldHideWhenAnimationComplete = true;
            StartHeightAnimation(false, this.Height);
            _statusLabel.Text = "フォームの外に出ました (1秒経過)";

            // 保留フラグをクリア
            _mouseLeavePending = false;
        }
        else
        {
            // マウスが戻ってきたか、他の条件により閉じないことに
            _mouseLeavePending = false;
        }
    }
    
    private void ShowDelayTimer_Tick(object sender, EventArgs e)
    {
        _showDelayTimer.Stop();

        if (!_showDelayPending)
            return;

        _showDelayPending = false;

        if (_pinMode != PinMode.None || _isFullScreen)
            return;

        if (!this.Visible && IsMouseInTriggerArea())
        {
            int currentHeight = Math.Max(0, Math.Min(this.Height, _originalHeight));
            this.Height = currentHeight;
            this.Show();
            this.BringToFront();
            StartHeightAnimation(true, currentHeight);
            _statusLabel.Text = "トリガーエリアに入りました (遅延後に開きました)";
        }
    }


    private void StopAnimationAndResumeLayoutIfNeeded()
    {
        if (_animationEngine != null && _animationEngine.IsRunning)
        {
            _animationEngine.Stop();
        }

        if (_suspendLayout)
        {
            ResumeLayout(true);
            _suspendLayout = false;
        }
    }

    private void StartHeightAnimation(bool isOpening, int? startHeightOverride = null)
    {
        // アニメーションが実行中なら一旦停止
        StopAnimationAndResumeLayoutIfNeeded();

        // 閉じるアニメーション実行中フラグの設定
        _isClosingAnimation = !isOpening;

        int initialHeight = startHeightOverride.HasValue ? startHeightOverride.Value : this.Height;
        initialHeight = Math.Max(0, Math.Min(initialHeight, _originalHeight));
        int targetHeight = isOpening ? _originalHeight : 0;

        if (initialHeight == targetHeight)
        {
            this.Height = targetHeight;
            _isClosingAnimation = false;

            if (!isOpening && _shouldHideWhenAnimationComplete)
            {
                this.Hide();
                _shouldHideWhenAnimationComplete = false;
            }

            return;
        }

        bool resumeLayoutAfterAnimation = false;

        if (!_suspendLayout)
        {
            // リサイズ中のレイアウト更新を一時停止
            SuspendLayout();
            _suspendLayout = true;
            resumeLayoutAfterAnimation = true;
        }

        _shouldHideWhenAnimationComplete = isOpening ? false : _shouldHideWhenAnimationComplete;

        _animationEngine.Start(
            // 更新処理
            (progress) => {
                int interpolatedHeight = (int)(initialHeight + (targetHeight - initialHeight) * progress);
                this.Height = interpolatedHeight;

                if (isOpening && _dockPosition == DockPosition.Top && !_isFullScreen)
                {
                    this.Location = new Point(this.Location.X, 0);
                }
            },
            // 完了時の処理
            () => {
                this.Height = targetHeight;
                _isClosingAnimation = false;

                if (!isOpening && _shouldHideWhenAnimationComplete)
                {
                    this.Hide();
                    _shouldHideWhenAnimationComplete = false;
                }

                _statusLabel.Text = isOpening ? "アニメーション完了" : "閉じるアニメーション完了";

                if (resumeLayoutAfterAnimation && _suspendLayout)
                {
                    ResumeLayout(true);
                    _suspendLayout = false;
                }
            }
        );
    }
    
    private void ToggleFullScreenMode()
    {
        // アニメーションが実行中なら一旦停止
        StopAnimationAndResumeLayoutIfNeeded();
        
        // リサイズ中のレイアウト更新を一時停止
        SuspendLayout();
        _suspendLayout = true;
        
        // 現在のサイズと位置を記録
        Rectangle startBounds = this.Bounds;
        Rectangle targetBounds;
        
        if (_isFullScreen)
        {
            // フルスクリーンサイズをターゲットに設定
            targetBounds = new Rectangle(
                0, 0, 
                _fullScreenSize.Width, 
                _fullScreenSize.Height
            );
            
            // ヘッダーパネルは常に表示
            _headerPanel.Visible = true;
        }
        else
        {
            // ドックサイズに戻す
            targetBounds = new Rectangle(
                _dockLocation.X, _dockLocation.Y,
                _dockSize.Width, _dockSize.Height
            );
        }
        
        // サイズ変更アニメーション
        _animationEngine.Start(
            // 更新処理
            (progress) => {
                int width = startBounds.Width + (int)((targetBounds.Width - startBounds.Width) * progress);
                int height = startBounds.Height + (int)((targetBounds.Height - startBounds.Height) * progress);
                int x = startBounds.X + (int)((targetBounds.X - startBounds.X) * progress);
                int y = startBounds.Y + (int)((targetBounds.Y - startBounds.Y) * progress);
                
                this.Bounds = new Rectangle(x, y, width, height);
            },
            // 完了時の処理
            () => {
                this.Bounds = targetBounds;
                if (!_isFullScreen)
                {
                    // ドックモードに戻る場合は元のサイズを保持
                    _originalHeight = _dockSize.Height;
                    _originalWidth = _dockSize.Width;
                    UpdateDockPosition();
                }
                _statusLabel.Text = _isFullScreen ? "フルスクリーンモードに切り替えました" : "ドックモードに切り替えました";
                
                // レイアウト更新を再開
                if (_suspendLayout)
                {
                    ResumeLayout(true);
                    _suspendLayout = false;
                }
            }
        );
    }

    private void InitializeForm()
    {
        // Basic form settings
        this.FormBorderStyle = FormBorderStyle.None;
        this.StartPosition = FormStartPosition.Manual;
        this.ShowInTaskbar = false;
        this.Size = new Size(500, 300);
        this.BackColor = Color.White;
        this.Opacity = 0.97;
        
        // ダブルバッファリングを有効化
        this.SetStyle(ControlStyles.DoubleBuffer, true);
        this.SetStyle(ControlStyles.UserPaint, true);
        this.SetStyle(ControlStyles.AllPaintingInWmPaint, true);
        this.SetStyle(ControlStyles.OptimizedDoubleBuffer, true);
        this.SetStyle(ControlStyles.ResizeRedraw, true);
        
        // Initialize position
        DockPosition = DockPosition.Top;
        CenterTriggerAreaOnScreen();
    }

    private void InitializeComponents()
    {
        // Create header panel
        _headerPanel = new Panel();
        _headerPanel.Dock = DockStyle.Top;
        _headerPanel.Height = 30;
        _headerPanel.BackColor = Color.FromArgb(40, 40, 40);
        
        // Title label
        Label titleLabel = new Label();
        titleLabel.Text = "DockForm";
        titleLabel.ForeColor = Color.White;
        titleLabel.Location = new Point(10, 7);
        titleLabel.AutoSize = true;
        _headerPanel.Controls.Add(titleLabel);
        
        // ボタンを右側に配置するパネル
        Panel buttonPanel = new Panel();
        buttonPanel.Dock = DockStyle.Right;
        buttonPanel.Height = _headerPanel.Height;
        buttonPanel.Width = 80; // ボタン3つ分の幅
        buttonPanel.BackColor = Color.Transparent;
        _headerPanel.Controls.Add(buttonPanel);
        
        // Pin button
        _pinButton = new Button();
        _pinButton.Size = new Size(24, 24);
        _pinButton.FlatStyle = FlatStyle.Flat;
        _pinButton.FlatAppearance.BorderSize = 0;
        _pinButton.Text = "📌";
        _pinButton.Click += new EventHandler(PinButton_Click);
        _pinButton.Location = new Point(5, 3);
        _pinButton.Cursor = Cursors.Hand;
        _pinButton.BackColor = Color.Transparent;
        _pinButton.ForeColor = Color.White;
        buttonPanel.Controls.Add(_pinButton);
        
        // FullScreen button right after topmost button
        _fullscreenButton = new Button();
        _fullscreenButton.Size = new Size(24, 24);
        _fullscreenButton.FlatStyle = FlatStyle.Flat;
        _fullscreenButton.FlatAppearance.BorderSize = 0;
        _fullscreenButton.Text = "🔍";
        _fullscreenButton.Click += new EventHandler(FullscreenButton_Click);
        _fullscreenButton.Location = new Point(30, 3);
        _fullscreenButton.Cursor = Cursors.Hand;
        _fullscreenButton.BackColor = Color.Transparent;
        _fullscreenButton.ForeColor = Color.White;
        buttonPanel.Controls.Add(_fullscreenButton);
        
        // Close button
        Button closeButton = new Button();
        closeButton.Size = new Size(24, 24);
        closeButton.FlatStyle = FlatStyle.Flat;
        closeButton.FlatAppearance.BorderSize = 0;
        closeButton.Text = "✕";
        closeButton.Click += (s, e) => { 
            if (_isClosingAnimation || _isFullScreen)
            {
                // フルスクリーンモードまたはアニメーション実行中は即時非表示
                this.Hide();
            }
            else
            {
                // アニメーションで徐々に閉じる
                _shouldHideWhenAnimationComplete = true;
                StartHeightAnimation(false, this.Height);
            }
        };
        closeButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        closeButton.Location = new Point(this.Width - 30, 3);
        closeButton.Cursor = Cursors.Hand;
        closeButton.BackColor = Color.Transparent;
        closeButton.ForeColor = Color.White;
        buttonPanel.Controls.Add(closeButton);
        
        // Make header draggable
        _headerPanel.MouseDown += new MouseEventHandler(Form_MouseDown);
        _headerPanel.MouseMove += new MouseEventHandler(Form_MouseMove);
        _headerPanel.MouseUp += new MouseEventHandler(Form_MouseUp);
        
        this.Controls.Add(_headerPanel);
        
        // Content panel - ダブルバッファリング対応のパネルを使用
        _contentPanel = new DoubleBufferedPanel();
        _contentPanel.Dock = DockStyle.Fill;
        _contentPanel.Padding = new Padding(10, 40, 10, 10);
        _contentPanel.BackColor = Color.White;
        this.Controls.Add(_contentPanel);
        
        // Status bar
        Panel statusBar = new Panel();
        statusBar.Dock = DockStyle.Bottom;
        statusBar.Height = 25;
        statusBar.BackColor = Color.FromArgb(240, 240, 240);
        
        // Status label
        _statusLabel = new Label();
        _statusLabel.AutoSize = true;
        _statusLabel.Location = new Point(5, 5);
        _statusLabel.Text = "Ready";
        statusBar.Controls.Add(_statusLabel);
        
        this.Controls.Add(statusBar);
    }
    
    // リサイズ機能の初期化
    private void InitializeResizeCapability()
    {
        // リサイズ時のマウスイベントを追加（フォーム自体）
        this.MouseMove += Form_ResizeMouseMove;
        
        // コンテントパネルにもイベントを追加
        _contentPanel.MouseMove += Form_ResizeMouseMove;
                
        if (_statusLabel != null && _statusLabel.Parent != null)
        {
            _statusLabel.Parent.MouseMove += Form_ResizeMouseMove;
        }
        
        SetupMouseHandlers();
    }

            
    private void SetupMouseHandlers()
    {
        // フォーム自体にイベントハンドラを割り当て
        this.MouseDown += Form_MouseDown;
        this.MouseMove += Form_MouseMove;
        this.MouseUp += Form_MouseUp;
        
        // コンテントパネルにイベントハンドラを割り当て
        _contentPanel.MouseDown += Form_MouseDown;
        _contentPanel.MouseMove += Form_MouseMove;
        _contentPanel.MouseUp += Form_MouseUp;
        
        if (_statusLabel != null && _statusLabel.Parent != null)
        {
            _statusLabel.Parent.MouseDown += Form_MouseDown;
            _statusLabel.Parent.MouseMove += Form_MouseMove;
            _statusLabel.Parent.MouseUp += Form_MouseUp;
        }
        
    }
    
    // マウスカーソル位置に基づいてリサイズ方向を判断
    private ResizeDirection GetResizeDirection(Point mousePosition)
    {
        // フルスクリーンモード時はリサイズ無効
        if (_isFullScreen)
            return ResizeDirection.None;
            
        // ピン留めされていない場合、フォームが表示されていない場合はリサイズ無効
        if (_pinMode == PinMode.None && !this.Visible)
            return ResizeDirection.None;
        
        int clientWidth = this.ClientSize.Width;
        int clientHeight = this.ClientSize.Height;
        
        bool onLeftEdge = mousePosition.X <= RESIZE_BORDER_SIZE;
        bool onRightEdge = mousePosition.X >= clientWidth - RESIZE_BORDER_SIZE;
        bool onBottomEdge = mousePosition.Y >= clientHeight - RESIZE_BORDER_SIZE;
        bool onTopEdge = mousePosition.Y <= RESIZE_BORDER_SIZE && _allowTopResize;
        
        // ヘッダーパネル内の場合はリサイズ無効（ドラッグ移動用）
        if (mousePosition.Y <= _headerPanel.Height)
            return ResizeDirection.None;
        
        if (onTopEdge && onLeftEdge)
            return ResizeDirection.TopLeft;
        else if (onTopEdge && onRightEdge)
            return ResizeDirection.TopRight;
        else if (onBottomEdge && onLeftEdge)
            return ResizeDirection.BottomLeft;
        else if (onBottomEdge && onRightEdge)
            return ResizeDirection.BottomRight;
        else if (onLeftEdge)
            return ResizeDirection.Left;
        else if (onRightEdge)
            return ResizeDirection.Right;
        else if (onBottomEdge)
            return ResizeDirection.Bottom;
        else if (onTopEdge)
            return ResizeDirection.Top;
        else
            return ResizeDirection.None;
    }
    
    // リサイズ時のマウスムーブイベントハンドラ（カーソルの形状変更専用）
    private void Form_ResizeMouseMove(object sender, MouseEventArgs e)
    {
        // ドラッグ中やリサイズ中は何もしない
        if (_isDragging || _isResizing)
            return;
            
        Control control = sender as Control;
        if (control == null) return;
        
        // マウス位置をフォームのクライアント座標に変換
        Point clientPoint;
        
        // 送信元がフォーム自体ならそのまま使用、そうでなければ変換
        if (control == this)
        {
            clientPoint = e.Location;
        }
        else
        {
            clientPoint = this.PointToClient(control.PointToScreen(e.Location));
        }
        
        // マウスカーソルの形状を変更
        UpdateCursorShape(clientPoint);
    }
    
    // マウスカーソルの形状を適切に変更
    private void UpdateCursorShape(Point mousePosition)
    {
        // 現在のマウス位置に基づいてリサイズ方向を取得
        ResizeDirection direction = GetResizeDirection(mousePosition);
        
        // リサイズ方向に応じてカーソル形状を変更
        switch (direction)
        {
            case ResizeDirection.Left:
            case ResizeDirection.Right:
                this.Cursor = Cursors.SizeWE;
                break;
            case ResizeDirection.Bottom:
            case ResizeDirection.Top:
                this.Cursor = Cursors.SizeNS;
                break;
            case ResizeDirection.BottomLeft:
            case ResizeDirection.TopRight:
                this.Cursor = Cursors.SizeNESW;
                break;
            case ResizeDirection.BottomRight:
            case ResizeDirection.TopLeft:
                this.Cursor = Cursors.SizeNWSE;
                break;
            default:
                this.Cursor = Cursors.Default;
                break;
        }
    }
    
    // 実際のリサイズ処理を行う - パフォーマンス改善版
    private void PerformResize(Point currentMousePosition)
    {
        int deltaX = currentMousePosition.X - _resizeStartPoint.X;
        int deltaY = currentMousePosition.Y - _resizeStartPoint.Y;
        
        // 最小サイズを定義
        int minWidth = 200;
        int minHeight = 100;
        
        // 現在の位置とサイズ
        Rectangle bounds = this.Bounds;
        int newWidth, newHeight, newX, newY;
        
        // リサイズ操作前にレイアウト更新を一時停止
        if (!_suspendLayout)
        {
            SuspendLayout();
            _suspendLayout = true;
        }
        
        // リサイズ方向に応じて処理
        switch (_currentResizeDirection)
        {
            case ResizeDirection.Left:
                // 左辺リサイズ
                newWidth = Math.Max(_originalResizeSize.Width - deltaX, minWidth);
                newX = bounds.Right - newWidth;
                this.SetBounds(newX, bounds.Y, newWidth, bounds.Height);
                break;
                
            case ResizeDirection.Right:
                // 右辺リサイズ
                this.Width = Math.Max(_originalResizeSize.Width + deltaX, minWidth);
                break;
                
            case ResizeDirection.Bottom:
                // 下辺リサイズ
                this.Height = Math.Max(_originalResizeSize.Height + deltaY, minHeight);
                break;
                
            case ResizeDirection.BottomLeft:
                // 左下角リサイズ
                newWidth = Math.Max(_originalResizeSize.Width - deltaX, minWidth);
                newX = bounds.Right - newWidth;
                this.SetBounds(newX, bounds.Y, newWidth, Math.Max(_originalResizeSize.Height + deltaY, minHeight));
                break;
                
            case ResizeDirection.BottomRight:
                // 右下角リサイズ
                this.Size = new Size(
                    Math.Max(_originalResizeSize.Width + deltaX, minWidth),
                    Math.Max(_originalResizeSize.Height + deltaY, minHeight)
                );
                break;
                
            case ResizeDirection.Top:
                // 上辺リサイズ（許可されている場合のみ）
                if (_allowTopResize)
                {
                    newWidth = bounds.Width;
                    newHeight = Math.Max(_originalResizeSize.Height - deltaY, minHeight);
                    newY = bounds.Bottom - newHeight;
                    this.SetBounds(bounds.X, newY, newWidth, newHeight);
                }
                break;
                
            case ResizeDirection.TopLeft:
                // 左上角リサイズ（上辺リサイズが許可されている場合のみ）
                if (_allowTopResize)
                {
                    newWidth = Math.Max(_originalResizeSize.Width - deltaX, minWidth);
                    newHeight = Math.Max(_originalResizeSize.Height - deltaY, minHeight);
                    newX = bounds.Right - newWidth;
                    newY = bounds.Bottom - newHeight;
                    this.SetBounds(newX, newY, newWidth, newHeight);
                }
                else
                {
                    // 上辺リサイズが許可されていない場合は左辺のみリサイズ
                    newWidth = Math.Max(_originalResizeSize.Width - deltaX, minWidth);
                    newX = bounds.Right - newWidth;
                    this.SetBounds(newX, bounds.Y, newWidth, bounds.Height);
                }
                break;
                
            case ResizeDirection.TopRight:
                // 右上角リサイズ（上辺リサイズが許可されている場合のみ）
                if (_allowTopResize)
                {
                    newWidth = Math.Max(_originalResizeSize.Width + deltaX, minWidth);
                    newHeight = Math.Max(_originalResizeSize.Height - deltaY, minHeight);
                    newY = bounds.Bottom - newHeight;
                    this.SetBounds(bounds.X, newY, newWidth, newHeight);
                }
                else
                {
                    // 上辺リサイズが許可されていない場合は右辺のみリサイズ
                    this.Width = Math.Max(_originalResizeSize.Width + deltaX, minWidth);
                }
                break;
        }
        
        // トリガーエリアの幅をフォームの幅に合わせる（上部ドック時）
        if (_dockPosition == DockPosition.Top)
        {
            _triggerWidth = this.Width;
            UpdateTriggerArea();
        }
        
        // ステータスバーに現在のサイズを表示
        _statusLabel.Text = "サイズ変更: " + this.Width + " x " + this.Height;
    }

    private void SetupMouseTracking()
    {
        // Set up the timer to check mouse position - 負荷軽減のため更新間隔を長く
        _mouseTrackTimer.Interval = 200; // 100ms → 200ms
        _mouseTrackTimer.Tick += new EventHandler(MouseTrackTimer_Tick);
        _mouseTrackTimer.Start();
        
        // Initialize the trigger area
        UpdateTriggerArea();
    }
    
    private void SetupTrayIcon()
    {
        _notifyIcon.Icon = SystemIcons.Application;
        _notifyIcon.Text = "Dockable Form";
        _notifyIcon.Visible = true;
        
        // Context menu for tray icon
        ContextMenuStrip menu = new ContextMenuStrip();
        ToolStripMenuItem showItem = new ToolStripMenuItem("Show");
        showItem.Click += new EventHandler(ShowItem_Click);
        menu.Items.Add(showItem);
        
        ToolStripMenuItem exitItem = new ToolStripMenuItem("Exit");
        exitItem.Click += new EventHandler(ExitItem_Click);
        menu.Items.Add(exitItem);
        
        _notifyIcon.ContextMenuStrip = menu;
        _notifyIcon.DoubleClick += new EventHandler(NotifyIcon_DoubleClick);
    }
    
    private void ShowItem_Click(object sender, EventArgs e)
    {
        this.Show();
        this.BringToFront();
        
        // 初回表示時以外でアニメーション実行中でなければ
        if (!_isInitialLoad && !_animationEngine.IsRunning)
        {
            // 通常のドックモードなら高さアニメーション実行
            if (!_isFullScreen)
            {
                // 閉じるアニメーション後に非表示にする予定だったらフラグを解除
                _shouldHideWhenAnimationComplete = false;
                this.Height = 0;
                PinMode = PinMode.Pinned;
                StartHeightAnimation(true);
            }
        }
    }
    
    private void UpdateTriggerArea()
    {
        // フルスクリーンモード時はトリガーエリアを無効化
        if (_isFullScreen)
            return;
            
        switch (_dockPosition)
        {
            case DockPosition.Top:
                // Use custom trigger position if set, otherwise center
                if (_triggerX == 0 && _triggerWidth > 0)
                {
                    _triggerX = (Screen.PrimaryScreen.WorkingArea.Width - _triggerWidth) / 2;
                }
                _formLeftPosition = _triggerX;
                _triggerArea = new Rectangle(
                    _triggerX,
                    _triggerY,
                    _triggerWidth,
                    _triggerHeight
                );
                break;
            case DockPosition.Left:
                if (_triggerY == 0 && _triggerWidth > 0)
                {
                    _triggerY = (Screen.PrimaryScreen.WorkingArea.Height - _triggerWidth) / 2;
                }
                _triggerArea = new Rectangle(
                    _triggerX,
                    _triggerY,
                    _triggerHeight,
                    _triggerWidth
                );
                break;
            case DockPosition.Right:
                if (_triggerY == 0 && _triggerWidth > 0)
                {
                    _triggerY = (Screen.PrimaryScreen.WorkingArea.Height - _triggerWidth) / 2;
                }
                if (_triggerX == 0)
                {
                    _triggerX = Screen.PrimaryScreen.WorkingArea.Width - _triggerHeight;
                }
                _triggerArea = new Rectangle(
                    _triggerX,
                    _triggerY,
                    _triggerHeight,
                    _triggerWidth
                );
                break;
            case DockPosition.None:
                // Custom trigger area regardless of docking
                _triggerArea = new Rectangle(
                    _triggerX,
                    _triggerY,
                    _triggerWidth,
                    _triggerHeight
                );
                break;
        }
    }
    
    private void UpdateDockPosition()
    {
        if (_isFullScreen) return; // フルスクリーンモード時は位置更新しない
        
        UpdateTriggerArea();
        
        switch (_dockPosition)
        {
            case DockPosition.Top:
                this.Location = new Point(_formLeftPosition, 0);
                break;
            case DockPosition.Left:
                this.Location = new Point(0, (Screen.PrimaryScreen.WorkingArea.Height - this.Height) / 2);
                break;
            case DockPosition.Right:
                this.Location = new Point(Screen.PrimaryScreen.WorkingArea.Width - this.Width, 
                                         (Screen.PrimaryScreen.WorkingArea.Height - this.Height) / 2);
                break;
            case DockPosition.None:
                // Not docked, can be placed anywhere
                break;
        }
        
        // 現在の位置をドック位置として保存
        _dockLocation = this.Location;
    }

    private void UpdateWindowButtonsVisibility()
    {
        if (_showStandardWindowButtons)
        {
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.ControlBox = true;
            this.MinimizeBox = true;
            this.MaximizeBox = true;
        }
        else
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.ControlBox = false;
            this.MinimizeBox = false;
            this.MaximizeBox = false;
        }
    }
    
    private void UpdateFullScreenButtonAppearance()
    {
        _fullscreenButton.ForeColor = _isFullScreen ? Color.Yellow : Color.White;
        _fullscreenButton.Text = _isFullScreen ? "🔽" : "🔍"; // 拡大と縮小アイコンを切り替え
    }
    
    private bool IsMouseInForm()
    {
        POINT cursorPos;
        GetCursorPos(out cursorPos);
        Point clientCursorPos = this.PointToClient(new Point(cursorPos.X, cursorPos.Y));
        
        return this.ClientRectangle.Contains(clientCursorPos);
    }
    
    private bool IsMouseInTriggerArea()
    {
        // フルスクリーンモード時はトリガーエリアチェックをスキップ
        if (_isFullScreen)
            return false;
            
        POINT cursorPos;
        GetCursorPos(out cursorPos);
        
        return _triggerArea.Contains(cursorPos.X, cursorPos.Y);
    }
    
    // Event handlers - 最適化版
    private void MouseTrackTimer_Tick(object sender, EventArgs e)
    {
        // リサイズ中は処理をスキップして負荷を軽減
        if (_isResizing)
            return;

        bool isAnimating = _animationEngine.IsRunning;

        // 開くアニメーション中は処理をスキップし、閉じるアニメーション中は再フォーカス検知のため継続
        if (isAnimating && !_isClosingAnimation)
            return;
            
        // カーソルがトリガーエリアに入ると最前面に表示
        if (IsMouseInTriggerArea())
        {
            this.TopMost = true;
            this.TopMost = false;
        }
            
        if (_pinMode != PinMode.None)
        {
            if (_showDelayTimer.Enabled || _showDelayPending)
            {
                _showDelayTimer.Stop();
                _showDelayPending = false;
            }

            // If pinned, always show
            if (!this.Visible)
            {
                this.Show();
            }
            
            return;
        }
        
        // フルスクリーンモード時はトリガーによる表示/非表示を無効化
        if (_isFullScreen)
        {
            if (_showDelayTimer.Enabled || _showDelayPending)
            {
                _showDelayTimer.Stop();
                _showDelayPending = false;
            }
            return;
        }
            
        // Check if mouse is in trigger area
        if (IsMouseInTriggerArea() && !this.Visible)
        {
            if (!_showDelayTimer.Enabled && !_showDelayPending)
            {
                _showDelayPending = true;
                _showDelayTimer.Interval = _showDelayMilliseconds;
                _showDelayTimer.Start();
                _statusLabel.Text = "トリガーエリアに入りました (1秒後に開きます)";
            }
        }
        else if (!IsMouseInTriggerArea() && _showDelayPending)
        {
            _showDelayTimer.Stop();
            _showDelayPending = false;
            _statusLabel.Text = "トリガーエリアから離れました (開く処理をキャンセル)";
        }
        // マウスがフォーム内にある場合
        else if (IsMouseInForm())
        {
            // 閉じるアニメーション中であれば、開くアニメーションに切り替え
            if (_isClosingAnimation && _animationEngine.IsRunning)
            {
                _shouldHideWhenAnimationComplete = false; // 非表示フラグをクリア
                StopAnimationAndResumeLayoutIfNeeded();
                _isClosingAnimation = false;

                int currentHeight = this.Height;
                StartHeightAnimation(true, currentHeight);
                _statusLabel.Text = "アニメーション方向を反転して開きました";
            }

            _hideDelayTimer.Stop();
            _mouseLeavePending = false;
            _showDelayTimer.Stop();
            _showDelayPending = false;
        }
        // マウスがフォーム外にある場合
        else if (!IsMouseInForm() && this.Visible && !_animationEngine.IsRunning && !_isClosingAnimation)
        {
            // まだ遅延タイマーが動いていなければ
            if (!_hideDelayTimer.Enabled && !_mouseLeavePending)
            {
                // 遅延タイマーを開始
                _mouseLeavePending = true;
                _hideDelayTimer.Start();
                _statusLabel.Text = "フォームの外に出ました (閉じるまで1秒)";
            }
        }
        else if (IsMouseInForm() && _mouseLeavePending)
        {
            // マウスがフォーム内に戻ってきたら遅延タイマーをキャンセル
            _hideDelayTimer.Stop();
            _mouseLeavePending = false;
            _showDelayTimer.Stop();
            _showDelayPending = false;
            _statusLabel.Text = "フォームに戻りました";
        }
    }

    private void UpdatePinModeSettings()
    {
        // ピンモードに応じて設定を変更
        switch (_pinMode)
        {
            case PinMode.None:
                this.TopMost = false;
                _pinButton.ForeColor = Color.White;
                _statusLabel.Text = "ピン留め解除: 自動表示/非表示モード";
                break;
            case PinMode.Pinned:
                this.TopMost = false;
                _pinButton.ForeColor = Color.Orange;
                _statusLabel.Text = "ピン留め: 常に表示モード";
                break;
            case PinMode.PinnedTopMost:
                this.TopMost = true;
                _pinButton.ForeColor = Color.LightGreen;
                _statusLabel.Text = "ピン留め＆最前面: 常に最前面表示モード";
                break;
        }
    }
    
    private void PinButton_Click(object sender, EventArgs e)
    {
        // ピンモードを切り替え（3段階）
        switch (_pinMode)
        {
            case PinMode.None:
                PinMode = PinMode.Pinned;
                break;
            case PinMode.Pinned:
                PinMode = PinMode.PinnedTopMost;
                break;
            case PinMode.PinnedTopMost:
                PinMode = PinMode.None;
                break;
        }
    }
    
    private void FullscreenButton_Click(object sender, EventArgs e)
    {
        IsFullScreen = !IsFullScreen;
    }
    
    private void ExitItem_Click(object sender, EventArgs e)
    {
        // Clean up and exit
        _mouseTrackTimer.Stop();
        StopAnimationAndResumeLayoutIfNeeded();
        _notifyIcon.Visible = false;
        Application.Exit();
    }
    
    private void NotifyIcon_DoubleClick(object sender, EventArgs e)
    {
        this.Show();
        this.BringToFront();
    }
    
    // Form_MouseDown イベントハンドラ（ドラッグとリサイズの両方に対応）
    private void Form_MouseDown(object sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            // マウス位置をフォームのクライアント座標に変換
            Control control = sender as Control;
            Point clientPoint;
            
            // 送信元がフォーム自体ならそのまま使用、そうでなければ変換
            if (control == this)
            {
                clientPoint = e.Location;
            }
            else
            {
                clientPoint = this.PointToClient(control.PointToScreen(e.Location));
            }
            
            // リサイズ方向を取得
            _currentResizeDirection = GetResizeDirection(clientPoint);
            
            if (_currentResizeDirection != ResizeDirection.None)
            {
                // リサイズ開始
                _isResizing = true;
                _resizeStartPoint = clientPoint;
                _originalResizeSize = this.Size;
                
                // アニメーション実行中ならば停止
                if (_animationEngine.IsRunning)
                {
                    StopAnimationAndResumeLayoutIfNeeded();
                    _isClosingAnimation = false;
                    _shouldHideWhenAnimationComplete = false;
                }
            }
            else
            {
                // リサイズ領域外ならドラッグ移動開始
                _isDragging = true;
                _dragStartPoint = new Point(e.X, e.Y);
            }
        }
    }

    // Form_MouseMove イベントハンドラ（ドラッグとリサイズの両方に対応）
    private void Form_MouseMove(object sender, MouseEventArgs e)
    {
        if (_isResizing)
        {
            // マウス位置をフォームのクライアント座標に変換
            Control control = sender as Control;
            Point clientPoint;
            
            // 送信元がフォーム自体ならそのまま使用、そうでなければ変換
            if (control == this)
            {
                clientPoint = e.Location;
            }
            else
            {
                clientPoint = this.PointToClient(control.PointToScreen(e.Location));
            }
            
            // リサイズ中
            PerformResize(clientPoint);
        }
        else if (_isDragging)
        {
            // フルスクリーンモード時はドラッグ無効
            if (_isFullScreen)
                return;
                
            Point newLocation = this.Location;
            
            if (_dockPosition == DockPosition.Top)
            {
                // When docked to top, only allow horizontal movement
                newLocation.X = this.Location.X + (e.X - _dragStartPoint.X);
                
                // Keep within screen bounds
                if (newLocation.X < 0)
                    newLocation.X = 0;
                if (newLocation.X + this.Width > Screen.PrimaryScreen.WorkingArea.Width)
                    newLocation.X = Screen.PrimaryScreen.WorkingArea.Width - this.Width;
                
                this.Location = newLocation;
                
                // Update trigger area position
                _formLeftPosition = newLocation.X;
                _triggerX = newLocation.X;
                UpdateTriggerArea();
            }
            else if (_dockPosition == DockPosition.Left || _dockPosition == DockPosition.Right)
            {
                // When docked to left or right, only allow vertical movement
                newLocation.Y = this.Location.Y + (e.Y - _dragStartPoint.Y);
                
                // Keep within screen bounds
                if (newLocation.Y < 0)
                    newLocation.Y = 0;
                if (newLocation.Y + this.Height > Screen.PrimaryScreen.WorkingArea.Height)
                    newLocation.Y = Screen.PrimaryScreen.WorkingArea.Height - this.Height;
                
                this.Location = newLocation;
                
                // Update trigger area position
                _triggerY = newLocation.Y;
                UpdateTriggerArea();
            }
            else
            {
                // Not docked, can move freely
                newLocation.X = this.Location.X + (e.X - _dragStartPoint.X);
                newLocation.Y = this.Location.Y + (e.Y - _dragStartPoint.Y);
                this.Location = newLocation;
            }
            
            // 現在の位置をドック位置として保存
            _dockLocation = this.Location;
        }
    }
    
    // Form_MouseUp イベントハンドラ（ドラッグとリサイズの両方に対応）
    private void Form_MouseUp(object sender, MouseEventArgs e)
    {
        if (_isResizing)
        {
            _isResizing = false;
            _currentResizeDirection = ResizeDirection.None;
            
            // リサイズ後のサイズを保存
            if (!_isFullScreen)
            {
                _originalHeight = this.Height;
                _originalWidth = this.Width;
                _dockSize = new Size(this.Width, this.Height);
            }
            
            // レイアウト更新を再開
            if (_suspendLayout)
            {
                ResumeLayout(true);
                _suspendLayout = false;
            }
            
            this.Cursor = Cursors.Default;
        }
        
        _isDragging = false;
    }
    
    // ディスポーズ処理のオーバーライド
    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            if (_mouseTrackTimer != null)
            {
                _mouseTrackTimer.Dispose();
            }
            StopAnimationAndResumeLayoutIfNeeded();
            if (_notifyIcon != null)
            {
                _notifyIcon.Dispose();
            }
            // 追加: 遅延タイマーの破棄
            if (_hideDelayTimer != null)
            {
                _hideDelayTimer.Stop();
                _hideDelayTimer.Dispose();
            }
            if (_showDelayTimer != null)
            {
                _showDelayTimer.Stop();
                _showDelayTimer.Dispose();
            }
        }
        base.Dispose(disposing);
    }
    
    [STAThread]
    public static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        
        AnimatedDockableForm form = new AnimatedDockableForm();
        Application.Run(form);
    }
}
"@

# Add the required assemblies
Add-Type -TypeDefinition $csCode -ReferencedAssemblies System.Windows.Forms, System.Drawing

# Launch the application using the static method
[AnimatedDockableForm]::Main()
