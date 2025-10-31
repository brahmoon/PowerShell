# This script combines the DockableForm and MetroUI components
# Coding guidelines: Use C# 5.0. (Features and syntax from C# 6.0 or later are not allowed.)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# C# code definition - Combined from both sources
$csCode = @"
using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Collections.Generic;
using System.Reflection;
using System.ComponentModel;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Threading;

#region DockableForm Classes

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
    private System.Windows.Forms.Timer _timer;
    private int _currentStep = 0;
    private int _totalSteps;
    private EasingFunctions.EasingType _easingType;
    private Action<float> _updateAction;
    private Action _completedAction;
    
    public bool IsRunning { get; private set; }
    
    // アニメーション速度の最適化（FPSを30に削減）
    public AnimationEngine(int durationMs, int fps = 30, EasingFunctions.EasingType easingType = EasingFunctions.EasingType.None)
    {
        _totalSteps = (int)(durationMs / (1000.0 / fps));
        _easingType = easingType;
        _timer = new System.Windows.Forms.Timer();
        _timer.Interval = (int)(1000.0 / fps);
        _timer.Tick += Timer_Tick;
    }
    
    private void Timer_Tick(object sender, EventArgs e)
    {
        // アニメーションステップを進める
        _currentStep++;
        
        // 進行度（0〜1の範囲）を計算
        float progress = (float)_currentStep / _totalSteps;
        
        // イージング効果を適用
        float easedProgress = EasingFunctions.ApplyEasing(progress, _easingType);
        
        // アニメーションを更新
        if (_updateAction != null)
        {
            _updateAction(easedProgress);
        }
        
        // アニメーション完了判定
        if (_currentStep >= _totalSteps)
        {
            _timer.Stop();
            IsRunning = false;
            
            // 完了アクションがあれば実行
            if (_completedAction != null)
            {
                _completedAction();
            }
        }
    }
    
    public void Start(Action<float> updateAction, Action completedAction = null)
    {
        if (IsRunning) return;
        
        _currentStep = 0;
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
    private System.Windows.Forms.Timer _mouseTrackTimer = new System.Windows.Forms.Timer();
    private System.Windows.Forms.Timer _hideDelayTimer = new System.Windows.Forms.Timer();
    private bool _mouseLeavePending = false;
    private System.Windows.Forms.Timer _triggerDelayTimer;
    private bool _triggerHoverPending = false;
    private int _hideDelayMilliseconds = 1000; // 1秒の遅延
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
        set { _hideDelayMilliseconds = value; }
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
        InitializeTriggerDelay();
        InitializeHideDelay();
        
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
    
    private void InitializeTriggerDelay()
    {
        // トリガー表示遅延タイマーの設定
        _triggerDelayTimer = new System.Windows.Forms.Timer();
        _triggerDelayTimer.Interval = 1000; // 1秒間の遅延
        _triggerDelayTimer.Tick += TriggerDelayTimer_Tick;
        _triggerDelayTimer.Enabled = false;
    }
    
    // トリガー遅延タイマーのTickイベントハンドラ
    private void TriggerDelayTimer_Tick(object sender, EventArgs e)
    {
        // タイマーを停止
        _triggerDelayTimer.Stop();
        
        // 遅延後もまだマウスがトリガーエリア内にあるか確認
        if (_triggerHoverPending && IsMouseInTriggerArea() && !this.Visible && !_animationEngine.IsRunning)
        {
            // フォームが非表示から表示されるときのアニメーション
            this.Height = 0;
            this.Show();
            this.BringToFront();
            StartHeightAnimation(true);
            _statusLabel.Text = "トリガーエリアに入りました (1秒経過)";
            
            // 保留フラグをクリア
            _triggerHoverPending = false;
        }
        else
        {
            // マウスが移動したか、他の条件により表示しないことに
            _triggerHoverPending = false;
        }
    }
    
    private void InitializeHideDelay()
    {
        // 遅延タイマーの設定
        _hideDelayTimer.Interval = _hideDelayMilliseconds;
        _hideDelayTimer.Tick += HideDelayTimer_Tick;
        _hideDelayTimer.Enabled = false;
    }
    
    // 遅延タイマーのTickイベントハンドラ
    private void HideDelayTimer_Tick(object sender, EventArgs e)
    {
        // タイマーを停止
        _hideDelayTimer.Stop();
        
        // 遅延後もまだマウスがフォーム外にあるか確認
        if (_mouseLeavePending && !IsMouseInForm() && this.Visible && !_animationEngine.IsRunning)
        {
            // アニメーションで閉じる
            _shouldHideWhenAnimationComplete = true;
            StartHeightAnimation(false);
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
    
    private void StartHeightAnimation(bool isOpening)
    {
        // アニメーションが実行中なら一旦停止
        if (_animationEngine.IsRunning)
        {
            _animationEngine.Stop();
        }
        
        // 閉じるアニメーション実行中フラグの設定
        _isClosingAnimation = !isOpening;
        
        if (isOpening)
        {
            // リサイズ中のレイアウト更新を一時停止
            SuspendLayout();
            _suspendLayout = true;
            
            // 表示（開く）アニメーション
            _animationEngine.Start(
                // 更新処理
                (progress) => {
                    int targetHeight = (int)(progress * _originalHeight);
                    this.Height = targetHeight;
                    
                    // 位置更新（上部に固定）
                    if (_dockPosition == DockPosition.Top && !_isFullScreen)
                    {
                        this.Location = new Point(this.Location.X, 0);
                    }
                },
                // アニメーション完了時の処理
                () => {
                    this.Height = _originalHeight;
                    _isClosingAnimation = false;
                    
                    // レイアウト更新を再開
                    if (_suspendLayout)
                    {
                        ResumeLayout(true);
                        _suspendLayout = false;
                    }
                }
            );
        }
        else
        {
            // リサイズ中のレイアウト更新を一時停止
            SuspendLayout();
            _suspendLayout = true;
            
            // 閉じるアニメーション
            _animationEngine.Start(
                // 更新処理
                (progress) => {
                    int targetHeight = (int)(_originalHeight * (1 - progress));
                    this.Height = targetHeight;
                },
                // クローズアニメーション完了時の処理
                () => {
                    _isClosingAnimation = false;
                    if (_shouldHideWhenAnimationComplete)
                    {
                        this.Hide();
                        _shouldHideWhenAnimationComplete = false;
                    }
                    
                    // レイアウト更新を再開
                    if (_suspendLayout)
                    {
                        ResumeLayout(true);
                        _suspendLayout = false;
                    }
                }
            );
        }
    }
    
    private void ToggleFullScreenMode()
    {
        // アニメーションが実行中なら一旦停止
        if (_animationEngine.IsRunning)
        {
            _animationEngine.Stop();
        }
        
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
                StartHeightAnimation(false);
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
        // リサイズ中やアニメーション中は処理をスキップして負荷を軽減
        if (_isResizing || _animationEngine.IsRunning)
            return;
            
        // カーソルがトリガーエリアに入ると最前面に表示
        if (IsMouseInTriggerArea())
        {
            this.TopMost = true;
            this.TopMost = false;
        }
            
        if (_pinMode != PinMode.None)
        {
            // If pinned, always show
            if (!this.Visible)
            {
                this.Show();
            }
            
            return;
        }
        
        // フルスクリーンモード時はトリガーによる表示/非表示を無効化
        if (_isFullScreen)
            return;
            
        // Check if mouse is in trigger area
        if (IsMouseInTriggerArea() && !this.Visible)
        {
            // トリガーによる表示の遅延処理
            if (!_triggerHoverPending && !_triggerDelayTimer.Enabled)
            {
                // 遅延タイマーを開始
                _triggerHoverPending = true;
                _triggerDelayTimer.Start();
                _statusLabel.Text = "トリガーエリアに入りました (表示まで1秒)";
            }
        }
        else if (!IsMouseInTriggerArea() && _triggerHoverPending)
        {
            // マウスがトリガーエリアから出た場合は遅延タイマーをキャンセル
            _triggerDelayTimer.Stop();
            _triggerHoverPending = false;
            _statusLabel.Text = "トリガーエリアから出ました";
        }
        // マウスがフォーム内にある場合
        else if (IsMouseInForm())
        {
            // 閉じるアニメーション中であれば、開くアニメーションに切り替え
            if (_isClosingAnimation && _animationEngine.IsRunning)
            {
                _shouldHideWhenAnimationComplete = false; // 非表示フラグをクリア
                
                // 現在のアニメーションを停止し、逆向きに開始
                float currentProgress = _animationEngine.GetCurrentProgress();
                float remainingProgress = 1.0f - currentProgress;
                
                _animationEngine.Stop();
                _isClosingAnimation = false;
                
                // 現在の高さから開くアニメーションを開始
                _animationEngine.Start(
                    // 更新処理
                    (progress) => {
                        int currentHeight = this.Height;
                        int targetHeight = (int)(_originalHeight * progress + currentHeight * (1.0f - progress));
                        this.Height = targetHeight;
                    },
                    // 完了時の処理
                    () => {
                        this.Height = _originalHeight;
                        _statusLabel.Text = "アニメーション方向を反転して開きました";
                    }
                );
            }
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
        _animationEngine.Stop();
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
                    _animationEngine.Stop();
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
            if (_animationEngine != null)
            {
                _animationEngine.Stop();
            }
            if (_notifyIcon != null)
            {
                _notifyIcon.Dispose();
            }
            if (_triggerDelayTimer != null)
            {
                _triggerDelayTimer.Stop();
                _triggerDelayTimer.Dispose();
            }
            if (_hideDelayTimer != null)
            {
                _hideDelayTimer.Stop();
                _hideDelayTimer.Dispose();
            }
        }
        base.Dispose(disposing);
    }
}

#endregion

#region MetroUI Classes

namespace MetroUI
{
    #region Models

    /// <summary>
    /// 予定を表すクラス
    /// </summary>
    public class Appointment
    {
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public string Title { get; set; }
        public int NotificationMinutesBefore { get; set; }
        public bool IsCompleted { get; set; }
        public bool IsNotified { get; set; }
        public AppointmentStatus Status { get; set; }
        
        private List<MetroTaskManager.SubTask> _subTasks;
        private string _memo;
        private bool _showSubTasks;
        public List<MetroTaskManager.SubTask> SubTasks
        { 
            get { return _subTasks; } 
            set { _subTasks = value; }
        }

        public string Memo 
        { 
            get { return _memo; } 
            set { _memo = value; }
        }

        public bool HasMemo 
        { 
            get { return !string.IsNullOrEmpty(Memo); } 
        }

        public bool ShowSubTasks 
        { 
            get { return _showSubTasks; } 
            set { _showSubTasks = value; }
        }
        
        public Appointment()
        {
            StartTime = DateTime.Now;
            EndTime = DateTime.Now.AddHours(1);
            Title = "新しい予定";
            NotificationMinutesBefore = 15;
            IsCompleted = false;
            IsNotified = false;
            Status = AppointmentStatus.Upcoming;
            
            // 追加
            _subTasks = new List<MetroTaskManager.SubTask>();
            _memo = "";
            _showSubTasks = false;
        }

        public DateTime NotificationTime 
        { 
            get { return StartTime.AddMinutes(-NotificationMinutesBefore); } 
        }

        public bool IsOnDate(DateTime date)
        {
            return StartTime.Date <= date.Date && EndTime.Date >= date.Date;
        }

        public void UpdateStatus()
        {
            DateTime now = DateTime.Now;

            if (IsCompleted)
            {
                Status = AppointmentStatus.Completed;
            }
            else if (now > EndTime)
            {
                Status = AppointmentStatus.Overdue;
            }
            else if (now >= StartTime)
            {
                Status = AppointmentStatus.InProgress;
            }
            else
            {
                Status = AppointmentStatus.Upcoming;
            }
        }
    }

    /// <summary>
    /// 予定の状態を表す列挙型
    /// </summary>
    public enum AppointmentStatus
    {
        Upcoming,
        InProgress,
        Completed,
        Overdue
    }

    #endregion

    #region Color Theme

    /// <summary>
    /// カラーテーマを管理するクラス
    /// </summary>
    public static class MetroColors
    {
        public static Color Primary = Color.FromArgb(0, 120, 215);
        public static Color Secondary = Color.FromArgb(0, 99, 177);
        public static Color Background = Color.White;
        public static Color Text = Color.FromArgb(51, 51, 51);
        public static Color TextLight = Color.FromArgb(153, 153, 153);
        public static Color Success = Color.FromArgb(92, 184, 92);
        public static Color Warning = Color.FromArgb(240, 173, 78);
        public static Color Danger = Color.FromArgb(217, 83, 79);
        public static Color Info = Color.FromArgb(91, 192, 222);
        public static Color InProgress = Color.FromArgb(135, 206, 250); // LightSkyBlue

        public static Color[] AccentColors = new Color[]
        {
            Color.FromArgb(0, 120, 215),    // Blue
            Color.FromArgb(231, 72, 86),    // Red
            Color.FromArgb(16, 124, 16),    // Green
            Color.FromArgb(180, 80, 0),     // Orange
            Color.FromArgb(106, 0, 255),    // Purple
            Color.FromArgb(170, 0, 170),    // Magenta
            Color.FromArgb(118, 118, 118)   // Gray
        };

        /// <summary>
        /// 予定の状態に応じた色を取得
        /// </summary>
        public static Color GetStatusColor(AppointmentStatus status)
        {
            switch (status)
            {
                case AppointmentStatus.InProgress:
                    return InProgress;
                case AppointmentStatus.Completed:
                    return Success;
                case AppointmentStatus.Overdue:
                    return Danger;
                case AppointmentStatus.Upcoming:
                default:
                    return Color.White;
            }
        }
    }

    #endregion

    #region Utilities

    /// <summary>
    /// 描画ユーティリティクラス
    /// </summary>
    internal static class DrawingUtils
    {
        public static GraphicsPath RoundedRect(Rectangle bounds, int radius)
        {
            int diameter = radius * 2;
            Size size = new Size(diameter, diameter);
            Rectangle arc = new Rectangle(bounds.Location, size);
            GraphicsPath path = new GraphicsPath();

            // 左上
            path.AddArc(arc, 180, 90);

            // 右上
            arc.X = bounds.Right - diameter;
            path.AddArc(arc, 270, 90);

            // 右下
            arc.Y = bounds.Bottom - diameter;
            path.AddArc(arc, 0, 90);

            // 左下
            arc.X = bounds.Left;
            path.AddArc(arc, 90, 90);

            path.CloseFigure();
            return path;
        }

        public static void DrawRoundedRectangle(Graphics graphics, Pen pen, Rectangle bounds, int radius)
        {
            using (GraphicsPath path = RoundedRect(bounds, radius))
            {
                graphics.DrawPath(pen, path);
            }
        }

        public static void FillRoundedRectangle(Graphics graphics, Brush brush, Rectangle bounds, int radius)
        {
            using (GraphicsPath path = RoundedRect(bounds, radius))
            {
                graphics.FillPath(brush, path);
            }
        }

        /// <summary>
        /// 円弧を描画
        /// </summary>
        public static void DrawProgressArc(Graphics g, Rectangle bounds, float startAngle, float sweepAngle, float thickness, Color color)
        {
            using (Pen pen = new Pen(color, thickness))
            {
                pen.StartCap = LineCap.Round;
                pen.EndCap = LineCap.Round;
                g.DrawArc(pen, bounds, startAngle, sweepAngle);
            }
        }
    }

    #endregion

    #region Notification System

    /// <summary>
    /// 通知ウィンドウ
    /// </summary>
    public class NotificationWindow : Form
    {
        private Appointment _appointment;
        private Button _closeButton;
        private Label _titleLabel;
        private Label _timeLabel;
        private Label _messageLabel;
        
        // 表示中の通知ウィンドウを管理する静的リスト
        public static List<NotificationWindow> ActiveNotifications = new List<NotificationWindow>();
        private System.Windows.Forms.Timer _closeTimer;
        
        public NotificationWindow(Appointment appointment)
        {
            _appointment = appointment;
            InitializeComponents();
            SetupTimer();
            
            // 静的リストに自分自身を追加
            ActiveNotifications.Add(this);
            
            // フォームが閉じられたときにリストから削除するためのイベントハンドラを追加
            this.FormClosed += (sender, e) => ActiveNotifications.Remove(this);
        }

        private void InitializeComponents()
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.Manual;
            this.ShowInTaskbar = false;
            this.TopMost = true;
            this.Size = new Size(300, 150);

            // 画面右下に配置
            Rectangle workingArea = Screen.PrimaryScreen.WorkingArea;
            this.Location = new Point(
                workingArea.Right - this.Width - 10,
                workingArea.Bottom - this.Height - 10);

            // ウィンドウの角を丸くする
            this.Region = Region.FromHrgn(NativeMethods.CreateRoundRectRgn(0, 0, this.Width, this.Height, 10, 10));
            this.BackColor = Color.FromArgb(240, 240, 240);

            // タイトルラベル
            _titleLabel = new Label();
            _titleLabel.Text = _appointment.Title;
            _titleLabel.Font = new Font("Segoe UI", 11, FontStyle.Bold);
            _titleLabel.ForeColor = MetroColors.Primary;
            _titleLabel.Location = new Point(15, 15);
            _titleLabel.Size = new Size(270, 23);
            _titleLabel.AutoEllipsis = true;
            this.Controls.Add(_titleLabel);

            // 時間ラベル
            _timeLabel = new Label();
            _timeLabel.Text = String.Format("{0:HH:mm} - {1:HH:mm}",
                _appointment.StartTime, _appointment.EndTime);
            _timeLabel.Font = new Font("Segoe UI", 9, FontStyle.Regular);
            _timeLabel.ForeColor = MetroColors.TextLight;
            _timeLabel.Location = new Point(15, 40);
            _timeLabel.Size = new Size(270, 18);
            this.Controls.Add(_timeLabel);

            // メッセージラベル
            _messageLabel = new Label();
            _messageLabel.Font = new Font("Segoe UI", 10, FontStyle.Regular);
            _messageLabel.ForeColor = MetroColors.Text;
            _messageLabel.Location = new Point(15, 70);
            _messageLabel.Size = new Size(270, 20);
            this.Controls.Add(_messageLabel);

            // 初期表示用に時間テキストを更新
            UpdateTimeRemainingText();

            // 閉じるボタン
            _closeButton = new Button();
            _closeButton.Text = "閉じる";
            _closeButton.Font = new Font("Segoe UI", 9, FontStyle.Regular);
            _closeButton.Size = new Size(80, 30);
            _closeButton.Location = new Point(this.Width - _closeButton.Width - 15, this.Height - _closeButton.Height - 15);
            _closeButton.FlatStyle = FlatStyle.Flat;
            _closeButton.FlatAppearance.BorderSize = 0;
            _closeButton.BackColor = MetroColors.Primary;
            _closeButton.ForeColor = Color.White;
            _closeButton.Click += (sender, e) => this.Close();
            this.Controls.Add(_closeButton);

            // マウスイベント
            this.MouseDown += NotificationWindow_MouseDown;
            this.MouseMove += NotificationWindow_MouseMove;
            this.Paint += NotificationWindow_Paint;
        }

        private void SetupTimer()
        {
            // 自動的に閉じる代わりに、1分ごとに残り時間表示を更新するタイマー
            _closeTimer = new System.Windows.Forms.Timer();
            _closeTimer.Interval = 60000; // 1分ごとに更新
            _closeTimer.Tick += (sender, e) => 
            {
                // 残り時間表示を更新
                UpdateTimeRemainingText();
            };
            _closeTimer.Start();
        }
       
        private Point _lastMousePos;

        private void NotificationWindow_MouseDown(object sender, MouseEventArgs e)
        {
            _lastMousePos = e.Location;
        }

        private void NotificationWindow_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.Location = new Point(
                    this.Location.X + (e.X - _lastMousePos.X),
                    this.Location.Y + (e.Y - _lastMousePos.Y));
            }
        }

        private void NotificationWindow_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            
            // 左端に色付きの縦線を描画
            using (SolidBrush lineBrush = new SolidBrush(MetroColors.Primary))
            {
                g.FillRectangle(lineBrush, new Rectangle(0, 0, 5, this.Height));
            }
        }
        
        public void UpdateTimeRemainingText()
        {
            DateTime now = DateTime.Now;
            
            if (now < _appointment.StartTime)
            {
                TimeSpan timeLeft = _appointment.StartTime - now;
                if (timeLeft.TotalMinutes < 60)
                {
                    _messageLabel.Text = String.Format("開始まであと{0}分です", (int)timeLeft.TotalMinutes);
                }
                else
                {
                    _messageLabel.Text = String.Format("開始まであと{0}時間{1}分です",
                        (int)timeLeft.TotalHours, timeLeft.Minutes);
                }
            }
            else
            {
                TimeSpan elapsed = now - _appointment.StartTime;
                if (_appointment.EndTime > now)
                {
                    // まだ終了時間前
                    TimeSpan remaining = _appointment.EndTime - now;
                    if (remaining.TotalMinutes < 60)
                    {
                        _messageLabel.Text = String.Format("進行中: 終了まであと{0}分", (int)remaining.TotalMinutes);
                    }
                    else
                    {
                        _messageLabel.Text = String.Format("進行中: 終了まであと{0}時間{1}分",
                            (int)remaining.TotalHours, remaining.Minutes);
                    }
                }
                else
                {
                    // 終了時間後
                    _messageLabel.Text = String.Format("終了時間から{0}分経過しています", (int)elapsed.TotalMinutes);
                }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_closeTimer != null)
                {
                    _closeTimer.Stop();
                    _closeTimer.Dispose();
                    _closeTimer = null;
                }
            }
            base.Dispose(disposing);
        }

        #region Native Methods

        private static class NativeMethods
        {
            [System.Runtime.InteropServices.DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
            public static extern IntPtr CreateRoundRectRgn(int nLeftRect, int nTopRect,
                int nRightRect, int nBottomRect, int nWidthEllipse, int nHeightEllipse);
        }

        #endregion
    }

    /// <summary>
    /// 通知マネージャー
    /// </summary>
    public class NotificationManager
    {
        private System.Windows.Forms.Timer _checkTimer;
        private List<Appointment> _appointments;

        public NotificationManager(List<Appointment> appointments)
        {
            _appointments = appointments;
            _checkTimer = new System.Windows.Forms.Timer();
            _checkTimer.Interval = 10000; // 10秒ごとにチェック
            _checkTimer.Tick += CheckTimer_Tick;
            _checkTimer.Start();
        }

        private void CheckTimer_Tick(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            foreach (Appointment appointment in _appointments)
            {
                // 通知時間に達したが、まだ通知されていない場合
                if (!appointment.IsNotified && 
                    now >= appointment.NotificationTime && 
                    now < appointment.StartTime.AddMinutes(5))
                {
                    ShowNotification(appointment);
                    appointment.IsNotified = true;
                }

                // 状態更新
                appointment.UpdateStatus();
            }
        }

        private void ShowNotification(Appointment appointment)
        {
            NotificationWindow notification = new NotificationWindow(appointment);
            notification.Show();
        }

        public void Dispose()
        {
            if (_checkTimer != null)
            {
                _checkTimer.Stop();
                _checkTimer.Dispose();
                _checkTimer = null;
            }
        }
    }

    #endregion
    
    #region Base Controls

    /// <summary>
    /// MetroUIコントロールの基底クラス
    /// </summary>
    public abstract class MetroControlBase : Control
    {
        private Font _metroFont;

        public MetroControlBase()
        {
            SetStyle(
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.ResizeRedraw |
                ControlStyles.UserPaint,
                true);

            BackColor = MetroColors.Background;
            ForeColor = MetroColors.Text;
            _metroFont = new Font("Segoe UI", 9F);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_metroFont != null)
                {
                    _metroFont.Dispose();
                    _metroFont = null;
                }
            }
            base.Dispose(disposing);
        }

        public Font MetroFont
        {
            get { return _metroFont; }
            set
            {
                if (_metroFont != null)
                {
                    _metroFont.Dispose();
                }
                _metroFont = value;
                this.Font = value;
                Invalidate();
            }
        }
    }

    #endregion

    #region Calendar Control

    /// <summary>
    /// メトロスタイルのカレンダーコントロール
    /// </summary>
    public class MetroCalendar : MetroControlBase
    {
        #region Fields and Properties

        private DateTime _currentDate;
        private DateTime _selectedDate;
        private List<Appointment> _appointments;
        private Rectangle[] _dayRects;
        private Rectangle _headerRect;
        private Rectangle _prevMonthRect;
        private Rectangle _nextMonthRect;
        private int _cellSize;
        private string[] _dayNames = { "日", "月", "火", "水", "木", "金", "土" };
        private bool _isProcessingMonthChange = false;

        [Browsable(true)]
        [Category("Metro Calendar")]
        [Description("カレンダーの現在日付")]
        public DateTime CurrentDate
        {
            get { return _currentDate; }
            set
            {
                if (_currentDate != value)
                {
                    _currentDate = value;
                    CalculateRectangles();
                    Invalidate();
                }
            }
        }

        [Browsable(true)]
        [Category("Metro Calendar")]
        [Description("カレンダーの選択された日付")]
        public DateTime SelectedDate
        {
            get { return _selectedDate; }
            set
            {
                if (_selectedDate != value)
                {
                    _selectedDate = value;
                    OnSelectedDateChanged(EventArgs.Empty);
                    Invalidate();
                }
            }
        }

        [Browsable(false)]
        public List<Appointment> Appointments
        {
            get { return _appointments; }
        }

        [Browsable(true)]
        [Category("Metro Calendar")]
        [Description("日付が選択された時に発生するイベント")]
        public event EventHandler SelectedDateChanged;

        [Browsable(true)]
        [Category("Metro Calendar")]
        [Description("予定が追加された時に発生するイベント")]
        public event EventHandler AppointmentAdded;

        [Browsable(true)]
        [Category("Metro Calendar")]
        [Description("予定が変更された時に発生するイベント")]
        public event EventHandler AppointmentChanged;

        #endregion

        #region Constructor

        public MetroCalendar()
        {
            _currentDate = DateTime.Now;
            _selectedDate = DateTime.Now;
            _appointments = new List<Appointment>();
            _dayRects = new Rectangle[42]; // 最大6週 x 7日
            Size = new Size(280, 320);
            CalculateRectangles();
        }

        #endregion

        #region Methods

        protected virtual void OnSelectedDateChanged(EventArgs e)
        {
            if (SelectedDateChanged != null)
            {
                SelectedDateChanged(this, e);
            }
        }

        protected virtual void OnAppointmentAdded(EventArgs e)
        {
            if (AppointmentAdded != null)
            {
                AppointmentAdded(this, e);
            }
        }

        protected virtual void OnAppointmentChanged(EventArgs e)
        {
            if (AppointmentChanged != null)
            {
                AppointmentChanged(this, e);
            }
        }

        /// <summary>
        /// 予定を追加する
        /// </summary>
        public void AddAppointment(Appointment appointment)
        {
            _appointments.Add(appointment);
            OnAppointmentAdded(EventArgs.Empty);
            Invalidate();
        }

        /// <summary>
        /// 予定を削除する
        /// </summary>
        public void RemoveAppointment(Appointment appointment)
        {
            if (_appointments.Contains(appointment))
            {
                _appointments.Remove(appointment);
                Invalidate();
            }
        }

        /// <summary>
        /// 指定した日付の予定を取得する
        /// </summary>
        public List<Appointment> GetAppointmentsForDate(DateTime date)
        {
            return _appointments.Where(a => a.IsOnDate(date)).ToList();
        }

        /// <summary>
        /// 前の月に移動する
        /// </summary>
        public void PreviousMonth()
        {
            if (_isProcessingMonthChange)
                return;

            _isProcessingMonthChange = true;
            try
            {
                CurrentDate = CurrentDate.AddMonths(-1);
            }
            finally
            {
                // 非同期で処理完了フラグをクリア
                ThreadPool.QueueUserWorkItem((state) => 
                {
                    Thread.Sleep(50);
                    _isProcessingMonthChange = false;
                });
            }
        }

        /// <summary>
        /// 次の月に移動する
        /// </summary>
        public void NextMonth()
        {
            if (_isProcessingMonthChange)
                return;

            _isProcessingMonthChange = true;
            try
            {
                CurrentDate = CurrentDate.AddMonths(1);
            }
            finally
            {
                // 非同期で処理完了フラグをクリア
                ThreadPool.QueueUserWorkItem((state) => 
                {
                    Thread.Sleep(50);
                    _isProcessingMonthChange = false;
                });
            }
        }

        /// <summary>
        /// 日付表示用の矩形を計算する
        /// </summary>
        private void CalculateRectangles()
        {
            // ヘッダー部分の矩形
            _headerRect = new Rectangle(0, 0, Width, 40);
            
            // 前月/次月ボタンの矩形
            _prevMonthRect = new Rectangle(10, 10, 20, 20);
            _nextMonthRect = new Rectangle(Width - 30, 10, 20, 20);

            // セルサイズの計算
            _cellSize = Math.Min((Width - 2) / 7, (Height - _headerRect.Height - 30) / 7);

            // 日付セルの矩形
            int startX = (Width - (_cellSize * 7)) / 2;
            int startY = _headerRect.Bottom + 5;

            // 曜日ヘッダー分を下げる
            startY += _cellSize;

            // 表示する日付の取得
            DateTime firstDayOfMonth = new DateTime(_currentDate.Year, _currentDate.Month, 1);
            int daysInMonth = DateTime.DaysInMonth(_currentDate.Year, _currentDate.Month);
            
            // 月初の曜日
            int startDayOfWeek = (int)firstDayOfMonth.DayOfWeek;
            
            // 1日ずらす（日本の場合は日曜が0）
            
            // 日付配置
            for (int i = 0; i < 42; i++)
            {
                int row = i / 7;
                int col = i % 7;
                _dayRects[i] = new Rectangle(
                    startX + (col * _cellSize),
                    startY + (row * _cellSize),
                    _cellSize,
                    _cellSize);
            }
        }

        /// <summary>
        /// 予定の状態を更新
        /// </summary>
        public void UpdateAppointmentStatuses()
        {
            foreach (Appointment appointment in _appointments)
            {
                appointment.UpdateStatus();
            }
            Invalidate();
        }

        #endregion

        #region Event Handlers

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            CalculateRectangles();
        }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            base.OnMouseClick(e);

            // 前月/次月ボタンのクリック
            if (_prevMonthRect.Contains(e.Location))
            {
                PreviousMonth();
                return;
            }

            if (_nextMonthRect.Contains(e.Location))
            {
                NextMonth();
                return;
            }

            // 日付のクリック
            for (int i = 0; i < _dayRects.Length; i++)
            {
                if (_dayRects[i].Contains(e.Location))
                {
                    // 表示している月の1日目
                    DateTime firstDayOfMonth = new DateTime(_currentDate.Year, _currentDate.Month, 1);
                    
                    // 1日の曜日（0=日曜）
                    int startDayOfWeek = (int)firstDayOfMonth.DayOfWeek;
                    
                    // クリックした位置の日付を計算
                    int dayOffset = i - startDayOfWeek;
                    
                    // 当月の有効な日付の場合
                    if (dayOffset >= 0 && dayOffset < DateTime.DaysInMonth(_currentDate.Year, _currentDate.Month))
                    {
                        // 選択日付を設定
                        SelectedDate = new DateTime(_currentDate.Year, _currentDate.Month, dayOffset + 1);
                    }
                    
                    break;
                }
            }
        }

        #endregion

        #region Paint Methods

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            // 背景描画
            using (SolidBrush backBrush = new SolidBrush(BackColor))
            {
                g.FillRectangle(backBrush, ClientRectangle);
            }

            // ヘッダー描画
            DrawHeader(g);
            
            // 曜日ヘッダー描画
            DrawDayHeaders(g);
            
            // 日付セル描画
            DrawDayCells(g);
        }

        /// <summary>
        /// ヘッダー部分を描画
        /// </summary>
        private void DrawHeader(Graphics g)
        {
            // ヘッダー背景
            using (SolidBrush headerBrush = new SolidBrush(MetroColors.Primary))
            {
                DrawingUtils.FillRoundedRectangle(g, headerBrush, new Rectangle(
                    _headerRect.X, _headerRect.Y, _headerRect.Width, _headerRect.Height), 5);
            }

            // 月表示
            string monthText = _currentDate.ToString("yyyy年 M月");
            using (Font titleFont = new Font(MetroFont.FontFamily, 12f, FontStyle.Regular))
            using (SolidBrush textBrush = new SolidBrush(Color.White))
            {
                SizeF textSize = g.MeasureString(monthText, titleFont);
                g.DrawString(monthText, titleFont, textBrush,
                    _headerRect.X + (_headerRect.Width - textSize.Width) / 2,
                    _headerRect.Y + (_headerRect.Height - textSize.Height) / 2);
            }

            // 前月/次月ボタン
            using (Pen buttonPen = new Pen(Color.White, 2f))
            {
                // 前月ボタン
                g.DrawLine(buttonPen,
                    _prevMonthRect.X + _prevMonthRect.Width * 0.7f, _prevMonthRect.Y + _prevMonthRect.Height * 0.3f,
                    _prevMonthRect.X + _prevMonthRect.Width * 0.3f, _prevMonthRect.Y + _prevMonthRect.Height * 0.5f);
                g.DrawLine(buttonPen,
                    _prevMonthRect.X + _prevMonthRect.Width * 0.3f, _prevMonthRect.Y + _prevMonthRect.Height * 0.5f,
                    _prevMonthRect.X + _prevMonthRect.Width * 0.7f, _prevMonthRect.Y + _prevMonthRect.Height * 0.7f);

                // 次月ボタン
                g.DrawLine(buttonPen,
                    _nextMonthRect.X + _nextMonthRect.Width * 0.3f, _nextMonthRect.Y + _nextMonthRect.Height * 0.3f,
                    _nextMonthRect.X + _nextMonthRect.Width * 0.7f, _nextMonthRect.Y + _nextMonthRect.Height * 0.5f);
                g.DrawLine(buttonPen,
                    _nextMonthRect.X + _nextMonthRect.Width * 0.7f, _nextMonthRect.Y + _nextMonthRect.Height * 0.5f,
                    _nextMonthRect.X + _nextMonthRect.Width * 0.3f, _nextMonthRect.Y + _nextMonthRect.Height * 0.7f);
            }
        }

        /// <summary>
        /// 曜日ヘッダーを描画
        /// </summary>
        private void DrawDayHeaders(Graphics g)
        {
            int startX = (Width - (_cellSize * 7)) / 2;
            int startY = _headerRect.Bottom + 5;

            using (Font dayHeaderFont = new Font(MetroFont.FontFamily, 9f, FontStyle.Regular))
            {
                for (int i = 0; i < 7; i++)
                {
                    Brush textBrush;

                    // 日曜は赤、土曜は青
                    if (i == 0)
                        textBrush = new SolidBrush(MetroColors.Danger);
                    else if (i == 6)
                        textBrush = new SolidBrush(MetroColors.Info);
                    else
                        textBrush = new SolidBrush(MetroColors.TextLight);

                    Rectangle rect = new Rectangle(startX + (i * _cellSize), startY, _cellSize, _cellSize);
                    StringFormat sf = new StringFormat {Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                    
                    g.DrawString(_dayNames[i], dayHeaderFont, textBrush, rect, sf);
                    
                    textBrush.Dispose();
                }
            }
        }

        /// <summary>
        /// 日付セルを描画
        /// </summary>
        private void DrawDayCells(Graphics g)
        {
            // 表示する日付の取得
            DateTime firstDayOfMonth = new DateTime(_currentDate.Year, _currentDate.Month, 1);
            int daysInMonth = DateTime.DaysInMonth(_currentDate.Year, _currentDate.Month);
            
            // 月初の曜日 (0 = 日曜)
            int startDayOfWeek = (int)firstDayOfMonth.DayOfWeek;

            using (Font dayFont = new Font(MetroFont.FontFamily, 9f, FontStyle.Regular))
            using (Font todayFont = new Font(MetroFont.FontFamily, 9f, FontStyle.Bold))
            using (StringFormat sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
            {
                // 前月の日付を取得
                DateTime prevMonth = _currentDate.AddMonths(-1);
                int daysInPrevMonth = DateTime.DaysInMonth(prevMonth.Year, prevMonth.Month);

                // 日付セル描画
                for (int i = 0; i < 42; i++)
                {
                    int dayOffset = i - startDayOfWeek;
                    int dayNumber;
                    Brush textBrush;
                    Brush backBrush = null;
                    bool isCurrentMonth = false;
                    DateTime cellDate;

                    if (dayOffset < 0)
                    {
                        // 前月の日付
                        dayNumber = daysInPrevMonth + dayOffset + 1;
                        textBrush = new SolidBrush(Color.FromArgb(100, MetroColors.TextLight));
                        cellDate = new DateTime(prevMonth.Year, prevMonth.Month, dayNumber);
                    }
                    else if (dayOffset < daysInMonth)
                    {
                        // 当月の日付
                        dayNumber = dayOffset + 1;
                        isCurrentMonth = true;

                        // 曜日に応じた色
                        int dayOfWeek = i % 7;
                        if (dayOfWeek == 0)
                            textBrush = new SolidBrush(MetroColors.Danger);
                        else if (dayOfWeek == 6)
                            textBrush = new SolidBrush(MetroColors.Info);
                        else
                            textBrush = new SolidBrush(MetroColors.Text);

                        cellDate = new DateTime(_currentDate.Year, _currentDate.Month, dayNumber);

                        // 今日の日付の場合
                        if (cellDate.Date == DateTime.Now.Date)
                        {
                            backBrush = new SolidBrush(Color.FromArgb(50, MetroColors.Primary));
                        }

                        // 選択中の日付の場合
                        if (cellDate.Date == _selectedDate.Date)
                        {
                            backBrush = new SolidBrush(Color.FromArgb(80, MetroColors.Primary));
                        }
                    }
                    else
                    {
                        // 翌月の日付
                        dayNumber = dayOffset - daysInMonth + 1;
                        textBrush = new SolidBrush(Color.FromArgb(100, MetroColors.TextLight));
                        cellDate = new DateTime(_currentDate.AddMonths(1).Year, _currentDate.AddMonths(1).Month, dayNumber);
                    }

                    // 背景描画
                    if (backBrush != null)
                    {
                        Rectangle backRect = new Rectangle(
                            _dayRects[i].X + 2,
                            _dayRects[i].Y + 2,
                            _dayRects[i].Width - 4,
                            _dayRects[i].Height - 4);
                        
                        DrawingUtils.FillRoundedRectangle(g, backBrush, backRect, 5);
                        backBrush.Dispose();
                    }

                    // 日付テキスト描画
                    Font useFont = (cellDate.Date == DateTime.Now.Date) ? todayFont : dayFont;
                    g.DrawString(dayNumber.ToString(), useFont, textBrush, _dayRects[i], sf);

                    // 予定マーク描画
                    if (isCurrentMonth)
                    {
                        List<Appointment> dayAppointments = GetAppointmentsForDate(cellDate);
                        
                        // 予定がある場合
                        if (dayAppointments.Count > 0)
                        {
                            bool hasOverdue = dayAppointments.Any(a => a.Status == AppointmentStatus.Overdue);
                            bool hasInProgress = !hasOverdue && dayAppointments.Any(a => a.Status == AppointmentStatus.InProgress);
                            bool hasCompleted = !hasOverdue && !hasInProgress && dayAppointments.All(a => a.Status == AppointmentStatus.Completed);

                            Color dotColor;
                            if (hasOverdue)
                                dotColor = MetroColors.Danger;
                            else if (hasInProgress)
                                dotColor = MetroColors.InProgress;
                            else if (hasCompleted)
                                dotColor = MetroColors.Success;
                            else
                                dotColor = MetroColors.Primary;

                            int markSize = Math.Max(4, _cellSize / 10);
                            Rectangle markRect = new Rectangle(
                                _dayRects[i].X + (_dayRects[i].Width - markSize) / 2,
                                _dayRects[i].Y + _dayRects[i].Height - markSize - 2,
                                markSize,
                                markSize);

                            using (SolidBrush markBrush = new SolidBrush(dotColor))
                            {
                                g.FillEllipse(markBrush, markRect);
                            }
                        }
                    }

                    textBrush.Dispose();
                }
            }
        }

        #endregion
    }

    #endregion

    #region Digital Clock Control

    /// <summary>
    /// メトロスタイルのデジタル時計コントロール（直線タイムライン版）
    /// </summary>
    public class MetroDigitalClock : MetroControlBase
    {
        #region Fields and Properties

        private System.Windows.Forms.Timer _timer;
        private MetroCalendar _calendar;
        private DateTime _currentTime;
        private Rectangle _timeRect;
        private Rectangle _dateRect;
        private Rectangle _nextTaskRect;
        private Rectangle _timelineRect;
        private Color _progressColor;
        private bool _showSeconds = true;
        private Color[] _taskColors = new Color[3]; // 予定の色（近い予定、やや遠い予定、遠い予定）
        private Dictionary<int, Appointment> _timelineAppointments = new Dictionary<int, Appointment>();
        // ポップアップ表示用の変数
        private bool _showMarkerTooltip = false;
        private Rectangle _tooltipRect;
        private Appointment _hoveredAppointment;
        private Point _tooltipPosition;

        [System.ComponentModel.Browsable(true)]
        [System.ComponentModel.Category("Metro Clock")]
        [System.ComponentModel.Description("現在の日時")]
        public DateTime CurrentTime 
        {
            get { return _currentTime; }
            private set
            {
                if (_currentTime != value)
                {
                    _currentTime = value;
                    Invalidate();
                }
            }
        }

        [System.ComponentModel.Browsable(true)]
        [System.ComponentModel.Category("Metro Clock")]
        [System.ComponentModel.Description("進捗バーの色")]
        public Color ProgressColor
        {
            get { return _progressColor; }
            set
            {
                if (_progressColor != value)
                {
                    _progressColor = value;
                    Invalidate();
                }
            }
        }

        [System.ComponentModel.Browsable(true)]
        [System.ComponentModel.Category("Metro Clock")]
        [System.ComponentModel.Description("秒表示の有無")]
        public bool ShowSeconds
        {
            get { return _showSeconds; }
            set
            {
                if (_showSeconds != value)
                {
                    _showSeconds = value;
                    Invalidate();
                }
            }
        }

        [System.ComponentModel.Browsable(true)]
        [System.ComponentModel.Category("Metro Clock")]
        [System.ComponentModel.Description("同期するカレンダーコントロール")]
        public MetroCalendar Calendar
        {
            get { return _calendar; }
            set
            {
                if (_calendar != value)
                {
                    _calendar = value;
                    Invalidate();
                }
            }
        }

        #endregion

        #region Constructor

        public MetroDigitalClock()
        {
            _currentTime = DateTime.Now;
            _progressColor = MetroColors.Info;
            Size = new Size(350, 150);
            
            // 予定の色を設定
            _taskColors[0] = Color.FromArgb(255, 87, 34);   // 近い予定（0～10日）：深いオレンジ
            _taskColors[1] = Color.FromArgb(255, 152, 0);   // やや遠い予定（11～20日）：オレンジ
            _taskColors[2] = Color.FromArgb(255, 193, 7);   // 遠い予定（21～30日）：黄色

            CalculateRectangles();

            // タイマー設定
            _timer = new System.Windows.Forms.Timer();
            _timer.Interval = 1000; // 1秒
            _timer.Tick += Timer_Tick;
            _timer.Start();

            // マウスイベント
            this.MouseClick += MetroDigitalClock_MouseClick;
            this.MouseMove += MetroDigitalClock_MouseMove;
            this.MouseLeave += MetroDigitalClock_MouseLeave;
        }

        #endregion

        #region Methods

        /// <summary>
        /// 表示用の矩形を計算する
        /// </summary>
        private void CalculateRectangles()
        {
            // 時計表示領域
            _timeRect = new Rectangle(0, 10, Width, 40);
            
            // 日付表示領域
            _dateRect = new Rectangle(0, _timeRect.Bottom - 5, Width, 20);
            
            // 次の予定表示領域
            _nextTaskRect = new Rectangle(0, _dateRect.Bottom + 5, Width, 20);
            
            // 直線タイムライン表示領域
            _timelineRect = new Rectangle(
                20,
                _nextTaskRect.Bottom + 15,
                Width - 40,
                30);
        }

        /// <summary>
        /// 次の予定を取得する
        /// </summary>
        private Appointment GetNextAppointment()
        {
            if (_calendar == null)
                return null;

            DateTime now = DateTime.Now;
            
            // 未完了かつ現在時刻以降の予定を開始時刻の昇順でフィルタリング
            var futureAppointments = _calendar.Appointments
                .Where(a => !a.IsCompleted && a.StartTime >= now)
                .OrderBy(a => a.StartTime);
            
            if (futureAppointments.Count() > 0)
                return futureAppointments.First();
            else
                return null;
        }

        /// <summary>
        /// 30日以内の予定を取得する
        /// </summary>
        private List<Appointment> GetFutureAppointments(int daysLimit)
        {
            if (_calendar == null)
                return new List<Appointment>();

            DateTime now = DateTime.Now;
            DateTime limit = now.AddDays(daysLimit);
            
            // 未完了かつ現在時刻から指定日数以内の予定を開始時刻の昇順でフィルタリング
            return _calendar.Appointments
                .Where(a => !a.IsCompleted && a.StartTime >= now && a.StartTime < limit)
                .OrderBy(a => a.StartTime)
                .ToList();
        }

        /// <summary>
        /// 日付を位置（X座標）に変換
        /// </summary>
        private float DateToPosition(DateTime date, DateTime baseDate, Rectangle bounds, DateTime limitDate)
        {
            TimeSpan totalSpan = limitDate - baseDate;
            TimeSpan dateSpan = date - baseDate;
            
            if (totalSpan.TotalDays <= 0)
                return bounds.Left;
                
            // スケールを計算（0～1の範囲）
            float scale = (float)(dateSpan.TotalDays / totalSpan.TotalDays);
            
            // X座標に変換
            return bounds.Left + (bounds.Width * scale);
        }

        /// <summary>
        /// 時間差を表す文字列を取得
        /// </summary>
        private string GetTimeDifferenceText(DateTime target, DateTime reference)
        {
            TimeSpan diff = target - reference;
            
            if (diff.TotalMinutes < 60)
            {
                // 60分未満
                return String.Format("{0}分後 ({1})", 
                    (int)diff.TotalMinutes, 
                    target.ToString("HH:mm"));
            }
            else if (diff.TotalHours < 24)
            {
                // 24時間未満
                return String.Format("{0}時間{1}分後 ({2})",
                    (int)diff.TotalHours,
                    diff.Minutes,
                    target.ToString("HH:mm"));
            }
            else
            {
                // 24時間以上
                return target.ToString("MM/dd HH:mm");
            }
        }

        #endregion

        #region Event Handlers

        private void Timer_Tick(object sender, EventArgs e)
        {
            // 前の分と現在の分を比較するために前の時間を保存
            DateTime previousTime = CurrentTime;
            
            // 現在時刻を更新
            CurrentTime = DateTime.Now;
            
            // 分が変わった場合のみ通知の更新を行う
            if (previousTime.Minute != CurrentTime.Minute)
            {
                // アクティブな通知ウィンドウの時間表示を更新
                foreach (NotificationWindow notification in NotificationWindow.ActiveNotifications)
                {
                    notification.UpdateTimeRemainingText();
                }
            }
            
            // カレンダーがある場合は予定の状態を更新
            if (_calendar != null)
            {
                _calendar.UpdateAppointmentStatuses();
            }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            CalculateRectangles();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_timer != null)
                {
                    _timer.Stop();
                    _timer.Dispose();
                    _timer = null;
                }
            }
            base.Dispose(disposing);
        }
        
        private void MetroDigitalClock_MouseClick(object sender, MouseEventArgs e)
        {
            // タイムライン上のクリック位置をチェック
            foreach (var kvp in _timelineAppointments)
            {
                Rectangle markerRect = new Rectangle(kvp.Key - 6, 
                    (int)(_timelineRect.Top + _timelineRect.Height / 2 - 6), 12, 12);
                
                if (markerRect.Contains(e.Location))
                {
                    // 対応する予定の日付をカレンダーで選択
                    if (_calendar != null)
                    {
                        DateTime appointmentDate = kvp.Value.StartTime.Date;
                        
                        // 現在のカレンダー表示月と異なる月の場合
                        if (appointmentDate.Year != _calendar.CurrentDate.Year || 
                            appointmentDate.Month != _calendar.CurrentDate.Month)
                        {
                            // カレンダーの表示月を予定の月に変更
                            _calendar.CurrentDate = new DateTime(appointmentDate.Year, appointmentDate.Month, 1);
                        }
                        
                        // 日付を選択
                        _calendar.SelectedDate = appointmentDate;
                    }
                    break;
                }
            }
        }
        
        // マウス移動イベントハンドラ - マーカーのホバー検出
        private void MetroDigitalClock_MouseMove(object sender, MouseEventArgs e)
        {
            bool found = false;
            
            // タイムライン上のマーカーにマウスが重なっているかチェック
            foreach (var kvp in _timelineAppointments)
            {
                Rectangle markerRect = new Rectangle(kvp.Key - 6, 
                    (int)(_timelineRect.Top + _timelineRect.Height / 2 - 6), 12, 12);
                
                if (markerRect.Contains(e.Location))
                {
                    _hoveredAppointment = kvp.Value;
                    
                    // ツールチップの位置とサイズを計算
                    using (Font titleFont = new Font(MetroFont.FontFamily, 9f))
                    using (Font dateFont = new Font(MetroFont.FontFamily, 9f, FontStyle.Bold))
                    using (Graphics g = this.CreateGraphics())
                    {
                        string dateTimeText = String.Format("{0:yyyy/MM/dd} {0:HH:mm}-{1:HH:mm}", 
                            _hoveredAppointment.StartTime, _hoveredAppointment.EndTime);
                        string titleText = _hoveredAppointment.Title;
                        
                        SizeF dateTimeSize = g.MeasureString(dateTimeText, dateFont);
                        SizeF titleSize = g.MeasureString(titleText, titleFont);
                        
                        // ツールチップの幅を計算（日付時刻とタイトルの長い方 + 余白）
                        int tooltipWidth = Math.Max((int)dateTimeSize.Width, (int)titleSize.Width) + 20;
                        int tooltipHeight = (int)(dateTimeSize.Height + titleSize.Height) + 20;
                        
                        // ツールチップが画面からはみ出さないよう位置を調整
                        int tooltipX = Math.Min(e.X - tooltipWidth / 2, Width - tooltipWidth - 5);
                        tooltipX = Math.Max(tooltipX, 5);
                        
                        // マーカーの上に表示
                        int tooltipY = markerRect.Y - tooltipHeight - 5;
                        if (tooltipY < 0) // 上に表示できない場合は下に表示
                            tooltipY = markerRect.Bottom + 5;
                        
                        _tooltipRect = new Rectangle(tooltipX, tooltipY, tooltipWidth, tooltipHeight);
                        _tooltipPosition = new Point(tooltipX + 10, tooltipY + 10);
                    }
                    
                    _showMarkerTooltip = true;
                    found = true;
                    this.Invalidate();
                    break;
                }
            }
            
            // マーカーから離れた場合、ツールチップを非表示に
            if (!found && _showMarkerTooltip)
            {
                _showMarkerTooltip = false;
                _hoveredAppointment = null;
                this.Invalidate();
            }
        }

        // マウス離脱イベントハンドラ
        private void MetroDigitalClock_MouseLeave(object sender, EventArgs e)
        {
            if (_showMarkerTooltip)
            {
                _showMarkerTooltip = false;
                _hoveredAppointment = null;
                this.Invalidate();
            }
        }

        #endregion

        #region Paint Methods

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            // 背景描画
            using (SolidBrush backBrush = new SolidBrush(BackColor))
            {
                g.FillRectangle(backBrush, ClientRectangle);
            }

            // 背景矩形
            using (SolidBrush brush = new SolidBrush(Color.FromArgb(20, MetroColors.Primary)))
            {
                DrawingUtils.FillRoundedRectangle(g, brush, new Rectangle(0, 0, Width, Height), 10);
            }

            // 直線タイムライン描画
            DrawLinearTimeline(g);

            // 時刻描画
            string timeText = _showSeconds 
                ? _currentTime.ToString("HH:mm:ss")
                : _currentTime.ToString("HH:mm");
                
            using (Font timeFont = new Font(MetroFont.FontFamily, 24f, FontStyle.Regular))
            using (SolidBrush timeBrush = new SolidBrush(MetroColors.Primary))
            {
                StringFormat sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                g.DrawString(timeText, timeFont, timeBrush, _timeRect, sf);
            }

            // 日付描画
            string dateText = _currentTime.ToString("yyyy年MM月dd日 (ddd)");
            using (Font dateFont = new Font(MetroFont.FontFamily, 9f, FontStyle.Regular))
            using (SolidBrush dateBrush = new SolidBrush(MetroColors.TextLight))
            {
                StringFormat sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                g.DrawString(dateText, dateFont, dateBrush, _dateRect, sf);
            }

            // 次の予定情報
            List<Appointment> nextAppointments = GetFutureAppointments(1); // 1日以内の予定
            string nextTaskText = "次の予定はありません";
            
            if (nextAppointments.Count > 0)
            {
                Appointment nextAppointment = nextAppointments[0];
                TimeSpan timeToNext = nextAppointment.StartTime - _currentTime;
                
                if (timeToNext.TotalMinutes < 60)
                {
                    nextTaskText = String.Format("次の予定: {0} (あと{1}分)", 
                        nextAppointment.Title, 
                        (int)timeToNext.TotalMinutes);
                }
                else
                {
                    nextTaskText = String.Format("次の予定: {0} (あと{1}時間{2}分)", 
                        nextAppointment.Title, 
                        (int)timeToNext.TotalHours,
                        timeToNext.Minutes);
                }
            }

            using (Font taskFont = new Font(MetroFont.FontFamily, 9f, FontStyle.Regular))
            using (SolidBrush taskBrush = new SolidBrush(MetroColors.Text))
            {
                StringFormat sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                g.DrawString(nextTaskText, taskFont, taskBrush, _nextTaskRect, sf);
            }
            
            // ツールチップの描画
            if (_showMarkerTooltip && _hoveredAppointment != null)
            {   
                // ツールチップの背景
                using (SolidBrush backBrush = new SolidBrush(Color.FromArgb(250, 250, 250)))
                using (Pen borderPen = new Pen(Color.FromArgb(200, 200, 200), 1))
                {
                    DrawingUtils.FillRoundedRectangle(g, backBrush, _tooltipRect, 5);
                    DrawingUtils.DrawRoundedRectangle(g, borderPen, _tooltipRect, 5);
                }
                
                // 日付と時間の描画
                string dateTimeText = String.Format("{0:yyyy/MM/dd} {0:HH:mm}-{1:HH:mm}", 
                    _hoveredAppointment.StartTime, _hoveredAppointment.EndTime);
                
                using (Font dateFont = new Font(MetroFont.FontFamily, 9f, FontStyle.Bold))
                using (SolidBrush textBrush = new SolidBrush(MetroColors.Primary))
                {
                    g.DrawString(dateTimeText, dateFont, textBrush, _tooltipPosition);
                }
                
                // タイトルの描画
                string titleText = _hoveredAppointment.Title;
                
                using (Font titleFont = new Font(MetroFont.FontFamily, 9f))
                using (SolidBrush textBrush = new SolidBrush(MetroColors.Text))
                {
                    SizeF dateSize = g.MeasureString(dateTimeText, new Font(MetroFont.FontFamily, 9f, FontStyle.Bold));
                    Point titlePos = new Point(_tooltipPosition.X, _tooltipPosition.Y + (int)dateSize.Height + 2);
                    g.DrawString(titleText, titleFont, textBrush, titlePos);
                }
            }
        }

        /// <summary>
        /// 直線タイムラインを描画
        /// </summary>
        private void DrawLinearTimeline(Graphics g)
        {
            // タイムラインの背景
            using (SolidBrush timelineBackBrush = new SolidBrush(Color.FromArgb(240, 240, 240)))
            {
                DrawingUtils.FillRoundedRectangle(g, timelineBackBrush, _timelineRect, 5);
            }
            
            // 現在時刻
            DateTime now = DateTime.Now;
            
            // 30日後の日付（これが右端）
            DateTime limitDate = now.AddDays(30);
            
            // 30日以内の予定を取得
            List<Appointment> futureAppointments = GetFutureAppointments(30);
            
            // タイムラインの中央Y座標
            float timelineY = _timelineRect.Top + _timelineRect.Height / 2;
            
            // タイムライン上の予定マッピングをクリア
            _timelineAppointments.Clear();
            
            if (futureAppointments.Count == 0)
            {
                // 予定がない場合はタイムラインの軸のみ描画
                using (Pen timelinePen = new Pen(Color.FromArgb(200, MetroColors.TextLight), 2))
                {
                    g.DrawLine(timelinePen,
                        _timelineRect.Left,
                        timelineY,
                        _timelineRect.Right,
                        timelineY);
                }
                
                // 現在位置に丸いマーカー
                using (SolidBrush currentBrush = new SolidBrush(MetroColors.Primary))
                {
                    g.FillEllipse(currentBrush, _timelineRect.Left - 5, timelineY - 5, 10, 10);
                }
                
                return;
            }
            
            // 最後の予定（これが右端になる）
            Appointment lastAppointment = futureAppointments.Last();
            
            // 現在位置のX座標 (常に左端)
            float startX = _timelineRect.Left;
            
            // タイムラインの軸
            using (Pen timelinePen = new Pen(Color.FromArgb(200, MetroColors.TextLight), 2))
            {
                g.DrawLine(timelinePen,
                    startX,
                    timelineY,
                    _timelineRect.Right,
                    timelineY);
            }
            
            // 現在から今後の予定までの区間
            float lastX = startX;
            List<TimelineItem> timelineItems = new List<TimelineItem>();
            
            foreach (Appointment appointment in futureAppointments)
            {
                // 時間に応じて色を選択
                TimeSpan daysUntil = appointment.StartTime - now;
                int dayBucket = Math.Min(2, (int)(daysUntil.TotalDays / 10));
                Color taskColor = _taskColors[dayBucket];
                
                // 予定のX座標
                float appointmentX = DateToPosition(appointment.StartTime, now, _timelineRect, limitDate);
                
                // 予定を線分を描画（タイムラインのパス）
                using (Pen linePen = new Pen(Color.FromArgb(100, taskColor), 2))
                {
                    g.DrawLine(linePen, lastX, timelineY, appointmentX, timelineY);
                }

                // タイムラインアイテムに追加（後で上にマーカーを描画するため）
                TimelineItem item = new TimelineItem();
                item.Appointment = appointment;
                item.Position = new PointF(appointmentX, timelineY);
                item.Color = taskColor;
                item.IsAbove = futureAppointments.IndexOf(appointment) % 2 == 0;
                timelineItems.Add(item);
                                
                // 現在位置を更新
                lastX = appointmentX;
                
                // クリック判定用に座標を保存
                _timelineAppointments[(int)appointmentX] = appointment;
            }
            
            // 日付目盛りを追加（5日ごと）
            for (int i = 5; i <= 30; i += 5)
            {
                DateTime tickDate = now.AddDays(i);
                float tickX = DateToPosition(tickDate, now, _timelineRect, limitDate);
                
                // 目盛り線
                using (Pen tickPen = new Pen(Color.FromArgb(100, MetroColors.TextLight), 1))
                {
                    g.DrawLine(tickPen,
                        tickX, _timelineRect.Top + 5,
                        tickX, _timelineRect.Bottom - 5);
                }
                
                // 日付ラベル
                string tickLabel = tickDate.ToString("MM/dd");
                using (Font tickFont = new Font(MetroFont.FontFamily, 7f, FontStyle.Regular))
                using (SolidBrush tickBrush = new SolidBrush(MetroColors.TextLight))
                {
                    SizeF textSize = g.MeasureString(tickLabel, tickFont);
                    g.DrawString(tickLabel, tickFont, tickBrush,
                        tickX - textSize.Width / 2,
                        _timelineRect.Bottom + 12);
                }
            }
            
            // 現在位置に丸いマーカー
            using (SolidBrush currentBrush = new SolidBrush(MetroColors.Primary))
            {
                g.FillEllipse(currentBrush, startX - 5, timelineY - 5, 10, 10);
            }
            
            // 最後に予定のマーカーを描画（最上位レイヤー）
            foreach (TimelineItem item in timelineItems)
            {
                // 予定のマーカー
                using (SolidBrush markerBrush = new SolidBrush(item.Color))
                {
                    g.FillEllipse(markerBrush, 
                        item.Position.X - 6, 
                        item.Position.Y - 6, 
                        12, 12);
                    
                    if (item.Appointment.StartTime.Date == _calendar.SelectedDate.Date)
                    {
                        using (Pen highlightPen = new Pen(Color.FromArgb(80, 127, 255), 2f))
                        {
                            g.DrawEllipse(highlightPen,
                                item.Position.X - 7,
                                item.Position.Y - 7,
                                14, 14);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// タイムラインアイテム
        /// </summary>
        private class TimelineItem
        {
            public Appointment Appointment { get; set; }
            public PointF Position { get; set; }
            public Color Color { get; set; }
            public int PositionPattern { get; set; }
            public bool IsAbove { get; set; } 
        }

        #endregion
    }

    #endregion

    #region Task Management Control

    /// <summary>
    /// メトロスタイルのタスク管理コントロール
    /// </summary>
    public class MetroTaskManager : MetroControlBase
    {
        #region Fields and Properties

        private MetroCalendar _calendar;
        private NotificationManager _notificationManager;
        private List<Appointment> _tasks;
        private List<TaskItemUI> _taskUIItems;
        private int _itemHeight = 60;
        private int _scrollPosition = 0;
        private Rectangle _addButtonRect;
        private bool _addButtonHover = false;
        private System.Windows.Forms.Timer _updateTimer;
        private Label _dateHeaderLabel;
        
        // コンテキストメニュー用の変数
        private Appointment _selectedTaskForMenu;
        private Rectangle _contextMenuRect;
        private bool _showContextMenu;
        private Rectangle _addSubTaskMenuItemRect;
        private Rectangle _addMemoMenuItemRect;
        private bool _isAddSubTaskHover;
        private bool _isAddMemoHover;

        // メモページ表示用の変数
        private Appointment _selectedTaskForMemo;
        private SubTask _selectedSubTaskForMemo;
        private Rectangle _memoPageRect;
        private Rectangle _memoCloseButtonRect;
        private bool _showMemoPage;
        private bool _isMemoCloseButtonHover;
        private TextBox _memoTextBox;
        
        private System.Windows.Forms.Timer _titleScrollTimer;

        [System.ComponentModel.Browsable(true)]
        [System.ComponentModel.Category("Metro Task Manager")]
        [System.ComponentModel.Description("同期するカレンダーコントロール")]
        public MetroCalendar Calendar
        {
            get { return _calendar; }
            set
            {
                if (_calendar != value)
                {
                    if (_calendar != null)
                    {
                        _calendar.AppointmentAdded -= Calendar_AppointmentAdded;
                        _calendar.AppointmentChanged -= Calendar_AppointmentChanged;
                        _calendar.SelectedDateChanged -= Calendar_SelectedDateChanged;
                    }

                    _calendar = value;

                    if (_calendar != null)
                    {
                        _calendar.AppointmentAdded += Calendar_AppointmentAdded;
                        _calendar.AppointmentChanged += Calendar_AppointmentChanged;
                        _calendar.SelectedDateChanged += Calendar_SelectedDateChanged;
                        
                        // 通知マネージャーを設定
                        if (_notificationManager != null)
                        {
                            _notificationManager.Dispose();
                        }
                        _notificationManager = new NotificationManager(_calendar.Appointments);
                        
                        LoadTasks();
                    }

                    Invalidate();
                }
            }
        }

        [System.ComponentModel.Browsable(true)]
        [System.ComponentModel.Category("Metro Task Manager")]
        [System.ComponentModel.Description("タスクが追加された時に発生するイベント")]
        public event EventHandler TaskAdded;

        [System.ComponentModel.Browsable(true)]
        [System.ComponentModel.Category("Metro Task Manager")]
        [System.ComponentModel.Description("タスクが変更された時に発生するイベント")]
        public event EventHandler TaskChanged;

        [System.ComponentModel.Browsable(true)]
        [System.ComponentModel.Category("Metro Task Manager")]
        [System.ComponentModel.Description("タスクが削除された時に発生するイベント")]
        public event EventHandler TaskDeleted;

        /// <summary>
        /// サブタスクを表すクラス
        /// </summary>
        public class SubTask
        {
            public string Title { get; set; }
            public bool IsCompleted { get; set; }
            public string Memo { get; set; }

            public SubTask()
            {
                Title = "新しいサブタスク";
                IsCompleted = false;
                Memo = "";
            }
        }
        
        #endregion

        #region Constructor

        public MetroTaskManager()
        {
            _tasks = new List<Appointment>();
            _taskUIItems = new List<TaskItemUI>();
            Size = new Size(350, 400);
            
            // 日付ヘッダー
            _dateHeaderLabel = new Label();
            _dateHeaderLabel.Text = "タスク一覧";
            _dateHeaderLabel.Font = new Font("Segoe UI", 12f, FontStyle.Bold);
            _dateHeaderLabel.ForeColor = MetroColors.Primary;
            _dateHeaderLabel.BackColor = Color.Transparent;
            _dateHeaderLabel.TextAlign = ContentAlignment.MiddleLeft;
            _dateHeaderLabel.Size = new Size(Width - 20, 30);
            _dateHeaderLabel.Location = new Point(10, 10);
            this.Controls.Add(_dateHeaderLabel);
            
            CalculateRectangles();

            // マウスイベント登録
            this.MouseMove += MetroTaskManager_MouseMove;
            this.MouseLeave += MetroTaskManager_MouseLeave;
            this.MouseWheel += MetroTaskManager_MouseWheel;
            this.MouseClick += MetroTaskManager_MouseClick;
            
            // タイマー設定
            _updateTimer = new System.Windows.Forms.Timer();
            _updateTimer.Interval = 10000; // 10秒ごとに状態更新
            _updateTimer.Tick += UpdateTimer_Tick;
            _updateTimer.Start();
            
    
            // タイトルスクロールタイマー設定
            _titleScrollTimer = new System.Windows.Forms.Timer();
            _titleScrollTimer.Interval = 50; // 50msごとに更新
            _titleScrollTimer.Tick += TitleScrollTimer_Tick;
            _titleScrollTimer.Start();
        }

        #endregion

        #region Methods
        protected virtual void OnTaskAdded(EventArgs e)
        {
            if (TaskAdded != null)
            {
                TaskAdded(this, e);
            }
        }

        protected virtual void OnTaskChanged(EventArgs e)
        {
            if (TaskChanged != null)
            {
                TaskChanged(this, e);
            }
        }

        protected virtual void OnTaskDeleted(EventArgs e)
        {
            if (TaskDeleted != null)
            {
                TaskDeleted(this, e);
            }
        }
        
        private void TitleScrollTimer_Tick(object sender, EventArgs e)
        {
            bool needRedraw = false;
            
            foreach (TaskItemUI item in _taskUIItems)
            {
                // タイトルテキストサイズとスクロール状態の確認
                using (Graphics g = this.CreateGraphics())
                using (Font titleFont = new Font(MetroFont.FontFamily, 10f, FontStyle.Bold))
                {
                    SizeF textSize = g.MeasureString(item.Task.Title, titleFont);
                    int titleWidth = 150; // 表示幅の制限
                    
                    // タイトルがスクロール対象かどうか確認
                    bool needsScrolling = textSize.Width > titleWidth;
                    
                    if (needsScrolling)
                    {
                        // カウンターを増加（タスクごとに独立）
                        item.ScrollCounter++;
                        
                        if (!item.IsTitleScrolling)
                        {
                            // スクロール待機中の場合
                            if (item.ScrollCounter >= item.ScrollDelay)
                            {
                                // 待機時間経過後、スクロール開始
                                item.IsTitleScrolling = true;
                                item.TitleScrollOffset = 0;
                                needRedraw = true;
                            }
                        }
                        else
                        {
                            // スクロール中の場合
                            item.TitleScrollOffset += item.ScrollSpeed;
                            
                            // スクロールが終わったら再初期化
                            if (item.TitleScrollOffset > textSize.Width - titleWidth + 20)
                            {
                                item.IsTitleScrolling = false;
                                item.TitleScrollOffset = 0;
                                item.ScrollCounter = 0; // カウンターリセット
                            }
                            
                            needRedraw = true;
                        }
                    }
                    else
                    {
                        // スクロール不要の場合
                        item.TitleScrollOffset = 0;
                        item.IsTitleScrolling = false;
                        item.ScrollCounter = 0;
                    }
                }
            }
            
            // 再描画が必要な場合のみ実行
            if (needRedraw)
            {
                Invalidate();
            }
        }

        /// <summary>
        /// 表示用の矩形を計算する
        /// </summary>
        private void CalculateRectangles()
        {
            if (this.Width <= 0 || this.Height <= 0)
                return; // サイズが有効でない場合は計算をスキップ

            _addButtonRect = new Rectangle(Width - 50, Height - 50, 40, 40);
            
            // _dateHeaderLabelがnullでないことを確認
            if (_dateHeaderLabel != null)
            {
                _dateHeaderLabel.Size = new Size(Width - 20, 30);
            }
            
            UpdateTaskUIItems();
        }

        /// <summary>
        /// カレンダーから選択された日付のタスクを読み込む
        /// </summary>
        private void LoadTasks()
        {
            if (_calendar == null)
                return;

            // タスクをクリア
            _tasks.Clear();
            
            // 選択された日付のタスクを取得
            _tasks = _calendar.GetAppointmentsForDate(_calendar.SelectedDate);
            
            // タスクの状態を更新
            foreach (Appointment task in _tasks)
            {
                task.UpdateStatus();
            }
            
            // 日付ヘッダーの更新
            UpdateDateHeader();
            
            // UIアイテムを更新
            UpdateTaskUIItems();
            
            // 表示更新
            Invalidate();
        }

        /// <summary>
        /// 日付ヘッダーを更新
        /// </summary>
        private void UpdateDateHeader()
        {
            if (!_showMemoPage)
            {
                if (_calendar == null)
                    return;
                    
                DateTime selectedDate = _calendar.SelectedDate;
                
                // 今日の日付かどうか
                if (selectedDate.Date == DateTime.Today)
                {
                    _dateHeaderLabel.Text = String.Format("今日のタスク ({0})", selectedDate.ToString("yyyy年MM月dd日 (ddd)"));
                }
                else
                {
                    _dateHeaderLabel.Text = String.Format("{0}のタスク", selectedDate.ToString("yyyy年MM月dd日 (ddd)"));
                }
            }
            else
            {
                _dateHeaderLabel.Text = "";
            }
        }

        /// <summary>
        /// タスクUIアイテムを更新する
        /// </summary>
        private void UpdateTaskUIItems()
        {
            if (_taskUIItems == null)
                _taskUIItems = new List<TaskItemUI>();
            else
                _taskUIItems.Clear();
            
            // _dateHeaderLabelがnullの場合の対策
            int headerHeight = (_dateHeaderLabel != null) ? _dateHeaderLabel.Bottom + 10 : 50;
            int y = headerHeight - _scrollPosition;
            
            if (_tasks != null)
            {
                foreach (Appointment task in _tasks)
                {
                    Rectangle itemRect = new Rectangle(10, y, Width - 20, _itemHeight);
                    TaskItemUI taskItem = new TaskItemUI(task, itemRect);
                    
                    // リストの初期化は不要（コンストラクタでされている）
                    
                    _taskUIItems.Add(taskItem);
                    y += _itemHeight + 5;
                    
                    // サブタスクの矩形を更新
                    UpdateSubTaskRects(taskItem, ref y);
                }
            }
        }

        /// <summary>
        /// タスクを追加する
        /// </summary>
        public void AddTask(Appointment task)
        {
            if (_calendar != null)
            {
                _calendar.AddAppointment(task);
                LoadTasks();
                OnTaskAdded(EventArgs.Empty);
            }
        }

        /// <summary>
        /// タスクを削除する
        /// </summary>
        public void DeleteTask(Appointment task)
        {
            if (_calendar != null && _tasks.Contains(task))
            {
                // 確認ダイアログ
                DialogResult result = MessageBox.Show(
                    String.Format("予定「{0}」を削除しますか？", task.Title),
                    "予定の削除",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                
                if (result == DialogResult.Yes)
                {
                    _calendar.RemoveAppointment(task);
                    LoadTasks();
                    OnTaskDeleted(EventArgs.Empty);
                }
            }
        }

        /// <summary>
        /// タスクを編集する
        /// </summary>
        public void EditTask(Appointment task)
        {
            if (_calendar == null || !_tasks.Contains(task))
                return;

            ShowTaskDialog(task, false);
        }

        /// <summary>
        /// 新しいタスクダイアログを表示する
        /// </summary>
        private void ShowNewTaskDialog()
        {
            if (_calendar == null)
                return;

            // 新しいタスクを作成
            Appointment newTask = new Appointment();
            newTask.StartTime = new DateTime(_calendar.SelectedDate.Year, _calendar.SelectedDate.Month, _calendar.SelectedDate.Day, DateTime.Now.Hour, 0, 0);
            newTask.EndTime = newTask.StartTime.AddHours(1);
            
            ShowTaskDialog(newTask, true);
        }
        
        /// <summary>
        /// タスクダイアログを表示（新規追加または編集）
        /// </summary>
        private void ShowTaskDialog(Appointment task, bool isNew)
        {
            // 簡易的なダイアログを使用
            using (Form dialog = new Form())
            {
                dialog.Text = isNew ? "新しいタスク" : "タスクの編集";
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.Size = new Size(400, 320);
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.MaximizeBox = false;
                dialog.MinimizeBox = false;
                dialog.BackColor = Color.White;
                dialog.Font = new Font("Segoe UI", 9f);

                // タイトル
                Label titleLabel = new Label { Text = "タイトル:", Left = 20, Top = 20, Width = 80 };
                TextBox titleTextBox = new TextBox { Text = task.Title, Left = 120, Top = 20, Width = 240 };

                // 開始時間
                Label startLabel = new Label { Text = "開始時間:", Left = 20, Top = 50, Width = 80 };
                DateTimePicker startTimePicker = new DateTimePicker { 
                    Value = task.StartTime, 
                    Left = 120, 
                    Top = 50, 
                    Width = 240, 
                    Format = DateTimePickerFormat.Custom, 
                    CustomFormat = "yyyy/MM/dd HH:mm" 
                };

                // 終了時間
                Label endLabel = new Label { Text = "終了時間:", Left = 20, Top = 80, Width = 80 };
                DateTimePicker endTimePicker = new DateTimePicker { 
                    Value = task.EndTime, 
                    Left = 120, 
                    Top = 80, 
                    Width = 240, 
                    Format = DateTimePickerFormat.Custom, 
                    CustomFormat = "yyyy/MM/dd HH:mm" 
                };

                // 通知設定（数値入力＋プルダウン）
                Label notifyLabel = new Label { Text = "通知:", Left = 20, Top = 110, Width = 80 };
                
                // 数値入力
                NumericUpDown notifyNumeric = new NumericUpDown { 
                    Left = 120, 
                    Top = 110, 
                    Width = 80,
                    Minimum = 0,
                    Maximum = 1440, // 最大24時間（分単位）
                    Value = task.NotificationMinutesBefore
                };
                
                // 単位選択
                ComboBox notifyUnitComboBox = new ComboBox { 
                    DropDownStyle = ComboBoxStyle.DropDownList, 
                    Left = 210, 
                    Top = 110, 
                    Width = 150
                };
                notifyUnitComboBox.Items.AddRange(new object[] { "分前", "時間前", "日前" });
                
                // 現在の通知設定に基づいて単位を設定
                int unitIndex = 0; // デフォルトは「分前」
                decimal value = task.NotificationMinutesBefore;
                
                if (task.NotificationMinutesBefore >= 1440) // 24時間以上
                {
                    unitIndex = 2; // 「日前」
                    value = task.NotificationMinutesBefore / 1440;
                }
                else if (task.NotificationMinutesBefore >= 60) // 1時間以上
                {
                    unitIndex = 1; // 「時間前」
                    value = task.NotificationMinutesBefore / 60;
                }
                
                notifyUnitComboBox.SelectedIndex = unitIndex;
                notifyNumeric.Value = value;
                
                // 通知単位変更時の処理
                notifyUnitComboBox.SelectedIndexChanged += (sender, e) => {
                    decimal currentValue = notifyNumeric.Value;
                    switch (notifyUnitComboBox.SelectedIndex)
                    {
                        case 0: // 分前
                            notifyNumeric.Maximum = 1440;
                            break;
                        case 1: // 時間前
                            notifyNumeric.Maximum = 24;
                            break;
                        case 2: // 日前
                            notifyNumeric.Maximum = 30;
                            break;
                    }
                };
                
                // 完了チェックボックス（新規タスクの場合は非表示）
                CheckBox completedCheckbox = null;
                if (!isNew)
                {
                    completedCheckbox = new CheckBox {
                        Text = "完了",
                        Checked = task.IsCompleted,
                        Left = 120,
                        Top = 145,
                        Width = 240
                    };
                }

                // ボタン
                Button saveButton = new Button { 
                    Text = "保存", 
                    Left = 120, 
                    Top = 220, 
                    Width = 100, 
                    DialogResult = DialogResult.OK,
                    BackColor = MetroColors.Primary,
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                saveButton.FlatAppearance.BorderSize = 0;
                
                Button cancelButton = new Button { 
                    Text = "キャンセル", 
                    Left = 240, 
                    Top = 220, 
                    Width = 100, 
                    DialogResult = DialogResult.Cancel,
                    BackColor = Color.LightGray,
                    FlatStyle = FlatStyle.Flat
                };
                cancelButton.FlatAppearance.BorderSize = 0;
                
                // メモ入力エリア
                Label memoLabel = new Label { Text = "メモ:", Left = 20, Top = 140, Width = 80 };
                TextBox memoTextBox = new TextBox { 
                    Text = task.Memo, 
                    Left = 120, 
                    Top = 140, 
                    Width = 240,
                    Height = 60,
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical
                };
                
                // コントロール追加
                List<Control> controls = new List<Control> {
                    titleLabel, titleTextBox,
                    startLabel, startTimePicker,
                    endLabel, endTimePicker,
                    notifyLabel, notifyNumeric, notifyUnitComboBox,
                    saveButton, cancelButton
                };
                
                if (completedCheckbox != null)
                {
                    controls.Add(completedCheckbox);
                }
                
                dialog.Controls.AddRange(controls.ToArray());

                dialog.AcceptButton = saveButton;
                dialog.CancelButton = cancelButton;

                if (dialog.ShowDialog(this.FindForm()) == DialogResult.OK)
                {
                    task.Title = titleTextBox.Text;
                    task.StartTime = startTimePicker.Value;
                    task.EndTime = endTimePicker.Value;
                    task.Memo = memoTextBox.Text;

                    // 通知時間の設定
                    int notificationValue = (int)notifyNumeric.Value;
                    switch (notifyUnitComboBox.SelectedIndex)
                    {
                        case 0: // 分前
                            task.NotificationMinutesBefore = notificationValue;
                            break;
                        case 1: // 時間前
                            task.NotificationMinutesBefore = notificationValue * 60;
                            break;
                        case 2: // 日前
                            task.NotificationMinutesBefore = notificationValue * 1440; // 24時間 * 60分
                            break;
                    }
                    
                    // 完了状態（既存タスクの編集時のみ）
                    if (!isNew && completedCheckbox != null)
                    {
                        task.IsCompleted = completedCheckbox.Checked;
                    }

                    // タスク追加/更新
                    if (isNew)
                    {
                        AddTask(task);
                    }
                    else
                    {
                        task.UpdateStatus();
                        OnTaskChanged(EventArgs.Empty);
                        LoadTasks();
                    }
                }
            }
        }
        
        /// <summary>
        /// サブタスクを追加する
        /// </summary>
        private void AddSubTask()
        {
            if (_selectedTaskForMenu != null)
            {
                SubTask subTask = new SubTask();
                _selectedTaskForMenu.SubTasks.Add(subTask);
                _selectedTaskForMenu.ShowSubTasks = true; // 追加時は自動的に表示
                
                UpdateTaskUIItems();
                Invalidate();
            }
        }

        /// <summary>
        /// サブタスク編集ダイアログを表示
        /// </summary>
        private void EditSubTask(Appointment parentTask, SubTask subTask)
        {
            using (Form dialog = new Form())
            {
                dialog.Text = "サブタスクの編集";
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.Size = new Size(400, 250);
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.MaximizeBox = false;
                dialog.MinimizeBox = false;
                dialog.BackColor = Color.White;
                dialog.Font = new Font("Segoe UI", 9f);

                // タイトル
                Label titleLabel = new Label { Text = "タイトル:", Left = 20, Top = 20, Width = 80 };
                TextBox titleTextBox = new TextBox { Text = subTask.Title, Left = 120, Top = 20, Width = 240 };

                // 完了状態
                CheckBox completedCheckbox = new CheckBox {
                    Text = "完了",
                    Checked = subTask.IsCompleted,
                    Left = 120,
                    Top = 50,
                    Width = 240
                };

                // メモ
                Label memoLabel = new Label { Text = "メモ:", Left = 20, Top = 80, Width = 80 };
                TextBox memoTextBox = new TextBox { 
                    Text = subTask.Memo, 
                    Left = 120, 
                    Top = 80, 
                    Width = 240,
                    Height = 60,
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical
                };

                // ボタン
                Button saveButton = new Button { 
                    Text = "保存", 
                    Left = 120, 
                    Top = 160, 
                    Width = 100, 
                    DialogResult = DialogResult.OK,
                    BackColor = MetroColors.Primary,
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                saveButton.FlatAppearance.BorderSize = 0;
                
                Button cancelButton = new Button { 
                    Text = "キャンセル", 
                    Left = 240, 
                    Top = 160, 
                    Width = 100, 
                    DialogResult = DialogResult.Cancel,
                    BackColor = Color.LightGray,
                    FlatStyle = FlatStyle.Flat
                };
                cancelButton.FlatAppearance.BorderSize = 0;

                // コントロール追加
                dialog.Controls.AddRange(new Control[] {
                    titleLabel, titleTextBox,
                    completedCheckbox,
                    memoLabel, memoTextBox,
                    saveButton, cancelButton
                });

                dialog.AcceptButton = saveButton;
                dialog.CancelButton = cancelButton;

                if (dialog.ShowDialog(this.FindForm()) == DialogResult.OK)
                {
                    subTask.Title = titleTextBox.Text;
                    subTask.IsCompleted = completedCheckbox.Checked;
                    subTask.Memo = memoTextBox.Text;
                    
                    UpdateTaskUIItems();
                    Invalidate();
                }
            }
        }

        /// <summary>
        /// メモページを表示する
        /// </summary>
        private void ShowMemoPage(Appointment task, SubTask subTask = null)
        {
            _selectedTaskForMemo = task;
            _selectedSubTaskForMemo = subTask;
            _showMemoPage = true;
            
            string currentMemo = subTask != null ? subTask.Memo : task.Memo;
            
            // メモテキストボックスの作成
            if (_memoTextBox == null)
            {
                _memoTextBox = new TextBox();
                _memoTextBox.Multiline = true;
                _memoTextBox.ScrollBars = ScrollBars.Vertical;
                _memoTextBox.BorderStyle = BorderStyle.None;
                _memoTextBox.Font = new Font(MetroFont.FontFamily, 10f);
                this.Controls.Add(_memoTextBox);
            }
            
            _memoTextBox.Text = currentMemo;
            
            // メモページの位置とサイズ
            _memoPageRect = new Rectangle(50, 50, Width - 100, Height - 100);
            _memoCloseButtonRect = new Rectangle(
                _memoPageRect.Right - 30, 
                _memoPageRect.Y + 10, 
                20, 20);
                
            _memoTextBox.Location = new Point(
                _memoPageRect.X + 20,
                _memoPageRect.Y + 40);
            _memoTextBox.Size = new Size(
                _memoPageRect.Width - 40,
                _memoPageRect.Height - 60);
            _memoTextBox.Visible = true;
            _memoTextBox.BringToFront();
            
            UpdateDateHeader();
            Invalidate();
        }

        /// <summary>
        /// メモページを閉じる
        /// </summary>
        private void CloseMemoPage()
        {
            if (_memoTextBox != null && _selectedTaskForMemo != null)
            {
                // メモを保存
                if (_selectedSubTaskForMemo != null)
                    _selectedSubTaskForMemo.Memo = _memoTextBox.Text;
                else
                    _selectedTaskForMemo.Memo = _memoTextBox.Text;
                
                _memoTextBox.Visible = false;
            }
            
            _showMemoPage = false;
            _selectedTaskForMemo = null;
            _selectedSubTaskForMemo = null;
            
            UpdateTaskUIItems();
            UpdateDateHeader();
            Invalidate();
        }

        /// <summary>
        /// サブタスクの表示/非表示を切り替える
        /// </summary>
        private void ToggleSubTasks(Appointment task)
        {
            task.ShowSubTasks = !task.ShowSubTasks;
            UpdateTaskUIItems();
            Invalidate();
        }

        /// <summary>
        /// TaskItemUI内のサブタスク関連の矩形を更新
        /// </summary>
        private void UpdateSubTaskRects(TaskItemUI item, ref int nextY)
        {
            if (item.Task.ShowSubTasks && item.Task.SubTasks.Count > 0)
            {
                item.SubTaskRects.Clear();
                item.SubTaskCheckBoxRects.Clear();
                item.SubTaskMemoIconRects.Clear();
                
                for (int i = 0; i < item.Task.SubTasks.Count; i++)
                {
                    // サブタスク全体の矩形
                    Rectangle subTaskRect = new Rectangle(
                        item.Bounds.X + 20, // インデント
                        nextY,
                        item.Bounds.Width - 20,
                        _itemHeight - 10); // 少し小さめ
                        
                    item.SubTaskRects.Add(subTaskRect);
                    
                    // チェックボックス
                    Rectangle checkBoxRect = new Rectangle(
                        subTaskRect.X + 10,
                        subTaskRect.Y + (subTaskRect.Height - 16) / 2,
                        16,
                        16);
                        
                    item.SubTaskCheckBoxRects.Add(checkBoxRect);
                    
                    // メモアイコン
                    Rectangle memoIconRect = new Rectangle(
                        subTaskRect.X,
                        subTaskRect.Y + (subTaskRect.Height - 16) / 2,
                        16,
                        16);
                        
                    item.SubTaskMemoIconRects.Add(memoIconRect);
                    
                    nextY += subTaskRect.Height + 5;
                }
            }
        }

        #endregion

        #region Event Handlers

        private void Calendar_AppointmentAdded(object sender, EventArgs e)
        {
            LoadTasks();
        }

        private void Calendar_AppointmentChanged(object sender, EventArgs e)
        {
            LoadTasks();
        }

        private void Calendar_SelectedDateChanged(object sender, EventArgs e)
        {
            LoadTasks();
            Invalidate();
            Update(); 
        }

        private void UpdateTimer_Tick(object sender, EventArgs e)
        {
            // 表示中のタスクの状態を更新
            foreach (Appointment task in _tasks)
            {
                task.UpdateStatus();
            }
            
            // UIも更新
            UpdateTaskUIItems();
            Invalidate();
        }

        private void MetroTaskManager_MouseMove(object sender, MouseEventArgs e)
        {
            bool redraw = false;

            // 追加ボタンのホバー状態
            bool newAddButtonHover = _addButtonRect.Contains(e.Location);
            if (_addButtonHover != newAddButtonHover)
            {
                _addButtonHover = newAddButtonHover;
                redraw = true;
            }

            // タスクアイテムのホバー状態
            foreach (TaskItemUI item in _taskUIItems)
            {
                bool newHover = item.Bounds.Contains(e.Location);
                bool newEditHover = item.EditButtonRect.Contains(e.Location);
                bool newDeleteHover = item.DeleteButtonRect.Contains(e.Location);
                bool newAddHover = item.AddButtonRect.Contains(e.Location);
                
                if (item.IsHovering != newHover || 
                    item.IsEditButtonHover != newEditHover || 
                    item.IsDeleteButtonHover != newDeleteHover ||
                    item.IsAddButtonHover != newAddHover)
                {
                    item.IsHovering = newHover;
                    item.IsEditButtonHover = newEditHover;
                    item.IsDeleteButtonHover = newDeleteHover;
                    item.IsAddButtonHover = newAddHover;
                    redraw = true;
                }
            }

            // コンテキストメニューのホバー状態
            if (_showContextMenu)
            {
                bool newAddSubTaskHover = _addSubTaskMenuItemRect.Contains(e.Location);
                bool newAddMemoHover = _addMemoMenuItemRect.Contains(e.Location);
                
                if (_isAddSubTaskHover != newAddSubTaskHover || _isAddMemoHover != newAddMemoHover)
                {
                    _isAddSubTaskHover = newAddSubTaskHover;
                    _isAddMemoHover = newAddMemoHover;
                    redraw = true;
                }
            }

            // メモページの閉じるボタンのホバー状態
            if (_showMemoPage)
            {
                bool newMemoCloseButtonHover = _memoCloseButtonRect.Contains(e.Location);
                if (_isMemoCloseButtonHover != newMemoCloseButtonHover)
                {
                    _isMemoCloseButtonHover = newMemoCloseButtonHover;
                    redraw = true;
                }
            }

            if (redraw)
            {
                Invalidate();
            }
        }

        private void MetroTaskManager_MouseLeave(object sender, EventArgs e)
        {
            bool redraw = false;

            if (_addButtonHover)
            {
                _addButtonHover = false;
                redraw = true;
            }

            foreach (TaskItemUI item in _taskUIItems)
            {
                if (item.IsHovering || item.IsEditButtonHover || item.IsDeleteButtonHover)
                {
                    item.IsHovering = false;
                    item.IsEditButtonHover = false;
                    item.IsDeleteButtonHover = false;
                    redraw = true;
                }
            }

            if (redraw)
            {
                Invalidate();
            }
        }

        private void MetroTaskManager_MouseWheel(object sender, MouseEventArgs e)
        {
            // スクロール処理
            int totalHeight = _dateHeaderLabel.Height + 10 + _tasks.Count * (_itemHeight + 5) + 10;
            if (totalHeight > Height)
            {
                int newScrollPosition = _scrollPosition - (e.Delta / 120 * 20);
                newScrollPosition = Math.Max(0, Math.Min(newScrollPosition, totalHeight - Height));

                if (newScrollPosition != _scrollPosition)
                {
                    _scrollPosition = newScrollPosition;
                    UpdateTaskUIItems();
                    Invalidate();
                }
            }
        }

        private void MetroTaskManager_MouseClick(object sender, MouseEventArgs e)
        {
            // 追加ボタンのクリック
            if (!_showMemoPage)
            {
                if (_addButtonRect.Contains(e.Location))
                {
                    ShowNewTaskDialog();
                    return;
                }
            }

            // コンテキストメニュー外のクリックで閉じる
            if (_showContextMenu && !_contextMenuRect.Contains(e.Location))
            {
                _showContextMenu = false;
                Invalidate();
                return;
            }

            // コンテキストメニュー項目のクリック
            if (_showContextMenu)
            {
                if (_addSubTaskMenuItemRect.Contains(e.Location))
                {
                    AddSubTask();
                    _showContextMenu = false;
                    Invalidate();
                    return;
                }
                
                if (_addMemoMenuItemRect.Contains(e.Location))
                {
                    ShowMemoPage(_selectedTaskForMenu);
                    _showContextMenu = false;
                    Invalidate();
                    return;
                }
            }

            // メモページの閉じるボタンクリック
            if (_showMemoPage && _memoCloseButtonRect.Contains(e.Location))
            {
                CloseMemoPage();
                return;
            }

            // タスクアイテムのクリック
            foreach (TaskItemUI item in _taskUIItems)
            {
                // チェックボックスのクリック
                if (item.CheckBoxRect.Contains(e.Location))
                {
                    item.Task.IsCompleted = !item.Task.IsCompleted;
                    item.Task.UpdateStatus();
                    OnTaskChanged(EventArgs.Empty);
                    Invalidate();
                    return;
                }
                
                // 編集ボタンのクリック
                if (item.EditButtonRect.Contains(e.Location))
                {
                    EditTask(item.Task);
                    return;
                }
                
                // 削除ボタンのクリック
                if (item.DeleteButtonRect.Contains(e.Location))
                {
                    DeleteTask(item.Task);
                    return;
                }
                
                // 追加ボタンのクリック（コンテキストメニュー表示）
                if (item.AddButtonRect.Contains(e.Location))
                {
                    _selectedTaskForMenu = item.Task;
                    _showContextMenu = true;
                    
                    // コンテキストメニューの位置
                    _contextMenuRect = new Rectangle(
                        item.AddButtonRect.X - 100,
                        item.AddButtonRect.Y + item.AddButtonRect.Height,
                        120, 80);
                        
                    _addSubTaskMenuItemRect = new Rectangle(
                        _contextMenuRect.X + 10,
                        _contextMenuRect.Y + 10,
                        _contextMenuRect.Width - 20, 30);
                        
                    _addMemoMenuItemRect = new Rectangle(
                        _contextMenuRect.X + 10,
                        _contextMenuRect.Y + 45,
                        _contextMenuRect.Width - 20, 30);
                        
                    Invalidate();
                    return;
                }
                
                // メモアイコンのクリック
                if (item.MemoIconRect.Contains(e.Location))
                {
                    ShowMemoPage(item.Task);
                    return;
                }
                
                // サブタスク表示切り替えボタンのクリック
                if (item.Task.SubTasks.Count > 0 && item.ToggleSubTasksButtonRect.Contains(e.Location))
                {
                    ToggleSubTasks(item.Task);
                    return;
                }
                
                // サブタスクのチェックボックスクリック
                for (int i = 0; i < item.SubTaskCheckBoxRects.Count; i++)
                {
                    if (i < item.Task.SubTasks.Count && item.SubTaskCheckBoxRects[i].Contains(e.Location))
                    {
                        item.Task.SubTasks[i].IsCompleted = !item.Task.SubTasks[i].IsCompleted;
                        Invalidate();
                        return;
                    }
                }
                
                // サブタスクのメモアイコンクリック
                for (int i = 0; i < item.SubTaskMemoIconRects.Count; i++)
                {
                    if (i < item.Task.SubTasks.Count && item.SubTaskMemoIconRects[i].Contains(e.Location))
                    {
                        ShowMemoPage(item.Task, item.Task.SubTasks[i]);
                        return;
                    }
                }
            }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            CalculateRectangles();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_updateTimer != null)
                {
                    _updateTimer.Stop();
                    _updateTimer.Dispose();
                    _updateTimer = null;
                }
            }
            base.Dispose(disposing);
        }

        #endregion

        #region Paint Methods

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            // 背景描画
            using (SolidBrush backBrush = new SolidBrush(BackColor))
            {
                g.FillRectangle(backBrush, ClientRectangle);
            }

            // タスクアイテム描画
            foreach (TaskItemUI item in _taskUIItems)
            {
                if (item.Bounds.IntersectsWith(new Rectangle(0, 0, Width, Height)))
                {
                    DrawTaskItem(g, item);
                }
            }

            // コンテキストメニューの描画
            if (_showContextMenu)
            {
                // 背景
                using (SolidBrush menuBrush = new SolidBrush(Color.FromArgb(240, 240, 240)))
                using (Pen menuPen = new Pen(Color.FromArgb(200, 200, 200), 1))
                {
                    DrawingUtils.FillRoundedRectangle(g, menuBrush, _contextMenuRect, 5);
                    DrawingUtils.DrawRoundedRectangle(g, menuPen, _contextMenuRect, 5);
                }
                
                // メニュー項目
                // サブタスク追加項目
                using (SolidBrush itemBrush = new SolidBrush(_isAddSubTaskHover ? 
                    Color.FromArgb(220, 220, 220) : Color.FromArgb(240, 240, 240)))
                using (Font itemFont = new Font(MetroFont.FontFamily, 9f))
                using (SolidBrush textBrush = new SolidBrush(MetroColors.Text))
                {
                    DrawingUtils.FillRoundedRectangle(g, itemBrush, _addSubTaskMenuItemRect, 3);
                    g.DrawString("サブタスク", itemFont, textBrush, 
                        new Rectangle(_addSubTaskMenuItemRect.X + 5, _addSubTaskMenuItemRect.Y + 5, 
                            _addSubTaskMenuItemRect.Width - 10, _addSubTaskMenuItemRect.Height - 10));
                }
                
                // メモ追加項目
                using (SolidBrush itemBrush = new SolidBrush(_isAddMemoHover ? 
                    Color.FromArgb(220, 220, 220) : Color.FromArgb(240, 240, 240)))
                using (Font itemFont = new Font(MetroFont.FontFamily, 9f))
                using (SolidBrush textBrush = new SolidBrush(MetroColors.Text))
                {
                    DrawingUtils.FillRoundedRectangle(g, itemBrush, _addMemoMenuItemRect, 3);
                    g.DrawString("メモ", itemFont, textBrush, 
                        new Rectangle(_addMemoMenuItemRect.X + 5, _addMemoMenuItemRect.Y + 5, 
                            _addMemoMenuItemRect.Width - 10, _addMemoMenuItemRect.Height - 10));
                }
            }
            
            // メモページの描画
            if (_showMemoPage && _memoTextBox != null)
            {
                // 半透明の背景オーバーレイ
                using (SolidBrush overlayBrush = new SolidBrush(Color.FromArgb(150, 0, 0, 0)))
                {
                    g.FillRectangle(overlayBrush, ClientRectangle);
                }
                
                // メモページ背景
                using (SolidBrush pageBrush = new SolidBrush(Color.FromArgb(250, 250, 250)))
                using (Pen pagePen = new Pen(Color.FromArgb(200, 200, 200), 1))
                {
                    DrawingUtils.FillRoundedRectangle(g, pageBrush, _memoPageRect, 10);
                    DrawingUtils.DrawRoundedRectangle(g, pagePen, _memoPageRect, 10);
                }
                
                // タイトル
                string title = "メモ: ";
                if (_selectedTaskForMemo != null)
                {
                    if (_selectedSubTaskForMemo != null)
                        title += _selectedSubTaskForMemo.Title;
                    else
                        title += _selectedTaskForMemo.Title;
                }

                using (Font titleFont = new Font(MetroFont.FontFamily, 11f, FontStyle.Bold))
                using (SolidBrush textBrush = new SolidBrush(MetroColors.Primary))
                {
                    g.DrawString(title, titleFont, textBrush, 
                        new Rectangle(_memoPageRect.X + 20, _memoPageRect.Y + 15, 
                            _memoPageRect.Width - 40, 20));
                }

                // 閉じるボタン
                Color closeBtnColor = _isMemoCloseButtonHover ? MetroColors.Danger : Color.FromArgb(180, 180, 180);
                using (SolidBrush btnBrush = new SolidBrush(closeBtnColor))
                using (Pen btnPen = new Pen(Color.White, 2f))
                {
                    g.FillEllipse(btnBrush, _memoCloseButtonRect);
                    
                    // ×印
                    int x = _memoCloseButtonRect.X;
                    int y = _memoCloseButtonRect.Y;
                    int w = _memoCloseButtonRect.Width;
                    int h = _memoCloseButtonRect.Height;
                    
                    g.DrawLine(btnPen, x + w * 0.3f, y + h * 0.3f, x + w * 0.7f, y + h * 0.7f);
                    g.DrawLine(btnPen, x + w * 0.7f, y + h * 0.3f, x + w * 0.3f, y + h * 0.7f);
                }
            }

            // 追加ボタン（右下の+ボタン）
            if (!_showMemoPage)
            {
                using (SolidBrush addBtnBrush = new SolidBrush(_addButtonHover ? 
                    Color.FromArgb(0, 99, 177) : MetroColors.Primary))
                {
                    DrawingUtils.FillRoundedRectangle(g, addBtnBrush, _addButtonRect, 20);
                    
                    // プラス記号
                    using (Pen addBtnPen = new Pen(Color.White, 2f))
                    {
                        int centerX = _addButtonRect.X + _addButtonRect.Width / 2;
                        int centerY = _addButtonRect.Y + _addButtonRect.Height / 2;
                        
                        g.DrawLine(addBtnPen, 
                            centerX - 8, centerY, 
                            centerX + 8, centerY);
                        g.DrawLine(addBtnPen, 
                            centerX, centerY - 8, 
                            centerX, centerY + 8);
                    }
                }
            }
        }

        /// <summary>
        /// タスクアイテムを描画する
        /// </summary>
        private void DrawTaskItem(Graphics g, TaskItemUI item)
        {
            // 背景色の決定（状態に応じて変更）
            Color backColor;
            
            if (item.IsHovering)
            {
                backColor = Color.FromArgb(240, 240, 240);
            }
            else
            {
                switch (item.Task.Status)
                {
                    case AppointmentStatus.InProgress:
                        backColor = Color.FromArgb(230, 245, 255); // 水色（進行中）
                        break;
                    case AppointmentStatus.Overdue:
                        backColor = Color.FromArgb(255, 235, 235); // 薄赤色（期限超過）
                        break;
                    case AppointmentStatus.Completed:
                        backColor = Color.FromArgb(235, 255, 235); // 薄緑色（完了）
                        break;
                    default:
                        backColor = Color.FromArgb(250, 250, 250); // 白色（通常）
                        break;
                }
            }
            
            // 背景
            using (SolidBrush backBrush = new SolidBrush(backColor))
            {
                DrawingUtils.FillRoundedRectangle(g, backBrush, item.Bounds, 5);
            }

            // 枠線
            Color borderColor;
            switch (item.Task.Status)
            {
                case AppointmentStatus.InProgress:
                    borderColor = MetroColors.InProgress;
                    break;
                case AppointmentStatus.Overdue:
                    borderColor = MetroColors.Danger;
                    break;
                case AppointmentStatus.Completed:
                    borderColor = MetroColors.Success;
                    break;
                default:
                    borderColor = Color.FromArgb(230, 230, 230);
                    break;
            }
            
            using (Pen borderPen = new Pen(borderColor, 1.5f))
            {
                DrawingUtils.DrawRoundedRectangle(g, borderPen, item.Bounds, 5);
            }

            // チェックボックス
            using (SolidBrush boxBrush = new SolidBrush(item.Task.IsCompleted ? MetroColors.Success : Color.White))
            using (Pen boxPen = new Pen(item.Task.IsCompleted ? MetroColors.Success : MetroColors.TextLight, 1.5f))
            {
                g.FillEllipse(boxBrush, item.CheckBoxRect);
                g.DrawEllipse(boxPen, item.CheckBoxRect);

                // チェックマーク
                if (item.Task.IsCompleted)
                {
                    using (Pen checkPen = new Pen(Color.White, 2f))
                    {
                        int x = item.CheckBoxRect.X;
                        int y = item.CheckBoxRect.Y;
                        int w = item.CheckBoxRect.Width;
                        int h = item.CheckBoxRect.Height;
                        
                        g.DrawLine(checkPen, x + w * 0.3f, y + h * 0.5f, x + w * 0.45f, y + h * 0.7f);
                        g.DrawLine(checkPen, x + w * 0.45f, y + h * 0.7f, x + w * 0.7f, y + h * 0.3f);
                    }
                }
            }
            
            // タイトル
            Font titleFont = new Font(MetroFont.FontFamily, 10f, FontStyle.Bold);
            Brush titleBrush = item.Task.IsCompleted
                ? new SolidBrush(MetroColors.TextLight)
                : new SolidBrush(MetroColors.Text);

            // タイトル表示用の縮小されたRect
            Rectangle titleDisplayRect = new Rectangle(
                item.TitleRect.X,
                item.TitleRect.Y,
                Math.Min(150, item.TitleRect.Width), // 幅を150pxに制限
                item.TitleRect.Height);

            // クリッピング領域を設定
            GraphicsState state = g.Save();
            g.SetClip(titleDisplayRect);

            // スクロールオフセットを適用してタイトルを描画
            g.DrawString(item.Task.Title, titleFont, titleBrush, 
                titleDisplayRect.X - item.TitleScrollOffset, 
                titleDisplayRect.Y);

            // クリッピングを元に戻す
            g.Restore(state);

            titleFont.Dispose();
            titleBrush.Dispose();

            // 時間
            string timeText = String.Format("{0} - {1}",
                item.Task.StartTime.ToString("HH:mm"),
                item.Task.EndTime.ToString("HH:mm"));
                
            Font timeFont = new Font(MetroFont.FontFamily, 8f, FontStyle.Regular);
            Brush timeBrush = new SolidBrush(MetroColors.TextLight);
            
            g.DrawString(timeText, timeFont, timeBrush, item.TimeRect);
            timeFont.Dispose();
            timeBrush.Dispose();

            // 状態表示と通知情報
            string statusText = "";
            Color statusColor = MetroColors.TextLight;
            
            DateTime now = DateTime.Now;
            
            if (item.Task.Status == AppointmentStatus.InProgress)
            {
                statusText = "実行中";
                statusColor = MetroColors.InProgress;
            }
            else if (item.Task.Status == AppointmentStatus.Overdue)
            {
                statusText = "期限超過";
                statusColor = MetroColors.Danger;
            }
            else if (item.Task.Status == AppointmentStatus.Completed)
            {
                statusText = "完了";
                statusColor = MetroColors.Success;
            }
            else if (item.Task.NotificationMinutesBefore > 0)
            {
                // 通知設定がある場合
                TimeSpan timeToStart = item.Task.StartTime - now;
                
                if (timeToStart.TotalMinutes <= item.Task.NotificationMinutesBefore && timeToStart.TotalMinutes > 0)
                {
                    statusText = "まもなく開始";
                    statusColor = MetroColors.Warning;
                }
                else if (timeToStart.TotalMinutes > 0)
                {
                    if (item.Task.NotificationMinutesBefore >= 1440) // 24時間以上
                    {
                        statusText = String.Format("{0}日前に通知", item.Task.NotificationMinutesBefore / 1440);
                    }
                    else if (item.Task.NotificationMinutesBefore >= 60)
                    {
                        statusText = String.Format("{0}時間前に通知", item.Task.NotificationMinutesBefore / 60);
                    }
                    else
                    {
                        statusText = String.Format("{0}分前に通知", item.Task.NotificationMinutesBefore);
                    }
                    statusColor = MetroColors.Info;
                }
            }
            
            if (!string.IsNullOrEmpty(statusText))
            {
                Font statusFont = new Font(MetroFont.FontFamily, 8f, FontStyle.Regular);
                Brush statusBrush = new SolidBrush(statusColor);
                
                g.DrawString(statusText, statusFont, statusBrush, item.NotifyRect);
                statusFont.Dispose();
                statusBrush.Dispose();
            }
            
            // 編集ボタン
            Color editBtnColor = item.IsEditButtonHover ? MetroColors.Secondary : MetroColors.Primary;
            using (SolidBrush editBtnBrush = new SolidBrush(editBtnColor))
            {
                DrawingUtils.FillRoundedRectangle(g, editBtnBrush, item.EditButtonRect, 3);
            }

            // 編集アイコン（鉛筆）- ペン描画を絵文字に置き換え
            using (Font emojiFont = new Font("Segoe UI Emoji", 6f))
            using (SolidBrush textBrush = new SolidBrush(Color.White))
            {
                g.DrawString("🖊", emojiFont, textBrush, 
                    item.EditButtonRect.X + 6, 
                    item.EditButtonRect.Y + 6);
            }
            
            // 削除ボタン
            Color deleteBtnColor = item.IsDeleteButtonHover ? MetroColors.Danger : Color.FromArgb(220, 220, 220);
            using (SolidBrush deleteBtnBrush = new SolidBrush(deleteBtnColor))
            {
                DrawingUtils.FillRoundedRectangle(g, deleteBtnBrush, item.DeleteButtonRect, 3);
            }
            
            // 削除アイコン（×）
            using (Pen iconPen = new Pen(item.IsDeleteButtonHover ? Color.White : Color.DimGray, 1.5f))
            {
                Rectangle iconRect = item.DeleteButtonRect;
                int centerX = iconRect.X + iconRect.Width / 2;
                int centerY = iconRect.Y + iconRect.Height / 2;
                
                g.DrawLine(iconPen, 
                    centerX - 3, centerY - 3, 
                    centerX + 3, centerY + 3);
                g.DrawLine(iconPen, 
                    centerX + 3, centerY - 3, 
                    centerX - 3, centerY + 3);
            }
            
            Color addBtnColor = item.IsAddButtonHover ? MetroColors.Secondary : MetroColors.Info;
            using (SolidBrush addBtnBrush = new SolidBrush(addBtnColor))
            {
                DrawingUtils.FillRoundedRectangle(g, addBtnBrush, item.AddButtonRect, 3);
            }
            
            // 追加ボタンのプラス記号
            using (Pen iconPen = new Pen(Color.White, 1.5f))
            {
                int centerX = item.AddButtonRect.X + item.AddButtonRect.Width / 2;
                int centerY = item.AddButtonRect.Y + item.AddButtonRect.Height / 2;
                
                g.DrawLine(iconPen, centerX - 3, centerY, centerX + 3, centerY);
                g.DrawLine(iconPen, centerX, centerY - 3, centerX, centerY + 3);
            }
            
            // サブタスク表示トグルボタン（サブタスクがある場合のみ）
            if (item.Task.SubTasks.Count > 0)
            {
                using (SolidBrush toggleBrush = new SolidBrush(MetroColors.TextLight))
                using (Pen togglePen = new Pen(MetroColors.TextLight, 1.5f))
                {
                    int x = item.ToggleSubTasksButtonRect.X;
                    int y = item.ToggleSubTasksButtonRect.Y;
                    int w = item.ToggleSubTasksButtonRect.Width;
                    int h = item.ToggleSubTasksButtonRect.Height;
                    
                    // 三角形（上向き/下向き）を描画
                    if (item.Task.ShowSubTasks)
                    {
                        // 下向き三角形（表示中）
                        Point[] points = {
                            new Point(x + w/2, y + h*2/3),
                            new Point(x + w/4, y + h/3),
                            new Point(x + w*3/4, y + h/3)
                        };
                        g.FillPolygon(toggleBrush, points);
                    }
                    else
                    {
                        // 右向き三角形（非表示）
                        Point[] points = {
                            new Point(x + w*2/3, y + h/2),
                            new Point(x + w/3, y + h/4),
                            new Point(x + w/3, y + h*3/4)
                        };
                        g.FillPolygon(toggleBrush, points);
                    }
                }
            }
            
            // メモアイコン（メモがある場合のみ）
            if (item.Task.HasMemo)
            {
                // クリップアイコン（📄）を描画 - この部分を修正
                using (Font emojiFont = new Font("Segoe UI Emoji", 10f))
                using (SolidBrush textBrush = new SolidBrush(MetroColors.Info))
                {
                    int x = item.MemoIconRect.X;
                    int y = item.MemoIconRect.Y;
                    g.DrawString("📄", emojiFont, textBrush, x, y);
                }
            }
            
            // サブタスク描画（表示設定がオンの場合のみ）
            if (item.Task.ShowSubTasks && item.Task.SubTasks.Count > 0)
            {
                for (int i = 0; i < item.SubTaskRects.Count; i++)
                {
                    if (i >= item.Task.SubTasks.Count)
                        break;
                        
                    SubTask subTask = item.Task.SubTasks[i];
                    Rectangle subTaskRect = item.SubTaskRects[i];
                    
                    // サブタスク背景
                    Color subTaskBackColor = Color.FromArgb(240, 240, 240);
                    using (SolidBrush backBrush = new SolidBrush(subTaskBackColor))
                    {
                        DrawingUtils.FillRoundedRectangle(g, backBrush, subTaskRect, 4);
                    }
                    
                    // サブタスクのタイトル
                    using (Font subTaskFont = new Font(MetroFont.FontFamily, 9f, 
                        subTask.IsCompleted ? FontStyle.Strikeout : FontStyle.Regular))
                    using (SolidBrush textBrush = new SolidBrush(subTask.IsCompleted ? 
                        MetroColors.TextLight : MetroColors.Text))
                    {
                        Rectangle titleRect = new Rectangle(
                            item.SubTaskCheckBoxRects[i].Right + 5,
                            subTaskRect.Y + (subTaskRect.Height - 16) / 2,
                            subTaskRect.Width - item.SubTaskCheckBoxRects[i].Width - 30,
                            16);
                            
                        g.DrawString(subTask.Title, subTaskFont, textBrush, titleRect);
                    }
                    
                    // サブタスクのチェックボックス
                    using (SolidBrush boxBrush = new SolidBrush(subTask.IsCompleted ? 
                        MetroColors.Success : Color.White))
                    using (Pen boxPen = new Pen(subTask.IsCompleted ? 
                        MetroColors.Success : MetroColors.TextLight, 1.5f))
                    {
                        g.FillEllipse(boxBrush, item.SubTaskCheckBoxRects[i]);
                        g.DrawEllipse(boxPen, item.SubTaskCheckBoxRects[i]);
                        
                        // チェックマーク
                        if (subTask.IsCompleted)
                        {
                            using (Pen checkPen = new Pen(Color.White, 1.5f))
                            {
                                Rectangle r = item.SubTaskCheckBoxRects[i];
                                g.DrawLine(checkPen, r.X + r.Width * 0.3f, r.Y + r.Height * 0.5f, 
                                    r.X + r.Width * 0.45f, r.Y + r.Height * 0.7f);
                                g.DrawLine(checkPen, r.X + r.Width * 0.45f, r.Y + r.Height * 0.7f, 
                                    r.X + r.Width * 0.7f, r.Y + r.Height * 0.3f);
                            }
                        }
                    }
                    
                    // サブタスクのメモアイコン（メモがある場合のみ）
                    if (!string.IsNullOrEmpty(subTask.Memo))
                    {
                        // クリップアイコン（📄）を描画 - この部分を修正
                        using (Font emojiFont = new Font("Segoe UI Emoji", 10f))
                        using (SolidBrush textBrush = new SolidBrush(MetroColors.Info))
                        {
                            Rectangle r = item.SubTaskMemoIconRects[i];
                            g.DrawString("📄", emojiFont, textBrush, r.X, r.Y);
                        }
                    }
                }
            }
        }

        #endregion

        #region Helper Classes

        /// <summary>
        /// タスクアイテムのUI情報
        /// </summary>
        internal class TaskItemUI
        {
            public Appointment Task { get; set; }
            public Rectangle Bounds { get; set; }
            public Rectangle CheckBoxRect { get; set; }
            public Rectangle TitleRect { get; set; }
            public Rectangle TimeRect { get; set; }
            public Rectangle NotifyRect { get; set; }
            public Rectangle EditButtonRect { get; set; }
            public Rectangle DeleteButtonRect { get; set; }
            public bool IsHovering { get; set; }
            public bool IsEditButtonHover { get; set; }
            public bool IsDeleteButtonHover { get; set; }
            
            public Rectangle AddButtonRect { get; set; }
            public bool IsAddButtonHover { get; set; }
            private List<Rectangle> _subTaskRects;
            private List<Rectangle> _subTaskCheckBoxRects;
            private List<Rectangle> _subTaskMemoIconRects;
            public Rectangle ToggleSubTasksButtonRect { get; set; }
            public Rectangle MemoIconRect { get; set; }
            
            public int TitleScrollOffset { get; set; } // タイトルのスクロールオフセット
            public bool IsTitleScrolling { get; set; } // スクロール中かどうか
            public DateTime TitleScrollStartTime { get; set; } // スクロール開始時間
            public int ScrollSpeed { get; set; } // スクロール速度（個別）
            public int ScrollDelay { get; set; } // スクロール開始前の待機時間（個別）
            public int ScrollCounter { get; set; } // スクロールカウンタ（個別）
            
            public List<Rectangle> SubTaskRects 
            { 
                get { return _subTaskRects; } 
                set { _subTaskRects = value; }
            }

            public List<Rectangle> SubTaskCheckBoxRects 
            { 
                get { return _subTaskCheckBoxRects; } 
                set { _subTaskCheckBoxRects = value; }
            }

            public List<Rectangle> SubTaskMemoIconRects 
            { 
                get { return _subTaskMemoIconRects; } 
                set { _subTaskMemoIconRects = value; }
            }

            public TaskItemUI(Appointment task, Rectangle bounds)
            {
                Task = task;
                Bounds = bounds;
                
                // リストの初期化
                _subTaskRects = new List<Rectangle>();
                _subTaskCheckBoxRects = new List<Rectangle>();
                _subTaskMemoIconRects = new List<Rectangle>();
                
                // タスク固有のスクロールパラメータを設定（ランダム化）
                // タスクのタイトルとIDを組み合わせてシードを生成
                int seed = task.Title.GetHashCode() + task.StartTime.GetHashCode();
                Random rand = new Random(seed);
                ScrollSpeed = rand.Next(1, 4); // 1～3のスクロール速度
                ScrollDelay = rand.Next(40, 120); // 2～6秒間（50msタイマーで40～120回分）
                ScrollCounter = 0;
                TitleScrollOffset = 0;
                IsTitleScrolling = false;
                TitleScrollStartTime = DateTime.MinValue;
                
                // 各要素の矩形を計算
                int checkBoxSize = 20;
                
                // サブタスク表示切り替えボタン - チェックボックスの左に移動
                ToggleSubTasksButtonRect = new Rectangle(
                    bounds.X + 4, // 左端に配置
                    bounds.Y + (bounds.Height - 16) / 2,
                    16,
                    16);
                
                // チェックボックス - 16px右に移動
                CheckBoxRect = new Rectangle(
                    bounds.X + 10 + 16, // トグルボタンの右側に間隔をあけて配置
                    bounds.Y + (bounds.Height - checkBoxSize) / 2,
                    checkBoxSize,
                    checkBoxSize);
                
                // 以下の要素も連動して調整
                int buttonSize = 24;
                
                // 編集ボタン
                EditButtonRect = new Rectangle(
                    bounds.Right - buttonSize - 10,
                    bounds.Y + (bounds.Height - buttonSize) / 2,
                    buttonSize,
                    buttonSize);
                
                // 削除ボタン
                DeleteButtonRect = new Rectangle(
                    EditButtonRect.X - buttonSize - 5,
                    bounds.Y + (bounds.Height - buttonSize) / 2,
                    buttonSize,
                    buttonSize);
                
                // 追加ボタン
                AddButtonRect = new Rectangle(
                    DeleteButtonRect.X - buttonSize - 5,
                    bounds.Y + (bounds.Height - buttonSize) / 2,
                    buttonSize,
                    buttonSize);
                
                // コントロールボタンの分だけ内容表示領域を調整
                int contentRight = DeleteButtonRect.X - 10;
                
                TitleRect = new Rectangle(
                    bounds.X + CheckBoxRect.Right + 10,
                    bounds.Y + 10,
                    contentRight - (bounds.X + CheckBoxRect.Right + 10),
                    20);
                
                TimeRect = new Rectangle(
                    TitleRect.X,
                    TitleRect.Bottom + 2,
                    TitleRect.Width / 2,
                    18);
                
                NotifyRect = new Rectangle(
                    TimeRect.Right + 5,
                    TimeRect.Y,
                    TitleRect.Width / 2 - 5,
                    18);
                
                // メモアイコン
                MemoIconRect = new Rectangle(
                    TitleRect.Right - 40,
                    TitleRect.Y + 2,
                    16,
                    16);
            }
        }

        #endregion
    }

    #endregion

    #region Sample Form

    /// <summary>
    /// サンプルフォーム
    /// </summary>
    public class MetroUIForm : Form
    {
        private MetroCalendar calendar;
        private MetroDigitalClock clock;
        private MetroTaskManager taskManager;
        private System.Windows.Forms.Timer statusUpdateTimer;

        public MetroUIForm()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            try
            {
                this.Text = "Metro UI Sample";
                this.Size = new Size(800, 600);
                this.BackColor = Color.WhiteSmoke;
                this.Font = new Font("Segoe UI", 9F);

                // カレンダー
                calendar = new MetroCalendar();
                calendar.Location = new Point(20, 20);
                calendar.Size = new Size(350, 300);
                this.Controls.Add(calendar);

                // デジタル時計
                clock = new MetroDigitalClock();
                clock.Location = new Point(20, 340);
                clock.Size = new Size(350, 220);
                clock.Calendar = calendar;
                this.Controls.Add(clock);

                // タスク管理
                taskManager = new MetroTaskManager();
                taskManager.Location = new Point(400, 20);
                taskManager.Size = new Size(360, 540);
                taskManager.Calendar = calendar;
                this.Controls.Add(taskManager);

                // ここで明示的にアプリケーションを初期化
                Application.DoEvents();

                // サンプルデータの作成
                AddSampleAppointments();

                // 状態更新タイマー
                statusUpdateTimer = new System.Windows.Forms.Timer();
                statusUpdateTimer.Interval = 60000; // 1分ごとに更新
                statusUpdateTimer.Tick += (sender, e) => 
                {
                    if (calendar != null)
                        calendar.UpdateAppointmentStatuses();
                };
                statusUpdateTimer.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show("初期化エラー: " + ex.Message + "\n" + ex.StackTrace, "エラー", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddSampleAppointments()
        {
            // 今日の予定
            Appointment today1 = new Appointment();
            today1.Title = "朝会";
            today1.StartTime = DateTime.Today.AddHours(9);
            today1.EndTime = DateTime.Today.AddHours(10);
            calendar.AddAppointment(today1);

            Appointment today2 = new Appointment();
            today2.Title = "プロジェクトミーティング";
            today2.StartTime = DateTime.Today.AddHours(13);
            today2.EndTime = DateTime.Today.AddHours(14);
            today2.NotificationMinutesBefore = 30;
            calendar.AddAppointment(today2);

            // 明日の予定
            Appointment tomorrow = new Appointment();
            tomorrow.Title = "クライアントミーティング";
            tomorrow.StartTime = DateTime.Today.AddDays(1).AddHours(15);
            tomorrow.EndTime = DateTime.Today.AddDays(1).AddHours(16);
            tomorrow.NotificationMinutesBefore = 60;
            calendar.AddAppointment(tomorrow);

            // 次週の予定
            Appointment nextWeek = new Appointment();
            nextWeek.Title = "プロジェクト納期";
            nextWeek.StartTime = DateTime.Today.AddDays(7).AddHours(18);
            nextWeek.EndTime = DateTime.Today.AddDays(7).AddHours(19);
            nextWeek.NotificationMinutesBefore = 1440; // 1日前
            calendar.AddAppointment(nextWeek);

            // 翌月の予定
            Appointment nextMonth = new Appointment();
            nextMonth.Title = "四半期レビュー";
            nextMonth.StartTime = DateTime.Today.AddDays(28).AddHours(10);
            nextMonth.EndTime = DateTime.Today.AddDays(28).AddHours(12);
            nextMonth.NotificationMinutesBefore = 120; // 2時間前
            calendar.AddAppointment(nextMonth);

            // 状態の更新
            calendar.UpdateAppointmentStatuses();
        }
    }

    #endregion
}

#endregion

#region DockableFormWithMetroUI

/// <summary>
/// DockableFormにMetroUIを組み込んだメインアプリケーションクラス
/// </summary>
public class DockableFormWithMetroUI
{
    // DockableFormインスタンス
    private static AnimatedDockableForm _dockForm;
    
    // MetroUIのコンポーネント
    private static MetroUI.MetroCalendar _calendar;
    private static MetroUI.MetroDigitalClock _clock;
    private static MetroUI.MetroTaskManager _taskManager;
    private static Panel _container;
    
    // メインエントリーポイント
    [STAThread]
    public static void Main()
    {
        try
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            // DockableFormの作成
            _dockForm = new AnimatedDockableForm();
            _dockForm.Text = "MetroUI Dashboard";
            _dockForm.Size = new Size(900, 600);
            
            // コンテンツパネルの設定
            _container = new Panel();
            _container.Dock = DockStyle.Fill;
            _container.BackColor = Color.WhiteSmoke;
            _dockForm.ContentPanel.Controls.Add(_container);
            
            // MetroUIコンポーネントの初期化
            InitializeMetroComponents();
            
            // サンプルデータの追加
            AddSampleAppointments();
            
            // フォームの表示
            _dockForm.PinMode = PinMode.None;
            
            Application.Run(_dockForm);
        }
        catch (Exception ex)
        {
            MessageBox.Show("アプリケーションエラー: " + ex.Message + "\n" + ex.StackTrace, "エラー", 
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
    
    private static void InitializeMetroComponents()
    {
        // カレンダー
        _calendar = new MetroUI.MetroCalendar();
        _calendar.Location = new Point(20, 20);
        _calendar.Size = new Size(350, 300);
        _container.Controls.Add(_calendar);

        // デジタル時計
        _clock = new MetroUI.MetroDigitalClock();
        _clock.Location = new Point(20, 340);
        _clock.Size = new Size(350, 220);
        _clock.Calendar = _calendar;
        _container.Controls.Add(_clock);

        // タスク管理
        _taskManager = new MetroUI.MetroTaskManager();
        _taskManager.Location = new Point(400, 20);
        _taskManager.Size = new Size(360, 540);
        _taskManager.Calendar = _calendar;
        _container.Controls.Add(_taskManager);
    }
    
    private static void AddSampleAppointments()
    {
        // 今日の予定
        MetroUI.Appointment today1 = new MetroUI.Appointment();
        today1.Title = "朝会";
        today1.StartTime = DateTime.Today.AddHours(9);
        today1.EndTime = DateTime.Today.AddHours(10);
        _calendar.AddAppointment(today1);

        MetroUI.Appointment today2 = new MetroUI.Appointment();
        today2.Title = "プロジェクトミーティング";
        today2.StartTime = DateTime.Today.AddHours(13);
        today2.EndTime = DateTime.Today.AddHours(14);
        today2.NotificationMinutesBefore = 30;
        _calendar.AddAppointment(today2);

        // 明日の予定
        MetroUI.Appointment tomorrow = new MetroUI.Appointment();
        tomorrow.Title = "クライアントミーティング";
        tomorrow.StartTime = DateTime.Today.AddDays(1).AddHours(15);
        tomorrow.EndTime = DateTime.Today.AddDays(1).AddHours(16);
        tomorrow.NotificationMinutesBefore = 60;
        _calendar.AddAppointment(tomorrow);

        // 次週の予定
        MetroUI.Appointment nextWeek = new MetroUI.Appointment();
        nextWeek.Title = "プロジェクト納期";
        nextWeek.StartTime = DateTime.Today.AddDays(7).AddHours(18);
        nextWeek.EndTime = DateTime.Today.AddDays(7).AddHours(19);
        nextWeek.NotificationMinutesBefore = 1440; // 1日前
        _calendar.AddAppointment(nextWeek);

        // 翌月の予定
        MetroUI.Appointment nextMonth = new MetroUI.Appointment();
        nextMonth.Title = "四半期レビュー";
        nextMonth.StartTime = DateTime.Today.AddDays(28).AddHours(10);
        nextMonth.EndTime = DateTime.Today.AddDays(28).AddHours(12);
        nextMonth.NotificationMinutesBefore = 120; // 2時間前
        _calendar.AddAppointment(nextMonth);

        // 状態の更新
        _calendar.UpdateAppointmentStatuses();
    }
}

#endregion
"@

# C#コードをコンパイルして実行
Add-Type -TypeDefinition $csCode -ReferencedAssemblies System.Windows.Forms, System.Drawing, System.ComponentModel, System.Core, System.Drawing.Design

# 統合されたアプリケーションを起動
[DockableFormWithMetroUI]::Main()