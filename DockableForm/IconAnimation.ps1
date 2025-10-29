[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') | Out-Null

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

$originalSize  = [System.Drawing.Size]::new(420, 280)
$collapsedSize = [System.Drawing.Size]::new(24, 24)
$headerHeight  = [Math]::Max([System.Windows.Forms.SystemInformation]::CaptionHeight + 8, 36)

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Dockable Form Icon Animation'
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
$form.StartPosition = 'Manual'
$form.ShowInTaskbar = $false
$form.TopMost = $false
$form.BackColor = [System.Drawing.Color]::White
$form.Opacity = 0.97
$form.Size = $originalSize
$form.Location = [System.Drawing.Point]::new(320, 180)

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = 'Top'
$headerPanel.Height = $headerHeight
$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(255, 45, 45, 70)
$headerPanel.ForeColor = [System.Drawing.Color]::White
$headerPanel.Padding = [System.Windows.Forms.Padding]::new(12, 8, 12, 8)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = 'DockableForm / Icon Hover Expansion'
$titleLabel.AutoSize = $true
$titleLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 11, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::White

$headerPanel.Controls.Add($titleLabel)

$contentPanel = New-Object System.Windows.Forms.Panel
$contentPanel.Dock = 'Fill'
$contentPanel.Padding = [System.Windows.Forms.Padding]::new(18)
$contentPanel.BackColor = [System.Drawing.Color]::FromArgb(255, 248, 248, 252)

$bodyLabel = New-Object System.Windows.Forms.Label
$bodyLabel.Text = @'
ホバーで正円が記憶された幅へと広がり、続けてフォーム本体が下方向にドローオープンします。
マウスカーソルが離れると、逆順のアニメーションで24x24の半透明な紫色アイコンへ折りたたまれます。
'@
$bodyLabel.AutoSize = $true
$bodyLabel.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$contentPanel.Controls.Add($bodyLabel)

$form.Controls.Add($contentPanel)
$form.Controls.Add($headerPanel)

$script:form = $form
$script:originalSize = $originalSize
$script:collapsedSize = $collapsedSize
$script:headerHeight = $headerHeight
$script:headerPanel = $headerPanel
$script:contentPanel = $contentPanel
$script:state = 'Collapsed'
$script:isAnimating = $false
$script:currentAnimation = $null
$script:anchorCenterX = $form.Location.X + [int]($originalSize.Width / 2)
$script:anchorTop = $form.Location.Y

$script:animationTimer = New-Object System.Windows.Forms.Timer
$script:animationTimer.Interval = 16

$script:hoverOutTimer = New-Object System.Windows.Forms.Timer
$script:hoverOutTimer.Interval = 180

function SetEllipseRegion {
    param(
        [int]$width,
        [int]$height
    )

    if ($width -lt 1 -or $height -lt 1) { return }

    if ($script:form.Region -ne $null) {
        $script:form.Region.Dispose()
        $script:form.Region = $null
    }

    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddEllipse(0, 0, $width, $height)
    $script:form.Region = New-Object System.Drawing.Region($path)
    $path.Dispose()
}

function SetCircleRegion {
    param([int]$diameter)

    SetEllipseRegion -width $diameter -height $diameter
}

function ClearRegion {
    if ($script:form.Region -ne $null) {
        $script:form.Region.Dispose()
        $script:form.Region = $null
    }
}

function UpdateFormBounds {
    param(
        [int]$width,
        [int]$height
    )

    $x = [int][Math]::Round($script:anchorCenterX - ($width / 2.0))

    $script:form.SuspendLayout()
    $script:form.Location = [System.Drawing.Point]::new($x, $script:anchorTop)
    $script:form.Size = [System.Drawing.Size]::new($width, $height)
    $script:form.ResumeLayout()

    $needsEllipse = $false

    switch ($script:state) {
        'Collapsed' { $needsEllipse = $true }
        'ExpandingWidth' { $needsEllipse = $true }
        'CollapsingIcon' { $needsEllipse = $true }
    }

    if ($needsEllipse) {
        SetEllipseRegion -width $width -height $height
    }
    elseif ($script:form.Region -ne $null) {
        ClearRegion
    }
}

function SetFormBoundsInstant {
    param(
        [int]$width,
        [int]$height
    )

    UpdateFormBounds -width $width -height $height
}

function SetCollapsedVisual {
    $script:headerPanel.Visible = $false
    $script:contentPanel.Visible = $false
    $script:form.BackColor = [System.Drawing.Color]::FromArgb(190, 120, 70, 190)
    $script:form.Opacity = 0.85
    SetCircleRegion $script:collapsedSize.Width
}

function SetHeaderVisual {
    $script:form.BackColor = [System.Drawing.Color]::FromArgb(255, 245, 243, 252)
    $script:form.Opacity = 0.92
    $script:headerPanel.Visible = $true
}

function SetExpandedVisual {
    $script:form.BackColor = [System.Drawing.Color]::White
    $script:form.Opacity = 0.97
    $script:headerPanel.Visible = $true
    $script:contentPanel.Visible = $true
}

function StopCurrentAnimation {
    if ($script:animationTimer.Enabled) {
        $script:animationTimer.Stop()
    }

    $script:currentAnimation = $null
    $script:isAnimating = $false
}

function StartSizeAnimation {
    param(
        [int]$targetWidth,
        [int]$targetHeight,
        [int]$duration,
        [ScriptBlock]$completed
    )

    $startWidth = $script:form.Width
    $startHeight = $script:form.Height

    if ($startWidth -eq $targetWidth -and $startHeight -eq $targetHeight) {
        if ($completed) { & $completed }
        return
    }

    $steps = [Math]::Max([Math]::Ceiling($duration / $script:animationTimer.Interval), 1)

    $script:currentAnimation = [ordered]@{
        Step = 0
        Steps = $steps
        StartWidth = $startWidth
        StartHeight = $startHeight
        TargetWidth = $targetWidth
        TargetHeight = $targetHeight
        Completed = $completed
    }

    $script:isAnimating = $true

    if (-not $script:animationTimer.Enabled) {
        $script:animationTimer.Start()
    }
}

$script:animationTimer.add_Tick({
    $state = $script:currentAnimation
    if ($null -eq $state) {
        $script:animationTimer.Stop()
        $script:isAnimating = $false
        return
    }

    $state.Step++
    $t = [Math]::Min(1.0, [double]$state.Step / [double]$state.Steps)
    $ease = $t * $t * (3 - 2 * $t)   # Smoothstep easing

    $width = [int][Math]::Round($state.StartWidth + ($state.TargetWidth - $state.StartWidth) * $ease)
    $height = [int][Math]::Round($state.StartHeight + ($state.TargetHeight - $state.StartHeight) * $ease)

    UpdateFormBounds -width $width -height $height

    if ($state.Step -ge $state.Steps) {
        UpdateFormBounds -width $state.TargetWidth -height $state.TargetHeight
        $completed = $state.Completed
        $script:currentAnimation = $null
        $script:animationTimer.Stop()
        $script:isAnimating = $false

        if ($completed) {
            & $completed
        }
    }
})

function CollapseInstant {
    SetFormBoundsInstant -width $script:collapsedSize.Width -height $script:collapsedSize.Height
    SetCollapsedVisual
    $script:state = 'Collapsed'
}

function StartCircleCollapse {
    $script:state = 'CollapsingIcon'
    StartSizeAnimation -targetWidth $script:collapsedSize.Width -targetHeight $script:collapsedSize.Height -duration 220 -completed {
        CollapseInstant
    }
}

function StartCollapse {
    if (TestCursorInsideForm) { return }

    switch ($script:state) {
        'Collapsed' { return }
        'CollapsingBody' { return }
        'CollapsingIcon' { return }
    }

    StopCurrentAnimation

    $script:state = 'CollapsingBody'
    $script:contentPanel.Visible = $false

    $targetHeight = [Math]::Max($script:headerHeight, $script:collapsedSize.Height)

    StartSizeAnimation -targetWidth $script:form.Width -targetHeight $targetHeight -duration 160 -completed {
        $script:headerPanel.Visible = $false
        StartCircleCollapse
    }
}

function StartExpand {
    switch ($script:state) {
        'Expanded' { return }
        'ExpandingWidth' { return }
        'ExpandingHeight' { return }
    }

    StopCurrentAnimation

    $script:state = 'ExpandingWidth'
    ClearRegion
    SetHeaderVisual

    $targetWidth = $script:originalSize.Width
    $targetHeight = [Math]::Max($script:headerHeight, $script:collapsedSize.Height)

    StartSizeAnimation -targetWidth $targetWidth -targetHeight $targetHeight -duration 220 -completed {
        $script:state = 'ExpandingHeight'
        $script:contentPanel.Visible = $true

        StartSizeAnimation -targetWidth $script:originalSize.Width -targetHeight $script:originalSize.Height -duration 260 -completed {
            $script:state = 'Expanded'
            SetExpandedVisual
            SetFormBoundsInstant -width $script:originalSize.Width -height $script:originalSize.Height
        }
    }
}

function TestCursorInsideForm {
    $cursor = [System.Windows.Forms.Control]::MousePosition
    return $script:form.Bounds.Contains($cursor)
}

function HandleMouseEnter {
    if ($script:hoverOutTimer.Enabled) {
        $script:hoverOutTimer.Stop()
    }

    StartExpand
}

function HandleMouseLeave {
    if (-not $script:hoverOutTimer.Enabled) {
        $script:hoverOutTimer.Start()
    }
}

function RegisterControlForHover {
    param([System.Windows.Forms.Control]$control)

    $control.add_MouseEnter({ param($sender, $args) HandleMouseEnter })
    $control.add_MouseLeave({ param($sender, $args) HandleMouseLeave })

    foreach ($child in $control.Controls) {
        RegisterControlForHover -control $child
    }
}

$script:hoverOutTimer.add_Tick({
    if (TestCursorInsideForm) { return }

    $script:hoverOutTimer.Stop()
    StartCollapse
})

$script:dragging = $false
$script:dragOffset = [System.Drawing.Point]::Empty

$headerPanel.add_MouseDown({
    param($sender, $args)
    if ($args.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
        $script:dragging = $true
        $screenPoint = $headerPanel.PointToScreen([System.Drawing.Point]::new($args.X, $args.Y))
        $script:dragOffset = [System.Drawing.Point]::new(
            $screenPoint.X - $script:form.Location.X,
            $screenPoint.Y - $script:form.Location.Y
        )
    }
})

$headerPanel.add_MouseMove({
    param($sender, $args)
    if ($script:dragging) {
        $cursor = [System.Windows.Forms.Control]::MousePosition
        $newX = $cursor.X - $script:dragOffset.X
        $newY = $cursor.Y - $script:dragOffset.Y
        $script:form.Location = [System.Drawing.Point]::new($newX, $newY)
    }
})

$headerPanel.add_MouseUp({
    param($sender, $args)
    if ($args.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
        $script:dragging = $false
    }
})

$form.add_MouseMove({
    if ($script:dragging) {
        $cursor = [System.Windows.Forms.Control]::MousePosition
        $newX = $cursor.X - $script:dragOffset.X
        $newY = $cursor.Y - $script:dragOffset.Y
        $script:form.Location = [System.Drawing.Point]::new($newX, $newY)
    }
})

$form.add_MouseUp({
    param($sender, $args)
    if ($args.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
        $script:dragging = $false
    }
})

$form.add_LocationChanged({
    if ($script:state -eq 'Expanded' -and -not $script:isAnimating) {
        $script:anchorCenterX = $script:form.Location.X + [int]($script:form.Width / 2)
        $script:anchorTop = $script:form.Location.Y
    }
})

RegisterControlForHover -control $form

CollapseInstant

[System.Windows.Forms.Application]::Run($form)
