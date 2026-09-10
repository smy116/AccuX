# Generates AccuX Ribbon icons: 32x32 PNG, unified visual language
# (dark blue rounded background + white symbol).
# Compact icon set for normal-size Ribbon buttons.
param(
    [string]$OutputDirectory = "$PSScriptRoot\..\src\AccuX.AddIn\Resources"
)

Add-Type -AssemblyName System.Drawing

if (-not (Test-Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

function New-AccuXIcon {
    param(
        [string]$Path,
        [System.Drawing.Color]$AccentColor,
        [scriptblock]$DrawSymbol
    )

    $size = 32
    # WPS 通过 IPictureDisp 显示图标时不会稳定保留 PNG Alpha，
    # 透明像素会被垫成灰色方块。使用 24bpp 不透明画布避免灰边。
    $bitmap = New-Object System.Drawing.Bitmap($size, $size, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $graphics.Clear([System.Drawing.Color]::White)

    # Fill the entire canvas so WPS does not show a gray/white matte around
    # transparent or rounded corners in the Ribbon.
    $background = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 20, 49, 70))
    $graphics.FillRectangle($background, 0, 0, $size, $size)

    $white = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 246, 250, 252), 2.5)
    $whiteBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 246, 250, 252))
    $accent = New-Object System.Drawing.Pen($AccentColor, 2.5)
    $accentBrush = New-Object System.Drawing.SolidBrush($AccentColor)
    $white.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $white.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $accent.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $accent.EndCap = [System.Drawing.Drawing2D.LineCap]::Round

    & $DrawSymbol $graphics $white $whiteBrush $accent $accentBrush

    $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)

    $accentBrush.Dispose()
    $accent.Dispose()
    $whiteBrush.Dispose()
    $white.Dispose()
    $background.Dispose()
    $graphics.Dispose()
    $bitmap.Dispose()
}

# Rounding: precision dial with a highlighted rounding arrow.
New-AccuXIcon -Path (Join-Path $OutputDirectory 'round.png') -AccentColor ([System.Drawing.Color]::FromArgb(255, 255, 193, 87)) -DrawSymbol {
    param($g, $pen, $brush, $accent, $accentBrush)
    $g.DrawArc($accent, 7, 7, 18, 18, 35, 275)
    $g.DrawLine($accent, 24, 8, 24, 13)
    $g.DrawLine($accent, 24, 8, 19, 9)
    $g.DrawLine($pen, 11, 14, 15, 14)
    $g.FillEllipse($accentBrush, 16.5, 12.5, 3, 3)
    $g.DrawLine($pen, 20, 14, 22, 14)
    $g.DrawLine($pen, 11, 19, 22, 19)
}

# Amount conversion: two clean, opposing transfer arrows.
New-AccuXIcon -Path (Join-Path $OutputDirectory 'convert.png') -AccentColor ([System.Drawing.Color]::FromArgb(255, 99, 211, 238)) -DrawSymbol {
    param($g, $pen, $brush, $accent, $accentBrush)
    $g.DrawLine($accent, 8, 11, 23, 11)
    $g.DrawLine($accent, 18, 7, 23, 11)
    $g.DrawLine($accent, 18, 15, 23, 11)
    $g.DrawLine($pen, 24, 21, 9, 21)
    $g.DrawLine($pen, 14, 17, 9, 21)
    $g.DrawLine($pen, 14, 25, 9, 21)
}

# Selection sum: a compact sigma symbol for aggregation.
New-AccuXIcon -Path (Join-Path $OutputDirectory 'sum.png') -AccentColor ([System.Drawing.Color]::FromArgb(255, 255, 193, 87)) -DrawSymbol {
    param($g, $pen, $brush, $accent, $accentBrush)
    $g.DrawLine($accent, 8, 8, 24, 8)
    $g.DrawLine($accent, 8, 8, 21, 16)
    $g.DrawLine($accent, 21, 16, 8, 24)
    $g.DrawLine($accent, 8, 24, 24, 24)
}

# Uppercase: a bold Chinese "大" mark with a small seal accent.
New-AccuXIcon -Path (Join-Path $OutputDirectory 'uppercase.png') -AccentColor ([System.Drawing.Color]::FromArgb(255, 166, 230, 183)) -DrawSymbol {
    param($g, $pen, $brush, $accent, $accentBrush)
    $font = New-Object System.Drawing.Font('Microsoft YaHei UI', 15, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $format = New-Object System.Drawing.StringFormat
    $format.Alignment = [System.Drawing.StringAlignment]::Center
    $format.LineAlignment = [System.Drawing.StringAlignment]::Center
    $g.DrawString('大', $font, $brush, (New-Object System.Drawing.RectangleF(4, 3, 24, 24)), $format)
    $g.DrawRectangle($accent, 21, 21, 5, 5)
    $font.Dispose()
    $format.Dispose()
}

# Directory: a document with a compact list marker.
New-AccuXIcon -Path (Join-Path $OutputDirectory 'directory.png') -AccentColor ([System.Drawing.Color]::FromArgb(255, 166, 230, 183)) -DrawSymbol {
    param($g, $pen, $brush, $accent, $accentBrush)
    $g.DrawRectangle($pen, 8, 6, 16, 20)
    $g.DrawLine($accent, 12, 12, 14, 12)
    $g.DrawLine($accent, 12, 17, 14, 17)
    $g.DrawLine($accent, 12, 22, 14, 22)
    $g.DrawLine($pen, 17, 12, 21, 12)
    $g.DrawLine($pen, 17, 17, 21, 17)
    $g.DrawLine($pen, 17, 22, 21, 22)
}

# Comment assistant: a speech bubble with three text dots.
New-AccuXIcon -Path (Join-Path $OutputDirectory 'comment.png') -AccentColor ([System.Drawing.Color]::FromArgb(255, 255, 193, 87)) -DrawSymbol {
    param($g, $pen, $brush, $accent, $accentBrush)
    # 尾巴并入轮廓路径，避免在 32px 下被描边填成实心楔形。
    # 注意：GraphicsPath.AddLine(x, y) 会被 PowerShell 误绑到带默认 0 的重载，
    # 这里统一用 AddLines + PointF 传入。
    $bubble = New-Object System.Drawing.Drawing2D.GraphicsPath
    $bubble.AddArc(6, 7, 8, 8, 180, 90)
    $bubble.AddLines([System.Drawing.PointF[]]@((New-Object System.Drawing.PointF(22, 7))))
    $bubble.AddArc(18, 7, 8, 8, 270, 90)
    $bubble.AddLines([System.Drawing.PointF[]]@((New-Object System.Drawing.PointF(26, 18))))
    $bubble.AddArc(18, 14, 8, 8, 0, 90)
    $bubble.AddLines([System.Drawing.PointF[]]@(
        (New-Object System.Drawing.PointF(15, 22)),
        (New-Object System.Drawing.PointF(10, 27)),
        (New-Object System.Drawing.PointF(9, 22))))
    $bubble.AddArc(6, 14, 8, 8, 90, 90)
    $bubble.CloseFigure()
    $g.DrawPath($pen, $bubble)
    $g.FillEllipse($accentBrush, 10, 13, 3, 3)
    $g.FillEllipse($accentBrush, 15, 13, 3, 3)
    $g.FillEllipse($accentBrush, 20, 13, 3, 3)
    $bubble.Dispose()
}

function New-AccuXTextIcon {
    param(
        [string]$Path,
        [System.Drawing.Color]$BackgroundColor,
        [string]$Symbol
    )

    $size = 32
    $bitmap = New-Object System.Drawing.Bitmap($size, $size, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $graphics.Clear($BackgroundColor)

    $whiteBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
    $font = New-Object System.Drawing.Font('Segoe UI', 24, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $format = New-Object System.Drawing.StringFormat
    $format.Alignment = [System.Drawing.StringAlignment]::Center
    $format.LineAlignment = [System.Drawing.StringAlignment]::Center
    $graphics.DrawString($Symbol, $font, $whiteBrush, (New-Object System.Drawing.RectangleF(0, 0, 32, 32)), $format)

    $format.Dispose()
    $font.Dispose()
    $whiteBrush.Dispose()
    $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    $graphics.Dispose()
    $bitmap.Dispose()
}

# Mark colors: opaque color tile with the requested white letter.
New-AccuXTextIcon -Path (Join-Path $OutputDirectory 'mark-green.png') -BackgroundColor ([System.Drawing.Color]::FromArgb(255, 24, 190, 106)) -Symbol 'T'
New-AccuXTextIcon -Path (Join-Path $OutputDirectory 'mark-red.png') -BackgroundColor ([System.Drawing.Color]::FromArgb(255, 237, 64, 21)) -Symbol 'F'
New-AccuXTextIcon -Path (Join-Path $OutputDirectory 'mark-yellow.png') -BackgroundColor ([System.Drawing.Color]::FromArgb(255, 254, 153, 0)) -Symbol 'W'
New-AccuXTextIcon -Path (Join-Path $OutputDirectory 'mark-blue.png') -BackgroundColor ([System.Drawing.Color]::FromArgb(255, 45, 183, 245)) -Symbol 'M'

Write-Host "Icons written to $OutputDirectory"
Get-ChildItem $OutputDirectory -Filter *.png | ForEach-Object { Write-Host " - $($_.Name)" }
