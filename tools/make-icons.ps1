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

Write-Host "Icons written to $OutputDirectory"
Get-ChildItem $OutputDirectory -Filter *.png | ForEach-Object { Write-Host " - $($_.Name)" }
