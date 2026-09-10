# Generates AccuX Ribbon icons: 32x32 PNG, based on the approved AI concept board.
# The final assets keep the concept's bold keyline, white surfaces, and vivid
# accents, while the full canvas remains a 99% transparent pure-white matte
# for WPS/Office compatibility.
param(
    [string]$OutputDirectory = "$PSScriptRoot\..\src\AccuX.AddIn\Resources"
)

Add-Type -AssemblyName System.Drawing

if (-not (Test-Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

$script:KeylineColor = [System.Drawing.Color]::FromArgb(255, 18, 26, 35)
$script:WhiteColor = [System.Drawing.Color]::FromArgb(255, 249, 251, 253)
$script:BlueColor = [System.Drawing.Color]::FromArgb(255, 42, 137, 245)
$script:CyanColor = [System.Drawing.Color]::FromArgb(255, 56, 216, 222)
$script:GreenColor = [System.Drawing.Color]::FromArgb(255, 53, 211, 121)
$script:RedColor = [System.Drawing.Color]::FromArgb(255, 255, 76, 76)
$script:YellowColor = [System.Drawing.Color]::FromArgb(255, 255, 209, 43)
$script:SlateColor = [System.Drawing.Color]::FromArgb(255, 88, 103, 116)

function New-PointF {
    param([float]$X, [float]$Y)
    return [System.Drawing.PointF]::new($X, $Y)
}

function New-PolygonPath {
    param([System.Drawing.PointF[]]$Points)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddPolygon($Points)
    return $path
}

function New-RoundedRectanglePath {
    param([float]$X, [float]$Y, [float]$Width, [float]$Height, [float]$Radius)
    $diameter = $Radius * 2
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc($X, $Y, $diameter, $diameter, 180, 90)
    $path.AddArc($X + $Width - $diameter, $Y, $diameter, $diameter, 270, 90)
    $path.AddArc($X + $Width - $diameter, $Y + $Height - $diameter, $diameter, $diameter, 0, 90)
    $path.AddArc($X, $Y + $Height - $diameter, $diameter, $diameter, 90, 90)
    $path.CloseFigure()
    return $path
}

function New-DocumentPath {
    param([float]$X, [float]$Y, [float]$Width, [float]$Height, [float]$Fold)
    $points = [System.Drawing.PointF[]]@(
        (New-PointF $X $Y),
        (New-PointF ($X + $Width - $Fold) $Y),
        (New-PointF ($X + $Width) ($Y + $Fold)),
        (New-PointF ($X + $Width) ($Y + $Height)),
        (New-PointF $X ($Y + $Height))
    )
    return New-PolygonPath $points
}

function New-ChatPath {
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc(5, 6, 22, 17, 180, 90)
    $path.AddArc(15, 6, 12, 17, 270, 90)
    $path.AddArc(15, 11, 12, 12, 0, 90)
    $path.AddLines([System.Drawing.PointF[]]@(
        (New-PointF 16 23),
        (New-PointF 10 28),
        (New-PointF 11 22)
    ))
    $path.AddArc(5, 11, 12, 12, 90, 90)
    $path.CloseFigure()
    return $path
}

function Draw-OutlinedLine {
    param($Graphics, $EdgePen, $ForegroundPen, [float]$X1, [float]$Y1, [float]$X2, [float]$Y2)
    $Graphics.DrawLine($EdgePen, $X1, $Y1, $X2, $Y2)
    $Graphics.DrawLine($ForegroundPen, $X1, $Y1, $X2, $Y2)
}

function Draw-OutlinedArc {
    param($Graphics, $EdgePen, $ForegroundPen, [float]$X, [float]$Y, [float]$Width, [float]$Height, [float]$StartAngle, [float]$SweepAngle)
    $Graphics.DrawArc($EdgePen, $X, $Y, $Width, $Height, $StartAngle, $SweepAngle)
    $Graphics.DrawArc($ForegroundPen, $X, $Y, $Width, $Height, $StartAngle, $SweepAngle)
}

function Fill-OutlinedPath {
    param($Graphics, $EdgePen, $ForegroundBrush, $Path)
    $Graphics.FillPath($ForegroundBrush, $Path)
    $Graphics.DrawPath($EdgePen, $Path)
}

function Draw-OutlinedPath {
    param($Graphics, $EdgePen, $ForegroundPen, $Path)
    $Graphics.DrawPath($EdgePen, $Path)
    $Graphics.DrawPath($ForegroundPen, $Path)
}

function Fill-OutlinedEllipse {
    param($Graphics, $EdgeBrush, $ForegroundBrush, [float]$X, [float]$Y, [float]$Width, [float]$Height)
    $Graphics.FillEllipse($EdgeBrush, $X, $Y, $Width, $Height)
    $Graphics.FillEllipse($ForegroundBrush, $X + 0.75, $Y + 0.75, $Width - 1.5, $Height - 1.5)
}

function Fill-OutlinedRectangle {
    param($Graphics, $EdgeBrush, $ForegroundBrush, [float]$X, [float]$Y, [float]$Width, [float]$Height)
    $Graphics.FillRectangle($EdgeBrush, $X, $Y, $Width, $Height)
    $Graphics.FillRectangle($ForegroundBrush, $X + 0.75, $Y + 0.75, $Width - 1.5, $Height - 1.5)
}

function Draw-OutlinedText {
    param($Graphics, [string]$Text, $Font, $OutlineBrush, $ForegroundBrush, [System.Drawing.RectangleF]$Rectangle, $Format)
    foreach ($dx in @(-1, 0, 1)) {
        foreach ($dy in @(-1, 0, 1)) {
            if ($dx -ne 0 -or $dy -ne 0) {
                $shadowRectangle = [System.Drawing.RectangleF]::new(
                    $Rectangle.X + $dx,
                    $Rectangle.Y + $dy,
                    $Rectangle.Width,
                    $Rectangle.Height)
                $Graphics.DrawString($Text, $Font, $OutlineBrush, $shadowRectangle, $Format)
            }
        }
    }
    $Graphics.DrawString($Text, $Font, $ForegroundBrush, $Rectangle, $Format)
}

function New-IconPalette {
    $edgePen = New-Object System.Drawing.Pen($script:KeylineColor, 3.2)
    $lineEdgePen = New-Object System.Drawing.Pen($script:KeylineColor, 4.6)
    $whitePen = New-Object System.Drawing.Pen($script:WhiteColor, 2.45)
    $bluePen = New-Object System.Drawing.Pen($script:BlueColor, 2.55)
    $cyanPen = New-Object System.Drawing.Pen($script:CyanColor, 2.55)
    $greenPen = New-Object System.Drawing.Pen($script:GreenColor, 2.55)
    $redPen = New-Object System.Drawing.Pen($script:RedColor, 2.55)
    $yellowPen = New-Object System.Drawing.Pen($script:YellowColor, 2.55)
    $slatePen = New-Object System.Drawing.Pen($script:SlateColor, 2.1)

    foreach ($pen in @($lineEdgePen, $whitePen, $bluePen, $cyanPen, $greenPen, $redPen, $yellowPen)) {
        $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
        $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
        $pen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
    }

    return [pscustomobject]@{
        EdgePen = $edgePen
        LineEdgePen = $lineEdgePen
        WhitePen = $whitePen
        BluePen = $bluePen
        CyanPen = $cyanPen
        GreenPen = $greenPen
        RedPen = $redPen
        YellowPen = $yellowPen
        SlatePen = $slatePen
        EdgeBrush = New-Object System.Drawing.SolidBrush($script:KeylineColor)
        WhiteBrush = New-Object System.Drawing.SolidBrush($script:WhiteColor)
        BlueBrush = New-Object System.Drawing.SolidBrush($script:BlueColor)
        CyanBrush = New-Object System.Drawing.SolidBrush($script:CyanColor)
        GreenBrush = New-Object System.Drawing.SolidBrush($script:GreenColor)
        RedBrush = New-Object System.Drawing.SolidBrush($script:RedColor)
        YellowBrush = New-Object System.Drawing.SolidBrush($script:YellowColor)
        SlateBrush = New-Object System.Drawing.SolidBrush($script:SlateColor)
    }
}

function Dispose-IconPalette {
    param($Palette)
    foreach ($property in $Palette.PSObject.Properties) {
        if ($property.Value -is [System.IDisposable]) {
            $property.Value.Dispose()
        }
    }
}

function New-AccuXIcon {
    param(
        [string]$Path,
        [scriptblock]$DrawSymbol
    )

    $size = 32
    $bitmap = New-Object System.Drawing.Bitmap($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $graphics.Clear([System.Drawing.Color]::FromArgb(0, 0, 0, 0))

    # 3/255 ≈ 1% opacity, i.e. 99% transparency. SourceCopy preserves
    # the non-zero matte instead of rounding it back to Alpha=0.
    $matte = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(3, 255, 255, 255))
    $graphics.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
    $graphics.FillRectangle($matte, 0, 0, $size, $size)
    $graphics.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceOver

    $palette = New-IconPalette
    & $DrawSymbol $graphics $palette

    $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)

    Dispose-IconPalette $palette
    $matte.Dispose()
    $graphics.Dispose()
    $bitmap.Dispose()
}

function Draw-Highlighter {
    param($Graphics, $Palette, $AccentBrush, $AccentPen)

    $body = New-PolygonPath ([System.Drawing.PointF[]]@(
        (New-PointF 14 4),
        (New-PointF 28 18),
        (New-PointF 21 25),
        (New-PointF 7 11)
    ))
    Fill-OutlinedPath $Graphics $Palette.EdgePen $AccentBrush $body
    $body.Dispose()

    Draw-OutlinedLine $Graphics $Palette.LineEdgePen $Palette.WhitePen 14 7 23 16

    $collar = New-PolygonPath ([System.Drawing.PointF[]]@(
        (New-PointF 7 11),
        (New-PointF 13 17),
        (New-PointF 8 22),
        (New-PointF 2 16)
    ))
    Fill-OutlinedPath $Graphics $Palette.EdgePen $Palette.WhiteBrush $collar
    $collar.Dispose()

    $tip = New-PolygonPath ([System.Drawing.PointF[]]@(
        (New-PointF 2 16),
        (New-PointF 8 22),
        (New-PointF 3 27),
        (New-PointF 1 23)
    ))
    Fill-OutlinedPath $Graphics $Palette.EdgePen $AccentBrush $tip
    $tip.Dispose()

    Draw-OutlinedLine $Graphics $Palette.LineEdgePen $AccentPen 4 28 27 28
}

# 1. Rounding: a white precision dial with a blue tick and arrow.
New-AccuXIcon -Path (Join-Path $OutputDirectory 'round.png') -DrawSymbol {
    param($g, $p)
    $g.FillEllipse($p.EdgeBrush, 8, 8, 16, 16)
    Draw-OutlinedArc $g $p.LineEdgePen $p.WhitePen 4 4 24 24 35 285
    $arrow = New-PolygonPath ([System.Drawing.PointF[]]@(
        (New-PointF 22 5),
        (New-PointF 28 7),
        (New-PointF 26 13),
        (New-PointF 23 10)
    ))
    Fill-OutlinedPath $g $p.EdgePen $p.WhiteBrush $arrow
    $arrow.Dispose()
    $font = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $format = New-Object System.Drawing.StringFormat
    $format.Alignment = [System.Drawing.StringAlignment]::Center
    $format.LineAlignment = [System.Drawing.StringAlignment]::Center
    Draw-OutlinedText $g '0.0' $font $p.EdgeBrush $p.WhiteBrush ([System.Drawing.RectangleF]::new(8, 10, 16, 8)) $format
    Draw-OutlinedLine $g $p.LineEdgePen $p.BluePen 16 21 16 25
    $font.Dispose()
    $format.Dispose()
}

# 2. Amount conversion: opposing blue and cyan arrows.
New-AccuXIcon -Path (Join-Path $OutputDirectory 'convert.png') -DrawSymbol {
    param($g, $p)
    $topArrow = New-PolygonPath ([System.Drawing.PointF[]]@(
        (New-PointF 7 8),
        (New-PointF 21 8),
        (New-PointF 21 5),
        (New-PointF 28 11),
        (New-PointF 21 17),
        (New-PointF 21 14),
        (New-PointF 7 14)
    ))
    Fill-OutlinedPath $g $p.EdgePen $p.BlueBrush $topArrow
    $topArrow.Dispose()

    $bottomArrow = New-PolygonPath ([System.Drawing.PointF[]]@(
        (New-PointF 25 18),
        (New-PointF 11 18),
        (New-PointF 11 15),
        (New-PointF 4 21),
        (New-PointF 11 27),
        (New-PointF 11 24),
        (New-PointF 25 24)
    ))
    Fill-OutlinedPath $g $p.EdgePen $p.CyanBrush $bottomArrow
    $bottomArrow.Dispose()
}

# 3. Selection sum: document/list plus a large sigma.
New-AccuXIcon -Path (Join-Path $OutputDirectory 'sum.png') -DrawSymbol {
    param($g, $p)
    $document = New-DocumentPath 4 5 23 22 5
    Fill-OutlinedPath $g $p.EdgePen $p.WhiteBrush $document
    $document.Dispose()
    Draw-OutlinedLine $g $p.LineEdgePen $p.BluePen 7 10 16 10
    Draw-OutlinedLine $g $p.LineEdgePen $p.BluePen 7 14 14 14
    Draw-OutlinedLine $g $p.LineEdgePen $p.BluePen 7 18 13 18
    $font = New-Object System.Drawing.Font('Segoe UI Symbol', 18, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $format = New-Object System.Drawing.StringFormat
    $format.Alignment = [System.Drawing.StringAlignment]::Center
    $format.LineAlignment = [System.Drawing.StringAlignment]::Center
    Draw-OutlinedText $g 'Σ' $font $p.EdgeBrush $p.BlueBrush ([System.Drawing.RectangleF]::new(14, 10, 12, 15)) $format
    $font.Dispose()
    $format.Dispose()
}

# 4. Chinese uppercase amount: a white document behind a bold blue 大.
New-AccuXIcon -Path (Join-Path $OutputDirectory 'uppercase.png') -DrawSymbol {
    param($g, $p)
    $document = New-DocumentPath 4 5 23 22 5
    Fill-OutlinedPath $g $p.EdgePen $p.WhiteBrush $document
    $document.Dispose()
    $font = New-Object System.Drawing.Font('Microsoft YaHei UI', 17, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $format = New-Object System.Drawing.StringFormat
    $format.Alignment = [System.Drawing.StringAlignment]::Center
    $format.LineAlignment = [System.Drawing.StringAlignment]::Center
    Draw-OutlinedText $g '大' $font $p.EdgeBrush $p.BlueBrush ([System.Drawing.RectangleF]::new(6, 7, 20, 18)) $format
    $font.Dispose()
    $format.Dispose()
}

# 5. Workbook directory: document, index bullets, and list lines.
New-AccuXIcon -Path (Join-Path $OutputDirectory 'directory.png') -DrawSymbol {
    param($g, $p)
    $document = New-DocumentPath 5 4 22 24 5
    Fill-OutlinedPath $g $p.EdgePen $p.WhiteBrush $document
    $document.Dispose()
    Fill-OutlinedRectangle $g $p.EdgeBrush $p.BlueBrush 8 10 3.5 3.5
    Fill-OutlinedRectangle $g $p.EdgeBrush $p.BlueBrush 8 16 3.5 3.5
    Fill-OutlinedRectangle $g $p.EdgeBrush $p.BlueBrush 8 22 3.5 3.5
    Draw-OutlinedLine $g $p.LineEdgePen $p.SlatePen 14 11.5 23 11.5
    Draw-OutlinedLine $g $p.LineEdgePen $p.SlatePen 14 17.5 23 17.5
    Draw-OutlinedLine $g $p.LineEdgePen $p.SlatePen 14 23.5 23 23.5
}

# 6. Comment assistant: white bubble with a blue outline and blue dots.
New-AccuXIcon -Path (Join-Path $OutputDirectory 'comment.png') -DrawSymbol {
    param($g, $p)
    $bubble = New-ChatPath
    $g.FillPath($p.WhiteBrush, $bubble)
    $g.DrawPath($p.EdgePen, $bubble)
    $g.DrawPath($p.BluePen, $bubble)
    $bubble.Dispose()
    Fill-OutlinedEllipse $g $p.EdgeBrush $p.BlueBrush 10 13 3.5 3.5
    Fill-OutlinedEllipse $g $p.EdgeBrush $p.BlueBrush 15 13 3.5 3.5
    Fill-OutlinedEllipse $g $p.EdgeBrush $p.BlueBrush 20 13 3.5 3.5
}

# 7. Region compare: two sheets with green/blue cells and center chevrons.
New-AccuXIcon -Path (Join-Path $OutputDirectory 'compare.png') -DrawSymbol {
    param($g, $p)
    # The compare control is a large Ribbon button, so use almost the full
    # 32px canvas and keep the comparison mark visually dominant.
    $leftDocument = New-DocumentPath 1 3 14 26 5
    $rightDocument = New-DocumentPath 17 3 14 26 5
    Fill-OutlinedPath $g $p.EdgePen $p.WhiteBrush $leftDocument
    Fill-OutlinedPath $g $p.EdgePen $p.WhiteBrush $rightDocument
    $leftDocument.Dispose()
    $rightDocument.Dispose()
    Fill-OutlinedRectangle $g $p.EdgeBrush $p.GreenBrush 4 8 4.5 4.5
    Fill-OutlinedRectangle $g $p.EdgeBrush $p.BlueBrush 23.5 8 4.5 4.5
    Draw-OutlinedLine $g $p.LineEdgePen $p.SlatePen 10 10.25 13 10.25
    Draw-OutlinedLine $g $p.LineEdgePen $p.SlatePen 20 10.25 23 10.25
    Draw-OutlinedLine $g $p.LineEdgePen $p.SlatePen 4 16 12 16
    Draw-OutlinedLine $g $p.LineEdgePen $p.SlatePen 20 16 28 16
    Draw-OutlinedLine $g $p.LineEdgePen $p.SlatePen 4 23 12 23
    Draw-OutlinedLine $g $p.LineEdgePen $p.SlatePen 20 23 28 23
    $leftChevron = New-PolygonPath ([System.Drawing.PointF[]]@(
        (New-PointF 15 12), (New-PointF 11 16), (New-PointF 15 20)
    ))
    $rightChevron = New-PolygonPath ([System.Drawing.PointF[]]@(
        (New-PointF 17 12), (New-PointF 21 16), (New-PointF 17 20)
    ))
    Fill-OutlinedPath $g $p.EdgePen $p.WhiteBrush $leftChevron
    Fill-OutlinedPath $g $p.EdgePen $p.WhiteBrush $rightChevron
    $leftChevron.Dispose()
    $rightChevron.Dispose()
}

# 8-11. Mark actions: the four approved color variants of a highlighter.
New-AccuXIcon -Path (Join-Path $OutputDirectory 'mark-green.png') -DrawSymbol {
    param($g, $p)
    Draw-Highlighter $g $p $p.GreenBrush $p.GreenPen
}
New-AccuXIcon -Path (Join-Path $OutputDirectory 'mark-red.png') -DrawSymbol {
    param($g, $p)
    Draw-Highlighter $g $p $p.RedBrush $p.RedPen
}
New-AccuXIcon -Path (Join-Path $OutputDirectory 'mark-yellow.png') -DrawSymbol {
    param($g, $p)
    Draw-Highlighter $g $p $p.YellowBrush $p.YellowPen
}
New-AccuXIcon -Path (Join-Path $OutputDirectory 'mark-blue.png') -DrawSymbol {
    param($g, $p)
    Draw-Highlighter $g $p $p.BlueBrush $p.BluePen
}

Write-Host "Icons written to $OutputDirectory"
Get-ChildItem $OutputDirectory -Filter *.png | Sort-Object Name | ForEach-Object { Write-Host " - $($_.Name)" }
