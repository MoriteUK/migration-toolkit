#Requires -Version 7.0
<#
.SYNOPSIS
    Builds FlyMigration.exe – a no-console Windows launcher for menu.ps1.
    Run once from this folder.  Output: FlyMigration.ico + FlyMigration.exe.  
#>
$ErrorActionPreference = 'Stop'
$Root = $PSScriptRoot

Add-Type -AssemblyName System.Drawing

# ─── 1. Generate FlyMigration.ico ─────────────────────────────────────────────
$icoPath = Join-Path $Root 'FlyMigration.ico'

function Build-Ico {
    # Paper-plane on a blue gradient rounded square
    $bmp = New-Object System.Drawing.Bitmap(256, 256)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $g.Clear([System.Drawing.Color]::Transparent)

    # Rounded-square background with diagonal gradient
    $r    = 44
    $path = [System.Drawing.Drawing2D.GraphicsPath]::new()
    $path.AddArc(  0,   0, $r*2, $r*2, 180, 90)
    $path.AddArc(256-$r*2,   0, $r*2, $r*2, 270, 90)
    $path.AddArc(256-$r*2, 256-$r*2, $r*2, $r*2,   0, 90)
    $path.AddArc(  0, 256-$r*2, $r*2, $r*2,  90, 90)
    $path.CloseFigure()

    $grad = [System.Drawing.Drawing2D.LinearGradientBrush]::new(
        [System.Drawing.Point]::new(0, 0),
        [System.Drawing.Point]::new(256, 256),
        [System.Drawing.Color]::FromArgb(30, 130, 210),
        [System.Drawing.Color]::FromArgb(0,  55, 130))
    $g.FillPath($grad, $path)
    $grad.Dispose(); $path.Dispose()

    # Polygon points for each part of the paper plane
    $upper = [System.Drawing.Point[]]@(
        [System.Drawing.Point]::new(228, 128),
        [System.Drawing.Point]::new(28,   28),
        [System.Drawing.Point]::new(92,  128))

    $lower = [System.Drawing.Point[]]@(
        [System.Drawing.Point]::new(228, 128),
        [System.Drawing.Point]::new(92,  128),
        [System.Drawing.Point]::new(28,  228))

    $fold  = [System.Drawing.Point[]]@(
        [System.Drawing.Point]::new(92,  128),
        [System.Drawing.Point]::new(28,  228),
        [System.Drawing.Point]::new(92,  178))

    # Upper and lower wing — white
    $g.FillPolygon([System.Drawing.Brushes]::White, $upper)
    $g.FillPolygon([System.Drawing.Brushes]::White, $lower)

    # Fold flap — light-blue tint to suggest depth
    $foldBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(190, 160, 205, 235))
    $g.FillPolygon($foldBrush, $fold)
    $foldBrush.Dispose()

    # Hairline outline for crispness at small sizes
    $pen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(60, 255, 255, 255), 1.5)
    $g.DrawPolygon($pen, $upper)
    $g.DrawPolygon($pen, $lower)
    $pen.Dispose()

    $g.Dispose()
    $src = $bmp

    # Render each size as a complete PNG byte array
    $sizes  = 16, 32, 48, 64, 128, 256
    $frames = foreach ($s in $sizes) {
        $b  = New-Object System.Drawing.Bitmap($s, $s)
        $g2 = [System.Drawing.Graphics]::FromImage($b)
        $g2.InterpolationMode  = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g2.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $g2.DrawImage($src, 0, 0, $s, $s)
        $g2.Dispose()
        $ms = [System.IO.MemoryStream]::new()
        $b.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        $b.Dispose()
        , $ms.ToArray()   # comma keeps each array as a single pipeline item
    }
    $src.Dispose()

    # Write Vista+ PNG-in-ICO container
    $ms = [System.IO.MemoryStream]::new()
    $bw = [System.IO.BinaryWriter]::new($ms)

    $bw.Write([uint16]0)                  # Reserved
    $bw.Write([uint16]1)                  # Type = ICO
    $bw.Write([uint16]$sizes.Count)       # Number of images

    $offset = [uint32](6 + $sizes.Count * 16)
    for ($i = 0; $i -lt $sizes.Count; $i++) {
        $s = $sizes[$i]; $d = $frames[$i]
        $wh = if ($s -ge 256) { [byte]0 } else { [byte]$s }
        $bw.Write($wh)                    # bWidth  (0 = 256)
        $bw.Write($wh)                    # bHeight (0 = 256)
        $bw.Write([byte]0)                # bColorCount  (0 = full colour)
        $bw.Write([byte]0)                # bReserved
        $bw.Write([uint16]1)              # wPlanes
        $bw.Write([uint16]32)             # wBitCount
        $bw.Write([uint32]$d.Length)      # dwBytesInRes
        $bw.Write($offset)                # dwImageOffset
        $offset += [uint32]$d.Length
    }
    foreach ($d in $frames) { $bw.Write($d) }

    $bw.Flush()
    [System.IO.File]::WriteAllBytes($icoPath, $ms.ToArray())
    $bw.Dispose(); $ms.Dispose()
}

Write-Host '==> Generating icon ...' -ForegroundColor Cyan
Build-Ico
Write-Host "    $icoPath" -ForegroundColor Green

# ─── 2. C# launcher source ────────────────────────────────────────────────────
$cs = @'
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;

class FlyLauncher {
    [STAThread]
    static void Main() {
        string dir  = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        string ps1  = Path.Combine(dir, "menu.ps1");
        string psExe = FindPwsh();

        var psi = new ProcessStartInfo {
            FileName         = psExe,
            Arguments        = "-NoProfile -ExecutionPolicy Bypass -File \"" + ps1 + "\"",
            WorkingDirectory = dir,
            UseShellExecute  = false,
            CreateNoWindow   = true   // no black console flicker; menu.ps1 shows its own GUI
        };
        using (var p = Process.Start(psi)) { p.WaitForExit(); }
    }

    static string FindPwsh() {
        // Prefer PowerShell 7 (pwsh.exe); fall back to Windows PowerShell
        foreach (string seg in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(';')) {
            try {
                string f = Path.Combine(seg.Trim(), "pwsh.exe");
                if (File.Exists(f)) return f;
            } catch { }
        }
        return "powershell.exe";
    }
}
'@

# ─── 3. Compile ───────────────────────────────────────────────────────────────
$outExe = Join-Path $Root 'FlyMigration.exe'
$csTmp  = Join-Path $env:TEMP 'FlyLauncher_build.cs'

Write-Host '==> Compiling FlyMigration.exe ...' -ForegroundColor Cyan

# csc.exe ships with .NET Framework on every Windows install
$csc = @(
    "$env:SystemRoot\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "$env:SystemRoot\Microsoft.NET\Framework\v4.0.30319\csc.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if ($csc) {
    $cs | Set-Content $csTmp -Encoding UTF8
    $result = & $csc /nologo /target:winexe /optimize+ `
                  "/win32icon:$icoPath" `
                  "/out:$outExe" `
                  $csTmp 2>&1
    Remove-Item $csTmp -ErrorAction SilentlyContinue
    if ($LASTEXITCODE -ne 0) { throw "Compile failed:`n$result" }
} else {
    Write-Warning 'csc.exe not found – falling back to Add-Type (Roslyn)'
    Add-Type -TypeDefinition $cs `
             -OutputAssembly $outExe `
             -OutputType WindowsApplication `
             -CompilerOptions "/win32icon:`"$icoPath`""
}

Write-Host "    $outExe" -ForegroundColor Green
Write-Host ''
Write-Host 'Done. Double-click FlyMigration.exe to launch the toolkit.' -ForegroundColor Cyan
