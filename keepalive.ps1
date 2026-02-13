Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;

public class IdleTimeHelper {
    [DllImport("user32.dll")]
    public static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

    public static uint GetIdleTime() {
        LASTINPUTINFO lastInputInfo = new LASTINPUTINFO();
        lastInputInfo.cbSize = (uint)Marshal.SizeOf(lastInputInfo);
        GetLastInputInfo(ref lastInputInfo);
        return ((uint)Environment.TickCount - lastInputInfo.dwTime);
    }
}

public struct LASTINPUTINFO {
    public uint cbSize;
    public uint dwTime;
}
"@

# ==================== CONFIGURATION ====================
$idleThresholdSeconds = 20  # Start jiggling after this many seconds
$maxRadius = 300            # Max radius of the spiral in pixels
$minRadius = 100            # Min radius (center point)
$radiusSteps = 15           # Number of radius increments from min to max
$pointsPerLoop = 60         # Points per full rotation (smoothness)
$stepDelayMs = 25           # Milliseconds between each step
# =======================================================

# Get screen center
$screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height
$centerX = [Math]::Floor($screenWidth / 2)
$centerY = [Math]::Floor($screenHeight / 2)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Spiral Mouse Jiggler v6.0" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Configuration:" -ForegroundColor Yellow
Write-Host "- Idle threshold: $idleThresholdSeconds seconds" -ForegroundColor Gray
Write-Host "- Spiral radius: $minRadius - $maxRadius pixels" -ForegroundColor Gray
Write-Host "- Screen center: ($centerX, $centerY)" -ForegroundColor Gray
Write-Host "- Press Ctrl+C to stop" -ForegroundColor Red
Write-Host ""
Write-Host "Status: Monitoring..." -ForegroundColor Green
Write-Host ""

$isJiggling = $false

# Build spiral path: outward then inward
function Get-SpiralPath {
    $path = @()
    $totalSteps = $radiusSteps * $pointsPerLoop

    # Spiral outward
    for ($i = 0; $i -lt $totalSteps; $i++) {
        $progress = $i / $totalSteps
        $radius = $minRadius + ($maxRadius - $minRadius) * $progress
        $angle = (2 * [Math]::PI * $i) / $pointsPerLoop
        $x = [Math]::Round($centerX + $radius * [Math]::Cos($angle))
        $y = [Math]::Round($centerY + $radius * [Math]::Sin($angle))
        $path += ,@($x, $y)
    }

    # Spiral inward
    for ($i = $totalSteps - 1; $i -ge 0; $i--) {
        $progress = $i / $totalSteps
        $radius = $minRadius + ($maxRadius - $minRadius) * $progress
        $angle = (2 * [Math]::PI * $i) / $pointsPerLoop
        $x = [Math]::Round($centerX + $radius * [Math]::Cos($angle))
        $y = [Math]::Round($centerY + $radius * [Math]::Sin($angle))
        $path += ,@($x, $y)
    }

    return ,$path
}

$spiralPath = Get-SpiralPath

while ($true) {
    # Get system idle time
    $systemIdleMs = [IdleTimeHelper]::GetIdleTime()
    $idleSeconds = $systemIdleMs / 1000

    # Check if user became active
    if ($systemIdleMs -lt 500 -and $isJiggling) {
        $isJiggling = $false
        Write-Host "`r[$(Get-Date -Format 'HH:mm:ss')] User active - Stopped                      " -ForegroundColor Yellow
        Write-Host ""
    }

    # Start jiggling if idle threshold exceeded
    if ($idleSeconds -ge $idleThresholdSeconds -and -not $isJiggling) {
        $isJiggling = $true
        Write-Host "`r[$(Get-Date -Format 'HH:mm:ss')] Idle detected - Starting spiral             " -ForegroundColor Green
        Write-Host ""
    }

    # Perform spiral if active
    if ($isJiggling) {
        foreach ($point in $spiralPath) {
            # Check if user became active mid-spiral
            $checkIdle = [IdleTimeHelper]::GetIdleTime()
            if ($checkIdle -lt 500) {
                $isJiggling = $false
                Write-Host "`r[$(Get-Date -Format 'HH:mm:ss')] User active - Stopped                      " -ForegroundColor Yellow
                Write-Host ""
                break
            }

            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($point[0], $point[1])
            Start-Sleep -Milliseconds $stepDelayMs
        }

        if ($isJiggling) {
            $idleNow = [Math]::Floor([IdleTimeHelper]::GetIdleTime() / 1000)
            Write-Host "`r[$(Get-Date -Format 'HH:mm:ss')] SPIRALING | Idle: ${idleNow}s    " -NoNewline -ForegroundColor Cyan
        }
    }
    else {
        # Monitor idle time
        $idleDisplay = [Math]::Floor($idleSeconds)
        $remaining = [Math]::Max(0, $idleThresholdSeconds - $idleDisplay)
        Write-Host "`r[$(Get-Date -Format 'HH:mm:ss')] MONITORING | Idle: $idleDisplay s | Jiggle starts in: $remaining s    " -NoNewline -ForegroundColor Gray
        Start-Sleep -Milliseconds 200
    }
}
