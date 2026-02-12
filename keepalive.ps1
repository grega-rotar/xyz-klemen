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
$jiggleDistance = 5         # Pixels to move
$jiggleInterval = 2000      # Milliseconds between jiggles
# =======================================================

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Simple Mouse Jiggler v4.0" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Configuration:" -ForegroundColor Yellow
Write-Host "- Idle threshold: $idleThresholdSeconds seconds" -ForegroundColor Gray
Write-Host "- Jiggle distance: $jiggleDistance pixels" -ForegroundColor Gray
Write-Host "- Jiggle interval: $jiggleInterval ms" -ForegroundColor Gray
Write-Host "- Press Ctrl+C to stop" -ForegroundColor Red
Write-Host ""
Write-Host "Status: Monitoring..." -ForegroundColor Green
Write-Host ""

$isJiggling = $false

while ($true) {
    # Get system idle time
    $systemIdleMs = [IdleTimeHelper]::GetIdleTime()
    $idleSeconds = $systemIdleMs / 1000
    
    # Check if user became active
    if ($systemIdleMs -lt 500 -and $isJiggling) {
        $isJiggling = $false
        Write-Host "`r[$(Get-Date -Format 'HH:mm:ss')] User active - Jiggling stopped          " -ForegroundColor Yellow
        Write-Host ""
    }
    
    # Start jiggling if idle threshold exceeded
    if ($idleSeconds -ge $idleThresholdSeconds -and -not $isJiggling) {
        $isJiggling = $true
        Write-Host "`r[$(Get-Date -Format 'HH:mm:ss')] Idle detected - Starting jiggle          " -ForegroundColor Green
        Write-Host ""
    }
    
    # Perform jiggle if active
    if ($isJiggling) {
        $currentPos = [System.Windows.Forms.Cursor]::Position
        
        # Move right
        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(($currentPos.X + $jiggleDistance), $currentPos.Y)
        Start-Sleep -Milliseconds 100
        
        # Move left
        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(($currentPos.X - $jiggleDistance), $currentPos.Y)
        Start-Sleep -Milliseconds 100
        
        # Move back to original
        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($currentPos.X, $currentPos.Y)
        
        Write-Host "`r[$(Get-Date -Format 'HH:mm:ss')] JIGGLING | Idle: $([Math]::Floor($idleSeconds))s    " -NoNewline -ForegroundColor Cyan
        
        Start-Sleep -Milliseconds $jiggleInterval
    }
    else {
        # Monitor idle time
        $idleDisplay = [Math]::Floor($idleSeconds)
        $remaining = [Math]::Max(0, $idleThresholdSeconds - $idleDisplay)
        Write-Host "`r[$(Get-Date -Format 'HH:mm:ss')] MONITORING | Idle: $idleDisplay s | Jiggle starts in: $remaining s    " -NoNewline -ForegroundColor Gray
        Start-Sleep -Milliseconds 200
    }
}
