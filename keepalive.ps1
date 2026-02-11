<#
.SYNOPSIS
    Smart Mouse Jiggler - Prevents system idle/sleep with intelligent activity detection
.DESCRIPTION
    Automatically jiggles mouse with natural movement when system is idle.
    Draws "GREGA ROTAR" on screen when unattended.
    Can simulate mouse clicks for drawing in Paint or similar applications.
.AUTHOR
    Grega Rotar
    Email: grega@etiam.si
.VERSION
    2.2
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class MouseOperations {
    [DllImport("user32.dll")]
    public static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);
    
    public const int MOUSEEVENTF_LEFTDOWN = 0x02;
    public const int MOUSEEVENTF_LEFTUP = 0x04;
    
    public static void MouseDown() {
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
    }
    
    public static void MouseUp() {
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
    }
}
"@

# ==================== CONFIGURATION ====================
$idleThresholdSeconds = 20
$scrollLockInterval = 30  # Seconds
$drawingSpeed = 8  # Milliseconds between drawing steps (lower = faster)

# WARNING: MOUSE CLICK SIMULATION (for Paint, drawing apps, etc.)
$enableMouseClicks = $true  # Set to $true to enable actual drawing with mouse clicks
# =======================================================

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Smart Mouse Jiggler v2.2" -ForegroundColor White
Write-Host "  Author: Grega Rotar" -ForegroundColor Gray
Write-Host "  Email: grega@etiam.si" -ForegroundColor Gray
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Configuration:" -ForegroundColor Yellow
Write-Host "- Idle threshold: $idleThresholdSeconds seconds" -ForegroundColor Gray
Write-Host "- Mode: Draws 'GREGA ROTAR' signature" -ForegroundColor Gray
Write-Host "- Mouse clicks: $(if($enableMouseClicks){'ENABLED (Paint mode)'} else {'DISABLED (Test mode)'})" -ForegroundColor $(if($enableMouseClicks){'Red'} else {'Green'})
Write-Host "- Drawing speed: $drawingSpeed ms" -ForegroundColor Gray
Write-Host "- Press Ctrl+C to stop" -ForegroundColor Red
Write-Host ""
if ($enableMouseClicks) {
    Write-Host "WARNING: Mouse clicks are ENABLED!" -ForegroundColor Red
    Write-Host "   This will actually draw in Paint or any active window!" -ForegroundColor Red
    Write-Host "   Starting in 5 seconds... Press Ctrl+C to cancel" -ForegroundColor Yellow
    Start-Sleep -Seconds 5
    Write-Host ""
}
Write-Host "Status: Monitoring mouse activity..." -ForegroundColor Green
Write-Host ""

$wshell = New-Object -com "Wscript.Shell"
$lastMousePos = [System.Windows.Forms.Cursor]::Position
$lastActivityTime = Get-Date
$isDrawing = $false
$scrollLockTimer = 0
$drawCount = 0

# Letter drawing patterns (relative coordinates)
function Get-LetterG { 
    param($startX, $startY, $scale)
    $s = [double]$scale
    return @(
        @([int]($startX + 30*$s), [int]($startY)),
        @([int]($startX), [int]($startY)),
        @([int]($startX), [int]($startY + 50*$s)),
        @([int]($startX + 30*$s), [int]($startY + 50*$s)),
        @([int]($startX + 30*$s), [int]($startY + 25*$s)),
        @([int]($startX + 15*$s), [int]($startY + 25*$s))
    )
}

function Get-LetterR {
    param($startX, $startY, $scale)
    $s = [double]$scale
    return @(
        @([int]($startX), [int]($startY + 50*$s)),
        @([int]($startX), [int]($startY)),
        @([int]($startX + 25*$s), [int]($startY)),
        @([int]($startX + 25*$s), [int]($startY + 25*$s)),
        @([int]($startX), [int]($startY + 25*$s)),
        @([int]($startX + 25*$s), [int]($startY + 50*$s))
    )
}

function Get-LetterE {
    param($startX, $startY, $scale)
    $s = [double]$scale
    return @(
        @([int]($startX + 25*$s), [int]($startY)),
        @([int]($startX), [int]($startY)),
        @([int]($startX), [int]($startY + 25*$s)),
        @([int]($startX + 20*$s), [int]($startY + 25*$s)),
        @([int]($startX), [int]($startY + 25*$s)),
        @([int]($startX), [int]($startY + 50*$s)),
        @([int]($startX + 25*$s), [int]($startY + 50*$s))
    )
}

function Get-LetterA {
    param($startX, $startY, $scale)
    $s = [double]$scale
    return @(
        @([int]($startX), [int]($startY + 50*$s)),
        @([int]($startX), [int]($startY + 10*$s)),
        @([int]($startX + 12*$s), [int]($startY)),
        @([int]($startX + 25*$s), [int]($startY + 10*$s)),
        @([int]($startX + 25*$s), [int]($startY + 50*$s)),
        @([int]($startX + 25*$s), [int]($startY + 30*$s)),
        @([int]($startX), [int]($startY + 30*$s))
    )
}

function Get-LetterT {
    param($startX, $startY, $scale)
    $s = [double]$scale
    return @(
        @([int]($startX), [int]($startY)),
        @([int]($startX + 25*$s), [int]($startY)),
        @([int]($startX + 12*$s), [int]($startY)),
        @([int]($startX + 12*$s), [int]($startY + 50*$s))
    )
}

function Get-LetterO {
    param($startX, $startY, $scale)
    $s = [double]$scale
    return @(
        @([int]($startX), [int]($startY)),
        @([int]($startX + 25*$s), [int]($startY)),
        @([int]($startX + 25*$s), [int]($startY + 50*$s)),
        @([int]($startX), [int]($startY + 50*$s)),
        @([int]($startX), [int]($startY))
    )
}

function Get-Word {
    param($word, $startX, $startY, $scale, $spacing)
    
    $pattern = @()
    $currentX = $startX
    
    for ($i = 0; $i -lt $word.Length; $i++) {
        $letter = $word[$i]
        
        switch ($letter) {
            'G' { $pattern += Get-LetterG -startX $currentX -startY $startY -scale $scale }
            'R' { $pattern += Get-LetterR -startX $currentX -startY $startY -scale $scale }
            'E' { $pattern += Get-LetterE -startX $currentX -startY $startY -scale $scale }
            'A' { $pattern += Get-LetterA -startX $currentX -startY $startY -scale $scale }
            'T' { $pattern += Get-LetterT -startX $currentX -startY $startY -scale $scale }
            'O' { $pattern += Get-LetterO -startX $currentX -startY $startY -scale $scale }
        }
        
        # Add pen lift marker between letters (null point)
        if ($i -lt $word.Length - 1) {
            $pattern += ,@($null, $null)
        }
        
        $currentX += $spacing
    }
    
    return $pattern
}

function Get-GregarotarPattern {
    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $centerX = $screen.Width / 2 - 250
    $centerY = $screen.Height / 2 - 50
    $scale = 2.0
    $spacing = 45 * $scale
    
    $pattern = @()
    
    # GREGA
    $pattern += Get-Word -word "GREGA" -startX $centerX -startY $centerY -scale $scale -spacing $spacing
    
    # Space between words - pen lift
    $pattern += ,@($null, $null)
    
    # ROTAR
    $pattern += Get-Word -word "ROTAR" -startX ($centerX + $spacing * 6) -startY $centerY -scale $scale -spacing $spacing
    
    return $pattern
}

function Move-MouseSmoothly {
    param($startPos, $endPos, $speed, $isDrawing)
    
    $deltaX = $endPos[0] - $startPos.X
    $deltaY = $endPos[1] - $startPos.Y
    $distance = [Math]::Sqrt($deltaX * $deltaX + $deltaY * $deltaY)
    $steps = [Math]::Max(10, [int]($distance / 5))
    
    for ($i = 1; $i -le $steps; $i++) {
        $interpX = [int]($startPos.X + ($deltaX * $i / $steps))
        $interpY = [int]($startPos.Y + ($deltaY * $i / $steps))
        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($interpX, $interpY)
        Start-Sleep -Milliseconds $speed
    }
}

function Draw-GregarotarSignature {
    param($withClicks)
    
    $pattern = Get-GregarotarPattern
    $currentPos = [System.Windows.Forms.Cursor]::Position
    $pointCount = 0
    $totalPoints = ($pattern | Where-Object { $_[0] -ne $null }).Count
    $isPenDown = $false
    
    Write-Host "`r[$(Get-Date -Format 'HH:mm:ss')] Drawing 'GREGA ROTAR' signature ($totalPoints points)...                    " -ForegroundColor Magenta
    if ($withClicks) {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [PAINT MODE] Actually drawing!" -ForegroundColor Red
    }
    
    foreach ($point in $pattern) {
        # Check for pen lift (null point)
        if ($point[0] -eq $null) {
            if ($isPenDown -and $withClicks) {
                [MouseOperations]::MouseUp()
                $isPenDown = $false
                Write-Host "`r[$(Get-Date -Format 'HH:mm:ss')] Pen lifted                    " -NoNewline -ForegroundColor Yellow
                Start-Sleep -Milliseconds 50
            }
            continue
        }
        
        $pointCount++
        $currentPos = [System.Windows.Forms.Cursor]::Position
        
        # Move to next point
        Move-MouseSmoothly -startPos $currentPos -endPos $point -speed $drawingSpeed -isDrawing $isPenDown
        
        # Put pen down if not already down
        if (-not $isPenDown -and $withClicks) {
            [MouseOperations]::MouseDown()
            $isPenDown = $true
            Write-Host "`r[$(Get-Date -Format 'HH:mm:ss')] Pen down - Drawing...                    " -NoNewline -ForegroundColor Green
            Start-Sleep -Milliseconds 50
        }
        
        # Progress indicator
        if ($pointCount % 15 -eq 0) {
            $progress = [int](($pointCount / $totalPoints) * 100)
            $modeText = if ($withClicks) { "[DRAWING]" } else { "[TESTING]" }
            Write-Host "`r[$(Get-Date -Format 'HH:mm:ss')] $modeText Progress: $progress% ($pointCount/$totalPoints)    " -NoNewline -ForegroundColor Magenta
        }
        
        # Check for user interruption
        Start-Sleep -Milliseconds 5
        $newPos = [System.Windows.Forms.Cursor]::Position
        $drift = [Math]::Abs($newPos.X - $point[0]) + [Math]::Abs($newPos.Y - $point[1])
        
        if ($drift -gt 15) {
            if ($isPenDown -and $withClicks) {
                [MouseOperations]::MouseUp()
            }
            Write-Host "`r[$(Get-Date -Format 'HH:mm:ss')] Drawing interrupted by user at point $pointCount                    " -ForegroundColor Yellow
            return $false
        }
    }
    
    # Make sure pen is up at the end
    if ($isPenDown -and $withClicks) {
        [MouseOperations]::MouseUp()
        Write-Host "`r[$(Get-Date -Format 'HH:mm:ss')] Pen lifted - Drawing complete!                    " -ForegroundColor Green
    }
    
    return $true
}

while ($true) {
    $currentPos = [System.Windows.Forms.Cursor]::Position
    
    # Check if mouse has moved (user activity detected)
    if ($currentPos.X -ne $lastMousePos.X -or $currentPos.Y -ne $lastMousePos.Y) {
        $timeSinceLastActivity = ((Get-Date) - $lastActivityTime).TotalMilliseconds
        
        if ($timeSinceLastActivity -gt 500 -or -not $isDrawing) {
            $lastActivityTime = Get-Date
            
            if ($isDrawing) {
                $isDrawing = $false
                $drawCount = 0
                Write-Host "`r[$(Get-Date -Format 'HH:mm:ss')] User activity detected - Drawing stopped                    " -ForegroundColor Yellow
                Write-Host ""
            }
        }
        
        $lastMousePos = $currentPos
    }
    
    # Calculate idle time
    $idleSeconds = ((Get-Date) - $lastActivityTime).TotalSeconds
    
    # Start drawing if idle threshold exceeded
    if ($idleSeconds -ge $idleThresholdSeconds -and -not $isDrawing) {
        $isDrawing = $true
        Write-Host "`r[$(Get-Date -Format 'HH:mm:ss')] Idle detected ($([Math]::Floor($idleSeconds))s) - Starting signature drawing       " -ForegroundColor Green
        Write-Host ""
    }
    
    # Perform drawing if active
    if ($isDrawing) {
        $success = Draw-GregarotarSignature -withClicks $enableMouseClicks
        
        if ($success) {
            $drawCount++
            $completeMsg = if ($enableMouseClicks) { "[DRAWN]" } else { "[COMPLETE]" }
            Write-Host "`r[$(Get-Date -Format 'HH:mm:ss')] $completeMsg Signature #$drawCount complete! Repeating in 3s...    " -ForegroundColor Cyan
            Write-Host ""
            
            # ScrollLock keepalive
            $scrollLockTimer += 1
            if ($scrollLockTimer -ge ($scrollLockInterval / 3)) {
                $wshell.sendkeys("{SCROLLLOCK}{SCROLLLOCK}")
                $scrollLockTimer = 0
            }
            
            Start-Sleep -Milliseconds 3000
        } else {
            # User interrupted
            $isDrawing = $false
            $lastActivityTime = Get-Date
        }
    }
    else {
        # When not drawing, monitor for idle timeout
        $idleDisplay = [Math]::Floor($idleSeconds)
        $remaining = [Math]::Max(0, $idleThresholdSeconds - $idleDisplay)
        Write-Host "`r[$(Get-Date -Format 'HH:mm:ss')] MONITORING | Idle: $idleDisplay s | Drawing starts in: $remaining s    " -NoNewline -ForegroundColor Gray
        Start-Sleep -Milliseconds 200
    }
}
