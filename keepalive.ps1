Add-Type -AssemblyName System.Windows.Forms

Write-Host "Starting in 10 seconds... Press Ctrl+C to stop anytime" -ForegroundColor Green
Start-Sleep -Seconds 10

Write-Host "Constant jiggling active! Press Ctrl+C to stop." -ForegroundColor Yellow

$wshell = New-Object -com "Wscript.Shell"
$counter = 0

while ($true) {
  # Constant mouse jiggle
  $currentPos = [System.Windows.Forms.Cursor]::Position
  [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(($currentPos.X + 1), $currentPos.Y)
  Start-Sleep -Milliseconds 100
  [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(($currentPos.X - 1), $currentPos.Y)
  Start-Sleep -Milliseconds 100
  
  # Send ScrollLock every 30 seconds (150 iterations × 200ms)
  $counter++
  if ($counter -ge 150) {
    $wshell.sendkeys("{SCROLLLOCK}{SCROLLLOCK}")
    $counter = 0
  }
}
