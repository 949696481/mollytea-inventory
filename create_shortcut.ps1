$name = [char]0x5E93 + [char]0x5B58 + [char]0x767B + [char]0x8BB0
$WshShell = New-Object -ComObject WScript.Shell
$desktop = $WshShell.SpecialFolders("Desktop")
$tempPath = "D:\DevTools\InventoryApp\InventoryAppTemp.lnk"
$finalPath = Join-Path $desktop ($name + ".lnk")

$Shortcut = $WshShell.CreateShortcut($tempPath)
$Shortcut.TargetPath = "C:\Python314\pythonw.exe"
$Shortcut.Arguments = '"D:\DevTools\InventoryApp\launch_app.pyw"'
$Shortcut.WorkingDirectory = "D:\DevTools\InventoryApp"
$Shortcut.IconLocation = "D:\DevTools\InventoryApp\icon.ico"
$Shortcut.Description = $name
$Shortcut.Save()

if (Test-Path $finalPath) { Remove-Item $finalPath -Force }
Move-Item -LiteralPath $tempPath -Destination $finalPath -Force
Write-Output ("Shortcut created at " + $finalPath)
