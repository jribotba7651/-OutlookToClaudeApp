Add-Type -AssemblyName System.Drawing

$pngPath = "C:\Users\juanr\Desktop\OutlookToClaudeApp\icon.png"
$icoPath = "C:\Users\juanr\Desktop\OutlookToClaudeApp\icon.ico"

# Load the PNG image
$png = [System.Drawing.Image]::FromFile($pngPath)

# Create a bitmap and resize if needed (ICO works best with 256x256 or smaller)
$bitmap = New-Object System.Drawing.Bitmap $png

# Create icon from bitmap
$icon = [System.Drawing.Icon]::FromHandle($bitmap.GetHicon())

# Save as ICO
$stream = [System.IO.File]::Create($icoPath)
$icon.Save($stream)
$stream.Close()

# Clean up
$icon.Dispose()
$bitmap.Dispose()
$png.Dispose()

Write-Host "Icon converted successfully to $icoPath"
