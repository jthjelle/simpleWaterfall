Add-Type -AssemblyName System.Drawing
$src = "C:/Users/jon/.gemini/antigravity/brain/30a01808-7bc5-4e83-b540-a42bbf0dfcce/uploaded_media_1769550707951.png"
$dest = "C:\Users\jon\.gemini\antigravity\scratch\simpleWaterfall\assets\icon.png"

try {
    $img = [System.Drawing.Image]::FromFile($src)
    $canvas = New-Object System.Drawing.Bitmap(20, 20)
    $graph = [System.Drawing.Graphics]::FromImage($canvas)
    # High quality resizing
    $graph.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graph.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $graph.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    
    $graph.DrawImage($img, 0, 0, 20, 20)
    $canvas.Save($dest, [System.Drawing.Imaging.ImageFormat]::Png)
    
    Write-Host "Success: Image resized to 20x20 at $dest"
} catch {
    Write-Error "Failed to resize image: $_"
} finally {
    if ($img) { $img.Dispose() }
    if ($canvas) { $canvas.Dispose() }
    if ($graph) { $graph.Dispose() }
}
