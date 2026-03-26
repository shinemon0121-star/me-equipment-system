$port = 8080
$root = "C:\Users\STH-ME001D\Desktop\ME機器管理システム_App"

$mimeTypes = @{
    '.html' = 'text/html; charset=utf-8'
    '.js'   = 'application/javascript; charset=utf-8'
    '.css'  = 'text/css; charset=utf-8'
    '.json' = 'application/json; charset=utf-8'
    '.png'  = 'image/png'
    '.jpg'  = 'image/jpeg'
    '.ico'  = 'image/x-icon'
    '.svg'  = 'image/svg+xml'
    '.woff' = 'font/woff'
    '.woff2'= 'font/woff2'
}

try {
    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://+:$port/")
    $listener.Start()
} catch {
    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://localhost:$port/")
    $listener.Start()
}

[Console]::Out.Flush()
Write-Host "Server running at http://localhost:$port/" -NoNewline
Write-Host ""
[Console]::Out.Flush()

while ($listener.IsListening) {
    try {
        $context = $listener.GetContext()
        $path = $context.Request.Url.LocalPath
        if ($path -eq '/') { $path = '/index.html' }

        $filePath = Join-Path $root $path.TrimStart('/')

        if (Test-Path $filePath -PathType Leaf) {
            $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
            $contentType = if ($mimeTypes.ContainsKey($ext)) { $mimeTypes[$ext] } else { 'application/octet-stream' }
            $context.Response.ContentType = $contentType
            $context.Response.Headers.Add("Access-Control-Allow-Origin", "*")
            $bytes = [System.IO.File]::ReadAllBytes($filePath)
            $context.Response.ContentLength64 = $bytes.Length
            $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
        } else {
            $context.Response.StatusCode = 404
            $msg = [System.Text.Encoding]::UTF8.GetBytes('Not Found')
            $context.Response.ContentLength64 = $msg.Length
            $context.Response.OutputStream.Write($msg, 0, $msg.Length)
        }
        $context.Response.Close()
    } catch {
        Write-Host "Error: $_"
    }
}
