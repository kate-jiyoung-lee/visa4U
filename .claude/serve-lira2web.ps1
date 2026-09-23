param([int]$Port = 8735)
$ErrorActionPreference = 'Stop'
$root = Join-Path $PSScriptRoot '..\..\LIRA2\apps\web\dist'
$root = (Resolve-Path $root).Path
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Start()
Write-Host "Serving $root on http://localhost:$Port/"
while ($listener.IsListening) {
  $ctx = $listener.GetContext()
  $req = $ctx.Request
  $res = $ctx.Response
  $rel = $req.Url.AbsolutePath.TrimStart('/')
  if ([string]::IsNullOrEmpty($rel)) { $rel = 'index.html' }
  $path = Join-Path $root $rel
  try {
    if (Test-Path $path -PathType Leaf) {
      $bytes = [System.IO.File]::ReadAllBytes($path)
      if ($path -match '\.html?$') { $res.ContentType = 'text/html; charset=utf-8' }
      elseif ($path -match '\.js$') { $res.ContentType = 'application/javascript; charset=utf-8' }
      elseif ($path -match '\.css$') { $res.ContentType = 'text/css; charset=utf-8' }
      elseif ($path -match '\.json$') { $res.ContentType = 'application/json; charset=utf-8' }
      elseif ($path -match '\.svg$') { $res.ContentType = 'image/svg+xml' }
      $res.OutputStream.Write($bytes, 0, $bytes.Length)
    } else {
      $res.StatusCode = 404
      $b = [System.Text.Encoding]::UTF8.GetBytes('not found')
      $res.OutputStream.Write($b, 0, $b.Length)
    }
  } catch {
    $res.StatusCode = 500
  } finally {
    $res.OutputStream.Close()
  }
}
