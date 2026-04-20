$ErrorActionPreference = "Stop"

$appDir = $PSScriptRoot
if (-not $appDir) {
    $candidatePath = $MyInvocation.MyCommand.Path
    if ($candidatePath) {
        $appDir = Split-Path -Parent $candidatePath
    } else {
        $appDir = [System.IO.Path]::GetDirectoryName([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
    }
}
$pythonEmbedded = Join-Path $appDir "pyembed\python.exe"
$pythonHidden = Join-Path $appDir "pyembed\pythonw.exe"
$appPy = Join-Path $appDir "app.py"
$url = "http://127.0.0.1:8765"

function Test-PortOpen {
    param([int]$Port)
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $iar = $client.BeginConnect("127.0.0.1", $Port, $null, $null)
        $ok = $iar.AsyncWaitHandle.WaitOne(600)
        if (-not $ok) {
            $client.Close()
            return $false
        }
        $client.EndConnect($iar) | Out-Null
        $client.Close()
        return $true
    } catch {
        return $false
    }
}

if (-not (Test-PortOpen -Port 8765)) {
    if (Test-Path $pythonHidden) {
        Start-Process -FilePath $pythonHidden -ArgumentList @($appPy) -WorkingDirectory $appDir -WindowStyle Hidden
    } elseif (Test-Path $pythonEmbedded) {
        Start-Process -FilePath $pythonEmbedded -ArgumentList @($appPy) -WorkingDirectory $appDir -WindowStyle Hidden
    } else {
        Start-Process -FilePath "python" -ArgumentList @($appPy) -WorkingDirectory $appDir
    }

    for ($i = 0; $i -lt 30; $i++) {
        Start-Sleep -Milliseconds 400
        if (Test-PortOpen -Port 8765) { break }
    }
}

Start-Process $url
