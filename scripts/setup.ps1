$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $projectRoot
$runtimePath = Join-Path $projectRoot '.runtime'
New-Item -ItemType Directory -Path $runtimePath -Force | Out-Null
if (-not (Test-Path -LiteralPath (Join-Path $runtimePath 'python.exe'))) {
    Write-Host 'Downloading portable Python 3.12.10 from python.org...'
    Invoke-WebRequest -UseBasicParsing -Uri 'https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip' -OutFile (Join-Path $runtimePath 'python.zip')
    Expand-Archive -LiteralPath (Join-Path $runtimePath 'python.zip') -DestinationPath $runtimePath -Force
}
Set-Content -LiteralPath (Join-Path $runtimePath 'python312._pth') -Encoding ASCII -Value @('python312.zip', '.', '..', 'Lib\site-packages', 'import site')
$pythonPath = Join-Path $runtimePath 'python.exe'
if (-not (Test-Path -LiteralPath (Join-Path $runtimePath 'Lib\site-packages\pip\__init__.py'))) {
    Invoke-WebRequest -UseBasicParsing -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile (Join-Path $runtimePath 'get-pip.py')
    & $pythonPath (Join-Path $runtimePath 'get-pip.py') --no-warn-script-location
    if ($LASTEXITCODE -ne 0) { throw 'pip installation failed.' }
}
$requirementsPath = if (Test-Path -LiteralPath 'requirements-lock.txt') { 'requirements-lock.txt' } else { 'requirements.txt' }
& $pythonPath -m pip install --no-warn-script-location --disable-pip-version-check -r $requirementsPath
if ($LASTEXITCODE -ne 0) { throw 'Dependency installation failed.' }
& $pythonPath scripts/make_test_data.py
if ($LASTEXITCODE -ne 0) { throw 'Fixture generation failed.' }
Write-Host 'Ready. Run start.cmd and open http://127.0.0.1:8765'
