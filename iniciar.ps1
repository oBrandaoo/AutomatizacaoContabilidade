$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$taskPython = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $taskPython)) {
    Write-Host 'Preparando o ambiente Python...'
    python -m venv .venv
    if ($LASTEXITCODE -ne 0) { throw 'Instale o Python 3.11 ou superior e tente novamente.' }
}
$taskRequirements = Join-Path $PSScriptRoot 'requirements.txt'
$taskStamp = Join-Path $PSScriptRoot '.venv\notas-requirements.sha256'
$taskHash = (Get-FileHash -LiteralPath $taskRequirements -Algorithm SHA256).Hash
$taskInstalled = if (Test-Path -LiteralPath $taskStamp) { (Get-Content -LiteralPath $taskStamp -Raw).Trim() } else { '' }
if ($taskInstalled -ne $taskHash) {
    Write-Host 'Instalando as dependencias (requer internet no primeiro uso)...'
    & $taskPython -m pip install -r $taskRequirements
    if ($LASTEXITCODE -ne 0) { throw 'Nao foi possivel instalar as dependencias.' }
    Set-Content -LiteralPath $taskStamp -Value $taskHash
}
Write-Host 'Abra http://127.0.0.1:8000 no navegador. Mantenha esta janela aberta.'
& $taskPython (Join-Path $PSScriptRoot 'app.py')
