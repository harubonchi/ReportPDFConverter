$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Resolve-Path (Join-Path $scriptDir "..")
Set-Location $projectRoot

$cargo = Join-Path $env:USERPROFILE ".cargo\bin\cargo.exe"
if (-not (Test-Path $cargo)) {
    throw "cargo.exe not found: $cargo"
}

Write-Host "[1/3] Building release binary..."
$vcvarsCandidates = @(
    "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat",
    "C:\Program Files\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat"
)
$vcvars = $vcvarsCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1

if ($vcvars) {
    cmd /c "`"$vcvars`" && `"$cargo`" build --release"
    if ($LASTEXITCODE -ne 0) {
        throw "cargo build failed (vcvars mode)"
    }
}
else {
    & $cargo build --release
    if ($LASTEXITCODE -ne 0) {
        throw "cargo build failed"
    }
}

$isccCandidates = @(
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    "C:\Program Files\Inno Setup 6\ISCC.exe",
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
)

$iscc = $isccCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) {
    throw "ISCC.exe not found. Install Inno Setup 6."
}

$iss = Join-Path $projectRoot "installer\ReportPDFConverter.iss"
if (-not (Test-Path $iss)) {
    throw "Installer script not found: $iss"
}

Write-Host "[2/3] Building installer with Inno Setup..."
& $iscc $iss
if ($LASTEXITCODE -ne 0) {
    throw "ISCC build failed"
}

Write-Host "[3/3] Done"
Write-Host "Output: $projectRoot\dist\installer\ReportPDFConverter-Setup.exe"
