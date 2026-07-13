# Verify that a published Windows OfficeCLI PE imports only Windows system DLLs.
# This is a read-only release gate and does not copy or rewrite dependencies.
# Exit codes: 2 invalid input, 3 inspection unavailable/failed, 4 non-portable.

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$BinaryPath,
    [switch]$SelfTest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Stop-WithCode {
    param([int]$Code, [string]$Message)
    [Console]::Error.WriteLine("ERROR: $Message")
    exit $Code
}

function Find-Dumpbin {
    $command = Get-Command dumpbin.exe -ErrorAction SilentlyContinue
    if ($null -ne $command) {
        return $command.Source
    }

    $programFilesX86 = [Environment]::GetFolderPath('ProgramFilesX86')
    $vswhere = Join-Path $programFilesX86 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (-not (Test-Path -LiteralPath $vswhere -PathType Leaf)) {
        return $null
    }

    $matches = @(
        & $vswhere -latest -products * `
            -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 `
            -find 'VC\Tools\MSVC\**\bin\Hostx64\x64\dumpbin.exe'
    )
    $candidate = $matches |
        Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } |
        Select-Object -Last 1
    return $candidate
}

# Baseline from the official .NET 10.0.9 win-x64/win-arm64 singlefilehost.
# API Set names are Windows system contracts, not app DLLs.
$allowedDlls = @(
    'ADVAPI32.dll',
    'KERNEL32.dll',
    'ole32.dll',
    'OLEAUT32.dll',
    'SHELL32.dll',
    'USER32.dll'
)

function Get-ImportedDlls {
    param([string]$DependentsOutput)
    return @(
        $DependentsOutput -split "\r?\n" |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -match '(?i)\.dll$' } |
            Sort-Object -Unique
    )
}

function Get-DependencyViolations {
    param([string[]]$Dependencies)
    return @(
        $Dependencies | Where-Object {
            ($allowedDlls -notcontains $_) -and
            ($_ -notmatch '(?i)^(api|ext)-ms-win-[a-z0-9-]+\.dll$')
        }
    )
}

if ($SelfTest) {
    $fixture = "KERNEL32.dll`napi-ms-win-crt-runtime-l1-1-0.dll`nlibcrypto-3-x64.dll"
    $fixtureViolations = @(Get-DependencyViolations (Get-ImportedDlls $fixture))
    if (($fixtureViolations.Count -ne 1) -or ($fixtureViolations[0] -ne 'libcrypto-3-x64.dll')) {
        throw "Windows portability parser self-test failed: $($fixtureViolations -join ', ')"
    }
    Write-Host 'Windows portability parser self-test passed'
    exit 0
}

if ([string]::IsNullOrWhiteSpace($BinaryPath)) {
    Stop-WithCode 2 'Usage: verify-windows-portability.ps1 <pe-binary>'
}
if (-not (Test-Path -LiteralPath $BinaryPath -PathType Leaf)) {
    Stop-WithCode 2 "Windows portability input does not exist: $BinaryPath"
}

$resolvedBinary = (Resolve-Path -LiteralPath $BinaryPath).Path
$dumpbin = Find-Dumpbin
if (-not $dumpbin) {
    Stop-WithCode 3 'Windows portability inspection requires Visual Studio dumpbin.exe'
}

try {
    $headers = (& $dumpbin /nologo /headers $resolvedBinary 2>&1 | Out-String)
    if ($LASTEXITCODE -ne 0) {
        Stop-WithCode 2 "Windows portability input is not a valid PE file: $resolvedBinary`n$headers"
    }

    $dependents = (& $dumpbin /nologo /dependents $resolvedBinary 2>&1 | Out-String)
    if ($LASTEXITCODE -ne 0) {
        Stop-WithCode 3 "dumpbin /dependents failed for $resolvedBinary`n$dependents"
    }
}
catch {
    Stop-WithCode 3 "dumpbin inspection failed for ${resolvedBinary}: $($_.Exception.Message)"
}

$dependencies = @(Get-ImportedDlls $dependents)
if ($dependencies.Count -eq 0) {
    Stop-WithCode 3 "dumpbin returned no PE dependencies for $resolvedBinary"
}

$violations = @(Get-DependencyViolations $dependencies)

if ($violations.Count -ne 0) {
    [Console]::Error.WriteLine("ERROR: non-portable Windows release artifact: $resolvedBinary")
    foreach ($dependency in $violations) {
        [Console]::Error.WriteLine("  - dependency: $dependency")
    }
    [Console]::Error.WriteLine(
        'Only audited Windows system DLLs and api-ms-win/ext-ms-win API Sets are allowed; publish with the official .NET SDK/CI.'
    )
    exit 4
}

Write-Host "Windows portability check passed: $resolvedBinary"
