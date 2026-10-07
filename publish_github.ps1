# Run from the extracted source folder on Windows with Git and GitHub CLI installed.
# Uploads to the user-provided empty repository; never forces or overwrites a branch.
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
$Repository = 'Raman0925/Edoc-generator'
$RemoteUrl = "https://github.com/$Repository.git"

function Invoke-Checked {
    param([string]$Program, [string[]]$Arguments)
    & $Program @Arguments
    if ($LASTEXITCODE -ne 0) { throw "$Program failed (exit $LASTEXITCODE)." }
}

try {
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) { throw 'Install Git from git-scm.com first.' }
    if (-not (Get-Command gh -ErrorAction SilentlyContinue)) { throw 'Install GitHub CLI from cli.github.com first.' }
    & gh auth status
    if ($LASTEXITCODE -ne 0) {
        Invoke-Checked gh @('auth', 'login', '--hostname', 'github.com', '--git-protocol', 'https', '--web')
    }
    Invoke-Checked gh @('auth', 'setup-git')
    $Python = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
    if (-not (Test-Path $Python)) { throw 'Run setup_windows.bat first.' }
    Invoke-Checked -Program $Python -Arguments @('-m', 'unittest', 'discover', '-s', 'tests', '-v')

    # The package is meant for a new empty repo. Existing remote contents require
    # a deliberate merge by the user; do not replace them automatically.
    $remoteHeads = & git ls-remote --heads $RemoteUrl
    if ($LASTEXITCODE -ne 0) { throw 'Cannot access the destination repository.' }
    if ($remoteHeads) { throw 'Destination already has branches. Clone it and merge this project; this script will not overwrite it.' }

    if (Test-Path '.git') { throw 'Local folder already contains Git history. Publish it manually after reviewing the repository.' }
    Invoke-Checked git @('init', '--initial-branch=main')
    Invoke-Checked git @('add', '.gitignore', '.github', 'edoc', 'examples', 'tests', 'launcher.py',
        'requirements.txt', 'requirements-build.txt', 'setup_windows.bat', 'run_windows.bat',
        'build_windows.bat', 'publish_github.ps1', 'README.md', 'LEARNING.md', 'VALIDATION.md')
    Invoke-Checked git @('commit', '-m', 'Build configurable Excel approval and certificate generator')
    Invoke-Checked git @('remote', 'add', 'origin', $RemoteUrl)
    Invoke-Checked git @('push', '-u', 'origin', 'main')
    Write-Host "Published to https://github.com/$Repository"
    Write-Host 'Open Actions, wait for the Windows build, and download EDocGenerator-Windows.'
}
catch {
    Write-Error $_
    exit 1
}
