param(
    [string]$Branch,
    [switch]$SkipOrigin,
    [switch]$SkipGithub
)

$ErrorActionPreference = "Stop"

$repoRoot = (git rev-parse --show-toplevel).Trim()
if (-not $repoRoot) {
    throw "Could not determine git repository root."
}

Set-Location $repoRoot

if (-not $Branch) {
    $Branch = (git branch --show-current).Trim()
}

if (-not $Branch) {
    throw "Could not determine current branch. Pass -Branch explicitly."
}

$prefix = "user/annolan/anythingReport"

Write-Host "Repo:    $repoRoot"
Write-Host "Branch:  $Branch"
Write-Host "Prefix:  $prefix"
Write-Host ""

if (-not $SkipOrigin) {
    Write-Host "Pushing branch to origin..."
    $env:ANYTHINGREPORT_SYNC_RUNNING = "1"
    git push origin "$Branch"
    Remove-Item Env:ANYTHINGREPORT_SYNC_RUNNING -ErrorAction SilentlyContinue
    if ($LASTEXITCODE -ne 0) {
        throw "Push to origin failed."
    }
}

if (-not $SkipGithub) {
    Write-Host "Mirroring anythingReport subtree to github/main..."
    $splitSha = (git subtree split --prefix="$prefix" HEAD).Trim()
    if (-not $splitSha) {
        throw "Failed to generate subtree split SHA."
    }

    # Force update is intentional: github/main is treated as a mirror of this subtree.
    git push --force github "$splitSha`:main"
    if ($LASTEXITCODE -ne 0) {
        throw "Mirror push to github/main failed."
    }
}

Write-Host ""
Write-Host "Done."
