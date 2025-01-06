param(
    [string]$RepoPath,
    [string]$UserName,
    [string]$PatToken,
    [string]$RepoURL
)

$LogFile = Join-Path $RepoPath "auto_push.log"
$CommitMessage = "Auto-commit on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

# ???亥?
Add-Content -Path $LogFile -Value "---- Start Push at $(Get-Date) ----"

Set-Location -Path $RepoPath

# 閮剖??垢 URL 撣?PAT
$tokenURL = $RepoURL -replace "https://", "https://$($PatToken)@"
git remote set-url origin $tokenURL

try {
    git pull origin main --rebase
    Add-Content -Path $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Pulled latest changes."
}
catch {
    Add-Content -Path $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Error pulling changes: $_"
    exit
}

git add .
$commitResult = git commit -m "$CommitMessage"

if ($commitResult -match "nothing to commit") {
    Add-Content -Path $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - No changes to commit."
    exit
}

try {
    git push origin main
    Add-Content -Path $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Pushed changes successfully."
}
catch {
    Add-Content -Path $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Error pushing changes: $_"
}

# 蝯??亥?
Add-Content -Path $LogFile -Value "---- End Push ----`n"
