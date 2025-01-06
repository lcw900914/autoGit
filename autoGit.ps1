[CmdletBinding()]
param()

### 載入必要的 .NET 組件 ###
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

### 建立主視窗 Form ###
$form = New-Object System.Windows.Forms.Form
$form.Text = "Auto Push to GitHub - Scheduler"
$form.Size = New-Object System.Drawing.Size(550, 400)
$form.StartPosition = "CenterScreen"

### 建立輸入元件 (Label、TextBox、ComboBox、Button) ###

# GitHub Username
$labelUser = New-Object System.Windows.Forms.Label
$labelUser.Text = "GitHub Username:"
$labelUser.Location = New-Object System.Drawing.Point(30,30)
$labelUser.Size = New-Object System.Drawing.Size(120,25)
$form.Controls.Add($labelUser)

$textUser = New-Object System.Windows.Forms.TextBox
$textUser.Location = New-Object System.Drawing.Point(160,28)
$textUser.Size = New-Object System.Drawing.Size(330,25)
$form.Controls.Add($textUser)

# GitHub PAT
$labelPAT = New-Object System.Windows.Forms.Label
$labelPAT.Text = "GitHub PAT:"
$labelPAT.Location = New-Object System.Drawing.Point(30,70)
$labelPAT.Size = New-Object System.Drawing.Size(120,25)
$form.Controls.Add($labelPAT)

$textPAT = New-Object System.Windows.Forms.TextBox
$textPAT.Location = New-Object System.Drawing.Point(160,68)
$textPAT.Size = New-Object System.Drawing.Size(330,25)
$textPAT.UseSystemPasswordChar = $true  # 顯示為●
$form.Controls.Add($textPAT)

# GitHub Repo URL
$labelRepo = New-Object System.Windows.Forms.Label
$labelRepo.Text = "GitHub Repo URL:"
$labelRepo.Location = New-Object System.Drawing.Point(30,110)
$labelRepo.Size = New-Object System.Drawing.Size(120,25)
$form.Controls.Add($labelRepo)

$textRepo = New-Object System.Windows.Forms.TextBox
$textRepo.Location = New-Object System.Drawing.Point(160,108)
$textRepo.Size = New-Object System.Drawing.Size(330,25)
$form.Controls.Add($textRepo)

# Local Folder (where script & files are)
$labelFolder = New-Object System.Windows.Forms.Label
$labelFolder.Text = "Local Folder:"
$labelFolder.Location = New-Object System.Drawing.Point(30,150)
$labelFolder.Size = New-Object System.Drawing.Size(120,25)
$form.Controls.Add($labelFolder)

$textFolder = New-Object System.Windows.Forms.TextBox
$textFolder.Location = New-Object System.Drawing.Point(160,148)
$textFolder.Size = New-Object System.Drawing.Size(240,25)
$form.Controls.Add($textFolder)

# Button to browse folder
$btnFolderSelect = New-Object System.Windows.Forms.Button
$btnFolderSelect.Text = "Browse..."
$btnFolderSelect.Location = New-Object System.Drawing.Point(410,148)
$btnFolderSelect.Size = New-Object System.Drawing.Size(80,25)
$form.Controls.Add($btnFolderSelect)

$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$btnFolderSelect.Add_Click({
    if($folderBrowser.ShowDialog() -eq "OK") {
        $textFolder.Text = $folderBrowser.SelectedPath
    }
})

# Execution Frequency
$labelFreq = New-Object System.Windows.Forms.Label
$labelFreq.Text = "Execution Frequency:"
$labelFreq.Location = New-Object System.Drawing.Point(30,190)
$labelFreq.Size = New-Object System.Drawing.Size(130,25)
$form.Controls.Add($labelFreq)

$comboFreq = New-Object System.Windows.Forms.ComboBox
$comboFreq.Location = New-Object System.Drawing.Point(160,188)
$comboFreq.Size = New-Object System.Drawing.Size(180,25)
$comboFreq.DropDownStyle = 'DropDownList'
$comboFreq.Items.Add("Every 5 minutes (/SC MINUTE /MO 5)")
$comboFreq.Items.Add("Every 15 minutes (/SC MINUTE /MO 15)")
$comboFreq.Items.Add("Hourly (/SC HOURLY /MO 1)")
$comboFreq.Items.Add("Daily (/SC DAILY /ST 00:00)")
$comboFreq.SelectedIndex = 0
$form.Controls.Add($comboFreq)

# Create Schedule button
$btnCreate = New-Object System.Windows.Forms.Button
$btnCreate.Text = "Create Schedule"
$btnCreate.Location = New-Object System.Drawing.Point(160,230)
$btnCreate.Size = New-Object System.Drawing.Size(120,40)
$form.Controls.Add($btnCreate)

# Status label
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Text = ""
$labelStatus.Location = New-Object System.Drawing.Point(30,290)
$labelStatus.Size = New-Object System.Drawing.Size(500,25)
$form.Controls.Add($labelStatus)

### 建立排程按鈕的事件 ###

$btnCreate.Add_Click({

    # 取得表單輸入
    $userName   = $textUser.Text.Trim()
    $patToken   = $textPAT.Text.Trim()
    $repoURL    = $textRepo.Text.Trim()
    $folderPath = $textFolder.Text.Trim()
    if (-not $folderPath) {
        # 若使用者沒選資料夾，就預設 C:\AutoGit
        $folderPath = "C:\AutoGit"
    }

    # 頻率處理
    $freqSelected = $comboFreq.SelectedItem
    switch -Wildcard ($freqSelected) {
        "*Every 5 minutes*"  { $trigger = "/SC MINUTE /MO 5" }
        "*Every 15 minutes*" { $trigger = "/SC MINUTE /MO 15" }
        "*Hourly*"           { $trigger = "/SC HOURLY /MO 1" }
        "*Daily*"            { $trigger = "/SC DAILY /ST 00:00" }
        default              { $trigger = "/SC MINUTE /MO 5" }
    }

    # 基本檢查
    if (-not $userName -or -not $patToken -or -not $repoURL) {
        $labelStatus.Text = "Please fill in Username, PAT, and Repo URL."
        return
    }

    # 檢查是否有 Git 安裝
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        $labelStatus.Text = "Git is not installed or not in PATH."
        return
    }

    # 產生 auto_push.ps1
    try {
        # 1) 創建資料夾（如果不存在）
        if (!(Test-Path $folderPath)) {
            New-Item -ItemType Directory -Path $folderPath | Out-Null
        }

        # 2) 設定 auto_push.ps1 路徑
        $scriptPath = Join-Path $folderPath "auto_push.ps1"

        # 3) 如果 auto_push.ps1 不存在，則創建
        if (!(Test-Path $scriptPath)) {
            $autoPushContent = @'
param(
    [string]$RepoPath,
    [string]$UserName,
    [string]$PatToken,
    [string]$RepoURL
)

$LogFile = Join-Path $RepoPath "auto_push.log"
$CommitMessage = "Auto-commit on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

# 開始日誌
Add-Content -Path $LogFile -Value "---- Start Push at $(Get-Date) ----"

Set-Location -Path $RepoPath

# 設定遠端 URL 帶 PAT
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

# 結束日誌
Add-Content -Path $LogFile -Value "---- End Push ----`n"
'@

            # 將內容寫入 auto_push.ps1
            Set-Content -Path $scriptPath -Value $autoPushContent -Encoding UTF8
        }

        # 4) 準備各參數
        $taskName   = "AutoGitPushTask"
        $psExePath  = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
        $repoPath   = $folderPath
        $userName   = $userName
        $patToken   = $patToken
        $repoUrl    = $repoURL

        # 5) 組合要執行的指令 (innerCommand)
        #    雙引號要用 反引號 ` 來轉義
        $innerCommand = "`"$psExePath`" -ExecutionPolicy Bypass -File `"$scriptPath`" -RepoPath `"$repoPath`" -UserName `"$userName`" -PatToken `"$patToken`" -RepoURL `"$repoUrl`""

        # 6) 再把 $innerCommand 用雙引號包起來供 /TR 使用
        $trValue = "`"$innerCommand`""

        # 7) 組合最終指令
        $taskCmd = "SCHTASKS /Create /TN `"$taskName`" /TR $trValue $trigger /F /RL HIGHEST"

        # Debug 檢查
        Write-Host "Final CMD: $taskCmd"

        # 執行 SCHTASKS
        Invoke-Expression $taskCmd

        # 成功訊息
        $labelStatus.Text = "Scheduled task created successfully: $taskName (Trigger: $freqSelected)"
    }
    catch {
        $labelStatus.Text = "Error: $($_.Exception.Message)"
    }

})

### 顯示視窗，讓使用者互動 ###
[void]$form.ShowDialog()
