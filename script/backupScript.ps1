
param(
    [string]$WsName # 워크스테이션 이름을 받는 파라미터 추가
)

# 필요한 네임스페이스 로드 (기존 코드 유지)
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# GUI 요소 생성 (기존 코드 유지)
$form = New-Object System.Windows.Forms.Form
$form.Text = "이벤트 로그 백업"
$form.Width = 450
$form.Height = 550

$startButton = New-Object System.Windows.Forms.Button
$startButton.Location = New-Object System.Drawing.Point(100, 20)
$startButton.Size = New-Object System.Drawing.Size(100, 30)
$startButton.Text = "백업 시작"
$form.Controls.Add($startButton)

$majorLogButton = New-Object System.Windows.Forms.Button
$majorLogButton.Location = New-Object System.Drawing.Point(210, 20)
$majorLogButton.Size = New-Object System.Drawing.Size(80, 30)
$majorLogButton.Text = "주요 로그"
$form.Controls.Add($majorLogButton)

$allLogButton = New-Object System.Windows.Forms.Button
$allLogButton.Location = New-Object System.Drawing.Point(300, 20)
$allLogButton.Size = New-Object System.Drawing.Size(80, 30)
$allLogButton.Text = "전체 로그"
$form.Controls.Add($allLogButton)

$logListLabel = New-Object System.Windows.Forms.Label
$logListLabel.Location = New-Object System.Drawing.Point(20, 70)
$logListLabel.AutoSize = $true
$logListLabel.Text = "백업할 로그 선택:"
$form.Controls.Add($logListLabel)

$logCheckBoxList = New-Object System.Windows.Forms.CheckedListBox
$logCheckBoxList.Location = New-Object System.Drawing.Point(20, 90)
$logCheckBoxList.Size = New-Object System.Drawing.Size(400, 150)
$form.Controls.Add($logCheckBoxList)

$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Location = New-Object System.Drawing.Point(20, 250)
$progressLabel.AutoSize = $true
$progressLabel.Text = "진행 상태:"
$form.Controls.Add($progressLabel)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 270)
$progressBar.Size = New-Object System.Drawing.Size(400, 20)
$form.Controls.Add($progressBar)

$savedLogListLabel = New-Object System.Windows.Forms.Label
$savedLogListLabel.Location = New-Object System.Drawing.Point(20, 300)
$savedLogListLabel.AutoSize = $true
$savedLogListLabel.Text = "저장된 로그:"
$form.Controls.Add($savedLogListLabel)

$logListView = New-Object System.Windows.Forms.ListBox
$logListView.Location = New-Object System.Drawing.Point(20, 320)
$logListView.Size = New-Object System.Drawing.Size(400, 150)
$form.Controls.Add($logListView)

# 백업 폴더 경로 생성 함수 (수정됨)
function Get-BackupFolderPath {
    param(
        [Parameter(Mandatory=$true)]
        [string]$WsName # 파라미터로 워크스테이션 이름 받음
    )
    $scriptPath = ""
    if ($MyInvocation.MyCommand.Path) {
        $scriptPath = Split-Path -parent $MyInvocation.MyCommand.Path
    } else {
        # 스크립트가 파일로 실행되지 않은 경우 현재 작업 디렉토리를 사용
        $scriptPath = Get-Location
    }
    $logDir = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "log"
    if (-not (Test-Path $logDir -PathType Container)) {
        try {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
            Write-Host "Log 디렉토리를 생성했습니다: $logDir"
        } catch {
            Write-Error "Log 디렉토리 생성 실패: $($_.Exception.Message)"
            return $null
        }
    }
    $wsLogDir = Join-Path -Path $logDir -ChildPath $WsName
    if (-not (Test-Path $wsLogDir -PathType Container)) {
        try {
            New-Item -ItemType Directory -Path $wsLogDir -Force | Out-Null
            Write-Host "워크스테이션 로그 디렉토리를 생성했습니다: $wsLogDir"
        } catch {
            Write-Error "워크스테이션 로그 디렉토리 생성 실패: $($_.Exception.Message)"
            return $null
        }
    }
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupPath = Join-Path -Path $wsLogDir -ChildPath $timestamp
    try {
        New-Item -ItemType Directory -Path $backupPath -Force | Out-Null
        Write-Host "백업 디렉토리를 생성했습니다: $backupPath"
        return $backupPath
    } catch {
        Write-Error "백업 디렉토리 생성 실패: $($_.Exception.Message)"
        return $null
    }
}

# 이벤트 로그 백업 함수 (일반 실행) (수정됨)
function Backup-SelectedLogs {
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.CheckedListBox]$LogCheckBoxList,
        [Parameter(Mandatory=$true)]
        [string]$BackupPath
    )
    if ([string]::IsNullOrEmpty($BackupPath)) {
        Write-Error "백업 경로가 유효하지 않습니다."
        return $false
    }
    $totalLogs = $LogCheckBoxList.CheckedItems.Count
    $processedLogs = 0

    foreach ($checkedItem in $LogCheckBoxList.CheckedItems) {
        $displayName = ""
        $realName = ""

        # 정규 표현식을 사용하여 괄호 안의 실제 이름 추출 (기존 코드 유지)
        if ($checkedItem -match '\(([^)]+)\)') {
            $realName = $Matches[1].Trim()
            $displayName = ($checkedItem -replace '\s*\([^)]+\)$', '').Trim()
            if ([string]::IsNullOrWhiteSpace($displayName)) {
                $displayName = $realName
            }
        } else {
            $realName = $checkedItem.Trim()
            $displayName = $realName
        }

        if (-not [string]::IsNullOrEmpty($realName)) {
            $fileName = "$realName" + "_" + (Get-Date -Format "yyyyMMdd") + ".evtx"
            $fullPath = Join-Path -Path $BackupPath -ChildPath $fileName
            try {
                wevtutil export-log $realName $fullPath /ow:true
                Write-Host "$displayName ($realName) 로그를 $fullPath 에 백업했습니다."
                $logListView.Items.Add("$displayName ($realName) - 백업 완료: $($fullPath.Split('\')[-1])") # ListBox 업데이트
            } catch {
                Write-Error "Backup-SelectedLogs 오류 ($realName): $($_.Exception.Message)"
                $logListView.Items.Add("$displayName ($realName) - 백업 실패: $($_.Exception.Message)") # ListBox 업데이트
                # Backup-SelectedLogs 내부 오류 발생 시에도 진행을 멈추지 않도록 수정
            }
        } else {
            Write-Warning "선택된 로그 항목에서 실제 이름을 찾을 수 없습니다: '$checkedItem'"
        }

        $processedLogs++
        if ($totalLogs -gt 0) {
            $progressPercentage = [int](($processedLogs / $totalLogs) * 100)
            $progressBar.Value = $progressPercentage
            $progressLabel.Text = "진행 상태: $displayName 백업 중... ($progressPercentage%)"
            [System.Windows.Forms.Application]::DoEvents() # GUI 업데이트
        }
    }
    return $true # 정상 종료 시 true 반환
}

# 폼 로드 시 이벤트 로그 목록을 CheckBoxList에 채우기 (기존 코드 유지)
$form.Add_Shown({
    Get-WinEvent -ListLog * | ForEach-Object {
        try {
            $displayName = $_.LogDisplayName
            $logName = $_.LogName
            if (-not [string]::IsNullOrWhiteSpace($displayName)) {
                $logCheckBoxList.Items.Add("$displayName ($logName)", $true)
            } else {
                $logCheckBoxList.Items.Add("($logName)", $true)
            }
        } catch {
            # 오류 발생 시 아무 동작도 하지 않음
        }
    }
    # Update-LogList # 이 함수는 정의되어 있지 않습니다.
})

# 주요 로그 버튼 클릭 이벤트 핸들러 (기존 코드 유지)
$majorLogButton.Add_Click({
    $indicesToSet = @()
    for ($i = 0; $i -lt $logCheckBoxList.Items.Count; $i++) {
        $item = $logCheckBoxList.Items[$i]
        $logNameMatch = $item -match '\(([^)]+)\)'
        if ($logNameMatch) {
            $realName = $Matches[1].Trim()
            if ($realName -ceq "Application" -or $realName -ceq "System" -or $realName -ceq "Security") {
                $indicesToSet += $i
            }
        }
    }
    for ($i = 0; $i -lt $logCheckBoxList.Items.Count; $i++) {
        $logCheckBoxList.SetItemChecked($i, $false)
    }
    foreach ($index in $indicesToSet) {
        $logCheckBoxList.SetItemChecked($index, $true)
    }
})

# 전체 로그 버튼 클릭 이벤트 핸들러 (기존 코드 유지)
$allLogButton.Add_Click({
    for ($i = 0; $i -lt $logCheckBoxList.Items.Count; $i++) {
        $logCheckBoxList.SetItemChecked($i, $true)
    }
})

# 백업 시작 버튼 클릭 이벤트 핸들러 (수정됨)
$startButton.Add_Click({
    $startButton.Enabled = $false
    $progressBar.Value = 0
    $progressLabel.Text = "진행 상태: 백업 준비 중..."

    if (-not [string]::IsNullOrEmpty($WsName)) { # 파라미터로 받은 WsName 사용
        $backupPath = Get-BackupFolderPath -WsName $WsName
        if ($backupPath) {
            try {
                $logListView.Items.Clear()
                $result = Backup-SelectedLogs -LogCheckBoxList $logCheckBoxList -BackupPath $backupPath
                if ($result -eq $true) {
                    [System.Windows.Forms.MessageBox]::Show("로그 백업이 완료되었습니다. 저장 위치: $backupPath", "완료", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                } else {
                    [System.Windows.Forms.MessageBox]::Show("백업 중 오류가 발생했습니다. 자세한 내용은 로그를 확인하세요.", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            } catch {
                [System.Windows.Forms.MessageBox]::Show("로그 백업 프로세스 중 예상치 못한 오류가 발생했습니다: $($_.Exception.Message)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            } finally {
                $startButton.Enabled = $true
                $progressBar.Value = 100
                $progressLabel.Text = "진행 상태: 완료"
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("백업 폴더를 생성하지 못했습니다.", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $startButton.Enabled = $true
            $progressBar.Value = 0
            $progressLabel.Text = "진행 상태: 오류"
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("워크스테이션 이름이 전달되지 않았습니다.", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $startButton.Enabled = $true
        $progressBar.Value = 0
        $progressLabel.Text = "진행 상태: 오류"
    }
})

# GUI 표시 (기존 코드 유지)
[void]$form.ShowDialog()