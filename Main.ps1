Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 버전 정보 설정
$programName = "ISSO_WIN_SCRIPT"
$version = "1.0"

# 폼 생성
$form = New-Object System.Windows.Forms.Form
$form.Text = "$programName v$version"
$form.Size = New-Object System.Drawing.Size(615, 400) # 폼 크기 유지
$form.StartPosition = "CenterScreen"

# GUI 이름 레이블
$guiNameLabel = New-Object System.Windows.Forms.Label
$guiNameLabel.Location = New-Object System.Drawing.Point(20, 10)
$guiNameLabel.Size = New-Object System.Drawing.Size(200, 25)
$guiNameLabel.Text = "ISSO WIN SCRIPT"
$guiNameLabel.Font = New-Object System.Drawing.Font("맑은 고딕", 12, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($guiNameLabel)

# 버전 정보 버튼
$versionButton = New-Object System.Windows.Forms.Button
$versionButton.Location = New-Object System.Drawing.Point(500, 10)
$versionButton.Size = New-Object System.Drawing.Size(80, 25)
$versionButton.Text = "버전 정보"
$versionButton.Add_Click({
    try {
        Start-Process "history.txt"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("history.txt 파일을 찾을 수 없습니다.", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$form.Controls.Add($versionButton)

# 설명 패널
$descriptionPanel = New-Object System.Windows.Forms.Panel
$descriptionPanel.Location = New-Object System.Drawing.Point(20, 50) # 왼쪽 여백 20 유지
$descriptionPanel.Size = New-Object System.Drawing.Size(560, 150) # 너비 560 유지
$descriptionPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($descriptionPanel)

# 설명 레이블
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Location = New-Object System.Drawing.Point(5, 5)
$descriptionLabel.Size = New-Object System.Drawing.Size(550, 140)
$descriptionLabel.Text = "이 도구는 Windows 시스템의 기본적인 보안 설정을 감사하고 수정하는 데 사용됩니다. 다음 기능을 제공합니다.`r`n- 방화벽 상태 확인 및 설정`r`n- 사용자 계정 관리`r`n- UAC 설정 관리`r`n- 로컬 보안 정책 감사 및 수정`r`n- Windows 업데이트 설정 관리`r`n- 바이러스 백신 상태 확인"
$descriptionPanel.Controls.Add($descriptionLabel)

# 경고 레이블
$warningLabel = New-Object System.Windows.Forms.Label
$warningLabel.Location = New-Object System.Drawing.Point(20, 210)
$warningLabel.Size = New-Object System.Drawing.Size(560, 60)
$warningLabel.Text = "경고: 이 도구를 잘못 사용하면 시스템 보안에 심각한 영향을 줄 수 있습니다. 신중하게 사용하시고, 변경 사항을 적용하기 전에 반드시 백업을 수행하십시오."
$warningLabel.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($warningLabel)

# 시작 버튼
$startButton = New-Object System.Windows.Forms.Button
$startButton.Location = New-Object System.Drawing.Point(200, 300)
$startButton.Size = New-Object System.Drawing.Size(100, 30)
$startButton.Text = "시작"
$startButton.Add_Click({
    # 시작 버튼 클릭 시 수행할 작업 (예: 다음 화면으로 이동)
    [System.Windows.Forms.MessageBox]::Show("감사 도구를 시작합니다.", "알림")
})
$form.Controls.Add($startButton)

# 종료 버튼
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Location = New-Object System.Drawing.Point(320, 300)
$exitButton.Size = New-Object System.Drawing.Size(100, 30)
$exitButton.Text = "종료"
$exitButton.Add_Click({
    $form.Close()
})
$form.Controls.Add($exitButton)

$form.ShowDialog()