
## TODO : 보고서 PDF로 출력하는 거

# Windows Forms 어셈블리 로드
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


# Update-ComputerComboBox 함수
function Update-ComputerComboBox {
    $computerComboBox.Items.Clear() # 기존 목록 초기화
    if (Test-Path $jsonFilePath) {
        try {
            $existingData = Get-Content $jsonFilePath | ConvertFrom-Json
            $existingData | ForEach-Object { $computerComboBox.Items.Add($_.ComputerName) }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("JSON 파일 읽기 오류: $($_.Exception.Message)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

# JSON 파일 경로
$jsonFilePath = "$PSScriptRoot\targetWorkstation.json"

# Form 생성
$form = New-Object System.Windows.Forms.Form
$form.Text = "보안 설정 감사"
$form.Size = New-Object System.Drawing.Size(1230, 850)
$form.StartPosition = "CenterScreen"

# 이미지 로드
$imagePath = "C:\Users\ace30\Desktop\WIN_INSPECTION_SCRIPT\background.png" # 실제 이미지 경로로 변경
try {
    $image = [System.Drawing.Image]::FromFile($imagePath)
    $form.BackgroundImage = $image
    $form.BackgroundImageLayout = "Stretch" # 이미지 늘리기 또는 "Tile"로 반복
} catch {
    Write-Error "이미지 로드 오류: $($_.Exception.Message)"
}

# TableLayoutPanel 컨트롤 생성
$tableLayout = New-Object System.Windows.Forms.TableLayoutPanel
$tableLayout.Dock = "Fill" # Form에 꽉 채우도록 설정
$form.Controls.Add($tableLayout)

# TableLayoutPanel 행 및 열 설정
$tableLayout.RowCount = 2
$null = $tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("AutoSize"))) # 툴바 행 자동 크기 조정
$null = $tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 100))) # 내용 영역 행 나머지 공간 채우기
$tableLayout.ColumnCount = 1
$null = $tableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100))) # 열 너비 100%

# ToolStrip 컨트롤 생성 및 TableLayoutPanel 첫 번째 행에 추가
$toolStrip = New-Object System.Windows.Forms.ToolStrip
$toolStrip.Dock = "Fill" # 행에 꽉 채우도록 설정
$tableLayout.Controls.Add($toolStrip, 0, 0)

# 편집 버튼 생성 및 추가
$editButton = New-Object System.Windows.Forms.ToolStripDropDownButton
$editButton.Text = "편집"

# 드롭다운 메뉴 항목 생성
$newCompMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("New Comp")
$editCompMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("Edit Comp")
$newItemMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("New Item")

# 구분선 추가
$separator = New-Object System.Windows.Forms.ToolStripSeparator

# 드롭다운 메뉴 항목 클릭 이벤트 핸들러 (예시)
$newCompMenuItem.Add_Click({ [System.Windows.Forms.MessageBox]::Show("New Comp 클릭!", "알림") })
$editCompMenuItem.Add_Click({ [System.Windows.Forms.MessageBox]::Show("Edit Comp 클릭!", "알림") })
$newItemMenuItem.Add_Click({ [System.Windows.Forms.MessageBox]::Show("New Item 클릭!", "알림") })

# 드롭다운 메뉴 항목을 편집 버튼에 추가
$null = $editButton.DropDownItems.AddRange(@($newCompMenuItem, $editCompMenuItem,$separator, $newItemMenuItem))
$null = $toolStrip.Items.Add($editButton)

# 정보 버튼 생성 및 추가
$infoButton = New-Object System.Windows.Forms.ToolStripButton
$infoButton.Text = "정보"
$infoButton.Add_Click({ [System.Windows.Forms.MessageBox]::Show("정보 버튼 클릭!", "알림") })
$null = $toolStrip.Items.Add($infoButton)

# 내용 영역 패널 생성 및 TableLayoutPanel 두 번째 행에 추가
$contentPanel = New-Object System.Windows.Forms.Panel
$contentPanel.Dock = "Fill" # 행에 꽉 채우도록 설정
$contentPanel.BackColor = [System.Drawing.Color]::Transparent # 배경 투명 설정

$contentPanel.BackgroundImage = $image
$contentPanel.BackgroundImageLayout = "Stretch" # 이미지 늘리기 또는 "Tile"로 반복
$tableLayout.Controls.Add($contentPanel, 0, 1)

# 기본 정보 영역 패널 생성
$basicInfoPanel = New-Object System.Windows.Forms.Panel
$basicInfoPanel.Location = New-Object System.Drawing.Point(0, 0)
$basicInfoPanel.Size = New-Object System.Drawing.Size(1200, 100)
$basicInfoPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$basicInfoPanel.BackColor = [System.Drawing.Color]::Transparent # 배경 투명 설정
$contentPanel.Controls.Add($basicInfoPanel)

# 제목 레이블 생성
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(10, 10)
$titleLabel.Size = New-Object System.Drawing.Size(1180, 30)
$titleLabel.Text = "Inspection Target"
$titleLabel.Font = New-Object System.Drawing.Font("맑은 고딕", 16, [System.Drawing.FontStyle]::Bold)
$basicInfoPanel.Controls.Add($titleLabel)

# 컴퓨터 선택 ComboBox 생성
$computerComboBox = New-Object System.Windows.Forms.ComboBox
$computerComboBox.Location = New-Object System.Drawing.Point(10, 50)
$computerComboBox.Size = New-Object System.Drawing.Size(200, 25)
$computerComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList # 직접 입력 방지
$basicInfoPanel.Controls.Add($computerComboBox)

# 세부정보 버튼 생성
$detailButton = New-Object System.Windows.Forms.Button
$detailButton.Location = New-Object System.Drawing.Point(220, 50)
$detailButton.Size = New-Object System.Drawing.Size(80, 25)
$detailButton.Text = "세부정보"
$basicInfoPanel.Controls.Add($detailButton)

# 새 감사대상 버튼 생성
$newTargetButton = New-Object System.Windows.Forms.Button
$newTargetButton.Location = New-Object System.Drawing.Point(310, 50)
$newTargetButton.Size = New-Object System.Drawing.Size(100, 25)
$newTargetButton.Text = "새 감사대상"
$basicInfoPanel.Controls.Add($newTargetButton)

# 세부정보 버튼 클릭 이벤트 핸들러
$detailButton.Add_Click({
    $selectedComputer = $computerComboBox.SelectedItem
    if ($selectedComputer) {
        # 새 감사대상 입력 Form 생성
        $detailForm = New-Object System.Windows.Forms.Form
        $detailForm.Text = "감사대상 세부정보"
        $detailForm.Size = New-Object System.Drawing.Size(300, 200)
        $detailForm.StartPosition = "CenterScreen"

        # 컴퓨터 이름 레이블 및 텍스트 박스
        $computerNameLabel = New-Object System.Windows.Forms.Label
        $computerNameLabel.Location = New-Object System.Drawing.Point(10, 20)
        $computerNameLabel.Size = New-Object System.Drawing.Size(80, 20)
        $computerNameLabel.Text = "컴퓨터 이름:"
        $detailForm.Controls.Add($computerNameLabel)

        $computerNameTextBox = New-Object System.Windows.Forms.TextBox
        $computerNameTextBox.Location = New-Object System.Drawing.Point(100, 20)
        $computerNameTextBox.Size = New-Object System.Drawing.Size(180, 20)
        $computerNameTextBox.Text = $selectedComputer # 선택된 컴퓨터 이름 표시
        $detailForm.Controls.Add($computerNameTextBox)

        # 설명 레이블 및 텍스트 박스
        $descriptionLabel = New-Object System.Windows.Forms.Label
        $descriptionLabel.Location = New-Object System.Drawing.Point(10, 50)
        $descriptionLabel.Size = New-Object System.Drawing.Size(80, 20)
        $descriptionLabel.Text = "설명:"
        $detailForm.Controls.Add($descriptionLabel)

        $descriptionTextBox = New-Object System.Windows.Forms.TextBox
        $descriptionTextBox.Location = New-Object System.Drawing.Point(100, 50)
        $descriptionTextBox.Size = New-Object System.Drawing.Size(180, 20)
        # 선택된 컴퓨터 설명 표시 (JSON 파일에서 읽어와야 함)
        $jsonFilePath = "$PSScriptRoot\targetWorkstation.json"
        if (Test-Path $jsonFilePath) {
            try {
                $existingData = Get-Content $jsonFilePath | ConvertFrom-Json
                $selectedData = $existingData | Where-Object { $_.ComputerName -eq $selectedComputer }
                if ($selectedData) {
                    $descriptionTextBox.Text = $selectedData.Description
                }
            } catch {
                [System.Windows.Forms.MessageBox]::Show("JSON 파일 읽기 오류: $($_.Exception.Message)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        $detailForm.Controls.Add($descriptionTextBox)

        # 저장 버튼
        $saveButton = New-Object System.Windows.Forms.Button
        $saveButton.Location = New-Object System.Drawing.Point(10, 120)
        $saveButton.Size = New-Object System.Drawing.Size(60, 30)
        $saveButton.Text = "저장"
        $detailForm.Controls.Add($saveButton)

        # 삭제 버튼
        $deleteButton = New-Object System.Windows.Forms.Button
        $deleteButton.Location = New-Object System.Drawing.Point(80, 120)
        $deleteButton.Size = New-Object System.Drawing.Size(60, 30)
        $deleteButton.Text = "삭제"
        $detailForm.Controls.Add($deleteButton)

        # 취소 버튼
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Location = New-Object System.Drawing.Point(150, 120)
        $cancelButton.Size = New-Object System.Drawing.Size(60, 30)
        $cancelButton.Text = "취소"
        $detailForm.Controls.Add($cancelButton)

        # 버튼 가운데 정렬
        $buttonWidth = $saveButton.Width + $deleteButton.Width + $cancelButton.Width + 20 # 버튼 사이 간격 10 * 2
        $formWidth = $detailForm.ClientSize.Width
        $startX = ($formWidth - $buttonWidth) / 2

        $saveButton.Location = New-Object System.Drawing.Point($startX, 120)
        $deletetmpLocation = $startX + 70
        $deleteButton.Location = New-Object System.Drawing.Point($deletetmpLocation, 120)
        $cancletmpLocation = $startX + 140
        $cancelButton.Location = New-Object System.Drawing.Point($cancletmpLocation, 120)

        # 저장 버튼 클릭 이벤트 핸들러
        $saveButton.Add_Click({
            # 수정된 정보 저장 로직 구현
            $newComputerName = $computerNameTextBox.Text
            $newDescription = $descriptionTextBox.Text

            # JSON 파일 경로
            $jsonFilePath = "$PSScriptRoot\targetWorkstation.json"

            # 기존 JSON 데이터 읽기
            if (Test-Path $jsonFilePath) {
                try {
                    $existingData = Get-Content $jsonFilePath | ConvertFrom-Json
                    # 수정된 데이터 찾기
                    for ($i = 0; $i -lt $existingData.Count; $i++) {
                        if ($existingData[$i].ComputerName -eq $selectedComputer) {
                            # 데이터 수정
                            $existingData[$i].ComputerName = $newComputerName
                            $existingData[$i].Description = $newDescription
                            break
                        }
                    }
                    # JSON 파일에 저장
                    $existingData | ConvertTo-Json | Set-Content $jsonFilePath
                    # ComboBox 업데이트
                    Update-ComputerComboBox
                    # titleLabel 업데이트
                    $titleLabel.Text = "[$newComputerName]"
                    # Form 닫기
                    $detailForm.Close()
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("JSON 파일 읽기/쓰기 오류: $($_.Exception.Message)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
            $detailForm.Close()
        })

        # 삭제 버튼 클릭 이벤트 핸들러
$deleteButton.Add_Click({
    # 삭제 확인 메시지 표시
    $result = [System.Windows.Forms.MessageBox]::Show("정말 삭제하시겠습니까?", "확인", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

    # 사용자가 '예'를 선택한 경우에만 삭제 진행
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        # 선택된 컴퓨터 삭제 로직 구현
        $selectedComputer = $computerNameTextBox.Text

        # JSON 파일 경로
        $jsonFilePath = "$PSScriptRoot\targetWorkstation.json"

        # 기존 JSON 데이터 읽기
        if (Test-Path $jsonFilePath) {
            try {
                $existingData = Get-Content $jsonFilePath | ConvertFrom-Json
                # 삭제할 데이터 찾기
                $newData = @()
                foreach ($item in $existingData) {
                    if ($item.ComputerName -ne $selectedComputer) {
                        $newData += $item
                    }
                }
                # JSON 파일에 저장
                $newData | ConvertTo-Json | Set-Content $jsonFilePath
                # ComboBox 업데이트
                Update-ComputerComboBox
                # titleLabel 업데이트
                $titleLabel.Text = "Inspection Target"
                # Form 닫기
                $detailForm.Close()
            } catch {
                [System.Windows.Forms.MessageBox]::Show("JSON 파일 읽기/쓰기 오류: $($_.Exception.Message)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    }
})

        # 취소 버튼 클릭 이벤트 핸들러
        $cancelButton.Add_Click({
            $detailForm.Close()
        })

        # Form 표시
        $detailForm.ShowDialog()
    }
})

# 새 감사대상 버튼 클릭 이벤트 핸들러
$newTargetButton.Add_Click({
    # 새 감사대상 입력 Form 생성
    $newTargetForm = New-Object System.Windows.Forms.Form
    $newTargetForm.Text = "새 감사대상 추가"
    $newTargetForm.Size = New-Object System.Drawing.Size(300, 200)
    $newTargetForm.StartPosition = "CenterScreen"

    # 컴퓨터 이름 레이블 및 텍스트 박스
    $computerNameLabel = New-Object System.Windows.Forms.Label
    $computerNameLabel.Location = New-Object System.Drawing.Point(10, 20)
    $computerNameLabel.Size = New-Object System.Drawing.Size(80, 20)
    $computerNameLabel.Text = "컴퓨터 이름:"
    $newTargetForm.Controls.Add($computerNameLabel)

    $computerNameTextBox = New-Object System.Windows.Forms.TextBox
    $computerNameTextBox.Location = New-Object System.Drawing.Point(100, 20)
    $computerNameTextBox.Size = New-Object System.Drawing.Size(180, 20)
    $newTargetForm.Controls.Add($computerNameTextBox)

    # 설명 레이블 및 텍스트 박스
    $descriptionLabel = New-Object System.Windows.Forms.Label
    $descriptionLabel.Location = New-Object System.Drawing.Point(10, 50)
    $descriptionLabel.Size = New-Object System.Drawing.Size(80, 20)
    $descriptionLabel.Text = "설명:"
    $newTargetForm.Controls.Add($descriptionLabel)

    $descriptionTextBox = New-Object System.Windows.Forms.TextBox
    $descriptionTextBox.Location = New-Object System.Drawing.Point(100, 50)
    $descriptionTextBox.Size = New-Object System.Drawing.Size(180, 20)
    $newTargetForm.Controls.Add($descriptionTextBox)

    # 확인 버튼
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(50, 120)
    $okButton.Size = New-Object System.Drawing.Size(80, 30)
    $okButton.Text = "확인"
    $newTargetForm.Controls.Add($okButton)

    # 취소 버튼
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(150, 120)
    $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
    $cancelButton.Text = "취소"
    $newTargetForm.Controls.Add($cancelButton)
    
   # 확인 버튼 클릭 이벤트 핸들러
    $okButton.Add_Click({
        # 입력된 정보 가져오기
        $computerName = $computerNameTextBox.Text
        $description = $descriptionTextBox.Text

        # JSON 파일 경로
        $jsonFilePath = "$PSScriptRoot\targetWorkstation.json"

        # 기존 JSON 데이터 읽기 (있는 경우)
        $existingData = if (Test-Path $jsonFilePath) {
            Get-Content $jsonFilePath | ConvertFrom-Json
        } else {
            @()
        }

        # 중복 검사
        $isDuplicate = $existingData | Where-Object { $_.ComputerName -eq $computerName }
        if ($isDuplicate) {
            [System.Windows.Forms.MessageBox]::Show("이미 존재하는 컴퓨터 이름입니다.", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return # 중복이면 추가하지 않고 종료
        }

        # JSON 데이터 생성
        $targetData = @{
            ComputerName = $computerName
            Description = $description
        }

        # ArrayList 생성
        $dataList = New-Object System.Collections.ArrayList

        # 기존 데이터 개별 추가
        $existingData | ForEach-Object {
            $dataList.Add($_) | Out-Null
        }

        # 새 데이터 추가
        $dataList.Add($targetData) | Out-Null

        # JSON 파일에 저장
        $dataList | ConvertTo-Json | Set-Content $jsonFilePath

        # Form 닫기
        $newTargetForm.Close()

        # ComboBox 업데이트
        Update-ComputerComboBox
    })

    # 취소 버튼 클릭 이벤트 핸들러
    $cancelButton.Add_Click({
        $newTargetForm.Close()
    })

    # Form 표시
    $newTargetForm.ShowDialog()
})

# Form 로드 시 이벤트 핸들러
$form.Add_Shown({
    $form.Activate()
    Update-ComputerComboBox
    $basicInfoChangeButton.Enabled = ($computerComboBox.SelectedItem -ne $null)

})

# computerComboBox 클릭 시 이벤트 핸들러
$computerComboBox.Add_DropDown({
    Update-ComputerComboBox
})

# computerComboBox 선택 변경 시 이벤트 핸들러
$computerComboBox.Add_SelectedIndexChanged({
    $basicInfoChangeButton.Enabled = ($computerComboBox.SelectedItem -ne $null)
    $selectedComputer = $computerComboBox.SelectedItem
    if ($selectedComputer) {
        $securityPanelTitleLabel.Text = "Security Panel [$($selectedComputer)]" # Security Panel 제목 업데이트
    } else {
        $titleLabel.Text = "Inspection Target"
        $securityPanelTitleLabel.Text = "Security Panel" # 초기 제목으로 복원
    }
})

# 보안 현황 영역 패널 생성
$securityStatusPanel = New-Object System.Windows.Forms.Panel
$securityStatusPanel.Location = New-Object System.Drawing.Point(0, 100)
$securityStatusPanel.Size = New-Object System.Drawing.Size(1200, 600)
$securityStatusPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$securityStatusPanel.BackColor = [System.Drawing.Color]::Transparent # 배경 투명 설정
$contentPanel.Controls.Add($securityStatusPanel)

# Security Panel 제목 레이블 생성
$securityPanelTitleLabel = New-Object System.Windows.Forms.Label
$securityPanelTitleLabel.Location = New-Object System.Drawing.Point(10, 10) # 패널 좌측 상단에 위치
$securityPanelTitleLabel.Size = New-Object System.Drawing.Size(500, 30) # 크기 설정
$securityPanelTitleLabel.Text = "Security Panel"
$securityPanelTitleLabel.Font = New-Object System.Drawing.Font("맑은 고딕", 16, [System.Drawing.FontStyle]::Bold) # 폰트 설정
$securityStatusPanel.Controls.Add($securityPanelTitleLabel)

# Basic Info 레이블 생성
$basicInfoLabel = New-Object System.Windows.Forms.Label
$basicInfoLabel.Location = New-Object System.Drawing.Point(10, 50)
$basicInfoLabel.Size = New-Object System.Drawing.Size(120, 30) # 크기 증가
$basicInfoLabel.Text = "Basic Info"
$basicInfoLabel.Font = New-Object System.Drawing.Font("맑은 고딕", 12, [System.Drawing.FontStyle]::Bold) # 폰트 크기 및 굵기 설정
$securityStatusPanel.Controls.Add($basicInfoLabel)

# Basic Info 변경 버튼 생성
$basicInfoChangeButton = New-Object System.Windows.Forms.Button
$basicInfoChangeButton.Size = New-Object System.Drawing.Size(100, 25)
$basicInfoChangeButton.Location = New-Object System.Drawing.Point(
    [int]($basicInfoTextArea.Right - $basicInfoChangeButton.Width), # Basic Info TextArea 우측 상단에 위치
    $basicInfoLabel.Location.Y
)
$basicInfoChangeButton.Text = "기본 정보 변경"
$securityStatusPanel.Controls.Add($basicInfoChangeButton)

# Basic Info 변경 버튼 클릭 이벤트 핸들러
$basicInfoChangeButton.Add_Click({
    if ($basicInfoTextArea.ReadOnly) {
        # 읽기 전용 모드 해제 및 버튼 텍스트 변경
        $basicInfoTextArea.ReadOnly = $false
        $basicInfoChangeButton.Text = "저장"
    } else {
        # 편집 내용 저장 및 읽기 전용 모드 설정
        # TODO: 편집 내용 저장 로직 구현
        $basicInfoTextArea.ReadOnly = $true
        $basicInfoChangeButton.Text = "기본 정보 변경"
    
    }
})

# 초기 버튼 비활성화 설정
$basicInfoChangeButton.Enabled = $false

# TextArea (TextBox) 생성
$basicInfoTextArea = New-Object System.Windows.Forms.TextBox
$basicInfoTextArea.Location = New-Object System.Drawing.Point(10, 80)
$basicInfoTextArea.Size = New-Object System.Drawing.Size(
    [int]($securityStatusPanel.Width * 2 / 7+30),
    [int]($securityStatusPanel.Height * 2.25 / 5)
)
$basicInfoTextArea.Multiline = $true
# 내용이 길어질 때만 스크롤 막대 표시
if ($basicInfoTextArea.Text.Length -gt 100) { # 예시: 내용 길이가 100자를 초과하면
    $basicInfoTextArea.ScrollBars = "Vertical"
} else {
    $basicInfoTextArea.ScrollBars = "None"
}
$basicInfoTextArea.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle # 검은 테두리 추가
$basicInfoTextArea.ReadOnly = $true # 편집 불가능 설정
$securityStatusPanel.Controls.Add($basicInfoTextArea)

# Audit History 레이블 생성
$auditHistoryLabel = New-Object System.Windows.Forms.Label
$auditHistoryLabel.Location = New-Object System.Drawing.Point(10, [int]($basicInfoTextArea.Bottom + 15)) # TextArea 아래에 위치
$auditHistoryLabel.Size = New-Object System.Drawing.Size(120, 30)
$auditHistoryLabel.Text = "Audit History"
$auditHistoryLabel.Font = New-Object System.Drawing.Font("맑은 고딕", 12, [System.Drawing.FontStyle]::Bold)
$securityStatusPanel.Controls.Add($auditHistoryLabel)


# Audit History 표 영역 (ListView) 생성
$auditHistoryListView = New-Object System.Windows.Forms.ListView
$auditHistoryListView.Location = New-Object System.Drawing.Point(10, [int]($auditHistoryLabel.Bottom))
$auditHistoryListView.Size = New-Object System.Drawing.Size(
    [int]($basicInfoTextArea.Size.Width),
    [int]($securityStatusPanel.Height * 0.25)
)
$auditHistoryListView.View = "Details" # 테이블 형태로 표시
$auditHistoryListView.FullRowSelect = $true # 행 전체 선택
$auditHistoryListView.GridLines = $true # 격자선 표시

# 테이블 컬럼 추가
$null = $auditHistoryListView.Columns.Add("날짜", [int]($auditHistoryListView.Width * 0.3))
$null = $auditHistoryListView.Columns.Add("감사 항목", [int]($auditHistoryListView.Width * 0.5))
$null = $auditHistoryListView.Columns.Add("보고서", [int]($auditHistoryListView.Width * 0.2))


# 열 너비 변경 이벤트 핸들러
$auditHistoryListView.Add_ColumnWidthChanging({
    param($sender, $e)
    $e.Cancel = $true # 열 너비 변경 취소
})

$securityStatusPanel.Controls.Add($auditHistoryListView)

# AccountList 레이블 생성
$accountListLabel = New-Object System.Windows.Forms.Label
$accountListLabel.Location = New-Object System.Drawing.Point(
    [int]($basicInfoTextArea.Right + 10), # Audit History ListView 오른쪽에 위치
    50 # BasicInfoTextArea의 Top보다 40픽셀 위에 위치
)
$accountListLabel.Size = New-Object System.Drawing.Size(120, 30)
$accountListLabel.Text = "Account List"
$accountListLabel.Font = New-Object System.Drawing.Font("맑은 고딕", 12, [System.Drawing.FontStyle]::Bold)
$securityStatusPanel.Controls.Add($accountListLabel)

# AccountList ListView 생성
$accountListListView = New-Object System.Windows.Forms.ListView
$accountListListView.Location = New-Object System.Drawing.Point(
    [int]($basicInfoTextArea.Right + 10), # Audit History ListView 오른쪽에 위치
    [int]($basicInfoTextArea.Top) # BasicInfoTextArea의 Top과 동일한 높이
)
$accountListListView.Size = New-Object System.Drawing.Size(
    [int]($securityStatusPanel.Width * 1 / 3 - 20), # 남은 너비의 1/3
    $basicInfoTextArea.Size.Height
)
$accountListListView.View = "Details"
$accountListListView.FullRowSelect = $true
$accountListListView.GridLines = $true

# AccountList 컬럼 추가
$null = $accountListListView.Columns.Add("계정 이름", [int]($accountListListView.Width * 0.6))
$null = $accountListListView.Columns.Add("권한", [int]($accountListListView.Width * 0.2))
$null = $accountListListView.Columns.Add("상태", [int]($accountListListView.Width * 0.2))

$securityStatusPanel.Controls.Add($accountListListView)


# Service List 레이블 생성
$serviceListLabel = New-Object System.Windows.Forms.Label
$serviceListLabel.Location = New-Object System.Drawing.Point(
    [int]($accountListListView.Right + 10), # AccountList ListView 오른쪽에 위치
    50 # BasicInfoTextArea의 Top보다 40픽셀 위에 위치
)
$serviceListLabel.Size = New-Object System.Drawing.Size(120, 30)
$serviceListLabel.Text = "Service List"
$serviceListLabel.Font = New-Object System.Drawing.Font("맑은 고딕", 12, [System.Drawing.FontStyle]::Bold)
$securityStatusPanel.Controls.Add($serviceListLabel)

# Service List ListView 생성
$serviceListListView = New-Object System.Windows.Forms.ListView
$serviceListListView.Location = New-Object System.Drawing.Point(
    [int]($accountListListView.Right + 10),
    [int]($basicInfoTextArea.Top) # Service List 레이블 아래에 위치
)
$serviceListListView.Size = New-Object System.Drawing.Size(
    $accountListListView.Width, # AccountList ListView와 동일한 너비
    $basicInfoTextArea.Size.Height
)
$serviceListListView.View = "Details"
$serviceListListView.FullRowSelect = $true
$serviceListListView.GridLines = $true

# Service List 컬럼 추가
$null = $serviceListListView.Columns.Add("서비스 이름", [int]($serviceListListView.Width * 0.5))
$null = $serviceListListView.Columns.Add("상태", [int]($serviceListListView.Width * 0.5))

$securityStatusPanel.Controls.Add($serviceListListView)


# Approved Media 레이블 생성
$approvedMediaLabel = New-Object System.Windows.Forms.Label
$approvedMediaLabel.Location = New-Object System.Drawing.Point(
    [int]($auditHistoryListView.Right + 10), # Service List ListView 오른쪽에 위치
    [int]($basicInfoTextArea.Bottom + 15) # Service List ListView 레이블 위에 위치
)
$approvedMediaLabel.Size = New-Object System.Drawing.Size(180, 30)
$approvedMediaLabel.Text = "Approved Media"
$approvedMediaLabel.Font = New-Object System.Drawing.Font("맑은 고딕", 12, [System.Drawing.FontStyle]::Bold)
$securityStatusPanel.Controls.Add($approvedMediaLabel)

# Approved Media ListView 생성
$approvedMediaListView = New-Object System.Windows.Forms.ListView
$approvedMediaListView.Location = New-Object System.Drawing.Point(
    [int]($basicInfoTextArea.Right + 10), # Audit History ListView 오른쪽에 위치
    [int]($approvedMediaLabel.Bottom)
)
$approvedMediaListView.Size = New-Object System.Drawing.Size(
    [int]($serviceListListView.Right - $accountListListView.Left), # Service List ListView와 동일한 높이
    [int]($auditHistoryListView.Height) # Service List ListView와 동일한 너비
)
$approvedMediaListView.View = "Details"
$approvedMediaListView.FullRowSelect = $true
$approvedMediaListView.GridLines = $true

# Approved Media 컬럼 추가
$null = $approvedMediaListView.Columns.Add("미디어", [int]($approvedMediaListView.Width * 0.3))
$null = $approvedMediaListView.Columns.Add("Reg", [int]($approvedMediaListView.Width * 0.7))

$securityStatusPanel.Controls.Add($approvedMediaListView)

# Audit ListView 아래 레이블 생성
$auditDateLabel = New-Object System.Windows.Forms.Label
$auditDateLabel.Location = New-Object System.Drawing.Point(10, [int]($auditHistoryListView.Bottom + 10))
$auditDateLabel.Size = New-Object System.Drawing.Size(800, 20)
$auditDateLabel.Text = "* 현황 기준일 : YYYY-MM-dd"
$auditDateLabel.Font = New-Object System.Drawing.Font("맑은 고딕", 10, [System.Drawing.FontStyle]::Bold)
$auditDateLabel.ForeColor = [System.Drawing.Color]::Green
$securityStatusPanel.Controls.Add($auditDateLabel)

# 버튼 영역 패널 생성
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Location = New-Object System.Drawing.Point(0, 700)
$buttonPanel.Size = New-Object System.Drawing.Size(1200, 60)
$buttonPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$buttonPanel.BackColor = [System.Drawing.Color]::Transparent # 배경 투명 설정
$contentPanel.Controls.Add($buttonPanel)


# ComboBox 생성
$compFolderComboBox = New-Object System.Windows.Forms.ComboBox
$compFolderComboBox.DropDownStyle = "DropDownList" # 입력 불가 설정
$compFolderComboBox.Size = New-Object System.Drawing.Size(200, 25)

# 폴더 목록 가져오기
$compFolderPath = Join-Path -Path $PSScriptRoot -ChildPath "comp"
if (Test-Path -Path $compFolderPath -PathType Container) {
    $compFolders = Get-ChildItem -Path $compFolderPath -Directory | Select-Object -ExpandProperty Name
    $compFolderComboBox.Items.AddRange($compFolders)
}

$buttonPanel.Controls.Add($compFolderComboBox)

# 감사수행 버튼 생성
$auditPerformButton = New-Object System.Windows.Forms.Button
$auditPerformButton.Size = New-Object System.Drawing.Size(100, 25)
$auditPerformButton.Text = "감사수행"
$buttonPanel.Controls.Add($auditPerformButton)

# 취소 버튼 생성
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Size = New-Object System.Drawing.Size(100, 25)
$cancelButton.Text = "취소"
$buttonPanel.Controls.Add($cancelButton)

# 컨트롤 위치 계산 및 설정
$totalWidth = $auditItemLabel.Width + $compFolderComboBox.Width + $auditPerformButton.Width + $cancelButton.Width + 30 # 컨트롤 사이 간격 30 추가
$startX = ($form.ClientSize.Width - $totalWidth) / 2

$compComboboxtmpLocation = $startX + 10
$compFolderComboBox.Location = New-Object System.Drawing.Point($compComboboxtmpLocation, 10)
$auditPerformtmpLocation = $startX + $compFolderComboBox.Width + 20
$auditPerformButton.Location = New-Object System.Drawing.Point($auditPerformtmpLocation, 10)
$canceltmpLocation = $startX + $compFolderComboBox.Width + $auditPerformButton.Width + 30
$cancelButton.Location = New-Object System.Drawing.Point($canceltmpLocation, 10)

# Form 표시
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()
