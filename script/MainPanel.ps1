
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
# 감사 수행 버튼 활성화 조건 설정
function Update-AuditPerformButtonState {
    if ($computerComboBox.SelectedItem -and $compFolderComboBox.SelectedItem) {
        $auditPerformButton.Enabled = $true
    } else {
        $auditPerformButton.Enabled = $false
    }
}

function Update-AccountListView {
    param(
        [Parameter(Mandatory=$true)]
        [string]$JsonFilePath,
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.ListView]$ListView
    )
    $ListView.Items.Clear()
    if (Test-Path $JsonFilePath) {
        try {
            $accounts = Get-Content $JsonFilePath | ConvertFrom-Json
            foreach ($account in $accounts) {
                $listItem = New-Object System.Windows.Forms.ListViewItem($account.SamAccountName)
                $listItem.SubItems.Add($account.Enabled.ToString())
                $null = $ListView.Items.Add($listItem)
            }
        } catch {
            Write-Warning "계정 정보 JSON 파일 처리 오류: $($_.Exception.Message) - 파일: $JsonFilePath"
        }
    } else {
        Write-Warning "계정 정보 JSON 파일을 찾을 수 없습니다: $JsonFilePath"
    }
}
function Update-ServiceListView {
    param(
        [Parameter(Mandatory=$true)]
        [string]$JsonFilePath,
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.ListView]$ListView
    )
    $ListView.Items.Clear()
    if (Test-Path $JsonFilePath) {
        try {
            $services = Get-Content $JsonFilePath | ConvertFrom-Json
            foreach ($service in $services) {
                $listItem = New-Object System.Windows.Forms.ListViewItem($service.DisplayName)
                $listItem.SubItems.Add($service.Status)
                $null = $ListView.Items.Add($listItem)
            }
        } catch {
            Write-Warning "서비스 정보 JSON 파일 처리 오류: $($_.Exception.Message) - 파일: $JsonFilePath"
        }
    } else {
        Write-Warning "서비스 정보 JSON 파일을 찾을 수 없습니다: $JsonFilePath"
    }
}
function Update-ApprovedMediaListView {
    param(
        [Parameter(Mandatory=$true)]
        [string]$JsonFilePath,
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.ListView]$ListView
    )
    $ListView.Items.Clear()
    if (Test-Path $JsonFilePath) {
        try {
            $mediaList = Get-Content $JsonFilePath | ConvertFrom-Json
            foreach ($media in $mediaList) {
                $listItem = New-Object System.Windows.Forms.ListViewItem($media."미디어")
                $listItem.SubItems.Add($media."Reg")
                $null = $ListView.Items.Add($listItem)
            }
        } catch {
            Write-Warning "승인된 미디어 정보 JSON 파일 처리 오류: $($_.Exception.Message) - 파일: $JsonFilePath"
        }
    } else {
        Write-Warning "승인된 미디어 정보 JSON 파일을 찾을 수 없습니다: $JsonFilePath"
    }
}

# JSON 파일 경로
$jsonFilePath = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "etc\targetWorkstation.json"

# Form 생성
$form = New-Object System.Windows.Forms.Form
$form.Text = "보안 설정 감사"
$form.Size = New-Object System.Drawing.Size(1230, 850)
$form.StartPosition = "CenterScreen"

# 이미지 로드
$imagePath  = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "etc\background1.png" # 실제 이미지 경로로 변경
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
$newCompMenuItem.Add_Click({ Show-NewCompSetForm })
$editCompMenuItem.Add_Click({ Show-CompListForm })
$newItemMenuItem.Add_Click({
   . "$PSScriptRoot\CreateJson.ps1"
})


# 드롭다운 메뉴 항목을 편집 버튼에 추가
$null = $editButton.DropDownItems.AddRange(@($newCompMenuItem, $editCompMenuItem, $separator, $newItemMenuItem))
$null = $toolStrip.Items.Add($editButton)

# 정보 버튼 생성 및 추가
$infoButton = New-Object System.Windows.Forms.ToolStripButton
$infoButton.Text = "정보"
# 정보 버튼 클릭 이벤트 핸들러 (MainPanel.ps1)
$infoButton.Add_Click({
    # 현재 스크립트 파일의 경로 가져오기
    $scriptFullPath = $MyInvocation.MyCommand.Path
    # 스크립트 파일이 실행되지 않은 경우 (예: 콘솔에서 직접 명령어 입력) 처리
    if (-not $scriptFullPath) {
        $scriptDirectory = Get-Location
    } else {
        # 스크립트 파일이 있는 폴더 경로 가져오기
        $scriptDirectory = Split-Path -parent $scriptFullPath
    }
    # 상위 폴더 경로 가져오기
    $parentDirectory = Split-Path -parent $scriptDirectory
    # readme.txt 파일의 전체 경로 생성
    $readmeFilePath = Join-Path -Path $parentDirectory -ChildPath "readme.txt"

    if (Test-Path $readmeFilePath -PathType Leaf) {
        try {
            # notepad.exe 프로세스 시작 및 readme.txt 파일 열기
            Start-Process -FilePath "notepad.exe" -ArgumentList "$readmeFilePath"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("readme.txt 파일을 여는 동안 오류가 발생했습니다: $($_.Exception.Message)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("상위 폴더에서 readme.txt 파일을 찾을 수 없습니다.", "알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})
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
#$basicInfoPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
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
        $jsonFilePath = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "etc\targetWorkstation.json"
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
            $jsonFilePath = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "etc\\targetWorkstation.json"

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
        $jsonFilePath = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "etc\targetWorkstation.json"

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
        $jsonFilePath = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "etc\targetWorkstation.json"

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

        # 로그 및 보고서 폴더 생성
        $reportFolderPath = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "Report\$computerName"

        try {
            if (-not (Test-Path $reportFolderPath -PathType Container)) {
                New-Item -Path $reportFolderPath -ItemType Directory -Force | Out-Null
                Write-Host "보고서 폴더 생성: $reportFolderPath"
            }
        } catch {
            Write-Error "보고서 폴더 생성 실패: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("로그 또는 보고서 폴더 생성에 실패했습니다: $($_.Exception.Message)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }

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
        Load-AuditHistory -WsName $selectedComputer

        # Basic Info 불러오기
        $reportPathRoot = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "Report"
        $wsReportPath = Join-Path -Path $reportPathRoot -ChildPath $selectedComputer
        $basicInfoFilePath = Join-Path -Path $wsReportPath -ChildPath "basic_info.txt"

        if (Test-Path $basicInfoFilePath -PathType Leaf) {
            try {
                $basicInfoTextArea.Text = Get-Content -Path $basicInfoFilePath -Raw -Encoding UTF8
            } catch {
                $basicInfoTextArea.Text = "Basic Info 파일을 읽는 데 실패했습니다."
            }
        } else {
            $basicInfoTextArea.Text = "" # 파일이 없으면 텍스트 영역을 비움
        }

        # Log Backup 버튼 활성화
        $logBackupButton.Enabled = $true

        # 가장 최근 감사 폴더 찾기 및 데이터 로드
        $reportPath = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "Report\$selectedComputer"
        if (Test-Path $reportPath -PathType Container) {
            $latestAuditFolder = Get-ChildItem -Path $reportPath -Directory | Where-Object {$_.Name -match '^\d{8}_\d{6}$'} | Sort-Object Name -Descending | Select-Object -First 1
            if ($latestAuditFolder) {
                $auditDateLabel.Text = "* 현황 기준일: $($latestAuditFolder.Name)"
                Load-LatestAuditData -WsName $selectedComputer -AuditFolderPath $latestAuditFolder.FullName
            } else {
                $auditDateLabel.Text = "감사기록이 없는 WS입니다."
                # Account, Service, Media ListView 비우기
                $accountListListView.Items.Clear()
                $serviceListListView.Items.Clear()
                $approvedMediaListView.Items.Clear()
            }
        } else {
            Write-Warning "Report\$selectedComputer 폴더를 찾을 수 없습니다."
            # Account, Service, Media ListView 비우기
            $accountListListView.Items.Clear()
            $serviceListListView.Items.Clear()
            $approvedMediaListView.Items.Clear()
        }
    } else {
        $titleLabel.Text = "Inspection Target"
        $securityPanelTitleLabel.Text = "Security Panel" # 초기 제목으로 복원
        # WS 선택 해제 시 모든 ListView 비우기
        $auditHistoryListView.Items.Clear()
        $accountListListView.Items.Clear()
        $serviceListListView.Items.Clear()
        $approvedMediaListView.Items.Clear()
        $basicInfoTextArea.Text = ""
        $basicInfoChangeButton.Text = "기본 정보 변경"
        $auditHistoryListView.Items.Clear()
    }
    Update-AuditPerformButtonState
})

# 보안 현황 영역 패널 생성
$securityStatusPanel = New-Object System.Windows.Forms.Panel
$securityStatusPanel.Location = New-Object System.Drawing.Point(0, 100)
$securityStatusPanel.Size = New-Object System.Drawing.Size(1200, 600)
#$securityStatusPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
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

# TODO
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
    if ($basicInfoChangeButton.Text -eq "저장") {
        $selectedComputer = $computerComboBox.SelectedItem
        if (-not [string]::IsNullOrEmpty($selectedComputer)) {
            $infoText = $basicInfoTextArea.Text

            $reportPathRoot = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "Report"
            $wsReportPath = Join-Path -Path $reportPathRoot -ChildPath $selectedComputer
            $basicInfoFilePath = Join-Path -Path $wsReportPath -ChildPath "basic_info.txt"

            # Report 폴더 및 워크스테이션 폴더가 없으면 생성
            if (-not (Test-Path $reportPathRoot -PathType Container)) {
                try {
                    New-Item -ItemType Directory -Path $reportPathRoot -Force | Out-Null
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Report 폴더를 생성하지 못했습니다: $($_.Exception.Message)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }
            }
            if (-not (Test-Path $wsReportPath -PathType Container)) {
                try {
                    New-Item -ItemType Directory -Path $wsReportPath -Force | Out-Null
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("$selectedComputer 폴더를 생성하지 못했습니다: $($_.Exception.Message)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }
            }

            # 정보 저장
            try {
                $infoText | Out-File -FilePath $basicInfoFilePath -Encoding UTF8
                [System.Windows.Forms.MessageBox]::Show("$selectedComputer 정보가 저장되었습니다: $basicInfoFilePath", "저장 완료", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                # 저장 후 버튼 텍스트를 원래대로 변경
                $basicInfoChangeButton.Text = "기본 정 변경"
            } catch {
                [System.Windows.Forms.MessageBox]::Show("정보 저장에 실패했습니다: $($_.Exception.Message)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("워크스테이션을 먼저 선택해주세요.", "알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    } else {
        # "저장" 버튼이 아닐 경우 (원래 "Basic Info 변경" 상태) 다른 동작이 있다면 여기에 구현
        # 현재는 아무 동작도 하지 않음
    }
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
#$basicInfoTextArea.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle # 검은 테두리 추가
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

# Log Backup 버튼 생성
$logBackupButton = New-Object System.Windows.Forms.Button
$logBackupButton.Text = "로그 백업"
$logBackupButton.Size = New-Object System.Drawing.Size(100, 25)
$logBackupButton.Enabled = $false
$logBackupButton.Location = New-Object System.Drawing.Point(
    [int]($auditHistoryListView.Right - $logBackupButton.Width), # Basic Info TextArea 우측 상단에 위치
    $auditHistoryLabel.Location.Y
)
$securityStatusPanel.Controls.Add($logBackupButton)

# Log Backup 버튼 클릭 이벤트 핸들러 (MainPanel.ps1 수정)
$logBackupButton.Add_Click({
    $scriptPath = Join-Path $PSScriptRoot "backupScript.ps1"
    $selectedComputer = $computerComboBox.SelectedItem

    if (Test-Path $scriptPath -PathType Leaf) {
        if ($selectedComputer) {
            try {
                & $scriptPath -WsName $selectedComputer
                # backupScript.ps1이 GUI를 표시하므로, 스크립트가 종료될 때까지 기다릴 수 있습니다.
                # 필요하다면 Invoke-Expression 실행 후 간단한 알림 메시지를 표시할 수 있습니다.
                # [System.Windows.Forms.MessageBox]::Show("Log Backup 스크립트 실행 완료.", "알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Log Backup 스크립트 실행 중 오류 발생: $($_.Exception.Message)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("워크스테이션을 먼저 선택해주세요.", "알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("backupScript.ps1 파일을 찾을 수 없습니다.", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Audit History ListView 컬럼 설정
if ($auditHistoryListView.Columns.Count -gt 0) {
    $auditHistoryListView.Columns.Clear()
}
$null = $auditHistoryListView.Columns.Add("날짜", 80)
$null = $auditHistoryListView.Columns.Add("감사 항목", 180) # 너비 조정 (선택 사항)
$null = $auditHistoryListView.Columns.Add("감사자", 120) # 너비 조정 (선택 사항)


# 열 너비 변경 이벤트 핸들러
$auditHistoryListView.Add_ColumnWidthChanging({
    param($sender, $e)
    $e.Cancel = $true # 열 너비 변경 취소
})

# ContextMenuStrip 생성
$auditHistoryContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$viewReportMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$viewReportMenuItem.Text = "보고서 보기"
$auditHistoryContextMenu.Items.Add($viewReportMenuItem)

# auditHistoryListView의 ContextMenuStrip 속성에 할당
$auditHistoryListView.ContextMenuStrip = $auditHistoryContextMenu

# "보고서 보기" 메뉴 아이템 클릭 이벤트 핸들러 연결 (아래 2번에서 구현)
$viewReportMenuItem.Add_Click({
    if ($auditHistoryListView.SelectedItems.Count -eq 1) {
        $selectedAuditItem = $auditHistoryListView.SelectedItems[0]
        $auditFolderName = $selectedAuditItem.SubItems[0].Text
        $selectedComputer = $computerComboBox.SelectedItem

        if ($selectedComputer -and $auditFolderName) {
            $scriptRootWithoutScript = Split-Path -parent $PSScriptRoot
            $auditFolderPath = Join-Path -Path (Join-Path -Path $scriptRootWithoutScript -ChildPath "Report\$selectedComputer") -ChildPath $auditFolderName

            if (Test-Path $auditFolderPath -PathType Container) {
                # HTML 보고서 생성 및 열기 함수 호출
                $htmlReportPath = Generate-HTMLAuditReport -AuditFolderPath $auditFolderPath -ComputerName $selectedComputer
                if ($htmlReportPath) {
                    # 웹 브라우저로 HTML 파일 열기
                    try {
                        Start-Process -FilePath $htmlReportPath
                    } catch {
                        [System.Windows.Forms.MessageBox]::Show("웹 브라우저를 여는 동안 오류가 발생했습니다: $($_.Exception.Message)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    }
                }
            } else {
                Write-Host "폴더 존재 확인: False"
                [System.Windows.Forms.MessageBox]::Show("해당 감사 폴더를 찾을 수 없습니다: $auditFolderPath", "알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("감사 기록을 선택하거나 워크스테이션을 선택해주세요.", "알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("하나의 감사 기록을 선택해주세요.", "알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
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

# "접근이력검사" 버튼 생성
$accessHistoryButton = New-Object System.Windows.Forms.Button
$accessHistoryButton.Size = New-Object System.Drawing.Size(100, 25)
$accessHistoryButton.Text = "접근이력"
$accessHistoryButton.Enabled = $false # 기본적으로 비활성화

# 버튼 위치 설정 (accountListLabel 및 accountListListView 우측 끝에 맞춤)
$accessHistoryButton.Location = New-Object System.Drawing.Point(
    [int]($accountListListView.Right - $accessHistoryButton.Width ), # ListView 우측에서 약간 떨어진 위치
    $accountListLabel.Location.Y # accountListLabel과 동일한 Y축
)
$securityStatusPanel.Controls.Add($accessHistoryButton)
$accessHistoryButton.Add_Click({ Show-AccessHistoryLog })

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
#$buttonPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$buttonPanel.BackColor = [System.Drawing.Color]::Transparent # 배경 투명 설정
$contentPanel.Controls.Add($buttonPanel)


# ComboBox 생성
$compFolderComboBox = New-Object System.Windows.Forms.ComboBox
$compFolderComboBox.DropDownStyle = "DropDownList" # 입력 불가 설정
$compFolderComboBox.Size = New-Object System.Drawing.Size(200, 25)

# 폴더 목록 가져오기
$compFolderPath = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "comp"
if (Test-Path -Path $compFolderPath -PathType Container) {
    $compFolders = Get-ChildItem -Path $compFolderPath -Directory | Select-Object -ExpandProperty Name
    $compFolderComboBox.Items.AddRange($compFolders)
}

# Comp 선택 ComboBox 값 변경 이벤트 핸들러
$compFolderComboBox.Add_SelectedIndexChanged({
    Update-AuditPerformButtonState
})

$buttonPanel.Controls.Add($compFolderComboBox)

# 감사수행 버튼 생성
$auditPerformButton = New-Object System.Windows.Forms.Button
$auditPerformButton.Size = New-Object System.Drawing.Size(100, 25)
$auditPerformButton.Text = "감사수행"
$buttonPanel.Controls.Add($auditPerformButton)
$auditPerformButton.Text = "감사수행"
$auditPerformButton.Enabled = $false

# 감사 수행 버튼 클릭 이벤트 핸들러 (수정됨 - MainPanel.ps1)
$auditPerformButton.Add_Click({
    if ($computerComboBox.SelectedItem) {
        $targetComputer = $computerComboBox.SelectedItem
        $accessHistoryButton.Enabled = $true
        if ($compFolderComboBox.SelectedItem) {
            $selectedCompSetName = $compFolderComboBox.SelectedItem
            $loggedInUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name # 로그인 계정명 가져오기
            $compSetPath = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "comp\$selectedCompSetName"

            Write-Host "감사 대상: $targetComputer"
            Write-Host "Comp Set 경로: $compSetPath"
            Write-Host "감사를 시작합니다..."
            
            # 현재 날짜와 시간으로 폴더 이름 생성 (YYYYMMDD_HHMMSS 형식)
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $reportSubFolderPath = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "Report\$targetComputer\$timestamp"

            try {

                # 보고서 폴더 생성
                if (-not (Test-Path $reportSubFolderPath -PathType Container)) {
                    New-Item -Path $reportSubFolderPath -ItemType Directory -Force | Out-Null
                    Write-Host "보고서 폴더 생성: $reportSubFolderPath"
                }

                # InspectionModule.ps1 스크립트 로드
                . "$PSScriptRoot\InspectionModule.ps1"
                
                # Comp Set 폴더 내의 모든 JSON 파일 경로 가져오기 (info.json 제외)
                $jsonFiles = Get-ChildItem -Path $compSetPath -Filter "*.json" | Where-Object {$_.BaseName -ne "info"} | Select-Object -ExpandProperty FullName
                
                # Start-SecurityScan 함수 호출 (Compliance 이름 파라미터 추가)
                Start-SecurityScan -JsonFiles $jsonFiles -ReportOutputPath $reportSubFolderPath -ComplianceName $selectedCompSetName -LoggedInUser $loggedInUser

                Write-Host "감사가 완료되었습니다."

                # 계정 및 서비스 ListView 업데이트
                $accountJsonPath = Join-Path $reportSubFolderPath "account.json"
                $serviceJsonPath = Join-Path $reportSubFolderPath "service.json"
                $mediaJsonPath = Join-Path $reportSubFolderPath "media.json"
                Update-AccountListView -JsonFilePath $accountJsonPath -ListView $accountListListView
                Update-ServiceListView -JsonFilePath $serviceJsonPath -ListView $serviceListListView
                Update-ApprovedMediaListView -JsonFilePath $mediaJsonPath -ListView $approvedMediaListView

                # Audit History ListView 갱신
                Load-AuditHistory -WsName $targetComputer
                
                # 감사 수행 후 감사 일자 레이블 업데이트
                $auditDateLabel.Text = "* 현황 기준 일자: $timestamp"

            } catch {
                Write-Error "감사 수행 중 오류 발생: $($_.Exception.Message)"
                [System.Windows.Forms.MessageBox]::Show("감사 수행 중 오류가 발생했습니다: $($_.Exception.Message)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }

        } else {
            [System.Windows.Forms.MessageBox]::Show("Comp Set을 선택해주세요.", "알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("감사 대상을 선택해주세요.", "알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

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

# root 폴더 경로 설정
$compFolder = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "comp"

$global:compListFormInstance = $null
$global:compListViewInstance = $null

function Load-CompListIntoListView ($listView) {
    $listView.Items.Clear()
    $compFolderPath = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "comp"
    if (Test-Path $compFolderPath -PathType Container) {
        $compFolders = Get-ChildItem -Path $compFolderPath -Directory
        foreach ($folder in $compFolders) {
            $folderName = $folder.Name
            $author = ""
            $itemCount = 0
            $infoFilePath = Join-Path $folder.FullName "info.json"
            if (Test-Path $infoFilePath -PathType Leaf) {
                try {
                    $info = Get-Content $infoFilePath | ConvertFrom-Json
                    $author = $info.author
                } catch {
                    Write-Warning "info.json 파일 읽기 오류: $($_.Exception.Message) - 폴더: $folderName"
                }
            }
            # 해당 폴더의 JSON 파일 개수 세기 (info.json 제외)
            $jsonFiles = Get-ChildItem -Path $folder.FullName -Filter "*.json" | Where-Object {$_.Name -ne "info.json"}
            $itemCount = $jsonFiles.Count
            $listItem = New-Object System.Windows.Forms.ListViewItem($folderName)
            $listItem.SubItems.Add($author)
            $listItem.SubItems.Add($itemCount)
            $null = $listView.Items.Add($listItem)
        }
        $listView.AutoResizeColumns("HeaderSize")
    } else {
        [System.Windows.Forms.MessageBox]::Show("comp 폴더를 찾을 수 없습니다.", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}
# 새로운 Compilation Set 생성 및 수정
function Show-NewCompSetForm {
    # 새로운 폼 생성
    $newCompSetForm = New-Object System.Windows.Forms.Form
    $newCompSetForm.Text = "새로운 Compilation Set 생성"
    $newCompSetForm.Size = New-Object System.Drawing.Size(350, 240)
    $newCompSetForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    # 제목 레이블 및 텍스트 박스
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(10, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(60, 20)
    $titleLabel.Text = "제목:"
    $newCompSetForm.Controls.Add($titleLabel)

    $titleTextBox = New-Object System.Windows.Forms.TextBox
    $titleTextBox.Location = New-Object System.Drawing.Point(80, 20)
    $titleTextBox.Size = New-Object System.Drawing.Size(240, 20)
    $newCompSetForm.Controls.Add($titleTextBox)

    # 작성자 레이블 및 텍스트 박스
    $authorLabel = New-Object System.Windows.Forms.Label
    $authorLabel.Location = New-Object System.Drawing.Point(10, 50)
    $authorLabel.Size = New-Object System.Drawing.Size(60, 20)
    $authorLabel.Text = "작성자:"
    $newCompSetForm.Controls.Add($authorLabel)

    $authorTextBox = New-Object System.Windows.Forms.TextBox
    $authorTextBox.Location = New-Object System.Drawing.Point(80, 50)
    $authorTextBox.Size = New-Object System.Drawing.Size(240, 20)
    $newCompSetForm.Controls.Add($authorTextBox)

    # 기반 레이블 및 텍스트 박스
    $baseLabel = New-Object System.Windows.Forms.Label
    $baseLabel.Location = New-Object System.Drawing.Point(10, 80)
    $baseLabel.Size = New-Object System.Drawing.Size(60, 20)
    $baseLabel.Text = "기반:"
    $newCompSetForm.Controls.Add($baseLabel)

    $baseTextBox = New-Object System.Windows.Forms.TextBox
    $baseTextBox.Location = New-Object System.Drawing.Point(80, 80)
    $baseTextBox.Size = New-Object System.Drawing.Size(240, 20)
    $newCompSetForm.Controls.Add($baseTextBox)

    # 동작 OS 레이블 및 텍스트 박스
    $osLabel = New-Object System.Windows.Forms.Label
    $osLabel.Location = New-Object System.Drawing.Point(10, 110)
    $osLabel.Size = New-Object System.Drawing.Size(60, 20)
    $osLabel.Text = "동작 OS:"
    $newCompSetForm.Controls.Add($osLabel)

    $osTextBox = New-Object System.Windows.Forms.TextBox
    $osTextBox.Location = New-Object System.Drawing.Point(80, 110)
    $osTextBox.Size = New-Object System.Drawing.Size(240, 20)
    $newCompSetForm.Controls.Add($osTextBox)

    # 생성 버튼
    $createButton = New-Object System.Windows.Forms.Button
    $createButton.Location = New-Object System.Drawing.Point(80, 160)
    $createButton.Size = New-Object System.Drawing.Size(80, 30)
    $createButton.Text = "생성"
    $newCompSetForm.Controls.Add($createButton)

    # 취소 버튼
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(180, 160)
    $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
    $cancelButton.Text = "취소"
    $newCompSetForm.Controls.Add($cancelButton)

    # 생성 버튼 클릭 이벤트 핸들러
    $createButton.Add_Click({
        $title = $titleTextBox.Text
        $author = $authorTextBox.Text
        $base = $baseTextBox.Text
        $os = $osTextBox.Text

        # 제목 유효성 검사 (비어 있는지 확인)
        if ([string]::IsNullOrEmpty($title)) {
            [System.Windows.Forms.MessageBox]::Show("제목을 입력해주세요.", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        # 폴더 이름으로 사용할 경로 생성
        $newFolderPath = Join-Path $compFolder $title

        # 폴더가 이미 존재하는지 확인
        if (Test-Path $newFolderPath -PathType Container) {
            [System.Windows.Forms.MessageBox]::Show("이미 존재하는 제목입니다.", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        # 폴더 생성
        try {
            New-Item -Path $newFolderPath -ItemType Directory -Force | Out-Null
            Write-Host "폴더 생성 완료: $newFolderPath"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("폴더 생성 실패: $($_.Exception.Message)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        # info.json에 저장할 데이터 생성
        $infoData = @{
            "title" = $title;
            "author" = $author;
            "base" = $base;
            "os" = $os
        } | ConvertTo-Json -Depth 2

        # info.json 파일 경로 생성
        $infoFilePath = Join-Path $newFolderPath "info.json"

        # info.json 파일 저장
        try {
            $infoData | Out-File -FilePath $infoFilePath -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Compilation Set 생성 완료!", "알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $newCompSetForm.Close()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("info.json 파일 저장 실패: $($_.Exception.Message)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            # 파일 저장 실패 시 생성된 폴더 삭제 (선택 사항)
            Remove-Item -Path $newFolderPath -Force -Recurse
        }
    })

    # 취소 버튼 클릭 이벤트 핸들러
    $cancelButton.Add_Click({
        $newCompSetForm.Close()
    })

    # 폼 표시
    $newCompSetForm.ShowDialog() | Out-Null
}
function Load-CompListIntoListView ($listView) {
    $listView.Items.Clear()
    $compFolderPath = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "comp"
    if (Test-Path $compFolderPath -PathType Container) {
        $compFolders = Get-ChildItem -Path $compFolderPath -Directory
        foreach ($folder in $compFolders) {
            $folderName = $folder.Name
            $author = ""
            $itemCount = 0
            $infoFilePath = Join-Path $folder.FullName "info.json"
            if (Test-Path $infoFilePath -PathType Leaf) {
                try {
                    $info = Get-Content $infoFilePath | ConvertFrom-Json
                    $author = $info.author
                } catch {
                    Write-Warning "info.json 파일 읽기 오류: $($_.Exception.Message) - 폴더: $folderName"
                }
            }
            # 해당 폴더의 JSON 파일 개수 세기 (info.json 제외)
            $jsonFiles = Get-ChildItem -Path $folder.FullName -Filter "*.json" | Where-Object {$_.Name -ne "info.json"}
            $itemCount = $jsonFiles.Count
            $listItem = New-Object System.Windows.Forms.ListViewItem($folderName)
            $listItem.SubItems.Add($author)
            $listItem.SubItems.Add($itemCount)
            $null = $listView.Items.Add($listItem)
        }
        $listView.AutoResizeColumns("HeaderSize")
    } else {
        [System.Windows.Forms.MessageBox]::Show("comp 폴더를 찾을 수 없습니다.", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}
function Show-CompListForm {
    if ($global:compListFormInstance -ne $null -and !$global:compListFormInstance.IsDisposed) {
        # 폼이 이미 열려 있다면 목록만 갱신
        Load-CompListIntoListView $global:compListViewInstance
        return
    }

    # 새로운 폼 생성
    $compListForm = New-Object System.Windows.Forms.Form
    $global:compListFormInstance = $compListForm # 폼 인스턴스 저장
    $compListForm.Text = "Compilation Set 목록"
    $compListForm.Size = New-Object System.Drawing.Size(600, 400)
    $compListForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $compListForm.Padding = New-Object System.Windows.Forms.Padding(10) # 폼에 패딩 추가

    # ListView 컨트롤 생성
    $compListViewInForm = New-Object System.Windows.Forms.ListView
    $global:compListViewInstance = $compListViewInForm # ListView 인스턴스 저장
    $compListViewInForm.Dock = "Fill"
    $compListViewInForm.View = "Details"
    $compListViewInForm.FullRowSelect = $true
    $compListViewInForm.GridLines = $true

    # 컬럼 추가
    $null = $compListViewInForm.Columns.Add("Comp", -1)
    $null = $compListViewInForm.Columns.Add("작성자", -1)
    $null = $compListViewInForm.Columns.Add("검사항목", -1)

    # ContextMenuStrip 생성
    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

    # 이름 변경 메뉴 아이템
    $renameMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("이름 변경")
    $renameMenuItem.Add_Click({
        if ($compListViewInForm.SelectedItems.Count -eq 1) {
            $selectedItem = $compListViewInForm.SelectedItems[0]
            $oldName = $selectedItem.Text
            # 사용자에게 새 이름을 입력받는 InputBox 표시 (MessageBox 활용)
            $newNameResult = [System.Windows.Forms.MessageBox]::Show("새로운 Comp 이름을 입력하세요:", "이름 변경", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Question, [System.Windows.Forms.MessageBoxDefaultButton]::Button1, 0, $oldName)
            if ($newNameResult -eq [System.Windows.Forms.DialogResult]::OK) {
                $newName = $InputBox.Text # InputBox.Text가 아닌 MessageBox에서 입력을 가져오는 것은 제한적입니다.

                # 임시적인 해결 방법: MessageBox의 텍스트를 직접 가져올 수 없으므로,
                # 사용자가 MessageBox에 입력한 내용을 $Input 변수에 저장하도록 유도하거나,
                # 사용자 정의 InputBox 폼을 만들어야 합니다.
                $userInput = Read-Host -Prompt "새로운 Comp 이름" # 콘솔에서 입력 받기 (임시)

                if (-not [string]::IsNullOrEmpty($userInput)) {
                    $newFolderName = $userInput
                    $oldFolderPath = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "comp" -ChildPath $oldName
                    $newFolderPath = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "comp" -ChildPath $newFolderName
                    if (Test-Path $newFolderPath) {
                        [System.Windows.Forms.MessageBox]::Show("이미 존재하는 Comp 이름입니다.", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        return
                    }
                    try {
                        Rename-Item -Path $oldFolderPath -NewName $newFolderName
                        Load-CompListIntoListView $compListViewInForm
                    } catch {
                        [System.Windows.Forms.MessageBox]::Show("이름 변경 실패: $($_.Exception.Message)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    }
                }
            }
        } elseif ($compListViewInForm.SelectedItems.Count -gt 1) {
            [System.Windows.Forms.MessageBox]::Show("하나의 Comp만 선택하여 이름을 변경할 수 있습니다.", "알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show("Comp을 선택해주세요.", "알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })

    # 삭제 메뉴 아이템
    $deleteMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("삭제")
    $deleteMenuItem.Add_Click({
        if ($compListViewInForm.SelectedItems.Count -gt 0) {
            $result = [System.Windows.Forms.MessageBox]::Show("선택한 Comp을 삭제하시겠습니까?", "삭제 확인", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                foreach ($selectedItem in $compListViewInForm.SelectedItems) {
                    # 선택된 ListView 아이템의 첫 번째 컬럼(Comp 이름)을 가져옵니다.
                    $compNameToDelete = $selectedItem.Text
                    # 삭제할 폴더의 전체 경로를 생성합니다.
                    $folderToDelete = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "comp\$compNameToDelete"
                    try {
                        Remove-Item -Path $folderToDelete -Force -Recurse
                    } catch {
                        [System.Windows.Forms.MessageBox]::Show("삭제 실패: $($_.Exception.Message) - Comp: $($compNameToDelete)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    }
                }
                Load-CompListIntoListView $compListViewInForm
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("삭제할 Comp을 선택해주세요.", "알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })

    # 새로운 보안 설정 만들기 메뉴 아이템 추가
    $newSecurityItemMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("새로운 보안 설정 만들기")
    $newSecurityItemMenuItem.Add_Click({
        if ($global:compListViewInstance.SelectedItems.Count -eq 1) { # ListViewInstance 사용
            $selectedCompName = $global:compListViewInstance.SelectedItems[0].Text # ListViewInstance 사용
            $targetFolderPath = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "comp\$selectedCompName"

            # CreateJson.ps1 스크립트 내용을 현재 프로세스에서 실행 (콘솔 창 안 보임)
            . "$PSScriptRoot\CreateJson.ps1" -OutputPath $targetFolderPath

            # 저장 후 Comp 목록 갱신
            if ($global:compListViewInstance -ne $null -and !$global:compListViewInstance.IsDisposed) {
                Load-CompListIntoListView $global:compListViewInstance
            }

        } elseif ($global:compListViewInstance.SelectedItems.Count -gt 1) { # ListViewInstance 사용
            [System.Windows.Forms.MessageBox]::Show("하나의 Comp만 선택해주세요.", "알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show("Comp을 선택해주세요.", "알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })

    # ContextMenuStrip에 메뉴 아이템 추가
    $null = $contextMenu.Items.AddRange(@($renameMenuItem, $deleteMenuItem, $newSecurityItemMenuItem))

    # ListView의 ContextMenuStrip 속성에 할당
    $compListViewInForm.ContextMenuStrip = $contextMenu

    # ListView를 폼에 추가
    $compListForm.Controls.Add($compListViewInForm)

    # Comp 목록 로드
    Load-CompListIntoListView $compListViewInForm

    # 폼 표시 (모달리스로 변경)
    $null = $compListForm.Show()
}
function Show-AccessHistoryLog {
    # 보안 로그에서 계정 관련 이벤트 검색
    $eventIDs = 4720, 4723, 4724 # 추가적인 계정 관련 이벤트 ID를 포함할 수 있습니다.
    $logEntries = Get-WinEvent -LogName Security -FilterXPath "*[System[EventID=$eventIDs]]" -MaxEvents 100 # 최근 100개 이벤트

    if ($logEntries) {
        # 새로운 폼 생성
        $logForm = New-Object System.Windows.Forms.Form
        $logForm.Text = "계정 접근 이력"
        $logForm.Size = New-Object System.Drawing.Size(800, 600)
        $logForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

        # 로그 내용을 표시할 TextBox 생성
        $logTextBox = New-Object System.Windows.Forms.TextBox
        $logTextBox.Multiline = $true
        $logTextBox.ReadOnly = $true
        $logTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
        $logTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill
        $logForm.Controls.Add($logTextBox)

        $logContent = ""
        foreach ($event in $logEntries) {
            $logContent += "Event ID: $($event.Id)`r`n"
            $logContent += "Time Created: $($event.TimeCreated)`r`n"
            $logContent += "Message: $($event.Message)`r`n"
            $logContent += "--------------------------------------------------`r`n"
        }
        $logTextBox.Text = $logContent

        # 새로운 폼 표시
        $null = $logForm.ShowDialog()
    } else {
        # 검색된 로그가 없는 경우 경고 메시지 박스 표시
        [System.Windows.Forms.MessageBox]::Show("계정 관련 이벤트 로그가 없습니다.", "경고", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
}
function Load-AuditHistory {
    param(
        [Parameter(Mandatory=$true)]
        [string]$WsName # 선택된 WS 이름 (폴더 구조 파악에 사용)
    )
    $auditHistoryListView.Items.Clear()
    $wsPath = Join-Path -Path (Split-Path -parent $PSScriptRoot) -ChildPath "Report\$WsName"
    if (Test-Path $wsPath -PathType Container) {
        $auditFolders = Get-ChildItem -Path $wsPath -Directory | Where-Object {$_.Name -match '^\d{8}_\d{6}$'} | Sort-Object -Descending

        foreach ($folder in $auditFolders) {
            # 폴더 이름 전체를 첫 번째 컬럼에 표시
            $listItem = New-Object System.Windows.Forms.ListViewItem($folder.Name)
            $loggedInUser = ""
            $complianceName = ""
            $logFilePath = Join-Path $folder.FullName "security_report.log"

            if (Test-Path $logFilePath) {
                $userLog = Get-Content -Path $logFilePath | Where-Object {$_ -like "* - Logged In User: *"} | Select-Object -First 1
                if ($userLog) {
                    $loggedInUser = $userLog.Substring($userLog.IndexOf(" - Logged In User: ") + 20).Trim()
                }
                $complianceLog = Get-Content -Path $logFilePath | Where-Object {$_ -like "* - Audit Compliance Set: *"} | Select-Object -First 1
                if ($complianceLog) {
                    $complianceName = $complianceLog.Substring($complianceLog.IndexOf(" - Audit Compliance Set: ") + 23).Trim()
                }
            }

            $listItem.SubItems.Add($complianceName)
            $listItem.SubItems.Add($loggedInUser)

            $reportInfo = @{
                "FullPath" = $folder.FullName
            }
            $listItem.Tag = $reportInfo
            $null = $auditHistoryListView.Items.Add($listItem)
        }
    }
}
function Load-LatestAuditData {
    param(
        [Parameter(Mandatory=$true)]
        [string]$WsName,
        [Parameter(Mandatory=$true)]
        [string]$AuditFolderPath
    )

    Write-Host "최근 감사 데이터 로드 시작 - WS: $WsName, 폴더: $AuditFolderPath"

    # Account 정보 로드
    $accountFilePath = Join-Path $AuditFolderPath "account.json" # 파일 이름 수정
    if (Test-Path $accountFilePath) {
        try {
            $accounts = Get-Content -Path $accountFilePath -Raw -Encoding UTF8 | ConvertFrom-Json
            $accountListListView.Items.Clear()
            Write-Host $account
            foreach ($account in $accounts) {
                $listItem = New-Object System.Windows.Forms.ListViewItem($account.SamAccountName)
                $listItem.SubItems.Add($account.Enabled.ToString())
                $null = $accountListListView.Items.Add($listItem)
            }
            Write-Host "Account 정보 로드 완료"
        } catch {
            Write-Error "Account 정보 로드 실패: $($_.Exception.Message) - 파일: $accountFilePath"
            $accountListListView.Items.Clear()
        }
    } else {
        Write-Warning "Account 정보 파일이 없습니다: $accountFilePath"
        $accountListListView.Items.Clear()
    }

    # Service 정보 로드
    $serviceFilePath = Join-Path $AuditFolderPath "service.json" # 파일 이름 수정
    if (Test-Path $serviceFilePath) {
        try {
            $services = Get-Content -Path $serviceFilePath -Raw -Encoding UTF8 | ConvertFrom-Json
            $serviceListListView.Items.Clear()
            foreach ($service in $services) {
                $listItem = New-Object System.Windows.Forms.ListViewItem($service.DisplayName)
                $listItem.SubItems.Add($service.Status)
                $null = $serviceListListView.Items.Add($listItem)
            }
            Write-Host "Service 정보 로드 완료"
        } catch {
            Write-Error "Service 정보 로드 실패: $($_.Exception.Message) - 파일: $serviceFilePath"
            $serviceListListView.Items.Clear()
        }
    } else {
        Write-Warning "Service 정보 파일이 없습니다: $serviceFilePath"
        $serviceListListView.Items.Clear()
    }

    # Media 정보 로드
    $mediaFilePath = Join-Path $AuditFolderPath "media.json"
    if (Test-Path $mediaFilePath) {
        try {
            $mediaList = Get-Content -Path $mediaFilePath -Raw -Encoding UTF8 | ConvertFrom-Json
            $approvedMediaListView.Items.Clear()
            foreach ($media in $mediaList) {
                $listItem = New-Object System.Windows.Forms.ListViewItem(@($media.미디어, $media.Reg))
                $null = $approvedMediaListView.Items.Add($listItem)
            }
            Write-Host "Media 정보 로드 완료"
        } catch {
            Write-Error "Media 정보 로드 실패: $($_.Exception.Message) - 파일: $mediaFilePath"
            $approvedMediaListView.Items.Clear()
        }
    } else {
        Write-Warning "Media 정보 파일이 없습니다: $mediaFilePath"
        $approvedMediaListView.Items.Clear()
    }

    Write-Host "최근 감사 데이터 로드 완료"
}

function Generate-HTMLAuditReport {
    param(
        [Parameter(Mandatory=$true)]
        [string]$AuditFolderPath,
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )

    $reportData = @{}
    # 감사 폴더에서 필요한 데이터 추출
    try {
        $accountData = Get-Content -Path (Join-Path $AuditFolderPath "account.json") -Raw -Encoding UTF8 | ConvertFrom-Json
        $serviceData = Get-Content -Path (Join-Path $AuditFolderPath "service.json") -Raw -Encoding UTF8 | ConvertFrom-Json
        # ... 다른 JSON 파일 로드

        $reportData.Account = $accountData
        $reportData.Service = $serviceData
        # ... 다른 데이터 저장

    } catch {
        Write-Error "감사 데이터 로드 실패: $($_.Exception.Message)"
        return $null
    }

    # security_report.log 내용 읽기
    $logFilePath = Join-Path $AuditFolderPath "security_report.log"
    $logContent = ""
    if (Test-Path $logFilePath -PathType Leaf) {
        try {
            $logContent = Get-Content -Path $logFilePath
            $logContent = $logContent -join "<br>" # 각 줄을 <br> 태그로 연결하여 HTML 줄바꿈 처리
        } catch {
            Write-Warning "security_report.log 파일 읽기 실패: $($_.Exception.Message)"
        }
    }

    # HTML 내용 생성
    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>감사 보고서 - $ComputerName - $((Split-Path $AuditFolderPath -Leaf))</title>
    <style>
        body { font-family: Arial, sans-serif; }
        h1, h2 { color: navy; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .log-section { margin-top: 20px; border: 1px solid #ccc; padding: 10px; background-color: #f9f9f9; white-space: pre-wrap; font-family: monospace; font-size: 0.9em; }
        .log-section h2 { color: darkred; }
    </style>
</head>
<body>
    <h1>감사 보고서</h1>
    <h2>워크스테이션: $ComputerName</h2>
    <p>감사 일자: $((Split-Path $AuditFolderPath -Leaf))</p>

    <h2>계정 정보</h2>
    <table>
        <thead>
            <tr>
                <th>SAM 계정명</th>
                <th>이름</th>
                <th>활성화됨</th>
            </tr>
        </thead>
        <tbody>
"@

    if ($reportData.Account) {
        foreach ($account in $reportData.Account) {
            $htmlContent += @"
            <tr>
                <td>$($account.SamAccountName)</td>
                <td>$($account.Name)</td>
                <td>$($account.Enabled)</td>
            </tr>
"@
        }
    } else {
        $htmlContent += @"
            <tr><td colspan="3">계정 정보가 없습니다.</td></tr>
"@
    }

    $htmlContent += @"
        </tbody>
    </table>

    <h2>서비스 정보</h2>
    <table>
        <thead>
            <tr>
                <th>서비스명</th>
                <th>상태</th>
            </tr>
        </thead>
        <tbody>
"@

    if ($reportData.Service) {
        foreach ($service in $reportData.Service) {
            $htmlContent += @"
            <tr>
                <td>$($service.DisplayName)</td>
                <td>$($service.Status)</td>
            </tr>
"@
        }
    } else {
        $htmlContent += @"
            <tr><td colspan="2">서비스 정보가 없습니다.</td></tr>
"@
    }

    $htmlContent += @"
        </tbody>
    </table>

    <div class="log-section">
        <h2>Security Report Log</h2>
        <pre>$logContent</pre>
    </div>

    <p>보고서 생성 일시: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
</body>
</html>
"@

    # HTML 파일 경로
    $htmlReportPath = Join-Path $AuditFolderPath "audit_report.html"

    try {
        $htmlContent | Out-File -FilePath $htmlReportPath -Encoding UTF8
        Write-Host "HTML 보고서 생성 완료: $htmlReportPath"
        return $htmlReportPath
    } catch {
        Write-Error "HTML 보고서 저장 실패: $($_.Exception.Message)"
        return $null
    }
}

# Form 표시
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()