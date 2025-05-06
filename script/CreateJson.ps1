param(
    [string]$OutputPath
)

Add-Type -AssemblyName System.Windows.Forms

# 폼 객체 생성
$form = New-Object System.Windows.Forms.Form
$form.Text = "검사 항목 선택"
$form.Width = 350
$form.Height = 110
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

# 버튼 너비 및 간격 설정
$buttonWidth = 90
$buttonSpacing = 15
$startX = 15
$startY = 20

# 버튼 1: 레지스트리
$buttonRegistry = New-Object System.Windows.Forms.Button
$buttonRegistry.Text = "레지스트리"
$buttonRegistry.Width = $buttonWidth
$buttonRegistry.Location = New-Object System.Drawing.Point($startX, $startY)
$form.Controls.Add($buttonRegistry)
$startX += $buttonWidth + $buttonSpacing

# 버튼 2: 파일 무결성
$buttonIntegrity = New-Object System.Windows.Forms.Button
$buttonIntegrity.Text = "파일무결성"
$buttonIntegrity.Width = $buttonWidth
$buttonIntegrity.Location = New-Object System.Drawing.Point($startX, $startY)
$form.Controls.Add($buttonIntegrity)
$startX += $buttonWidth + $buttonSpacing

# 버튼 3: 보안설정
$buttonSecurity = New-Object System.Windows.Forms.Button
$buttonSecurity.Text = "보안설정검사"
$buttonSecurity.Width = $buttonWidth
$buttonSecurity.Location = New-Object System.Drawing.Point($startX, $startY)
$form.Controls.Add($buttonSecurity)


function Show-RegistryInputForm {
    # 새로운 폼 생성
    $registryForm = New-Object System.Windows.Forms.Form
    $registryForm.Text = "레지스트리 검사 항목 입력"
    $registryForm.Width = 400
    $registryForm.Height = 340 # 높이 변경
    $registryForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    # 레이블 및 텍스트 박스 생성
    $labelTitle = New-Object System.Windows.Forms.Label
    $labelTitle.Text = "제목:"
    $labelTitle.Location = New-Object System.Drawing.Point(20, 20)
    $labelTitle.AutoSize = $true
    $registryForm.Controls.Add($labelTitle)

    $textBoxTitle = New-Object System.Windows.Forms.TextBox
    $textBoxTitle.Location = New-Object System.Drawing.Point(120, 20)
    $textBoxTitle.Width = 240
    $registryForm.Controls.Add($textBoxTitle)

    $labelItem = New-Object System.Windows.Forms.Label
    $labelItem.Text = "검사항목:"
    $labelItem.Location = New-Object System.Drawing.Point(20, 60)
    $labelItem.AutoSize = $true
    $registryForm.Controls.Add($labelItem)

    $textBoxItem = New-Object System.Windows.Forms.TextBox
    $textBoxItem.Location = New-Object System.Drawing.Point(120, 60)
    $textBoxItem.Width = 240
    $registryForm.Controls.Add($textBoxItem)

    $labelDescription = New-Object System.Windows.Forms.Label
    $labelDescription.Text = "설명:"
    $labelDescription.Location = New-Object System.Drawing.Point(20, 100)
    $labelDescription.AutoSize = $true
    $registryForm.Controls.Add($labelDescription)

    $textBoxDescription = New-Object System.Windows.Forms.TextBox
    $textBoxDescription.Location = New-Object System.Drawing.Point(120, 100)
    $textBoxDescription.Multiline = $true
    $textBoxDescription.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $textBoxDescription.Width = 240
    $textBoxDescription.Height = 60
    $registryForm.Controls.Add($textBoxDescription)

    # 비교 연산자 레이블 및 콤보 박스
    $labelOperatorRegistry = New-Object System.Windows.Forms.Label
    $labelOperatorRegistry.Text = "비교 연산자:"
    $labelOperatorRegistry.Location = New-Object System.Drawing.Point(20, 180)
    $labelOperatorRegistry.AutoSize = $true
    $registryForm.Controls.Add($labelOperatorRegistry)

    $comboBoxOperatorRegistry = New-Object System.Windows.Forms.ComboBox
    $comboBoxOperatorRegistry.Items.AddRange(@("Equals", "NotEquals", "GreaterThan", "LessThan", "Contains", "NotContains"))
    $comboBoxOperatorRegistry.Location = New-Object System.Drawing.Point(120, 180)
    $comboBoxOperatorRegistry.Width = 240
    $registryForm.Controls.Add($comboBoxOperatorRegistry)

    # 비교 값 레이블 및 텍스트 박스
    $labelCompareValueRegistry = New-Object System.Windows.Forms.Label
    $labelCompareValueRegistry.Text = "비교 값:"
    $labelCompareValueRegistry.Location = New-Object System.Drawing.Point(20, 220)
    $labelCompareValueRegistry.AutoSize = $true
    $registryForm.Controls.Add($labelCompareValueRegistry)

    $textBoxCompareValueRegistry = New-Object System.Windows.Forms.TextBox
    $textBoxCompareValueRegistry.Location = New-Object System.Drawing.Point(120, 220)
    $textBoxCompareValueRegistry.Width = 240
    $registryForm.Controls.Add($textBoxCompareValueRegistry)

    # 저장 버튼 생성
    $buttonSave = New-Object System.Windows.Forms.Button
    $buttonSave.Text = "저장"
    $buttonSave.Location = New-Object System.Drawing.Point(280, 250) # 위치 조정
    $registryForm.Controls.Add($buttonSave)

    # 저장 버튼 클릭 이벤트 핸들러
    $buttonSave.Add_Click({
        # 저장 기능 구현
        $title = $textBoxTitle.Text
        $item = $textBoxItem.Text
        $description = $textBoxDescription.Text
        $operator = $comboBoxOperatorRegistry.SelectedItem
        $compareValue = $textBoxCompareValueRegistry.Text

        # 제목 유효성 검사 (언더바 제외 특수문자 불가)
        if ($title -match '[^\w\s-]') {
            [System.Windows.Forms.MessageBox]::Show("제목에 언더바(_)를 제외한 특수문자를 사용할 수 없습니다.", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        # 저장 경로 설정 (전달받은 $OutputPath 사용)
        $filePath = Join-Path -Path $OutputPath -ChildPath "$title.json"

        $data = @{
            "type"       = "registry";
            "item"       = $item;
            "desc"       = $description;
            "Operator"   = $operator;
            "CompareValue" = $compareValue
        } | ConvertTo-Json -Depth 3

        try {
            $data | Out-File -FilePath $filePath -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("저장 완료: $filePath", "알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $registryForm.Close() # 저장 후 입력 폼 닫기
        } catch {
            [System.Windows.Forms.MessageBox]::Show("저장 실패: $($_.Exception.Message)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    # 새로운 폼을 모달로 표시
    $registryForm.ShowDialog() | Out-Null
}

# 레지스트리 버튼 클릭 이벤트 핸들러
$buttonRegistry.Add_Click({
    Show-RegistryInputForm
})

# 파일 무결성 파일 선택 다이얼로그를 표시하는 함수
function Show-IntegrityInputForm {
    # 새로운 폼 생성
    $integrityForm = New-Object System.Windows.Forms.Form
    $integrityForm.Text = "파일 무결성 검사 항목 입력"
    $integrityForm.Width = 400
    $integrityForm.Height = 200
    $integrityForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    # 제목 레이블 및 텍스트 박스
    $labelTitleIntegrity = New-Object System.Windows.Forms.Label
    $labelTitleIntegrity.Text = "제목:"
    $labelTitleIntegrity.Location = New-Object System.Drawing.Point(20, 20)
    $labelTitleIntegrity.AutoSize = $true
    $integrityForm.Controls.Add($labelTitleIntegrity)

    $textBoxTitleIntegrity = New-Object System.Windows.Forms.TextBox
    $textBoxTitleIntegrity.Location = New-Object System.Drawing.Point(80, 20)
    $textBoxTitleIntegrity.Width = 280
    $integrityForm.Controls.Add($textBoxTitleIntegrity)

    # 파일 선택 레이블 및 텍스트 박스, 버튼
    $labelFileIntegrity = New-Object System.Windows.Forms.Label
    $labelFileIntegrity.Text = "파일:"
    $labelFileIntegrity.Location = New-Object System.Drawing.Point(20, 60)
    $labelFileIntegrity.AutoSize = $true
    $integrityForm.Controls.Add($labelFileIntegrity)

    $textBoxFileIntegrity = New-Object System.Windows.Forms.TextBox
    $textBoxFileIntegrity.Location = New-Object System.Drawing.Point(80, 60)
    $textBoxFileIntegrity.Width = 180
    $integrityForm.Controls.Add($textBoxFileIntegrity)
    $textBoxFileIntegrity.ReadOnly = $true # 직접 수정 방지

    $buttonBrowseFile = New-Object System.Windows.Forms.Button
    $buttonBrowseFile.Text = "찾아보기"
    $buttonBrowseFile.Location = New-Object System.Drawing.Point(270, 60)
    $buttonBrowseFile.Width = 90
    $integrityForm.Controls.Add($buttonBrowseFile)

    # 저장 버튼
    $buttonSaveIntegrity = New-Object System.Windows.Forms.Button
    $buttonSaveIntegrity.Text = "저장"
    $buttonSaveIntegrity.Location = New-Object System.Drawing.Point(280, 110)
    $buttonBrowseFile.Width = 90
    $integrityForm.Controls.Add($buttonSaveIntegrity)

    # "찾아보기" 버튼 클릭 이벤트 핸들러
    $buttonBrowseFile.Add_Click({
        $openFileDialogIntegrity = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialogIntegrity.Title = "검사할 파일 선택"
        $openFileDialogIntegrity.Filter = "모든 파일 (*.*)|*.*"
        $openFileDialogIntegrity.Multiselect = $false # 단일 파일 선택

        if ($openFileDialogIntegrity.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $textBoxFileIntegrity.Text = $openFileDialogIntegrity.FileName
        }
    })

    # "저장" 버튼 클릭 이벤트 핸들러
    $buttonSaveIntegrity.Add_Click({
        $title = $textBoxTitleIntegrity.Text
        $selectedFile = $textBoxFileIntegrity.Text

        # 제목 유효성 검사 (언더바 제외 특수문자 불가)
        if ($title -match '[^\w\s-]') {
            [System.Windows.Forms.MessageBox]::Show("제목에 언더바(_)를 제외한 특수문자를 사용할 수 없습니다.", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        if (-not $selectedFile) {
            [System.Windows.Forms.MessageBox]::Show("검사할 파일을 선택해주세요.", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        # 파일의 SHA256 해시값 계산
        try {
            $fileHash = Get-FileHash -Path $selectedFile -Algorithm SHA256
            $hashValue = $fileHash.Hash
        } catch {
            [System.Windows.Forms.MessageBox]::Show("파일 해시값 계산 오류: $($_.Exception.Message)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        # 저장 위치 선택 다이얼로그 표시
        $saveFileDialogIntegrity = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialogIntegrity.Filter = "JSON 파일 (*.json)|*.json|모든 파일 (*.*)|*.*"
        $saveFileDialogIntegrity.Title = "저장 위치 선택"
        $saveFileDialogIntegrity.FileName = "$title.json"

        if ($saveFileDialogIntegrity.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $filePath = $saveFileDialogIntegrity.FileName
            $data = @{
                "type" = "file_integrity";
                "title" = $title;
                "file" = $selectedFile
                "hashValue" = $hashValue
            } | ConvertTo-Json -Depth 3

            try {
                $data | Out-File -FilePath $filePath -Encoding UTF8
                [System.Windows.Forms.MessageBox]::Show("저장 완료: $filePath", "알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $integrityForm.Close() # 저장 후 입력 폼 닫기
            } catch {
                [System.Windows.Forms.MessageBox]::Show("저장 실패: $($_.Exception.Message)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })

    # 새로운 폼을 모달로 표시
    $integrityForm.ShowDialog() | Out-Null
}

# 파일 무결성 버튼 클릭 이벤트 핸들러
$buttonIntegrity.Add_Click({
    Show-IntegrityInputForm
})

# 보안 설정 검사 항목 입력 폼을 표시하는 함수
# 보안 설정 검사 항목 입력 폼을 표시하는 함수
function Show-SecurityPolicyInputForm {
    # 새로운 폼 생성
    $securityPolicyForm = New-Object System.Windows.Forms.Form
    $securityPolicyForm.Text = "보안 설정 검사 항목 입력"
    $securityPolicyForm.Width = 400
    $securityPolicyForm.Height = 340 # 높이 변경
    $securityPolicyForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    # 레이블 및 텍스트 박스 생성 (레지스트리 폼과 동일)
    $labelTitleSecurity = New-Object System.Windows.Forms.Label
    $labelTitleSecurity.Text = "제목:"
    $labelTitleSecurity.Location = New-Object System.Drawing.Point(20, 20)
    $labelTitleSecurity.AutoSize = $true
    $securityPolicyForm.Controls.Add($labelTitleSecurity)

    $textBoxTitleSecurity = New-Object System.Windows.Forms.TextBox
    $textBoxTitleSecurity.Location = New-Object System.Drawing.Point(120, 20)
    $textBoxTitleSecurity.Width = 240
    $securityPolicyForm.Controls.Add($textBoxTitleSecurity)

    $labelItemSecurity = New-Object System.Windows.Forms.Label
    $labelItemSecurity.Text = "검사항목:"
    $labelItemSecurity.Location = New-Object System.Drawing.Point(20, 60)
    $labelItemSecurity.AutoSize = $true
    $securityPolicyForm.Controls.Add($labelItemSecurity)

    $textBoxItemSecurity = New-Object System.Windows.Forms.TextBox
    $textBoxItemSecurity.Location = New-Object System.Drawing.Point(120, 60)
    $textBoxItemSecurity.Width = 240
    $securityPolicyForm.Controls.Add($textBoxItemSecurity)

    $labelDescriptionSecurity = New-Object System.Windows.Forms.Label
    $labelDescriptionSecurity.Text = "설명:"
    $labelDescriptionSecurity.Location = New-Object System.Drawing.Point(20, 100)
    $labelDescriptionSecurity.AutoSize = $true
    $securityPolicyForm.Controls.Add($labelDescriptionSecurity)

    $textBoxDescriptionSecurity = New-Object System.Windows.Forms.TextBox
    $textBoxDescriptionSecurity.Location = New-Object System.Drawing.Point(120, 100)
    $textBoxDescriptionSecurity.Multiline = $true
    $textBoxDescriptionSecurity.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $textBoxDescriptionSecurity.Width = 240
    $textBoxDescriptionSecurity.Height = 60
    $securityPolicyForm.Controls.Add($textBoxDescriptionSecurity)

    # 비교 연산자 레이블 및 콤보 박스
    $labelOperatorSecurity = New-Object System.Windows.Forms.Label
    $labelOperatorSecurity.Text = "비교 연산자:"
    $labelOperatorSecurity.Location = New-Object System.Drawing.Point(20, 180)
    $labelOperatorSecurity.AutoSize = $true
    $securityPolicyForm.Controls.Add($labelOperatorSecurity)

    $comboBoxOperatorSecurity = New-Object System.Windows.Forms.ComboBox
    $comboBoxOperatorSecurity.Items.AddRange(@("Equals", "NotEquals", "GreaterThan", "LessThan", "Contains", "NotContains"))
    $comboBoxOperatorSecurity.Location = New-Object System.Drawing.Point(120, 180)
    $comboBoxOperatorSecurity.Width = 240
    $securityPolicyForm.Controls.Add($comboBoxOperatorSecurity)

    # 비교 값 레이블 및 텍스트 박스
    $labelCompareValueSecurity = New-Object System.Windows.Forms.Label
    $labelCompareValueSecurity.Text = "비교 값:"
    $labelCompareValueSecurity.Location = New-Object System.Drawing.Point(20, 220)
    $labelCompareValueSecurity.AutoSize = $true
    $securityPolicyForm.Controls.Add($labelCompareValueSecurity)

    $textBoxCompareValueSecurity = New-Object System.Windows.Forms.TextBox
    $textBoxCompareValueSecurity.Location = New-Object System.Drawing.Point(120, 220)
    $textBoxCompareValueSecurity.Width = 240
    $securityPolicyForm.Controls.Add($textBoxCompareValueSecurity)

    # 저장 버튼 생성
    $buttonSaveSecurity = New-Object System.Windows.Forms.Button
    $buttonSaveSecurity.Text = "저장"
    $buttonSaveSecurity.Location = New-Object System.Drawing.Point(280, 250) # 위치 조정
    $securityPolicyForm.Controls.Add($buttonSaveSecurity)

    # 저장 버튼 클릭 이벤트 핸들러 (type만 변경)
    $buttonSaveSecurity.Add_Click({
        # 저장 기능 구현
        $title = $textBoxTitleSecurity.Text
        $item = $textBoxItemSecurity.Text
        $description = $textBoxDescriptionSecurity.Text
        $operator = $comboBoxOperatorSecurity.SelectedItem
        $compareValue = $textBoxCompareValueSecurity.Text

        # 제목 유효성 검사 (언더바 제외 특수문자 불가)
        if ($title -match '[^\w\s-]') {
            [System.Windows.Forms.MessageBox]::Show("제목에 언더바(_)를 제외한 특수문자를 사용할 수 없습니다.", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        # 저장 위치 선택 다이얼로그 표시
        $saveFileDialogSecurity = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialogSecurity.Filter = "JSON 파일 (*.json)|*.json|모든 파일 (*.*)|*.*"
        $saveFileDialogSecurity.Title = "저장 위치 선택"
        $saveFileDialogSecurity.FileName = "$title.json"

        if ($saveFileDialogSecurity.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $filePath = $saveFileDialogSecurity.FileName
            $data = @{
                "type"       = "security_policy"; # type 변경
                "item" = $item;
                "desc"     = $description;
                "Operator"   = $operator;
                "CompareValue" = $compareValue
            } | ConvertTo-Json -Depth 3

            try {
                $data | Out-File -FilePath $filePath -Encoding UTF8
                [System.Windows.Forms.MessageBox]::Show("저장 완료: $filePath", "알림", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $securityPolicyForm.Close() # 저장 후 입력 폼 닫기
            } catch {
                [System.Windows.Forms.MessageBox]::Show("저장 실패: $($_.Exception.Message)", "오류", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })

    # 새로운 폼을 모달로 표시
    $securityPolicyForm.ShowDialog() | Out-Null
}
# 보안설정 버튼 클릭 이벤트 핸들러
$buttonSecurity.Add_Click({
    Show-SecurityPolicyInputForm
})

# 폼 표시
$form.ShowDialog()