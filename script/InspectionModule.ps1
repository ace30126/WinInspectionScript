# 보안 검사를 시작하는 함수
function Start-SecurityScan {
    param(
        [Parameter(Mandatory=$true)]
        [string[]]$JsonFiles,
        [Parameter(Mandatory=$true)]
        [string]$ReportOutputPath,
        [Parameter(Mandatory=$true)]
        [string]$ComplianceName,
        [Parameter(Mandatory=$true)]
        [string]$LoggedInUser # 새로운 파라미터
    )

    Write-Host "보안 감사 시작..."

    if (-not $ReportOutputPath) {
        Write-Warning "보고서 출력 폴더가 설정되지 않았습니다. 콘솔에만 결과를 출력합니다."
        return
    }

    $logFilePath = Join-Path $ReportOutputPath "security_report.log"
    "[$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] - 보안 감사 시작" | Out-File -FilePath $logFilePath -Encoding UTF8 -Append
    
    # 감사 대상 Compliance 이름 로그에 기록
    "[$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] - Audit Compliance Set: $ComplianceName" | Out-File -FilePath $logFilePath -Encoding UTF8 -Append

    # 로그인 계정명 로그에 기록 (키워드 변경)
    "[$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] - Logged In User: $LoggedInUser" | Out-File -FilePath $logFilePath -Encoding UTF8 -Append


    # 현재 계정 정보 추출 및 저장
    $accounts = Get-CurrentAccounts
    Save-ToJson -Data $accounts -OutputPath $ReportOutputPath -FileName "account.json"
    Write-Host "현재 계정 정보가 $($ReportOutputPath)\account.json에 저장되었습니다."

    # 현재 서비스 정보 추출 및 저장
    $services = Get-CurrentServices
    Save-ToJson -Data $services -OutputPath $ReportOutputPath -FileName "service.json"
    Write-Host "현재 서비스 정보가 $($ReportOutputPath)\service.json에 저장되었습니다."

    # Approved Media 정보 추출 및 저장 (새로운 코드)
    $approvedMedia = Get-ApprovedMedia
    Save-ToJson -Data $approvedMedia -OutputPath $ReportOutputPath -FileName "media.json"
    Write-Host "승인된 미디어 정보가 $($ReportOutputPath)\media.json에 저장되었습니다."


    foreach ($jsonFile in $JsonFiles) {
        try {
            Write-Host "파일 읽는 중: $jsonFile"
            $content = Get-Content -Path $jsonFile -Raw -Encoding UTF8
            $data = ConvertFrom-Json -InputObject $content

            switch ($data.type) {
                "registry" {
                    Write-Host "레지스트리 검사 항목 발견: $($data.desc)"
                    $registryResult = Test-RegistrySetting -Item $data.item -Operator $data.Operator -CompareValue $data.CompareValue
                    Write-Host "  결과:" ($registryResult | ConvertTo-Json)
                    "[$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] - 레지스트리 검사: $($data.desc) - 결과: $($registryResult | ConvertTo-Json)" | Out-File -FilePath $logFilePath -Encoding UTF8 -Append
                }
                "file_integrity" {
                    Write-Host "파일 무결성 검사 항목 발견: $($data.title) - $($data.file)"
                    $integrityResult = Test-FileIntegrity -FilePath $data.file -ExpectedHash $data.hashValue
                    Write-Host "  결과:" ($integrityResult | ConvertTo-Json)
                    "[$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] - 파일 무결성 검사: $($data.title) - 파일: $($data.file) - 결과: $($integrityResult | ConvertTo-Json)" | Out-File -FilePath $logFilePath -Encoding UTF8 -Append
                }
                "security_policy" {
                    Write-Host "보안 설정 검사 항목 발견: $($data.desc)"
                    $securityPolicyResult = Test-SecurityPolicy -Item $data.item -Operator $data.Operator -CompareValue $data.CompareValue
                    Write-Host "    결과:" ($securityPolicyResult | ConvertTo-Json)
                    "[$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] - 보안 설정 검사: $($data.desc) - 결과: $($securityPolicyResult | ConvertTo-Json)" | Out-File -FilePath $logFilePath -Encoding UTF8 -Append
                }
                default {
                    Write-Warning "알 수 없는 검사 유형: $($data.type) - 파일: $jsonFile"
                    "[$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] - 알 수 없는 검사 유형: $($data.type) - 파일: $jsonFile" | Out-File -FilePath $logFilePath -Encoding UTF8 -Append
                }
            }
        } catch {
            Write-Error "JSON 파일 처리 오류: $($_.Exception.Message) - 파일: $jsonFile"
            "[$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] - JSON 파일 처리 오류: $($_.Exception.Message) - 파일: $jsonFile" | Out-File -FilePath $logFilePath -Encoding UTF8 -Append
        }
    }

    "[$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] - 검사 완료." | Out-File -FilePath $logFilePath -Encoding UTF8 -Append
    Write-Host "검사 완료. 결과는 $($logFilePath)에 기록되었습니다."
}

function Test-RegistrySetting {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Item,

        [Parameter(Mandatory=$false)]
        [string]$Operator, # 비교 연산자 (Equals, NotEquals, GreaterThan, LessThan, Contains, NotContains)

        [Parameter(Mandatory=$false)]
        [string]$CompareValue
    )

    $keyPath = ""
    $valueName = $null

    # 마지막 백슬래시 위치를 찾음
    $lastBackslashIndex = $Item.LastIndexOf('\')

    if ($lastBackslashIndex -gt 4) { # 'HKLM:' 또는 'HKCU:' 길이보다 커야 값 이름이 존재할 가능성이 있음
        $keyPath = $Item.Substring(0, $lastBackslashIndex)
        $valueName = $Item.Substring($lastBackslashIndex + 1)
    } else {
        $keyPath = $Item
    }

    
    # Registry:: 접두사 제거
    $registryValue = Get-ItemProperty -Path $keyPath -Name $valueName -ErrorAction SilentlyContinue
    if ($registryValue) {
        $currentValue = if ($valueName) { $registryValue.$valueName } else { $true }

        $result = @{
            Item = $Item
            CurrentValue = $currentValue
            Result = "Not Tested"
            Details = ""
        }

        if ($Operator) {
            switch ($Operator.ToLower()) {
                "equals" {
                    if ($currentValue -ceq $CompareValue) {
                        $result.Result = "Passed"
                        $result.Details = "Current value '$currentValue' equals '$CompareValue'."
                    } else {
                        $result.Result = "Failed"
                        $result.Details = "Current value '$currentValue' does not equal '$CompareValue'."
                    }
                }
                "notequals" {
                    if ($currentValue -cne $CompareValue) {
                        $result.Result = "Passed"
                        $result.Details = "Current value '$currentValue' does not equal '$CompareValue'."
                    } else {
                        $result.Result = "Failed"
                        $result.Details = "Current value '$currentValue' equals '$CompareValue'."
                    }
                }
                "greaterthan" {
                    if ($currentValue -gt $CompareValue) {
                        $result.Result = "Passed"
                        $result.Details = "Current value '$currentValue' is greater than '$CompareValue'."
                    } else {
                        $result.Result = "Failed"
                        $result.Details = "Current value '$currentValue' is not greater than '$CompareValue'."
                    }
                }
                "lessthan" {
                    if ($currentValue -lt $CompareValue) {
                        $result.Result = "Passed"
                        $result.Details = "Current value '$currentValue' is less than '$CompareValue'."
                    } else {
                        $result.Result = "Failed"
                        $result.Details = "Current value '$currentValue' is not less than '$CompareValue'."
                    }
                }
                "contains" {
                    if ($currentValue -like "*$CompareValue*") {
                        $result.Result = "Passed"
                        $result.Details = "Current value '$currentValue' contains '$CompareValue'."
                    } else {
                        $result.Result = "Failed"
                        $result.Details = "Current value '$currentValue' does not contain '$CompareValue'."
                    }
                }
                "notcontains" {
                    if ($currentValue -notlike "*$CompareValue*") {
                        $result.Result = "Passed"
                        $result.Details = "Current value '$currentValue' does not contain '$CompareValue'."
                    } else {
                        $result.Result = "Failed"
                        $result.Details = "Current value '$currentValue' contains '$CompareValue'."
                    }
                }
                default {
                    $result.Result = "Error"
                    $result.Details = "Invalid operator '$Operator'."
                }
            }
        } else {
            $result.Result = "Present" # Operator가 없으면 단순히 존재 여부만 확인
            $result.Details = "Registry item '$Item' is present."
        }
        return $result
    } else {
        return @{
            Item = $Item
            CurrentValue = $null
            Result = "Not Present"
            Details = "Registry item '$Item' does not exist."
        }
    }
}

function Test-FileIntegrity {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,

        [Parameter(Mandatory=$true)]
        [string]$ExpectedHash
    )

    if (-not (Test-Path -Path $FilePath)) {
        return @{
            FilePath = $FilePath
            Result = "Not Present"
            Details = "File not found."
        }
    }

    try {
        $currentHashObject = Get-FileHash -Path $FilePath -Algorithm SHA256
        $currentHash = $currentHashObject.Hash

        if ($currentHash -ceq $ExpectedHash) {
            return @{
                FilePath = $FilePath
                Result = "Passed"
                Details = "File hash matches the expected value."
                ExpectedHash = $ExpectedHash
                CurrentHash = $currentHash
            }
        } else {
            return @{
                FilePath = $FilePath
                Result = "Failed"
                Details = "File hash does not match the expected value."
                ExpectedHash = $ExpectedHash
                CurrentHash = $currentHash
            }
        }
    } catch {
        return @{
            FilePath = $FilePath
            Result = "Error"
            Details = "Error calculating file hash: $($_.Exception.Message)"
            ExpectedHash = $ExpectedHash
        }
    }
}

function Test-SecurityPolicy {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Item,

        [Parameter(Mandatory=$false)]
        [string]$Operator,

        [Parameter(Mandatory=$false)]
        [string]$CompareValue
    )

    $tempFilePath = Join-Path $env:TEMP "security_settings.inf"

    try {
        secedit.exe /export /cfg $tempFilePath /quiet

        if (Test-Path $tempFilePath) {
            $content = Get-Content -Path $tempFilePath

            $foundValue = $content | Where-Object { $_ -like "$Item = *" } | ForEach-Object {
                $_.Split('=')[1].Trim()
            }

            $result = @{
                Item = $Item
                CurrentValue = $foundValue
                Result = "Not Tested"
                Details = ""
            }

            if ($Operator) {
                switch ($Operator.ToLower()) {
                    "equals" {
                        if ([int]$foundValue -ceq [int]$CompareValue) {
                            $result.Result = "Passed"
                            $result.Details = "Current value '$foundValue' equals '$CompareValue'."
                        } else {
                            $result.Result = "Failed"
                            $result.Details = "Current value '$foundValue' does not equal '$CompareValue'."
                        }
                    }
                    "notequals" {
                        if ([int]$foundValue -cne [int]$CompareValue) {
                            $result.Result = "Passed"
                            $result.Details = "Current value '$foundValue' does not equal '$CompareValue'."
                        } else {
                            $result.Result = "Failed"
                            $result.Details = "Current value '$foundValue' equals '$CompareValue'."
                        }
                    }
                    "greaterthan" {
                        if ([int]$foundValue -gt [int]$CompareValue) {
                            $result.Result = "Passed"
                            $result.Details = "Current value '$foundValue' is greater than '$CompareValue'."
                        } else {
                            $result.Result = "Failed"
                            $result.Details = "Current value '$foundValue' is not greater than '$CompareValue'."
                        }
                    }
                    "lessthan" {
                        if ([int]$foundValue -lt [int]$CompareValue) {
                            $result.Result = "Passed"
                            $result.Details = "Current value '$foundValue' is less than '$CompareValue'."
                        } else {
                            $result.Result = "Failed"
                            $result.Details = "Current value '$foundValue' is not less than '$CompareValue'."
                        }
                    }
                    "contains" {
                        if ($foundValue -like "*$CompareValue*") {
                            $result.Result = "Passed"
                            $result.Details = "Current value '$foundValue' contains '$CompareValue'."
                        } else {
                            $result.Result = "Failed"
                            $result.Details = "Current value '$foundValue' does not contain '$CompareValue'."
                        }
                    }
                    "notcontains" {
                        if ($foundValue -notlike "*$CompareValue*") {
                            $result.Result = "Passed"
                            $result.Details = "Current value '$foundValue' does not contain '$CompareValue'."
                        } else {
                            $result.Result = "Failed"
                            $result.Details = "Current value '$foundValue' contains '$CompareValue'."
                        }
                    }
                    default {
                        $result.Result = "Error"
                        $result.Details = "Invalid operator '$Operator'."
                    }
                }
            } else {
                if ($foundValue) {
                    $result.Result = "Present"
                    $result.Details = "Security policy item '$Item' is present."
                } else {
                    $result.Result = "Not Present"
                    $result.Details = "Security policy item '$Item' is not present."
                }
            }
            return $result
        } else {
            return @{
                Item = $Item
                CurrentValue = $null
                Result = "Error"
                Details = "Failed to export security settings."
            }
        }
    } catch {
        return @{
            Item = $Item
            CurrentValue = $null
            Result = "Error"
            Details = "Error during security policy check: $($_.Exception.Message)"
        }
    } finally {
        if (Test-Path $tempFilePath) {
            Remove-Item $tempFilePath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-CurrentAccounts {
    $accountsData = @()
    foreach ($user in Get-LocalUser) {
        $accountInfo = [PSCustomObject]@{
            SamAccountName = $user.Name
            Name = $user.PrincipalSource
            Enabled = $user.Enabled
        }
        $accountsData += $accountInfo
    }
    return $accountsData
}
function Get-CurrentServices {
    $services = Get-Service | Select-Object -Property DisplayName, Status | ForEach-Object {
        $statusText = switch ($_.Status) {
            'Running'         { 'Running' }
            'Stopped'         { 'Stopped' }
            'StartPending'    { 'Start Pending' }
            'StopPending'     { 'Stop Pending' }
            'PausePending'    { 'Pause Pending' }
            'ContinuePending' { 'Continue Pending' }
            'Paused'          { 'Paused' }
            default           { $_.Status } # 알 수 없는 상태는 그대로 표시
        }
        [PSCustomObject]@{
            "DisplayName" = $_.DisplayName
            "Status"      = $statusText
        }
    }
    return $services
}
function Get-ApprovedMedia {
    $approvedMedia = @()
    $usbStorPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\USBSTOR"
    $usbPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\USB"

    # 1. USBSTOR 경로 확인 및 처리
    if (Test-Path $usbStorPath) {
        $usbStorDevices = Get-ChildItem -Path "$usbStorPath\*" -ErrorAction SilentlyContinue | ForEach-Object {
            $friendlyName = Get-ItemProperty -Path $_.PSPath -Name FriendlyName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FriendlyName
            if (-not [string]::IsNullOrEmpty($friendlyName)) {
                [PSCustomObject]@{
                    "미디어" = $friendlyName
                    "Reg"   = $_.PSChildName
                }
            }
        }
        if ($usbStorDevices) {
            $approvedMedia += $usbStorDevices
        }
    } else {
        Write-Warning "레지스트리 경로 '$usbStorPath'를 찾을 수 없습니다."
    }

    # 2. USB 경로 확인 및 처리
    if (Test-Path $usbPath) {
        $usbDevices = Get-ChildItem -Path "$usbPath\*" -ErrorAction SilentlyContinue | ForEach-Object {
            $properties = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
            $mediaName = $properties.FriendlyName -or $properties.DeviceDesc -or $properties.Mfg
            if (-not [string]::IsNullOrEmpty($mediaName)) {
                [PSCustomObject]@{
                    "미디어" = $mediaName
                    "Reg"   = $_.PSChildName
                }
            }
        }
        if ($usbDevices) {
            $approvedMedia += $usbDevices
        }
    } else {
        Write-Warning "레지스트리 경로 '$usbPath'를 찾을 수 없습니다."
    }

    return $approvedMedia | Where-Object {$_.미디어 -ne $null} | Select-Object -Unique
}
function Save-ToJson {
    param(
        [Parameter(Mandatory=$true)]
        [object]$Data,
        [Parameter(Mandatory=$true)]
        [string]$OutputPath,
        [Parameter(Mandatory=$true)]
        [string]$FileName
    )
    $fullPath = Join-Path $OutputPath $FileName
    $Data | ConvertTo-Json -Depth 2 | Out-File -FilePath $fullPath -Encoding UTF8
}
function Get-JsonObjectFromUser {
    $openFileDialogJson = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialogJson.Title = "불러올 JSON 파일 선택"
    $openFileDialogJson.Filter = "JSON 파일 (*.json)|*.json|모든 파일 (*.*)|*.*"
    $openFileDialogJson.Multiselect = $false # 일단 하나의 파일만 선택

    if ($openFileDialogJson.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedJsonFile = $openFileDialogJson.FileName
        try {
            Write-Host "파일 읽는 중: $selectedJsonFile"
            $content = Get-Content -Path $selectedJsonFile -Raw -Encoding UTF8
            $jsonObject = ConvertFrom-Json -InputObject $content
            return $jsonObject
        } catch {
            Write-Error "JSON 파일 처리 오류: $($_.Exception.Message) - 파일: $selectedJsonFile"
            return $null
        }
    }
    return $null # 파일 선택이 취소된 경우
}
