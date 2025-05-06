WIN_INSPECTION_SCRIPT 해설서 (인수인계용)

본 문서는 WIN_INSPECTION_SCRIPT 프로젝트의 코드 구조, 기능, 설정 방법 및 유지보수 정보를 담고 있습니다. 이 문서를 통해 스크립트를 이해하고, 필요한 경우 수정하거나 확장하는 데 도움이 되기를 바랍니다.

1. 프로젝트 개요
WIN_INSPECTION_SCRIPT는 Windows 시스템의 정보를 수집하고 감사하는 PowerShell 기반 스크립트 모음입니다. GUI 인터페이스를 통해 사용자는 특정 워크스테이션을 선택하고, 기본 정보 확인, 감사 기록 열람, 로그 백업 등의 작업을 수행할 수 있습니다.

2. 파일 구조
프로젝트의 주요 파일 및 폴더 구조는 다음과 같습니다.
WIN_INSPECTION_SCRIPT/
├── script/
│   ├── MainPanel.ps1         # 메인 GUI 인터페이스 및 로직
│   ├── InspectionModule.ps1  # (현재 사용되지 않음) 시스템 검사 관련 모듈 (향후 확장 가능)
│   └── backupScript.ps1      # 이벤트 로그 백업 GUI 및 로직
├── Report/                 # 워크스테이션별 보고서 및 정보 저장 폴더
│   └── [워크스테이션 이름]/
│       ├── [YYYYMMDD_HHmmss]/ # 감사 기록 폴더 (날짜 및 시간별)
│       │   ├── account.json      # 계정 정보 (예시)
│       │   ├── service.json      # 서비스 정보 (예시)
│       │   └── security_report.log # 보안 로그 (예시)
│       └── basic_info.txt      # 워크스테이션별 기본 정보 (노트 형식)
├── etc/                    # 기타 설정 파일 또는 리소스 (현재 background 이미지 폴더)
│   └── background/
│       └── [배경 이미지 파일]
└── readme.txt              # 본 해설서
3. 주요 스크립트 설명


3.1. script/MainPanel.ps1
역할: 스크립트의 메인 GUI 인터페이스를 제공하며, 사용자 상호작용 및 주요 로직을 처리합니다.
주요 기능:
워크스테이션 선택 ($computerComboBox): 콤보 박스를 통해 검사할 워크스테이션을 선택합니다.
Basic Info ($basicInfoTextArea, $basicInfoChangeButton): 선택된 워크스테이션에 대한 기본 정보를 기록하고 저장/불러옵니다. 텍스트 영역에 노트를 작성하고 "저장" 버튼을 클릭하면 Report\[워크스테이션 이름]\basic_info.txt 파일에 저장됩니다. 워크스테이션 선택 시 해당 파일이 존재하면 내용을 불러와 표시합니다.

Audit History ($auditHistoryListView): 선택된 워크스테이션의 감사 기록 폴더 (Report\[워크스테이션 이름]\[YYYYMMDD_HHmmss]) 목록을 표시합니다. 각 감사 폴더의 이름 (YYYYMMDD_HHmmss), Compliance 정보, 로그인 사용자 정보를 ListView에 보여줍니다.
보고서 보기 (우클릭 메뉴): Audit History ListView의 아이템을 오른쪽 클릭하고 "보고서 보기"를 선택하면 해당 감사 폴더의 정보를 기반으로 HTML 보고서 (audit_report.html)를 생성하여 웹 브라우저로 열람합니다. 보고서에는 계정 정보 (account.json), 서비스 정보 (service.json), 그리고 security_report.log 파일 내용이 포함됩니다.

Log Backup ($logBackupButton): script\backupScript.ps1 스크립트를 실행하여 이벤트 로그 백업 GUI를 표시합니다. 선택된 워크스테이션 이름을 파라미터로 전달하여 백업 로그가 log\[워크스테이션 이름]\[YYYYMMDD_HHmmss] 폴더에 저장되도록 합니다.
이미지 배경: etc 폴더의 이미지를 GUI 배경으로 설정합니다.

3.2. script/backupScript.ps1
역할: 이벤트 로그 백업을 위한 GUI 인터페이스 및 백업 로직을 제공합니다.
주요 기능:
백업할 이벤트 로그 선택 ($logCheckBoxList): CheckBoxList를 통해 백업할 이벤트 로그를 선택합니다.
주요 로그/전체 로그 선택 버튼: 자주 사용되는 주요 로그 또는 전체 로그를 빠르게 선택할 수 있는 버튼을 제공합니다.
백업 시작 ($startButton): 선택된 로그를 백업합니다.
백업 폴더 구조: 백업된 로그는 log 폴더 아래에 선택된 워크스테이션 이름으로 된 폴더를 생성하고, 그 하위에 타임스탬프 (YYYYMMDD_HHmmss) 형식의 폴더를 생성하여 저장됩니다.

3.3. script/InspectionModule.ps1
역할: (현재 사용되지 않음) 시스템 검사 관련 함수 및 로직을 포함할 예정이었던 모듈입니다. 향후 시스템 정보 수집, 설정 확인 등의 기능을 확장할 때 활용할 수 있습니다.

4. 설정 및 실행 방법

파일 구조 확인: 위에 설명된 파일 및 폴더 구조가 올바르게 구성되어 있는지 확인합니다.
배경 이미지: GUI 배경 이미지를 변경하려면 etc\background 폴더에 원하는 이미지 파일을 넣고 MainPanel.ps1 스크립트 내에서 해당 파일명을 수정합니다.
스크립트 실행: MainPanel.ps1 파일을 PowerShell에서 실행하면 GUI 인터페이스가 나타납니다.

5. 유지보수 및 확장
코드 스타일: PowerShell 코딩 규칙을 준수하여 일관성을 유지하는 것이 좋습니다.
오류 처리: try-catch 블록을 적극적으로 사용하여 예외 상황에 대비하고 사용자에게 적절한 오류 메시지를 제공해야 합니다.
로깅: 필요한 경우 스크립트 실행 과정, 오류 등을 로그 파일에 기록하여 추후 문제 해결에 도움을 받을 수 있습니다.
모듈화: 기능별로 함수를 분리하고, 필요하다면 모듈 (.psm1)로 만들어 관리하는 것이 좋습니다.
확장 계획:
InspectionModule.ps1을 활용하여 다양한 시스템 정보 수집 및 검사 기능을 추가할 수 있습니다.
보고서 형식을 다양화 (CSV, JSON 등)하거나, 보고서에 더 많은 정보를 포함할 수 있습니다.
원격 워크스테이션 검사 기능을 추가할 수 있습니다.
자동화된 감사 스케줄링 기능을 추가할 수 있습니다.

6. 주의 사항

스크립트를 실행하는 환경에 필요한 권한이 있는지 확인해야 합니다. 특히 시스템 정보 수집 및 로그 백업 관련 기능은 관리자 권한이 필요할 수 있습니다.
스크립트 수정 시에는 충분한 테스트를 거쳐 예상치 못한 동작이나 오류가 발생하지 않도록 주의해야 합니다.