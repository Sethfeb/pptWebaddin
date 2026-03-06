# PPT-EquipSpec-Addin GitHub 초기화 스크립트
# 실행 방법: PowerShell에서 프로젝트 폴더로 이동 후
#   cd "D:\development\PPT-EquipSpec-Addin"
#   .\init_github.ps1

$projectName = "PPT-EquipSpec-Addin"
$githubUser  = "sethfeb"

Write-Host "=== $projectName GitHub 초기화 ===" -ForegroundColor Cyan

# 현재 디렉토리 확인
Write-Host "현재 경로: $(Get-Location)"

# git 초기화
git init
git add .
git commit -m "feat: v1.0.0 초기 릴리즈 - PPT 설비사양 Add-in

- SharePoint REST API 연동 (modSharePoint.bas)
- Collection 기반 로컬 캐시 (modCache.bas)
- 단축어 치환 로직 (modShortCode.bas)
- 검색 패널 UserForm (frmSearch.frm)
- 설정 대화상자 (frmSettings.frm)
- 리본 UI 콜백 (modRibbon.bas, customUI.xml)
- 단위 테스트 (modTests.bas)
- SharePoint List 초기 데이터 템플릿 CSV
- 설치 가이드 (docs/setup_guide.md)"

# GitHub 원격 저장소 연결 (저장소가 이미 생성되어 있어야 함)
# GitHub에서 먼저 빈 저장소 생성: https://github.com/new
# 저장소 이름: PPT-EquipSpec-Addin

$remoteUrl = "https://github.com/$githubUser/$projectName.git"
Write-Host ""
Write-Host "원격 저장소 URL: $remoteUrl" -ForegroundColor Yellow
Write-Host "GitHub에서 먼저 빈 저장소를 생성하세요: https://github.com/new" -ForegroundColor Yellow
Write-Host ""

$confirm = Read-Host "GitHub 저장소 생성 완료 후 Enter를 누르세요 (취소: Ctrl+C)"

git remote add origin $remoteUrl
git branch -M main
git push -u origin main

Write-Host ""
Write-Host "=== 완료 ===" -ForegroundColor Green
Write-Host "저장소 URL: https://github.com/$githubUser/$projectName"
