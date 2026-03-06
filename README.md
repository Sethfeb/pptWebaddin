# PPT-EquipSpec-Addin

PowerPoint 설비사양 자동 입력 Add-in (.ppam)

## 개요

SharePoint List(`EquipmentSpecs`)를 설비 사양 DB로 활용하여,
PowerPoint 슬라이드 작성 시 단축어 치환 및 검색 패널을 통해
설비 사양을 일관성 있게 입력할 수 있도록 지원하는 VBA Add-in입니다.

## 주요 기능

1. **단축어 치환**: 텍스트박스에 `!단축어` 입력 후 `Ctrl+Shift+E` → 해당 사양값으로 자동 치환
2. **검색 패널**: 리본 탭 > "설비사양 검색" 버튼 → 키워드 검색 후 선택 삽입

## 파일 구조

```
PPT-EquipSpec-Addin/
├── src/
│   ├── modules/
│   │   ├── modSharePoint.bas   # SharePoint REST API 호출 + JSON 파싱
│   │   ├── modShortCode.bas    # 단축어 치환 로직
│   │   ├── modRibbon.bas       # 리본 콜백 함수
│   │   └── modCache.bas        # 로컬 캐시 (Collection 기반)
│   ├── forms/
│   │   └── frmSearch.frm       # 검색 패널 UserForm
│   └── ribbon/
│       └── customUI.xml        # 리본 탭/버튼 정의
├── sharepoint/
│   └── EquipmentSpecs_template.csv  # SharePoint List 초기 데이터 템플릿
├── tests/
│   └── modTests.bas            # 단위 테스트
├── docs/
│   └── setup_guide.md          # 설치 및 SharePoint 설정 가이드
└── README.md
```

## 프로젝트 경로

```
D:\development\PPT-EquipSpec-Addin\
```

## 빠른 시작

1. `docs/setup_guide.md` 참조하여 SharePoint List 생성
2. `src/ribbon/customUI.xml` + 모든 `.bas`/`.frm` 파일을 PowerPoint VBA 편집기에 임포트
3. `.ppam`으로 저장 후 `%APPDATA%\Microsoft\AddIns\`에 복사
4. PowerPoint → 파일 → 옵션 → 추가 기능 → PowerPoint 추가 기능 → 찾아보기 → 등록

## 버전

- v1.0.0 - 초기 릴리즈
- 대상 환경: Microsoft 365 데스크톱, SharePoint Online (Windows SSO)
- VBA 호환: Office 2016 이상
