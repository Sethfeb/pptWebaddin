# PPT 설비사양 Add-in 설치 및 설정 가이드

## 사전 요구사항

- Microsoft 365 데스크톱 (PowerPoint 2016 이상)
- SharePoint Online 또는 SharePoint On-Premises 접근 권한
- Windows 통합 인증(SSO) 환경

---

## 1단계: SharePoint List 생성

### 1.1 SharePoint 사이트 접속

1. SharePoint 사이트 접속 (예: `https://contoso.sharepoint.com/sites/factory`)
2. **설정(톱니바퀴)** → **사이트 콘텐츠** → **새로 만들기** → **목록**

### 1.2 List 생성

- **이름**: `EquipmentSpecs`
- **설명**: 설비 사양 데이터베이스

### 1.3 열(컬럼) 추가

기본 `Title` 열 외에 아래 열을 추가합니다:

| 열 이름 | 유형 | 필수 | 설명 |
|---------|------|------|------|
| `EquipID` | 한 줄 텍스트 | 예 | 설비 식별자 (예: CONV-001) |
| `ShortCode` | 한 줄 텍스트 | 예 | 단축어 (예: !conv001_speed) |
| `SpecName` | 한 줄 텍스트 | 예 | 사양 항목명 (예: 컨베이어 속도) |
| `SpecValue` | 한 줄 텍스트 | 예 | 사양 값 (예: 15) |
| `Unit` | 한 줄 텍스트 | 아니오 | 단위 (예: m/min) |
| `Revision` | 숫자 | 아니오 | 개정 번호 (기본값: 1) |

> **주의**: `Title` 열은 삭제하지 말고 그대로 두세요.

### 1.4 인덱스 설정 (성능 최적화)

- `ShortCode` 열 → 인덱스 추가 (목록 설정 → 인덱스 열)

### 1.5 초기 데이터 입력

`sharepoint/EquipmentSpecs_template.csv` 파일을 참고하여 데이터를 입력합니다.

**빠른 방법**: SharePoint List → **빠른 편집** 모드에서 Excel처럼 붙여넣기 가능

---

## 2단계: Add-in 파일 준비

> **변경된 설치 방식**: `.frm` 파일 직접 임포트는 VBA 편집기에서 지원되지 않습니다.
> `Install.bas` 하나만 임포트하면 모든 모듈과 폼이 자동으로 생성됩니다.

### 2.1 자동 설치 스크립트 실행 (권장)

1. PowerPoint 실행 (`Alt+F11` → VBA 편집기 열기)
2. **파일 → 파일 가져오기** → `src\Install.bas` 선택 (이 파일 하나만 임포트)
3. `Ctrl+G` (직접 실행 창 열기)
4. 아래 명령 입력 후 Enter:
   ```
   Install.RunInstall
   ```
5. 완료 메시지 확인 → 프로젝트 탐색기에 모듈 5개 + 폼 2개 자동 생성됨

> **설치 후**: `Install` 모듈은 우클릭 → 제거해도 됩니다.

### 2.2 리본 XML 적용 (Custom UI Editor 사용)

> **필요 도구**: [Custom UI Editor for Microsoft Office](https://github.com/fernandreu/office-ribbonx-editor/releases)
> (무료, 오픈소스)

1. 현재 파일을 **파일 → 다른 이름으로 저장 → PowerPoint 매크로 사용 프레젠테이션(*.pptm)** 으로 저장
2. PowerPoint **닫기**
3. Custom UI Editor 실행
4. **파일 → 열기** → 저장한 `.pptm` 파일 선택
5. **삽입 → Office 2010+ Custom UI Part**
6. `src\ribbon\customUI.xml` 내용 전체를 붙여넣기
7. **저장** 후 Custom UI Editor 닫기

### 2.3 .ppam으로 저장

> **"서버에 저장할 수 없습니다" 오류 발생 시**: PowerPoint가 OneDrive/SharePoint 동기화 폴더를 기본 저장 위치로 사용하고 있기 때문입니다. 아래 방법으로 **로컬 경로에 직접 저장**하세요.

**저장 방법**:

1. PowerPoint → **파일** → **다른 이름으로 저장** → **이 PC**
2. 주소창에 아래 경로를 직접 입력 후 Enter:
   ```
   %APPDATA%\Microsoft\AddIns
   ```
3. 파일 이름: `EquipSpecAddin`
4. 파일 형식: **PowerPoint 추가 기능 (*.ppam)**
5. **저장** 클릭

> **경로가 없는 경우**: 탐색기에서 `%APPDATA%\Microsoft\AddIns` 폴더를 먼저 생성하세요.

> **여전히 저장 안 될 경우**: 바탕화면 등 임의 로컬 경로에 먼저 저장 후, 탐색기로 `%APPDATA%\Microsoft\AddIns\` 폴더에 복사하세요.

---

## 3단계: Add-in 등록

1. PowerPoint 실행
2. **파일** → **옵션** → **추가 기능**
3. 하단 **관리**: **PowerPoint 추가 기능** → **이동**
4. **추가** → `EquipSpecAddin.ppam` 선택
5. 체크박스 활성화 확인

---

## 4단계: SharePoint URL 설정

1. PowerPoint 리본에서 **설비사양** 탭 확인
2. **SharePoint 설정** 버튼 클릭
3. **SharePoint 사이트 URL** 입력:
   ```
   https://contoso.sharepoint.com/sites/factory
   ```
   (끝에 `/` 없이 입력)
4. **List 이름**: `EquipmentSpecs` (기본값 유지)
5. **연결 테스트** → "연결 성공" 확인
6. **저장**

---

## 5단계: 기능 사용

### 단축어 치환 방법

1. PowerPoint 슬라이드에서 텍스트박스 선택 또는 편집 모드 진입
2. `!conv001_speed` 와 같이 단축어 입력
3. 리본 **설비사양** 탭 → **단축어 치환** 클릭
4. 단축어가 `15 m/min` 으로 자동 치환됨

> **팁**: 한 텍스트박스에 여러 단축어를 입력한 후 한 번에 치환 가능

### 검색 패널 사용 방법

1. 리본 **설비사양** 탭 → **사양 검색** 클릭
2. 검색창에 설비명, 단축어, 사양명 등 키워드 입력
3. **검색** 버튼 또는 Enter
4. 결과 목록에서 항목 선택
5. **슬라이드에 삽입** (값+단위) 또는 **전체 정보 삽입** (사양명: 값 단위)

---

## 전사 배포 방법

### 방법 A: 공유 드라이브 경로

1. `.ppam` 파일을 공유 드라이브에 복사 (예: `\\server\share\AddIns\EquipSpecAddin.ppam`)
2. 각 사용자 PC에서 3단계 등록 시 공유 드라이브 경로 지정

### 방법 B: GPO (그룹 정책)

IT 관리자가 GPO를 통해 레지스트리 키 자동 배포:

```
HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\EquipSpecAddin
  Path = \\server\share\AddIns\EquipSpecAddin.ppam
  AutoLoad = 1
```

---

## 문제 해결

### "SharePoint 연결 오류" 발생 시

1. SharePoint URL 끝에 `/` 가 없는지 확인
2. 회사 VPN 연결 여부 확인
3. 브라우저에서 `{URL}/_api/web/lists` 접속하여 JSON 응답 확인
4. MFA(다단계 인증) 강제 적용 환경: IT 관리자에게 앱 전용 인증 정책 문의

### 단축어가 치환되지 않을 때

1. 단축어 앞에 `!` 가 있는지 확인
2. 단축어가 SharePoint List의 `ShortCode` 열 값과 정확히 일치하는지 확인 (대소문자 무관)
3. **캐시 새로고침** 후 재시도

### 리본 탭이 보이지 않을 때

1. PowerPoint 재시작
2. **파일** → **옵션** → **추가 기능** → **PowerPoint 추가 기능** → Add-in 체크 여부 확인

---

## 버전 이력

| 버전 | 날짜 | 변경 내용 |
|------|------|----------|
| 1.0.0 | 2026-03-05 | 초기 릴리즈: 단축어 치환 + 검색 패널 |
