Attribute VB_Name = "modSharePoint"
'==============================================================================
' modSharePoint.bas
' SharePoint REST API v1 호출 및 JSON 파싱 모듈
'
' 의존: WinHttp.WinHttpRequest.5.1 (Windows 내장 COM)
'       VBScript.RegExp (Windows 내장 COM)
'
' 스레드 안전성: 단일 스레드(VBA 기본). 동시 호출 없음.
' 메모리 관리: WinHttpRequest는 함수 종료 시 Set obj = Nothing 으로 해제.
'==============================================================================
Option Explicit

' ─── 공개 상수 ───────────────────────────────────────────────────────────────
Public Const SP_LIST_NAME As String = "EquipmentSpecs"
Public Const SP_REQUEST_TIMEOUT_MS As Long = 15000   ' 15초

' ─── 내부 상수 ───────────────────────────────────────────────────────────────
Private Const HTTP_OK As Integer = 200
Private Const SP_SELECT_FIELDS As String = _
    "EquipID,ShortCode,SpecName,SpecValue,Unit,Revision"

'==============================================================================
' GetSpecByShortCode
' 단축어 1건으로 사양 레코드를 조회한다.
'
' 매개변수:
'   siteUrl   - SharePoint 사이트 루트 URL (예: https://contoso.sharepoint.com/sites/factory)
'   shortCode - 조회할 단축어 (예: !conv001_speed)
'   result    - [출력] 조회된 SpecRecord
'
' 반환값: True=성공, False=실패(result는 빈 값)
'==============================================================================
Public Function GetSpecByShortCode( _
    ByVal siteUrl As String, _
    ByVal shortCode As String, _
    ByRef result As SpecRecord) As Boolean

    Dim endpoint As String
    Dim encoded  As String
    Dim json     As String
    Dim items()  As SpecRecord

    GetSpecByShortCode = False

    ' 단축어의 작은따옴표를 이스케이프 (OData 규칙: '' 로 치환)
    encoded = Replace(shortCode, "'", "''")

    endpoint = siteUrl & "/_api/web/lists/getbytitle('" & SP_LIST_NAME & "')" & _
               "/items?$select=" & SP_SELECT_FIELDS & _
               "&$filter=ShortCode eq '" & encoded & "'" & _
               "&$top=1"

    json = ExecuteGetRequest(endpoint)
    If Len(json) = 0 Then Exit Function

    items = ParseSpecItems(json)
    If UBound(items) < 0 Then Exit Function

    result = items(0)
    GetSpecByShortCode = True
End Function

'==============================================================================
' GetAllSpecs
' 전체 사양 목록을 배열로 반환한다. (캐시 프리페치용)
'
' 매개변수:
'   siteUrl - SharePoint 사이트 루트 URL
'
' 반환값: SpecRecord 배열 (0-based). 실패 시 빈 배열(UBound = -1).
'==============================================================================
Public Function GetAllSpecs(ByVal siteUrl As String) As SpecRecord()
    Dim endpoint As String
    Dim json     As String
    Dim empty()  As SpecRecord

    ReDim empty(-1 To -1)   ' 빈 배열 초기화

    endpoint = siteUrl & "/_api/web/lists/getbytitle('" & SP_LIST_NAME & "')" & _
               "/items?$select=" & SP_SELECT_FIELDS & _
               "&$orderby=EquipID,ShortCode" & _
               "&$top=5000"

    json = ExecuteGetRequest(endpoint)
    If Len(json) = 0 Then
        GetAllSpecs = empty
        Exit Function
    End If

    GetAllSpecs = ParseSpecItems(json)
End Function

'==============================================================================
' SearchSpecs
' 키워드로 ShortCode 또는 SpecName을 부분 검색한다.
'
' 매개변수:
'   siteUrl - SharePoint 사이트 루트 URL
'   keyword - 검색어 (부분 일치)
'
' 반환값: SpecRecord 배열 (0-based). 실패 시 빈 배열.
'==============================================================================
Public Function SearchSpecs( _
    ByVal siteUrl As String, _
    ByVal keyword As String) As SpecRecord()

    Dim endpoint As String
    Dim json     As String
    Dim encoded  As String
    Dim empty()  As SpecRecord

    ReDim empty(-1 To -1)

    encoded = Replace(keyword, "'", "''")

    ' OData v3 substringof 함수 사용 (SharePoint REST API v1 지원)
    endpoint = siteUrl & "/_api/web/lists/getbytitle('" & SP_LIST_NAME & "')" & _
               "/items?$select=" & SP_SELECT_FIELDS & _
               "&$filter=substringof('" & encoded & "',ShortCode) or " & _
               "substringof('" & encoded & "',SpecName)" & _
               "&$top=100"

    json = ExecuteGetRequest(endpoint)
    If Len(json) = 0 Then
        SearchSpecs = empty
        Exit Function
    End If

    SearchSpecs = ParseSpecItems(json)
End Function

'==============================================================================
' ExecuteGetRequest  [Private]
' WinHttpRequest로 GET 요청을 실행하고 응답 본문을 반환한다.
'
' 인증: SetAutoLogonPolicy(0) → Windows 통합 인증(NTLM/Kerberos) 자동 적용
' 오류: 실패 시 빈 문자열 반환 + MsgBox 알림
'==============================================================================
Private Function ExecuteGetRequest(ByVal url As String) As String
    Dim http As Object
    Dim resp As String

    ExecuteGetRequest = ""

    On Error GoTo ErrHandler
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "GET", url, False
    http.SetAutoLogonPolicy 0          ' AUTOLOGON_POLICY_ALWAYS: Windows 자격증명 자동 전송
    http.SetRequestHeader "Accept", "application/json;odata=verbose"
    http.SetRequestHeader "Content-Type", "application/json;odata=verbose"
    http.SetTimeouts SP_REQUEST_TIMEOUT_MS, SP_REQUEST_TIMEOUT_MS, _
                     SP_REQUEST_TIMEOUT_MS, SP_REQUEST_TIMEOUT_MS

    http.Send

    If http.Status = HTTP_OK Then
        resp = http.ResponseText
        ExecuteGetRequest = resp
    Else
        Dim m1 As String
        m1 = "SharePoint request failed." & vbCrLf
        m1 = m1 & "HTTP " & http.Status & ": " & http.StatusText & vbCrLf
        m1 = m1 & "URL: " & url
        MsgBox m1, vbExclamation, "EquipSpec Add-in"
    End If

    Set http = Nothing
    Exit Function

ErrHandler:
    If Not http Is Nothing Then Set http = Nothing
    Dim m2 As String
    m2 = "SharePoint connection error: " & Err.Description & vbCrLf & "URL: " & url
    MsgBox m2, vbCritical, "EquipSpec Add-in"
End Function

'==============================================================================
' ParseSpecItems  [Private]
' SharePoint REST API JSON 응답(odata=verbose)에서 SpecRecord 배열을 추출한다.
'
' JSON 구조 (odata=verbose):
'   {"d":{"results":[{"EquipID":"...","ShortCode":"...","SpecName":"...",
'                     "SpecValue":"...","Unit":"...","Revision":1}, ...]}}
'
' 파싱 전략: VBScript.RegExp으로 "results":[...] 블록을 추출한 뒤
'            각 객체 {}를 순회하며 필드값을 추출한다.
'            중첩 깊이 1단계만 처리하므로 현재 스키마에 충분함.
'==============================================================================
Private Function ParseSpecItems(ByVal json As String) As SpecRecord()
    Dim re       As Object
    Dim matches  As Object
    Dim items()  As SpecRecord
    Dim empty()  As SpecRecord
    Dim i        As Long
    Dim objJson  As String

    ReDim empty(-1 To -1)
    ParseSpecItems = empty

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.MultiLine = True
    re.IgnoreCase = False

    ' results 배열 전체 추출
    re.Pattern = """results""\s*:\s*\[(.+)\]\s*\}\s*\}"
    Set matches = re.Execute(json)
    If matches.Count = 0 Then
        Set re = Nothing
        Exit Function
    End If

    Dim resultsBlock As String
    resultsBlock = matches(0).SubMatches(0)

    ' 개별 객체 {} 추출 (최상위 레벨만, 중첩 없음)
    Dim objPattern As String
    objPattern = "\{[^{}]+\}"
    re.Pattern = objPattern
    Set matches = re.Execute(resultsBlock)

    If matches.Count = 0 Then
        Set re = Nothing
        Exit Function
    End If

    ReDim items(0 To matches.Count - 1)

    For i = 0 To matches.Count - 1
        objJson = matches(i).Value
        items(i).EquipID   = ExtractJsonString(objJson, "EquipID")
        items(i).ShortCode = ExtractJsonString(objJson, "ShortCode")
        items(i).SpecName  = ExtractJsonString(objJson, "SpecName")
        items(i).SpecValue = ExtractJsonString(objJson, "SpecValue")
        items(i).Unit      = ExtractJsonString(objJson, "Unit")
        items(i).Revision  = CLng(Val(ExtractJsonNumber(objJson, "Revision")))
    Next i

    Set re = Nothing
    ParseSpecItems = items
End Function

'==============================================================================
' ExtractJsonString  [Private]
' JSON 객체 문자열에서 지정 키의 문자열 값을 추출한다.
' 예: {"EquipID":"CONV-001"} → "CONV-001"
'==============================================================================
Private Function ExtractJsonString(ByVal objJson As String, ByVal key As String) As String
    Dim re      As Object
    Dim matches As Object

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = False
    re.Pattern = """" & key & """\s*:\s*""([^""]*)"""

    Set matches = re.Execute(objJson)
    If matches.Count > 0 Then
        ExtractJsonString = matches(0).SubMatches(0)
    Else
        ExtractJsonString = ""
    End If

    Set re = Nothing
End Function

'==============================================================================
' ExtractJsonNumber  [Private]
' JSON 객체 문자열에서 지정 키의 숫자 값을 문자열로 추출한다.
' 예: {"Revision":2} → "2"
'==============================================================================
Private Function ExtractJsonNumber(ByVal objJson As String, ByVal key As String) As String
    Dim re      As Object
    Dim matches As Object

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = False
    re.Pattern = """" & key & """\s*:\s*([0-9]+)"

    Set matches = re.Execute(objJson)
    If matches.Count > 0 Then
        ExtractJsonNumber = matches(0).SubMatches(0)
    Else
        ExtractJsonNumber = "0"
    End If

    Set re = Nothing
End Function
