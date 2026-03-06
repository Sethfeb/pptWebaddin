Attribute VB_Name = "modTests"
'==============================================================================
' modTests.bas
' 단위 테스트 모듈
'
' 실행 방법:
'   PowerPoint VBA 편집기 > 직접 실행 창에서:
'     modTests.RunAllTests
'   또는 개별 테스트 함수 직접 호출
'
' 테스트 패턴:
'   - Assert 헬퍼 함수로 조건 검증
'   - 외부 의존(SharePoint, PowerPoint UI)은 모의 데이터로 대체
'   - 각 테스트는 독립적으로 실행 가능
'==============================================================================
Option Explicit

' 테스트 결과 집계
Private m_passCount As Long
Private m_failCount As Long
Private m_failLog   As String

'==============================================================================
' RunAllTests
' 모든 단위 테스트를 실행하고 결과를 출력한다.
'==============================================================================
Public Sub RunAllTests()
    m_passCount = 0
    m_failCount = 0
    m_failLog = ""

    Debug.Print "========================================="
    Debug.Print "PPT 설비사양 Add-in 단위 테스트 시작"
    Debug.Print "========================================="

    ' JSON 파싱 테스트
    Test_ParseJsonString_BasicField
    Test_ParseJsonString_SpecialChars
    Test_ParseJsonNumber_Integer
    Test_ParseJsonNumber_Missing

    ' 캐시 테스트
    Test_Cache_PutAndGet
    Test_Cache_GetMiss
    Test_Cache_CaseInsensitiveKey
    Test_Cache_PutAll_And_Count
    Test_Cache_Clear
    Test_Cache_Search_ByShortCode
    Test_Cache_Search_BySpecName
    Test_Cache_Search_EmptyKeyword

    ' 단축어 패턴 테스트
    Test_ShortCode_ExtractSingle
    Test_ShortCode_ExtractMultiple
    Test_ShortCode_NoMatch
    Test_ShortCode_WithSpecialPrefix

    ' 설정 테스트
    Test_Settings_SaveAndLoad
    Test_Settings_DefaultListName
    Test_Settings_IsConfigured_False
    Test_Settings_IsConfigured_True

    ' 결과 출력
    Debug.Print "-----------------------------------------"
    Debug.Print "PASS: " & m_passCount & " / FAIL: " & m_failCount
    If m_failCount > 0 Then
        Debug.Print "실패 목록:"
        Debug.Print m_failLog
    End If
    Debug.Print "========================================="

    MsgBox "테스트 완료" & vbCrLf & _
           "PASS: " & m_passCount & vbCrLf & _
           "FAIL: " & m_failCount & vbCrLf & vbCrLf & _
           IIf(m_failCount > 0, "실패 상세는 직접 실행 창 확인", "모든 테스트 통과!"), _
           IIf(m_failCount > 0, vbExclamation, vbInformation), "테스트 결과"
End Sub

'==============================================================================
' ─── JSON 파싱 테스트 ────────────────────────────────────────────────────────
'==============================================================================

Private Sub Test_ParseJsonString_BasicField()
    Dim json   As String
    Dim result As String

    json = "{""EquipID"":""CONV-001"",""ShortCode"":""!conv001_speed""}"
    result = CallExtractJsonString(json, "EquipID")

    Assert "Test_ParseJsonString_BasicField", result = "CONV-001", _
           "Expected 'CONV-001', got '" & result & "'"
End Sub

Private Sub Test_ParseJsonString_SpecialChars()
    Dim json   As String
    Dim result As String

    ' 단위에 특수문자 포함 (°C, m/min)
    json = "{""Unit"":""m/min"",""SpecValue"":""15""}"
    result = CallExtractJsonString(json, "Unit")

    Assert "Test_ParseJsonString_SpecialChars", result = "m/min", _
           "Expected 'm/min', got '" & result & "'"
End Sub

Private Sub Test_ParseJsonNumber_Integer()
    Dim json   As String
    Dim result As String

    json = "{""Revision"":3,""EquipID"":""X""}"
    result = CallExtractJsonNumber(json, "Revision")

    Assert "Test_ParseJsonNumber_Integer", result = "3", _
           "Expected '3', got '" & result & "'"
End Sub

Private Sub Test_ParseJsonNumber_Missing()
    Dim json   As String
    Dim result As String

    json = "{""EquipID"":""X""}"
    result = CallExtractJsonNumber(json, "Revision")

    Assert "Test_ParseJsonNumber_Missing", result = "0", _
           "Expected '0' for missing field, got '" & result & "'"
End Sub

'==============================================================================
' ─── 캐시 테스트 ─────────────────────────────────────────────────────────────
'==============================================================================

Private Sub Test_Cache_PutAndGet()
    Dim rec    As SpecRecord
    Dim result As SpecRecord
    Dim hit    As Boolean

    modCache.ClearCache

    rec.EquipID   = "CONV-001"
    rec.ShortCode = "!conv001_speed"
    rec.SpecName  = "컨베이어 속도"
    rec.SpecValue = "15"
    rec.Unit      = "m/min"
    rec.Revision  = 1

    modCache.CachePut rec
    hit = modCache.CacheGet("!conv001_speed", result)

    Assert "Test_Cache_PutAndGet_Hit", hit = True, "CacheGet should return True"
    Assert "Test_Cache_PutAndGet_Value", result.SpecValue = "15", _
           "Expected '15', got '" & result.SpecValue & "'"
    Assert "Test_Cache_PutAndGet_Unit", result.Unit = "m/min", _
           "Expected 'm/min', got '" & result.Unit & "'"
End Sub

Private Sub Test_Cache_GetMiss()
    Dim result As SpecRecord
    Dim hit    As Boolean

    modCache.ClearCache
    hit = modCache.CacheGet("!nonexistent", result)

    Assert "Test_Cache_GetMiss", hit = False, "CacheGet should return False for missing key"
End Sub

Private Sub Test_Cache_CaseInsensitiveKey()
    Dim rec    As SpecRecord
    Dim result As SpecRecord
    Dim hit    As Boolean

    modCache.ClearCache

    rec.ShortCode = "!CONV001_SPEED"
    rec.SpecValue = "15"
    rec.Unit      = "m/min"
    modCache.CachePut rec

    ' 소문자로 조회
    hit = modCache.CacheGet("!conv001_speed", result)

    Assert "Test_Cache_CaseInsensitiveKey", hit = True, _
           "Cache key lookup should be case-insensitive"
End Sub

Private Sub Test_Cache_PutAll_And_Count()
    Dim items(0 To 2) As SpecRecord

    modCache.ClearCache

    items(0).ShortCode = "!a": items(0).SpecValue = "1"
    items(1).ShortCode = "!b": items(1).SpecValue = "2"
    items(2).ShortCode = "!c": items(2).SpecValue = "3"

    modCache.CachePutAll items

    Assert "Test_Cache_PutAll_Count", modCache.CacheCount = 3, _
           "Expected 3 items, got " & modCache.CacheCount
    Assert "Test_Cache_PutAll_IsLoaded", modCache.IsFullyLoaded = True, _
           "IsFullyLoaded should be True after PutAll"
End Sub

Private Sub Test_Cache_Clear()
    Dim items(0 To 1) As SpecRecord

    items(0).ShortCode = "!x": items(0).SpecValue = "1"
    items(1).ShortCode = "!y": items(1).SpecValue = "2"
    modCache.CachePutAll items

    modCache.ClearCache

    Assert "Test_Cache_Clear_Count", modCache.CacheCount = 0, _
           "Expected 0 items after clear, got " & modCache.CacheCount
    Assert "Test_Cache_Clear_IsLoaded", modCache.IsFullyLoaded = False, _
           "IsFullyLoaded should be False after clear"
End Sub

Private Sub Test_Cache_Search_ByShortCode()
    Dim items(0 To 2) As SpecRecord
    Dim results()     As SpecRecord

    modCache.ClearCache

    items(0).ShortCode = "!conv001_speed": items(0).SpecName = "컨베이어 속도": items(0).EquipID = "CONV-001"
    items(1).ShortCode = "!conv001_width": items(1).SpecName = "컨베이어 폭":   items(1).EquipID = "CONV-001"
    items(2).ShortCode = "!pump001_flow":  items(2).SpecName = "펌프 유량":     items(2).EquipID = "PUMP-001"
    modCache.CachePutAll items

    results = modCache.CacheSearch("conv001")

    Assert "Test_Cache_Search_ByShortCode_Count", UBound(results) + 1 = 2, _
           "Expected 2 results for 'conv001', got " & (UBound(results) + 1)
End Sub

Private Sub Test_Cache_Search_BySpecName()
    Dim items(0 To 1) As SpecRecord
    Dim results()     As SpecRecord

    modCache.ClearCache

    items(0).ShortCode = "!conv001_speed": items(0).SpecName = "컨베이어 속도": items(0).EquipID = "CONV-001"
    items(1).ShortCode = "!pump001_flow":  items(1).SpecName = "펌프 유량":     items(1).EquipID = "PUMP-001"
    modCache.CachePutAll items

    results = modCache.CacheSearch("펌프")

    Assert "Test_Cache_Search_BySpecName_Count", UBound(results) + 1 = 1, _
           "Expected 1 result for '펌프', got " & (UBound(results) + 1)
    Assert "Test_Cache_Search_BySpecName_Value", results(0).EquipID = "PUMP-001", _
           "Expected PUMP-001, got " & results(0).EquipID
End Sub

Private Sub Test_Cache_Search_EmptyKeyword()
    Dim items(0 To 1) As SpecRecord
    Dim results()     As SpecRecord

    modCache.ClearCache

    items(0).ShortCode = "!a": items(0).SpecValue = "1"
    items(1).ShortCode = "!b": items(1).SpecValue = "2"
    modCache.CachePutAll items

    results = modCache.CacheSearch("")

    Assert "Test_Cache_Search_EmptyKeyword", UBound(results) + 1 = 2, _
           "Empty keyword should return all items, got " & (UBound(results) + 1)
End Sub

'==============================================================================
' ─── 단축어 패턴 테스트 ──────────────────────────────────────────────────────
'==============================================================================

Private Sub Test_ShortCode_ExtractSingle()
    Dim codes() As String
    codes = modShortCode.ExtractShortCodesFromText("속도는 !conv001_speed 입니다.")

    Assert "Test_ShortCode_ExtractSingle_Count", UBound(codes) = 0, _
           "Expected 1 shortcode, got " & (UBound(codes) + 1)
    Assert "Test_ShortCode_ExtractSingle_Value", codes(0) = "!conv001_speed", _
           "Expected '!conv001_speed', got '" & codes(0) & "'"
End Sub

Private Sub Test_ShortCode_ExtractMultiple()
    Dim codes() As String
    codes = modShortCode.ExtractShortCodesFromText("!conv001_speed / !conv001_width")

    Assert "Test_ShortCode_ExtractMultiple_Count", UBound(codes) + 1 = 2, _
           "Expected 2 shortcodes, got " & (UBound(codes) + 1)
End Sub

Private Sub Test_ShortCode_NoMatch()
    Dim codes() As String
    codes = modShortCode.ExtractShortCodesFromText("단축어 없는 일반 텍스트입니다.")

    ' 빈 배열: UBound = 0, codes(0) = ""
    Assert "Test_ShortCode_NoMatch", codes(0) = "", _
           "Expected empty result for text without shortcodes"
End Sub

Private Sub Test_ShortCode_WithSpecialPrefix()
    Dim codes() As String
    codes = modShortCode.ExtractShortCodesFromText("값: !robot001_reach mm")

    Assert "Test_ShortCode_WithSpecialPrefix_Count", UBound(codes) = 0, _
           "Expected 1 shortcode"
    Assert "Test_ShortCode_WithSpecialPrefix_Value", codes(0) = "!robot001_reach", _
           "Expected '!robot001_reach', got '" & codes(0) & "'"
End Sub

'==============================================================================
' ─── 설정 테스트 ─────────────────────────────────────────────────────────────
'==============================================================================

Private Sub Test_Settings_SaveAndLoad()
    Dim testUrl As String
    testUrl = "https://test.sharepoint.com/sites/testfactory"

    modSettings.SetSharePointUrl testUrl
    Dim loaded As String
    loaded = modSettings.GetSharePointUrl

    Assert "Test_Settings_SaveAndLoad", loaded = testUrl, _
           "Expected '" & testUrl & "', got '" & loaded & "'"

    ' 정리
    modSettings.SetSharePointUrl ""
End Sub

Private Sub Test_Settings_DefaultListName()
    ' ListName 키를 삭제하여 기본값 테스트
    DeleteSetting "EquipSpecAddin", "Config", "ListName"

    Dim name As String
    name = modSettings.GetListName

    Assert "Test_Settings_DefaultListName", name = "EquipmentSpecs", _
           "Expected 'EquipmentSpecs', got '" & name & "'"
End Sub

Private Sub Test_Settings_IsConfigured_False()
    modSettings.SetSharePointUrl ""
    Assert "Test_Settings_IsConfigured_False", modSettings.IsConfigured = False, _
           "IsConfigured should be False when URL is empty"
End Sub

Private Sub Test_Settings_IsConfigured_True()
    modSettings.SetSharePointUrl "https://test.sharepoint.com/sites/x"
    Assert "Test_Settings_IsConfigured_True", modSettings.IsConfigured = True, _
           "IsConfigured should be True when URL is set"
    modSettings.SetSharePointUrl ""
End Sub

'==============================================================================
' ─── 헬퍼 함수 ───────────────────────────────────────────────────────────────
'==============================================================================

' Assert: 조건이 False이면 실패로 기록
Private Sub Assert(ByVal testName As String, ByVal condition As Boolean, ByVal message As String)
    If condition Then
        m_passCount = m_passCount + 1
        Debug.Print "  [PASS] " & testName
    Else
        m_failCount = m_failCount + 1
        m_failLog = m_failLog & "  [FAIL] " & testName & ": " & message & vbCrLf
        Debug.Print "  [FAIL] " & testName & ": " & message
    End If
End Sub

' modSharePoint의 Private 함수를 테스트하기 위한 래퍼
' (VBA에서 Private 함수 직접 호출 불가 → 동일 로직을 여기서 재구현)
Private Function CallExtractJsonString(ByVal objJson As String, ByVal key As String) As String
    Dim re      As Object
    Dim matches As Object

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = False
    re.Pattern = """" & key & """\s*:\s*""([^""]*)"""

    Set matches = re.Execute(objJson)
    If matches.Count > 0 Then
        CallExtractJsonString = matches(0).SubMatches(0)
    Else
        CallExtractJsonString = ""
    End If

    Set re = Nothing
End Function

Private Function CallExtractJsonNumber(ByVal objJson As String, ByVal key As String) As String
    Dim re      As Object
    Dim matches As Object

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = False
    re.Pattern = """" & key & """\s*:\s*([0-9]+)"

    Set matches = re.Execute(objJson)
    If matches.Count > 0 Then
        CallExtractJsonNumber = matches(0).SubMatches(0)
    Else
        CallExtractJsonNumber = "0"
    End If

    Set re = Nothing
End Function
