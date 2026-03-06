Attribute VB_Name = "modCache"
'==============================================================================
' modCache.bas
' 설비사양 로컬 캐시 모듈 (Collection 기반, 세션 내 유지)
'
' 설계:
'   - 모듈 수준 Collection을 싱글턴으로 사용
'   - 키: ShortCode (소문자 정규화)
'   - 값: SpecRecord (UDT → Variant 직렬화 불가 → 배열로 래핑하여 저장)
'   - 스레드 안전성: VBA 단일 스레드 모델이므로 별도 잠금 불필요
'   - 수명: Add-in 세션 동안 유지. ClearCache 또는 RefreshCache로 초기화.
'==============================================================================
Option Explicit

' ─── 모듈 수준 상태 ──────────────────────────────────────────────────────────
Private m_cache       As Collection   ' Key=ShortCode(소문자), Value=Variant()(6개 필드)
Private m_isLoaded    As Boolean      ' 전체 목록 프리페치 완료 여부
Private m_loadedAt    As Date         ' 마지막 프리페치 시각

' 캐시 만료 시간 (분 단위). 0 = 만료 없음.
Private Const CACHE_TTL_MINUTES As Long = 60

'==============================================================================
' EnsureCache  [Private]
' Collection이 초기화되지 않은 경우 생성한다.
'==============================================================================
Private Sub EnsureCache()
    If m_cache Is Nothing Then
        Set m_cache = New Collection
    End If
End Sub

'==============================================================================
' CacheGet
' 캐시에서 ShortCode로 SpecRecord를 조회한다.
'
' 반환값: True=히트, False=미스
'==============================================================================
Public Function CacheGet( _
    ByVal shortCode As String, _
    ByRef result As SpecRecord) As Boolean

    Dim key  As String
    Dim arr  As Variant

    CacheGet = False
    EnsureCache

    key = LCase(Trim(shortCode))

    On Error Resume Next
    arr = m_cache(key)
    On Error GoTo 0

    If IsEmpty(arr) Then Exit Function
    If Not IsArray(arr) Then Exit Function

    result.EquipID   = arr(0)
    result.ShortCode = arr(1)
    result.SpecName  = arr(2)
    result.SpecValue = arr(3)
    result.Unit      = arr(4)
    result.Revision  = CLng(arr(5))

    CacheGet = True
End Function

'==============================================================================
' CachePut
' SpecRecord를 캐시에 저장한다. 이미 존재하면 덮어쓴다.
'==============================================================================
Public Sub CachePut(ByRef rec As SpecRecord)
    Dim key As String
    Dim arr(0 To 5) As Variant

    EnsureCache

    key    = LCase(Trim(rec.ShortCode))
    arr(0) = rec.EquipID
    arr(1) = rec.ShortCode
    arr(2) = rec.SpecName
    arr(3) = rec.SpecValue
    arr(4) = rec.Unit
    arr(5) = rec.Revision

    ' 기존 키 제거 후 재삽입 (Collection은 Update 미지원)
    On Error Resume Next
    m_cache.Remove key
    On Error GoTo 0

    m_cache.Add arr, key
End Sub

'==============================================================================
' CachePutAll
' SpecRecord 배열 전체를 캐시에 저장한다. (프리페치용)
'==============================================================================
Public Sub CachePutAll(ByRef items() As SpecRecord)
    Dim i As Long

    If UBound(items) < 0 Then Exit Sub

    For i = 0 To UBound(items)
        CachePut items(i)
    Next i

    m_isLoaded = True
    m_loadedAt = Now
End Sub

'==============================================================================
' CacheGetAll
' 캐시에 저장된 모든 항목을 SpecRecord 배열로 반환한다.
' 검색 패널의 로컬 필터링에 사용한다.
'==============================================================================
Public Function CacheGetAll() As SpecRecord()
    Dim result()  As SpecRecord
    Dim empty()   As SpecRecord
    Dim i         As Long
    Dim arr       As Variant

    ReDim empty(-1 To -1)
    EnsureCache

    If m_cache.Count = 0 Then
        CacheGetAll = empty
        Exit Function
    End If

    ReDim result(0 To m_cache.Count - 1)
    i = 0

    Dim item As Variant
    For Each item In m_cache
        If IsArray(item) Then
            result(i).EquipID   = item(0)
            result(i).ShortCode = item(1)
            result(i).SpecName  = item(2)
            result(i).SpecValue = item(3)
            result(i).Unit      = item(4)
            result(i).Revision  = CLng(item(5))
            i = i + 1
        End If
    Next item

    If i = 0 Then
        CacheGetAll = empty
    Else
        ReDim Preserve result(0 To i - 1)
        CacheGetAll = result
    End If
End Function

'==============================================================================
' CacheSearch
' 캐시 내에서 keyword로 ShortCode 또는 SpecName을 부분 검색한다.
' 캐시가 프리페치된 경우 SharePoint 호출 없이 로컬 필터링 가능.
'==============================================================================
Public Function CacheSearch(ByVal keyword As String) As SpecRecord()
    Dim all()     As SpecRecord
    Dim result()  As SpecRecord
    Dim empty()   As SpecRecord
    Dim i         As Long
    Dim count     As Long
    Dim kw        As String

    ReDim empty(-1 To -1)
    kw = LCase(Trim(keyword))

    If Len(kw) = 0 Then
        CacheSearch = CacheGetAll
        Exit Function
    End If

    all = CacheGetAll
    If UBound(all) < 0 Then
        CacheSearch = empty
        Exit Function
    End If

    ReDim result(0 To UBound(all))
    count = 0

    For i = 0 To UBound(all)
        If InStr(1, LCase(all(i).ShortCode), kw) > 0 Or _
           InStr(1, LCase(all(i).SpecName), kw) > 0 Or _
           InStr(1, LCase(all(i).EquipID), kw) > 0 Then
            result(count) = all(i)
            count = count + 1
        End If
    Next i

    If count = 0 Then
        CacheSearch = empty
    Else
        ReDim Preserve result(0 To count - 1)
        CacheSearch = result
    End If
End Function

'==============================================================================
' SearchVariant
' UserForm에서 SpecRecord 타입 없이 사용할 수 있도록 Variant 배열로 반환.
' 각 원소: Array(EquipID, ShortCode, SpecName, SpecValue, Unit, Revision)
' siteUrl이 있으면 캐시 미스 시 SharePoint에서 직접 조회.
'==============================================================================
Public Function SearchVariant(ByVal keyword As String, ByVal siteUrl As String) As Variant()
    Dim recs()    As SpecRecord
    Dim result()  As Variant
    Dim empty()   As Variant
    Dim i         As Long

    ReDim empty(-1 To -1)
    SearchVariant = empty

    If IsFullyLoaded Then
        recs = CacheSearch(keyword)
    ElseIf Len(siteUrl) > 0 Then
        If Len(keyword) = 0 Then
            recs = modSharePoint.GetAllSpecs(siteUrl)
        Else
            recs = modSharePoint.SearchSpecs(siteUrl, keyword)
        End If
    Else
        Exit Function
    End If

    If UBound(recs) < 0 Then Exit Function

    ReDim result(0 To UBound(recs))
    For i = 0 To UBound(recs)
        result(i) = Array(recs(i).EquipID, recs(i).ShortCode, recs(i).SpecName, _
                          recs(i).SpecValue, recs(i).Unit, recs(i).Revision)
    Next i

    SearchVariant = result
End Function

'==============================================================================
' ClearCache
' 캐시를 완전히 초기화한다.
'==============================================================================
Public Sub ClearCache()
    Set m_cache = New Collection
    m_isLoaded = False
    m_loadedAt = CDate(0)
End Sub

'==============================================================================
' IsFullyLoaded
' 전체 목록 프리페치 완료 여부를 반환한다.
' TTL이 설정된 경우 만료 여부도 확인한다.
'==============================================================================
Public Function IsFullyLoaded() As Boolean
    If Not m_isLoaded Then
        IsFullyLoaded = False
        Exit Function
    End If

    If CACHE_TTL_MINUTES > 0 Then
        Dim elapsedMin As Long
        elapsedMin = DateDiff("n", m_loadedAt, Now)
        If elapsedMin >= CACHE_TTL_MINUTES Then
            m_isLoaded = False
            IsFullyLoaded = False
            Exit Function
        End If
    End If

    IsFullyLoaded = True
End Function

'==============================================================================
' CacheCount
' 현재 캐시에 저장된 항목 수를 반환한다.
'==============================================================================
Public Function CacheCount() As Long
    EnsureCache
    CacheCount = m_cache.Count
End Function
