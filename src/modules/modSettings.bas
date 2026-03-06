Attribute VB_Name = "modSettings"
'==============================================================================
' modSettings.bas
' Add-in 설정 저장/불러오기 모듈
'
' 저장 위치: Windows 레지스트리 HKCU\Software\EquipSpecAddin
'   - SharePointUrl : SharePoint 사이트 루트 URL
'   - ListName      : 사양 List 이름 (기본값: EquipmentSpecs)
'   - PrefetchOnLoad: 시작 시 전체 목록 프리페치 여부 (0/1)
'
' 레지스트리 사용 근거: VBA GetSetting/SaveSetting은 HKCU\Software\VB and VBA Program Settings
'   경로에 저장되며 별도 COM 없이 사용 가능. (VBA 내장 함수)
'==============================================================================
Option Explicit

Private Const APP_NAME    As String = "EquipSpecAddin"
Private Const SECTION_CFG As String = "Config"

'==============================================================================
' GetSharePointUrl
' 저장된 SharePoint URL을 반환한다. 미설정 시 빈 문자열.
'==============================================================================
Public Function GetSharePointUrl() As String
    GetSharePointUrl = GetSetting(APP_NAME, SECTION_CFG, "SharePointUrl", "")
End Function

'==============================================================================
' SetSharePointUrl
' SharePoint URL을 저장한다.
'==============================================================================
Public Sub SetSharePointUrl(ByVal url As String)
    SaveSetting APP_NAME, SECTION_CFG, "SharePointUrl", Trim(url)
End Sub

'==============================================================================
' GetListName
' 사양 List 이름을 반환한다. 기본값: EquipmentSpecs
'==============================================================================
Public Function GetListName() As String
    GetListName = GetSetting(APP_NAME, SECTION_CFG, "ListName", "EquipmentSpecs")
End Function

'==============================================================================
' SetListName
'==============================================================================
Public Sub SetListName(ByVal listName As String)
    SaveSetting APP_NAME, SECTION_CFG, "ListName", Trim(listName)
End Sub

'==============================================================================
' GetPrefetchOnLoad
' 시작 시 전체 목록 프리페치 여부. 기본값: True
'==============================================================================
Public Function GetPrefetchOnLoad() As Boolean
    Dim val As String
    val = GetSetting(APP_NAME, SECTION_CFG, "PrefetchOnLoad", "1")
    GetPrefetchOnLoad = (val = "1")
End Function

'==============================================================================
' SetPrefetchOnLoad
'==============================================================================
Public Sub SetPrefetchOnLoad(ByVal enabled As Boolean)
    SaveSetting APP_NAME, SECTION_CFG, "PrefetchOnLoad", IIf(enabled, "1", "0")
End Sub

'==============================================================================
' IsConfigured
' 최소 설정(SharePoint URL)이 완료되었는지 확인한다.
'==============================================================================
Public Function IsConfigured() As Boolean
    IsConfigured = Len(GetSharePointUrl) > 0
End Function
