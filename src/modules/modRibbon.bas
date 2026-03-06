Attribute VB_Name = "modRibbon"
Option Explicit

Private m_ribbon As Object

Public Sub RibbonOnLoad(ByVal ribbon As Object)
    Set m_ribbon = ribbon
    If Not modSettings.IsConfigured Then
        Dim m As String
        m = "EquipSpec Add-in loaded." & vbCrLf & vbCrLf
        m = m & "Go to [EquipSpec] tab > SharePoint Settings" & vbCrLf
        m = m & "and enter your SharePoint URL first."
        MsgBox m, vbInformation, "EquipSpec Add-in"
        Exit Sub
    End If
    If modSettings.GetPrefetchOnLoad Then PrefetchAllSpecs
End Sub

Public Sub RibbonOpenSearch(ByVal control As Object)
    If Not modSettings.IsConfigured Then
        MsgBox "Configure SharePoint URL first." & vbCrLf & "[EquipSpec] tab > SharePoint Settings", vbExclamation, "EquipSpec Add-in"
        Exit Sub
    End If
    If Not modCache.IsFullyLoaded Then PrefetchAllSpecs
    OpenFormByType False
End Sub

Public Sub RibbonApplyShortCode(ByVal control As Object)
    If Not modSettings.IsConfigured Then
        MsgBox "Configure SharePoint URL first.", vbExclamation, "EquipSpec Add-in"
        Exit Sub
    End If
    modShortCode.ApplyShortCodeInSelection
End Sub

Public Sub RibbonRefreshCache(ByVal control As Object)
    If Not modSettings.IsConfigured Then
        MsgBox "Configure SharePoint URL first.", vbExclamation, "EquipSpec Add-in"
        Exit Sub
    End If
    modCache.ClearCache
    PrefetchAllSpecs
End Sub

Public Sub RibbonClearCache(ByVal control As Object)
    modCache.ClearCache
    MsgBox "Cache cleared. (" & modCache.CacheCount & " items removed)", vbInformation, "EquipSpec Add-in"
End Sub

Public Sub RibbonOpenSettings(ByVal control As Object)
    OpenFormByType True
End Sub

Public Sub RibbonAbout(ByVal control As Object)
    Dim m As String
    m = "EquipSpec Add-in  v1.0.0" & vbCrLf & vbCrLf
    m = m & "SharePoint-linked equipment spec auto-insert tool" & vbCrLf
    m = m & "ShortCode replace + Search panel" & vbCrLf & vbCrLf
    m = m & "Cache: " & modCache.CacheCount & " items" & vbCrLf
    m = m & "URL: " & modSettings.GetSharePointUrl
    MsgBox m, vbInformation, "About EquipSpec Add-in"
End Sub

'==============================================================================
' OpenFormByType
' Finds UserForms in the VBProject by scan order and shows them.
' isSettings=False -> 1st UserForm (Search)
' isSettings=True  -> 2nd UserForm (Settings)
' Uses Application.Run with string to avoid compile-time form reference.
'==============================================================================
Private Sub OpenFormByType(ByVal isSettings As Boolean)
    Dim vbComps As Object
    Dim c       As Object
    Dim nm      As String
    Dim cnt     As Integer

    On Error GoTo ErrHandler

    Set vbComps = ActivePresentation.VBProject.VBComponents
    cnt = 0

    For Each c In vbComps
        If c.Type = 3 Then
            cnt = cnt + 1
            If (isSettings And cnt = 2) Or (Not isSettings And cnt = 1) Then
                nm = c.Name
                Exit For
            End If
            ' Also match by expected name directly
            If Not isSettings And (c.Name = "frmSearch") Then
                nm = c.Name: Exit For
            End If
            If isSettings And (c.Name = "frmSettings") Then
                nm = c.Name: Exit For
            End If
        End If
    Next c

    If Len(nm) = 0 Then
        MsgBox "Form not found. Please re-run Install.RunInstall.", vbExclamation, "EquipSpec Add-in"
        Exit Sub
    End If

    ' Show via Application.Run - no compile-time dependency on form name
    If isSettings Then
        Application.Run nm & ".Show", vbModal
    Else
        Application.Run nm & ".Show", vbModeless
    End If
    Exit Sub

ErrHandler:
    MsgBox "Cannot open form. Error: " & Err.Description, vbExclamation, "EquipSpec Add-in"
End Sub

Private Sub PrefetchAllSpecs()
    Dim items() As SpecRecord
    Dim siteUrl As String
    siteUrl = modSettings.GetSharePointUrl
    On Error GoTo ErrHandler
    items = modSharePoint.GetAllSpecs(siteUrl)
    If UBound(items) >= 0 Then modCache.CachePutAll items
    Exit Sub
ErrHandler:
End Sub

Public Sub OnKeyboardShortcut()
    modShortCode.ApplyShortCodeInSelection
End Sub

Public Sub Auto_Open()
    If modSettings.IsConfigured And modSettings.GetPrefetchOnLoad Then
        PrefetchAllSpecs
    End If
End Sub
