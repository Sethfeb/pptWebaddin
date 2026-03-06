Attribute VB_Name = "Install"
'==============================================================================
' Install.bas  -  EquipSpec Add-in Auto Installer
'
' Usage:
'   1. Alt+F11 -> VBA Editor
'   2. File -> Import File -> Install.bas
'   3. Ctrl+G -> type:  Install.RunInstall  -> Enter
'==============================================================================
Option Explicit

Private Const PROJ_PATH As String = "D:\development\PPT-EquipSpec-Addin\src\"

' Actual form names after creation (may differ from "frmSearch" if rename fails)
Public g_searchFormName   As String
Public g_settingsFormName As String

Public Sub RunInstall()
    Dim vbComps As Object

    On Error GoTo NeedTrust
    Set vbComps = ActivePresentation.VBProject.VBComponents
    On Error GoTo 0

    Dim msg1 As String
    msg1 = "EquipSpec Add-in installer started." & vbCrLf & "Path: " & PROJ_PATH
    MsgBox msg1, vbInformation, "EquipSpec Add-in"

    RemoveIfExists vbComps, "modTypes"
    RemoveIfExists vbComps, "modSettings"
    RemoveIfExists vbComps, "modCache"
    RemoveIfExists vbComps, "modSharePoint"
    RemoveIfExists vbComps, "modShortCode"
    RemoveIfExists vbComps, "modRibbon"
    RemoveIfExists vbComps, "frmSearch"
    RemoveIfExists vbComps, "frmSettings"
    RemoveIfExists vbComps, "UserForm1"
    RemoveIfExists vbComps, "UserForm2"

    ImportBas vbComps, PROJ_PATH & "modules\modTypes.bas"
    ImportBas vbComps, PROJ_PATH & "modules\modSettings.bas"
    ImportBas vbComps, PROJ_PATH & "modules\modCache.bas"
    ImportBas vbComps, PROJ_PATH & "modules\modSharePoint.bas"
    ImportBas vbComps, PROJ_PATH & "modules\modShortCode.bas"

    BuildFormSearch vbComps
    BuildFormSettings vbComps

    ImportBas vbComps, PROJ_PATH & "modules\modRibbon.bas"

    Dim msg2 As String
    msg2 = "Installation complete!" & vbCrLf & vbCrLf
    msg2 = msg2 & "Search form name  : " & g_searchFormName & vbCrLf
    msg2 = msg2 & "Settings form name: " & g_settingsFormName & vbCrLf & vbCrLf
    msg2 = msg2 & "Next steps:" & vbCrLf
    msg2 = msg2 & "1. File -> Save As -> .pptm" & vbCrLf
    msg2 = msg2 & "2. Apply customUI.xml via Custom UI Editor" & vbCrLf
    msg2 = msg2 & "3. Save As .ppam -> copy to %APPDATA%\Microsoft\AddIns\"
    MsgBox msg2, vbInformation, "EquipSpec Add-in"
    Exit Sub

NeedTrust:
    Dim errMsg As String
    errMsg = "[Error] Cannot access VBA project." & vbCrLf & vbCrLf
    errMsg = errMsg & "Fix steps:" & vbCrLf
    errMsg = errMsg & "1. File -> Options -> Trust Center -> Trust Center Settings" & vbCrLf
    errMsg = errMsg & "2. Macro Settings tab" & vbCrLf
    errMsg = errMsg & "3. Check [Trust access to the VBA project object model]" & vbCrLf
    errMsg = errMsg & "4. OK -> Restart PowerPoint -> Retry"
    MsgBox errMsg, vbCritical, "EquipSpec Add-in - Trust Required"
End Sub

Private Sub RemoveIfExists(ByVal comps As Object, ByVal nm As String)
    Dim c As Object
    On Error Resume Next
    Set c = comps(nm)
    On Error GoTo 0
    If Not c Is Nothing Then comps.Remove c
End Sub

Private Sub ImportBas(ByVal comps As Object, ByVal fp As String)
    If Dir(fp) = "" Then
        MsgBox "File not found: " & fp, vbExclamation, "Install Error"
        Exit Sub
    End If
    comps.Import fp
End Sub

'==============================================================================
' BuildFormSearch
'==============================================================================
Private Sub BuildFormSearch(ByVal comps As Object)
    Dim frm As Object
    Dim d   As Object
    Dim c   As Object

    Set frm = comps.Add(3)

    ' Try to rename - ignore error if not supported
    On Error Resume Next
    frm.Name = "frmSearch"
    On Error GoTo 0

    g_searchFormName = frm.Name   ' Save actual name

    Set d = frm.Designer
    d.Caption = "Equip Spec Search"

    Set c = d.Controls.Add("Forms.Label.1")
    c.Name = "lblSearch": c.Caption = "Keyword:"
    c.Left = 6: c.Top = 10: c.Width = 48: c.Height = 14

    Set c = d.Controls.Add("Forms.TextBox.1")
    c.Name = "txtSearch"
    c.Left = 60: c.Top = 8: c.Width = 280: c.Height = 18

    Set c = d.Controls.Add("Forms.CommandButton.1")
    c.Name = "btnSearch": c.Caption = "Search"
    c.Left = 348: c.Top = 7: c.Width = 54: c.Height = 20

    Set c = d.Controls.Add("Forms.CommandButton.1")
    c.Name = "btnClear": c.Caption = "Clear"
    c.Left = 408: c.Top = 7: c.Width = 54: c.Height = 20

    Set c = d.Controls.Add("Forms.ListBox.1")
    c.Name = "lstResults": c.ColumnCount = 4
    c.ColumnWidths = "60 pt;90 pt;120 pt;60 pt"
    c.Left = 6: c.Top = 32: c.Width = 456: c.Height = 150

    Set c = d.Controls.Add("Forms.Label.1")
    c.Name = "lblDetail": c.Caption = "--- Selected Item ---"
    c.Left = 6: c.Top = 188: c.Width = 456: c.Height = 14

    Set c = d.Controls.Add("Forms.Label.1")
    c.Name = "lblEquipID": c.Caption = "EquipID:"
    c.Left = 6: c.Top = 206: c.Width = 54: c.Height = 14

    Set c = d.Controls.Add("Forms.TextBox.1")
    c.Name = "txtEquipID": c.Locked = True
    c.Left = 66: c.Top = 204: c.Width = 100: c.Height = 18

    Set c = d.Controls.Add("Forms.Label.1")
    c.Name = "lblSpecName": c.Caption = "Spec Name:"
    c.Left = 174: c.Top = 206: c.Width = 60: c.Height = 14

    Set c = d.Controls.Add("Forms.TextBox.1")
    c.Name = "txtSpecName": c.Locked = True
    c.Left = 240: c.Top = 204: c.Width = 150: c.Height = 18

    Set c = d.Controls.Add("Forms.Label.1")
    c.Name = "lblSpecValue": c.Caption = "Value:"
    c.Left = 6: c.Top = 228: c.Width = 54: c.Height = 14

    Set c = d.Controls.Add("Forms.TextBox.1")
    c.Name = "txtSpecValue": c.Locked = True
    c.Left = 66: c.Top = 226: c.Width = 100: c.Height = 18

    Set c = d.Controls.Add("Forms.Label.1")
    c.Name = "lblUnit": c.Caption = "Unit:"
    c.Left = 174: c.Top = 228: c.Width = 40: c.Height = 14

    Set c = d.Controls.Add("Forms.TextBox.1")
    c.Name = "txtUnit": c.Locked = True
    c.Left = 220: c.Top = 226: c.Width = 80: c.Height = 18

    Set c = d.Controls.Add("Forms.CommandButton.1")
    c.Name = "btnInsert": c.Caption = "Insert to Slide": c.Enabled = False
    c.Left = 6: c.Top = 252: c.Width = 110: c.Height = 22

    Set c = d.Controls.Add("Forms.CommandButton.1")
    c.Name = "btnInsertFull": c.Caption = "Insert Full Info": c.Enabled = False
    c.Left = 122: c.Top = 252: c.Width = 110: c.Height = 22

    Set c = d.Controls.Add("Forms.CommandButton.1")
    c.Name = "btnClose": c.Caption = "Close"
    c.Left = 396: c.Top = 252: c.Width = 66: c.Height = 22

    Set c = d.Controls.Add("Forms.Label.1")
    c.Name = "lblStatus": c.Caption = "Enter keyword and click Search."
    c.ForeColor = &H808080
    c.Left = 6: c.Top = 282: c.Width = 456: c.Height = 14

    InjectSearchCode frm.CodeModule
End Sub

Private Sub InjectSearchCode(ByVal cm As Object)
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim L As Long
    L = 1

    cm.InsertLines L, "Option Explicit":                                                           L = L + 1
    cm.InsertLines L, "' m_results: Variant array of SpecRecord fields":                           L = L + 1
    cm.InsertLines L, "' Each element: Array(EquipID, ShortCode, SpecName, SpecValue, Unit, Rev)": L = L + 1
    cm.InsertLines L, "Private m_results() As Variant":                                            L = L + 1
    cm.InsertLines L, "Private m_hasResults As Boolean":                                           L = L + 1
    cm.InsertLines L, "":                                                                          L = L + 1
    cm.InsertLines L, "Private Sub UserForm_Initialize()":                                        L = L + 1
    cm.InsertLines L, "    m_hasResults = False":                                                  L = L + 1
    cm.InsertLines L, "    With lstResults":                                                       L = L + 1
    cm.InsertLines L, "        .ColumnCount = 4":                                                  L = L + 1
    cm.InsertLines L, "        .ColumnWidths = ""60 pt;90 pt;120 pt;60 pt""":                      L = L + 1
    cm.InsertLines L, "        .ColumnHeads = False":                                              L = L + 1
    cm.InsertLines L, "    End With":                                                              L = L + 1
    cm.InsertLines L, "    ClearDetail":                                                           L = L + 1
    cm.InsertLines L, "    lblStatus.Caption = ""Enter keyword and click Search.""":               L = L + 1
    cm.InsertLines L, "    btnInsert.Enabled = False":                                             L = L + 1
    cm.InsertLines L, "    btnInsertFull.Enabled = False":                                         L = L + 1
    cm.InsertLines L, "    If modCache.IsFullyLoaded Then":                                        L = L + 1
    cm.InsertLines L, "        lblStatus.Caption = ""Cache loaded ("" & modCache.CacheCount & "" items). Enter keyword.""": L = L + 1
    cm.InsertLines L, "    End If":                                                                L = L + 1
    cm.InsertLines L, "End Sub":                                                                   L = L + 1
    cm.InsertLines L, "":                                                                          L = L + 1
    cm.InsertLines L, "Private Sub txtSearch_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)": L = L + 1
    cm.InsertLines L, "    If KeyCode = 13 Then DoSearch":                                         L = L + 1
    cm.InsertLines L, "End Sub":                                                                   L = L + 1
    cm.InsertLines L, "":                                                                          L = L + 1
    cm.InsertLines L, "Private Sub btnSearch_Click()":                                            L = L + 1
    cm.InsertLines L, "    DoSearch":                                                              L = L + 1
    cm.InsertLines L, "End Sub":                                                                   L = L + 1
    cm.InsertLines L, "":                                                                          L = L + 1
    cm.InsertLines L, "Private Sub btnClear_Click()":                                             L = L + 1
    cm.InsertLines L, "    txtSearch.Text = """"":                                                 L = L + 1
    cm.InsertLines L, "    lstResults.Clear":                                                      L = L + 1
    cm.InsertLines L, "    ClearDetail":                                                           L = L + 1
    cm.InsertLines L, "    m_hasResults = False":                                                  L = L + 1
    cm.InsertLines L, "    btnInsert.Enabled = False":                                             L = L + 1
    cm.InsertLines L, "    btnInsertFull.Enabled = False":                                         L = L + 1
    cm.InsertLines L, "    lblStatus.Caption = ""Enter keyword.""":                                L = L + 1
    cm.InsertLines L, "    txtSearch.SetFocus":                                                    L = L + 1
    cm.InsertLines L, "End Sub":                                                                   L = L + 1
    cm.InsertLines L, "":                                                                          L = L + 1
    cm.InsertLines L, "Private Sub lstResults_Click()":                                           L = L + 1
    cm.InsertLines L, "    Dim idx As Long":                                                       L = L + 1
    cm.InsertLines L, "    idx = lstResults.ListIndex":                                            L = L + 1
    cm.InsertLines L, "    If idx < 0 Then":                                                       L = L + 1
    cm.InsertLines L, "        ClearDetail":                                                       L = L + 1
    cm.InsertLines L, "        btnInsert.Enabled = False":                                         L = L + 1
    cm.InsertLines L, "        btnInsertFull.Enabled = False":                                     L = L + 1
    cm.InsertLines L, "        Exit Sub":                                                          L = L + 1
    cm.InsertLines L, "    End If":                                                                L = L + 1
    cm.InsertLines L, "    If Not m_hasResults Then Exit Sub":                                     L = L + 1
    cm.InsertLines L, "    If idx > UBound(m_results) Then Exit Sub":                             L = L + 1
    cm.InsertLines L, "    txtEquipID.Text  = m_results(idx)(0)":                                  L = L + 1
    cm.InsertLines L, "    txtSpecName.Text  = m_results(idx)(2)":                                 L = L + 1
    cm.InsertLines L, "    txtSpecValue.Text = m_results(idx)(3)":                                 L = L + 1
    cm.InsertLines L, "    txtUnit.Text      = m_results(idx)(4)":                                 L = L + 1
    cm.InsertLines L, "    btnInsert.Enabled = True":                                              L = L + 1
    cm.InsertLines L, "    btnInsertFull.Enabled = True":                                          L = L + 1
    cm.InsertLines L, "End Sub":                                                                   L = L + 1
    cm.InsertLines L, "":                                                                          L = L + 1
    cm.InsertLines L, "Private Sub lstResults_DblClick(ByVal Cancel As Integer)":                 L = L + 1
    cm.InsertLines L, "    If btnInsert.Enabled Then btnInsert_Click":                             L = L + 1
    cm.InsertLines L, "End Sub":                                                                   L = L + 1
    cm.InsertLines L, "":                                                                          L = L + 1
    cm.InsertLines L, "Private Sub btnInsert_Click()":                                            L = L + 1
    cm.InsertLines L, "    Dim idx As Long":                                                       L = L + 1
    cm.InsertLines L, "    idx = lstResults.ListIndex":                                            L = L + 1
    cm.InsertLines L, "    If idx < 0 Or Not m_hasResults Then":                                   L = L + 1
    cm.InsertLines L, "        MsgBox ""Please select an item."", vbInformation, ""EquipSpec""":   L = L + 1
    cm.InsertLines L, "        Exit Sub":                                                          L = L + 1
    cm.InsertLines L, "    End If":                                                                L = L + 1
    cm.InsertLines L, "    Dim specVal As String":                                                 L = L + 1
    cm.InsertLines L, "    Dim unitVal As String":                                                 L = L + 1
    cm.InsertLines L, "    specVal = m_results(idx)(3)":                                           L = L + 1
    cm.InsertLines L, "    unitVal = m_results(idx)(4)":                                           L = L + 1
    cm.InsertLines L, "    modShortCode.InsertTextToSelection specVal, unitVal":                   L = L + 1
    cm.InsertLines L, "    lblStatus.Caption = ""Inserted: "" & specVal & "" "" & unitVal":        L = L + 1
    cm.InsertLines L, "End Sub":                                                                   L = L + 1
    cm.InsertLines L, "":                                                                          L = L + 1
    cm.InsertLines L, "Private Sub btnInsertFull_Click()":                                        L = L + 1
    cm.InsertLines L, "    Dim idx As Long":                                                       L = L + 1
    cm.InsertLines L, "    idx = lstResults.ListIndex":                                            L = L + 1
    cm.InsertLines L, "    If idx < 0 Or Not m_hasResults Then":                                   L = L + 1
    cm.InsertLines L, "        MsgBox ""Please select an item."", vbInformation, ""EquipSpec""":   L = L + 1
    cm.InsertLines L, "        Exit Sub":                                                          L = L + 1
    cm.InsertLines L, "    End If":                                                                L = L + 1
    cm.InsertLines L, "    Dim specName As String":                                                L = L + 1
    cm.InsertLines L, "    Dim specVal  As String":                                                L = L + 1
    cm.InsertLines L, "    Dim unitVal  As String":                                                L = L + 1
    cm.InsertLines L, "    specName = m_results(idx)(2)":                                          L = L + 1
    cm.InsertLines L, "    specVal  = m_results(idx)(3)":                                          L = L + 1
    cm.InsertLines L, "    unitVal  = m_results(idx)(4)":                                          L = L + 1
    cm.InsertLines L, "    modShortCode.InsertTextToSelection specName & "": "" & specVal, unitVal": L = L + 1
    cm.InsertLines L, "    lblStatus.Caption = ""Inserted: "" & specName & "": "" & specVal & "" "" & unitVal": L = L + 1
    cm.InsertLines L, "End Sub":                                                                   L = L + 1
    cm.InsertLines L, "":                                                                          L = L + 1
    cm.InsertLines L, "Private Sub btnClose_Click()":                                             L = L + 1
    cm.InsertLines L, "    Me.Hide":                                                               L = L + 1
    cm.InsertLines L, "End Sub":                                                                   L = L + 1
    cm.InsertLines L, "":                                                                          L = L + 1
    cm.InsertLines L, "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)":  L = L + 1
    cm.InsertLines L, "    If CloseMode = vbFormControlMenu Then":                                 L = L + 1
    cm.InsertLines L, "        Cancel = True":                                                     L = L + 1
    cm.InsertLines L, "        Me.Hide":                                                           L = L + 1
    cm.InsertLines L, "    End If":                                                                L = L + 1
    cm.InsertLines L, "End Sub":                                                                   L = L + 1
    cm.InsertLines L, "":                                                                          L = L + 1
    cm.InsertLines L, "Private Sub DoSearch()":                                                   L = L + 1
    cm.InsertLines L, "    Dim keyword As String":                                                 L = L + 1
    cm.InsertLines L, "    Dim siteUrl As String":                                                 L = L + 1
    cm.InsertLines L, "    Dim i As Long":                                                         L = L + 1
    cm.InsertLines L, "    keyword = Trim(txtSearch.Text)":                                        L = L + 1
    cm.InsertLines L, "    lstResults.Clear":                                                      L = L + 1
    cm.InsertLines L, "    ClearDetail":                                                           L = L + 1
    cm.InsertLines L, "    m_hasResults = False":                                                  L = L + 1
    cm.InsertLines L, "    btnInsert.Enabled = False":                                             L = L + 1
    cm.InsertLines L, "    btnInsertFull.Enabled = False":                                         L = L + 1
    cm.InsertLines L, "    lblStatus.Caption = ""Searching...""":                                  L = L + 1
    cm.InsertLines L, "    Me.Repaint":                                                            L = L + 1
    cm.InsertLines L, "    On Error GoTo ErrHandler":                                              L = L + 1
    cm.InsertLines L, "    siteUrl = modSettings.GetSharePointUrl":                                L = L + 1
    cm.InsertLines L, "    ' SearchVariant returns Variant() each element=Array(EquipID,ShortCode,SpecName,SpecValue,Unit,Rev)": L = L + 1
    cm.InsertLines L, "    m_results = modCache.SearchVariant(keyword, siteUrl)":                  L = L + 1
    cm.InsertLines L, "    If UBound(m_results) < 0 Then":                                         L = L + 1
    cm.InsertLines L, "        lblStatus.Caption = ""No results found.""":                         L = L + 1
    cm.InsertLines L, "        Exit Sub":                                                          L = L + 1
    cm.InsertLines L, "    End If":                                                                L = L + 1
    cm.InsertLines L, "    For i = 0 To UBound(m_results)":                                        L = L + 1
    cm.InsertLines L, "        lstResults.AddItem m_results(i)(0)":                                L = L + 1
    cm.InsertLines L, "        lstResults.List(i, 1) = m_results(i)(1)":                           L = L + 1
    cm.InsertLines L, "        lstResults.List(i, 2) = m_results(i)(2)":                           L = L + 1
    cm.InsertLines L, "        lstResults.List(i, 3) = m_results(i)(3) & "" "" & m_results(i)(4)": L = L + 1
    cm.InsertLines L, "    Next i":                                                                L = L + 1
    cm.InsertLines L, "    m_hasResults = True":                                                   L = L + 1
    cm.InsertLines L, "    lblStatus.Caption = UBound(m_results) + 1 & "" item(s) found.""":       L = L + 1
    cm.InsertLines L, "    Exit Sub":                                                              L = L + 1
    cm.InsertLines L, "ErrHandler:":                                                               L = L + 1
    cm.InsertLines L, "    lblStatus.Caption = ""Search error: "" & Err.Description":              L = L + 1
    cm.InsertLines L, "End Sub":                                                                   L = L + 1
    cm.InsertLines L, "":                                                                          L = L + 1
    cm.InsertLines L, "Private Sub ClearDetail()":                                                L = L + 1
    cm.InsertLines L, "    txtEquipID.Text   = """"":                                              L = L + 1
    cm.InsertLines L, "    txtSpecName.Text  = """"":                                              L = L + 1
    cm.InsertLines L, "    txtSpecValue.Text = """"":                                              L = L + 1
    cm.InsertLines L, "    txtUnit.Text      = """"":                                              L = L + 1
    cm.InsertLines L, "End Sub":                                                                   L = L + 1
End Sub

'==============================================================================
' BuildFormSettings
'==============================================================================
Private Sub BuildFormSettings(ByVal comps As Object)
    Dim frm As Object
    Dim d   As Object
    Dim c   As Object

    Set frm = comps.Add(3)

    On Error Resume Next
    frm.Name = "frmSettings"
    On Error GoTo 0

    g_settingsFormName = frm.Name

    Set d = frm.Designer
    d.Caption = "SharePoint Settings"

    Set c = d.Controls.Add("Forms.Label.1")
    c.Name = "lblUrl": c.Caption = "SharePoint Site URL:"
    c.Left = 6: c.Top = 10: c.Width = 150: c.Height = 14

    Set c = d.Controls.Add("Forms.TextBox.1")
    c.Name = "txtUrl"
    c.Left = 6: c.Top = 26: c.Width = 330: c.Height = 18

    Set c = d.Controls.Add("Forms.Label.1")
    c.Name = "lblListName": c.Caption = "List Name:"
    c.Left = 6: c.Top = 52: c.Width = 80: c.Height = 14

    Set c = d.Controls.Add("Forms.TextBox.1")
    c.Name = "txtListName"
    c.Left = 6: c.Top = 68: c.Width = 160: c.Height = 18

    Set c = d.Controls.Add("Forms.CheckBox.1")
    c.Name = "chkPrefetch": c.Caption = "Prefetch all specs on startup"
    c.Value = True
    c.Left = 6: c.Top = 94: c.Width = 220: c.Height = 16

    Set c = d.Controls.Add("Forms.CommandButton.1")
    c.Name = "btnTest": c.Caption = "Test Connection"
    c.Left = 6: c.Top = 118: c.Width = 100: c.Height = 22

    Set c = d.Controls.Add("Forms.CommandButton.1")
    c.Name = "btnSave": c.Caption = "Save"
    c.Left = 200: c.Top = 118: c.Width = 60: c.Height = 22

    Set c = d.Controls.Add("Forms.CommandButton.1")
    c.Name = "btnCancel": c.Caption = "Cancel"
    c.Left = 270: c.Top = 118: c.Width = 60: c.Height = 22

    Set c = d.Controls.Add("Forms.Label.1")
    c.Name = "lblStatus": c.Caption = ""
    c.ForeColor = &HFF0000
    c.Left = 6: c.Top = 148: c.Width = 330: c.Height = 14

    InjectSettingsCode frm.CodeModule
End Sub

Private Sub InjectSettingsCode(ByVal cm As Object)
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines

    Dim L As Long
    L = 1

    cm.InsertLines L, "Option Explicit":                                                           L = L + 1
    cm.InsertLines L, "":                                                                          L = L + 1
    cm.InsertLines L, "Private Sub UserForm_Initialize()":                                        L = L + 1
    cm.InsertLines L, "    txtUrl.Text = modSettings.GetSharePointUrl":                            L = L + 1
    cm.InsertLines L, "    txtListName.Text = modSettings.GetListName":                            L = L + 1
    cm.InsertLines L, "    chkPrefetch.Value = modSettings.GetPrefetchOnLoad":                     L = L + 1
    cm.InsertLines L, "    lblStatus.Caption = """"":                                              L = L + 1
    cm.InsertLines L, "End Sub":                                                                   L = L + 1
    cm.InsertLines L, "":                                                                          L = L + 1
    cm.InsertLines L, "Private Sub btnTest_Click()":                                              L = L + 1
    cm.InsertLines L, "    Dim url As String":                                                     L = L + 1
    cm.InsertLines L, "    Dim listName As String":                                                L = L + 1
    cm.InsertLines L, "    Dim endpoint As String":                                                L = L + 1
    cm.InsertLines L, "    Dim http As Object":                                                    L = L + 1
    cm.InsertLines L, "    url = Trim(txtUrl.Text)":                                               L = L + 1
    cm.InsertLines L, "    listName = Trim(txtListName.Text)":                                     L = L + 1
    cm.InsertLines L, "    If Len(url) = 0 Then lblStatus.Caption = ""Enter URL."": Exit Sub":     L = L + 1
    cm.InsertLines L, "    lblStatus.Caption = ""Testing connection...""":                         L = L + 1
    cm.InsertLines L, "    Me.Repaint":                                                            L = L + 1
    cm.InsertLines L, "    On Error GoTo ErrHandler":                                              L = L + 1
    cm.InsertLines L, "    endpoint = url & ""/_api/web/lists/getbytitle('"" & listName & ""')?$select=Title""": L = L + 1
    cm.InsertLines L, "    Set http = CreateObject(""WinHttp.WinHttpRequest.5.1"")":               L = L + 1
    cm.InsertLines L, "    http.Open ""GET"", endpoint, False":                                    L = L + 1
    cm.InsertLines L, "    http.SetAutoLogonPolicy 0":                                             L = L + 1
    cm.InsertLines L, "    http.SetRequestHeader ""Accept"", ""application/json;odata=verbose""":  L = L + 1
    cm.InsertLines L, "    http.SetTimeouts 10000, 10000, 10000, 10000":                           L = L + 1
    cm.InsertLines L, "    http.Send":                                                             L = L + 1
    cm.InsertLines L, "    If http.Status = 200 Then":                                             L = L + 1
    cm.InsertLines L, "        lblStatus.Caption = ""OK - List '"" & listName & ""' found.""":     L = L + 1
    cm.InsertLines L, "    Else":                                                                  L = L + 1
    cm.InsertLines L, "        lblStatus.Caption = ""Failed: HTTP "" & http.Status & "" "" & http.StatusText": L = L + 1
    cm.InsertLines L, "    End If":                                                                L = L + 1
    cm.InsertLines L, "    Set http = Nothing":                                                    L = L + 1
    cm.InsertLines L, "    Exit Sub":                                                              L = L + 1
    cm.InsertLines L, "ErrHandler:":                                                               L = L + 1
    cm.InsertLines L, "    If Not http Is Nothing Then Set http = Nothing":                        L = L + 1
    cm.InsertLines L, "    lblStatus.Caption = ""Error: "" & Err.Description":                     L = L + 1
    cm.InsertLines L, "End Sub":                                                                   L = L + 1
    cm.InsertLines L, "":                                                                          L = L + 1
    cm.InsertLines L, "Private Sub btnSave_Click()":                                              L = L + 1
    cm.InsertLines L, "    Dim url As String":                                                     L = L + 1
    cm.InsertLines L, "    Dim listName As String":                                                L = L + 1
    cm.InsertLines L, "    url = Trim(txtUrl.Text)":                                               L = L + 1
    cm.InsertLines L, "    listName = Trim(txtListName.Text)":                                     L = L + 1
    cm.InsertLines L, "    If Len(url) = 0 Then":                                                  L = L + 1
    cm.InsertLines L, "        MsgBox ""Enter SharePoint URL."", vbExclamation, ""Settings""":     L = L + 1
    cm.InsertLines L, "        txtUrl.SetFocus: Exit Sub":                                         L = L + 1
    cm.InsertLines L, "    End If":                                                                L = L + 1
    cm.InsertLines L, "    If Len(listName) = 0 Then listName = ""EquipmentSpecs""":               L = L + 1
    cm.InsertLines L, "    If Right(url, 1) = ""/"" Then url = Left(url, Len(url) - 1)":           L = L + 1
    cm.InsertLines L, "    modSettings.SetSharePointUrl url":                                      L = L + 1
    cm.InsertLines L, "    modSettings.SetListName listName":                                      L = L + 1
    cm.InsertLines L, "    modSettings.SetPrefetchOnLoad chkPrefetch.Value":                       L = L + 1
    cm.InsertLines L, "    modCache.ClearCache":                                                   L = L + 1
    cm.InsertLines L, "    MsgBox ""Settings saved."", vbInformation, ""Settings""":               L = L + 1
    cm.InsertLines L, "    Me.Hide":                                                               L = L + 1
    cm.InsertLines L, "End Sub":                                                                   L = L + 1
    cm.InsertLines L, "":                                                                          L = L + 1
    cm.InsertLines L, "Private Sub btnCancel_Click()":                                            L = L + 1
    cm.InsertLines L, "    Me.Hide":                                                               L = L + 1
    cm.InsertLines L, "End Sub":                                                                   L = L + 1
End Sub
