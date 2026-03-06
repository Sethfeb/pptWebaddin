Attribute VB_Name = "modShortCode"
Option Explicit

Private m_isProcessing As Boolean
Private Const SHORTCODE_PATTERN As String = "![A-Za-z0-9_\-]+"

Public Function ApplyShortCodeInSelection() As Long
    Dim siteUrl  As String
    Dim sel      As Object
    Dim tf       As Object
    Dim replaced As Long

    ApplyShortCodeInSelection = 0
    If m_isProcessing Then Exit Function
    m_isProcessing = True

    On Error GoTo ErrHandler

    siteUrl = modSettings.GetSharePointUrl
    If Len(siteUrl) = 0 Then
        MsgBox "SharePoint URL is not configured." & vbCrLf & "[EquipSpec] tab > SharePoint Settings", vbExclamation, "EquipSpec Add-in"
        GoTo CleanUp
    End If

    If ActiveWindow Is Nothing Then GoTo CleanUp
    Set sel = ActiveWindow.Selection

    Select Case sel.Type
        Case ppSelectionText
            Set tf = sel.TextRange
            replaced = ReplaceShortCodesInRange(tf, siteUrl)

        Case ppSelectionShapes
            Dim shp As Object
            For Each shp In sel.ShapeRange
                If shp.HasTextFrame Then
                    Set tf = shp.TextFrame.TextRange
                    replaced = replaced + ReplaceShortCodesInRange(tf, siteUrl)
                End If
            Next shp

        Case Else
            MsgBox "Select a text box or enter text edit mode first.", vbInformation, "EquipSpec Add-in"
    End Select

    If replaced > 0 Then
        MsgBox replaced & " shortcode(s) replaced.", vbInformation, "EquipSpec Add-in"
    ElseIf sel.Type = ppSelectionText Or sel.Type = ppSelectionShapes Then
        MsgBox "No shortcodes (!xxx) found.", vbInformation, "EquipSpec Add-in"
    End If

    ApplyShortCodeInSelection = replaced

CleanUp:
    m_isProcessing = False
    Exit Function

ErrHandler:
    m_isProcessing = False
    MsgBox "Error during shortcode replace: " & Err.Description, vbCritical, "EquipSpec Add-in"
End Function

Private Function ReplaceShortCodesInRange(ByRef tf As Object, ByVal siteUrl As String) As Long
    Dim re        As Object
    Dim matches   As Object
    Dim fullText  As String
    Dim i         As Long
    Dim count     As Long
    Dim rec       As SpecRecord
    Dim newText   As String

    ReplaceShortCodesInRange = 0

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = SHORTCODE_PATTERN

    fullText = tf.Text
    Set matches = re.Execute(fullText)

    If matches.Count = 0 Then
        Set re = Nothing
        Exit Function
    End If

    Dim codeList() As String
    Dim posList()  As Long
    Dim lenList()  As Long

    ReDim codeList(0 To matches.Count - 1)
    ReDim posList(0 To matches.Count - 1)
    ReDim lenList(0 To matches.Count - 1)

    For i = 0 To matches.Count - 1
        codeList(i) = matches(i).Value
        posList(i)  = matches(i).FirstIndex
        lenList(i)  = matches(i).Length
    Next i

    Set re = Nothing

    For i = matches.Count - 1 To 0 Step -1
        Dim shortCode As String
        shortCode = codeList(i)

        Dim hit As Boolean
        hit = modCache.CacheGet(shortCode, rec)

        If Not hit Then
            hit = modSharePoint.GetSpecByShortCode(siteUrl, shortCode, rec)
            If hit Then modCache.CachePut rec
        End If

        If hit Then
            newText = rec.SpecValue
            If Len(Trim(rec.Unit)) > 0 Then newText = newText & " " & rec.Unit
            Dim startPos As Long
            startPos = posList(i) + 1
            tf.Characters(startPos, lenList(i)).Text = newText
            count = count + 1
        End If
    Next i

    ReplaceShortCodesInRange = count
End Function

Public Sub InsertSpecToSelection(ByRef rec As SpecRecord)
    Dim sel      As Object
    Dim tf       As Object
    Dim insertTx As String

    insertTx = rec.SpecValue
    If Len(Trim(rec.Unit)) > 0 Then insertTx = insertTx & " " & rec.Unit

    On Error GoTo ErrHandler

    If ActiveWindow Is Nothing Then Exit Sub
    Set sel = ActiveWindow.Selection

    Select Case sel.Type
        Case ppSelectionText
            sel.TextRange.Text = insertTx

        Case ppSelectionShapes
            If sel.ShapeRange.Count > 0 Then
                If sel.ShapeRange(1).HasTextFrame Then
                    Set tf = sel.ShapeRange(1).TextFrame.TextRange
                    tf.Text = tf.Text & insertTx
                End If
            End If

        Case ppSelectionNone, ppSelectionSlides
            Dim slide As Object
            Set slide = ActiveWindow.View.Slide
            Dim newShape As Object
            Set newShape = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 400, 50)
            newShape.TextFrame.TextRange.Text = insertTx

        Case Else
            MsgBox "Select a text box or slide first.", vbInformation, "EquipSpec Add-in"
    End Select

    Exit Sub

ErrHandler:
    MsgBox "Error inserting spec: " & Err.Description, vbCritical, "EquipSpec Add-in"
End Sub

'==============================================================================
' InsertTextToSelection
' UserForm에서 SpecRecord 없이 문자열만으로 삽입할 수 있는 헬퍼 함수
'==============================================================================
Public Sub InsertTextToSelection(ByVal specValue As String, ByVal unitValue As String)
    Dim rec As SpecRecord
    rec.SpecValue = specValue
    rec.Unit = unitValue
    InsertSpecToSelection rec
End Sub

Public Function ExtractShortCodesFromText(ByVal text As String) As String()
    Dim re       As Object
    Dim matches  As Object
    Dim result() As String
    Dim empty()  As String
    Dim i        As Long

    ReDim empty(0)
    empty(0) = ""

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = SHORTCODE_PATTERN

    Set matches = re.Execute(text)

    If matches.Count = 0 Then
        Set re = Nothing
        ExtractShortCodesFromText = empty
        Exit Function
    End If

    ReDim result(0 To matches.Count - 1)
    For i = 0 To matches.Count - 1
        result(i) = matches(i).Value
    Next i

    Set re = Nothing
    ExtractShortCodesFromText = result
End Function
