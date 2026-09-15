Attribute VB_Name = "modOutlookTemplateTool"
Option Explicit

' Outlook template tool for Excel and classic Outlook.
' This source file is intentionally ASCII-only so it imports reliably in VBE.
' Localized UI text is stored in the Settings sheet.

Private Const TOOL_SHEET_INDEX As Long = 1
Private Const TEMPLATE_SHEET_INDEX As Long = 2
Private Const SETTINGS_SHEET_INDEX As Long = 3

Private Const CELL_TEMPLATE As String = "B7"
Private Const CELL_AUTO_SEND As String = "B6"
Private Const CELL_PREVIEW_TO As String = "E7"
Private Const CELL_PREVIEW_CC As String = "E8"
Private Const CELL_PREVIEW_SUBJECT As String = "E9"
Private Const CELL_PREVIEW_BODY As String = "D12"
Private Const CELL_STATUS As String = "B32"

Private Const VARIABLE_FIRST_ROW As Long = 12
Private Const VARIABLE_LAST_ROW As Long = 30

Private mBusy As Boolean

Public Sub SetupWorkbook()
    On Error GoTo Failed

    Dim ws As Worksheet
    Set ws = ToolSheet()

    Application.ScreenUpdating = False
    DeleteControlIfPresent ws, "btnLoadTemplate"
    DeleteControlIfPresent ws, "btnPreview"
    DeleteControlIfPresent ws, "btnGenerateMail"
    DeleteControlIfPresent ws, "chkAutoSend"

    ws.Range("B8").ClearContents

    AddButton ws, "btnLoadTemplate", UiText("button.load"), ws.Range("A36"), 118, "LoadSelectedTemplate"
    AddButton ws, "btnPreview", UiText("button.preview"), ws.Range("B36"), 118, "RefreshPreview"
    AddButton ws, "btnGenerateMail", UiText("button.generate"), ws.Range("D36"), 150, "GenerateMail"

    Dim cb As CheckBox
    Set cb = ws.CheckBoxes.Add(ws.Range("B8").Left + 4, ws.Range("B8").Top + 3, 190, 18)
    With cb
        .Name = "chkAutoSend"
        .Caption = UiText("checkbox.autosend")
        .LinkedCell = "'" & SettingsSheet().Name & "'!$" & CELL_AUTO_SEND
        .Value = IIf(CBool(SettingsSheet().Range(CELL_AUTO_SEND).Value), 1, -4146)
        .OnAction = "UpdateActionHint"
    End With

    ws.Rows(36).RowHeight = 28
    LoadSelectedTemplate
    SetStatus UiText("status.initialized"), False

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

Failed:
    MsgBox UiText("error.setup") & Err.Description, vbCritical, UiText("app.title")
    Resume CleanExit
End Sub

Public Sub LoadSelectedTemplate()
    On Error GoTo Failed

    Dim rowNumber As Long
    rowNumber = GetTemplateRow()

    Dim wsTool As Worksheet
    Dim wsTemplates As Worksheet
    Set wsTool = ToolSheet()
    Set wsTemplates = TemplateSheet()

    Dim previousValues As Object
    Set previousValues = ReadManualValues()

    Dim orderedVariables As Collection
    Dim seen As Object
    Set orderedVariables = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1

    AddTokensFromText CStr(wsTemplates.Cells(rowNumber, 2).Value2), orderedVariables, seen
    AddTokensFromText CStr(wsTemplates.Cells(rowNumber, 3).Value2), orderedVariables, seen
    AddTokensFromText CStr(wsTemplates.Cells(rowNumber, 4).Value2), orderedVariables, seen
    AddTokensFromText CStr(wsTemplates.Cells(rowNumber, 5).Value2), orderedVariables, seen

    If orderedVariables.Count > VARIABLE_LAST_ROW - VARIABLE_FIRST_ROW + 1 Then
        Err.Raise vbObjectError + 701, , UiText("error.variables")
    End If

    wsTool.Range("A" & VARIABLE_FIRST_ROW & ":B" & VARIABLE_LAST_ROW).ClearContents
    wsTool.Range("A11").Value = "batch_date"
    wsTool.Range("B11").Value = DateAdd("d", -1, Date)
    wsTool.Range("B11").NumberFormat = "yyyy-mm-dd"

    Dim i As Long
    Dim variableName As String
    For i = 1 To orderedVariables.Count
        variableName = CStr(orderedVariables(i))
        wsTool.Cells(VARIABLE_FIRST_ROW + i - 1, 1).Value = variableName
        If previousValues.Exists(variableName) Then
            wsTool.Cells(VARIABLE_FIRST_ROW + i - 1, 2).Value = previousValues(variableName)
        End If
    Next i

    RefreshPreview
    Exit Sub

Failed:
    SetStatus UiText("error.load") & Err.Description, True
    MsgBox UiText("error.load") & Err.Description, vbExclamation, UiText("app.title")
End Sub

Public Sub RefreshPreview()
    On Error GoTo Failed

    Dim renderedTo As String
    Dim renderedCc As String
    Dim renderedSubject As String
    Dim renderedBody As String
    Dim missing As String

    RenderSelectedTemplate renderedTo, renderedCc, renderedSubject, renderedBody, missing

    Dim ws As Worksheet
    Set ws = ToolSheet()
    ws.Range(CELL_PREVIEW_TO).Value = renderedTo
    ws.Range(CELL_PREVIEW_CC).Value = renderedCc
    ws.Range(CELL_PREVIEW_SUBJECT).Value = renderedSubject
    ws.Range(CELL_PREVIEW_BODY).Value = renderedBody

    If Len(missing) > 0 Then
        SetStatus UiText("status.missing") & missing, True
    Else
        SetStatus UiText("status.preview"), False
    End If
    Exit Sub

Failed:
    SetStatus UiText("error.preview") & Err.Description, True
    MsgBox UiText("error.preview") & Err.Description, vbExclamation, UiText("app.title")
End Sub

Public Sub GenerateMail()
    If mBusy Then Exit Sub
    mBusy = True
    On Error GoTo Failed

    Dim renderedTo As String
    Dim renderedCc As String
    Dim renderedSubject As String
    Dim renderedBody As String
    Dim missing As String

    RenderSelectedTemplate renderedTo, renderedCc, renderedSubject, renderedBody, missing

    If Len(missing) > 0 Then
        Err.Raise vbObjectError + 702, , UiText("status.missing") & missing
    End If
    If Len(Trim$(renderedTo)) = 0 Then
        Err.Raise vbObjectError + 703, , UiText("error.to")
    End If

    Dim outlookApp As Object
    Set outlookApp = GetOutlookApplication()
    If outlookApp Is Nothing Then
        Err.Raise vbObjectError + 704, , UiText("error.outlook")
    End If

    Dim mail As Object
    Set mail = outlookApp.CreateItem(0)
    With mail
        .To = renderedTo
        .CC = renderedCc
        .Subject = renderedSubject
        .Body = renderedBody
    End With

    If Not mail.Recipients.ResolveAll Then
        Err.Raise vbObjectError + 705, , UiText("error.resolve")
    End If

    If AutoSendEnabled() Then
        Dim prompt As String
        prompt = BuildConfirmation(renderedTo, renderedCc, renderedSubject, renderedBody)
        If MsgBox(prompt, vbYesNo + vbExclamation + vbDefaultButton2, UiText("confirm.title")) <> vbYes Then
            SetStatus UiText("status.cancelled"), False
            GoTo CleanExit
        End If
        mail.Send
        SetStatus UiText("status.sent"), False
    Else
        mail.Display
        SetStatus UiText("status.opened"), False
    End If

CleanExit:
    mBusy = False
    Exit Sub

Failed:
    SetStatus UiText("error.action") & Err.Description, True
    MsgBox UiText("error.action") & Err.Description, vbCritical, UiText("app.title")
    Resume CleanExit
End Sub

Public Sub UpdateActionHint()
    If AutoSendEnabled() Then
        SetStatus UiText("status.autosend.on"), False
    Else
        SetStatus UiText("status.autosend.off"), False
    End If
End Sub

Private Sub RenderSelectedTemplate(ByRef renderedTo As String, _
                                   ByRef renderedCc As String, _
                                   ByRef renderedSubject As String, _
                                   ByRef renderedBody As String, _
                                   ByRef missingText As String)
    Dim rowNumber As Long
    rowNumber = GetTemplateRow()

    Dim ws As Worksheet
    Set ws = TemplateSheet()

    Dim values As Object
    Dim missing As Object
    Set values = ReadManualValues()
    Set missing = CreateObject("Scripting.Dictionary")
    missing.CompareMode = 1

    renderedTo = RenderText(CStr(ws.Cells(rowNumber, 2).Value2), values, missing)
    renderedCc = RenderText(CStr(ws.Cells(rowNumber, 3).Value2), values, missing)
    renderedSubject = RenderText(CStr(ws.Cells(rowNumber, 4).Value2), values, missing)
    renderedBody = RenderText(CStr(ws.Cells(rowNumber, 5).Value2), values, missing)

    Dim key As Variant
    For Each key In missing.Keys
        If Len(missingText) > 0 Then missingText = missingText & "; "
        missingText = missingText & CStr(key)
    Next key
End Sub

Private Function RenderText(ByVal sourceText As String, ByVal values As Object, ByVal missing As Object) As String
    Dim result As String
    result = sourceText

    Dim startAt As Long
    Dim endAt As Long
    Dim token As String
    Dim replacement As String
    Dim recognizedDate As Boolean

    startAt = InStr(1, result, "{{", vbBinaryCompare)
    Do While startAt > 0
        endAt = InStr(startAt + 2, result, "}}", vbBinaryCompare)
        If endAt = 0 Then
            Err.Raise vbObjectError + 706, , UiText("error.unclosed")
        End If

        token = Trim$(Mid$(result, startAt + 2, endAt - startAt - 2))
        If Len(token) = 0 Then
            Err.Raise vbObjectError + 707, , UiText("error.emptytoken")
        End If

        recognizedDate = False
        replacement = RenderDateToken(token, recognizedDate)
        If Not recognizedDate Then
            If values.Exists(token) And Len(Trim$(CStr(values(token)))) > 0 Then
                replacement = CStr(values(token))
            Else
                replacement = "{{" & token & "}}"
                If Not missing.Exists(token) Then missing.Add token, True
            End If
        End If

        result = Left$(result, startAt - 1) & replacement & Mid$(result, endAt + 2)
        startAt = InStr(startAt + Len(replacement), result, "{{", vbBinaryCompare)
    Loop

    RenderText = result
End Function

Private Function RenderDateToken(ByVal token As String, ByRef recognized As Boolean) As String
    recognized = False

    Dim expression As String
    Dim formatCode As String
    Dim colonAt As Long
    expression = Trim$(token)
    formatCode = "yyyy-mm-dd"

    colonAt = InStrRev(expression, ":")
    If colonAt > 0 Then
        formatCode = Trim$(Mid$(expression, colonAt + 1))
        expression = Trim$(Left$(expression, colonAt - 1))
        If Len(formatCode) = 0 Then
            Err.Raise vbObjectError + 708, , UiText("error.dateformat") & " {{" & token & "}}"
        End If
    End If

    Dim lowered As String
    lowered = LCase$(expression)
    If lowered = "batch_date" Then
        recognized = True
        RenderDateToken = Format$(DateAdd("d", -1, Date), formatCode)
        Exit Function
    End If

    Dim rest As String
    If Left$(lowered, 5) = "today" Then
        recognized = True
        rest = Mid$(expression, 6)
    ElseIf Left$(expression, Len(UiText("token.today"))) = UiText("token.today") Then
        recognized = True
        rest = Mid$(expression, Len(UiText("token.today")) + 1)
    Else
        Exit Function
    End If

    rest = Trim$(rest)
    If Len(rest) > 0 Then
        If Right$(rest, Len(UiText("token.day"))) = UiText("token.day") Or LCase$(Right$(rest, 1)) = "d" Then
            If Right$(rest, Len(UiText("token.day"))) = UiText("token.day") Then
                rest = Left$(rest, Len(rest) - Len(UiText("token.day")))
            Else
                rest = Left$(rest, Len(rest) - 1)
            End If
        End If
    End If

    Dim offsetDays As Long
    offsetDays = 0
    If Len(rest) > 0 Then
        If (Left$(rest, 1) <> "+" And Left$(rest, 1) <> "-") Or Not IsNumeric(rest) Then
            Err.Raise vbObjectError + 709, , UiText("error.date") & " {{" & token & "}}"
        End If
        offsetDays = CLng(rest)
        If Abs(offsetDays) > 36525 Then
            Err.Raise vbObjectError + 710, , UiText("error.offset") & " {{" & token & "}}"
        End If
    End If

    RenderDateToken = Format$(DateAdd("d", offsetDays, Date), formatCode)
End Function

Private Sub AddTokensFromText(ByVal sourceText As String, ByVal ordered As Collection, ByVal seen As Object)
    Dim startAt As Long
    Dim endAt As Long
    Dim token As String
    Dim recognizedDate As Boolean
    Dim ignored As String

    startAt = InStr(1, sourceText, "{{", vbBinaryCompare)
    Do While startAt > 0
        endAt = InStr(startAt + 2, sourceText, "}}", vbBinaryCompare)
        If endAt = 0 Then
            Err.Raise vbObjectError + 711, , UiText("error.unclosed")
        End If

        token = Trim$(Mid$(sourceText, startAt + 2, endAt - startAt - 2))
        If Len(token) = 0 Then
            Err.Raise vbObjectError + 712, , UiText("error.emptytoken")
        End If

        recognizedDate = False
        ignored = RenderDateToken(token, recognizedDate)
        If Not recognizedDate Then
            If Not seen.Exists(token) Then
                seen.Add token, True
                ordered.Add token
            End If
        End If
        startAt = InStr(endAt + 2, sourceText, "{{", vbBinaryCompare)
    Loop
End Sub

Private Function ReadManualValues() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    Dim ws As Worksheet
    Set ws = ToolSheet()

    Dim rowNumber As Long
    Dim variableName As String
    For rowNumber = VARIABLE_FIRST_ROW To VARIABLE_LAST_ROW
        variableName = Trim$(CStr(ws.Cells(rowNumber, 1).Value2))
        If Len(variableName) > 0 Then
            result(variableName) = CStr(ws.Cells(rowNumber, 2).Value2)
        End If
    Next rowNumber

    Set ReadManualValues = result
End Function

Private Function GetTemplateRow() As Long
    Dim wsTool As Worksheet
    Dim wsTemplates As Worksheet
    Set wsTool = ToolSheet()
    Set wsTemplates = TemplateSheet()

    Dim templateName As String
    templateName = Trim$(CStr(wsTool.Range(CELL_TEMPLATE).Value2))
    If Len(templateName) = 0 Then
        Err.Raise vbObjectError + 713, , UiText("error.select")
    End If

    Dim matchResult As Variant
    matchResult = Application.Match(templateName, wsTemplates.Range("A5:A104"), 0)
    If IsError(matchResult) Then
        Err.Raise vbObjectError + 714, , UiText("error.notfound")
    End If

    GetTemplateRow = CLng(matchResult) + 4
End Function

Private Function AutoSendEnabled() As Boolean
    AutoSendEnabled = CBool(SettingsSheet().Range(CELL_AUTO_SEND).Value)
End Function

Private Function GetOutlookApplication() As Object
    On Error Resume Next
    Set GetOutlookApplication = GetObject(, "Outlook.Application")
    If GetOutlookApplication Is Nothing Then
        Set GetOutlookApplication = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
End Function

Private Function BuildConfirmation(ByVal recipients As String, _
                                   ByVal cc As String, _
                                   ByVal subject As String, _
                                   ByVal body As String) As String
    Dim bodyPreview As String
    bodyPreview = body
    If Len(bodyPreview) > 500 Then bodyPreview = Left$(bodyPreview, 500) & vbCrLf & UiText("ellipsis")

    BuildConfirmation = UiText("confirm.intro") & vbCrLf & vbCrLf & _
                        UiText("confirm.to") & recipients & vbCrLf & _
                        UiText("confirm.cc") & IIf(Len(Trim$(cc)) = 0, UiText("confirm.none"), cc) & vbCrLf & _
                        UiText("confirm.subject") & subject & vbCrLf & vbCrLf & _
                        UiText("confirm.body") & vbCrLf & bodyPreview & vbCrLf & vbCrLf & _
                        UiText("confirm.question")
End Function

Private Sub AddButton(ByVal ws As Worksheet, _
                      ByVal controlName As String, _
                      ByVal caption As String, _
                      ByVal anchor As Range, _
                      ByVal buttonWidth As Double, _
                      ByVal macroName As String)
    Dim button As Button
    Set button = ws.Buttons.Add(anchor.Left, anchor.Top, buttonWidth, 24)
    With button
        .Name = controlName
        .Caption = caption
        .OnAction = macroName
        .Font.Name = "Arial"
        .Font.Size = 10
    End With
End Sub

Private Sub DeleteControlIfPresent(ByVal ws As Worksheet, ByVal controlName As String)
    On Error Resume Next
    ws.Shapes(controlName).Delete
    On Error GoTo 0
End Sub

Private Sub SetStatus(ByVal message As String, ByVal isError As Boolean)
    With ToolSheet().Range(CELL_STATUS)
        .Value = message
        If isError Then
            .Interior.Color = RGB(255, 199, 206)
            .Font.Color = RGB(156, 0, 6)
        Else
            .Interior.Color = RGB(226, 240, 217)
            .Font.Color = RGB(31, 78, 42)
        End If
    End With
End Sub

Private Function ToolSheet() As Worksheet
    Set ToolSheet = ThisWorkbook.Worksheets(TOOL_SHEET_INDEX)
End Function

Private Function TemplateSheet() As Worksheet
    Set TemplateSheet = ThisWorkbook.Worksheets(TEMPLATE_SHEET_INDEX)
End Function

Private Function SettingsSheet() As Worksheet
    Set SettingsSheet = ThisWorkbook.Worksheets(SETTINGS_SHEET_INDEX)
End Function

Private Function UiText(ByVal key As String) As String
    Dim matchResult As Variant
    matchResult = Application.Match(key, SettingsSheet().Range("H5:H80"), 0)
    If IsError(matchResult) Then
        UiText = key
    Else
        UiText = CStr(SettingsSheet().Cells(CLng(matchResult) + 4, 9).Value2)
    End If
End Function
