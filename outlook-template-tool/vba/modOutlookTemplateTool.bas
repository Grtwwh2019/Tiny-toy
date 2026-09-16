Attribute VB_Name = "modOutlookTemplateTool"
Option Explicit

' Excel VBA mail template tool for classic Outlook.
' ASCII-only source imports reliably in VBE. Localized text is stored in the workbook.

Private Const TOOL_SHEET_INDEX As Long = 1
Private Const TEMPLATE_SHEET_INDEX As Long = 2
Private Const SCHEDULE_SHEET_INDEX As Long = 3
Private Const SETTINGS_SHEET_INDEX As Long = 4

Private Const CELL_TEMPLATE As String = "B7"
Private Const CELL_AUTO_SEND As String = "B6"
Private Const CELL_SCROLL_OFFSET As String = "B10"
Private Const CELL_PREVIEW_TO As String = "E7"
Private Const CELL_PREVIEW_CC As String = "E8"
Private Const CELL_PREVIEW_SUBJECT As String = "E9"
Private Const CELL_PREVIEW_BODY As String = "D12"
Private Const CELL_STATUS As String = "B32"

Private Const PARAM_FIRST_ROW As Long = 11
Private Const PARAM_LAST_ROW As Long = 16
Private Const PARAM_PAGE_SIZE As Long = 6
Private Const MAX_PARAMETERS As Long = 100

Private Const TASK_FIRST_ROW As Long = 6
Private Const TASK_LAST_ROW As Long = 104
Private Const COL_TASK_ID As Long = 1
Private Const COL_TASK_NAME As Long = 2
Private Const COL_ENABLED As Long = 3
Private Const COL_TEMPLATE As Long = 4
Private Const COL_FREQUENCY As Long = 5
Private Const COL_WEEKDAY As Long = 6
Private Const COL_SEND_TIME As Long = 7
Private Const COL_PARAMETERS As Long = 8
Private Const COL_SIGNATURE As Long = 9
Private Const COL_NEXT_TIME As Long = 10
Private Const COL_LAST_PLANNED As Long = 11
Private Const COL_STATUS As Long = 12
Private Const COL_DETAIL As Long = 13
Private Const COL_OCCURRENCE_KEY As Long = 14
Private Const COL_UPDATED_AT As Long = 15

Private mBusy As Boolean
Private mSchedulerBusy As Boolean
Private mTimerScheduled As Boolean
Private mNextTimer As Date
Private mEditingTaskRow As Long

Private mParamNames As Collection
Private mParamValues As Object
Private mParamComputed As Object
Private mParamOffset As Long
Private mLoadedTemplateName As String

Public Sub Auto_Open()
    On Error Resume Next
    LoadSelectedTemplate
    InitializeScheduler
    On Error GoTo 0
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    CancelScheduledTimer
    On Error GoTo 0
End Sub

Public Sub SetupWorkbook()
    On Error GoTo Failed

    Dim ws As Worksheet
    Dim wsTasks As Worksheet
    Dim cb As CheckBox
    Dim sb As ScrollBar

    Set ws = ToolSheet()
    Set wsTasks = ScheduleSheet()
    Application.ScreenUpdating = False

    DeleteControlIfPresent ws, "btnLoadTemplate"
    DeleteControlIfPresent ws, "btnPreview"
    DeleteControlIfPresent ws, "btnGenerateMail"
    DeleteControlIfPresent ws, "chkAutoSend"
    DeleteControlIfPresent ws, "scrVariables"

    ws.Range("B8").ClearContents
    ws.Range("A18:B18").UnMerge
    ws.Range("A18:B18").ClearContents

    AddButton ws, "btnLoadTemplate", UiText("button.load"), ws.Range("A20"), 118, "LoadSelectedTemplate"
    AddButton ws, "btnPreview", UiText("button.preview"), ws.Range("B20"), 118, "RefreshPreview"
    AddButton ws, "btnGenerateMail", UiText("button.generate"), ws.Range("D20"), 150, "GenerateMail"

    Set cb = ws.CheckBoxes.Add(ws.Range("B8").Left + 4, ws.Range("B8").Top + 3, 190, 18)
    With cb
        .Name = "chkAutoSend"
        .Caption = UiText("checkbox.autosend")
        .LinkedCell = "'" & SettingsSheet().Name & "'!$" & CELL_AUTO_SEND
        .Value = IIf(CBool(SettingsSheet().Range(CELL_AUTO_SEND).Value), 1, -4146)
        .OnAction = "UpdateActionHint"
    End With

    Set sb = ws.ScrollBars.Add(ws.Range("B18").Left, ws.Range("B18").Top + 2, ws.Range("B18").Width, 18)
    With sb
        .Name = "scrVariables"
        .Min = 0
        .Max = 0
        .SmallChange = 1
        .LargeChange = PARAM_PAGE_SIZE
        .LinkedCell = "'" & SettingsSheet().Name & "'!$" & CELL_SCROLL_OFFSET
        .OnAction = "VariableScrollChanged"
    End With

    DeleteControlIfPresent wsTasks, "btnTaskLoad"
    DeleteControlIfPresent wsTasks, "btnTaskSave"
    DeleteControlIfPresent wsTasks, "btnTaskPause"
    DeleteControlIfPresent wsTasks, "btnTaskRefresh"
    DeleteControlIfPresent wsTasks, "btnTaskTest"
    AddButton wsTasks, "btnTaskLoad", UiText("button.task.load"), wsTasks.Range("G2"), 125, "LoadSelectedTaskForEditing"
    AddButton wsTasks, "btnTaskSave", UiText("button.task.save"), wsTasks.Range("H2"), 125, "SaveTaskFromCurrent"
    AddButton wsTasks, "btnTaskPause", UiText("button.task.pause"), wsTasks.Range("I2"), 105, "PauseSelectedTask"
    AddButton wsTasks, "btnTaskRefresh", UiText("button.task.refresh"), wsTasks.Range("J2"), 105, "RefreshScheduler"
    AddButton wsTasks, "btnTaskTest", UiText("button.task.test"), wsTasks.Range("K2"), 145, "TestSelectedTask"

    ws.Rows(20).RowHeight = 28
    LoadSelectedTemplate
    InitializeScheduler
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
    Dim wsTemplates As Worksheet
    Dim oldValues As Object
    Dim ordered As Collection
    Dim seen As Object
    Dim computed As Object
    Dim i As Long
    Dim token As String

    EnsureParameterState
    SaveVisibleValues
    Set oldValues = CloneDictionary(mParamValues)
    rowNumber = GetTemplateRowByName(CStr(ToolSheet().Range(CELL_TEMPLATE).Value2))
    Set wsTemplates = TemplateSheet()
    Set ordered = New Collection
    Set seen = NewDictionary()
    Set computed = NewDictionary()

    CollectTokensFromText CStr(wsTemplates.Cells(rowNumber, 2).Value2), ordered, seen, computed
    CollectTokensFromText CStr(wsTemplates.Cells(rowNumber, 3).Value2), ordered, seen, computed
    CollectTokensFromText CStr(wsTemplates.Cells(rowNumber, 4).Value2), ordered, seen, computed
    CollectTokensFromText CStr(wsTemplates.Cells(rowNumber, 5).Value2), ordered, seen, computed

    If ordered.Count > MAX_PARAMETERS Then Err.Raise vbObjectError + 701, , UiText("error.variables")

    Set mParamNames = ordered
    Set mParamComputed = computed
    Set mParamValues = NewDictionary()
    For i = 1 To mParamNames.Count
        token = CStr(mParamNames(i))
        If Not CBool(mParamComputed(token)) Then
            If oldValues.Exists(token) Then mParamValues(token) = oldValues(token)
        End If
    Next i

    mLoadedTemplateName = Trim$(CStr(ToolSheet().Range(CELL_TEMPLATE).Value2))
    mParamOffset = 0
    SettingsSheet().Range(CELL_SCROLL_OFFSET).Value = 0
    ConfigureVariableScrollbar
    ShowVariablePage
    RefreshPreview
    Exit Sub

Failed:
    SetStatus UiText("error.load") & Err.Description, True
    MsgBox UiText("error.load") & Err.Description, vbExclamation, UiText("app.title")
End Sub

Public Sub VariableScrollChanged()
    On Error GoTo Failed
    EnsureLoadedTemplate
    SaveVisibleValues
    mParamOffset = CLng(Val(SettingsSheet().Range(CELL_SCROLL_OFFSET).Value2))
    If mParamOffset < 0 Then mParamOffset = 0
    If mParamOffset > MaxParameterOffset() Then mParamOffset = MaxParameterOffset()
    ShowVariablePage
    RefreshPreview
    Exit Sub
Failed:
    SetStatus UiText("error.action") & Err.Description, True
End Sub

Public Sub RefreshPreview()
    On Error GoTo Failed

    Dim renderedTo As String
    Dim renderedCc As String
    Dim renderedSubject As String
    Dim renderedBody As String
    Dim missing As String

    EnsureLoadedTemplate
    SaveVisibleValues
    RenderSelectedTemplate renderedTo, renderedCc, renderedSubject, renderedBody, missing, Date

    With ToolSheet()
        .Range(CELL_PREVIEW_TO).Value = renderedTo
        .Range(CELL_PREVIEW_CC).Value = renderedCc
        .Range(CELL_PREVIEW_SUBJECT).Value = renderedSubject
        .Range(CELL_PREVIEW_BODY).Value = renderedBody
    End With

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
    Dim mail As Object

    EnsureLoadedTemplate
    SaveVisibleValues
    RenderSelectedTemplate renderedTo, renderedCc, renderedSubject, renderedBody, missing, Date
    ValidateRenderedMail renderedTo, missing
    Set mail = BuildOutlookMail(renderedTo, renderedCc, renderedSubject, renderedBody)

    If AutoSendEnabled() Then
        If MsgBox(BuildConfirmation(renderedTo, renderedCc, renderedSubject, renderedBody), _
                  vbYesNo + vbExclamation + vbDefaultButton2, UiText("confirm.title")) <> vbYes Then
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

Public Sub LoadSelectedTaskForEditing()
    On Error GoTo Failed

    Dim taskRow As Long
    Dim templateName As String
    Dim storedValues As Object
    Dim key As Variant

    taskRow = SelectedTaskRow()
    templateName = Trim$(CStr(ScheduleSheet().Cells(taskRow, COL_TEMPLATE).Value2))
    If Len(templateName) = 0 Then Err.Raise vbObjectError + 721, , UiText("error.select")
    GetTemplateRowByName templateName

    ToolSheet().Range(CELL_TEMPLATE).Value = templateName
    LoadSelectedTemplate
    Set storedValues = DecodeParameters(CStr(ScheduleSheet().Cells(taskRow, COL_PARAMETERS).Value2))
    For Each key In storedValues.Keys
        If mParamValues.Exists(CStr(key)) Then mParamValues(CStr(key)) = storedValues(key)
    Next key
    mParamOffset = 0
    SettingsSheet().Range(CELL_SCROLL_OFFSET).Value = 0
    ShowVariablePage
    RefreshPreview
    mEditingTaskRow = taskRow
    ToolSheet().Activate
    SetStatus UiText("status.task.loaded"), False
    Exit Sub

Failed:
    MsgBox UiText("error.action") & Err.Description, vbExclamation, UiText("app.title")
End Sub

Public Sub SaveTaskFromCurrent()
    On Error GoTo Failed

    Dim taskRow As Long
    Dim ws As Worksheet
    Dim taskName As String
    Dim templateName As String
    Dim frequency As Long
    Dim weekdayNumber As Long
    Dim sendTime As Date
    Dim renderedTo As String
    Dim renderedCc As String
    Dim renderedSubject As String
    Dim renderedBody As String
    Dim missing As String
    Dim prompt As String

    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 722, , UiText("error.task.readonly")
    If mEditingTaskRow >= TASK_FIRST_ROW Then
        taskRow = mEditingTaskRow
    Else
        taskRow = SelectedTaskRow()
    End If
    Set ws = ScheduleSheet()
    taskName = Trim$(CStr(ws.Cells(taskRow, COL_TASK_NAME).Value2))
    If Len(taskName) = 0 Then Err.Raise vbObjectError + 723, , UiText("error.task.name")

    EnsureLoadedTemplate
    SaveVisibleValues
    templateName = Trim$(CStr(ToolSheet().Range(CELL_TEMPLATE).Value2))
    frequency = ValidatedTaskFrequency(taskRow)
    weekdayNumber = ValidatedTaskWeekday(taskRow, frequency)
    sendTime = ValidatedTaskTime(taskRow)
    RenderSelectedTemplate renderedTo, renderedCc, renderedSubject, renderedBody, missing, Date
    ValidateRenderedMail renderedTo, missing

    prompt = BuildTaskConfirmation(taskName, templateName, frequency, weekdayNumber, sendTime, renderedTo, renderedCc, renderedSubject)
    If MsgBox(prompt, vbYesNo + vbExclamation + vbDefaultButton2, UiText("confirm.task.title")) <> vbYes Then Exit Sub

    CancelScheduledTimer
    If Len(Trim$(CStr(ws.Cells(taskRow, COL_TASK_ID).Value2))) = 0 Then
        ws.Cells(taskRow, COL_TASK_ID).Value = Format$(Now, "yyyymmddhhnnss") & "-" & CStr(taskRow)
    End If
    ws.Cells(taskRow, COL_ENABLED).Value = "YES"
    ws.Cells(taskRow, COL_TEMPLATE).Value = templateName
    ws.Cells(taskRow, COL_PARAMETERS).Value = EncodeParameters(mParamValues)
    ws.Cells(taskRow, COL_SIGNATURE).Value = TemplateSignature(templateName)
    ws.Cells(taskRow, COL_NEXT_TIME).Value = NextOccurrence(taskRow, Now)
    ws.Cells(taskRow, COL_STATUS).Value = UiText("task.status.ready")
    ws.Cells(taskRow, COL_DETAIL).ClearContents
    ws.Cells(taskRow, COL_OCCURRENCE_KEY).ClearContents
    ws.Cells(taskRow, COL_UPDATED_AT).Value = Now
    SaveWorkbookSafely
    mEditingTaskRow = 0
    InitializeScheduler
    ws.Activate
    ws.Cells(taskRow, COL_TASK_NAME).Select
    SetStatus UiText("status.task.saved"), False
    Exit Sub

Failed:
    MsgBox UiText("error.action") & Err.Description, vbCritical, UiText("app.title")
End Sub

Public Sub PauseSelectedTask()
    On Error GoTo Failed
    Dim taskRow As Long
    taskRow = SelectedTaskRow()
    CancelScheduledTimer
    With ScheduleSheet()
        .Cells(taskRow, COL_ENABLED).Value = "NO"
        .Cells(taskRow, COL_NEXT_TIME).ClearContents
        .Cells(taskRow, COL_STATUS).Value = UiText("task.status.paused")
        .Cells(taskRow, COL_UPDATED_AT).Value = Now
    End With
    SaveWorkbookSafely
    InitializeScheduler
    Exit Sub
Failed:
    MsgBox UiText("error.action") & Err.Description, vbExclamation, UiText("app.title")
End Sub

Public Sub RefreshScheduler()
    On Error GoTo Failed
    InitializeScheduler
    Exit Sub
Failed:
    MsgBox UiText("error.action") & Err.Description, vbExclamation, UiText("app.title")
End Sub

Public Sub TestSelectedTask()
    On Error GoTo Failed

    Dim taskRow As Long
    Dim renderedTo As String
    Dim renderedCc As String
    Dim renderedSubject As String
    Dim renderedBody As String
    Dim missing As String
    Dim mail As Object

    taskRow = SelectedTaskRow()
    RenderTask taskRow, Date, renderedTo, renderedCc, renderedSubject, renderedBody, missing
    ValidateRenderedMail renderedTo, missing
    Set mail = BuildOutlookMail(renderedTo, renderedCc, renderedSubject, renderedBody)
    mail.Display
    Exit Sub
Failed:
    MsgBox UiText("error.action") & Err.Description, vbExclamation, UiText("app.title")
End Sub

Public Sub SchedulerTick()
    If mSchedulerBusy Then Exit Sub
    mSchedulerBusy = True
    mTimerScheduled = False
    On Error GoTo Failed

    Dim tickTime As Date
    Dim taskRow As Long
    Dim plannedTime As Date
    Dim changed As Boolean
    tickTime = Now

    For taskRow = TASK_FIRST_ROW To TASK_LAST_ROW
        If TaskEnabled(taskRow) And IsDate(ScheduleSheet().Cells(taskRow, COL_NEXT_TIME).Value) Then
            plannedTime = CDate(ScheduleSheet().Cells(taskRow, COL_NEXT_TIME).Value)
            If plannedTime <= tickTime Then
                If DateDiff("s", plannedTime, tickTime) = 0 Then
                    SendScheduledTask taskRow, plannedTime
                Else
                    MarkTaskMissed taskRow, plannedTime
                End If
                ScheduleSheet().Cells(taskRow, COL_NEXT_TIME).Value = NextOccurrence(taskRow, tickTime)
                ScheduleSheet().Cells(taskRow, COL_UPDATED_AT).Value = Now
                changed = True
            End If
        End If
    Next taskRow

    If changed Then SaveWorkbookSafely
    InitializeScheduler
CleanExit:
    mSchedulerBusy = False
    Exit Sub
Failed:
    On Error Resume Next
    InitializeScheduler
    On Error GoTo 0
    Resume CleanExit
End Sub

Public Sub InitializeScheduler()
    On Error GoTo Failed
    If mSchedulerBusy Then Exit Sub
    CancelScheduledTimer
    If ThisWorkbook.ReadOnly Then
        SetStatus UiText("error.task.readonly"), True
        Exit Sub
    End If

    Dim taskRow As Long
    Dim nextTime As Date
    Dim earliest As Date
    Dim hasEarliest As Boolean
    Dim templateName As String
    Dim currentSignature As String
    Dim plannedTime As Date
    Dim changed As Boolean

    For taskRow = TASK_FIRST_ROW To TASK_LAST_ROW
        If TaskEnabled(taskRow) Then
            templateName = Trim$(CStr(ScheduleSheet().Cells(taskRow, COL_TEMPLATE).Value2))
            currentSignature = SafeTemplateSignature(templateName)
            If Len(currentSignature) = 0 Or _
               CStr(ScheduleSheet().Cells(taskRow, COL_SIGNATURE).Value2) <> currentSignature Then
                DisableChangedTask taskRow
                changed = True
            Else
                If IsDate(ScheduleSheet().Cells(taskRow, COL_NEXT_TIME).Value) Then
                    plannedTime = CDate(ScheduleSheet().Cells(taskRow, COL_NEXT_TIME).Value)
                    If DateDiff("s", plannedTime, Now) > 0 Then
                        MarkTaskMissed taskRow, plannedTime
                        ScheduleSheet().Cells(taskRow, COL_NEXT_TIME).Value = NextOccurrence(taskRow, Now)
                        changed = True
                    End If
                Else
                    ScheduleSheet().Cells(taskRow, COL_NEXT_TIME).Value = NextOccurrence(taskRow, Now)
                    ScheduleSheet().Cells(taskRow, COL_STATUS).Value = UiText("task.status.ready")
                    changed = True
                End If

                nextTime = CDate(ScheduleSheet().Cells(taskRow, COL_NEXT_TIME).Value)
                If Not hasEarliest Or nextTime < earliest Then
                    earliest = nextTime
                    hasEarliest = True
                End If
            End If
        End If
    Next taskRow

    If changed Then SaveWorkbookSafely
    If hasEarliest Then RegisterScheduledTimer earliest
    Exit Sub
Failed:
    SetStatus UiText("error.action") & Err.Description, True
End Sub

Private Function SafeTemplateSignature(ByVal templateName As String) As String
    On Error GoTo NotAvailable
    If Len(Trim$(templateName)) = 0 Then Exit Function
    SafeTemplateSignature = TemplateSignature(templateName)
    Exit Function
NotAvailable:
    SafeTemplateSignature = vbNullString
End Function

Private Sub SendScheduledTask(ByVal taskRow As Long, ByVal plannedTime As Date)
    Dim ws As Worksheet
    Dim occurrenceKey As String
    Dim renderedTo As String
    Dim renderedCc As String
    Dim renderedSubject As String
    Dim renderedBody As String
    Dim missing As String
    Dim mail As Object
    Dim sendInvoked As Boolean

    Set ws = ScheduleSheet()
    occurrenceKey = Format$(plannedTime, "yyyymmddhhnnss")
    If CStr(ws.Cells(taskRow, COL_OCCURRENCE_KEY).Value2) = occurrenceKey Then
        ws.Cells(taskRow, COL_DETAIL).Value = UiText("task.detail.duplicate")
        Exit Sub
    End If

    On Error GoTo Failed
    RenderTask taskRow, DateValue(plannedTime), renderedTo, renderedCc, renderedSubject, renderedBody, missing
    ValidateRenderedMail renderedTo, missing
    Set mail = BuildOutlookMail(renderedTo, renderedCc, renderedSubject, renderedBody)

    ws.Cells(taskRow, COL_LAST_PLANNED).Value = plannedTime
    ws.Cells(taskRow, COL_STATUS).Value = UiText("task.status.sending")
    ws.Cells(taskRow, COL_DETAIL).ClearContents
    ws.Cells(taskRow, COL_OCCURRENCE_KEY).Value = occurrenceKey
    ws.Cells(taskRow, COL_UPDATED_AT).Value = Now
    SaveWorkbookSafely

    sendInvoked = True
    mail.Send
    ws.Cells(taskRow, COL_STATUS).Value = UiText("task.status.sent")
    ws.Cells(taskRow, COL_DETAIL).Value = UiText("task.detail.sent")
    ws.Cells(taskRow, COL_UPDATED_AT).Value = Now
    SaveWorkbookSafely
    Exit Sub

Failed:
    If sendInvoked Then
        ws.Cells(taskRow, COL_STATUS).Value = UiText("task.status.uncertain")
    Else
        ws.Cells(taskRow, COL_STATUS).Value = UiText("task.status.failed")
    End If
    ws.Cells(taskRow, COL_DETAIL).Value = Err.Description
    ws.Cells(taskRow, COL_UPDATED_AT).Value = Now
    On Error Resume Next
    SaveWorkbookSafely
    On Error GoTo 0
End Sub

Private Sub MarkTaskMissed(ByVal taskRow As Long, ByVal plannedTime As Date)
    With ScheduleSheet()
        .Cells(taskRow, COL_LAST_PLANNED).Value = plannedTime
        .Cells(taskRow, COL_STATUS).Value = UiText("task.status.missed")
        .Cells(taskRow, COL_DETAIL).Value = UiText("task.detail.missed")
        .Cells(taskRow, COL_OCCURRENCE_KEY).Value = Format$(plannedTime, "yyyymmddhhnnss")
        .Cells(taskRow, COL_UPDATED_AT).Value = Now
    End With
End Sub

Private Sub DisableChangedTask(ByVal taskRow As Long)
    With ScheduleSheet()
        .Cells(taskRow, COL_ENABLED).Value = "NO"
        .Cells(taskRow, COL_NEXT_TIME).ClearContents
        .Cells(taskRow, COL_STATUS).Value = UiText("task.status.changed")
        .Cells(taskRow, COL_UPDATED_AT).Value = Now
    End With
End Sub

Private Function NextOccurrence(ByVal taskRow As Long, ByVal afterTime As Date) As Date
    Dim frequency As Long
    Dim weekdayNumber As Long
    Dim sendTime As Date
    Dim candidate As Date
    Dim daysToAdd As Long

    frequency = ValidatedTaskFrequency(taskRow)
    weekdayNumber = ValidatedTaskWeekday(taskRow, frequency)
    sendTime = ValidatedTaskTime(taskRow)
    candidate = DateValue(afterTime) + TimeValue(sendTime)

    If frequency = 1 Then
        If candidate <= afterTime Then candidate = DateAdd("d", 1, candidate)
    Else
        daysToAdd = (weekdayNumber - Weekday(DateValue(afterTime), vbMonday) + 7) Mod 7
        candidate = DateAdd("d", daysToAdd, candidate)
        If candidate <= afterTime Then candidate = DateAdd("d", 7, candidate)
    End If
    NextOccurrence = candidate
End Function

Private Function ValidatedTaskFrequency(ByVal taskRow As Long) As Long
    Dim value As Long
    value = CLng(Val(ScheduleSheet().Cells(taskRow, COL_FREQUENCY).Value2))
    If value <> 1 And value <> 2 Then Err.Raise vbObjectError + 724, , UiText("error.task.frequency")
    ValidatedTaskFrequency = value
End Function

Private Function ValidatedTaskWeekday(ByVal taskRow As Long, ByVal frequency As Long) As Long
    Dim value As Long
    If frequency = 1 Then
        ValidatedTaskWeekday = 1
        Exit Function
    End If
    value = CLng(Val(ScheduleSheet().Cells(taskRow, COL_WEEKDAY).Value2))
    If value < 1 Or value > 7 Then Err.Raise vbObjectError + 725, , UiText("error.task.weekday")
    ValidatedTaskWeekday = value
End Function

Private Function ValidatedTaskTime(ByVal taskRow As Long) As Date
    Dim value As Variant
    value = ScheduleSheet().Cells(taskRow, COL_SEND_TIME).Value
    If Not IsDate(value) Then Err.Raise vbObjectError + 726, , UiText("error.task.time")
    ValidatedTaskTime = TimeValue(CDate(value))
End Function

Private Function SelectedTaskRow() As Long
    If ActiveSheet.Index <> SCHEDULE_SHEET_INDEX Or ActiveCell.Row < TASK_FIRST_ROW Or ActiveCell.Row > TASK_LAST_ROW Then
        Err.Raise vbObjectError + 727, , UiText("error.task.row")
    End If
    SelectedTaskRow = ActiveCell.Row
End Function

Private Function TaskEnabled(ByVal taskRow As Long) As Boolean
    TaskEnabled = (UCase$(Trim$(CStr(ScheduleSheet().Cells(taskRow, COL_ENABLED).Value2))) = "YES")
End Function

Private Sub RegisterScheduledTimer(ByVal runTime As Date)
    mNextTimer = runTime
    Application.OnTime EarliestTime:=mNextTimer, Procedure:=QualifiedMacroName("SchedulerTick"), Schedule:=True
    mTimerScheduled = True
End Sub

Private Sub CancelScheduledTimer()
    If Not mTimerScheduled Then Exit Sub
    On Error Resume Next
    Application.OnTime EarliestTime:=mNextTimer, Procedure:=QualifiedMacroName("SchedulerTick"), Schedule:=False
    On Error GoTo 0
    mTimerScheduled = False
End Sub

Private Function QualifiedMacroName(ByVal procedureName As String) As String
    QualifiedMacroName = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & procedureName
End Function

Private Sub EnsureParameterState()
    If mParamNames Is Nothing Then Set mParamNames = New Collection
    If mParamValues Is Nothing Then Set mParamValues = NewDictionary()
    If mParamComputed Is Nothing Then Set mParamComputed = NewDictionary()
End Sub

Private Sub EnsureLoadedTemplate()
    EnsureParameterState
    If mLoadedTemplateName <> Trim$(CStr(ToolSheet().Range(CELL_TEMPLATE).Value2)) Then LoadSelectedTemplate
End Sub

Private Sub SaveVisibleValues()
    EnsureParameterState
    Dim visibleRow As Long
    Dim itemIndex As Long
    Dim token As String
    For visibleRow = PARAM_FIRST_ROW To PARAM_LAST_ROW
        itemIndex = mParamOffset + visibleRow - PARAM_FIRST_ROW + 1
        If itemIndex <= mParamNames.Count Then
            token = CStr(mParamNames(itemIndex))
            If Not CBool(mParamComputed(token)) Then
                mParamValues(token) = CStr(ToolSheet().Cells(visibleRow, 2).Value2)
            End If
        End If
    Next visibleRow
End Sub

Private Sub ShowVariablePage()
    EnsureParameterState
    Dim ws As Worksheet
    Dim visibleRow As Long
    Dim itemIndex As Long
    Dim token As String
    Dim recognized As Boolean
    Dim firstShown As Long
    Dim lastShown As Long

    Set ws = ToolSheet()
    ws.Range("A" & PARAM_FIRST_ROW & ":B" & PARAM_LAST_ROW).ClearContents
    ws.Range("A" & PARAM_FIRST_ROW & ":B" & PARAM_LAST_ROW).Interior.Color = RGB(255, 255, 255)
    ws.Range("B" & PARAM_FIRST_ROW & ":B" & PARAM_LAST_ROW).Interior.Color = RGB(255, 242, 204)

    If mParamNames.Count = 0 Then
        ws.Cells(PARAM_FIRST_ROW, 1).Value = UiText("status.no.parameters")
        ws.Range("A18").Value = "0 / 0"
        Exit Sub
    End If

    For visibleRow = PARAM_FIRST_ROW To PARAM_LAST_ROW
        itemIndex = mParamOffset + visibleRow - PARAM_FIRST_ROW + 1
        If itemIndex <= mParamNames.Count Then
            token = CStr(mParamNames(itemIndex))
            ws.Cells(visibleRow, 1).Value = token
            If CBool(mParamComputed(token)) Then
                recognized = False
                ws.Cells(visibleRow, 2).Value = RenderDateToken(token, recognized, Date)
                ws.Range("A" & visibleRow & ":B" & visibleRow).Interior.Color = RGB(231, 230, 230)
            ElseIf mParamValues.Exists(token) Then
                ws.Cells(visibleRow, 2).Value = mParamValues(token)
            End If
        End If
    Next visibleRow

    firstShown = mParamOffset + 1
    lastShown = mParamOffset + PARAM_PAGE_SIZE
    If lastShown > mParamNames.Count Then lastShown = mParamNames.Count
    ws.Range("A18").Value = UiText("scroll.position") & " " & firstShown & "-" & lastShown & " / " & mParamNames.Count
End Sub

Private Sub ConfigureVariableScrollbar()
    On Error Resume Next
    With ToolSheet().ScrollBars("scrVariables")
        .Min = 0
        .Max = MaxParameterOffset()
        .SmallChange = 1
        .LargeChange = PARAM_PAGE_SIZE
        .Value = mParamOffset
        .Enabled = (mParamNames.Count > PARAM_PAGE_SIZE)
    End With
    On Error GoTo 0
End Sub

Private Function MaxParameterOffset() As Long
    If mParamNames Is Nothing Then
        MaxParameterOffset = 0
    ElseIf mParamNames.Count > PARAM_PAGE_SIZE Then
        MaxParameterOffset = mParamNames.Count - PARAM_PAGE_SIZE
    Else
        MaxParameterOffset = 0
    End If
End Function

Private Sub RenderSelectedTemplate(ByRef renderedTo As String, ByRef renderedCc As String, _
                                   ByRef renderedSubject As String, ByRef renderedBody As String, _
                                   ByRef missingText As String, ByVal baseDate As Date)
    Dim templateName As String
    templateName = Trim$(CStr(ToolSheet().Range(CELL_TEMPLATE).Value2))
    RenderTemplateByName templateName, mParamValues, baseDate, renderedTo, renderedCc, renderedSubject, renderedBody, missingText
End Sub

Private Sub RenderTask(ByVal taskRow As Long, ByVal baseDate As Date, _
                       ByRef renderedTo As String, ByRef renderedCc As String, _
                       ByRef renderedSubject As String, ByRef renderedBody As String, _
                       ByRef missingText As String)
    Dim templateName As String
    Dim values As Object
    templateName = Trim$(CStr(ScheduleSheet().Cells(taskRow, COL_TEMPLATE).Value2))
    Set values = DecodeParameters(CStr(ScheduleSheet().Cells(taskRow, COL_PARAMETERS).Value2))
    RenderTemplateByName templateName, values, baseDate, renderedTo, renderedCc, renderedSubject, renderedBody, missingText
End Sub

Private Sub RenderTemplateByName(ByVal templateName As String, ByVal values As Object, ByVal baseDate As Date, _
                                 ByRef renderedTo As String, ByRef renderedCc As String, _
                                 ByRef renderedSubject As String, ByRef renderedBody As String, _
                                 ByRef missingText As String)
    Dim rowNumber As Long
    Dim ws As Worksheet
    Dim missing As Object
    Dim key As Variant
    rowNumber = GetTemplateRowByName(templateName)
    Set ws = TemplateSheet()
    Set missing = NewDictionary()
    renderedTo = RenderText(CStr(ws.Cells(rowNumber, 2).Value2), values, missing, baseDate)
    renderedCc = RenderText(CStr(ws.Cells(rowNumber, 3).Value2), values, missing, baseDate)
    renderedSubject = RenderText(CStr(ws.Cells(rowNumber, 4).Value2), values, missing, baseDate)
    renderedBody = RenderText(CStr(ws.Cells(rowNumber, 5).Value2), values, missing, baseDate)
    For Each key In missing.Keys
        If Len(missingText) > 0 Then missingText = missingText & "; "
        missingText = missingText & CStr(key)
    Next key
End Sub

Private Function RenderText(ByVal sourceText As String, ByVal values As Object, ByVal missing As Object, ByVal baseDate As Date) As String
    Dim result As String
    Dim startAt As Long
    Dim endAt As Long
    Dim token As String
    Dim replacement As String
    Dim recognizedDate As Boolean
    result = sourceText
    startAt = InStr(1, result, "{{", vbBinaryCompare)
    Do While startAt > 0
        endAt = InStr(startAt + 2, result, "}}", vbBinaryCompare)
        If endAt = 0 Then Err.Raise vbObjectError + 706, , UiText("error.unclosed")
        token = Trim$(Mid$(result, startAt + 2, endAt - startAt - 2))
        If Len(token) = 0 Then Err.Raise vbObjectError + 707, , UiText("error.emptytoken")
        recognizedDate = False
        replacement = RenderDateToken(token, recognizedDate, baseDate)
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

Private Function RenderDateToken(ByVal token As String, ByRef recognized As Boolean, ByVal baseDate As Date) As String
    recognized = False
    Dim expression As String
    Dim formatCode As String
    Dim colonAt As Long
    Dim lowered As String
    Dim rest As String
    Dim offsetDays As Long
    expression = Trim$(token)
    formatCode = "yyyy-mm-dd"
    colonAt = InStrRev(expression, ":")
    If colonAt > 0 Then
        formatCode = Trim$(Mid$(expression, colonAt + 1))
        expression = Trim$(Left$(expression, colonAt - 1))
        If Len(formatCode) = 0 Then Err.Raise vbObjectError + 708, , UiText("error.dateformat") & " {{" & token & "}}"
    End If
    lowered = LCase$(expression)
    If lowered = "batch_date" Then
        recognized = True
        RenderDateToken = Format$(DateAdd("d", -1, DateValue(baseDate)), formatCode)
        Exit Function
    End If
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
        If Right$(rest, Len(UiText("token.day"))) = UiText("token.day") Then
            rest = Left$(rest, Len(rest) - Len(UiText("token.day")))
        ElseIf LCase$(Right$(rest, 1)) = "d" Then
            rest = Left$(rest, Len(rest) - 1)
        End If
    End If
    offsetDays = 0
    If Len(rest) > 0 Then
        If (Left$(rest, 1) <> "+" And Left$(rest, 1) <> "-") Or Not IsNumeric(rest) Then
            Err.Raise vbObjectError + 709, , UiText("error.date") & " {{" & token & "}}"
        End If
        offsetDays = CLng(rest)
        If Abs(offsetDays) > 36525 Then Err.Raise vbObjectError + 710, , UiText("error.offset") & " {{" & token & "}}"
    End If
    RenderDateToken = Format$(DateAdd("d", offsetDays, DateValue(baseDate)), formatCode)
End Function

Private Sub CollectTokensFromText(ByVal sourceText As String, ByVal ordered As Collection, ByVal seen As Object, ByVal computed As Object)
    Dim startAt As Long
    Dim endAt As Long
    Dim token As String
    Dim recognizedDate As Boolean
    Dim ignored As String
    startAt = InStr(1, sourceText, "{{", vbBinaryCompare)
    Do While startAt > 0
        endAt = InStr(startAt + 2, sourceText, "}}", vbBinaryCompare)
        If endAt = 0 Then Err.Raise vbObjectError + 711, , UiText("error.unclosed")
        token = Trim$(Mid$(sourceText, startAt + 2, endAt - startAt - 2))
        If Len(token) = 0 Then Err.Raise vbObjectError + 712, , UiText("error.emptytoken")
        If Not seen.Exists(token) Then
            recognizedDate = False
            ignored = RenderDateToken(token, recognizedDate, Date)
            seen.Add token, True
            computed.Add token, recognizedDate
            ordered.Add token
        End If
        startAt = InStr(endAt + 2, sourceText, "{{", vbBinaryCompare)
    Loop
End Sub

Private Function TemplateSignature(ByVal templateName As String) As String
    Dim rowNumber As Long
    Dim source As String
    Dim ws As Worksheet
    rowNumber = GetTemplateRowByName(templateName)
    Set ws = TemplateSheet()
    source = templateName & Chr$(30) & CStr(ws.Cells(rowNumber, 2).Value2) & Chr$(30) & _
             CStr(ws.Cells(rowNumber, 3).Value2) & Chr$(30) & CStr(ws.Cells(rowNumber, 4).Value2) & Chr$(30) & _
             CStr(ws.Cells(rowNumber, 5).Value2)
    TemplateSignature = AdlerChecksum(source)
End Function

Private Function AdlerChecksum(ByVal value As String) As String
    Dim a As Double
    Dim b As Double
    Dim i As Long
    Dim codePoint As Long
    a = 1
    b = 0
    For i = 1 To Len(value)
        codePoint = AscW(Mid$(value, i, 1))
        If codePoint < 0 Then codePoint = codePoint + 65536
        a = (a + codePoint) Mod 65521
        b = (b + a) Mod 65521
    Next i
    AdlerChecksum = Format$(a, "0") & "-" & Format$(b, "0")
End Function

Private Function EncodeParameters(ByVal values As Object) As String
    Dim key As Variant
    Dim result As String
    For Each key In values.Keys
        If Len(result) > 0 Then result = result & "|"
        result = result & EscapeParameter(CStr(key)) & "=" & EscapeParameter(CStr(values(key)))
    Next key
    EncodeParameters = result
End Function

Private Function DecodeParameters(ByVal encoded As String) As Object
    Dim result As Object
    Dim parts As Variant
    Dim item As Variant
    Dim equalsAt As Long
    Dim key As String
    Set result = NewDictionary()
    If Len(encoded) > 0 Then
        parts = Split(encoded, "|")
        For Each item In parts
            equalsAt = InStr(1, CStr(item), "=", vbBinaryCompare)
            If equalsAt > 0 Then
                key = UnescapeParameter(Left$(CStr(item), equalsAt - 1))
                result(key) = UnescapeParameter(Mid$(CStr(item), equalsAt + 1))
            End If
        Next item
    End If
    Set DecodeParameters = result
End Function

Private Function EscapeParameter(ByVal value As String) As String
    value = Replace(value, "%", "%25")
    value = Replace(value, "|", "%7C")
    value = Replace(value, "=", "%3D")
    value = Replace(value, vbCr, "%0D")
    value = Replace(value, vbLf, "%0A")
    EscapeParameter = value
End Function

Private Function UnescapeParameter(ByVal value As String) As String
    value = Replace(value, "%0A", vbLf)
    value = Replace(value, "%0D", vbCr)
    value = Replace(value, "%3D", "=")
    value = Replace(value, "%7C", "|")
    value = Replace(value, "%25", "%")
    UnescapeParameter = value
End Function

Private Function CloneDictionary(ByVal source As Object) As Object
    Dim result As Object
    Dim key As Variant
    Set result = NewDictionary()
    For Each key In source.Keys
        result(key) = source(key)
    Next key
    Set CloneDictionary = result
End Function

Private Function NewDictionary() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1
    Set NewDictionary = result
End Function

Private Function GetTemplateRowByName(ByVal templateName As String) As Long
    Dim matchResult As Variant
    templateName = Trim$(templateName)
    If Len(templateName) = 0 Then Err.Raise vbObjectError + 713, , UiText("error.select")
    matchResult = Application.Match(templateName, TemplateSheet().Range("A5:A104"), 0)
    If IsError(matchResult) Then Err.Raise vbObjectError + 714, , UiText("error.notfound")
    GetTemplateRowByName = CLng(matchResult) + 4
End Function

Private Function AutoSendEnabled() As Boolean
    AutoSendEnabled = CBool(SettingsSheet().Range(CELL_AUTO_SEND).Value)
End Function

Private Function BuildOutlookMail(ByVal recipients As String, ByVal cc As String, ByVal subject As String, ByVal body As String) As Object
    Dim outlookApp As Object
    Dim mail As Object
    Set outlookApp = GetOutlookApplication()
    If outlookApp Is Nothing Then Err.Raise vbObjectError + 704, , UiText("error.outlook")
    Set mail = outlookApp.CreateItem(0)
    With mail
        .To = recipients
        .CC = cc
        .Subject = subject
        .Body = body
    End With
    If Not mail.Recipients.ResolveAll Then Err.Raise vbObjectError + 705, , UiText("error.resolve")
    Set BuildOutlookMail = mail
End Function

Private Function GetOutlookApplication() As Object
    On Error Resume Next
    Set GetOutlookApplication = GetObject(, "Outlook.Application")
    If GetOutlookApplication Is Nothing Then Set GetOutlookApplication = CreateObject("Outlook.Application")
    On Error GoTo 0
End Function

Private Sub ValidateRenderedMail(ByVal renderedTo As String, ByVal missing As String)
    If Len(missing) > 0 Then Err.Raise vbObjectError + 702, , UiText("status.missing") & missing
    If Len(Trim$(renderedTo)) = 0 Then Err.Raise vbObjectError + 703, , UiText("error.to")
End Sub

Private Function BuildConfirmation(ByVal recipients As String, ByVal cc As String, ByVal subject As String, ByVal body As String) As String
    Dim bodyPreview As String
    bodyPreview = body
    If Len(bodyPreview) > 500 Then bodyPreview = Left$(bodyPreview, 500) & vbCrLf & UiText("ellipsis")
    BuildConfirmation = UiText("confirm.intro") & vbCrLf & vbCrLf & UiText("confirm.to") & recipients & vbCrLf & _
                        UiText("confirm.cc") & IIf(Len(Trim$(cc)) = 0, UiText("confirm.none"), cc) & vbCrLf & _
                        UiText("confirm.subject") & subject & vbCrLf & vbCrLf & UiText("confirm.body") & vbCrLf & _
                        bodyPreview & vbCrLf & vbCrLf & UiText("confirm.question")
End Function

Private Function BuildTaskConfirmation(ByVal taskName As String, ByVal templateName As String, ByVal frequency As Long, _
                                       ByVal weekdayNumber As Long, ByVal sendTime As Date, ByVal recipients As String, _
                                       ByVal cc As String, ByVal subject As String) As String
    Dim frequencyText As String
    If frequency = 1 Then
        frequencyText = UiText("confirm.task.frequency.daily")
    Else
        frequencyText = UiText("confirm.task.frequency.weekly") & " " & CStr(weekdayNumber)
    End If
    BuildTaskConfirmation = UiText("confirm.task.intro") & vbCrLf & vbCrLf & _
                            UiText("confirm.task.name") & taskName & vbCrLf & _
                            UiText("confirm.task.template") & templateName & vbCrLf & _
                            frequencyText & vbCrLf & UiText("confirm.task.time") & Format$(sendTime, "hh:nn:ss") & vbCrLf & _
                            UiText("confirm.to") & recipients & vbCrLf & _
                            UiText("confirm.cc") & IIf(Len(Trim$(cc)) = 0, UiText("confirm.none"), cc) & vbCrLf & _
                            UiText("confirm.subject") & subject
End Function

Private Sub SaveWorkbookSafely()
    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 728, , UiText("error.task.readonly")
    ThisWorkbook.Save
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal controlName As String, ByVal caption As String, _
                      ByVal anchor As Range, ByVal buttonWidth As Double, ByVal macroName As String)
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
Private Function ScheduleSheet() As Worksheet
    Set ScheduleSheet = ThisWorkbook.Worksheets(SCHEDULE_SHEET_INDEX)
End Function
Private Function SettingsSheet() As Worksheet
    Set SettingsSheet = ThisWorkbook.Worksheets(SETTINGS_SHEET_INDEX)
End Function

Private Function UiText(ByVal key As String) As String
    Dim matchResult As Variant
    matchResult = Application.Match(key, SettingsSheet().Range("H5:H120"), 0)
    If IsError(matchResult) Then
        UiText = key
    Else
        UiText = CStr(SettingsSheet().Cells(CLng(matchResult) + 4, 9).Value2)
    End If
End Function
