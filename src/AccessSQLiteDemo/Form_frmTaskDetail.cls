Attribute VB_Name = "Form_frmTaskDetail"
Attribute VB_Base = "0{79F9AFDB-B980-47E3-B95C-38A48AF3E740}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

' =============================================================================
' Form_frmTaskDetail
' Unbound Single-Page Form bound directly to SQLite via clsTaskItem
' (Implements AccessSQLite.ISQLiteEntity) and AccessSQLite.clsSQLiteRepository.
' =============================================================================

Private m_CurrentTask As clsTaskItem

Private Sub Form_Load()
    On Error GoTo LoadErr
    Dim strArg As String
    Dim lngFirstId As Long

    strArg = Trim$(CStr(Nz(Me.OpenArgs, vbNullString)))
    If StrComp(strArg, "NEW", vbTextCompare) = 0 Then
        BeginNewTask
    ElseIf IsNumeric(strArg) And CLng(Val(strArg)) > 0 Then
        LoadTaskFromSQLite CLng(Val(strArg))
    Else
        lngFirstId = FindTargetSqliteTaskId("FIRST")
        If lngFirstId > 0 Then
            LoadTaskFromSQLite lngFirstId
        Else
            BeginNewTask
        End If
    End If
    Exit Sub
LoadErr:
End Sub

Public Sub BeginNewTask()
    Set m_CurrentTask = New clsTaskItem
    Me!txtTaskID.Value = "(New)"
    Me!txtTaskTitle.Value = vbNullString
    Me!cboCategoryName.Value = "General"
    Me!cboPriorityLevel.Value = "Medium"
    Me!txtEstimateCost.Value = 0@
    Me!txtDueDate.Value = Null
    Me!chkIsCompleted.Value = False
    Me!txtTaskNotes.Value = vbNullString
    Me!txtAuditStamp.Value = "New unsaved record (will insert directly into AccessSQLiteDemo.db)"
    Me!cmdDeleteTask.Enabled = False
    Me!cmdDeleteTask.Caption = "Soft Delete"
    On Error Resume Next
    Me!txtTaskTitle.SetFocus
    On Error GoTo 0
End Sub

Public Sub LoadTaskFromSQLite(ByVal lngTaskId As Long)
    Dim objLoaded As clsTaskItem
    Set objLoaded = TasksRepo().FindEntity(lngTaskId, New clsTaskItem, True)
    If objLoaded Is Nothing Then
        BeginNewTask
        Exit Sub
    End If

    Set m_CurrentTask = objLoaded
    Me!txtTaskID.Value = CStr(m_CurrentTask.TaskID)
    Me!txtTaskTitle.Value = m_CurrentTask.TaskTitle
    Me!cboCategoryName.Value = m_CurrentTask.CategoryName
    Me!cboPriorityLevel.Value = m_CurrentTask.PriorityLevel
    Me!txtEstimateCost.Value = m_CurrentTask.EstimateCost
    Me!txtDueDate.Value = m_CurrentTask.DueDate
    Me!chkIsCompleted.Value = m_CurrentTask.IsCompleted
    Me!txtTaskNotes.Value = m_CurrentTask.TaskNotes

    If Not IsNull(m_CurrentTask.DeletedAt) Then
        Me!txtAuditStamp.Value = "TRASHED: " & FormatTimestampDisplay(m_CurrentTask.DeletedAt) & _
                                 "  |  Created: " & FormatTimestampDisplay(m_CurrentTask.CreatedAt)
        Me!cmdDeleteTask.Caption = "Restore Task"
    Else
        Me!txtAuditStamp.Value = "Created: " & FormatTimestampDisplay(m_CurrentTask.CreatedAt) & _
                                 "   |   Updated: " & FormatTimestampDisplay(m_CurrentTask.UpdatedAt)
        Me!cmdDeleteTask.Caption = "Soft Delete"
    End If
    Me!cmdDeleteTask.Enabled = True
End Sub

Public Function SaveCurrentTaskToSQLite(Optional ByVal blnSilent As Boolean = False) As Long
    On Error GoTo SaveErr
    Dim varDue As Variant
    Dim strCost As String

    If m_CurrentTask Is Nothing Then Set m_CurrentTask = New clsTaskItem

    m_CurrentTask.TaskTitle = Trim$(CStr(Nz(Me!txtTaskTitle.Value, vbNullString)))
    m_CurrentTask.CategoryName = IIf(Len(Trim$(CStr(Nz(Me!cboCategoryName.Value, vbNullString)))) = 0, "General", CStr(Me!cboCategoryName.Value))
    m_CurrentTask.PriorityLevel = IIf(Len(Trim$(CStr(Nz(Me!cboPriorityLevel.Value, vbNullString)))) = 0, "Medium", CStr(Me!cboPriorityLevel.Value))

    strCost = Trim$(CStr(Nz(Me!txtEstimateCost.Value, "0")))
    If IsNumeric(strCost) Then
        m_CurrentTask.EstimateCost = CCur(strCost)
    Else
        m_CurrentTask.EstimateCost = 0@
    End If

    varDue = Me!txtDueDate.Value
    If IsNull(varDue) Or IsEmpty(varDue) Or Len(Trim$(CStr(Nz(varDue, vbNullString)))) = 0 Then
        m_CurrentTask.DueDate = Null
    ElseIf IsDate(varDue) Then
        m_CurrentTask.DueDate = CDate(varDue)
    Else
        Err.Raise sqErrValidation, "frmTaskDetail", "Due Date must be a valid date."
    End If

    m_CurrentTask.IsCompleted = CBool(Nz(Me!chkIsCompleted.Value, False))
    m_CurrentTask.TaskNotes = CStr(Nz(Me!txtTaskNotes.Value, vbNullString))

    TasksRepo().SaveEntity m_CurrentTask, blnDirtyOnly:=True
    LoadTaskFromSQLite m_CurrentTask.TaskID
    NotifyTaskListChanged m_CurrentTask.TaskID
    SaveCurrentTaskToSQLite = m_CurrentTask.TaskID
    Exit Function

SaveErr:
    If blnSilent Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        MsgBox Err.Description, vbExclamation, "SQLite ORM Validation / Save Error"
    End If
End Function

Private Sub cmdPrevTask_Click()
    Dim lngCur As Long
    Dim lngPrev As Long
    If Not m_CurrentTask Is Nothing Then lngCur = m_CurrentTask.TaskID
    If lngCur <= 0 Then
        lngPrev = FindTargetSqliteTaskId("LAST")
    Else
        lngPrev = FindTargetSqliteTaskId("PREV", lngCur)
    End If
    If lngPrev > 0 Then LoadTaskFromSQLite lngPrev
End Sub

Private Sub cmdNextTask_Click()
    Dim lngCur As Long
    Dim lngNext As Long
    If Not m_CurrentTask Is Nothing Then lngCur = m_CurrentTask.TaskID
    If lngCur <= 0 Then
        lngNext = FindTargetSqliteTaskId("FIRST")
    Else
        lngNext = FindTargetSqliteTaskId("NEXT", lngCur)
    End If
    If lngNext > 0 Then LoadTaskFromSQLite lngNext
End Sub

Private Sub cmdNewTask_Click()
    BeginNewTask
End Sub

Private Sub cmdSaveTask_Click()
    SaveCurrentTaskToSQLite False
End Sub

Private Sub cmdDeleteTask_Click()
    Dim lngDeletedId As Long
    Dim lngNeighborId As Long

    If m_CurrentTask Is Nothing Then Exit Sub
    If m_CurrentTask.TaskID <= 0 Then Exit Sub

    lngDeletedId = m_CurrentTask.TaskID
    If Not IsNull(m_CurrentTask.DeletedAt) Then
        TasksRepo().RestoreRecord lngDeletedId
        LoadTaskFromSQLite lngDeletedId
        NotifyTaskListChanged lngDeletedId
    Else
        TasksRepo().DeleteRecord lngDeletedId, blnForceHardDelete:=False
        NotifyTaskListChanged 0
        lngNeighborId = FindTargetSqliteTaskId("NEXT", lngDeletedId)
        If lngNeighborId <= 0 Or lngNeighborId = lngDeletedId Then
            lngNeighborId = FindTargetSqliteTaskId("PREV", lngDeletedId)
        End If
        If lngNeighborId > 0 And lngNeighborId <> lngDeletedId Then
            LoadTaskFromSQLite lngNeighborId
        Else
            BeginNewTask
        End If
    End If
End Sub

Private Sub cmdCloseForm_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

