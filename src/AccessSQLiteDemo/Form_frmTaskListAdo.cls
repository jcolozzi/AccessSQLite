Attribute VB_Name = "Form_frmTaskListAdo"
Attribute VB_Base = "0{92A3AF81-613B-4A68-BDD8-4722F9847098}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

' =============================================================================
' Form_frmTaskListAdo
' Continuous Form dashboard bound directly to a disconnected client-side
' in-memory ADODB.Recordset (zero temp table, zero front-end bloat) with a
' virtual client-only [IsSelected] checkbox column per row and inline ORM
' writeback to AccessSQLiteDemo.db via TasksRepo().SaveFromAdoRow.
' =============================================================================

Private m_SuppressSync As Boolean
Private m_SelectionOnlyEdit As Boolean

Private Sub Form_Load()
    On Error GoTo LoadErr
    m_SuppressSync = True
    Me!txtSearch.Value = vbNullString
    Me!cboFilterCategory.Value = "(All)"
    Me!cboFilterStatus.Value = "Active"
    m_SuppressSync = False
    RefreshFromSQLite 0
    Exit Sub
LoadErr:
    m_SuppressSync = False
End Sub

Public Sub RefreshFromSQLite(Optional ByVal lngSelectTaskId As Long = 0)
    On Error GoTo RefreshErr
    Dim strSearch As String
    Dim strCat As String
    Dim strStatus As String
    Dim objAdoRs As Object
    Dim objClone As Object

    CommitPendingEditQuiet

    strSearch = CStr(Nz(Me!txtSearch.Value, vbNullString))
    strCat = CStr(Nz(Me!cboFilterCategory.Value, "(All)"))
    strStatus = CStr(Nz(Me!cboFilterStatus.Value, "Active"))

    Set objAdoRs = QueryTasksAdoFromSQLite(strSearch, strCat, strStatus)
    Set Me.Recordset = objAdoRs

    Me!lblBanner.Caption = "[In-Memory ADO Recordset]  " & GetSqliteSummaryBanner()
    If StrComp(strStatus, "Trash (Soft-Deleted)", vbTextCompare) = 0 Then
        Me!cmdDeleteSelected.Caption = "Restore Task"
    Else
        Me!cmdDeleteSelected.Caption = "Soft Delete"
    End If

    If lngSelectTaskId > 0 And Not objAdoRs Is Nothing Then
        If objAdoRs.RecordCount > 0 Then
            Set objClone = objAdoRs.Clone
            objClone.MoveFirst
            objClone.Find "TaskID = " & lngSelectTaskId
            If Not objClone.EOF Then
                Me.Bookmark = objClone.Bookmark
            End If
        End If
    End If

    UpdateSelectionButtonsCaption
    Exit Sub

RefreshErr:
    MsgBox "Failed to bind in-memory ADO recordset from SQLite: " & Err.Description, vbExclamation, "AccessSQLite ADO Sync Error"
End Sub

Private Sub Form_Current()
    On Error Resume Next
    If HasActiveAdoRow() Then
        If IsCurrentRowTrashed() Then
            Me!cmdDeleteSelected.Caption = "Restore Task"
        Else
            Me!cmdDeleteSelected.Caption = "Soft Delete"
        End If
    End If
    On Error GoTo 0
End Sub

' Writes inline Continuous Form edits directly back into AccessSQLiteDemo.db via SaveFromAdoRow.
' Pure virtual [IsSelected] checkbox clicks stay 100% in RAM without touching disk.
Private Sub Form_AfterUpdate()
    On Error GoTo SaveRowErr
    If m_SuppressSync Or m_SelectionOnlyEdit Then Exit Sub
    If Not HasActiveAdoRow() Then Exit Sub

    TasksRepo().SaveFromAdoRow Me.Recordset, True
    Me!lblBanner.Caption = "[In-Memory ADO Recordset]  " & GetSqliteSummaryBanner()
    If CurrentProject.AllForms("frmTaskList").IsLoaded Then
        Forms("frmTaskList").RefreshFromSQLite CurrentSelectedTaskId()
    End If
    Exit Sub

SaveRowErr:
    MsgBox Err.Description, vbExclamation, "SQLite Inline ADO Save Error"
    RefreshFromSQLite CurrentSelectedTaskId()
End Sub

Private Sub chkRowSelect_AfterUpdate()
    m_SelectionOnlyEdit = True
    CommitPendingEditQuiet
    m_SelectionOnlyEdit = False
    UpdateSelectionButtonsCaption
End Sub

Private Sub chkRowDone_AfterUpdate()
    CommitPendingEditQuiet
End Sub

Private Sub txtSearch_AfterUpdate()
    If Not m_SuppressSync Then RefreshFromSQLite CurrentSelectedTaskId()
End Sub

Private Sub cboFilterCategory_AfterUpdate()
    If Not m_SuppressSync Then RefreshFromSQLite 0
End Sub

Private Sub cboFilterStatus_AfterUpdate()
    If Not m_SuppressSync Then RefreshFromSQLite 0
End Sub

Private Sub cmdApplyFilter_Click()
    RefreshFromSQLite 0
End Sub

Private Sub cmdResetFilter_Click()
    m_SuppressSync = True
    Me!txtSearch.Value = vbNullString
    Me!cboFilterCategory.Value = "(All)"
    Me!cboFilterStatus.Value = "Active"
    m_SuppressSync = False
    RefreshFromSQLite 0
End Sub

Private Sub cmdRowOpenDetail_Click()
    OpenDetailForSelectedRow
End Sub

Private Sub cmdEditSelected_Click()
    OpenDetailForSelectedRow
End Sub

Private Sub cmdNewTask_Click()
    CommitPendingEditQuiet
    If CurrentProject.AllForms("frmTaskDetail").IsLoaded Then
        Forms("frmTaskDetail").BeginNewTask
        Forms("frmTaskDetail").SetFocus
    Else
        DoCmd.OpenForm "frmTaskDetail", acNormal, , , , , "NEW"
    End If
End Sub

Private Sub cmdToggleSelected_Click()
    Dim lngId As Long
    lngId = CurrentSelectedTaskId()
    If lngId <= 0 Then Exit Sub
    CommitPendingEditQuiet
    ToggleTaskCompletedInSQLite lngId
    RefreshFromSQLite lngId
    If CurrentProject.AllForms("frmTaskList").IsLoaded Then
        Forms("frmTaskList").RefreshFromSQLite lngId
    End If
End Sub

Private Sub cmdDeleteSelected_Click()
    Dim lngId As Long
    lngId = CurrentSelectedTaskId()
    If lngId <= 0 Then Exit Sub
    CommitPendingEditQuiet

    If IsCurrentRowTrashed() Then
        TasksRepo().RestoreRecord lngId
        RefreshFromSQLite lngId
    Else
        TasksRepo().DeleteRecord lngId, blnForceHardDelete:=False
        RefreshFromSQLite 0
    End If
    If CurrentProject.AllForms("frmTaskList").IsLoaded Then
        Forms("frmTaskList").RefreshFromSQLite 0
    End If
End Sub

Private Sub cmdSelectAll_Click()
    Dim colIds As Collection
    Set colIds = SelectedTaskIds()
    If Me.Recordset Is Nothing Then Exit Sub
    If Me.Recordset.RecordCount = 0 Then Exit Sub

    If colIds.Count < Me.Recordset.RecordCount Then
        SetAllRowsSelected True
    Else
        SetAllRowsSelected False
    End If
End Sub

Private Sub cmdBatchComplete_Click()
    Dim lngUpdated As Long
    lngUpdated = BatchCompleteSelectedRows()
    If lngUpdated = 0 Then
        MsgBox "Check one or more [Sel] boxes first to batch-complete tasks.", vbInformation, "In-Memory ADO Multi-Select"
    End If
End Sub

Private Sub cmdReloadSqlite_Click()
    RefreshFromSQLite CurrentSelectedTaskId()
End Sub

Private Sub cmdOpenStudio_Click()
    CommitPendingEditQuiet
    AccessSQLite.OpenSQLiteStudio DemoSqliteFilePath()
End Sub

' Returns a Collection of checked TaskIDs from the virtual in-memory [IsSelected] column.
Public Function SelectedTaskIds() As Collection
    m_SelectionOnlyEdit = True
    CommitPendingEditQuiet
    m_SelectionOnlyEdit = False

    If Me.Recordset Is Nothing Then
        Set SelectedTaskIds = New Collection
    Else
        Set SelectedTaskIds = GetDemoConnection().NewAdoBridge().GetSelectedKeys(Me.Recordset, "TaskID", "IsSelected")
    End If
End Function

Public Sub SetAllRowsSelected(ByVal blnSelected As Boolean)
    m_SelectionOnlyEdit = True
    CommitPendingEditQuiet
    m_SelectionOnlyEdit = False

    If Me.Recordset Is Nothing Then Exit Sub
    GetDemoConnection().NewAdoBridge().SetAllSelected Me.Recordset, blnSelected, "IsSelected"
    Me.Refresh
    UpdateSelectionButtonsCaption
End Sub

Public Function BatchCompleteSelectedRows() As Long
    Dim colIds As Collection
    Dim lngCnt As Long
    Set colIds = SelectedTaskIds()
    If colIds.Count = 0 Then Exit Function

    lngCnt = BatchSetCompletedInSQLite(colIds, True)
    RefreshFromSQLite CurrentSelectedTaskId()
    If CurrentProject.AllForms("frmTaskList").IsLoaded Then
        Forms("frmTaskList").RefreshFromSQLite 0
    End If
    BatchCompleteSelectedRows = lngCnt
End Function

Private Sub UpdateSelectionButtonsCaption()
    On Error Resume Next
    Dim colIds As Collection
    Dim lngTotal As Long
    Set colIds = SelectedTaskIds()
    If Not Me.Recordset Is Nothing Then lngTotal = Me.Recordset.RecordCount

    If colIds.Count > 0 Then
        Me!cmdBatchComplete.Caption = "Complete (" & colIds.Count & ")"
    Else
        Me!cmdBatchComplete.Caption = "Complete Checked"
    End If

    If lngTotal > 0 And colIds.Count >= lngTotal Then
        Me!cmdSelectAll.Caption = "Clear Sel."
    Else
        Me!cmdSelectAll.Caption = "Select All"
    End If
    On Error GoTo 0
End Sub

Private Sub OpenDetailForSelectedRow()
    Dim lngId As Long
    lngId = CurrentSelectedTaskId()
    If lngId <= 0 Then
        cmdNewTask_Click
        Exit Sub
    End If
    CommitPendingEditQuiet

    If CurrentProject.AllForms("frmTaskDetail").IsLoaded Then
        Forms("frmTaskDetail").LoadTaskFromSQLite lngId
        Forms("frmTaskDetail").SetFocus
    Else
        DoCmd.OpenForm "frmTaskDetail", acNormal, , , , , CStr(lngId)
    End If
End Sub

Private Sub CommitPendingEditQuiet()
    On Error Resume Next
    If HasActiveAdoRow() Then
        If Me.Dirty Then Me.Dirty = False
    End If
    On Error GoTo 0
End Sub

Private Function HasActiveAdoRow() As Boolean
    On Error Resume Next
    If Not Me.Recordset Is Nothing Then
        If Me.Recordset.RecordCount > 0 Then
            If Not (Me.Recordset.BOF Or Me.Recordset.EOF) Then
                HasActiveAdoRow = True
            End If
        End If
    End If
    On Error GoTo 0
End Function

Private Function IsCurrentRowTrashed() As Boolean
    On Error Resume Next
    If HasActiveAdoRow() Then
        IsCurrentRowTrashed = Not IsNull(Me.Recordset.Fields.Item("DeletedAt").Value)
    End If
    On Error GoTo 0
End Function

Private Function CurrentSelectedTaskId() As Long
    On Error Resume Next
    If HasActiveAdoRow() Then
        If Not IsNull(Me!txtRowTaskID.Value) Then
            CurrentSelectedTaskId = CLng(Me!txtRowTaskID.Value)
        End If
    End If
    On Error GoTo 0
End Function

