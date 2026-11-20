Attribute VB_Name = "Form_frmTaskList"
Attribute VB_Base = "0{5A9820A6-0EFE-442D-B414-22DC798F1A64}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

' =============================================================================
' Form_frmTaskList
' Continuous Form dashboard backed by AccessSQLiteDemo.db via the referenced
' AccessSQLite Class Library & ORM. Syncs filtered SQLite query results into
' tblTaskBuffer and writes inline row edits & actions straight back to SQLite.
' =============================================================================

Private m_SuppressSync As Boolean

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
    Dim rstClone As DAO.Recordset

    CommitPendingEditQuiet

    strSearch = CStr(Nz(Me!txtSearch.Value, vbNullString))
    strCat = CStr(Nz(Me!cboFilterCategory.Value, "(All)"))
    strStatus = CStr(Nz(Me!cboFilterStatus.Value, "Active"))

    SyncBufferFromSQLite strSearch, strCat, strStatus
    Me.Requery

    Me!lblBanner.Caption = GetSqliteSummaryBanner()
    If StrComp(strStatus, "Trash (Soft-Deleted)", vbTextCompare) = 0 Then
        Me!cmdDeleteSelected.Caption = "Restore Task"
    Else
        Me!cmdDeleteSelected.Caption = "Soft Delete"
    End If

    If lngSelectTaskId > 0 Then
        Set rstClone = Me.RecordsetClone
        If Not (rstClone.BOF And rstClone.EOF) Then
            rstClone.FindFirst "[TaskID] = " & lngSelectTaskId
            If Not rstClone.NoMatch Then
                Me.Bookmark = rstClone.Bookmark
            End If
        End If
        rstClone.Close
    End If
    Exit Sub

RefreshErr:
    MsgBox "Failed to sync from SQLite: " & Err.Description, vbExclamation, "AccessSQLite Sync Error"
End Sub

Private Sub Form_Current()
    On Error Resume Next
    If Not (Me.RecordsetClone.BOF And Me.RecordsetClone.EOF) Then
        If CBool(Nz(Me!IsDeletedRow, False)) Then
            Me!cmdDeleteSelected.Caption = "Restore Task"
        Else
            Me!cmdDeleteSelected.Caption = "Soft Delete"
        End If
    End If
    On Error GoTo 0
End Sub

' Writes any inline Continuous Form edit directly back into AccessSQLiteDemo.db via ORM Repository.
Private Sub Form_AfterUpdate()
    On Error GoTo SaveRowErr
    If m_SuppressSync Then Exit Sub
    If IsNull(Me!txtRowTaskID.Value) Then Exit Sub

    PushBufferRowToSQLite _
        CLng(Me!txtRowTaskID.Value), _
        CStr(Nz(Me!txtRowTitle.Value, vbNullString)), _
        CStr(Nz(Me!cboRowCategory.Value, "General")), _
        CStr(Nz(Me!cboRowPriority.Value, "Medium")), _
        CCur(Nz(Me!txtRowCost.Value, 0@)), _
        Me!txtRowDueDate.Value, _
        CBool(Nz(Me!chkRowDone.Value, False))

    Me!lblBanner.Caption = GetSqliteSummaryBanner()
    Exit Sub

SaveRowErr:
    MsgBox Err.Description, vbExclamation, "SQLite Inline Save Error"
    RefreshFromSQLite CurrentSelectedTaskId()
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
End Sub

Private Sub cmdDeleteSelected_Click()
    Dim lngId As Long
    lngId = CurrentSelectedTaskId()
    If lngId <= 0 Then Exit Sub
    CommitPendingEditQuiet

    If CBool(Nz(Me!IsDeletedRow, False)) Then
        TasksRepo().RestoreRecord lngId
        RefreshFromSQLite lngId
    Else
        TasksRepo().DeleteRecord lngId, blnForceHardDelete:=False
        RefreshFromSQLite 0
    End If
End Sub

Private Sub cmdReloadSqlite_Click()
    RefreshFromSQLite CurrentSelectedTaskId()
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
    If Not (Me.RecordsetClone.BOF And Me.RecordsetClone.EOF) Then
        If Me.Dirty Then Me.Dirty = False
    End If
    On Error GoTo 0
End Sub

Private Function CurrentSelectedTaskId() As Long
    On Error Resume Next
    If Not (Me.RecordsetClone.BOF And Me.RecordsetClone.EOF) Then
        If Not IsNull(Me!txtRowTaskID.Value) Then
            CurrentSelectedTaskId = CLng(Me!txtRowTaskID.Value)
        End If
    End If
    On Error GoTo 0
End Function

