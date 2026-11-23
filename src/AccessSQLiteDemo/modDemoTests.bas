Attribute VB_Name = "modDemoTests"
Option Compare Database
Option Explicit

Public Function RunDemoFormsVerification() As String
    Dim strStep As String
    Dim lngInitialSqliteCnt As Long
    Dim lngNewId As Long
    Dim strSavedTitle As String
    Dim curSavedCost As Currency
    Dim lngFilteredCnt As Long
    Dim lngAdoFilteredCnt As Long
    Dim objAdoRs As Object
    Dim colSelIds As Collection

    On Error GoTo TestErr

    strStep = "0-InstallDemoRibbon"
    InstallDemoRibbon
    If CurrentDb.Properties("CustomRibbonID").Value <> "AccessSQLiteDemo" Then
        Err.Raise vbObjectError + 3000, "modDemoTests", "CustomRibbonID was not set to AccessSQLiteDemo."
    End If

    strStep = "1-OpenDaoTaskList"
    DoCmd.OpenForm "frmTaskList", acNormal

    strStep = "1b-DaoCountCheck"
    lngInitialSqliteCnt = TasksRepo().CountAll(False)
    Forms("frmTaskList").RecordsetClone.MoveLast
    If Forms("frmTaskList").RecordsetClone.RecordCount <> lngInitialSqliteCnt Then
        Err.Raise vbObjectError + 3001, "modDemoTests", "Initial DAO Continuous Form row count mismatch."
    End If

    strStep = "2-DaoFilterCategory"
    Dim lngExpectedEng As Long
    lngExpectedEng = TasksRepo().NewQuery().WhereEq("CategoryName", "Engineering").CountRows()
    Forms("frmTaskList")!cboFilterCategory.Value = "Engineering"
    Forms("frmTaskList").RefreshFromSQLite 0
    If lngExpectedEng > 0 Then Forms("frmTaskList").RecordsetClone.MoveLast
    lngFilteredCnt = Forms("frmTaskList").RecordsetClone.RecordCount
    If lngFilteredCnt <> lngExpectedEng Then
        Err.Raise vbObjectError + 3002, "modDemoTests", "Expected " & lngExpectedEng & " DAO Engineering tasks, got " & lngFilteredCnt
    End If

    strStep = "2b-DaoEmptySearchCheck"
    Forms("frmTaskList")!txtSearch.Value = "zzzz_no_match_check"
    Forms("frmTaskList").RefreshFromSQLite 0
    If Forms("frmTaskList").RecordsetClone.RecordCount <> 0 Then
        Err.Raise vbObjectError + 3006, "modDemoTests", "Expected 0 DAO rows for non-matching search."
    End If

    strStep = "2c-DaoResetFilter"
    Forms("frmTaskList")!txtSearch.Value = vbNullString
    Forms("frmTaskList")!cboFilterCategory.Value = "(All)"
    Forms("frmTaskList").RefreshFromSQLite 0

    strStep = "3-OpenAdoTaskList"
    DoCmd.OpenForm "frmTaskListAdo", acNormal
    If Len(Forms("frmTaskListAdo").RecordSource) <> 0 Then
        Err.Raise vbObjectError + 3007, "modDemoTests", "frmTaskListAdo RecordSource must be empty (pure in-memory ADO)."
    End If
    If Forms("frmTaskListAdo").Recordset.RecordCount <> lngInitialSqliteCnt Then
        Err.Raise vbObjectError + 3008, "modDemoTests", "frmTaskListAdo initial ADO RecordCount mismatch."
    End If

    strStep = "3b-AdoFilterCheck"
    Forms("frmTaskListAdo")!cboFilterCategory.Value = "Engineering"
    Forms("frmTaskListAdo").RefreshFromSQLite 0
    lngAdoFilteredCnt = Forms("frmTaskListAdo").Recordset.RecordCount
    If lngAdoFilteredCnt <> lngExpectedEng Then
        Err.Raise vbObjectError + 3009, "modDemoTests", "Expected " & lngExpectedEng & " ADO Engineering tasks, got " & lngAdoFilteredCnt
    End If
    Forms("frmTaskListAdo")!cboFilterCategory.Value = "(All)"
    Forms("frmTaskListAdo").RefreshFromSQLite 0

    strStep = "4-OpenTaskDetailNew"
    DoCmd.OpenForm "frmTaskDetail", acNormal, , , , , "NEW"

    strStep = "4b-PopulateDetailFields"
    Forms("frmTaskDetail")!txtTaskTitle.Value = "E2E Verified SQLite Form Task"
    Forms("frmTaskDetail")!cboCategoryName.Value = "Analytics"
    Forms("frmTaskDetail")!cboPriorityLevel.Value = "High"
    Forms("frmTaskDetail")!txtEstimateCost.Value = 777.25@
    Forms("frmTaskDetail")!txtDueDate.Value = DateSerial(2026, 12, 31)
    Forms("frmTaskDetail")!txtTaskNotes.Value = "Created via frmTaskDetail unbound ORM single-page form."

    strStep = "4c-SaveDetailToSQLite"
    lngNewId = Forms("frmTaskDetail").SaveCurrentTaskToSQLite(True)
    If lngNewId <= 0 Then
        Err.Raise vbObjectError + 3003, "modDemoTests", "frmTaskDetail failed to return new SQLite TaskID."
    End If

    strStep = "4d-VerifyBothContinuousFormsAutoSynced"
    Forms("frmTaskList").RecordsetClone.MoveLast
    If Forms("frmTaskList").RecordsetClone.RecordCount <> (lngInitialSqliteCnt + 1) Then
        Err.Raise vbObjectError + 3004, "modDemoTests", "frmTaskList (DAO) did not auto-sync after frmTaskDetail save."
    End If
    If Forms("frmTaskListAdo").Recordset.RecordCount <> (lngInitialSqliteCnt + 1) Then
        Err.Raise vbObjectError + 3010, "modDemoTests", "frmTaskListAdo (ADO) did not auto-sync after frmTaskDetail save."
    End If

    strStep = "5-AdoVirtualCheckboxAndBatchComplete"
    Forms("frmTaskListAdo").SetAllRowsSelected False
    Set objAdoRs = Forms("frmTaskListAdo").Recordset.Clone
    objAdoRs.MoveFirst
    objAdoRs.Find "TaskID = " & lngNewId
    If objAdoRs.EOF Then
        Err.Raise vbObjectError + 3011, "modDemoTests", "Could not locate newly inserted TaskID in ADO Recordset."
    End If
    objAdoRs.Fields.Item("IsSelected").Value = True
    objAdoRs.Update
    Set colSelIds = Forms("frmTaskListAdo").SelectedTaskIds()
    If colSelIds.Count <> 1 Or CLng(colSelIds(1)) <> lngNewId Then
        Err.Raise vbObjectError + 3012, "modDemoTests", "Virtual ADO [IsSelected] checkbox state did not return expected TaskID."
    End If
    If Forms("frmTaskListAdo").BatchCompleteSelectedRows() <> 1 Then
        Err.Raise vbObjectError + 3013, "modDemoTests", "BatchCompleteSelectedRows did not update 1 selected ADO row."
    End If

    strStep = "5b-AdoInlineSaveToSQLite"
    Set objAdoRs = Forms("frmTaskListAdo").Recordset
    objAdoRs.MoveFirst
    objAdoRs.Find "TaskID = " & lngNewId
    objAdoRs.Fields.Item("TaskTitle").Value = "E2E Verified ADO Inline Update"
    objAdoRs.Fields.Item("EstimateCost").Value = 888.5@
    objAdoRs.Update
    TasksRepo().SaveFromAdoRow objAdoRs, True

    strStep = "5c-ReloadDetailCheck"
    Forms("frmTaskDetail").LoadTaskFromSQLite lngNewId
    strSavedTitle = CStr(Forms("frmTaskDetail")!txtTaskTitle.Value)
    curSavedCost = CCur(Forms("frmTaskDetail")!txtEstimateCost.Value)
    If strSavedTitle <> "E2E Verified ADO Inline Update" Or curSavedCost <> 888.5@ Or Not CBool(Forms("frmTaskDetail")!chkIsCompleted.Value) Then
        Err.Raise vbObjectError + 3005, "modDemoTests", "Single-page form did not read ADO inline + batch SQLite updates."
    End If

    strStep = "6-CrossLibraryAxLiteCheck"
    Dim strAxKey As String
    strAxKey = AccessSQLite.OpenAxLite(DemoSqliteFilePath())
    If Len(strAxKey) = 0 Or AccessSQLite.OpenAxLiteCount() < 1 Then
        Err.Raise vbObjectError + 3014, "modDemoTests", "OpenAxLite cross-library window failed to open."
    End If
    AccessSQLite.CloseAllAxLiteWindows

    strStep = "7-Cleanup"
    DoCmd.Close acForm, "frmTaskDetail", acSaveNo
    TasksRepo().DeleteRecord lngNewId, blnForceHardDelete:=True
    Forms("frmTaskList").RefreshFromSQLite 0
    Forms("frmTaskListAdo").RefreshFromSQLite 0
    DoCmd.Close acForm, "frmTaskListAdo", acSaveNo
    DoCmd.Close acForm, "frmTaskList", acSaveNo

    RunDemoFormsVerification = "PASS | DAO=" & lngInitialSqliteCnt & " | ADO=" & lngInitialSqliteCnt & " | AxLite=" & strAxKey & " | RoundTripID=" & lngNewId
    Exit Function

TestErr:
    Dim strErr As String
    strErr = "FAIL at [" & strStep & "] #" & Err.Number & " (" & Err.Source & "): " & Err.Description
    On Error Resume Next
    If lngNewId > 0 Then TasksRepo().DeleteRecord lngNewId, True
    DoCmd.Close acForm, "frmTaskDetail", acSaveNo
    DoCmd.Close acForm, "frmTaskListAdo", acSaveNo
    DoCmd.Close acForm, "frmTaskList", acSaveNo
    On Error GoTo 0
    RunDemoFormsVerification = strErr
End Function
