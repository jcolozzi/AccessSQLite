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

    On Error GoTo TestErr

    strStep = "1-OpenTaskList"
    DoCmd.OpenForm "frmTaskList", acNormal

    strStep = "1b-CountCheck"
    lngInitialSqliteCnt = TasksRepo().CountAll(False)
    Forms("frmTaskList").RecordsetClone.MoveLast
    If Forms("frmTaskList").RecordsetClone.RecordCount <> lngInitialSqliteCnt Then
        Err.Raise vbObjectError + 3001, "modDemoTests", "Initial Continuous Form row count mismatch."
    End If

    strStep = "2-FilterCategory"
    Dim lngExpectedEng As Long
    lngExpectedEng = TasksRepo().NewQuery().WhereEq("CategoryName", "Engineering").CountRows()
    Forms("frmTaskList")!cboFilterCategory.Value = "Engineering"
    Forms("frmTaskList").RefreshFromSQLite 0
    If lngExpectedEng > 0 Then Forms("frmTaskList").RecordsetClone.MoveLast
    lngFilteredCnt = Forms("frmTaskList").RecordsetClone.RecordCount
    If lngFilteredCnt <> lngExpectedEng Then
        Err.Raise vbObjectError + 3002, "modDemoTests", "Expected " & lngExpectedEng & " Engineering tasks, got " & lngFilteredCnt
    End If

    strStep = "2b-EmptySearchCheck"
    Forms("frmTaskList")!txtSearch.Value = "zzzz_no_match_check"
    Forms("frmTaskList").RefreshFromSQLite 0
    If Forms("frmTaskList").RecordsetClone.RecordCount <> 0 Then
        Err.Raise vbObjectError + 3006, "modDemoTests", "Expected 0 rows for non-matching search, got " & Forms("frmTaskList").RecordsetClone.RecordCount
    End If

    strStep = "2c-ResetFilterFromEmpty"
    Forms("frmTaskList")!txtSearch.Value = vbNullString
    Forms("frmTaskList")!cboFilterCategory.Value = "(All)"
    Forms("frmTaskList").RefreshFromSQLite 0

    strStep = "3-OpenTaskDetailNew"
    DoCmd.OpenForm "frmTaskDetail", acNormal, , , , , "NEW"

    strStep = "3b-PopulateDetailFields"
    Forms("frmTaskDetail")!txtTaskTitle.Value = "E2E Verified SQLite Form Task"
    Forms("frmTaskDetail")!cboCategoryName.Value = "Analytics"
    Forms("frmTaskDetail")!cboPriorityLevel.Value = "High"
    Forms("frmTaskDetail")!txtEstimateCost.Value = 777.25@
    Forms("frmTaskDetail")!txtDueDate.Value = DateSerial(2026, 12, 31)
    Forms("frmTaskDetail")!txtTaskNotes.Value = "Created via frmTaskDetail unbound ORM single-page form."

    strStep = "3c-SaveDetailToSQLite"
    lngNewId = Forms("frmTaskDetail").SaveCurrentTaskToSQLite(True)
    If lngNewId <= 0 Then
        Err.Raise vbObjectError + 3003, "modDemoTests", "frmTaskDetail failed to return new SQLite TaskID."
    End If

    strStep = "3d-VerifyListAutoSynced"
    Forms("frmTaskList").RecordsetClone.MoveLast
    If Forms("frmTaskList").RecordsetClone.RecordCount <> (lngInitialSqliteCnt + 1) Then
        Err.Raise vbObjectError + 3004, "modDemoTests", "frmTaskList did not auto-sync after frmTaskDetail save."
    End If

    strStep = "4-InlineUpdateSync"
    PushBufferRowToSQLite lngNewId, "E2E Verified Inline Update", "Analytics", "Critical", 888.5@, DateSerial(2026, 12, 31), True
    Forms("frmTaskList").RefreshFromSQLite lngNewId

    strStep = "4b-ReloadDetailCheck"
    Forms("frmTaskDetail").LoadTaskFromSQLite lngNewId
    strSavedTitle = CStr(Forms("frmTaskDetail")!txtTaskTitle.Value)
    curSavedCost = CCur(Forms("frmTaskDetail")!txtEstimateCost.Value)
    If strSavedTitle <> "E2E Verified Inline Update" Or curSavedCost <> 888.5@ Then
        Err.Raise vbObjectError + 3005, "modDemoTests", "Single-page form did not read inline SQLite update."
    End If

    strStep = "5-Cleanup"
    DoCmd.Close acForm, "frmTaskDetail", acSaveNo
    TasksRepo().DeleteRecord lngNewId, blnForceHardDelete:=True
    Forms("frmTaskList").RefreshFromSQLite 0
    DoCmd.Close acForm, "frmTaskList", acSaveNo

    RunDemoFormsVerification = "PASS | Initial=" & lngInitialSqliteCnt & " | FilteredEng=" & lngFilteredCnt & " | RoundTripID=" & lngNewId
    Exit Function

TestErr:
    Dim strErr As String
    strErr = "FAIL at [" & strStep & "] #" & Err.Number & " (" & Err.Source & "): " & Err.Description
    On Error Resume Next
    If lngNewId > 0 Then TasksRepo().DeleteRecord lngNewId, True
    DoCmd.Close acForm, "frmTaskDetail", acSaveNo
    DoCmd.Close acForm, "frmTaskList", acSaveNo
    On Error GoTo 0
    RunDemoFormsVerification = strErr
End Function
