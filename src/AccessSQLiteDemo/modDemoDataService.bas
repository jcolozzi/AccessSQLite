Attribute VB_Name = "modDemoDataService"
Option Compare Database
Option Explicit

' =============================================================================
' modDemoDataService
' Data service connecting AccessSQLiteDemo.accdb forms:
'   1. frmTaskList    - Continuous Form backed by DAO staging table (tblTaskBuffer)
'   2. frmTaskListAdo - Continuous Form backed by In-Memory ADODB.Recordset (zero temp table)
'   3. frmTaskDetail  - Unbound Single-Page Form backed by clsTaskItem (ISQLiteEntity)
'   4. AxLite         - Monaco SQL front-end client hosted in AccessSQLite.accdb
' =============================================================================

Private m_SqliteConn As AccessSQLite.clsSQLiteConnection
Private m_Initialized As Boolean

Public Function DemoSqliteFilePath() As String
    DemoSqliteFilePath = CurrentProject.Path & "\AccessSQLiteDemo.db"
End Function

Public Function GetDemoConnection() As AccessSQLite.clsSQLiteConnection
    If m_SqliteConn Is Nothing Then
        Set m_SqliteConn = AccessSQLite.OpenSQLite(DemoSqliteFilePath(), sqProfileAuto, sqDriverAuto)
        m_Initialized = False
    ElseIf Not m_SqliteConn.IsOpen Then
        m_SqliteConn.OpenConnection DemoSqliteFilePath(), sqProfileAuto, sqDriverAuto
        m_Initialized = False
    End If

    If Not m_Initialized Then
        ConfigureSchemaAndSeed m_SqliteConn
        m_Initialized = True
    End If

    Set GetDemoConnection = m_SqliteConn
End Function

Public Sub CloseDemoConnection()
    On Error Resume Next
    If Not m_SqliteConn Is Nothing Then
        m_SqliteConn.CloseConnection
        Set m_SqliteConn = Nothing
    End If
    m_Initialized = False
    On Error GoTo 0
End Sub

Public Function TasksRepo() As AccessSQLite.clsSQLiteRepository
    Set TasksRepo = GetDemoConnection().Repository("Tasks")
End Function

Private Sub ConfigureSchemaAndSeed(ByVal objConn As AccessSQLite.clsSQLiteConnection)
    Dim objMig As AccessSQLite.clsSQLiteMigrator
    Dim objMap As AccessSQLite.clsSQLiteTableMap

    ' 1. Run versioned schema migration in AccessSQLiteDemo.db
    Set objMig = objConn.NewMigrator()
    objMig.AddMigration 1, "Create Tasks table and indexes", _
        "CREATE TABLE IF NOT EXISTS Tasks (" & _
        "  TaskID INTEGER PRIMARY KEY AUTOINCREMENT," & _
        "  TaskTitle TEXT NOT NULL," & _
        "  CategoryName TEXT NOT NULL DEFAULT 'General'," & _
        "  PriorityLevel TEXT NOT NULL DEFAULT 'Medium'," & _
        "  EstimateCost CURRENCY NOT NULL DEFAULT 0," & _
        "  DueDate DATE," & _
        "  IsCompleted BOOLEAN NOT NULL DEFAULT 0," & _
        "  TaskNotes MEMO," & _
        "  CreatedAt DATETIME," & _
        "  UpdatedAt DATETIME," & _
        "  DeletedAt DATETIME" & _
        ");" & _
        "CREATE INDEX IF NOT EXISTS idx_Tasks_Category ON Tasks (CategoryName);" & _
        "CREATE INDEX IF NOT EXISTS idx_Tasks_Completed ON Tasks (IsCompleted);", _
        "DROP TABLE IF EXISTS Tasks;"
    objMig.MigrateUp

    ' 2. Register ORM TableMap with validation rules, timestamps & soft-delete
    Set objMap = objConn.DefineTable("Tasks", "TaskID") _
        .AddPrimaryKey("TaskID", sqColInt64, True) _
        .AddColumn("TaskTitle", sqColText, "Task Title", blnRequired:=True, lngMaxLength:=150) _
        .AddColumn("CategoryName", sqColText, "Category", blnRequired:=True, lngMaxLength:=50, varDefault:="General") _
        .AddColumn("PriorityLevel", sqColText, "Priority", blnRequired:=True, lngMaxLength:=30, varDefault:="Medium") _
        .AddColumn("EstimateCost", sqColCurrency, "Estimated Cost", blnRequired:=True, varDefault:=0@) _
        .AddColumn("DueDate", sqColDate, "Due Date") _
        .AddColumn("IsCompleted", sqColBoolean, "Completed", blnRequired:=True, varDefault:=False) _
        .AddColumn("TaskNotes", sqColMemo, "Notes") _
        .AddTimestamps("CreatedAt", "UpdatedAt") _
        .AddSoftDelete("DeletedAt")

    ' 3. Seed realistic demo records on first launch
    If CLng(objConn.QueryScalar("SELECT COUNT(*) FROM Tasks;")) = 0 Then
        SeedInitialTasks objConn.Repository("Tasks")
    End If
End Sub

Private Sub SeedInitialTasks(ByVal objRepo As AccessSQLite.clsSQLiteRepository)
    Dim colSeed As New Collection

    colSeed.Add MakeSeedRow("Configure SMB-safe SQLite pragma profile", "Engineering", "Critical", 1450@, DateSerial(2026, 11, 25), True, _
        "Verified DELETE journal mode on UNC/mapped shares and WAL on local SSD drives.")
    colSeed.Add MakeSeedRow("Build Continuous Form task dashboard", "Engineering", "High", 1800@, DateSerial(2026, 11, 28), False, _
        "Syncs filtered SQLite query results into tblTaskBuffer and writes inline edits straight back to SQLite.")
    colSeed.Add MakeSeedRow("Design Single-Page ORM detail editor", "Engineering", "High", 1600@, DateSerial(2026, 12, 2), False, _
        "Binds unbound controls directly to clsTaskItem (ISQLiteEntity) via clsSQLiteRepository.")
    colSeed.Add MakeSeedRow("Audit Q4 vendor contract renewals", "Finance", "Critical", 3200@, DateSerial(2026, 12, 5), False, _
        "Review service-level tiers and compare annual support estimates before fiscal cutoff.")
    colSeed.Add MakeSeedRow("Publish monthly clinical throughput KPIs", "Analytics", "Medium", 950@, DateSerial(2026, 12, 8), False, _
        "Aggregate regional encounter counts in SQLite and export summary snapshot to Access.")
    colSeed.Add MakeSeedRow("Validate 64-bit ROWID & Currency precision", "Analytics", "High", 750@, DateSerial(2026, 11, 22), True, _
        "Tested 5-billion integer keys and 4-decimal fixed-point Currency values across WinSQLite and Bundled drivers.")
    colSeed.Add MakeSeedRow("Standardize front-end reference deployment", "Operations", "Medium", 620@, DateSerial(2026, 12, 12), False, _
        "Verify CodeProject.Path resolution when front-end .accdb runs from user workstation folders.")
    colSeed.Add MakeSeedRow("Archive legacy CSV staging scripts", "Operations", "Low", 300@, DateSerial(2026, 12, 18), True, _
        "Replaced ad-hoc text imports with clsSQLiteAccessBridge transactional imports.")
    colSeed.Add MakeSeedRow("Prepare team walkthrough & demo deck", "General", "Medium", 480@, DateSerial(2026, 12, 15), False, _
        "Demonstrate live filtering, inline completion toggles, soft-delete trash view, and ORM validation.")
    colSeed.Add MakeSeedRow("Schedule nightly online SQLite backup check", "Operations", "Low", 540@, DateSerial(2026, 12, 20), False, _
        "Use clsSQLiteConnection.BackupToFile and PRAGMA integrity_check during off-peak window.")

    objRepo.InsertMany colSeed
End Sub

Private Function MakeSeedRow( _
    ByVal strTitle As String, _
    ByVal strCat As String, _
    ByVal strPrio As String, _
    ByVal curCost As Currency, _
    ByVal dtmDue As Date, _
    ByVal blnDone As Boolean, _
    ByVal strNotes As String) As Scripting.Dictionary

    Dim dictRow As Scripting.Dictionary
    Set dictRow = AccessSQLite.NewTextDict()
    dictRow("TaskTitle") = strTitle
    dictRow("CategoryName") = strCat
    dictRow("PriorityLevel") = strPrio
    dictRow("EstimateCost") = curCost
    dictRow("DueDate") = dtmDue
    dictRow("IsCompleted") = blnDone
    dictRow("TaskNotes") = strNotes
    Set MakeSeedRow = dictRow
End Function

' Shared filtered QueryBuilder used by both the DAO Staging Form and the ADO In-Memory Form.
Public Function BuildFilteredTasksQuery( _
    Optional ByVal strSearchText As String = vbNullString, _
    Optional ByVal strCategoryFilter As String = vbNullString, _
    Optional ByVal strStatusFilter As String = "Active") As AccessSQLite.clsSQLiteQueryBuilder

    Dim objQb As AccessSQLite.clsSQLiteQueryBuilder
    Dim strCleanSearch As String
    Dim strCleanCat As String
    Dim strCleanStatus As String

    Set objQb = TasksRepo().NewQuery()

    strCleanStatus = Trim$(strStatusFilter)
    Select Case UCase$(strCleanStatus)
        Case "OPEN ONLY"
            objQb.WhereEq "IsCompleted", False
        Case "COMPLETED ONLY"
            objQb.WhereEq "IsCompleted", True
        Case "TRASH (SOFT-DELETED)"
            objQb.OnlyTrashed
        Case Else ' "Active" (all non-deleted rows)
            ' Default soft-delete scope excludes trashed rows automatically
    End Select

    strCleanCat = Trim$(strCategoryFilter)
    If Len(strCleanCat) > 0 And StrComp(strCleanCat, "(All)", vbTextCompare) <> 0 Then
        objQb.WhereEq "CategoryName", strCleanCat
    End If

    strCleanSearch = Trim$(strSearchText)
    If Len(strCleanSearch) > 0 Then
        objQb.WhereContains "TaskTitle", strCleanSearch
    End If

    Set BuildFilteredTasksQuery = objQb.OrderByCol("IsCompleted", False).OrderByCol("TaskID", False)
End Function

' Approach B: Builds a pure in-memory disconnected ADODB.Recordset with a virtual [IsSelected]
' checkbox column -- zero temp tables, zero front-end bloat!
Public Function QueryTasksAdoFromSQLite( _
    Optional ByVal strSearchText As String = vbNullString, _
    Optional ByVal strCategoryFilter As String = vbNullString, _
    Optional ByVal strStatusFilter As String = "Active") As Object

    Dim objQb As AccessSQLite.clsSQLiteQueryBuilder
    Set objQb = BuildFilteredTasksQuery(strSearchText, strCategoryFilter, strStatusFilter)
    Set QueryTasksAdoFromSQLite = objQb.ToAdoRecordset(True, "IsSelected")
End Function

Public Function BatchSetCompletedInSQLite(ByVal colTaskIds As Collection, ByVal blnCompleted As Boolean) As Long
    Dim objConn As AccessSQLite.clsSQLiteConnection
    Dim varId As Variant
    Dim dictUpd As Scripting.Dictionary
    Dim lngCount As Long

    If colTaskIds Is Nothing Then Exit Function
    If colTaskIds.Count = 0 Then Exit Function

    Set objConn = GetDemoConnection()
    objConn.BeginTrans
    On Error GoTo BatchErr

    Set dictUpd = AccessSQLite.NewTextDict()
    dictUpd("IsCompleted") = blnCompleted

    For Each varId In colTaskIds
        TasksRepo().UpdateRecord CLng(varId), dictUpd, False
        lngCount = lngCount + 1
    Next varId

    objConn.CommitTrans
    BatchSetCompletedInSQLite = lngCount
    Exit Function

BatchErr:
    Dim lngErr As Long
    Dim strDesc As String
    lngErr = Err.Number
    strDesc = Err.Description
    objConn.RollbackTrans
    Err.Raise lngErr, "modDemoDataService.BatchSetCompletedInSQLite", strDesc
End Function

' Approach A: Queries SQLite via clsSQLiteQueryBuilder and refreshes local tblTaskBuffer for frmTaskList.
Public Function SyncBufferFromSQLite( _
    Optional ByVal strSearchText As String = vbNullString, _
    Optional ByVal strCategoryFilter As String = vbNullString, _
    Optional ByVal strStatusFilter As String = "Active") As Long

    Dim objQb As AccessSQLite.clsSQLiteQueryBuilder
    Dim colRows As Collection
    Dim dictRow As Scripting.Dictionary
    Dim dbLocal As DAO.Database
    Dim rstBuf As DAO.Recordset

    Set objQb = BuildFilteredTasksQuery(strSearchText, strCategoryFilter, strStatusFilter)
    Set colRows = objQb.ToRowsRaw()

    On Error GoTo SyncError
    Set dbLocal = CurrentDb
    dbLocal.Execute "DELETE FROM [tblTaskBuffer];", dbFailOnError
    If colRows.Count > 0 Then
        Set rstBuf = dbLocal.OpenRecordset("tblTaskBuffer", dbOpenDynaset, dbAppendOnly)
        For Each dictRow In colRows
            rstBuf.AddNew
            rstBuf!TaskID = CLng(dictRow("TaskID"))
            rstBuf!TaskTitle = Left$(CStr(Nz(dictRow("TaskTitle"), vbNullString)), 150)
            rstBuf!CategoryName = Left$(CStr(Nz(dictRow("CategoryName"), "General")), 50)
            rstBuf!PriorityLevel = Left$(CStr(Nz(dictRow("PriorityLevel"), "Medium")), 30)
            rstBuf!EstimateCost = CCur(Nz(dictRow("EstimateCost"), 0@))
            If Not IsNull(dictRow("DueDate")) And Not IsEmpty(dictRow("DueDate")) Then
                rstBuf!DueDate = CDate(dictRow("DueDate"))
            End If
            rstBuf!IsCompleted = CBool(Nz(dictRow("IsCompleted"), False))
            rstBuf!UpdatedAtText = Left$(FormatTimestampDisplay(dictRow("UpdatedAt")), 40)
            rstBuf!IsDeletedRow = Not IsNull(dictRow("DeletedAt"))
            rstBuf.Update
        Next dictRow
        rstBuf.Close
        Set rstBuf = Nothing
    End If

    SyncBufferFromSQLite = colRows.Count
    Exit Function

SyncError:
    Dim lngErr As Long
    Dim strDesc As String
    lngErr = Err.Number
    strDesc = Err.Description
    On Error Resume Next
    If Not rstBuf Is Nothing Then rstBuf.Close
    On Error GoTo 0
    Err.Raise lngErr, "modDemoDataService.SyncBufferFromSQLite", strDesc
End Function

' Pushes an inline Continuous Form row edit directly back into SQLite via the ORM Repository.
Public Sub PushBufferRowToSQLite( _
    ByVal lngTaskId As Long, _
    ByVal strTitle As String, _
    ByVal strCategory As String, _
    ByVal strPriority As String, _
    ByVal curEstimate As Currency, _
    ByVal varDueDate As Variant, _
    ByVal blnCompleted As Boolean)

    Dim dictChanges As Scripting.Dictionary
    If lngTaskId <= 0 Then Exit Sub

    Set dictChanges = AccessSQLite.NewTextDict()
    dictChanges("TaskTitle") = Trim$(strTitle)
    dictChanges("CategoryName") = IIf(Len(Trim$(strCategory)) = 0, "General", Trim$(strCategory))
    dictChanges("PriorityLevel") = IIf(Len(Trim$(strPriority)) = 0, "Medium", Trim$(strPriority))
    dictChanges("EstimateCost") = curEstimate
    If IsNull(varDueDate) Or IsEmpty(varDueDate) Or Len(Trim$(CStr(Nz(varDueDate, vbNullString)))) = 0 Then
        dictChanges("DueDate") = Null
    Else
        dictChanges("DueDate") = CDate(varDueDate)
    End If
    dictChanges("IsCompleted") = blnCompleted

    TasksRepo().UpdateRecord lngTaskId, dictChanges, blnDirtyOnly:=True
End Sub

Public Sub ToggleTaskCompletedInSQLite(ByVal lngTaskId As Long)
    Dim dictCur As Scripting.Dictionary
    Dim dictUpd As Scripting.Dictionary
    If lngTaskId <= 0 Then Exit Sub
    Set dictCur = TasksRepo().FindByIdRaw(lngTaskId, True)
    If dictCur Is Nothing Then Exit Sub
    Set dictUpd = AccessSQLite.NewTextDict()
    dictUpd("IsCompleted") = Not CBool(Nz(dictCur("IsCompleted"), False))
    TasksRepo().UpdateRecord lngTaskId, dictUpd, False
End Sub

' Returns live SQLite engine & KPI metrics for the form header status banner.
Public Function GetSqliteSummaryBanner() As String
    Dim objConn As AccessSQLite.clsSQLiteConnection
    Dim dictStats As Scripting.Dictionary
    Dim lngActive As Long
    Dim lngOpen As Long
    Dim lngTrashed As Long
    Dim curOpenTotal As Currency

    Set objConn = GetDemoConnection()
    Set dictStats = objConn.QueryOneRow( _
        "SELECT " & _
        "  SUM(CASE WHEN DeletedAt IS NULL THEN 1 ELSE 0 END) AS ActiveCnt," & _
        "  SUM(CASE WHEN DeletedAt IS NULL AND IsCompleted = 0 THEN 1 ELSE 0 END) AS OpenCnt," & _
        "  SUM(CASE WHEN DeletedAt IS NOT NULL THEN 1 ELSE 0 END) AS TrashCnt," & _
        "  COALESCE(SUM(CASE WHEN DeletedAt IS NULL AND IsCompleted = 0 THEN EstimateCost ELSE 0 END), 0) AS OpenEst " & _
        "FROM Tasks;")

    If Not dictStats Is Nothing Then
        lngActive = CLng(Nz(dictStats("ActiveCnt"), 0))
        lngOpen = CLng(Nz(dictStats("OpenCnt"), 0))
        lngTrashed = CLng(Nz(dictStats("TrashCnt"), 0))
        curOpenTotal = CCur(Nz(dictStats("OpenEst"), 0@))
    End If

    GetSqliteSummaryBanner = _
        "Engine: SQLite " & objConn.Driver.LibraryVersion & " (" & objConn.Driver.DriverTitle & ")   |   " & _
        lngActive & " Active (" & lngOpen & " Open, " & Format$(curOpenTotal, "Currency") & " Open Est.)   |   " & _
        lngTrashed & " In Trash"
End Function

' Finds the First, Previous, Next, or Last active TaskID directly in SQLite for Single-Page navigation.
Public Function FindTargetSqliteTaskId( _
    ByVal strDirection As String, _
    Optional ByVal lngCurrentTaskId As Long = 0) As Long

    Dim objConn As AccessSQLite.clsSQLiteConnection
    Dim varId As Variant

    Set objConn = GetDemoConnection()
    Select Case UCase$(Trim$(strDirection))
        Case "FIRST"
            varId = objConn.QueryScalar("SELECT MIN(TaskID) FROM Tasks WHERE DeletedAt IS NULL;")
        Case "LAST"
            varId = objConn.QueryScalar("SELECT MAX(TaskID) FROM Tasks WHERE DeletedAt IS NULL;")
        Case "PREV"
            varId = objConn.QueryScalar("SELECT MAX(TaskID) FROM Tasks WHERE DeletedAt IS NULL AND TaskID < ?;", lngCurrentTaskId)
            If IsNull(varId) Then varId = lngCurrentTaskId
        Case "NEXT"
            varId = objConn.QueryScalar("SELECT MIN(TaskID) FROM Tasks WHERE DeletedAt IS NULL AND TaskID > ?;", lngCurrentTaskId)
            If IsNull(varId) Then varId = lngCurrentTaskId
    End Select

    If Not IsNull(varId) And Not IsEmpty(varId) Then
        FindTargetSqliteTaskId = CLng(varId)
    End If
End Function

Public Function FormatTimestampDisplay(ByVal varStamp As Variant) As String
    If IsNull(varStamp) Or IsEmpty(varStamp) Then
        FormatTimestampDisplay = "(not saved)"
    ElseIf VarType(varStamp) = vbDate Then
        FormatTimestampDisplay = Format$(CDate(varStamp), "yyyy-mm-dd hh:nn:ss")
    Else
        FormatTimestampDisplay = Replace(CStr(varStamp), "T", " ")
    End If
End Function

' Notifies both Continuous Forms (DAO frmTaskList and In-Memory ADO frmTaskListAdo) whenever data changes.
Public Sub NotifyTaskListChanged(Optional ByVal lngFocusTaskId As Long = 0)
    On Error Resume Next
    If CurrentProject.AllForms.Count > 0 Then
        If CurrentProject.AllForms("frmTaskList").IsLoaded Then
            Forms("frmTaskList").RefreshFromSQLite lngFocusTaskId
        End If
        If CurrentProject.AllForms("frmTaskListAdo").IsLoaded Then
            Forms("frmTaskListAdo").RefreshFromSQLite lngFocusTaskId
        End If
    End If
    On Error GoTo 0
End Sub

Public Function EnsureDemoProjectSetup() As String
    Dim objProj As Object
    Dim objComp As Object
    Dim objFrm As Object
    Dim strTargetFile As String
    On Error Resume Next
    strTargetFile = LCase$(CurrentProject.Name)
    For Each objProj In Application.VBE.VBProjects
        If Right$(LCase$(CStr(objProj.FileName)), Len(strTargetFile)) = strTargetFile Then
            If objProj.Name <> "AccessSQLiteDemo" Then
                objProj.Name = "AccessSQLiteDemo"
            End If
            For Each objComp In objProj.VBComponents
                If objComp.Type = 1 Or objComp.Type = 2 Then
                    DoCmd.Save acModule, objComp.Name
                End If
            Next objComp
            EnsureDemoProjectSetup = objProj.Name
            Exit For
        End If
    Next objProj

    For Each objFrm In CurrentProject.AllForms
        DoCmd.Save acForm, objFrm.Name
    Next objFrm

    DoCmd.RunCommand 126 ' acCmdCompileAndSaveAllModules
    On Error GoTo 0
End Function

