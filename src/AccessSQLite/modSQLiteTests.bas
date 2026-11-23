Attribute VB_Name = "modSQLiteTests"
Option Compare Database
Option Explicit

' =============================================================================
' modSQLiteTests
' Comprehensive automated test suite for AccessSQLite VBA Class Library & ORM.
' Run via Immediate Window: ? RunAllSQLiteTests()
' Or via SQL: SELECT RunAllSQLiteTests() AS TestSummary FROM tblDual;
' =============================================================================

Private m_Passed As Long
Private m_Failed As Long
Private m_FailureNotes As String

Public Function RunAllSQLiteTests() As String
    On Error GoTo UnexpectedFail
    Dim curStart As Currency
    Dim dblMs As Double

    m_Passed = 0
    m_Failed = 0
    m_FailureNotes = vbNullString
    curStart = StartPerfTimer()

    TestDriversAndPragmas
    TestDataTypesOnDriver sqDriverWinSqlite, "WinSQLite"
    TestDataTypesOnDriver sqDriverBundled, "Bundled"
    TestTransactionsAndSavepoints
    TestSchemaAndQueryBuilder
    TestOrmRepoEntityEagerAndUow
    TestMigratorAndAccessBridge
    TestAdoBridgeAndInMemoryRecordsets
    TestStudioHostHeadless

    dblMs = ElapsedPerfMs(curStart)
    If m_Failed = 0 Then
        RunAllSQLiteTests = "PASS | Assertions=" & m_Passed & " | Failed=0 | Elapsed=" & Format$(dblMs, "0.0") & " ms"
    Else
        RunAllSQLiteTests = "FAIL | Passed=" & m_Passed & " | Failed=" & m_Failed & " | Details: " & m_FailureNotes
    End If
    Exit Function

UnexpectedFail:
    RunAllSQLiteTests = "UNHANDLED_ERR #" & Err.Number & " (" & Err.Source & "): " & Err.Description & _
                        " | PassedSoFar=" & m_Passed & " | Notes=" & m_FailureNotes
End Function

Private Sub AssertTrue(ByVal blnCond As Boolean, ByVal strLabel As String)
    If blnCond Then
        m_Passed = m_Passed + 1
    Else
        m_Failed = m_Failed + 1
        If Len(m_FailureNotes) > 0 Then m_FailureNotes = m_FailureNotes & " ; "
        m_FailureNotes = m_FailureNotes & strLabel
    End If
End Sub

' -----------------------------------------------------------------------------
' 1. Drivers & Pragma Profiles
' -----------------------------------------------------------------------------
Private Sub TestDriversAndPragmas()
    Dim dbWin As clsSQLiteConnection
    Dim dbBundled As clsSQLiteConnection
    Dim dbFile As clsSQLiteConnection
    Dim strTestDbPath As String
    Dim strJournal As String
    Dim lngFk As Long

    Set dbWin = OpenSQLiteMemory(sqDriverWinSqlite)
    AssertTrue Left$(dbWin.Driver.LibraryVersion, 2) = "3.", "WinSQLite version 3.x"
    AssertTrue CLng(dbWin.QueryScalar("SELECT 20 + 22;")) = 42, "WinSQLite scalar 42"
    dbWin.CloseConnection

    AssertTrue LCase$(Right$(GetBundledDllDirectory(), 4)) = "\lib", "Bundled DLL directory resolves to \lib"
    Set dbBundled = OpenSQLiteMemory(sqDriverBundled)
    AssertTrue Left$(dbBundled.Driver.LibraryVersion, 2) = "3.", "Bundled version 3.x"
    AssertTrue CLng(dbBundled.QueryScalar("SELECT 19 + 23;")) = 42, "Bundled scalar 42"
    dbBundled.CloseConnection

    strTestDbPath = GetLibraryDirectory() & "\AccessSQLite_Test.db"
    Set dbFile = OpenSQLite(strTestDbPath, sqProfileAuto, sqDriverAuto)
    AssertTrue dbFile.ActivePragmaProfile = sqProfileNetworkSafe, "Auto profile selects NetworkSafe on SMB drive"
    strJournal = CStr(dbFile.QueryScalar("PRAGMA journal_mode;"))
    lngFk = CLng(dbFile.QueryScalar("PRAGMA foreign_keys;"))
    AssertTrue StrComp(strJournal, "delete", vbTextCompare) = 0, "SMB journal_mode is DELETE"
    AssertTrue lngFk = 1, "foreign_keys is ON"
    AssertTrue dbFile.IntegrityCheck(), "IntegrityCheck ok on AccessSQLite_Test.db"
    dbFile.CloseConnection
End Sub

' -----------------------------------------------------------------------------
' 2. Data Types, Empty-String vs Null, 64-bit Ints, 0-byte BLOBs & Named Params
' -----------------------------------------------------------------------------
Private Sub TestDataTypesOnDriver(ByVal enmDriver As SQLiteDriverKind, ByVal strTag As String)
    Dim objDb As clsSQLiteConnection
    Dim strLongMemo As String
    Dim abytZero() As Byte
    Dim abytSample(0 To 3) As Byte
    Dim abytOutZero() As Byte
    Dim abytOutSample() As Byte
    Dim dictRow As Scripting.Dictionary
    Dim dtmSample As Date

    Set objDb = OpenSQLiteMemory(enmDriver)
    objDb.ExecuteScript _
        "CREATE TABLE TypeCheck (" & _
        "  ItemID INTEGER PRIMARY KEY AUTOINCREMENT," & _
        "  ReqEmptyText TEXT NOT NULL," & _
        "  OptNullText TEXT," & _
        "  UnicodeText TEXT," & _
        "  BigMemo TEXT," & _
        "  BigIntVal INTEGER," & _
        "  AmountVal CURRENCY," & _
        "  FlagVal BOOLEAN," & _
        "  EventAt DATETIME," & _
        "  ZeroBlob BLOB," & _
        "  DataBlob BLOB" & _
        ");"

    strLongMemo = String$(12000, "M") & "-END"
    abytZero = StrConv(vbNullString, vbFromUnicode)
    abytSample(0) = 0: abytSample(1) = 1: abytSample(2) = 127: abytSample(3) = 255
    dtmSample = DateSerial(2026, 11, 20) + TimeSerial(10, 15, 30)

    objDb.ExecuteSql _
        "INSERT INTO TypeCheck (ReqEmptyText, OptNullText, UnicodeText, BigMemo, BigIntVal, AmountVal, FlagVal, EventAt, ZeroBlob, DataBlob) " & _
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?);", _
        vbNullString, Null, "Cafe - " & ChrW$(228) & ChrW$(246) & ChrW$(252) & " - " & ChrW$(26085) & ChrW$(26412), _
        strLongMemo, 5000000007@, 123456.7895@, True, dtmSample, abytZero, abytSample

    AssertTrue CLng(objDb.LastInsertRowId) = 1, strTag & ": LastInsertRowId = 1"

    Set dictRow = objDb.QueryOneRowRaw("SELECT * FROM TypeCheck WHERE ItemID = @id;", "@id", 1)
    AssertTrue Not dictRow Is Nothing, strTag & ": fetched row"
    AssertTrue Not IsNull(dictRow("ReqEmptyText")) And CStr(dictRow("ReqEmptyText")) = vbNullString, strTag & ": empty string preserved as '' not NULL"
    AssertTrue IsNull(dictRow("OptNullText")), strTag & ": Null preserved as NULL"
    AssertTrue InStr(1, CStr(dictRow("UnicodeText")), ChrW$(26085), vbBinaryCompare) > 0, strTag & ": Unicode UTF-16 preserved"
    AssertTrue Len(CStr(dictRow("BigMemo"))) = 12004 And Right$(CStr(dictRow("BigMemo")), 4) = "-END", strTag & ": 12KB Memo preserved"
    AssertTrue CCur(dictRow("BigIntVal")) = 5000000007@, strTag & ": 64-bit integer 5000000007 preserved"
    AssertTrue CCur(dictRow("AmountVal")) = 123456.7895@, strTag & ": Currency 4-decimal preserved"
    AssertTrue VarType(dictRow("FlagVal")) = vbBoolean And CBool(dictRow("FlagVal")) = True, strTag & ": Boolean True hydrated"
    AssertTrue VarType(dictRow("EventAt")) = vbDate And CDate(dictRow("EventAt")) = dtmSample, strTag & ": DateTime hydrated"

    abytOutZero = dictRow("ZeroBlob")
    abytOutSample = dictRow("DataBlob")
    AssertTrue Not IsNull(dictRow("ZeroBlob")) And GetByteArrayLength(abytOutZero) = 0, strTag & ": 0-byte BLOB round-tripped"
    AssertTrue GetByteArrayLength(abytOutSample) = 4 And abytOutSample(3) = 255, strTag & ": 4-byte BLOB round-tripped"

    ' Positional parameter starting with '@' must NOT be confused with a named parameter
    AssertTrue CStr(objDb.QueryScalar("SELECT ? || '|' || ?;", "@literalUser", "ok")) = "@literalUser|ok", _
        strTag & ": positional @literal disambiguated"

    objDb.CloseConnection
End Sub

' -----------------------------------------------------------------------------
' 3. Transactions & Nested Savepoints
' -----------------------------------------------------------------------------
Private Sub TestTransactionsAndSavepoints()
    Dim objDb As clsSQLiteConnection
    Set objDb = OpenSQLiteMemory()
    objDb.ExecuteScript "CREATE TABLE Ledger (EntryNo INTEGER PRIMARY KEY, LabelText TEXT);"

    objDb.BeginTrans
    AssertTrue objDb.TransactionDepth = 1, "Outer tx depth = 1"
    objDb.ExecuteSql "INSERT INTO Ledger (EntryNo, LabelText) VALUES (?, ?);", 1, "KeepOuter"

    objDb.BeginTrans
    AssertTrue objDb.TransactionDepth = 2, "Nested savepoint depth = 2"
    objDb.ExecuteSql "INSERT INTO Ledger (EntryNo, LabelText) VALUES (?, ?);", 2, "DiscardInner"
    objDb.RollbackTrans
    AssertTrue objDb.TransactionDepth = 1, "After inner rollback depth = 1"

    objDb.BeginTrans
    objDb.ExecuteSql "INSERT INTO Ledger (EntryNo, LabelText) VALUES (?, ?);", 3, "KeepInner2"
    objDb.CommitTrans
    objDb.CommitTrans
    AssertTrue Not objDb.InTransaction, "All transactions committed"

    AssertTrue CLng(objDb.QueryScalar("SELECT COUNT(*) FROM Ledger;")) = 2, "Savepoint kept 2 rows"
    AssertTrue CLng(objDb.QueryScalar("SELECT COUNT(*) FROM Ledger WHERE EntryNo = 2;")) = 0, "Rolled-back savepoint row absent"

    objDb.BeginTrans
    objDb.ExecuteSql "INSERT INTO Ledger (EntryNo, LabelText) VALUES (?, ?);", 99, "RollbackAll"
    objDb.RollbackTrans
    AssertTrue CLng(objDb.QueryScalar("SELECT COUNT(*) FROM Ledger WHERE EntryNo = 99;")) = 0, "Full rollback discarded row 99"

    objDb.CloseConnection
End Sub

' -----------------------------------------------------------------------------
' 4. Schema DDL, STRICT Tables, Additive Sync & Fluent Query Builder
' -----------------------------------------------------------------------------
Private Sub TestSchemaAndQueryBuilder()
    Dim objDb As clsSQLiteConnection
    Dim mapProj As clsSQLiteTableMap
    Dim lngAddedCols As Long
    Dim colPage As Collection
    Dim colCodes As Collection

    Set objDb = OpenSQLiteMemory()

    Set mapProj = objDb.DefineTable("Projects", "ProjectID") _
        .SetStrict(True) _
        .AddPrimaryKey("ProjectID", sqColInt64, True) _
        .AddColumn("ProjectCode", sqColText, "Project code", blnRequired:=True, lngMaxLength:=20, blnUnique:=True) _
        .AddColumn("ProjectTitle", sqColText, "Project title", blnRequired:=True, lngMaxLength:=100) _
        .AddColumn("BudgetAmount", sqColCurrency, "Budget", varDefault:=0@) _
        .AddIndex("idx_Projects_Budget", "BudgetAmount DESC")

    objDb.SchemaManager().EnsureTable mapProj

    ' Add a new column to the map and verify SyncAdditiveColumns applies ALTER TABLE ADD COLUMN
    mapProj.AddColumn "OwnerDept", sqColText, "Department", varDefault:="R&D"
    lngAddedCols = objDb.SchemaManager().SyncAdditiveColumns(mapProj)
    AssertTrue lngAddedCols = 1, "SyncAdditiveColumns added OwnerDept"

    objDb.ExecuteSql "INSERT INTO Projects (ProjectCode, ProjectTitle, BudgetAmount) VALUES (?, ?, ?);", "PRJ-01", "Alpha_100% Core", 15000@
    objDb.ExecuteSql "INSERT INTO Projects (ProjectCode, ProjectTitle, BudgetAmount) VALUES (?, ?, ?);", "PRJ-02", "Beta Analytics", 35000@
    objDb.ExecuteSql "INSERT INTO Projects (ProjectCode, ProjectTitle, BudgetAmount) VALUES (?, ?, ?);", "PRJ-03", "Gamma Portal", 55000@
    objDb.ExecuteSql "INSERT INTO Projects (ProjectCode, ProjectTitle, BudgetAmount) VALUES (?, ?, ?);", "PRJ-04", "Delta Warehouse", 75000@

    AssertTrue objDb.FromTable("Projects").WhereContains("ProjectTitle", "100%").CountRows() = 1, "WhereContains escapes % wildcard"
    AssertTrue objDb.FromTable("Projects").WhereIn("ProjectCode", New Collection).CountRows() = 0, "Empty WhereIn returns 0 rows safely"
    AssertTrue objDb.FromTable("Projects").WhereBetween("BudgetAmount", 20000@, 60000@).CountRows() = 2, "WhereBetween matched 2 rows"
    AssertTrue objDb.FromTable("Projects").WhereEq("ProjectCode", "PRJ-03").ExistsAny(), "ExistsAny True"

    Set colPage = objDb.FromTable("Projects").OrderByCol("ProjectID").Page(2, 2).ToRows()
    AssertTrue colPage.Count = 2 And CStr(colPage(1)("ProjectCode")) = "PRJ-03", "Page(2, 2) returned PRJ-03..04"

    Set colCodes = objDb.FromTable("Projects").WhereOp("BudgetAmount", ">=", 50000@).OrderByDesc("BudgetAmount").PluckCol("ProjectCode")
    AssertTrue colCodes.Count = 2 And CStr(colCodes(1)) = "PRJ-04", "PluckCol returned ordered values"

    objDb.CloseConnection
End Sub

' -----------------------------------------------------------------------------
' 5. ORM Repository, Typed Entity, Soft Deletes, Eager Loading & Unit of Work
' -----------------------------------------------------------------------------
Private Sub TestOrmRepoEntityEagerAndUow()
    Dim objDb As clsSQLiteConnection
    Dim mapProj As clsSQLiteTableMap
    Dim mapTask As clsSQLiteTableMap
    Dim repoProj As clsSQLiteRepository
    Dim repoTask As clsSQLiteRepository
    Dim dictP1 As Scripting.Dictionary
    Dim dictP2 As Scripting.Dictionary
    Dim lngP1 As Long
    Dim lngP2 As Long
    Dim objTask As clsDemoTask
    Dim objLoaded As clsDemoTask
    Dim lngBeforeQueries As Long
    Dim blnCaughtValidation As Boolean
    Dim colProjects As Collection
    Dim colSubTasks As Collection
    Dim objUow As clsSQLiteUnitOfWork
    Dim blnCaughtConstraint As Boolean
    Dim strGenCode As String

    Set objDb = OpenSQLiteMemory()

    Set mapProj = objDb.DefineTable("Projects", "ProjectID") _
        .AddPrimaryKey("ProjectID", sqColInt64, True) _
        .AddColumn("ProjectCode", sqColText, "Project code", blnRequired:=True, lngMaxLength:=20, blnUnique:=True) _
        .AddColumn("ProjectTitle", sqColText, "Project title", blnRequired:=True, lngMaxLength:=100) _
        .AddTimestamps("CreatedAt", "UpdatedAt")

    Set mapTask = objDb.DefineTable("DemoTasks", "TaskID") _
        .AddPrimaryKey("TaskID", sqColInt64, True) _
        .AddColumn("ProjectID", sqColLong, "Project", blnRequired:=True, strFkTable:="Projects", strFkColumn:="ProjectID", strOnDelete:="CASCADE") _
        .AddColumn("TitleText", sqColText, "Task title", blnRequired:=True, lngMaxLength:=80) _
        .AddColumn("PriorityRank", sqColLong, "Priority", blnRequired:=True, varDefault:=3) _
        .AddColumn("EstimateCost", sqColCurrency, "Estimate cost", blnRequired:=True, varDefault:=0@) _
        .AddColumn("IsCompleted", sqColBoolean, "Completed", blnRequired:=True, varDefault:=False) _
        .AddColumn("DueDate", sqColDate, "Due date") _
        .AddTimestamps("CreatedAt", "UpdatedAt") _
        .AddSoftDelete("DeletedAt") _
        .AddIndex("idx_DemoTasks_ProjectID", "ProjectID")

    objDb.SchemaManager().EnsureTable mapProj
    objDb.SchemaManager().EnsureTable mapTask

    Set repoProj = objDb.Repository("Projects")
    Set repoTask = objDb.Repository("DemoTasks")

    ' Upsert test on Projects
    Set dictP1 = NewTextDict()
    dictP1("ProjectCode") = "ORM-1": dictP1("ProjectTitle") = "Initial Title"
    lngP1 = CLng(repoProj.UpsertRecord("ProjectCode", dictP1))
    dictP1("ProjectTitle") = "Updated Via Upsert"
    repoProj.UpsertRecord "ProjectCode", dictP1
    AssertTrue repoProj.CountAll() = 1, "Upsert kept single Project row"
    AssertTrue CStr(repoProj.FindById(lngP1)("ProjectTitle")) = "Updated Via Upsert", "Upsert updated ProjectTitle"

    Set dictP2 = NewTextDict()
    dictP2("ProjectCode") = "ORM-2": dictP2("ProjectTitle") = "Second Project"
    lngP2 = CLng(repoProj.InsertRecord(dictP2))

    ' Validation error test (missing TitleText)
    On Error Resume Next
    Err.Clear
    Dim dictBad As Scripting.Dictionary
    Set dictBad = NewTextDict()
    dictBad("ProjectID") = lngP1
    repoTask.InsertRecord dictBad
    blnCaughtValidation = (Err.Number = sqErrValidation)
    On Error GoTo 0
    AssertTrue blnCaughtValidation, "Missing required column raises sqErrValidation"

    ' Strongly-typed ISQLiteEntity CRUD & No-Op Dirty Check
    Set objTask = New clsDemoTask
    objTask.ProjectID = lngP1
    objTask.TitleText = "Build SQLite ORM"
    objTask.PriorityRank = 1
    objTask.EstimateCost = 1250.5@
    objTask.IsCompleted = False
    objTask.DueDate = DateSerial(2026, 12, 1)

    repoTask.SaveEntity objTask
    AssertTrue objTask.TaskID > 0, "SaveEntity populated generated TaskID"
    AssertTrue Not IsNull(objTask.CreatedAt), "SaveEntity populated CreatedAt"

    ' Saving unchanged entity with blnDirtyOnly=True executes only 1 diff SELECT and 0 UPDATEs
    Dim lngTotalChangesBefore As Long
    lngTotalChangesBefore = objDb.TotalChangesCount
    objDb.ResetExecutedQueryCount
    repoTask.SaveEntity objTask, True
    lngBeforeQueries = objDb.ExecutedQueryCount
    AssertTrue lngBeforeQueries = 1 And objDb.TotalChangesCount = lngTotalChangesBefore, "Unchanged SaveEntity skipped UPDATE statement"

    ' Mutate entity & save
    objTask.IsCompleted = True
    objTask.TitleText = "Build SQLite ORM (Done)"
    repoTask.SaveEntity objTask, True

    Set objLoaded = repoTask.FindEntity(objTask.TaskID, New clsDemoTask)
    AssertTrue Not objLoaded Is Nothing, "FindEntity hydrated clsDemoTask"
    AssertTrue objLoaded.IsCompleted = True And objLoaded.EstimateCost = 1250.5@, "Hydrated entity matches saved fields"

    ' Add two more tasks for eager-loading & soft-delete tests
    Dim objTask2 As clsDemoTask
    Set objTask2 = New clsDemoTask
    objTask2.ProjectID = lngP1: objTask2.TitleText = "Write Unit Tests": objTask2.PriorityRank = 2: objTask2.EstimateCost = 400@
    repoTask.SaveEntity objTask2

    Dim objTask3 As clsDemoTask
    Set objTask3 = New clsDemoTask
    objTask3.ProjectID = lngP2: objTask3.TitleText = "Deploy Package": objTask3.PriorityRank = 1: objTask3.EstimateCost = 300@
    repoTask.SaveEntity objTask3

    ' Soft delete & restore
    repoTask.DeleteEntity objTask2, False
    AssertTrue repoTask.CountAll(False) = 2, "Soft-deleted task hidden from active CountAll"
    AssertTrue repoTask.CountAll(True) = 3, "CountAll(includeTrashed:=True) sees 3 tasks"
    AssertTrue repoTask.NewQuery().OnlyTrashed().CountRows() = 1, "OnlyTrashed finds 1 soft-deleted task"
    repoTask.RestoreRecord objTask2.TaskID
    AssertTrue repoTask.CountAll(False) = 3, "RestoreRecord restored soft-deleted task"

    ' Batch Eager Loading (IncludeHasMany must run in exactly 2 SQL queries, NOT N+1!)
    objDb.ResetExecutedQueryCount
    Set colProjects = repoProj.ListAllRaw("ProjectID")
    repoProj.IncludeHasMany colProjects, "DemoTasks", "ProjectID", "Tasks", "ProjectID", "PriorityRank"
    AssertTrue objDb.ExecutedQueryCount = 2, "Eager IncludeHasMany executed exactly 2 queries (zero N+1)"
    Set colSubTasks = colProjects(1)("Tasks")
    AssertTrue colSubTasks.Count = 2 And colProjects(2)("Tasks").Count = 1, "IncludeHasMany attached 2 and 1 child tasks"

    ' Unit of Work atomic rollback on failure + atomic commit on success
    Set objUow = objDb.NewUnitOfWork()
    Dim dictP3 As Scripting.Dictionary
    Set dictP3 = NewTextDict()
    dictP3("ProjectCode") = "ORM-3": dictP3("ProjectTitle") = "Will Rollback"
    objUow.RegisterInsert "Projects", dictP3
    ' Duplicate ProjectCode "ORM-1" triggers UNIQUE constraint rollback of entire UoW!
    Dim dictDup As Scripting.Dictionary
    Set dictDup = NewTextDict()
    dictDup("ProjectCode") = "ORM-1": dictDup("ProjectTitle") = "Duplicate Code"
    objUow.RegisterInsert "Projects", dictDup

    On Error Resume Next
    Err.Clear
    objUow.CommitWork
    blnCaughtConstraint = (Err.Number = sqErrConstraint)
    On Error GoTo 0
    AssertTrue blnCaughtConstraint, "UnitOfWork raised sqErrConstraint on duplicate key"
    AssertTrue repoProj.CountAll() = 2, "UnitOfWork rolled back partial insert ORM-3"

    objUow.ClearPending
    objUow.RegisterInsert "Projects", dictP3
    AssertTrue objUow.CommitWork() = 1 And repoProj.CountAll() = 3, "Valid UnitOfWork committed ORM-3"

    ' Entity Code Generator reserved-word protection check
    objDb.ExecuteScript "CREATE TABLE LegacyAudit ([Date] TEXT, [Value] REAL, [Select] TEXT);"
    strGenCode = GenerateEntityClassCode(objDb, "LegacyAudit")
    AssertTrue InStr(1, strGenCode, "Property Get DateValue()", vbBinaryCompare) > 0 And _
               InStr(1, strGenCode, "Property Get ValueValue()", vbBinaryCompare) > 0 And _
               InStr(1, strGenCode, "Implements ISQLiteEntity", vbBinaryCompare) > 0, _
               "Entity generator sanitizes reserved words Date/Value/Select"

    objDb.CloseConnection
End Sub

' -----------------------------------------------------------------------------
' 6. Migrator, Online Backup & Access DAO <-> SQLite Bridge
' -----------------------------------------------------------------------------
Private Sub TestMigratorAndAccessBridge()
    Dim strDbPath As String
    Dim strBackupPath As String
    Dim objDb As clsSQLiteConnection
    Dim objMig As clsSQLiteMigrator
    Dim objBridge As clsSQLiteAccessBridge
    Dim tdfSrc As DAO.TableDef
    Dim rstSrc As DAO.Recordset
    Dim strSrcTable As String
    Dim strDstTable As String
    Dim strEmptyTable As String
    Dim rstImported As DAO.Recordset

    strDbPath = GetLibraryDirectory() & "\AccessSQLite_Test.db"
    strBackupPath = Environ$("TEMP") & "\AccessSQLite_BackupCheck.db"
    If Len(Dir$(strBackupPath)) > 0 Then Kill strBackupPath

    Set objDb = OpenSQLite(strDbPath)
    objDb.ExecuteScript _
        "DROP TABLE IF EXISTS BridgeWidgets; " & _
        "DROP TABLE IF EXISTS MigItems; " & _
        "DROP TABLE IF EXISTS MigCategories; " & _
        "DROP TABLE IF EXISTS _schema_migrations;"

    Set objMig = objDb.NewMigrator()
    objMig.AddMigration 1, "Create MigCategories", _
        "CREATE TABLE MigCategories (CatID INTEGER PRIMARY KEY, CatTitle TEXT NOT NULL);", _
        "DROP TABLE MigCategories;"
    objMig.AddMigration 2, "Create MigItems", _
        "CREATE TABLE MigItems (ItemID INTEGER PRIMARY KEY, CatID INTEGER REFERENCES MigCategories(CatID), ItemTitle TEXT NOT NULL);", _
        "DROP TABLE MigItems;"

    AssertTrue objMig.MigrateUp() = 2 And objMig.CurrentVersion() = 2, "Migrator applied v1..v2"
    AssertTrue objMig.RollbackLast(1) = 1 And objMig.CurrentVersion() = 1, "Migrator rolled back v2 -> v1"
    AssertTrue Not objDb.TableExists("MigItems"), "MigItems dropped on rollback"
    AssertTrue objMig.MigrateUp() = 1 And objMig.CurrentVersion() = 2, "Migrator re-applied v2"

    ' Access DAO -> SQLite -> Access DAO round-trip (testing Null-binding reset + >255 char Memo!)
    strSrcTable = "tblTmpBridgeSource"
    strDstTable = "tblTmpBridgeImported"
    strEmptyTable = "tblTmpBridgeEmpty"
    DeleteAccessTableQuiet strSrcTable
    DeleteAccessTableQuiet strDstTable
    DeleteAccessTableQuiet strEmptyTable

    Set tdfSrc = CurrentDb.CreateTableDef(strSrcTable)
    tdfSrc.Fields.Append tdfSrc.CreateField("WidgetID", dbLong)
    tdfSrc.Fields.Append tdfSrc.CreateField("WidgetLabel", dbText, 100)
    tdfSrc.Fields.Append tdfSrc.CreateField("OptionalNote", dbMemo)
    tdfSrc.Fields.Append tdfSrc.CreateField("UnitPrice", dbCurrency)
    tdfSrc.Fields.Append tdfSrc.CreateField("IsActive", dbBoolean)
    CurrentDb.TableDefs.Append tdfSrc
    CurrentDb.TableDefs.Refresh

    Set rstSrc = CurrentDb.OpenRecordset(strSrcTable, dbOpenDynaset)
    rstSrc.AddNew
    rstSrc!WidgetID = 10
    rstSrc!WidgetLabel = "First Widget"
    rstSrc!OptionalNote = String$(400, "A") ' > 255 chars!
    rstSrc!UnitPrice = 49.95@
    rstSrc!IsActive = True
    rstSrc.Update

    rstSrc.AddNew
    rstSrc!WidgetID = 20
    rstSrc!WidgetLabel = "Second Widget"
    rstSrc!OptionalNote = Null ' Must NOT inherit row 1's 400-char string!
    rstSrc!UnitPrice = 19.5@
    rstSrc!IsActive = False
    rstSrc.Update
    rstSrc.Close

    Set objBridge = objDb.NewAccessBridge()
    AssertTrue objBridge.ExportAccessToSQLite(strSrcTable, "BridgeWidgets", True) = 2, "ExportAccessToSQLite exported 2 rows"
    AssertTrue IsNull(objDb.QueryScalar("SELECT OptionalNote FROM BridgeWidgets WHERE WidgetID = 20;")), _
        "ExportAccessToSQLite cleared bindings so row 2 OptionalNote is NULL"

    AssertTrue objBridge.ImportSQLiteToAccess("BridgeWidgets", strDstTable, True) = 2, "ImportSQLiteToAccess imported 2 rows"
    AssertTrue objBridge.ImportQueryToAccess("SELECT * FROM BridgeWidgets WHERE WidgetID > ?;", strEmptyTable, True, 999) = 0, _
        "ImportQueryToAccess 0-row query succeeded"
    AssertTrue AccessTableExistsInCurrentDb(strEmptyTable), "0-row import still created Access TableDef schema"

    Set rstImported = CurrentDb.OpenRecordset("SELECT * FROM [" & strDstTable & "] ORDER BY WidgetID;", dbOpenSnapshot)
    AssertTrue Len(CStr(rstImported!OptionalNote)) = 400 And CCur(rstImported!UnitPrice) = 49.95@, "Row 1 400-char Memo & Currency imported"
    rstImported.MoveNext
    AssertTrue IsNull(rstImported!OptionalNote) And CBool(rstImported!IsActive) = False, "Row 2 Null & Boolean imported"
    rstImported.Close

    DeleteAccessTableQuiet strSrcTable
    DeleteAccessTableQuiet strDstTable
    DeleteAccessTableQuiet strEmptyTable

    ' Online Backup verification
    objDb.BackupToFile strBackupPath
    objDb.CloseConnection

    Dim objBck As clsSQLiteConnection
    Set objBck = OpenSQLite(strBackupPath, sqProfileReadOnly)
    AssertTrue CLng(objBck.QueryScalar("SELECT COUNT(*) FROM BridgeWidgets;")) = 2, "Online backup preserved BridgeWidgets"
    objBck.CloseConnection
    If Len(Dir$(strBackupPath)) > 0 Then Kill strBackupPath
End Sub

Private Sub DeleteAccessTableQuiet(ByVal strTbl As String)
    On Error Resume Next
    CurrentDb.TableDefs.Delete strTbl
    CurrentDb.TableDefs.Refresh
    On Error GoTo 0
End Sub

Private Function AccessTableExistsInCurrentDb(ByVal strTbl As String) As Boolean
    Dim tdf As DAO.TableDef
    For Each tdf In CurrentDb.TableDefs
        If StrComp(tdf.Name, strTbl, vbTextCompare) = 0 Then
            AccessTableExistsInCurrentDb = True
            Exit Function
        End If
    Next tdf
End Function


' -----------------------------------------------------------------------------
' 7. In-Memory Disconnected ADODB.Recordset Bridge & Virtual Checkbox Selection
' -----------------------------------------------------------------------------
Private Sub TestAdoBridgeAndInMemoryRecordsets()
    Dim objDb As clsSQLiteConnection
    Dim mapItems As clsSQLiteTableMap
    Dim repoItems As clsSQLiteRepository
    Dim objAdoBridge As clsSQLiteAdoBridge
    Dim objRs As Object
    Dim dictSeed As Scripting.Dictionary
    Dim colSelected As Collection
    Dim lngNewId As Long

    Set objDb = OpenSQLiteMemory()
    Set mapItems = objDb.DefineTable("AdoTasks", "TaskID") _
        .AddPrimaryKey("TaskID", sqColInt64, True) _
        .AddColumn("TaskTitle", sqColText, "Title", blnRequired:=True, lngMaxLength:=120) _
        .AddColumn("BudgetCost", sqColCurrency, "Budget", blnRequired:=True, varDefault:=0@) _
        .AddColumn("DueDate", sqColDate, "Due date") _
        .AddColumn("IsDone", sqColBoolean, "Done", blnRequired:=True, varDefault:=False) _
        .AddColumn("LongNotes", sqColMemo, "Notes") _
        .AddTimestamps("CreatedAt", "UpdatedAt")
    objDb.SchemaManager().EnsureTable mapItems
    Set repoItems = objDb.Repository("AdoTasks")

    Set dictSeed = NewTextDict()
    dictSeed("TaskTitle") = "First ADO Row": dictSeed("BudgetCost") = 250.5@
    dictSeed("DueDate") = DateSerial(2026, 12, 10): dictSeed("IsDone") = False
    dictSeed("LongNotes") = String$(350, "N")
    repoItems.InsertRecord dictSeed

    Set dictSeed = NewTextDict()
    dictSeed("TaskTitle") = "Second ADO Row": dictSeed("BudgetCost") = 750.25@
    dictSeed("DueDate") = DateSerial(2026, 12, 20): dictSeed("IsDone") = True
    dictSeed("LongNotes") = "Short note"
    repoItems.InsertRecord dictSeed

    ' Build disconnected ADODB.Recordset with virtual IsSelected checkbox column
    Set objRs = repoItems.NewQuery().OrderByCol("TaskID").ToAdoRecordset(True, "IsSelected")
    AssertTrue Not objRs Is Nothing, "ToAdoRecordset returned ADODB.Recordset"
    AssertTrue objRs.RecordCount = 2, "ADO RecordCount = 2 without MoveLast"
    AssertTrue CBool(objRs.Fields.Item("IsSelected").Value) = False, "Virtual IsSelected starts False"
    AssertTrue Len(CStr(objRs.Fields.Item("LongNotes").Value)) = 350, "ADO >255-char Memo preserved"
    AssertTrue CCur(objRs.Fields.Item("BudgetCost").Value) = 250.5@, "ADO Currency preserved"
    AssertTrue CDate(objRs.Fields.Item("DueDate").Value) = DateSerial(2026, 12, 10), "ADO Date hydrated"

    ' Toggle virtual IsSelected on row 1 & edit BudgetCost inline -> write back to SQLite
    objRs.Fields.Item("IsSelected").Value = True
    objRs.Fields.Item("BudgetCost").Value = 300.75@
    objRs.Update
    repoItems.SaveFromAdoRow objRs, True, "IsSelected"

    AssertTrue CCur(objDb.QueryScalar("SELECT BudgetCost FROM AdoTasks WHERE TaskID = 1;")) = 300.75@, _
        "SaveFromAdoRow persisted edited BudgetCost to SQLite"

    ' Add a brand-new row directly in the in-memory ADODB.Recordset and save to SQLite
    objRs.AddNew
    objRs.Fields.Item("IsSelected").Value = True
    objRs.Fields.Item("TaskID").Value = 0
    objRs.Fields.Item("TaskTitle").Value = "Inserted Via Disconnected ADO"
    objRs.Fields.Item("BudgetCost").Value = 999@
    objRs.Fields.Item("IsDone").Value = False
    objRs.Update
    lngNewId = CLng(repoItems.SaveFromAdoRow(objRs, True, "IsSelected"))
    AssertTrue lngNewId = 3 And CLng(objRs.Fields.Item("TaskID").Value) = 3, _
        "SaveFromAdoRow inserted new ADO row and hydrated TaskID=3 back onto ADO recordset"

    ' Verify GetSelectedKeys via Clone (rows 1 and 3 were checked)
    Set objAdoBridge = objDb.NewAdoBridge()
    Set colSelected = objAdoBridge.GetSelectedKeys(objRs, "TaskID", "IsSelected")
    AssertTrue colSelected.Count = 2 And CLng(colSelected(1)) = 1 And CLng(colSelected(2)) = 3, _
        "GetSelectedKeys returned checked rows [1, 3]"

    objAdoBridge.SetAllSelected objRs, False, "IsSelected"
    AssertTrue objAdoBridge.GetSelectedKeys(objRs, "TaskID", "IsSelected").Count = 0, _
        "SetAllSelected(False) cleared virtual checkboxes"

    ' Import disconnected ADODB.Recordset into a brand-new SQLite table (excluding IsSelected)
    AssertTrue objAdoBridge.ImportAdoRecordsetToSQLite(objRs, "AdoTasksCopy", True, "IsSelected") = 3, _
        "ImportAdoRecordsetToSQLite imported 3 rows"
    AssertTrue CLng(objDb.QueryScalar("SELECT COUNT(*) FROM AdoTasksCopy;")) = 3, _
        "AdoTasksCopy row count = 3"

    objRs.Close
    objDb.CloseConnection
End Sub




' -----------------------------------------------------------------------------
' 8. Monaco SQLite Studio Host Bridge (Schema, Stored Queries, Multi-SQL, Explain)
' -----------------------------------------------------------------------------
Private Sub TestStudioHostHeadless()
    Dim strReply As String

    ResetHeadlessSQLiteStudio
    strReply = SQLiteStudioDispatchHeadless(":memory:", "{""id"":1,""action"":""studio.init"",""payload"":{}}")
    AssertTrue InStr(1, strReply, """ok"":true", vbBinaryCompare) > 0, "Studio studio.init ok"

    strReply = SQLiteStudioDispatchHeadless("", _
        "{""id"":2,""action"":""studio.execSql"",""payload"":{""sql"":""CREATE TABLE StudioItems (ItemID INTEGER PRIMARY KEY, ItemName TEXT NOT NULL); INSERT INTO StudioItems (ItemName) VALUES ('Monaco'),('Ribbon'); SELECT ItemID, ItemName FROM StudioItems ORDER BY ItemID;"",""rowLimit"":100}}")
    AssertTrue InStr(1, strReply, """hasResultSet"":true", vbBinaryCompare) > 0 And _
               InStr(1, strReply, """schemaChanged"":true", vbBinaryCompare) > 0 And _
               InStr(1, strReply, "Monaco", vbBinaryCompare) > 0, _
               "Studio multi-statement DDL + INSERT + SELECT returned grid & schemaChanged"

    strReply = SQLiteStudioDispatchHeadless("", _
        "{""id"":3,""action"":""studio.getSchema"",""payload"":{}}")
    AssertTrue InStr(1, strReply, """StudioItems""", vbBinaryCompare) > 0 And _
               InStr(1, strReply, """ItemName""", vbBinaryCompare) > 0, _
               "Studio studio.getSchema returned StudioItems & columns"

    strReply = SQLiteStudioDispatchHeadless("", _
        "{""id"":4,""action"":""studio.saveQuery"",""payload"":{""name"":""All Studio Items"",""category"":""Demo"",""sql"":""SELECT * FROM StudioItems;"",""description"":""Test saved query""}}")
    AssertTrue InStr(1, strReply, """All Studio Items""", vbBinaryCompare) > 0, "Studio studio.saveQuery persisted stored query"

    strReply = SQLiteStudioDispatchHeadless("", _
        "{""id"":5,""action"":""studio.explainSql"",""payload"":{""sql"":""SELECT * FROM StudioItems WHERE ItemID = 1;""}}")
    AssertTrue InStr(1, strReply, """nodes"":", vbBinaryCompare) > 0, "Studio studio.explainSql returned query plan nodes"

    strReply = SQLiteStudioDispatchHeadless("", _
        "{""id"":6,""action"":""studio.deleteQuery"",""payload"":{""name"":""All Studio Items""}}")
    AssertTrue InStr(1, strReply, """savedQueries"":[]", vbBinaryCompare) > 0, "Studio studio.deleteQuery deleted stored query"

    ResetHeadlessSQLiteStudio
End Sub


