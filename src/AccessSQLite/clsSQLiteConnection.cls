Attribute VB_Name = "clsSQLiteConnection"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

' =============================================================================
' clsSQLiteConnection
' High-level SQLite connection manager with automatic local-WAL vs SMB-safe
' pragma profiles, prepared statement cache, nested SAVEPOINT transactions,
' sub-millisecond SQL trace, online backup, and ORM entry points.
' =============================================================================

Private Const MAX_STMT_CACHE_SIZE As Long = 128
Private Const DEFAULT_BUSY_TIMEOUT_MS As Long = 5000

Private m_Driver As ISQLiteDriver
Private m_DbHandle As LongPtr
Private m_FilePath As String
Private m_DateStorage As SQLiteDateStorage
Private m_ActiveProfile As SQLitePragmaProfile
Private m_BusyTimeoutMs As Long
Private m_CacheEnabled As Boolean
Private m_TxDepth As Long

Private m_StmtCache As Scripting.Dictionary
Private m_TableMaps As Scripting.Dictionary
Private m_Repos As Scripting.Dictionary
Private m_TraceLog As Collection
Private m_QueryCount As Long

Private Sub Class_Initialize()
    m_DateStorage = sqDateIso8601
    m_BusyTimeoutMs = DEFAULT_BUSY_TIMEOUT_MS
    m_CacheEnabled = True
    Set m_StmtCache = New Scripting.Dictionary
    m_StmtCache.CompareMode = BinaryCompare
    Set m_TableMaps = NewTextDict()
    Set m_Repos = NewTextDict()
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    CloseConnection
    On Error GoTo 0
End Sub

Public Property Get IsOpen() As Boolean
    IsOpen = (m_DbHandle <> 0 And Not m_Driver Is Nothing)
End Property

Public Property Get FilePath() As String
    FilePath = m_FilePath
End Property

Public Property Get DbHandle() As LongPtr
    DbHandle = m_DbHandle
End Property

Public Property Get Driver() As ISQLiteDriver
    Set Driver = m_Driver
End Property

Public Property Get DateStorage() As SQLiteDateStorage
    DateStorage = m_DateStorage
End Property

Public Property Let DateStorage(ByVal enmMode As SQLiteDateStorage)
    m_DateStorage = enmMode
    ClearStatementCache
End Property

Public Property Get ActivePragmaProfile() As SQLitePragmaProfile
    ActivePragmaProfile = m_ActiveProfile
End Property

Public Property Get StatementCacheEnabled() As Boolean
    StatementCacheEnabled = m_CacheEnabled
End Property

Public Property Let StatementCacheEnabled(ByVal blnEnabled As Boolean)
    m_CacheEnabled = blnEnabled
    If Not blnEnabled Then ClearStatementCache
End Property

Public Property Get CachedStatementCount() As Long
    CachedStatementCount = m_StmtCache.Count
End Property

Public Property Get InTransaction() As Boolean
    InTransaction = (m_TxDepth > 0)
End Property

Public Property Get TransactionDepth() As Long
    TransactionDepth = m_TxDepth
End Property

Public Property Get ExecutedQueryCount() As Long
    ExecutedQueryCount = m_QueryCount
End Property

Public Sub ResetExecutedQueryCount()
    m_QueryCount = 0
End Sub

Public Property Get LastInsertRowId() As Variant
    EnsureOpen
    LastInsertRowId = m_Driver.GetLastInsertRowId(m_DbHandle)
End Property

Public Property Get ChangesCount() As Long
    EnsureOpen
    ChangesCount = m_Driver.GetChanges(m_DbHandle)
End Property

Public Property Get TotalChangesCount() As Long
    EnsureOpen
    TotalChangesCount = m_Driver.GetTotalChanges(m_DbHandle)
End Property

' Opens a file-backed or :memory: SQLite database with automatic driver & pragma profile selection.
Public Sub OpenConnection( _
    ByVal strFilePath As String, _
    Optional ByVal enmProfile As SQLitePragmaProfile = sqProfileAuto, _
    Optional ByVal enmDriver As SQLiteDriverKind = sqDriverAuto, _
    Optional ByVal lngOpenFlags As Long = 0, _
    Optional ByVal strDllDirectory As String = vbNullString)

    Dim strCleanPath As String
    Dim lngFlags As Long
    Dim lngRc As Long
    Dim strErr As String

    strCleanPath = Trim$(strFilePath)
    If Len(strCleanPath) = 0 Then
        RaiseSQLiteError sqErrInvalidArgument, "clsSQLiteConnection.OpenConnection", "Database file path cannot be empty."
    End If

    If IsOpen Then CloseConnection

    Set m_Driver = ResolveDriver(enmDriver, strDllDirectory)

    If lngOpenFlags = 0 Then
        If enmProfile = sqProfileReadOnly Then
            lngFlags = sqOpenReadOnly
        Else
            lngFlags = sqOpenReadWrite Or sqOpenCreate
        End If
    Else
        lngFlags = lngOpenFlags
    End If

    If StrComp(strCleanPath, ":memory:", vbTextCompare) = 0 And lngOpenFlags = 0 Then
        lngRc = m_Driver.OpenDb(":memory:", m_DbHandle)
    Else
        lngRc = m_Driver.OpenDbV2(strCleanPath, m_DbHandle, lngFlags)
    End If

    If lngRc <> sqOk Or m_DbHandle = 0 Then
        If m_DbHandle <> 0 Then
            strErr = m_Driver.GetErrMsg(m_DbHandle)
            m_Driver.CloseDb m_DbHandle
            m_DbHandle = 0
        End If
        If Len(strErr) = 0 Then strErr = "Unable to open SQLite database '" & strCleanPath & "'."
        RaiseSQLiteError sqErrInitFailed, "clsSQLiteConnection.OpenConnection", strErr, lngRc
    End If

    m_FilePath = strCleanPath
    m_TxDepth = 0
    m_QueryCount = 0
    ApplyPragmaProfile enmProfile, strCleanPath, lngFlags
End Sub

Public Sub OpenInMemory( _
    Optional ByVal enmDriver As SQLiteDriverKind = sqDriverAuto, _
    Optional ByVal strDllDirectory As String = vbNullString)

    OpenConnection ":memory:", sqProfileInMemory, enmDriver, 0, strDllDirectory
End Sub

Public Sub CloseConnection()
    If m_DbHandle = 0 Then Exit Sub
    On Error Resume Next
    If m_TxDepth > 0 Then
        m_TxDepth = 1
        RollbackTrans
    End If
    ClearStatementCache
    Set m_Repos = NewTextDict()
    If Not m_Driver Is Nothing Then
        m_Driver.CloseDb m_DbHandle
    End If
    m_DbHandle = 0
    m_FilePath = vbNullString
    m_TxDepth = 0
    On Error GoTo 0
End Sub

Public Sub ClearStatementCache()
    Dim varKey As Variant
    Dim objStmt As clsSQLiteStatement
    If m_StmtCache Is Nothing Then Exit Sub
    For Each varKey In m_StmtCache.Keys
        Set objStmt = m_StmtCache(varKey)
        If Not objStmt Is Nothing Then
            objStmt.FinalizeStatement
        End If
    Next varKey
    m_StmtCache.RemoveAll
End Sub

' Prepares a single SQL statement, returning a cached clsSQLiteStatement when eligible.
Public Function PrepareStatement( _
    ByVal strSql As String, _
    Optional ByVal blnAllowCache As Boolean = True) As clsSQLiteStatement

    Dim strTrimmed As String
    Dim blnShouldCache As Boolean
    Dim ptrStmt As LongPtr
    Dim lngPrepFlags As Long
    Dim lngRc As Long
    Dim strRem As String
    Dim objStmt As clsSQLiteStatement
    Dim varFirstKey As Variant
    Dim objEvicted As clsSQLiteStatement

    EnsureOpen
    strTrimmed = Trim$(strSql)
    If Len(strTrimmed) = 0 Then
        RaiseSQLiteError sqErrInvalidArgument, "clsSQLiteConnection.PrepareStatement", "SQL statement cannot be empty."
    End If

    blnShouldCache = (blnAllowCache And m_CacheEnabled And IsCacheableSql(strTrimmed))

    If blnShouldCache And m_StmtCache.Exists(strTrimmed) Then
        Set objStmt = m_StmtCache(strTrimmed)
        If objStmt.IsValid Then
            objStmt.ResetForReuse
            Set PrepareStatement = objStmt
            Exit Function
        Else
            m_StmtCache.Remove strTrimmed
        End If
    End If

    lngPrepFlags = IIf(blnShouldCache, sqPreparePersistent, sqPrepareNormal)
    lngRc = m_Driver.PrepareStmt(m_DbHandle, strTrimmed, ptrStmt, lngPrepFlags, strRem)
    If lngRc <> sqOk Or ptrStmt = 0 Then
        RaiseConnectionSqlError "PrepareStatement", lngRc, strTrimmed
    End If

    ' Do not cache multi-statement strings where trailing SQL remains.
    If Len(Trim$(strRem)) > 0 Then
        blnShouldCache = False
    End If

    Set objStmt = New clsSQLiteStatement
    objStmt.InitStatement m_Driver, m_DbHandle, ptrStmt, strTrimmed, m_DateStorage, blnShouldCache

    If blnShouldCache Then
        If m_StmtCache.Count >= MAX_STMT_CACHE_SIZE Then
            For Each varFirstKey In m_StmtCache.Keys
                Set objEvicted = m_StmtCache(varFirstKey)
                If Not objEvicted Is Nothing Then objEvicted.FinalizeStatement
                m_StmtCache.Remove varFirstKey
                Exit For
            Next varFirstKey
        End If
        Set m_StmtCache(strTrimmed) = objStmt
    End If

    Set PrepareStatement = objStmt
End Function

' Executes a non-query SQL statement (INSERT/UPDATE/DELETE/DDL) and returns affected rows.
Public Function ExecuteSql(ByVal strSql As String, ParamArray varParams() As Variant) As Long
    Dim varPassed As Variant
    varPassed = varParams
    ExecuteSql = ExecuteParams(strSql, varPassed)
End Function

Public Function ExecuteNamed(ByVal strSql As String, ByVal dictParams As Scripting.Dictionary) As Long
    ExecuteNamed = ExecuteParams(strSql, dictParams)
End Function

Public Function ExecuteParams(ByVal strSql As String, ByVal varParamList As Variant) As Long
    Dim objStmt As clsSQLiteStatement
    Dim curStart As Currency
    Dim strTraceSql As String

    EnsureOpen
    curStart = StartPerfTimer()
    Set objStmt = PrepareStatement(strSql, True)

    On Error GoTo ExecFail
    objStmt.BindFlexibleParams varParamList
    If Not m_TraceLog Is Nothing Then strTraceSql = objStmt.ExpandedSql

    Do While objStmt.StepOnce() = sqRow
        ' Consume any rows if caller passed a RETURNING/PRAGMA statement to ExecuteSql.
    Loop

    objStmt.ResetForReuse
    If Not objStmt.IsCached Then objStmt.FinalizeStatement

    m_QueryCount = m_QueryCount + 1
    RecordTrace IIf(Len(strTraceSql) > 0, strTraceSql, strSql), curStart
    ExecuteParams = m_Driver.GetChanges(m_DbHandle)
    Exit Function

ExecFail:
    Dim lngErrNum As Long
    Dim strErrSrc As String
    Dim strErrDesc As String
    lngErrNum = Err.Number
    strErrSrc = Err.Source
    strErrDesc = Err.Description
    If Not objStmt Is Nothing Then
        objStmt.ResetForReuse
        If Not objStmt.IsCached Then objStmt.FinalizeStatement
    End If
    Err.Raise lngErrNum, strErrSrc, strErrDesc
End Function

' Executes a multi-statement SQL script (DDL migrations, PRAGMAs, seed batches) using SQLite's native C lexer.
Public Function ExecuteScript(ByVal strMultiSql As String) As Long
    Dim strPending As String
    Dim strRemainder As String
    Dim ptrStmt As LongPtr
    Dim lngRc As Long
    Dim lngStmtCount As Long
    Dim curStart As Currency

    EnsureOpen
    strPending = strMultiSql
    curStart = StartPerfTimer()

    Do While Len(Trim$(strPending)) > 0
        ptrStmt = 0
        strRemainder = vbNullString
        lngRc = m_Driver.PrepareStmt(m_DbHandle, strPending, ptrStmt, sqPrepareNormal, strRemainder)
        If lngRc <> sqOk Then
            If ptrStmt <> 0 Then m_Driver.FinalizeStmt ptrStmt
            RaiseConnectionSqlError "ExecuteScript", lngRc, strPending
        End If

        If ptrStmt <> 0 Then
            Do
                lngRc = m_Driver.StepStmt(ptrStmt)
            Loop While lngRc = sqRow

            If lngRc <> sqDone And lngRc <> sqOk Then
                m_Driver.FinalizeStmt ptrStmt
                RaiseConnectionSqlError "ExecuteScript", lngRc, strPending
            End If
            m_Driver.FinalizeStmt ptrStmt
            lngStmtCount = lngStmtCount + 1
            m_QueryCount = m_QueryCount + 1
        End If

        If StrComp(strPending, strRemainder, vbBinaryCompare) = 0 Then Exit Do
        strPending = strRemainder
    Loop

    RecordTrace "SCRIPT (" & lngStmtCount & " statements)", curStart
    ExecuteScript = lngStmtCount
End Function

' Queries rows as Collection of Scripting.Dictionary (JSON-friendly ISO date strings).
Public Function QueryRows(ByVal strSql As String, ParamArray varParams() As Variant) As Collection
    Dim varPassed As Variant
    varPassed = varParams
    Set QueryRows = QueryParams(strSql, varPassed, False, Nothing)
End Function

' Queries rows preserving native VBA types (Date, Boolean, Currency) where declared or mapped.
Public Function QueryRowsRaw(ByVal strSql As String, ParamArray varParams() As Variant) As Collection
    Dim varPassed As Variant
    varPassed = varParams
    Set QueryRowsRaw = QueryParams(strSql, varPassed, True, Nothing)
End Function

Public Function QueryParams( _
    ByVal strSql As String, _
    ByVal varParamList As Variant, _
    Optional ByVal blnRawTypes As Boolean = False, _
    Optional ByVal objTableMap As clsSQLiteTableMap = Nothing) As Collection

    Dim objStmt As clsSQLiteStatement
    Dim colResult As Collection
    Dim curStart As Currency
    Dim strTraceSql As String

    EnsureOpen
    curStart = StartPerfTimer()
    Set objStmt = PrepareStatement(strSql, True)

    On Error GoTo QueryFail
    objStmt.BindFlexibleParams varParamList
    If Not m_TraceLog Is Nothing Then strTraceSql = objStmt.ExpandedSql
    Set colResult = objStmt.ReadAllRows(blnRawTypes, objTableMap)
    If Not objStmt.IsCached Then objStmt.FinalizeStatement

    m_QueryCount = m_QueryCount + 1
    RecordTrace IIf(Len(strTraceSql) > 0, strTraceSql, strSql), curStart
    Set QueryParams = colResult
    Exit Function

QueryFail:
    Dim lngErrNum As Long
    Dim strErrSrc As String
    Dim strErrDesc As String
    lngErrNum = Err.Number
    strErrSrc = Err.Source
    strErrDesc = Err.Description
    If Not objStmt Is Nothing Then
        objStmt.ResetForReuse
        If Not objStmt.IsCached Then objStmt.FinalizeStatement
    End If
    Err.Raise lngErrNum, strErrSrc, strErrDesc
End Function

Public Function QueryOneRow(ByVal strSql As String, ParamArray varParams() As Variant) As Scripting.Dictionary
    Dim varPassed As Variant
    Dim colRows As Collection
    varPassed = varParams
    Set colRows = QueryParams(strSql, varPassed, False, Nothing)
    If colRows.Count > 0 Then Set QueryOneRow = colRows(1)
End Function

Public Function QueryOneRowRaw(ByVal strSql As String, ParamArray varParams() As Variant) As Scripting.Dictionary
    Dim varPassed As Variant
    Dim colRows As Collection
    varPassed = varParams
    Set colRows = QueryParams(strSql, varPassed, True, Nothing)
    If colRows.Count > 0 Then Set QueryOneRowRaw = colRows(1)
End Function

Public Function QueryScalar(ByVal strSql As String, ParamArray varParams() As Variant) As Variant
    Dim varPassed As Variant
    varPassed = varParams
    QueryScalar = QueryScalarParams(strSql, varPassed)
End Function

Public Function QueryScalarParams(ByVal strSql As String, ByVal varParamList As Variant) As Variant
    Dim objStmt As clsSQLiteStatement
    Dim lngRc As Long
    Dim varResult As Variant
    Dim curStart As Currency
    Dim strTraceSql As String

    EnsureOpen
    curStart = StartPerfTimer()
    Set objStmt = PrepareStatement(strSql, True)

    On Error GoTo ScalarFail
    objStmt.BindFlexibleParams varParamList
    If Not m_TraceLog Is Nothing Then strTraceSql = objStmt.ExpandedSql

    lngRc = objStmt.StepOnce()
    If lngRc = sqRow And objStmt.ColumnCount() > 0 Then
        varResult = objStmt.ReadColumnValue(0, False)
    Else
        varResult = Null
    End If

    objStmt.ResetForReuse
    If Not objStmt.IsCached Then objStmt.FinalizeStatement

    m_QueryCount = m_QueryCount + 1
    RecordTrace IIf(Len(strTraceSql) > 0, strTraceSql, strSql), curStart
    QueryScalarParams = varResult
    Exit Function

ScalarFail:
    Dim lngErrNum As Long
    Dim strErrSrc As String
    Dim strErrDesc As String
    lngErrNum = Err.Number
    strErrSrc = Err.Source
    strErrDesc = Err.Description
    If Not objStmt Is Nothing Then
        objStmt.ResetForReuse
        If Not objStmt.IsCached Then objStmt.FinalizeStatement
    End If
    Err.Raise lngErrNum, strErrSrc, strErrDesc
End Function

' =============================================================================
' Transactions & Nested Savepoints
' =============================================================================

Public Sub BeginTrans()
    EnsureOpen
    If m_TxDepth = 0 Then
        ExecuteParams "BEGIN IMMEDIATE TRANSACTION;", Empty
        m_TxDepth = 1
    Else
        m_TxDepth = m_TxDepth + 1
        ExecuteParams "SAVEPOINT sp_orm_" & m_TxDepth & ";", Empty
    End If
End Sub

Public Sub CommitTrans()
    EnsureOpen
    If m_TxDepth <= 0 Then
        RaiseSQLiteError sqErrSqlFailed, "clsSQLiteConnection.CommitTrans", "No active transaction to commit."
    End If
    If m_TxDepth = 1 Then
        ExecuteParams "COMMIT TRANSACTION;", Empty
        m_TxDepth = 0
    Else
        ExecuteParams "RELEASE SAVEPOINT sp_orm_" & m_TxDepth & ";", Empty
        m_TxDepth = m_TxDepth - 1
    End If
End Sub

Public Sub RollbackTrans()
    Dim lngCurDepth As Long
    If m_DbHandle = 0 Or m_TxDepth <= 0 Then Exit Sub
    If m_TxDepth = 1 Then
        m_TxDepth = 0
        On Error Resume Next
        ExecuteParams "ROLLBACK TRANSACTION;", Empty
        On Error GoTo 0
    Else
        lngCurDepth = m_TxDepth
        m_TxDepth = m_TxDepth - 1
        ExecuteParams "ROLLBACK TO SAVEPOINT sp_orm_" & lngCurDepth & ";", Empty
        ExecuteParams "RELEASE SAVEPOINT sp_orm_" & lngCurDepth & ";", Empty
    End If
End Sub

' =============================================================================
' Diagnostics & Maintenance
' =============================================================================

Public Sub StartTrace()
    Set m_TraceLog = New Collection
End Sub

Public Function StopTrace() As String
    Dim varLine As Variant
    Dim strOut As String
    If m_TraceLog Is Nothing Then Exit Function
    For Each varLine In m_TraceLog
        strOut = strOut & CStr(varLine) & vbCrLf
    Next varLine
    Set m_TraceLog = Nothing
    StopTrace = strOut
End Function

Public Property Get TraceEntries() As Collection
    Set TraceEntries = m_TraceLog
End Property

Public Sub BackupToFile(ByVal strDestFilePath As String)
    Dim ptrDestDb As LongPtr
    Dim lngRc As Long
    Dim strErr As String

    EnsureOpen
    If Len(Trim$(strDestFilePath)) = 0 Then
        RaiseSQLiteError sqErrInvalidArgument, "clsSQLiteConnection.BackupToFile", "Destination backup path cannot be empty."
    End If

    lngRc = m_Driver.OpenDb(strDestFilePath, ptrDestDb)
    If lngRc <> sqOk Or ptrDestDb = 0 Then
        If ptrDestDb <> 0 Then m_Driver.CloseDb ptrDestDb
        RaiseSQLiteError sqErrSqlFailed, "clsSQLiteConnection.BackupToFile", _
            "Failed to open backup destination '" & strDestFilePath & "'.", lngRc
    End If

    lngRc = m_Driver.BackupDatabase(ptrDestDb, "main", m_DbHandle, "main")
    If lngRc <> sqOk Then
        strErr = m_Driver.GetErrMsg(ptrDestDb)
    End If
    m_Driver.CloseDb ptrDestDb

    If lngRc <> sqOk Then
        RaiseSQLiteError sqErrSqlFailed, "clsSQLiteConnection.BackupToFile", _
            "Online backup failed: " & strErr, lngRc
    End If
End Sub

Public Sub VacuumDb()
    EnsureOpen
    If m_TxDepth > 0 Then
        RaiseSQLiteError sqErrSqlFailed, "clsSQLiteConnection.VacuumDb", "VACUUM cannot run inside an active transaction."
    End If
    ClearStatementCache
    ExecuteParams "VACUUM;", Empty
End Sub

Public Function IntegrityCheck() As Boolean
    Dim varRes As Variant
    EnsureOpen
    varRes = QueryScalarParams("PRAGMA integrity_check;", Empty)
    If Not IsNull(varRes) Then
        IntegrityCheck = (StrComp(CStr(varRes), "ok", vbTextCompare) = 0)
    End If
End Function

Public Function TableExists(ByVal strTableName As String) As Boolean
    Dim varCnt As Variant
    EnsureOpen
    varCnt = QueryScalar( _
        "SELECT COUNT(*) FROM sqlite_master WHERE type IN ('table','view') AND LOWER(name) = LOWER(?);", _
        Trim$(strTableName))
    If Not IsNull(varCnt) Then TableExists = (CLng(varCnt) > 0)
End Function

' =============================================================================
' Schema, Query Builder & ORM Entry Points
' =============================================================================

Public Function DefineTable(ByVal strTableName As String, Optional ByVal strPrimaryKey As String = "ID") As clsSQLiteTableMap
    Dim objMap As clsSQLiteTableMap
    Set objMap = New clsSQLiteTableMap
    objMap.InitMap strTableName, strPrimaryKey
    RegisterTableMap objMap
    Set DefineTable = objMap
End Function

Public Sub RegisterTableMap(ByVal objMap As clsSQLiteTableMap)
    If objMap Is Nothing Then Exit Sub
    Set m_TableMaps(objMap.TableName) = objMap
    If m_Repos.Exists(objMap.TableName) Then
        m_Repos.Remove objMap.TableName
    End If
End Sub

Public Function HasTableMap(ByVal strTableName As String) As Boolean
    HasTableMap = m_TableMaps.Exists(Trim$(strTableName))
End Function

Public Function GetTableMap(ByVal strTableName As String) As clsSQLiteTableMap
    Dim strKey As String
    strKey = Trim$(strTableName)
    If m_TableMaps.Exists(strKey) Then
        Set GetTableMap = m_TableMaps(strKey)
    End If
End Function

Public Function SchemaManager() As clsSQLiteSchema
    Dim objSchema As clsSQLiteSchema
    Set objSchema = New clsSQLiteSchema
    objSchema.InitSchema Me
    Set SchemaManager = objSchema
End Function

Public Function FromTable( _
    ByVal strTableName As String, _
    Optional ByVal strAlias As String = vbNullString) As clsSQLiteQueryBuilder

    Dim objQb As clsSQLiteQueryBuilder
    Dim objMap As clsSQLiteTableMap
    Set objMap = GetTableMap(strTableName)
    Set objQb = New clsSQLiteQueryBuilder
    objQb.InitBuilder Me, strTableName, strAlias, objMap
    Set FromTable = objQb
End Function

Public Function Repository(ByVal strTableName As String) As clsSQLiteRepository
    Dim strKey As String
    Dim objMap As clsSQLiteTableMap
    Dim objRepo As clsSQLiteRepository

    EnsureOpen
    strKey = Trim$(strTableName)
    If m_Repos.Exists(strKey) Then
        Set Repository = m_Repos(strKey)
        Exit Function
    End If

    Set objMap = GetTableMap(strKey)
    If objMap Is Nothing Then
        Set objMap = SchemaManager().ReflectTableMap(strKey)
        RegisterTableMap objMap
    End If

    Set objRepo = New clsSQLiteRepository
    objRepo.InitRepo Me, objMap
    Set m_Repos(strKey) = objRepo
    Set Repository = objRepo
End Function

Public Function NewUnitOfWork() As clsSQLiteUnitOfWork
    Dim objUow As clsSQLiteUnitOfWork
    Set objUow = New clsSQLiteUnitOfWork
    objUow.InitUnitOfWork Me
    Set NewUnitOfWork = objUow
End Function

Public Function NewMigrator() As clsSQLiteMigrator
    Dim objMig As clsSQLiteMigrator
    Set objMig = New clsSQLiteMigrator
    objMig.InitMigrator Me
    Set NewMigrator = objMig
End Function

Public Function NewAccessBridge(Optional ByVal objDaoDb As DAO.Database = Nothing) As clsSQLiteAccessBridge
    Dim objBridge As clsSQLiteAccessBridge
    Set objBridge = New clsSQLiteAccessBridge
    objBridge.InitBridge Me, objDaoDb
    Set NewAccessBridge = objBridge
End Function

Public Function NewAdoBridge() As clsSQLiteAdoBridge
    Dim objAdo As clsSQLiteAdoBridge
    Set objAdo = New clsSQLiteAdoBridge
    objAdo.InitAdoBridge Me
    Set NewAdoBridge = objAdo
End Function

' Materializes a parameterized SQLite SELECT into a disconnected in-memory ADODB.Recordset.
Public Function QueryAdoRecordset(ByVal strSql As String, ParamArray varParams() As Variant) As Object
    Dim varPassed As Variant
    varPassed = varParams
    Set QueryAdoRecordset = QueryAdoParams(strSql, varPassed, Nothing, False, "IsSelected")
End Function

Public Function QueryAdoParams( _
    ByVal strSql As String, _
    ByVal varParamList As Variant, _
    Optional ByVal objTableMap As clsSQLiteTableMap = Nothing, _
    Optional ByVal blnIncludeSelectColumn As Boolean = False, _
    Optional ByVal strSelectColumnName As String = "IsSelected") As Object

    Dim curStart As Currency
    EnsureOpen
    curStart = StartPerfTimer()
    Set QueryAdoParams = NewAdoBridge().QueryToAdoRecordset( _
        strSql, varParamList, objTableMap, blnIncludeSelectColumn, strSelectColumnName)
    m_QueryCount = m_QueryCount + 1
    RecordTrace "ADO: " & strSql, curStart
End Function

' =============================================================================
' Internal Helpers
' =============================================================================

Private Function ResolveDriver( _
    ByVal enmKind As SQLiteDriverKind, _
    ByVal strDllDirectory As String) As ISQLiteDriver

    Dim objWin As clsSQLiteWinDriver
    Dim objWinIface As ISQLiteDriver
    Dim objBundled As clsSQLiteBundledDriver
    Dim objBundledIface As ISQLiteDriver

    Select Case enmKind
        Case sqDriverWinSqlite
            Set objWin = New clsSQLiteWinDriver
            Set objWinIface = objWin
            If Not objWinIface.InitializeDriver(strDllDirectory) Then
                RaiseSQLiteError sqErrInitFailed, "clsSQLiteConnection.ResolveDriver", objWinIface.LastInitError
            End If
            Set ResolveDriver = objWinIface

        Case sqDriverBundled
            Set objBundled = New clsSQLiteBundledDriver
            Set objBundledIface = objBundled
            If Not objBundledIface.InitializeDriver(strDllDirectory) Then
                RaiseSQLiteError sqErrInitFailed, "clsSQLiteConnection.ResolveDriver", objBundledIface.LastInitError
            End If
            Set ResolveDriver = objBundledIface

        Case Else ' sqDriverAuto
            Set objWin = New clsSQLiteWinDriver
            Set objWinIface = objWin
            If objWinIface.InitializeDriver(strDllDirectory) Then
                Set ResolveDriver = objWinIface
                Exit Function
            End If

            Set objBundled = New clsSQLiteBundledDriver
            Set objBundledIface = objBundled
            If objBundledIface.InitializeDriver(strDllDirectory) Then
                Set ResolveDriver = objBundledIface
                Exit Function
            End If

            RaiseSQLiteError sqErrInitFailed, "clsSQLiteConnection.ResolveDriver", _
                "Both SQLite drivers failed to initialize. WinSQLite: [" & objWinIface.LastInitError & _
                "] | Bundled: [" & objBundledIface.LastInitError & "]"
    End Select
End Function

Private Sub ApplyPragmaProfile( _
    ByVal enmRequested As SQLitePragmaProfile, _
    ByVal strFilePath As String, _
    ByVal lngOpenFlags As Long)

    Dim enmResolved As SQLitePragmaProfile

    If enmRequested = sqProfileAuto Then
        If StrComp(strFilePath, ":memory:", vbTextCompare) = 0 Or (lngOpenFlags And sqOpenMemory) <> 0 Then
            enmResolved = sqProfileInMemory
        ElseIf (lngOpenFlags And sqOpenReadOnly) <> 0 Then
            enmResolved = sqProfileReadOnly
        ElseIf IsNetworkOrUncPath(strFilePath) Then
            enmResolved = sqProfileNetworkSafe
        Else
            enmResolved = sqProfileLocalWal
        End If
    Else
        enmResolved = enmRequested
    End If

    m_ActiveProfile = enmResolved
    m_Driver.SetBusyTimeout m_DbHandle, m_BusyTimeoutMs

    Select Case enmResolved
        Case sqProfileLocalWal
            ExecuteScript "PRAGMA journal_mode = WAL; PRAGMA synchronous = NORMAL; PRAGMA foreign_keys = ON; PRAGMA temp_store = MEMORY;"
        Case sqProfileNetworkSafe
            ExecuteScript "PRAGMA journal_mode = DELETE; PRAGMA synchronous = FULL; PRAGMA foreign_keys = ON; PRAGMA temp_store = MEMORY;"
        Case sqProfileReadOnly
            ExecuteScript "PRAGMA query_only = ON; PRAGMA temp_store = MEMORY;"
        Case sqProfileInMemory
            ExecuteScript "PRAGMA foreign_keys = ON; PRAGMA temp_store = MEMORY;"
    End Select
    m_QueryCount = 0
End Sub

Private Function IsCacheableSql(ByVal strTrimmed As String) As Boolean
    Dim strPrefix As String
    strPrefix = UCase$(Left$(strTrimmed, 10))
    If Left$(strPrefix, 7) = "SELECT " Or Left$(strPrefix, 5) = "WITH " Or _
       Left$(strPrefix, 7) = "INSERT " Or Left$(strPrefix, 7) = "UPDATE " Or _
       Left$(strPrefix, 7) = "DELETE " Or Left$(strPrefix, 8) = "REPLACE " Then
        IsCacheableSql = True
    End If
End Function

Private Sub RecordTrace(ByVal strSql As String, ByVal curStart As Currency)
    Dim dblMs As Double
    Dim strOneLine As String
    If m_TraceLog Is Nothing Then Exit Sub
    dblMs = ElapsedPerfMs(curStart)
    strOneLine = Replace(Replace(strSql, vbCrLf, " "), vbLf, " ")
    m_TraceLog.Add Format$(dblMs, "0.00") & " ms | " & Left$(strOneLine, 220)
End Sub

Private Sub EnsureOpen()
    If Not IsOpen Then
        RaiseSQLiteError sqErrConnectionClosed, "clsSQLiteConnection", "SQLite database connection is not open."
    End If
End Sub

Private Sub RaiseConnectionSqlError(ByVal strMethod As String, ByVal lngRc As Long, ByVal strSql As String)
    Dim lngExt As Long
    Dim strMsg As String
    Dim enmErr As SQLiteErrorNumber

    lngExt = m_Driver.GetExtendedErrCode(m_DbHandle)
    strMsg = m_Driver.GetErrMsg(m_DbHandle)
    If Len(strMsg) = 0 Then strMsg = "SQLite operation failed."
    strMsg = strMsg & " | SQL: " & Left$(strSql, 200)

    If lngRc = sqConstraint Or (lngExt And &HFF&) = sqConstraint Then
        enmErr = sqErrConstraint
    Else
        enmErr = sqErrSqlFailed
    End If
    RaiseSQLiteError enmErr, "clsSQLiteConnection." & strMethod, strMsg, lngRc, lngExt
End Sub

