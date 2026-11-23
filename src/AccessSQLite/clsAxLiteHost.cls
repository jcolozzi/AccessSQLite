Attribute VB_Name = "clsAxLiteHost"
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
' clsAxLiteHost
' Host controller for the Monaco-based AxLite SQLite front-end client.
' Drives both the live Edge Browser Control form (frmAxLite) and
' headless JSON dispatch for automated verification.
' =============================================================================

Private Const SAVED_QUERIES_TABLE As String = "_axlite_saved_queries"
Private Const LEGACY_SAVED_QUERIES_TABLE As String = "_studio_saved_queries"
Private Const ACTIVE_TIMER_MS As Long = 40
Private Const IDLE_TIMER_MS As Long = 150
Private Const INITIAL_LOAD_DELAY_MS As Long = 15
Private Const ACTIVE_WINDOW_SEC As Double = 3#

Private Const MSO_FILE_DIALOG_SAVEAS As Long = 2
Private Const MSO_FILE_DIALOG_OPEN As Long = 3

Private m_frmHost As Access.Form
Private m_objEdge As Object
Private m_Bridge As clsSQLiteWebBridge
Private m_Conn As clsSQLiteConnection
Private m_CurrentDbPath As String

Private m_blnProcessing As Boolean
Private m_blnLoadPending As Boolean
Private m_dblLastActivity As Double

Private Sub Class_Terminate()
    ShutdownHost
End Sub

Public Property Get Connection() As clsSQLiteConnection
    Set Connection = m_Conn
End Property

Public Property Get CurrentDatabasePath() As String
    CurrentDatabasePath = m_CurrentDbPath
End Property

Public Sub StartHost( _
    ByVal frmHost As Access.Form, _
    ByVal objEdge As Object, _
    Optional ByVal strInitialDbPath As String = vbNullString)

    Set m_frmHost = frmHost
    Set m_objEdge = objEdge
    OpenTargetDatabase ResolveInitialDbPath(strInitialDbPath)

    Set m_Bridge = New clsSQLiteWebBridge
    m_Bridge.InitBridge objEdge

    ' Defer Navigate & event wiring to first Timer tick to avoid break-mode errors when launched from Immediate/Ribbon
    m_blnLoadPending = True
    m_frmHost.TimerInterval = INITIAL_LOAD_DELAY_MS
End Sub

Public Sub StartHeadless(Optional ByVal strInitialDbPath As String = vbNullString)
    OpenTargetDatabase ResolveInitialDbPath(strInitialDbPath)
End Sub

Public Sub ShutdownHost()
    On Error Resume Next
    If Not m_frmHost Is Nothing Then m_frmHost.TimerInterval = 0
    If Not m_Conn Is Nothing Then
        m_Conn.CloseConnection
        Set m_Conn = Nothing
    End If
    Set m_Bridge = Nothing
    Set m_objEdge = Nothing
    Set m_frmHost = Nothing
    On Error GoTo 0
End Sub

Public Sub OnBeforeNavigate(ByRef intCancel As Integer, ByVal strUrl As String)
    If m_Bridge Is Nothing Then Exit Sub
    If Not m_Bridge.IsAppUrl(strUrl) Then
        intCancel = True
    ElseIf strUrl <> "about:blank" Then
        m_Bridge.MarkNavigating
    End If
End Sub

Public Sub OnDocumentComplete()
    If m_Bridge Is Nothing Then Exit Sub
    m_Bridge.MarkReady
    m_dblLastActivity = Timer
End Sub

Public Sub OnTimer()
    If m_blnProcessing Or m_Bridge Is Nothing Then Exit Sub
    m_blnProcessing = True
    On Error GoTo TimerErr

    If m_blnLoadPending Then
        m_blnLoadPending = False
        m_dblLastActivity = Timer
        m_objEdge.Properties("OnBeforeNavigation").Value = "[Event Procedure]"
        m_objEdge.Properties("OnDocumentComplete").Value = "[Event Procedure]"
        m_Bridge.LoadPage AxLiteHtmlFilePath()
    Else
        ProcessBridgeInbox
    End If

TimerDone:
    If Not m_frmHost Is Nothing Then
        m_frmHost.TimerInterval = IIf(SecondsElapsed(m_dblLastActivity) < ACTIVE_WINDOW_SEC, ACTIVE_TIMER_MS, IDLE_TIMER_MS)
    End If
    m_blnProcessing = False
    Exit Sub

TimerErr:
    If Not m_Bridge Is Nothing Then
        m_Bridge.EmitEvent "bridge.error", Err.Description
    End If
    Resume TimerDone
End Sub

Public Function DispatchJson(ByVal strEnvelopeJson As String) As String
    Dim udtParsed As ParseResult
    Dim dicErr As Scripting.Dictionary
    Dim dicReply As Scripting.Dictionary

    udtParsed = LibJSON.Parse(strEnvelopeJson)
    If Not udtParsed.IsValid Or Not IsObject(udtParsed.Value) Then
        Set dicReply = New Scripting.Dictionary
        dicReply("id") = Null
        dicReply("ok") = False
        Set dicErr = New Scripting.Dictionary
        dicErr("code") = "BadRequest"
        dicErr("message") = "Invalid JSON request envelope."
        Set dicReply("error") = dicErr
        DispatchJson = LibJSON.Serialize(dicReply)
        Exit Function
    End If

    DispatchJson = LibJSON.Serialize(DispatchEnvelope(udtParsed.Value))
End Function

' =============================================================================
' Envelope Dispatcher
' =============================================================================

Private Sub ProcessBridgeInbox()
    Dim colEnvelopes As Collection
    Dim varEnv As Variant

    Set colEnvelopes = m_Bridge.PollOutbox()
    If colEnvelopes.Count = 0 Then Exit Sub

    m_dblLastActivity = Timer
    m_Bridge.AckOutbox

    For Each varEnv In colEnvelopes
        If IsObject(varEnv) Then
            If TypeOf varEnv Is Scripting.Dictionary Then
                m_Bridge.SendReply DispatchEnvelope(varEnv)
            End If
        End If
    Next varEnv
End Sub

Public Function DispatchEnvelope(ByVal dicEnv As Scripting.Dictionary) As Scripting.Dictionary
    Dim dicReply As New Scripting.Dictionary
    Dim dicPayload As Scripting.Dictionary
    Dim strAction As String
    Dim varData As Variant
    Dim dicErr As Scripting.Dictionary

    If dicEnv.Exists("id") Then
        dicReply("id") = dicEnv("id")
    Else
        dicReply("id") = Null
    End If

    On Error GoTo ActionFailed
    strAction = Trim$(CStr(Nz(dicEnv("action"), vbNullString)))
    If dicEnv.Exists("payload") And IsObject(dicEnv("payload")) Then
        Set dicPayload = dicEnv("payload")
    Else
        Set dicPayload = NewTextDict()
    End If

    Select Case strAction
        Case "axlite.init"
            Set varData = BuildInitPayload()

        Case "axlite.getSchema"
            Set varData = BuildSchemaPayload()

        Case "axlite.openPath"
            OpenTargetDatabase CStr(Nz(dicPayload("dbPath"), ":memory:"))
            Set varData = BuildInitPayload()

        Case "axlite.openDialog"
            Set varData = HandleOpenDatabaseDialog(CStr(Nz(dicPayload("mode"), "open")))

        Case "axlite.execSql"
            Set varData = ExecuteAxLiteSql( _
                CStr(Nz(dicPayload("sql"), vbNullString)), _
                CLng(Nz(dicPayload("rowLimit"), 1000)))

        Case "axlite.explainSql"
            Set varData = ExplainAxLiteSql(CStr(Nz(dicPayload("sql"), vbNullString)))

        Case "axlite.listQueries"
            Set varData = New Scripting.Dictionary
            Set varData("savedQueries") = ListSavedQueries()

        Case "axlite.saveQuery"
            SaveStoredQuery _
                CStr(Nz(dicPayload("name"), vbNullString)), _
                CStr(Nz(dicPayload("category"), "General")), _
                CStr(Nz(dicPayload("sql"), vbNullString)), _
                CStr(Nz(dicPayload("description"), vbNullString))
            Set varData = New Scripting.Dictionary
            Set varData("savedQueries") = ListSavedQueries()

        Case "axlite.deleteQuery"
            DeleteStoredQuery CStr(Nz(dicPayload("name"), vbNullString))
            Set varData = New Scripting.Dictionary
            Set varData("savedQueries") = ListSavedQueries()

        Case "axlite.dbAction"
            Set varData = HandleDatabaseAction(dicPayload)

        Case Else
            RaiseSQLiteError sqErrInvalidArgument, "clsAxLiteHost.DispatchEnvelope", _
                "Unsupported AxLite action: " & strAction
    End Select

    dicReply("ok") = True
    If IsObject(varData) Then
        Set dicReply("data") = varData
    Else
        dicReply("data") = varData
    End If
    Set DispatchEnvelope = dicReply
    Exit Function

ActionFailed:
    Set dicErr = New Scripting.Dictionary
    dicErr("code") = MapErrNumberToCode(Err.Number)
    dicErr("message") = Err.Description
    dicReply("ok") = False
    Set dicReply("error") = dicErr
    Set DispatchEnvelope = dicReply
End Function

' =============================================================================
' Database Connection, Schema Introspection & Saved Queries
' =============================================================================

Private Sub OpenTargetDatabase(ByVal strPath As String)
    Dim strClean As String
    strClean = Trim$(strPath)
    If Len(strClean) = 0 Then strClean = ":memory:"

    If Not m_Conn Is Nothing Then
        m_Conn.CloseConnection
        Set m_Conn = Nothing
    End If

    Set m_Conn = New clsSQLiteConnection
    If StrComp(strClean, ":memory:", vbTextCompare) = 0 Then
        m_Conn.OpenInMemory sqDriverAuto
    Else
        m_Conn.OpenConnection strClean, sqProfileAuto, sqDriverAuto
    End If
    m_CurrentDbPath = strClean
    MigrateLegacySavedQueriesIfPresent
    EnsureStarterQueriesIfDemo

    If Not m_frmHost Is Nothing Then
        On Error Resume Next
        m_frmHost.Caption = "AxLite - " & ExtractLeafName(strClean)
        On Error GoTo 0
    End If
End Sub

Private Sub MigrateLegacySavedQueriesIfPresent()
    On Error Resume Next
    If m_Conn.TableExists(LEGACY_SAVED_QUERIES_TABLE) Then
        If Not m_Conn.TableExists(SAVED_QUERIES_TABLE) Then
            m_Conn.ExecuteSql "ALTER TABLE """ & LEGACY_SAVED_QUERIES_TABLE & """ RENAME TO """ & SAVED_QUERIES_TABLE & """;"
        Else
            m_Conn.ExecuteSql "DROP TABLE IF EXISTS """ & LEGACY_SAVED_QUERIES_TABLE & """;"
        End If
    End If
    On Error GoTo 0
End Sub

Private Function BuildInitPayload() As Scripting.Dictionary
    Dim dicOut As New Scripting.Dictionary
    EnsureOpenConn

    dicOut("dbPath") = m_CurrentDbPath
    dicOut("dbName") = ExtractLeafName(m_CurrentDbPath)
    dicOut("engineVersion") = m_Conn.Driver.LibraryVersion
    dicOut("driverTitle") = m_Conn.Driver.DriverTitle
    dicOut("journalMode") = CStr(Nz(m_Conn.QueryScalar("PRAGMA journal_mode;"), ""))
    dicOut("foreignKeys") = CLng(Nz(m_Conn.QueryScalar("PRAGMA foreign_keys;"), 0))
    dicOut("fileSizeBytes") = GetSafeFileBytes(m_CurrentDbPath)
    dicOut("inTransaction") = Not m_Conn.Driver.GetAutocommit(m_Conn.DbHandle)
    Set dicOut("schema") = BuildSchemaPayload()
    Set dicOut("savedQueries") = ListSavedQueries()
    Set BuildInitPayload = dicOut
End Function

Private Function BuildSchemaPayload() As Scripting.Dictionary
    Dim dicSchema As New Scripting.Dictionary
    Dim colTables As New Collection
    Dim colViews As New Collection
    Dim colTriggers As New Collection
    Dim colMaster As Collection
    Dim dicObj As Scripting.Dictionary
    Dim strType As String
    Dim strName As String
    Dim strCreateSql As String
    Dim dicTableItem As Scripting.Dictionary
    Dim dicViewItem As Scripting.Dictionary
    Dim dicTrigItem As Scripting.Dictionary

    EnsureOpenConn
    Set colMaster = m_Conn.QueryRows( _
        "SELECT type, name, tbl_name, sql FROM sqlite_master " & _
        "WHERE name NOT LIKE 'sqlite_%' AND name <> ? AND name <> ? " & _
        "ORDER BY CASE type WHEN 'table' THEN 1 WHEN 'view' THEN 2 ELSE 3 END, LOWER(name);", _
        SAVED_QUERIES_TABLE, LEGACY_SAVED_QUERIES_TABLE)

    For Each dicObj In colMaster
        strType = LCase$(CStr(Nz(dicObj("type"), vbNullString)))
        strName = CStr(Nz(dicObj("name"), vbNullString))
        strCreateSql = CStr(Nz(dicObj("sql"), vbNullString))

        Select Case strType
            Case "table"
                Set dicTableItem = New Scripting.Dictionary
                dicTableItem("name") = strName
                dicTableItem("sql") = strCreateSql
                dicTableItem("strict") = (InStr(1, UCase$(strCreateSql), ") STRICT", vbBinaryCompare) > 0)
                dicTableItem("rowCount") = SafeTableRowCount(strName)
                Set dicTableItem("columns") = ReadObjectColumns(strName)
                Set dicTableItem("foreignKeys") = ReadTableForeignKeys(strName)
                Set dicTableItem("indexes") = ReadTableIndexes(strName)
                colTables.Add dicTableItem

            Case "view"
                Set dicViewItem = New Scripting.Dictionary
                dicViewItem("name") = strName
                dicViewItem("sql") = strCreateSql
                Set dicViewItem("columns") = ReadObjectColumns(strName)
                colViews.Add dicViewItem

            Case "trigger"
                Set dicTrigItem = New Scripting.Dictionary
                dicTrigItem("name") = strName
                dicTrigItem("tbl_name") = CStr(Nz(dicObj("tbl_name"), vbNullString))
                dicTrigItem("sql") = strCreateSql
                colTriggers.Add dicTrigItem
        End Select
    Next dicObj

    Set dicSchema("tables") = colTables
    Set dicSchema("views") = colViews
    Set dicSchema("triggers") = colTriggers
    Set BuildSchemaPayload = dicSchema
End Function

Private Function ReadObjectColumns(ByVal strObjName As String) As Collection
    Dim colOut As New Collection
    Dim colRows As Collection
    Dim dicRow As Scripting.Dictionary
    Dim dicCol As Scripting.Dictionary

    On Error GoTo ColErr
    Set colRows = m_Conn.QueryRows("PRAGMA table_xinfo(" & QuoteSQLiteIdent(strObjName) & ");")
    For Each dicRow In colRows
        If CLng(Nz(dicRow("hidden"), 0)) = 0 Then
            Set dicCol = New Scripting.Dictionary
            dicCol("cid") = CLng(Nz(dicRow("cid"), 0))
            dicCol("name") = CStr(Nz(dicRow("name"), vbNullString))
            dicCol("type") = CStr(Nz(dicRow("type"), vbNullString))
            dicCol("notnull") = (CLng(Nz(dicRow("notnull"), 0)) <> 0)
            dicCol("dflt_value") = dicRow("dflt_value")
            dicCol("pk") = (CLng(Nz(dicRow("pk"), 0)) > 0)
            colOut.Add dicCol
        End If
    Next dicRow
    Set ReadObjectColumns = colOut
    Exit Function
ColErr:
    Set ReadObjectColumns = colOut
End Function

Private Function ReadTableForeignKeys(ByVal strTableName As String) As Collection
    Dim colOut As New Collection
    Dim colRows As Collection
    Dim dicRow As Scripting.Dictionary
    Dim dicFk As Scripting.Dictionary

    On Error GoTo FkErr
    Set colRows = m_Conn.QueryRows("PRAGMA foreign_key_list(" & QuoteSQLiteIdent(strTableName) & ");")
    For Each dicRow In colRows
        Set dicFk = New Scripting.Dictionary
        dicFk("from") = CStr(Nz(dicRow("from"), vbNullString))
        dicFk("table") = CStr(Nz(dicRow("table"), vbNullString))
        dicFk("to") = CStr(Nz(dicRow("to"), vbNullString))
        dicFk("on_delete") = CStr(Nz(dicRow("on_delete"), vbNullString))
        colOut.Add dicFk
    Next dicRow
    Set ReadTableForeignKeys = colOut
    Exit Function
FkErr:
    Set ReadTableForeignKeys = colOut
End Function

Private Function ReadTableIndexes(ByVal strTableName As String) As Collection
    Dim colOut As New Collection
    Dim colIdxRows As Collection
    Dim colInfoRows As Collection
    Dim dicIdx As Scripting.Dictionary
    Dim dicInfo As Scripting.Dictionary
    Dim dicItem As Scripting.Dictionary
    Dim strIdxName As String
    Dim strCols As String

    On Error GoTo IdxErr
    Set colIdxRows = m_Conn.QueryRows("PRAGMA index_list(" & QuoteSQLiteIdent(strTableName) & ");")
    For Each dicIdx In colIdxRows
        strIdxName = CStr(Nz(dicIdx("name"), vbNullString))
        If Left$(strIdxName, 17) <> "sqlite_autoindex_" Then
            strCols = vbNullString
            Set colInfoRows = m_Conn.QueryRows("PRAGMA index_info(" & QuoteSQLiteIdent(strIdxName) & ");")
            For Each dicInfo In colInfoRows
                If Len(strCols) > 0 Then strCols = strCols & ", "
                strCols = strCols & CStr(Nz(dicInfo("name"), vbNullString))
            Next dicInfo

            Set dicItem = New Scripting.Dictionary
            dicItem("name") = strIdxName
            dicItem("unique") = (CLng(Nz(dicIdx("unique"), 0)) <> 0)
            dicItem("columns") = strCols
            dicItem("sql") = CStr(Nz(m_Conn.QueryScalar("SELECT sql FROM sqlite_master WHERE type = 'index' AND name = ?;", strIdxName), vbNullString))
            colOut.Add dicItem
        End If
    Next dicIdx
    Set ReadTableIndexes = colOut
    Exit Function
IdxErr:
    Set ReadTableIndexes = colOut
End Function

Private Function SafeTableRowCount(ByVal strTableName As String) As Long
    On Error GoTo CntErr
    Dim varCnt As Variant
    varCnt = m_Conn.QueryScalar("SELECT COUNT(*) FROM " & QuoteSQLiteIdent(strTableName) & ";")
    If Not IsNull(varCnt) Then SafeTableRowCount = CLng(varCnt)
    Exit Function
CntErr:
    SafeTableRowCount = 0
End Function

' =============================================================================
' Multi-Statement SQL Executor & Explain Plan
' =============================================================================

Private Function ExecuteAxLiteSql(ByVal strSql As String, ByVal lngRowLimit As Long) As Scripting.Dictionary
    Dim dicOut As New Scripting.Dictionary
    Dim colStatements As New Collection
    Dim colCols As Collection
    Dim colRows As Collection
    Dim colOneRow As Collection
    Dim dicColMeta As Scripting.Dictionary
    Dim dicStmtLog As Scripting.Dictionary
    Dim strPending As String
    Dim strRemainder As String
    Dim strStmtPreview As String
    Dim ptrStmt As LongPtr
    Dim lngRc As Long
    Dim lngStmtIndex As Long
    Dim lngColCount As Long
    Dim lngCol As Long
    Dim lngFetched As Long
    Dim blnTruncated As Boolean
    Dim blnHasResultSet As Boolean
    Dim blnSchemaChanged As Boolean
    Dim lngTotalChanges As Long
    Dim varLastRowId As Variant
    Dim curTotalStart As Currency
    Dim curStmtStart As Currency
    Dim dblStmtMs As Double

    EnsureOpenConn
    If Len(Trim$(strSql)) = 0 Then
        RaiseSQLiteError sqErrInvalidArgument, "clsAxLiteHost.ExecuteAxLiteSql", "SQL text cannot be empty."
    End If

    curTotalStart = StartPerfTimer()
    strPending = strSql
    Set colCols = New Collection
    Set colRows = New Collection
    varLastRowId = 0

    Do While Len(Trim$(strPending)) > 0
        ptrStmt = 0
        strRemainder = vbNullString
        curStmtStart = StartPerfTimer()

        lngRc = m_Conn.Driver.PrepareStmt(m_Conn.DbHandle, strPending, ptrStmt, sqPrepareNormal, strRemainder)
        If lngRc <> sqOk Then
            If ptrStmt <> 0 Then m_Conn.Driver.FinalizeStmt ptrStmt
            RaiseAxLiteSqliteError lngRc, strPending
        End If

        If ptrStmt <> 0 Then
            lngStmtIndex = lngStmtIndex + 1
            strStmtPreview = ExtractCurrentStmtPreview(strPending, strRemainder)
            If LooksLikeDdlStatement(strStmtPreview) Then
                m_Conn.ClearStatementCache
                blnSchemaChanged = True
            End If

            lngColCount = m_Conn.Driver.GetColumnCount(ptrStmt)
            lngFetched = 0
            blnTruncated = False

            lngRc = m_Conn.Driver.StepStmt(ptrStmt)
            If lngRc = sqRow And lngColCount = 0 Then
                lngColCount = m_Conn.Driver.GetColumnCount(ptrStmt)
            End If

            If lngColCount > 0 And (lngRc = sqRow Or lngRc = sqDone) Then
                blnHasResultSet = True
                Set colCols = New Collection
                Set colRows = New Collection

                For lngCol = 0 To lngColCount - 1
                    Set dicColMeta = New Scripting.Dictionary
                    dicColMeta("name") = m_Conn.Driver.GetColumnName(ptrStmt, lngCol)
                    If Len(CStr(dicColMeta("name"))) = 0 Then dicColMeta("name") = "Col" & (lngCol + 1)
                    dicColMeta("declType") = m_Conn.Driver.GetColumnDeclType(ptrStmt, lngCol)
                    colCols.Add dicColMeta
                Next lngCol

                Do While lngRc = sqRow
                    If lngRowLimit > 0 And lngFetched >= lngRowLimit Then
                        blnTruncated = True
                        Exit Do
                    End If
                    Set colOneRow = New Collection
                    For lngCol = 0 To lngColCount - 1
                        colOneRow.Add ReadJsonSafeColumn(ptrStmt, lngCol)
                    Next lngCol
                    colRows.Add colOneRow
                    lngFetched = lngFetched + 1
                    lngRc = m_Conn.Driver.StepStmt(ptrStmt)
                Loop

                If Not blnTruncated And lngRc <> sqDone And lngRc <> sqOk Then
                    m_Conn.Driver.FinalizeStmt ptrStmt
                    RaiseAxLiteSqliteError lngRc, strStmtPreview
                End If

                m_Conn.Driver.FinalizeStmt ptrStmt
                ptrStmt = 0
                dblStmtMs = ElapsedPerfMs(curStmtStart)

                Set dicStmtLog = New Scripting.Dictionary
                dicStmtLog("index") = lngStmtIndex
                dicStmtLog("sqlPreview") = Left$(strStmtPreview, 140)
                dicStmtLog("isQuery") = True
                dicStmtLog("rowCount") = lngFetched
                dicStmtLog("rowsAffected") = 0
                dicStmtLog("lastInsertRowId") = 0
                dicStmtLog("elapsedMs") = dblStmtMs
                colStatements.Add dicStmtLog
            Else
                If lngRc <> sqDone And lngRc <> sqOk Then
                    m_Conn.Driver.FinalizeStmt ptrStmt
                    RaiseAxLiteSqliteError lngRc, strStmtPreview
                End If
                m_Conn.Driver.FinalizeStmt ptrStmt
                ptrStmt = 0
                dblStmtMs = ElapsedPerfMs(curStmtStart)

                Dim lngAffected As Long
                lngAffected = m_Conn.ChangesCount
                lngTotalChanges = lngTotalChanges + lngAffected
                varLastRowId = m_Conn.LastInsertRowId

                Set dicStmtLog = New Scripting.Dictionary
                dicStmtLog("index") = lngStmtIndex
                dicStmtLog("sqlPreview") = Left$(strStmtPreview, 140)
                dicStmtLog("isQuery") = False
                dicStmtLog("rowCount") = 0
                dicStmtLog("rowsAffected") = lngAffected
                dicStmtLog("lastInsertRowId") = varLastRowId
                dicStmtLog("elapsedMs") = dblStmtMs
                colStatements.Add dicStmtLog
            End If
        End If

        If StrComp(strPending, strRemainder, vbBinaryCompare) = 0 Then Exit Do
        strPending = strRemainder
    Loop

    If blnSchemaChanged Then m_Conn.ClearStatementCache

    dicOut("hasResultSet") = blnHasResultSet
    Set dicOut("columns") = colCols
    Set dicOut("rows") = colRows
    dicOut("rowCount") = colRows.Count
    dicOut("truncated") = blnTruncated
    dicOut("totalChanges") = lngTotalChanges
    dicOut("lastInsertRowId") = varLastRowId
    dicOut("totalElapsedMs") = ElapsedPerfMs(curTotalStart)
    dicOut("inTransaction") = Not m_Conn.Driver.GetAutocommit(m_Conn.DbHandle)
    dicOut("schemaChanged") = blnSchemaChanged
    dicOut("fileSizeBytes") = GetSafeFileBytes(m_CurrentDbPath)
    Set dicOut("statements") = colStatements
    Set ExecuteAxLiteSql = dicOut
End Function

Private Function ExplainAxLiteSql(ByVal strSql As String) As Scripting.Dictionary
    Dim dicOut As New Scripting.Dictionary
    Dim colNodes As New Collection
    Dim colRows As Collection
    Dim dicRow As Scripting.Dictionary
    Dim dicNode As Scripting.Dictionary

    EnsureOpenConn
    Set colRows = m_Conn.QueryRows("EXPLAIN QUERY PLAN " & Trim$(strSql))
    For Each dicRow In colRows
        Set dicNode = New Scripting.Dictionary
        dicNode("id") = CLng(Nz(dicRow("id"), 0))
        dicNode("parent") = CLng(Nz(dicRow("parent"), 0))
        dicNode("notused") = CLng(Nz(dicRow("notused"), 0))
        dicNode("detail") = CStr(Nz(dicRow("detail"), vbNullString))
        colNodes.Add dicNode
    Next dicRow
    Set dicOut("nodes") = colNodes
    Set ExplainAxLiteSql = dicOut
End Function

Private Function ReadJsonSafeColumn(ByVal ptrStmt As LongPtr, ByVal lngZeroCol As Long) As Variant
    Dim lngAff As Long
    Dim abytBlob() As Byte
    Dim lngBytes As Long
    Dim lngPreview As Long
    Dim lngIdx As Long
    Dim strHex As String

    lngAff = m_Conn.Driver.GetColumnType(ptrStmt, lngZeroCol)
    Select Case lngAff
        Case sqAffinityNull
            ReadJsonSafeColumn = Null
        Case sqAffinityInteger
            ReadJsonSafeColumn = m_Conn.Driver.GetColumnInt64(ptrStmt, lngZeroCol)
        Case sqAffinityFloat
            ReadJsonSafeColumn = m_Conn.Driver.GetColumnDouble(ptrStmt, lngZeroCol)
        Case sqAffinityText
            ReadJsonSafeColumn = m_Conn.Driver.GetColumnText(ptrStmt, lngZeroCol)
        Case sqAffinityBlob
            abytBlob = m_Conn.Driver.GetColumnBlob(ptrStmt, lngZeroCol)
            lngBytes = GetByteArrayLength(abytBlob)
            If lngBytes <= 0 Then
                ReadJsonSafeColumn = "[BLOB: 0 bytes]"
            Else
                lngPreview = IIf(lngBytes > 16, 16, lngBytes)
                For lngIdx = 0 To lngPreview - 1
                    strHex = strHex & Right$("0" & Hex$(abytBlob(LBound(abytBlob) + lngIdx)), 2)
                Next lngIdx
                If lngBytes > 16 Then strHex = strHex & "..."
                ReadJsonSafeColumn = "[BLOB: " & lngBytes & " B] 0x" & strHex
            End If
        Case Else
            ReadJsonSafeColumn = Null
    End Select
End Function

' =============================================================================
' Stored Queries Persistence (_axlite_saved_queries)
' =============================================================================

Private Sub EnsureSavedQueriesTable()
    EnsureOpenConn
    m_Conn.ExecuteScript _
        "CREATE TABLE IF NOT EXISTS """ & SAVED_QUERIES_TABLE & """ (" & _
        "  QueryID INTEGER PRIMARY KEY AUTOINCREMENT," & _
        "  QueryName TEXT NOT NULL UNIQUE," & _
        "  CategoryName TEXT NOT NULL DEFAULT 'General'," & _
        "  SqlText TEXT NOT NULL," & _
        "  DescriptionText TEXT," & _
        "  UpdatedAt TEXT NOT NULL" & _
        ");"
End Sub

Private Sub EnsureStarterQueriesIfDemo()
    On Error Resume Next
    If Not m_Conn.TableExists("Tasks") Then Exit Sub
    If m_Conn.TableExists(SAVED_QUERIES_TABLE) Then Exit Sub
    EnsureSavedQueriesTable

    SaveStoredQuery "Open Tasks by Category", "Analytics", _
        "SELECT CategoryName," & vbCrLf & _
        "       COUNT(*) AS OpenTasks," & vbCrLf & _
        "       ROUND(SUM(EstimateCost), 2) AS TotalOpenBudget" & vbCrLf & _
        "FROM Tasks" & vbCrLf & _
        "WHERE DeletedAt IS NULL AND IsCompleted = 0" & vbCrLf & _
        "GROUP BY CategoryName" & vbCrLf & _
        "ORDER BY TotalOpenBudget DESC;", _
        "Summarizes active open tasks and estimated cost by category"

    SaveStoredQuery "Critical & High Priority Schedule", "Operations", _
        "SELECT TaskID, TaskTitle, CategoryName, PriorityLevel, EstimateCost, DueDate" & vbCrLf & _
        "FROM Tasks" & vbCrLf & _
        "WHERE DeletedAt IS NULL" & vbCrLf & _
        "  AND IsCompleted = 0" & vbCrLf & _
        "  AND PriorityLevel IN ('Critical', 'High')" & vbCrLf & _
        "ORDER BY CASE PriorityLevel WHEN 'Critical' THEN 1 ELSE 2 END, DueDate ASC;", _
        "Lists open Critical and High priority tasks ordered by urgency"

    SaveStoredQuery "SQLite Engine & Active Pragmas", "Diagnostics", _
        "SELECT sqlite_version() AS SQLiteVersion," & vbCrLf & _
        "       (SELECT journal_mode FROM pragma_journal_mode) AS JournalMode," & vbCrLf & _
        "       (SELECT foreign_keys FROM pragma_foreign_keys) AS ForeignKeys," & vbCrLf & _
        "       (SELECT page_size FROM pragma_page_size) AS PageSize," & vbCrLf & _
        "       (SELECT page_count FROM pragma_page_count) AS PageCount;", _
        "Inspects active SQLite engine version and connection pragma settings"
    On Error GoTo 0
End Sub

Private Function ListSavedQueries() As Collection
    Dim colOut As New Collection
    Dim colRows As Collection
    Dim dicRow As Scripting.Dictionary
    Dim dicItem As Scripting.Dictionary

    EnsureOpenConn
    If Not m_Conn.TableExists(SAVED_QUERIES_TABLE) Then
        Set ListSavedQueries = colOut
        Exit Function
    End If

    Set colRows = m_Conn.QueryRows( _
        "SELECT QueryID, QueryName, CategoryName, SqlText, DescriptionText, UpdatedAt " & _
        "FROM """ & SAVED_QUERIES_TABLE & """ ORDER BY LOWER(CategoryName), LOWER(QueryName);")
    For Each dicRow In colRows
        Set dicItem = New Scripting.Dictionary
        dicItem("id") = CLng(Nz(dicRow("QueryID"), 0))
        dicItem("name") = CStr(Nz(dicRow("QueryName"), vbNullString))
        dicItem("category") = CStr(Nz(dicRow("CategoryName"), "General"))
        dicItem("sql") = CStr(Nz(dicRow("SqlText"), vbNullString))
        dicItem("description") = CStr(Nz(dicRow("DescriptionText"), vbNullString))
        dicItem("updatedAt") = CStr(Nz(dicRow("UpdatedAt"), vbNullString))
        colOut.Add dicItem
    Next dicRow
    Set ListSavedQueries = colOut
End Function

Private Sub SaveStoredQuery( _
    ByVal strName As String, _
    ByVal strCategory As String, _
    ByVal strSqlText As String, _
    ByVal strDesc As String)

    Dim strCleanName As String
    Dim strCleanCat As String

    strCleanName = Trim$(strName)
    If Len(strCleanName) = 0 Then
        RaiseSQLiteError sqErrValidation, "clsAxLiteHost.SaveStoredQuery", "Stored query name is required."
    End If
    If Len(Trim$(strSqlText)) = 0 Then
        RaiseSQLiteError sqErrValidation, "clsAxLiteHost.SaveStoredQuery", "Stored query SQL cannot be empty."
    End If
    strCleanCat = IIf(Len(Trim$(strCategory)) = 0, "General", Trim$(strCategory))

    EnsureSavedQueriesTable
    m_Conn.ExecuteSql _
        "INSERT INTO """ & SAVED_QUERIES_TABLE & """ (QueryName, CategoryName, SqlText, DescriptionText, UpdatedAt) " & _
        "VALUES (?, ?, ?, ?, ?) " & _
        "ON CONFLICT(QueryName) DO UPDATE SET " & _
        "  CategoryName = excluded.CategoryName," & _
        "  SqlText = excluded.SqlText," & _
        "  DescriptionText = excluded.DescriptionText," & _
        "  UpdatedAt = excluded.UpdatedAt;", _
        strCleanName, strCleanCat, strSqlText, Trim$(strDesc), UtcNowIsoString()
End Sub

Private Sub DeleteStoredQuery(ByVal strName As String)
    EnsureOpenConn
    If Not m_Conn.TableExists(SAVED_QUERIES_TABLE) Then Exit Sub
    m_Conn.ExecuteSql "DELETE FROM """ & SAVED_QUERIES_TABLE & """ WHERE QueryName = ?;", Trim$(strName)
End Sub

' =============================================================================
' File Dialogs & Database Maintenance Actions
' =============================================================================

Private Function HandleOpenDatabaseDialog(ByVal strMode As String) As Scripting.Dictionary
    Dim dicOut As New Scripting.Dictionary
    Dim objDlg As Object
    Dim strPicked As String

    On Error GoTo DialogErr
    If StrComp(strMode, "create", vbTextCompare) = 0 Then
        Set objDlg = Application.FileDialog(MSO_FILE_DIALOG_SAVEAS)
        objDlg.Title = "Create New SQLite Database (.db)"
        objDlg.InitialFileName = CurrentProject.Path & "\NewDatabase.db"
    Else
        Set objDlg = Application.FileDialog(MSO_FILE_DIALOG_OPEN)
        objDlg.Title = "Open SQLite Database in AxLite"
        objDlg.AllowMultiSelect = False
        objDlg.InitialFileName = CurrentProject.Path & "\"
        On Error Resume Next
        objDlg.Filters.Clear
        objDlg.Filters.Add "SQLite Databases (*.db; *.sqlite; *.sqlite3)", "*.db;*.sqlite;*.sqlite3"
        objDlg.Filters.Add "All Files (*.*)", "*.*"
        On Error GoTo DialogErr
    End If

    If Not objDlg.Show Then
        dicOut("cancelled") = True
        Set HandleOpenDatabaseDialog = dicOut
        Exit Function
    End If

    strPicked = CStr(objDlg.SelectedItems(1))
    If StrComp(strMode, "create", vbTextCompare) = 0 Then
        If InStr(1, ExtractLeafName(strPicked), ".", vbBinaryCompare) = 0 Then
            strPicked = strPicked & ".db"
        End If
    End If

    OpenTargetDatabase strPicked
    Set HandleOpenDatabaseDialog = BuildInitPayload()
    Exit Function

DialogErr:
    RaiseSQLiteError sqErrInvalidArgument, "clsAxLiteHost.HandleOpenDatabaseDialog", Err.Description
End Function

Private Function HandleDatabaseAction(ByVal dicPayload As Scripting.Dictionary) As Scripting.Dictionary
    Dim dicOut As New Scripting.Dictionary
    Dim strKind As String
    Dim strRes As String
    Dim objDlg As Object
    Dim strTarget As String
    Dim abytUtf8() As Byte
    Dim intFileNum As Integer

    EnsureOpenConn
    strKind = LCase$(Trim$(CStr(Nz(dicPayload("kind"), vbNullString))))

    Select Case strKind
        Case "integrity_check"
            strRes = CStr(Nz(m_Conn.QueryScalar("PRAGMA integrity_check;"), "unknown"))
            dicOut("ok") = (StrComp(strRes, "ok", vbTextCompare) = 0)
            dicOut("message") = "PRAGMA integrity_check: " & strRes

        Case "vacuum"
            m_Conn.VacuumDb
            dicOut("ok") = True
            dicOut("fileSizeBytes") = GetSafeFileBytes(m_CurrentDbPath)
            dicOut("message") = "VACUUM completed successfully."

        Case "backup"
            Set objDlg = Application.FileDialog(MSO_FILE_DIALOG_SAVEAS)
            objDlg.Title = "Save Online SQLite Backup As"
            objDlg.InitialFileName = CurrentProject.Path & "\" & Replace(ExtractLeafName(m_CurrentDbPath), ".db", "") & "_backup.db"
            If Not objDlg.Show Then
                dicOut("cancelled") = True
                Set HandleDatabaseAction = dicOut
                Exit Function
            End If
            strTarget = CStr(objDlg.SelectedItems(1))
            If InStr(1, ExtractLeafName(strTarget), ".", vbBinaryCompare) = 0 Then strTarget = strTarget & ".db"
            m_Conn.BackupToFile strTarget
            dicOut("ok") = True
            dicOut("message") = "Online backup written to: " & strTarget

        Case "export_csv"
            Set objDlg = Application.FileDialog(MSO_FILE_DIALOG_SAVEAS)
            objDlg.Title = "Export Query Results as CSV"
            objDlg.InitialFileName = CurrentProject.Path & "\" & CStr(Nz(dicPayload("suggestedName"), "query_results.csv"))
            If Not objDlg.Show Then
                dicOut("cancelled") = True
                Set HandleDatabaseAction = dicOut
                Exit Function
            End If
            strTarget = CStr(objDlg.SelectedItems(1))
            If InStr(1, ExtractLeafName(strTarget), ".", vbBinaryCompare) = 0 Then strTarget = strTarget & ".csv"
            abytUtf8 = VbaStringToUtf8Bytes(CStr(Nz(dicPayload("csvText"), vbNullString)))
            If Len(Dir$(strTarget)) > 0 Then Kill strTarget
            intFileNum = FreeFile
            Open strTarget For Binary Access Write As #intFileNum
            If UBound(abytUtf8) > 0 Then
                ReDim Preserve abytUtf8(0 To UBound(abytUtf8) - 1) ' strip trailing C null byte
                Put #intFileNum, , abytUtf8
            End If
            Close #intFileNum
            dicOut("ok") = True
            dicOut("message") = "Exported CSV to: " & strTarget

        Case Else
            RaiseSQLiteError sqErrInvalidArgument, "clsAxLiteHost.HandleDatabaseAction", _
                "Unknown database maintenance action: " & strKind
    End Select

    Set HandleDatabaseAction = dicOut
End Function

' =============================================================================
' Utility Helpers
' =============================================================================

Private Function ResolveInitialDbPath(ByVal strCandidate As String) As String
    Dim strClean As String
    Dim strCurDir As String
    Dim strLibDir As String

    strClean = Trim$(strCandidate)
    If Len(strClean) > 0 Then
        ResolveInitialDbPath = strClean
        Exit Function
    End If

    On Error Resume Next
    strCurDir = CurrentProject.Path
    strLibDir = GetLibraryDirectory()
    On Error GoTo 0

    If Len(strCurDir) > 0 And Len(Dir$(strCurDir & "\AccessSQLiteDemo.db")) > 0 Then
        ResolveInitialDbPath = strCurDir & "\AccessSQLiteDemo.db"
    ElseIf Len(strLibDir) > 0 And Len(Dir$(strLibDir & "\AccessSQLiteDemo.db")) > 0 Then
        ResolveInitialDbPath = strLibDir & "\AccessSQLiteDemo.db"
    ElseIf Len(strLibDir) > 0 And Len(Dir$(strLibDir & "\AccessSQLite_Test.db")) > 0 Then
        ResolveInitialDbPath = strLibDir & "\AccessSQLite_Test.db"
    Else
        ResolveInitialDbPath = ":memory:"
    End If
End Function

Private Function AxLiteHtmlFilePath() As String
    Dim strRoot As String
    strRoot = GetLibraryDirectory()
    If Right$(strRoot, 1) = "\" Then strRoot = Left$(strRoot, Len(strRoot) - 1)
    AxLiteHtmlFilePath = strRoot & "\web\axlite\index.html"
End Function

Private Function ExtractLeafName(ByVal strPath As String) As String
    Dim lngSlash As Long
    If StrComp(strPath, ":memory:", vbTextCompare) = 0 Then
        ExtractLeafName = ":memory:"
        Exit Function
    End If
    lngSlash = InStrRev(strPath, "\")
    If lngSlash > 0 Then
        ExtractLeafName = Mid$(strPath, lngSlash + 1)
    Else
        ExtractLeafName = strPath
    End If
End Function

Private Function GetSafeFileBytes(ByVal strPath As String) As Double
    On Error GoTo SizeErr
    If Len(strPath) = 0 Or StrComp(strPath, ":memory:", vbTextCompare) = 0 Then Exit Function
    If Len(Dir$(strPath)) = 0 Then Exit Function
    GetSafeFileBytes = CDbl(FileLen(strPath))
    Exit Function
SizeErr:
    GetSafeFileBytes = 0#
End Function

Private Function ExtractCurrentStmtPreview(ByVal strFull As String, ByVal strRemainder As String) As String
    Dim lngConsumed As Long
    Dim strChunk As String
    lngConsumed = Len(strFull) - Len(strRemainder)
    If lngConsumed > 0 Then
        strChunk = Left$(strFull, lngConsumed)
    Else
        strChunk = strFull
    End If
    strChunk = Replace(Replace(Trim$(strChunk), vbCrLf, " "), vbLf, " ")
    ExtractCurrentStmtPreview = strChunk
End Function

Private Function LooksLikeDdlStatement(ByVal strSqlPreview As String) As Boolean
    Dim strUpper As String
    strUpper = UCase$(LTrim$(strSqlPreview))
    If Left$(strUpper, 7) = "CREATE " Or Left$(strUpper, 6) = "ALTER " Or _
       Left$(strUpper, 5) = "DROP " Or Left$(strUpper, 7) = "ATTACH " Or _
       Left$(strUpper, 7) = "DETACH " Or Left$(strUpper, 6) = "VACUUM" Or _
       Left$(strUpper, 7) = "REINDEX" Then
        LooksLikeDdlStatement = True
    End If
End Function

Private Sub RaiseAxLiteSqliteError(ByVal lngRc As Long, ByVal strSqlContext As String)
    Dim lngExt As Long
    Dim strMsg As String
    lngExt = m_Conn.Driver.GetExtendedErrCode(m_Conn.DbHandle)
    strMsg = m_Conn.Driver.GetErrMsg(m_Conn.DbHandle)
    If Len(strMsg) = 0 Then strMsg = "SQLite statement failed."
    RaiseSQLiteError IIf(lngRc = sqConstraint, sqErrConstraint, sqErrSqlFailed), _
        "clsAxLiteHost", strMsg & " | SQL: " & Left$(Trim$(strSqlContext), 180), lngRc, lngExt
End Sub

Private Function MapErrNumberToCode(ByVal lngErrNum As Long) As String
    Select Case lngErrNum
        Case sqErrConstraint: MapErrNumberToCode = "ConstraintError"
        Case sqErrValidation: MapErrNumberToCode = "ValidationError"
        Case sqErrNotFound: MapErrNumberToCode = "NotFound"
        Case sqErrInvalidArgument: MapErrNumberToCode = "BadRequest"
        Case Else: MapErrNumberToCode = "SQLiteError"
    End Select
End Function

Private Function SecondsElapsed(ByVal dblStarted As Double) As Double
    SecondsElapsed = Timer - dblStarted
    If SecondsElapsed < 0# Then SecondsElapsed = SecondsElapsed + 86400#
End Function

Private Sub EnsureOpenConn()
    If m_Conn Is Nothing Or Not m_Conn.IsOpen Then
        OpenTargetDatabase ResolveInitialDbPath(m_CurrentDbPath)
    End If
End Sub
