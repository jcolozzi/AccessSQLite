Attribute VB_Name = "clsSQLiteMigrator"
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
' clsSQLiteMigrator
' Versioned SQLite schema migration runner with FNV-1a checksum verification,
' transactional up/down execution, and _schema_migrations tracking table.
' =============================================================================

Private Const MIGRATIONS_TABLE As String = "_schema_migrations"

Private m_Conn As clsSQLiteConnection
Private m_StepsByVer As Scripting.Dictionary

Private Sub Class_Initialize()
    Set m_StepsByVer = New Scripting.Dictionary
End Sub

Public Sub InitMigrator(ByVal objConn As clsSQLiteConnection)
    Set m_Conn = objConn
    Set m_StepsByVer = New Scripting.Dictionary
End Sub

Public Function AddMigration( _
    ByVal lngVersionNum As Long, _
    ByVal strTitle As String, _
    ByVal strUpSql As String, _
    Optional ByVal strDownSql As String = vbNullString) As clsSQLiteMigrator

    Dim dictStep As Scripting.Dictionary
    If lngVersionNum <= 0 Then
        RaiseSQLiteError sqErrMigration, "clsSQLiteMigrator.AddMigration", "Migration version must be >= 1."
    End If
    If Len(Trim$(strUpSql)) = 0 Then
        RaiseSQLiteError sqErrMigration, "clsSQLiteMigrator.AddMigration", "Migration UpSql cannot be empty for version " & lngVersionNum & "."
    End If
    If m_StepsByVer.Exists(lngVersionNum) Then
        RaiseSQLiteError sqErrMigration, "clsSQLiteMigrator.AddMigration", "Duplicate migration version " & lngVersionNum & "."
    End If

    Set dictStep = NewTextDict()
    dictStep("VersionNum") = lngVersionNum
    dictStep("MigrationTitle") = Trim$(strTitle)
    dictStep("UpSql") = strUpSql
    dictStep("DownSql") = strDownSql
    dictStep("ChecksumHex") = ComputeFnv1aHex(strUpSql)
    Set m_StepsByVer(lngVersionNum) = dictStep
    Set AddMigration = Me
End Function

Public Function CurrentVersion() As Long
    Dim varMax As Variant
    EnsureHistoryTable
    varMax = m_Conn.QueryScalar("SELECT COALESCE(MAX(""VersionNum""), 0) FROM """ & MIGRATIONS_TABLE & """;")
    If Not IsNull(varMax) Then CurrentVersion = CLng(varMax)
End Function

Public Function AppliedMigrations() As Collection
    EnsureHistoryTable
    Set AppliedMigrations = m_Conn.QueryRows( _
        "SELECT ""VersionNum"", ""MigrationTitle"", ""ChecksumHex"", ""AppliedAtIso"", ""DurationMs"" " & _
        "FROM """ & MIGRATIONS_TABLE & """ ORDER BY ""VersionNum"" ASC;")
End Function

Public Sub VerifyChecksums()
    Dim colApplied As Collection
    Dim dictRow As Scripting.Dictionary
    Dim lngVer As Long
    Dim dictStep As Scripting.Dictionary
    Dim strExpected As String
    Dim strActual As String

    Set colApplied = AppliedMigrations()
    For Each dictRow In colApplied
        lngVer = CLng(dictRow("VersionNum"))
        If m_StepsByVer.Exists(lngVer) Then
            Set dictStep = m_StepsByVer(lngVer)
            strExpected = CStr(dictStep("ChecksumHex"))
            strActual = CStr(dictRow("ChecksumHex"))
            If StrComp(strExpected, strActual, vbTextCompare) <> 0 Then
                RaiseSQLiteError sqErrMigration, "clsSQLiteMigrator.VerifyChecksums", _
                    "Checksum mismatch on applied migration v" & lngVer & " ('" & CStr(dictRow("MigrationTitle")) & _
                    "'): stored=" & strActual & ", code=" & strExpected
            End If
        End If
    Next dictRow
End Sub

Public Function MigrateUp(Optional ByVal lngTargetVersion As Long = -1) As Long
    Dim colApplied As Collection
    Dim dictAppliedSet As Scripting.Dictionary
    Dim dictRow As Scripting.Dictionary
    Dim alngVersions() As Long
    Dim lngIdx As Long
    Dim lngVer As Long
    Dim dictStep As Scripting.Dictionary
    Dim curTimer As Currency
    Dim lngElapsedMs As Long
    Dim lngAppliedCount As Long

    EnsureHistoryTable
    VerifyChecksums
    If m_StepsByVer.Count = 0 Then Exit Function

    Set colApplied = AppliedMigrations()
    Set dictAppliedSet = New Scripting.Dictionary
    For Each dictRow In colApplied
        dictAppliedSet(CLng(dictRow("VersionNum"))) = True
    Next dictRow

    alngVersions = GetSortedRegisteredVersions()
    For lngIdx = LBound(alngVersions) To UBound(alngVersions)
        lngVer = alngVersions(lngIdx)
        If lngTargetVersion <= 0 Or lngVer <= lngTargetVersion Then
            If Not dictAppliedSet.Exists(lngVer) Then
                Set dictStep = m_StepsByVer(lngVer)
                curTimer = StartPerfTimer()

                m_Conn.BeginTrans
                On Error GoTo UpFailed
                m_Conn.ExecuteScript CStr(dictStep("UpSql"))
                lngElapsedMs = CLng(Fix(ElapsedPerfMs(curTimer) + 0.5))
                m_Conn.ExecuteSql _
                    "INSERT INTO """ & MIGRATIONS_TABLE & """ " & _
                    "(""VersionNum"", ""MigrationTitle"", ""ChecksumHex"", ""AppliedAtIso"", ""DurationMs"") " & _
                    "VALUES (?, ?, ?, ?, ?);", _
                    lngVer, CStr(dictStep("MigrationTitle")), CStr(dictStep("ChecksumHex")), UtcNowIsoString(), lngElapsedMs
                m_Conn.CommitTrans
                On Error GoTo 0

                lngAppliedCount = lngAppliedCount + 1
            End If
        End If
    Next lngIdx

    MigrateUp = lngAppliedCount
    Exit Function

UpFailed:
    Dim lngErrNum As Long
    Dim strErrSrc As String
    Dim strErrDesc As String
    lngErrNum = Err.Number
    strErrSrc = Err.Source
    strErrDesc = Err.Description
    m_Conn.RollbackTrans
    RaiseSQLiteError sqErrMigration, "clsSQLiteMigrator.MigrateUp", _
        "Migration v" & lngVer & " failed: " & strErrDesc & " (source=" & strErrSrc & ", err=" & lngErrNum & ")"
End Function

Public Function RollbackLast(Optional ByVal lngStepCount As Long = 1) As Long
    Dim colAppliedDesc As Collection
    Dim dictRow As Scripting.Dictionary
    Dim lngVer As Long
    Dim dictStep As Scripting.Dictionary
    Dim strDownSql As String
    Dim lngRolledBack As Long

    If lngStepCount <= 0 Then Exit Function
    EnsureHistoryTable

    Set colAppliedDesc = m_Conn.QueryRows( _
        "SELECT ""VersionNum"", ""MigrationTitle"" FROM """ & MIGRATIONS_TABLE & """ " & _
        "ORDER BY ""VersionNum"" DESC LIMIT ?;", lngStepCount)

    For Each dictRow In colAppliedDesc
        lngVer = CLng(dictRow("VersionNum"))
        If Not m_StepsByVer.Exists(lngVer) Then
            RaiseSQLiteError sqErrMigration, "clsSQLiteMigrator.RollbackLast", _
                "Cannot rollback v" & lngVer & ": migration definition is not registered."
        End If
        Set dictStep = m_StepsByVer(lngVer)
        strDownSql = Trim$(CStr(dictStep("DownSql")))
        If Len(strDownSql) = 0 Then
            RaiseSQLiteError sqErrMigration, "clsSQLiteMigrator.RollbackLast", _
                "Cannot rollback v" & lngVer & " ('" & CStr(dictStep("MigrationTitle")) & "'): DownSql is empty."
        End If

        m_Conn.BeginTrans
        On Error GoTo DownFailed
        m_Conn.ExecuteScript strDownSql
        m_Conn.ExecuteSql "DELETE FROM """ & MIGRATIONS_TABLE & """ WHERE ""VersionNum"" = ?;", lngVer
        m_Conn.CommitTrans
        On Error GoTo 0

        lngRolledBack = lngRolledBack + 1
    Next dictRow

    RollbackLast = lngRolledBack
    Exit Function

DownFailed:
    Dim lngErrNum As Long
    Dim strErrSrc As String
    Dim strErrDesc As String
    lngErrNum = Err.Number
    strErrSrc = Err.Source
    strErrDesc = Err.Description
    m_Conn.RollbackTrans
    RaiseSQLiteError sqErrMigration, "clsSQLiteMigrator.RollbackLast", _
        "Rollback of v" & lngVer & " failed: " & strErrDesc & " (source=" & strErrSrc & ", err=" & lngErrNum & ")"
End Function

Private Sub EnsureHistoryTable()
    If m_Conn Is Nothing Then
        RaiseSQLiteError sqErrConnectionClosed, "clsSQLiteMigrator", "Migrator is not attached to a clsSQLiteConnection."
    End If
    m_Conn.ExecuteScript _
        "CREATE TABLE IF NOT EXISTS """ & MIGRATIONS_TABLE & """ (" & _
        """VersionNum"" INTEGER PRIMARY KEY, " & _
        """MigrationTitle"" TEXT NOT NULL, " & _
        """ChecksumHex"" TEXT NOT NULL, " & _
        """AppliedAtIso"" TEXT NOT NULL, " & _
        """DurationMs"" INTEGER NOT NULL);"
End Sub

Private Function GetSortedRegisteredVersions() As Long()
    Dim alngOut() As Long
    Dim varKey As Variant
    Dim lngCount As Long
    Dim lngI As Long
    Dim lngJ As Long
    Dim lngTmp As Long

    ReDim alngOut(0 To m_StepsByVer.Count - 1)
    For Each varKey In m_StepsByVer.Keys
        alngOut(lngCount) = CLng(varKey)
        lngCount = lngCount + 1
    Next varKey

    For lngI = 0 To lngCount - 2
        For lngJ = lngI + 1 To lngCount - 1
            If alngOut(lngJ) < alngOut(lngI) Then
                lngTmp = alngOut(lngI)
                alngOut(lngI) = alngOut(lngJ)
                alngOut(lngJ) = lngTmp
            End If
        Next lngJ
    Next lngI

    GetSortedRegisteredVersions = alngOut
End Function

