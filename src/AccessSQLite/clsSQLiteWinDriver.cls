Attribute VB_Name = "clsSQLiteWinDriver"
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
' clsSQLiteWinDriver
' Zero-deployment native SQLite driver targeting Windows winsqlite3.dll.
' Works identically on 32-bit (SysWOW64) and 64-bit (System32) Microsoft Access.
' =============================================================================

Implements ISQLiteDriver

Private Const SQLITE_TRANSIENT As LongPtr = -1

Private Declare PtrSafe Function LoadLibraryW Lib "kernel32" (ByVal ptrLibFileName As LongPtr) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal ptrLibModule As LongPtr) As Long

Private Declare PtrSafe Function ws_libversion Lib "winsqlite3.dll" Alias "sqlite3_libversion" () As LongPtr
Private Declare PtrSafe Function ws_open16 Lib "winsqlite3.dll" Alias "sqlite3_open16" ( _
    ByVal ptrFileName As LongPtr, ByRef ptrDbOut As LongPtr) As Long
Private Declare PtrSafe Function ws_open_v2 Lib "winsqlite3.dll" Alias "sqlite3_open_v2" ( _
    ByVal ptrFileNameUtf8 As LongPtr, ByRef ptrDbOut As LongPtr, ByVal lngFlags As Long, ByVal ptrVfsUtf8 As LongPtr) As Long
Private Declare PtrSafe Function ws_close_v2 Lib "winsqlite3.dll" Alias "sqlite3_close_v2" (ByVal ptrDb As LongPtr) As Long

Private Declare PtrSafe Function ws_prepare16_v3 Lib "winsqlite3.dll" Alias "sqlite3_prepare16_v3" ( _
    ByVal ptrDb As LongPtr, ByVal ptrSql As LongPtr, ByVal lngByteLen As Long, _
    ByVal lngPrepFlags As Long, ByRef ptrStmtOut As LongPtr, ByRef ptrTailOut As LongPtr) As Long
Private Declare PtrSafe Function ws_step Lib "winsqlite3.dll" Alias "sqlite3_step" (ByVal ptrStmt As LongPtr) As Long
Private Declare PtrSafe Function ws_reset Lib "winsqlite3.dll" Alias "sqlite3_reset" (ByVal ptrStmt As LongPtr) As Long
Private Declare PtrSafe Function ws_clear_bindings Lib "winsqlite3.dll" Alias "sqlite3_clear_bindings" (ByVal ptrStmt As LongPtr) As Long
Private Declare PtrSafe Function ws_finalize Lib "winsqlite3.dll" Alias "sqlite3_finalize" (ByVal ptrStmt As LongPtr) As Long

Private Declare PtrSafe Function ws_bind_null Lib "winsqlite3.dll" Alias "sqlite3_bind_null" ( _
    ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long) As Long
Private Declare PtrSafe Function ws_bind_int Lib "winsqlite3.dll" Alias "sqlite3_bind_int" ( _
    ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByVal lngVal As Long) As Long
#If Win64 Then
Private Declare PtrSafe Function ws_bind_int64 Lib "winsqlite3.dll" Alias "sqlite3_bind_int64" ( _
    ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByVal lngVal64 As LongLong) As Long
#Else
Private Declare PtrSafe Function ws_bind_int64 Lib "winsqlite3.dll" Alias "sqlite3_bind_int64" ( _
    ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByVal curRawBits As Currency) As Long
#End If
Private Declare PtrSafe Function ws_bind_double Lib "winsqlite3.dll" Alias "sqlite3_bind_double" ( _
    ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByVal dblVal As Double) As Long
Private Declare PtrSafe Function ws_bind_text16 Lib "winsqlite3.dll" Alias "sqlite3_bind_text16" ( _
    ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByVal ptrWideText As LongPtr, _
    ByVal lngByteLen As Long, ByVal ptrDestructor As LongPtr) As Long
Private Declare PtrSafe Function ws_bind_blob Lib "winsqlite3.dll" Alias "sqlite3_bind_blob" ( _
    ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByVal ptrData As LongPtr, _
    ByVal lngByteLen As Long, ByVal ptrDestructor As LongPtr) As Long
Private Declare PtrSafe Function ws_bind_zeroblob Lib "winsqlite3.dll" Alias "sqlite3_bind_zeroblob" ( _
    ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByVal lngByteLen As Long) As Long

Private Declare PtrSafe Function ws_bind_parameter_count Lib "winsqlite3.dll" Alias "sqlite3_bind_parameter_count" ( _
    ByVal ptrStmt As LongPtr) As Long
Private Declare PtrSafe Function ws_bind_parameter_index Lib "winsqlite3.dll" Alias "sqlite3_bind_parameter_index" ( _
    ByVal ptrStmt As LongPtr, ByVal ptrNameUtf8 As LongPtr) As Long
Private Declare PtrSafe Function ws_bind_parameter_name Lib "winsqlite3.dll" Alias "sqlite3_bind_parameter_name" ( _
    ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long) As LongPtr

Private Declare PtrSafe Function ws_column_count Lib "winsqlite3.dll" Alias "sqlite3_column_count" (ByVal ptrStmt As LongPtr) As Long
Private Declare PtrSafe Function ws_column_type Lib "winsqlite3.dll" Alias "sqlite3_column_type" ( _
    ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As Long
Private Declare PtrSafe Function ws_column_name16 Lib "winsqlite3.dll" Alias "sqlite3_column_name16" ( _
    ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As LongPtr
Private Declare PtrSafe Function ws_column_decltype16 Lib "winsqlite3.dll" Alias "sqlite3_column_decltype16" ( _
    ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As LongPtr
Private Declare PtrSafe Function ws_column_int Lib "winsqlite3.dll" Alias "sqlite3_column_int" ( _
    ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As Long
#If Win64 Then
Private Declare PtrSafe Function ws_column_int64 Lib "winsqlite3.dll" Alias "sqlite3_column_int64" ( _
    ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As LongLong
#Else
Private Declare PtrSafe Function ws_column_int64 Lib "winsqlite3.dll" Alias "sqlite3_column_int64" ( _
    ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As Currency
#End If
Private Declare PtrSafe Function ws_column_double Lib "winsqlite3.dll" Alias "sqlite3_column_double" ( _
    ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As Double
Private Declare PtrSafe Function ws_column_text16 Lib "winsqlite3.dll" Alias "sqlite3_column_text16" ( _
    ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As LongPtr
Private Declare PtrSafe Function ws_column_bytes16 Lib "winsqlite3.dll" Alias "sqlite3_column_bytes16" ( _
    ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As Long
Private Declare PtrSafe Function ws_column_blob Lib "winsqlite3.dll" Alias "sqlite3_column_blob" ( _
    ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As LongPtr
Private Declare PtrSafe Function ws_column_bytes Lib "winsqlite3.dll" Alias "sqlite3_column_bytes" ( _
    ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As Long

Private Declare PtrSafe Function ws_changes Lib "winsqlite3.dll" Alias "sqlite3_changes" (ByVal ptrDb As LongPtr) As Long
Private Declare PtrSafe Function ws_total_changes Lib "winsqlite3.dll" Alias "sqlite3_total_changes" (ByVal ptrDb As LongPtr) As Long
#If Win64 Then
Private Declare PtrSafe Function ws_last_insert_rowid Lib "winsqlite3.dll" Alias "sqlite3_last_insert_rowid" (ByVal ptrDb As LongPtr) As LongLong
#Else
Private Declare PtrSafe Function ws_last_insert_rowid Lib "winsqlite3.dll" Alias "sqlite3_last_insert_rowid" (ByVal ptrDb As LongPtr) As Currency
#End If
Private Declare PtrSafe Function ws_get_autocommit Lib "winsqlite3.dll" Alias "sqlite3_get_autocommit" (ByVal ptrDb As LongPtr) As Long
Private Declare PtrSafe Function ws_busy_timeout Lib "winsqlite3.dll" Alias "sqlite3_busy_timeout" (ByVal ptrDb As LongPtr, ByVal lngMs As Long) As Long

Private Declare PtrSafe Function ws_errcode Lib "winsqlite3.dll" Alias "sqlite3_errcode" (ByVal ptrDb As LongPtr) As Long
Private Declare PtrSafe Function ws_extended_errcode Lib "winsqlite3.dll" Alias "sqlite3_extended_errcode" (ByVal ptrDb As LongPtr) As Long
Private Declare PtrSafe Function ws_errmsg16 Lib "winsqlite3.dll" Alias "sqlite3_errmsg16" (ByVal ptrDb As LongPtr) As LongPtr
Private Declare PtrSafe Function ws_expanded_sql Lib "winsqlite3.dll" Alias "sqlite3_expanded_sql" (ByVal ptrStmt As LongPtr) As LongPtr
Private Declare PtrSafe Sub ws_free Lib "winsqlite3.dll" Alias "sqlite3_free" (ByVal ptrMem As LongPtr)

Private Declare PtrSafe Function ws_backup_init Lib "winsqlite3.dll" Alias "sqlite3_backup_init" ( _
    ByVal ptrDestDb As LongPtr, ByVal ptrDestNameUtf8 As LongPtr, _
    ByVal ptrSourceDb As LongPtr, ByVal ptrSourceNameUtf8 As LongPtr) As LongPtr
Private Declare PtrSafe Function ws_backup_step Lib "winsqlite3.dll" Alias "sqlite3_backup_step" ( _
    ByVal ptrBackup As LongPtr, ByVal lngPages As Long) As Long
Private Declare PtrSafe Function ws_backup_finish Lib "winsqlite3.dll" Alias "sqlite3_backup_finish" ( _
    ByVal ptrBackup As LongPtr) As Long

Private m_LibHandle As LongPtr
Private m_LastInitErr As String
Private m_ZeroWideChar As Integer

Private Sub Class_Terminate()
    ISQLiteDriver_ReleaseDriver
End Sub

Private Property Get ISQLiteDriver_DriverTitle() As String
    ISQLiteDriver_DriverTitle = "WinSQLite3 (Windows System DLL)"
End Property

Private Property Get ISQLiteDriver_LibraryVersion() As String
    If Not ISQLiteDriver_IsInitialized Then
        If Not ISQLiteDriver_InitializeDriver() Then Exit Property
    End If
    ISQLiteDriver_LibraryVersion = Utf8PtrToVbaString(ws_libversion())
End Property

Private Property Get ISQLiteDriver_IsInitialized() As Boolean
    ISQLiteDriver_IsInitialized = (m_LibHandle <> 0)
End Property

Private Property Get ISQLiteDriver_LastInitError() As String
    ISQLiteDriver_LastInitError = m_LastInitErr
End Property

Private Function ISQLiteDriver_InitializeDriver(Optional ByVal strLibDirectory As String = vbNullString) As Boolean
    Dim lngWinErr As Long
    If m_LibHandle <> 0 Then
        ISQLiteDriver_InitializeDriver = True
        Exit Function
    End If
    m_LibHandle = LoadLibraryW(StrPtr("winsqlite3.dll"))
    lngWinErr = Err.LastDllError
    If m_LibHandle = 0 Then
        m_LastInitErr = "Failed to load winsqlite3.dll: " & FormatWin32Error(lngWinErr)
        ISQLiteDriver_InitializeDriver = False
        Exit Function
    End If
    m_LastInitErr = vbNullString
    ISQLiteDriver_InitializeDriver = True
End Function

Private Sub ISQLiteDriver_ReleaseDriver()
    If m_LibHandle <> 0 Then
        FreeLibrary m_LibHandle
        m_LibHandle = 0
    End If
End Sub

Private Function ISQLiteDriver_OpenDb(ByVal strFilePath As String, ByRef ptrDbOut As LongPtr) As Long
    ISQLiteDriver_OpenDb = ws_open16(StrPtr(strFilePath), ptrDbOut)
End Function

Private Function ISQLiteDriver_OpenDbV2( _
    ByVal strFilePath As String, _
    ByRef ptrDbOut As LongPtr, _
    ByVal lngOpenFlags As Long, _
    Optional ByVal strVfsName As String = vbNullString) As Long

    Dim abytPath() As Byte
    Dim abytVfs() As Byte
    abytPath = VbaStringToUtf8Bytes(strFilePath)
    If Len(strVfsName) = 0 Then
        ISQLiteDriver_OpenDbV2 = ws_open_v2(VarPtr(abytPath(0)), ptrDbOut, lngOpenFlags, 0)
    Else
        abytVfs = VbaStringToUtf8Bytes(strVfsName)
        ISQLiteDriver_OpenDbV2 = ws_open_v2(VarPtr(abytPath(0)), ptrDbOut, lngOpenFlags, VarPtr(abytVfs(0)))
    End If
End Function

Private Function ISQLiteDriver_CloseDb(ByVal ptrDb As LongPtr) As Long
    If ptrDb = 0 Then Exit Function
    ISQLiteDriver_CloseDb = ws_close_v2(ptrDb)
End Function

Private Function ISQLiteDriver_PrepareStmt( _
    ByVal ptrDb As LongPtr, _
    ByVal strSql As String, _
    ByRef ptrStmtOut As LongPtr, _
    Optional ByVal lngPrepFlags As Long = 0, _
    Optional ByRef strRemainingSql As String = vbNullString) As Long

    Dim ptrTail As LongPtr
    ptrStmtOut = 0
    If Len(strSql) = 0 Then
        strRemainingSql = vbNullString
        ISQLiteDriver_PrepareStmt = sqOk
        Exit Function
    End If

    ISQLiteDriver_PrepareStmt = ws_prepare16_v3(ptrDb, StrPtr(strSql), -1, lngPrepFlags, ptrStmtOut, ptrTail)
    If ptrTail <> 0 Then
        strRemainingSql = Utf16PtrToVbaString(ptrTail, -1)
    Else
        strRemainingSql = vbNullString
    End If
End Function

Private Function ISQLiteDriver_StepStmt(ByVal ptrStmt As LongPtr) As Long
    ISQLiteDriver_StepStmt = ws_step(ptrStmt)
End Function

Private Function ISQLiteDriver_ResetStmt(ByVal ptrStmt As LongPtr) As Long
    ISQLiteDriver_ResetStmt = ws_reset(ptrStmt)
End Function

Private Function ISQLiteDriver_ClearStmtBindings(ByVal ptrStmt As LongPtr) As Long
    ISQLiteDriver_ClearStmtBindings = ws_clear_bindings(ptrStmt)
End Function

Private Function ISQLiteDriver_FinalizeStmt(ByVal ptrStmt As LongPtr) As Long
    If ptrStmt = 0 Then Exit Function
    ISQLiteDriver_FinalizeStmt = ws_finalize(ptrStmt)
End Function

Private Function ISQLiteDriver_BindNull(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long) As Long
    ISQLiteDriver_BindNull = ws_bind_null(ptrStmt, lngParamIndex)
End Function

Private Function ISQLiteDriver_BindInt32(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByVal lngVal As Long) As Long
    ISQLiteDriver_BindInt32 = ws_bind_int(ptrStmt, lngParamIndex, lngVal)
End Function

Private Function ISQLiteDriver_BindInt64(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByVal varInt64 As Variant) As Long
    #If Win64 Then
        ISQLiteDriver_BindInt64 = ws_bind_int64(ptrStmt, lngParamIndex, CLngLng(varInt64))
    #Else
        Dim curWhole As Currency
        curWhole = CCur(varInt64)
        ISQLiteDriver_BindInt64 = ws_bind_int64(ptrStmt, lngParamIndex, curWhole / 10000@)
    #End If
End Function

Private Function ISQLiteDriver_BindDouble(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByVal dblVal As Double) As Long
    ISQLiteDriver_BindDouble = ws_bind_double(ptrStmt, lngParamIndex, dblVal)
End Function

Private Function ISQLiteDriver_BindText(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByVal strVal As String) As Long
    If Len(strVal) = 0 Then
        ' Non-null pointer + 0 byte length binds SQL empty string '' rather than NULL.
        ISQLiteDriver_BindText = ws_bind_text16(ptrStmt, lngParamIndex, VarPtr(m_ZeroWideChar), 0, SQLITE_TRANSIENT)
    Else
        ISQLiteDriver_BindText = ws_bind_text16(ptrStmt, lngParamIndex, StrPtr(strVal), Len(strVal) * 2, SQLITE_TRANSIENT)
    End If
End Function

Private Function ISQLiteDriver_BindBlob(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByRef abytVal() As Byte) As Long
    Dim lngLen As Long
    lngLen = GetByteArrayLength(abytVal)
    If lngLen <= 0 Then
        ISQLiteDriver_BindBlob = ws_bind_zeroblob(ptrStmt, lngParamIndex, 0)
    Else
        ISQLiteDriver_BindBlob = ws_bind_blob(ptrStmt, lngParamIndex, VarPtr(abytVal(LBound(abytVal))), lngLen, SQLITE_TRANSIENT)
    End If
End Function

Private Function ISQLiteDriver_GetParamCount(ByVal ptrStmt As LongPtr) As Long
    ISQLiteDriver_GetParamCount = ws_bind_parameter_count(ptrStmt)
End Function

Private Function ISQLiteDriver_GetParamIndex(ByVal ptrStmt As LongPtr, ByVal strParamName As String) As Long
    Dim abytName() As Byte
    abytName = VbaStringToUtf8Bytes(strParamName)
    ISQLiteDriver_GetParamIndex = ws_bind_parameter_index(ptrStmt, VarPtr(abytName(0)))
End Function

Private Function ISQLiteDriver_GetParamName(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long) As String
    ISQLiteDriver_GetParamName = Utf8PtrToVbaString(ws_bind_parameter_name(ptrStmt, lngParamIndex))
End Function

Private Function ISQLiteDriver_GetColumnCount(ByVal ptrStmt As LongPtr) As Long
    ISQLiteDriver_GetColumnCount = ws_column_count(ptrStmt)
End Function

Private Function ISQLiteDriver_GetColumnType(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As Long
    ISQLiteDriver_GetColumnType = ws_column_type(ptrStmt, lngZeroBasedCol)
End Function

Private Function ISQLiteDriver_GetColumnName(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As String
    ISQLiteDriver_GetColumnName = Utf16PtrToVbaString(ws_column_name16(ptrStmt, lngZeroBasedCol), -1)
End Function

Private Function ISQLiteDriver_GetColumnDeclType(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As String
    ISQLiteDriver_GetColumnDeclType = Utf16PtrToVbaString(ws_column_decltype16(ptrStmt, lngZeroBasedCol), -1)
End Function

Private Function ISQLiteDriver_GetColumnInt32(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As Long
    ISQLiteDriver_GetColumnInt32 = ws_column_int(ptrStmt, lngZeroBasedCol)
End Function

Private Function ISQLiteDriver_GetColumnInt64(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As Variant
    #If Win64 Then
        Dim lng64 As LongLong
        lng64 = ws_column_int64(ptrStmt, lngZeroBasedCol)
        If lng64 >= -2147483648^ And lng64 <= 2147483647^ Then
            ISQLiteDriver_GetColumnInt64 = CLng(lng64)
        Else
            ISQLiteDriver_GetColumnInt64 = lng64
        End If
    #Else
        Dim curRaw As Currency
        Dim curWhole As Currency
        curRaw = ws_column_int64(ptrStmt, lngZeroBasedCol)
        curWhole = curRaw * 10000@
        If curWhole >= -2147483648@ And curWhole <= 2147483647@ Then
            ISQLiteDriver_GetColumnInt64 = CLng(curWhole)
        Else
            ISQLiteDriver_GetColumnInt64 = curWhole
        End If
    #End If
End Function

Private Function ISQLiteDriver_GetColumnDouble(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As Double
    ISQLiteDriver_GetColumnDouble = ws_column_double(ptrStmt, lngZeroBasedCol)
End Function

Private Function ISQLiteDriver_GetColumnText(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As String
    Dim ptrWide As LongPtr
    Dim lngBytes As Long
    ptrWide = ws_column_text16(ptrStmt, lngZeroBasedCol)
    lngBytes = ws_column_bytes16(ptrStmt, lngZeroBasedCol)
    ISQLiteDriver_GetColumnText = Utf16PtrToVbaString(ptrWide, lngBytes)
End Function

Private Function ISQLiteDriver_GetColumnBlob(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As Byte()
    Dim ptrData As LongPtr
    Dim lngBytes As Long
    ptrData = ws_column_blob(ptrStmt, lngZeroBasedCol)
    lngBytes = ws_column_bytes(ptrStmt, lngZeroBasedCol)
    ISQLiteDriver_GetColumnBlob = PtrToByteArray(ptrData, lngBytes)
End Function

Private Function ISQLiteDriver_GetChanges(ByVal ptrDb As LongPtr) As Long
    ISQLiteDriver_GetChanges = ws_changes(ptrDb)
End Function

Private Function ISQLiteDriver_GetTotalChanges(ByVal ptrDb As LongPtr) As Long
    ISQLiteDriver_GetTotalChanges = ws_total_changes(ptrDb)
End Function

Private Function ISQLiteDriver_GetLastInsertRowId(ByVal ptrDb As LongPtr) As Variant
    #If Win64 Then
        Dim lng64 As LongLong
        lng64 = ws_last_insert_rowid(ptrDb)
        If lng64 >= -2147483648^ And lng64 <= 2147483647^ Then
            ISQLiteDriver_GetLastInsertRowId = CLng(lng64)
        Else
            ISQLiteDriver_GetLastInsertRowId = lng64
        End If
    #Else
        Dim curRaw As Currency
        Dim curWhole As Currency
        curRaw = ws_last_insert_rowid(ptrDb)
        curWhole = curRaw * 10000@
        If curWhole >= -2147483648@ And curWhole <= 2147483647@ Then
            ISQLiteDriver_GetLastInsertRowId = CLng(curWhole)
        Else
            ISQLiteDriver_GetLastInsertRowId = curWhole
        End If
    #End If
End Function

Private Function ISQLiteDriver_GetAutocommit(ByVal ptrDb As LongPtr) As Boolean
    ISQLiteDriver_GetAutocommit = (ws_get_autocommit(ptrDb) <> 0)
End Function

Private Function ISQLiteDriver_SetBusyTimeout(ByVal ptrDb As LongPtr, ByVal lngTimeoutMs As Long) As Long
    ISQLiteDriver_SetBusyTimeout = ws_busy_timeout(ptrDb, lngTimeoutMs)
End Function

Private Function ISQLiteDriver_GetErrCode(ByVal ptrDb As LongPtr) As Long
    ISQLiteDriver_GetErrCode = ws_errcode(ptrDb)
End Function

Private Function ISQLiteDriver_GetExtendedErrCode(ByVal ptrDb As LongPtr) As Long
    ISQLiteDriver_GetExtendedErrCode = ws_extended_errcode(ptrDb)
End Function

Private Function ISQLiteDriver_GetErrMsg(ByVal ptrDb As LongPtr) As String
    If ptrDb = 0 Then Exit Function
    ISQLiteDriver_GetErrMsg = Utf16PtrToVbaString(ws_errmsg16(ptrDb), -1)
End Function

Private Function ISQLiteDriver_GetExpandedSql(ByVal ptrStmt As LongPtr) As String
    Dim ptrExp As LongPtr
    If ptrStmt = 0 Then Exit Function
    ptrExp = ws_expanded_sql(ptrStmt)
    If ptrExp <> 0 Then
        ISQLiteDriver_GetExpandedSql = Utf8PtrToVbaString(ptrExp)
        ws_free ptrExp
    End If
End Function

Private Function ISQLiteDriver_BackupDatabase( _
    ByVal ptrDestDb As LongPtr, _
    ByVal strDestSchema As String, _
    ByVal ptrSourceDb As LongPtr, _
    ByVal strSourceSchema As String) As Long

    Dim abytDest() As Byte
    Dim abytSrc() As Byte
    Dim ptrBackup As LongPtr
    Dim lngStepRc As Long
    Dim lngFinishRc As Long

    abytDest = VbaStringToUtf8Bytes(IIf(Len(strDestSchema) = 0, "main", strDestSchema))
    abytSrc = VbaStringToUtf8Bytes(IIf(Len(strSourceSchema) = 0, "main", strSourceSchema))

    ptrBackup = ws_backup_init(ptrDestDb, VarPtr(abytDest(0)), ptrSourceDb, VarPtr(abytSrc(0)))
    If ptrBackup = 0 Then
        ISQLiteDriver_BackupDatabase = ws_errcode(ptrDestDb)
        If ISQLiteDriver_BackupDatabase = sqOk Then ISQLiteDriver_BackupDatabase = sqError
        Exit Function
    End If

    lngStepRc = ws_backup_step(ptrBackup, -1)
    lngFinishRc = ws_backup_finish(ptrBackup)
    If lngFinishRc <> sqOk Then
        ISQLiteDriver_BackupDatabase = lngFinishRc
    ElseIf lngStepRc = sqDone Then
        ISQLiteDriver_BackupDatabase = sqOk
    Else
        ISQLiteDriver_BackupDatabase = lngStepRc
    End If
End Function

