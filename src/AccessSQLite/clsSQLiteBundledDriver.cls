Attribute VB_Name = "clsSQLiteBundledDriver"
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
' clsSQLiteBundledDriver
' Portable bundled SQLite driver:
'   - 32-bit Access: loads <libDir>\lib\sqlite3.dll + <libDir>\lib\SQLite3_StdCall.dll
'   - 64-bit Access: loads <libDir>\lib\x64\sqlite3.dll (with root/direct fallbacks)
' Preserves Err.LastDllError before formatting and frees libraries in reverse order.
' =============================================================================

Implements ISQLiteDriver

Private Const SQLITE_TRANSIENT As LongPtr = -1

Private Declare PtrSafe Function LoadLibraryW Lib "kernel32" (ByVal ptrLibFileName As LongPtr) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal ptrLibModule As LongPtr) As Long
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal ptrModule As LongPtr, ByVal strProcName As String) As LongPtr

#If Win64 Then
Private Declare PtrSafe Function bd_libversion Lib "sqlite3" Alias "sqlite3_libversion" () As LongPtr
Private Declare PtrSafe Function bd_open16 Lib "sqlite3" Alias "sqlite3_open16" (ByVal ptrFileName As LongPtr, ByRef ptrDbOut As LongPtr) As Long
Private Declare PtrSafe Function bd_open_v2 Lib "sqlite3" Alias "sqlite3_open_v2" (ByVal ptrFileNameUtf8 As LongPtr, ByRef ptrDbOut As LongPtr, ByVal lngFlags As Long, ByVal ptrVfsUtf8 As LongPtr) As Long
Private Declare PtrSafe Function bd_close Lib "sqlite3" Alias "sqlite3_close_v2" (ByVal ptrDb As LongPtr) As Long
Private Declare PtrSafe Function bd_prepare16_v2 Lib "sqlite3" Alias "sqlite3_prepare16_v2" (ByVal ptrDb As LongPtr, ByVal ptrSql As LongPtr, ByVal lngByteLen As Long, ByRef ptrStmtOut As LongPtr, ByRef ptrTailOut As LongPtr) As Long
Private Declare PtrSafe Function bd_step Lib "sqlite3" Alias "sqlite3_step" (ByVal ptrStmt As LongPtr) As Long
Private Declare PtrSafe Function bd_reset Lib "sqlite3" Alias "sqlite3_reset" (ByVal ptrStmt As LongPtr) As Long
Private Declare PtrSafe Function bd_clear_bindings Lib "sqlite3" Alias "sqlite3_clear_bindings" (ByVal ptrStmt As LongPtr) As Long
Private Declare PtrSafe Function bd_finalize Lib "sqlite3" Alias "sqlite3_finalize" (ByVal ptrStmt As LongPtr) As Long
Private Declare PtrSafe Function bd_bind_null Lib "sqlite3" Alias "sqlite3_bind_null" (ByVal ptrStmt As LongPtr, ByVal lngIdx As Long) As Long
Private Declare PtrSafe Function bd_bind_int Lib "sqlite3" Alias "sqlite3_bind_int" (ByVal ptrStmt As LongPtr, ByVal lngIdx As Long, ByVal lngVal As Long) As Long
Private Declare PtrSafe Function bd_bind_int64 Lib "sqlite3" Alias "sqlite3_bind_int64" (ByVal ptrStmt As LongPtr, ByVal lngIdx As Long, ByVal lngVal64 As LongLong) As Long
Private Declare PtrSafe Function bd_bind_double Lib "sqlite3" Alias "sqlite3_bind_double" (ByVal ptrStmt As LongPtr, ByVal lngIdx As Long, ByVal dblVal As Double) As Long
Private Declare PtrSafe Function bd_bind_text16 Lib "sqlite3" Alias "sqlite3_bind_text16" (ByVal ptrStmt As LongPtr, ByVal lngIdx As Long, ByVal ptrWide As LongPtr, ByVal lngByteLen As Long, ByVal ptrDel As LongPtr) As Long
Private Declare PtrSafe Function bd_bind_blob Lib "sqlite3" Alias "sqlite3_bind_blob" (ByVal ptrStmt As LongPtr, ByVal lngIdx As Long, ByVal ptrData As LongPtr, ByVal lngByteLen As Long, ByVal ptrDel As LongPtr) As Long
Private Declare PtrSafe Function bd_bind_zeroblob Lib "sqlite3" Alias "sqlite3_bind_zeroblob" (ByVal ptrStmt As LongPtr, ByVal lngIdx As Long, ByVal lngByteLen As Long) As Long
Private Declare PtrSafe Function bd_bind_parameter_count Lib "sqlite3" Alias "sqlite3_bind_parameter_count" (ByVal ptrStmt As LongPtr) As Long
Private Declare PtrSafe Function bd_bind_parameter_index Lib "sqlite3" Alias "sqlite3_bind_parameter_index" (ByVal ptrStmt As LongPtr, ByVal ptrNameUtf8 As LongPtr) As Long
Private Declare PtrSafe Function bd_bind_parameter_name Lib "sqlite3" Alias "sqlite3_bind_parameter_name" (ByVal ptrStmt As LongPtr, ByVal lngIdx As Long) As LongPtr
Private Declare PtrSafe Function bd_column_count Lib "sqlite3" Alias "sqlite3_column_count" (ByVal ptrStmt As LongPtr) As Long
Private Declare PtrSafe Function bd_column_type Lib "sqlite3" Alias "sqlite3_column_type" (ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As Long
Private Declare PtrSafe Function bd_column_name16 Lib "sqlite3" Alias "sqlite3_column_name16" (ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As LongPtr
Private Declare PtrSafe Function bd_column_decltype16 Lib "sqlite3" Alias "sqlite3_column_decltype16" (ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As LongPtr
Private Declare PtrSafe Function bd_column_int Lib "sqlite3" Alias "sqlite3_column_int" (ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As Long
Private Declare PtrSafe Function bd_column_int64 Lib "sqlite3" Alias "sqlite3_column_int64" (ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As LongLong
Private Declare PtrSafe Function bd_column_double Lib "sqlite3" Alias "sqlite3_column_double" (ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As Double
Private Declare PtrSafe Function bd_column_text16 Lib "sqlite3" Alias "sqlite3_column_text16" (ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As LongPtr
Private Declare PtrSafe Function bd_column_bytes16 Lib "sqlite3" Alias "sqlite3_column_bytes16" (ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As Long
Private Declare PtrSafe Function bd_column_blob Lib "sqlite3" Alias "sqlite3_column_blob" (ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As LongPtr
Private Declare PtrSafe Function bd_column_bytes Lib "sqlite3" Alias "sqlite3_column_bytes" (ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As Long
Private Declare PtrSafe Function bd_changes Lib "sqlite3" Alias "sqlite3_changes" (ByVal ptrDb As LongPtr) As Long
Private Declare PtrSafe Function bd_total_changes Lib "sqlite3" Alias "sqlite3_total_changes" (ByVal ptrDb As LongPtr) As Long
Private Declare PtrSafe Function bd_last_insert_rowid Lib "sqlite3" Alias "sqlite3_last_insert_rowid" (ByVal ptrDb As LongPtr) As LongLong
Private Declare PtrSafe Function bd_get_autocommit Lib "sqlite3" Alias "sqlite3_get_autocommit" (ByVal ptrDb As LongPtr) As Long
Private Declare PtrSafe Function bd_busy_timeout Lib "sqlite3" Alias "sqlite3_busy_timeout" (ByVal ptrDb As LongPtr, ByVal lngMs As Long) As Long
Private Declare PtrSafe Function bd_errcode Lib "sqlite3" Alias "sqlite3_errcode" (ByVal ptrDb As LongPtr) As Long
Private Declare PtrSafe Function bd_extended_errcode Lib "sqlite3" Alias "sqlite3_extended_errcode" (ByVal ptrDb As LongPtr) As Long
Private Declare PtrSafe Function bd_errmsg16 Lib "sqlite3" Alias "sqlite3_errmsg16" (ByVal ptrDb As LongPtr) As LongPtr
Private Declare PtrSafe Function bd_backup_init Lib "sqlite3" Alias "sqlite3_backup_init" (ByVal ptrDestDb As LongPtr, ByVal ptrDestName As LongPtr, ByVal ptrSourceDb As LongPtr, ByVal ptrSourceName As LongPtr) As LongPtr
Private Declare PtrSafe Function bd_backup_step Lib "sqlite3" Alias "sqlite3_backup_step" (ByVal ptrBackup As LongPtr, ByVal lngPages As Long) As Long
Private Declare PtrSafe Function bd_backup_finish Lib "sqlite3" Alias "sqlite3_backup_finish" (ByVal ptrBackup As LongPtr) As Long
#Else
Private Declare PtrSafe Function bd_libversion Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_libversion@0" () As LongPtr
Private Declare PtrSafe Function bd_open16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_open16@8" (ByVal ptrFileName As LongPtr, ByRef ptrDbOut As LongPtr) As Long
Private Declare PtrSafe Function bd_open_v2 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_open_v2@16" (ByVal ptrFileNameUtf8 As LongPtr, ByRef ptrDbOut As LongPtr, ByVal lngFlags As Long, ByVal ptrVfsUtf8 As LongPtr) As Long
Private Declare PtrSafe Function bd_close Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_close@4" (ByVal ptrDb As LongPtr) As Long
Private Declare PtrSafe Function bd_prepare16_v2 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_prepare16_v2@20" (ByVal ptrDb As LongPtr, ByVal ptrSql As LongPtr, ByVal lngByteLen As Long, ByRef ptrStmtOut As LongPtr, ByRef ptrTailOut As LongPtr) As Long
Private Declare PtrSafe Function bd_step Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_step@4" (ByVal ptrStmt As LongPtr) As Long
Private Declare PtrSafe Function bd_reset Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_reset@4" (ByVal ptrStmt As LongPtr) As Long
Private Declare PtrSafe Function bd_clear_bindings Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_clear_bindings@4" (ByVal ptrStmt As LongPtr) As Long
Private Declare PtrSafe Function bd_finalize Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_finalize@4" (ByVal ptrStmt As LongPtr) As Long
Private Declare PtrSafe Function bd_bind_null Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_null@8" (ByVal ptrStmt As LongPtr, ByVal lngIdx As Long) As Long
Private Declare PtrSafe Function bd_bind_int Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_int@12" (ByVal ptrStmt As LongPtr, ByVal lngIdx As Long, ByVal lngVal As Long) As Long
Private Declare PtrSafe Function bd_bind_int64 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_int64@16" (ByVal ptrStmt As LongPtr, ByVal lngIdx As Long, ByVal curRawBits As Currency) As Long
Private Declare PtrSafe Function bd_bind_double Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_double@16" (ByVal ptrStmt As LongPtr, ByVal lngIdx As Long, ByVal dblVal As Double) As Long
Private Declare PtrSafe Function bd_bind_text16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_text16@20" (ByVal ptrStmt As LongPtr, ByVal lngIdx As Long, ByVal ptrWide As LongPtr, ByVal lngByteLen As Long, ByVal ptrDel As LongPtr) As Long
Private Declare PtrSafe Function bd_bind_blob Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_blob@20" (ByVal ptrStmt As LongPtr, ByVal lngIdx As Long, ByVal ptrData As LongPtr, ByVal lngByteLen As Long, ByVal ptrDel As LongPtr) As Long
Private Declare PtrSafe Function bd_bind_zeroblob Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_zeroblob@12" (ByVal ptrStmt As LongPtr, ByVal lngIdx As Long, ByVal lngByteLen As Long) As Long
Private Declare PtrSafe Function bd_bind_parameter_count Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_parameter_count@4" (ByVal ptrStmt As LongPtr) As Long
Private Declare PtrSafe Function bd_bind_parameter_index Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_parameter_index@8" (ByVal ptrStmt As LongPtr, ByVal ptrNameUtf8 As LongPtr) As Long
Private Declare PtrSafe Function bd_bind_parameter_name Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_parameter_name@8" (ByVal ptrStmt As LongPtr, ByVal lngIdx As Long) As LongPtr
Private Declare PtrSafe Function bd_column_count Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_count@4" (ByVal ptrStmt As LongPtr) As Long
Private Declare PtrSafe Function bd_column_type Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_type@8" (ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As Long
Private Declare PtrSafe Function bd_column_name16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_name16@8" (ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As LongPtr
Private Declare PtrSafe Function bd_column_int Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_int@8" (ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As Long
Private Declare PtrSafe Function bd_column_int64 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_int64@8" (ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As Currency
Private Declare PtrSafe Function bd_column_double Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_double@8" (ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As Double
Private Declare PtrSafe Function bd_column_text16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_text16@8" (ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As LongPtr
Private Declare PtrSafe Function bd_column_bytes16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_bytes16@8" (ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As Long
Private Declare PtrSafe Function bd_column_blob Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_blob@8" (ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As LongPtr
Private Declare PtrSafe Function bd_column_bytes Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_bytes@8" (ByVal ptrStmt As LongPtr, ByVal lngCol As Long) As Long
Private Declare PtrSafe Function bd_changes Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_changes@4" (ByVal ptrDb As LongPtr) As Long
Private Declare PtrSafe Function bd_total_changes Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_total_changes@4" (ByVal ptrDb As LongPtr) As Long
Private Declare PtrSafe Function bd_errcode Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_errcode@4" (ByVal ptrDb As LongPtr) As Long
Private Declare PtrSafe Function bd_extended_errcode Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_extended_errcode@4" (ByVal ptrDb As LongPtr) As Long
Private Declare PtrSafe Function bd_errmsg16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_errmsg16@4" (ByVal ptrDb As LongPtr) As LongPtr
Private Declare PtrSafe Function bd_backup_init Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_init@16" (ByVal ptrDestDb As LongPtr, ByVal ptrDestName As LongPtr, ByVal ptrSourceDb As LongPtr, ByVal ptrSourceName As LongPtr) As LongPtr
Private Declare PtrSafe Function bd_backup_step Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_step@8" (ByVal ptrBackup As LongPtr, ByVal lngPages As Long) As Long
Private Declare PtrSafe Function bd_backup_finish Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_finish@4" (ByVal ptrBackup As LongPtr) As Long
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" (ByVal ptrInstance As LongPtr, ByVal ptrFuncAddr As LongPtr, ByVal lngCallConv As Long, ByVal intVtReturn As Integer, ByVal lngArgCount As Long, ByRef intArgTypes As Integer, ByRef ptrArgValues As LongPtr, ByRef varReturnVal As Variant) As Long
#End If

Private m_SqliteHandle As LongPtr
Private m_StdCallHandle As LongPtr
Private m_HasOpenV2Export As Boolean
Private m_FnColDeclType16 As LongPtr
Private m_FnGetAutocommit As LongPtr
Private m_FnBusyTimeout As LongPtr
Private m_FnExpandedSql As LongPtr
Private m_FnSqliteFree As LongPtr
Private m_LastInitErr As String
Private m_ZeroWideChar As Integer

Private Sub Class_Terminate()
    ISQLiteDriver_ReleaseDriver
End Sub

Private Property Get ISQLiteDriver_DriverTitle() As String
    #If Win64 Then
        ISQLiteDriver_DriverTitle = "Bundled SQLite3 (x64 Direct)"
    #Else
        ISQLiteDriver_DriverTitle = "Bundled SQLite3 + SQLite3_StdCall (x86)"
    #End If
End Property

Private Property Get ISQLiteDriver_LibraryVersion() As String
    If Not ISQLiteDriver_IsInitialized Then
        If Not ISQLiteDriver_InitializeDriver() Then Exit Property
    End If
    ISQLiteDriver_LibraryVersion = Utf8PtrToVbaString(bd_libversion())
End Property

Private Property Get ISQLiteDriver_IsInitialized() As Boolean
    #If Win64 Then
        ISQLiteDriver_IsInitialized = (m_SqliteHandle <> 0)
    #Else
        ISQLiteDriver_IsInitialized = (m_SqliteHandle <> 0 And m_StdCallHandle <> 0)
    #End If
End Property

Private Property Get ISQLiteDriver_LastInitError() As String
    ISQLiteDriver_LastInitError = m_LastInitErr
End Property

Private Function ISQLiteDriver_InitializeDriver(Optional ByVal strLibDirectory As String = vbNullString) As Boolean
    Dim strBaseDir As String
    Dim strEnginePath As String
    Dim strThunkPath As String
    Dim lngWinErr As Long

    If ISQLiteDriver_IsInitialized Then
        ISQLiteDriver_InitializeDriver = True
        Exit Function
    End If

    strBaseDir = Trim$(strLibDirectory)
    If Len(strBaseDir) = 0 Then
        strBaseDir = GetBundledDllDirectory()
    Else
        If Right$(strBaseDir, 1) = "\" Then strBaseDir = Left$(strBaseDir, Len(strBaseDir) - 1)
        ' If caller passed the parent project directory, step into \lib when DLLs reside there
        If Len(Dir$(strBaseDir & "\sqlite3.dll")) = 0 And Len(Dir$(strBaseDir & "\x64\sqlite3.dll")) = 0 Then
            If Len(Dir$(strBaseDir & "\lib\sqlite3.dll")) > 0 Or Len(Dir$(strBaseDir & "\lib\x64\sqlite3.dll")) > 0 Then
                strBaseDir = strBaseDir & "\lib"
            End If
        End If
    End If
    If Right$(strBaseDir, 1) <> "\" Then strBaseDir = strBaseDir & "\"

    #If Win64 Then
        strEnginePath = strBaseDir & "x64\sqlite3.dll"
        m_SqliteHandle = LoadLibraryW(StrPtr(strEnginePath))
        lngWinErr = Err.LastDllError
        If m_SqliteHandle = 0 Then
            strEnginePath = strBaseDir & "lib\x64\sqlite3.dll"
            m_SqliteHandle = LoadLibraryW(StrPtr(strEnginePath))
            lngWinErr = Err.LastDllError
        End If
        If m_SqliteHandle = 0 Then
            strEnginePath = strBaseDir & "sqlite3.dll"
            m_SqliteHandle = LoadLibraryW(StrPtr(strEnginePath))
            lngWinErr = Err.LastDllError
        End If
        If m_SqliteHandle = 0 Then
            m_LastInitErr = "Failed to load 64-bit sqlite3.dll from '" & strBaseDir & "x64\sqlite3.dll': " & FormatWin32Error(lngWinErr)
            Exit Function
        End If
        m_HasOpenV2Export = True
    #Else
        If m_SqliteHandle = 0 Then
            strEnginePath = strBaseDir & "sqlite3.dll"
            m_SqliteHandle = LoadLibraryW(StrPtr(strEnginePath))
            lngWinErr = Err.LastDllError
            If m_SqliteHandle = 0 Then
                m_LastInitErr = "Failed to load 32-bit sqlite3.dll from '" & strEnginePath & "': " & FormatWin32Error(lngWinErr)
                Exit Function
            End If
        End If
        If m_StdCallHandle = 0 Then
            strThunkPath = strBaseDir & "SQLite3_StdCall.dll"
            m_StdCallHandle = LoadLibraryW(StrPtr(strThunkPath))
            lngWinErr = Err.LastDllError
            If m_StdCallHandle = 0 Then
                FreeLibrary m_SqliteHandle
                m_SqliteHandle = 0
                m_LastInitErr = "Failed to load SQLite3_StdCall.dll from '" & strThunkPath & "': " & FormatWin32Error(lngWinErr)
                Exit Function
            End If
            m_HasOpenV2Export = (GetProcAddress(m_StdCallHandle, "_sqlite3_stdcall_open_v2@16") <> 0)
            m_FnColDeclType16 = GetProcAddress(m_SqliteHandle, "sqlite3_column_decltype16")
            m_FnGetAutocommit = GetProcAddress(m_SqliteHandle, "sqlite3_get_autocommit")
            m_FnBusyTimeout = GetProcAddress(m_SqliteHandle, "sqlite3_busy_timeout")
            m_FnExpandedSql = GetProcAddress(m_SqliteHandle, "sqlite3_expanded_sql")
            m_FnSqliteFree = GetProcAddress(m_SqliteHandle, "sqlite3_free")
        End If
    #End If

    m_LastInitErr = vbNullString
    ISQLiteDriver_InitializeDriver = True
End Function

Private Sub ISQLiteDriver_ReleaseDriver()
    If m_StdCallHandle <> 0 Then
        FreeLibrary m_StdCallHandle
        m_StdCallHandle = 0
    End If
    If m_SqliteHandle <> 0 Then
        FreeLibrary m_SqliteHandle
        m_SqliteHandle = 0
    End If
    m_HasOpenV2Export = False
    m_FnColDeclType16 = 0
    m_FnGetAutocommit = 0
    m_FnBusyTimeout = 0
    m_FnExpandedSql = 0
    m_FnSqliteFree = 0
End Sub

Private Function ISQLiteDriver_OpenDb(ByVal strFilePath As String, ByRef ptrDbOut As LongPtr) As Long
    ISQLiteDriver_OpenDb = bd_open16(StrPtr(strFilePath), ptrDbOut)
End Function

Private Function ISQLiteDriver_OpenDbV2(ByVal strFilePath As String, ByRef ptrDbOut As LongPtr, ByVal lngOpenFlags As Long, Optional ByVal strVfsName As String = vbNullString) As Long
    Dim abytPath() As Byte
    Dim abytVfs() As Byte
    Dim lngRc As Long

    If m_HasOpenV2Export Then
        abytPath = VbaStringToUtf8Bytes(strFilePath)
        If Len(strVfsName) = 0 Then
            ISQLiteDriver_OpenDbV2 = bd_open_v2(VarPtr(abytPath(0)), ptrDbOut, lngOpenFlags, 0)
        Else
            abytVfs = VbaStringToUtf8Bytes(strVfsName)
            ISQLiteDriver_OpenDbV2 = bd_open_v2(VarPtr(abytPath(0)), ptrDbOut, lngOpenFlags, VarPtr(abytVfs(0)))
        End If
        Exit Function
    End If

    If (lngOpenFlags And sqOpenMemory) <> 0 Then
        ISQLiteDriver_OpenDbV2 = bd_open16(StrPtr(":memory:"), ptrDbOut)
        Exit Function
    End If
    If (lngOpenFlags And sqOpenReadOnly) <> 0 And StrComp(strFilePath, ":memory:", vbTextCompare) <> 0 Then
        If Len(Dir$(strFilePath)) = 0 Then
            ISQLiteDriver_OpenDbV2 = sqCantOpen
            Exit Function
        End If
        lngRc = bd_open16(StrPtr(strFilePath), ptrDbOut)
        If lngRc = sqOk And ptrDbOut <> 0 Then
            ExecInternalSql ptrDbOut, "PRAGMA query_only = ON;"
        End If
        ISQLiteDriver_OpenDbV2 = lngRc
        Exit Function
    End If
    ISQLiteDriver_OpenDbV2 = bd_open16(StrPtr(strFilePath), ptrDbOut)
End Function

Private Function ISQLiteDriver_CloseDb(ByVal ptrDb As LongPtr) As Long
    If ptrDb = 0 Then Exit Function
    ISQLiteDriver_CloseDb = bd_close(ptrDb)
End Function

Private Function ISQLiteDriver_PrepareStmt(ByVal ptrDb As LongPtr, ByVal strSql As String, ByRef ptrStmtOut As LongPtr, Optional ByVal lngPrepFlags As Long = 0, Optional ByRef strRemainingSql As String = vbNullString) As Long
    Dim ptrTail As LongPtr
    ptrStmtOut = 0
    If Len(strSql) = 0 Then
        strRemainingSql = vbNullString
        ISQLiteDriver_PrepareStmt = sqOk
        Exit Function
    End If

    ISQLiteDriver_PrepareStmt = bd_prepare16_v2(ptrDb, StrPtr(strSql), -1, ptrStmtOut, ptrTail)
    If ptrTail <> 0 Then
        strRemainingSql = Utf16PtrToVbaString(ptrTail, -1)
    Else
        strRemainingSql = vbNullString
    End If
End Function

Private Function ISQLiteDriver_StepStmt(ByVal ptrStmt As LongPtr) As Long
    ISQLiteDriver_StepStmt = bd_step(ptrStmt)
End Function

Private Function ISQLiteDriver_ResetStmt(ByVal ptrStmt As LongPtr) As Long
    ISQLiteDriver_ResetStmt = bd_reset(ptrStmt)
End Function

Private Function ISQLiteDriver_ClearStmtBindings(ByVal ptrStmt As LongPtr) As Long
    ISQLiteDriver_ClearStmtBindings = bd_clear_bindings(ptrStmt)
End Function

Private Function ISQLiteDriver_FinalizeStmt(ByVal ptrStmt As LongPtr) As Long
    If ptrStmt = 0 Then Exit Function
    ISQLiteDriver_FinalizeStmt = bd_finalize(ptrStmt)
End Function

Private Function ISQLiteDriver_BindNull(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long) As Long
    ISQLiteDriver_BindNull = bd_bind_null(ptrStmt, lngParamIndex)
End Function

Private Function ISQLiteDriver_BindInt32(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByVal lngVal As Long) As Long
    ISQLiteDriver_BindInt32 = bd_bind_int(ptrStmt, lngParamIndex, lngVal)
End Function

Private Function ISQLiteDriver_BindInt64(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByVal varInt64 As Variant) As Long
    #If Win64 Then
        ISQLiteDriver_BindInt64 = bd_bind_int64(ptrStmt, lngParamIndex, CLngLng(varInt64))
    #Else
        Dim curWhole As Currency
        curWhole = CCur(varInt64)
        ISQLiteDriver_BindInt64 = bd_bind_int64(ptrStmt, lngParamIndex, curWhole / 10000@)
    #End If
End Function

Private Function ISQLiteDriver_BindDouble(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByVal dblVal As Double) As Long
    ISQLiteDriver_BindDouble = bd_bind_double(ptrStmt, lngParamIndex, dblVal)
End Function

Private Function ISQLiteDriver_BindText(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByVal strVal As String) As Long
    If Len(strVal) = 0 Then
        ISQLiteDriver_BindText = bd_bind_text16(ptrStmt, lngParamIndex, VarPtr(m_ZeroWideChar), 0, SQLITE_TRANSIENT)
    Else
        ISQLiteDriver_BindText = bd_bind_text16(ptrStmt, lngParamIndex, StrPtr(strVal), Len(strVal) * 2, SQLITE_TRANSIENT)
    End If
End Function

Private Function ISQLiteDriver_BindBlob(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByRef abytVal() As Byte) As Long
    Dim lngLen As Long
    lngLen = GetByteArrayLength(abytVal)
    If lngLen <= 0 Then
        ISQLiteDriver_BindBlob = bd_bind_zeroblob(ptrStmt, lngParamIndex, 0)
    Else
        ISQLiteDriver_BindBlob = bd_bind_blob(ptrStmt, lngParamIndex, VarPtr(abytVal(LBound(abytVal))), lngLen, SQLITE_TRANSIENT)
    End If
End Function

Private Function ISQLiteDriver_GetParamCount(ByVal ptrStmt As LongPtr) As Long
    ISQLiteDriver_GetParamCount = bd_bind_parameter_count(ptrStmt)
End Function

Private Function ISQLiteDriver_GetParamIndex(ByVal ptrStmt As LongPtr, ByVal strParamName As String) As Long
    Dim abytName() As Byte
    abytName = VbaStringToUtf8Bytes(strParamName)
    ISQLiteDriver_GetParamIndex = bd_bind_parameter_index(ptrStmt, VarPtr(abytName(0)))
End Function

Private Function ISQLiteDriver_GetParamName(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long) As String
    ISQLiteDriver_GetParamName = Utf8PtrToVbaString(bd_bind_parameter_name(ptrStmt, lngParamIndex))
End Function

Private Function ISQLiteDriver_GetColumnCount(ByVal ptrStmt As LongPtr) As Long
    ISQLiteDriver_GetColumnCount = bd_column_count(ptrStmt)
End Function

Private Function ISQLiteDriver_GetColumnType(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As Long
    ISQLiteDriver_GetColumnType = bd_column_type(ptrStmt, lngZeroBasedCol)
End Function

Private Function ISQLiteDriver_GetColumnName(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As String
    ISQLiteDriver_GetColumnName = Utf16PtrToVbaString(bd_column_name16(ptrStmt, lngZeroBasedCol), -1)
End Function

Private Function ISQLiteDriver_GetColumnDeclType(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As String
    #If Win64 Then
        ISQLiteDriver_GetColumnDeclType = Utf16PtrToVbaString(bd_column_decltype16(ptrStmt, lngZeroBasedCol), -1)
    #Else
        If m_FnColDeclType16 <> 0 Then
            ISQLiteDriver_GetColumnDeclType = Utf16PtrToVbaString(CallCdeclLong2(m_FnColDeclType16, CLng(ptrStmt), lngZeroBasedCol), -1)
        End If
    #End If
End Function

Private Function ISQLiteDriver_GetColumnInt32(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As Long
    ISQLiteDriver_GetColumnInt32 = bd_column_int(ptrStmt, lngZeroBasedCol)
End Function

Private Function ISQLiteDriver_GetColumnInt64(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As Variant
    #If Win64 Then
        Dim lng64 As LongLong
        lng64 = bd_column_int64(ptrStmt, lngZeroBasedCol)
        If lng64 >= -2147483648^ And lng64 <= 2147483647^ Then
            ISQLiteDriver_GetColumnInt64 = CLng(lng64)
        Else
            ISQLiteDriver_GetColumnInt64 = lng64
        End If
    #Else
        Dim curRaw As Currency
        Dim curWhole As Currency
        curRaw = bd_column_int64(ptrStmt, lngZeroBasedCol)
        curWhole = curRaw * 10000@
        If curWhole >= -2147483648@ And curWhole <= 2147483647@ Then
            ISQLiteDriver_GetColumnInt64 = CLng(curWhole)
        Else
            ISQLiteDriver_GetColumnInt64 = curWhole
        End If
    #End If
End Function

Private Function ISQLiteDriver_GetColumnDouble(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As Double
    ISQLiteDriver_GetColumnDouble = bd_column_double(ptrStmt, lngZeroBasedCol)
End Function

Private Function ISQLiteDriver_GetColumnText(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As String
    Dim ptrWide As LongPtr
    Dim lngBytes As Long
    ptrWide = bd_column_text16(ptrStmt, lngZeroBasedCol)
    lngBytes = bd_column_bytes16(ptrStmt, lngZeroBasedCol)
    ISQLiteDriver_GetColumnText = Utf16PtrToVbaString(ptrWide, lngBytes)
End Function

Private Function ISQLiteDriver_GetColumnBlob(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As Byte()
    Dim ptrData As LongPtr
    Dim lngBytes As Long
    ptrData = bd_column_blob(ptrStmt, lngZeroBasedCol)
    lngBytes = bd_column_bytes(ptrStmt, lngZeroBasedCol)
    ISQLiteDriver_GetColumnBlob = PtrToByteArray(ptrData, lngBytes)
End Function

Private Function ISQLiteDriver_GetChanges(ByVal ptrDb As LongPtr) As Long
    ISQLiteDriver_GetChanges = bd_changes(ptrDb)
End Function

Private Function ISQLiteDriver_GetTotalChanges(ByVal ptrDb As LongPtr) As Long
    ISQLiteDriver_GetTotalChanges = bd_total_changes(ptrDb)
End Function

Private Function ISQLiteDriver_GetLastInsertRowId(ByVal ptrDb As LongPtr) As Variant
    #If Win64 Then
        Dim lng64 As LongLong
        lng64 = bd_last_insert_rowid(ptrDb)
        If lng64 >= -2147483648^ And lng64 <= 2147483647^ Then
            ISQLiteDriver_GetLastInsertRowId = CLng(lng64)
        Else
            ISQLiteDriver_GetLastInsertRowId = lng64
        End If
    #Else
        Dim ptrStmt As LongPtr
        Dim ptrTail As LongPtr
        Dim lngRc As Long
        lngRc = bd_prepare16_v2(ptrDb, StrPtr("SELECT last_insert_rowid();"), -1, ptrStmt, ptrTail)
        If lngRc = sqOk And ptrStmt <> 0 Then
            If bd_step(ptrStmt) = sqRow Then
                ISQLiteDriver_GetLastInsertRowId = ISQLiteDriver_GetColumnInt64(ptrStmt, 0)
            End If
            bd_finalize ptrStmt
        End If
    #End If
End Function

Private Function ISQLiteDriver_GetAutocommit(ByVal ptrDb As LongPtr) As Boolean
    #If Win64 Then
        ISQLiteDriver_GetAutocommit = (bd_get_autocommit(ptrDb) <> 0)
    #Else
        If m_FnGetAutocommit <> 0 Then
            ISQLiteDriver_GetAutocommit = (CallCdeclLong1(m_FnGetAutocommit, CLng(ptrDb)) <> 0)
        Else
            ISQLiteDriver_GetAutocommit = True
        End If
    #End If
End Function

Private Function ISQLiteDriver_SetBusyTimeout(ByVal ptrDb As LongPtr, ByVal lngTimeoutMs As Long) As Long
    #If Win64 Then
        ISQLiteDriver_SetBusyTimeout = bd_busy_timeout(ptrDb, lngTimeoutMs)
    #Else
        If m_FnBusyTimeout <> 0 Then
            ISQLiteDriver_SetBusyTimeout = CallCdeclLong2(m_FnBusyTimeout, CLng(ptrDb), lngTimeoutMs)
        Else
            ISQLiteDriver_SetBusyTimeout = ExecInternalSql(ptrDb, "PRAGMA busy_timeout = " & lngTimeoutMs & ";")
        End If
    #End If
End Function

Private Function ISQLiteDriver_GetErrCode(ByVal ptrDb As LongPtr) As Long
    ISQLiteDriver_GetErrCode = bd_errcode(ptrDb)
End Function

Private Function ISQLiteDriver_GetExtendedErrCode(ByVal ptrDb As LongPtr) As Long
    ISQLiteDriver_GetExtendedErrCode = bd_extended_errcode(ptrDb)
End Function

Private Function ISQLiteDriver_GetErrMsg(ByVal ptrDb As LongPtr) As String
    If ptrDb = 0 Then Exit Function
    ISQLiteDriver_GetErrMsg = Utf16PtrToVbaString(bd_errmsg16(ptrDb), -1)
End Function

Private Function ISQLiteDriver_GetExpandedSql(ByVal ptrStmt As LongPtr) As String
    #If Win64 Then
        ISQLiteDriver_GetExpandedSql = vbNullString
    #Else
        Dim ptrExp As LongPtr
        If ptrStmt = 0 Or m_FnExpandedSql = 0 Then Exit Function
        ptrExp = CallCdeclLong1(m_FnExpandedSql, CLng(ptrStmt))
        If ptrExp <> 0 Then
            ISQLiteDriver_GetExpandedSql = Utf8PtrToVbaString(ptrExp)
            If m_FnSqliteFree <> 0 Then CallCdeclVoid1 m_FnSqliteFree, CLng(ptrExp)
        End If
    #End If
End Function

Private Function ISQLiteDriver_BackupDatabase(ByVal ptrDestDb As LongPtr, ByVal strDestSchema As String, ByVal ptrSourceDb As LongPtr, ByVal strSourceSchema As String) As Long
    Dim abytDest() As Byte
    Dim abytSrc() As Byte
    Dim ptrBackup As LongPtr
    Dim lngStepRc As Long
    Dim lngFinishRc As Long

    abytDest = VbaStringToUtf8Bytes(IIf(Len(strDestSchema) = 0, "main", strDestSchema))
    abytSrc = VbaStringToUtf8Bytes(IIf(Len(strSourceSchema) = 0, "main", strSourceSchema))

    ptrBackup = bd_backup_init(ptrDestDb, VarPtr(abytDest(0)), ptrSourceDb, VarPtr(abytSrc(0)))
    If ptrBackup = 0 Then
        ISQLiteDriver_BackupDatabase = bd_errcode(ptrDestDb)
        If ISQLiteDriver_BackupDatabase = sqOk Then ISQLiteDriver_BackupDatabase = sqError
        Exit Function
    End If

    lngStepRc = bd_backup_step(ptrBackup, -1)
    lngFinishRc = bd_backup_finish(ptrBackup)
    If lngFinishRc <> sqOk Then
        ISQLiteDriver_BackupDatabase = lngFinishRc
    ElseIf lngStepRc = sqDone Then
        ISQLiteDriver_BackupDatabase = sqOk
    Else
        ISQLiteDriver_BackupDatabase = lngStepRc
    End If
End Function

Private Function ExecInternalSql(ByVal ptrDb As LongPtr, ByVal strSql As String) As Long
    Dim ptrStmt As LongPtr
    Dim ptrTail As LongPtr
    Dim lngRc As Long
    lngRc = bd_prepare16_v2(ptrDb, StrPtr(strSql), -1, ptrStmt, ptrTail)
    If lngRc <> sqOk Then
        ExecInternalSql = lngRc
        Exit Function
    End If
    If ptrStmt <> 0 Then
        lngRc = bd_step(ptrStmt)
        bd_finalize ptrStmt
        If lngRc = sqDone Or lngRc = sqRow Then lngRc = sqOk
    End If
    ExecInternalSql = lngRc
End Function

#If Win64 Then
#Else
Private Function CallCdeclLong1(ByVal ptrFunc As LongPtr, ByVal lngArg0 As Long) As Long
    Dim vntArg0 As Variant
    Dim aintTypes(0 To 0) As Integer
    Dim aptrArgs(0 To 0) As LongPtr
    Dim vntRet As Variant
    If ptrFunc = 0 Then Exit Function
    vntArg0 = lngArg0
    aintTypes(0) = vbLong
    aptrArgs(0) = VarPtr(vntArg0)
    If DispCallFunc(0, ptrFunc, 1, vbLong, 1, aintTypes(0), aptrArgs(0), vntRet) = 0 Then
        CallCdeclLong1 = CLng(vntRet)
    End If
End Function

Private Function CallCdeclLong2(ByVal ptrFunc As LongPtr, ByVal lngArg0 As Long, ByVal lngArg1 As Long) As Long
    Dim vntArg0 As Variant
    Dim vntArg1 As Variant
    Dim aintTypes(0 To 1) As Integer
    Dim aptrArgs(0 To 1) As LongPtr
    Dim vntRet As Variant
    If ptrFunc = 0 Then Exit Function
    vntArg0 = lngArg0
    vntArg1 = lngArg1
    aintTypes(0) = vbLong
    aintTypes(1) = vbLong
    aptrArgs(0) = VarPtr(vntArg0)
    aptrArgs(1) = VarPtr(vntArg1)
    If DispCallFunc(0, ptrFunc, 1, vbLong, 2, aintTypes(0), aptrArgs(0), vntRet) = 0 Then
        CallCdeclLong2 = CLng(vntRet)
    End If
End Function

Private Sub CallCdeclVoid1(ByVal ptrFunc As LongPtr, ByVal lngArg0 As Long)
    Dim vntArg0 As Variant
    Dim aintTypes(0 To 0) As Integer
    Dim aptrArgs(0 To 0) As LongPtr
    Dim vntRet As Variant
    If ptrFunc = 0 Then Exit Sub
    vntArg0 = lngArg0
    aintTypes(0) = vbLong
    aptrArgs(0) = VarPtr(vntArg0)
    DispCallFunc 0, ptrFunc, 1, vbEmpty, 1, aintTypes(0), aptrArgs(0), vntRet
End Sub
#End If
