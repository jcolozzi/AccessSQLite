Attribute VB_Name = "ISQLiteDriver"
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
' ISQLiteDriver
' Pluggable native SQLite C-API driver interface (32-bit and 64-bit VBA7).
' Implemented by clsSQLiteWinDriver (winsqlite3.dll) and clsSQLiteBundledDriver.
' =============================================================================

Public Property Get DriverTitle() As String
End Property

Public Property Get LibraryVersion() As String
End Property

Public Property Get IsInitialized() As Boolean
End Property

Public Property Get LastInitError() As String
End Property

Public Function InitializeDriver(Optional ByVal strLibDirectory As String = vbNullString) As Boolean
End Function

Public Sub ReleaseDriver()
End Sub

Public Function OpenDb(ByVal strFilePath As String, ByRef ptrDbOut As LongPtr) As Long
End Function

Public Function OpenDbV2( _
    ByVal strFilePath As String, _
    ByRef ptrDbOut As LongPtr, _
    ByVal lngOpenFlags As Long, _
    Optional ByVal strVfsName As String = vbNullString) As Long
End Function

Public Function CloseDb(ByVal ptrDb As LongPtr) As Long
End Function

Public Function PrepareStmt( _
    ByVal ptrDb As LongPtr, _
    ByVal strSql As String, _
    ByRef ptrStmtOut As LongPtr, _
    Optional ByVal lngPrepFlags As Long = 0, _
    Optional ByRef strRemainingSql As String = vbNullString) As Long
End Function

Public Function StepStmt(ByVal ptrStmt As LongPtr) As Long
End Function

Public Function ResetStmt(ByVal ptrStmt As LongPtr) As Long
End Function

Public Function ClearStmtBindings(ByVal ptrStmt As LongPtr) As Long
End Function

Public Function FinalizeStmt(ByVal ptrStmt As LongPtr) As Long
End Function

Public Function BindNull(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long) As Long
End Function

Public Function BindInt32(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByVal lngVal As Long) As Long
End Function

Public Function BindInt64(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByVal varInt64 As Variant) As Long
End Function

Public Function BindDouble(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByVal dblVal As Double) As Long
End Function

Public Function BindText(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByVal strVal As String) As Long
End Function

Public Function BindBlob(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long, ByRef abytVal() As Byte) As Long
End Function

Public Function GetParamCount(ByVal ptrStmt As LongPtr) As Long
End Function

Public Function GetParamIndex(ByVal ptrStmt As LongPtr, ByVal strParamName As String) As Long
End Function

Public Function GetParamName(ByVal ptrStmt As LongPtr, ByVal lngParamIndex As Long) As String
End Function

Public Function GetColumnCount(ByVal ptrStmt As LongPtr) As Long
End Function

Public Function GetColumnType(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As Long
End Function

Public Function GetColumnName(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As String
End Function

Public Function GetColumnDeclType(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As String
End Function

Public Function GetColumnInt32(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As Long
End Function

Public Function GetColumnInt64(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As Variant
End Function

Public Function GetColumnDouble(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As Double
End Function

Public Function GetColumnText(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As String
End Function

Public Function GetColumnBlob(ByVal ptrStmt As LongPtr, ByVal lngZeroBasedCol As Long) As Byte()
End Function

Public Function GetChanges(ByVal ptrDb As LongPtr) As Long
End Function

Public Function GetTotalChanges(ByVal ptrDb As LongPtr) As Long
End Function

Public Function GetLastInsertRowId(ByVal ptrDb As LongPtr) As Variant
End Function

Public Function GetAutocommit(ByVal ptrDb As LongPtr) As Boolean
End Function

Public Function SetBusyTimeout(ByVal ptrDb As LongPtr, ByVal lngTimeoutMs As Long) As Long
End Function

Public Function GetErrCode(ByVal ptrDb As LongPtr) As Long
End Function

Public Function GetExtendedErrCode(ByVal ptrDb As LongPtr) As Long
End Function

Public Function GetErrMsg(ByVal ptrDb As LongPtr) As String
End Function

Public Function GetExpandedSql(ByVal ptrStmt As LongPtr) As String
End Function

Public Function BackupDatabase( _
    ByVal ptrDestDb As LongPtr, _
    ByVal strDestSchema As String, _
    ByVal ptrSourceDb As LongPtr, _
    ByVal strSourceSchema As String) As Long
End Function

