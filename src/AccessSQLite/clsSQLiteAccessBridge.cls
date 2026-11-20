Attribute VB_Name = "clsSQLiteAccessBridge"
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
' clsSQLiteAccessBridge
' High-speed transactional data & schema bridge between Access DAO.Database
' and SQLite (ExportAccessToSQLite and ImportSQLiteToAccess).
' Fixes Null-binding bleed, 0-row schema creation, and >255-char Memo fields.
' =============================================================================

Private m_Conn As clsSQLiteConnection
Private m_DaoDb As DAO.Database

Public Sub InitBridge(ByVal objConn As clsSQLiteConnection, Optional ByVal objDaoDb As DAO.Database = Nothing)
    Set m_Conn = objConn
    If objDaoDb Is Nothing Then
        Set m_DaoDb = CurrentDb
    Else
        Set m_DaoDb = objDaoDb
    End If
End Sub

' Exports an Access table or SELECT query into a SQLite table inside a single transaction.
Public Function ExportAccessToSQLite( _
    ByVal strAccessTableOrSql As String, _
    ByVal strTargetSqliteTable As String, _
    Optional ByVal blnOverwriteTable As Boolean = False) As Long

    Dim strSource As String
    Dim strOpenSql As String
    Dim strPkField As String
    Dim rstSrc As DAO.Recordset
    Dim lngFieldCount As Long
    Dim lngIdx As Long
    Dim strColsDdl As String
    Dim strInsertCols As String
    Dim strPlaceholders As String
    Dim objFld As DAO.Field
    Dim objStmt As clsSQLiteStatement
    Dim lngExported As Long

    EnsureReady
    strSource = Trim$(strAccessTableOrSql)
    If Len(strSource) = 0 Or Len(Trim$(strTargetSqliteTable)) = 0 Then
        RaiseSQLiteError sqErrInvalidArgument, "clsSQLiteAccessBridge.ExportAccessToSQLite", _
            "Source Access table/SQL and target SQLite table name are required."
    End If

    If IsSelectSql(strSource) Then
        strOpenSql = strSource
    Else
        strPkField = FindAccessSinglePkField(strSource)
        strOpenSql = "SELECT * FROM " & QuoteAccessIdent(strSource) & ";"
    End If

    Set rstSrc = m_DaoDb.OpenRecordset(strOpenSql, dbOpenSnapshot)
    On Error GoTo ExportCleanFail

    lngFieldCount = rstSrc.Fields.Count
    For lngIdx = 0 To lngFieldCount - 1
        Set objFld = rstSrc.Fields(lngIdx)
        If lngIdx > 0 Then
            strColsDdl = strColsDdl & ", "
            strInsertCols = strInsertCols & ", "
            strPlaceholders = strPlaceholders & ", "
        End If
        strColsDdl = strColsDdl & QuoteSQLiteIdent(objFld.Name) & " " & MapDaoFieldToSqliteDdl(objFld, strPkField)
        strInsertCols = strInsertCols & QuoteSQLiteIdent(objFld.Name)
        strPlaceholders = strPlaceholders & "?"
    Next lngIdx

    m_Conn.BeginTrans
    If blnOverwriteTable Then
        m_Conn.SchemaManager().DropTableIfExists strTargetSqliteTable
    End If
    m_Conn.ExecuteParams _
        "CREATE TABLE IF NOT EXISTS " & QuoteSQLiteIdent(strTargetSqliteTable) & " (" & strColsDdl & ");", Empty

    Set objStmt = m_Conn.PrepareStatement( _
        "INSERT INTO " & QuoteSQLiteIdent(strTargetSqliteTable) & _
        " (" & strInsertCols & ") VALUES (" & strPlaceholders & ");", False)

    Do Until rstSrc.EOF
        objStmt.ResetForReuse
        For lngIdx = 0 To lngFieldCount - 1
            BindDaoFieldToSqlite objStmt, lngIdx + 1, rstSrc.Fields(lngIdx)
        Next lngIdx
        objStmt.StepOnce
        lngExported = lngExported + 1
        rstSrc.MoveNext
    Loop

    objStmt.FinalizeStatement
    Set objStmt = Nothing
    m_Conn.CommitTrans

    rstSrc.Close
    Set rstSrc = Nothing
    ExportAccessToSQLite = lngExported
    Exit Function

ExportCleanFail:
    Dim lngErrNum As Long
    Dim strErrSrc As String
    Dim strErrDesc As String
    lngErrNum = Err.Number
    strErrSrc = Err.Source
    strErrDesc = Err.Description
    On Error Resume Next
    If Not objStmt Is Nothing Then objStmt.FinalizeStatement
    If m_Conn.InTransaction Then m_Conn.RollbackTrans
    If Not rstSrc Is Nothing Then rstSrc.Close
    On Error GoTo 0
    RaiseSQLiteError sqErrBridge, "clsSQLiteAccessBridge.ExportAccessToSQLite", _
        "Export failed: " & strErrDesc & " (" & strErrSrc & " #" & lngErrNum & ")"
End Function

' Imports a SQLite table or parameterized SELECT query into a local Access table.
Public Function ImportQueryToAccess( _
    ByVal strSqliteSql As String, _
    ByVal strTargetAccessTable As String, _
    ByVal blnOverwriteTable As Boolean, _
    ParamArray varParams() As Variant) As Long

    Dim varPassed As Variant
    varPassed = varParams
    ImportQueryToAccess = ImportSQLiteToAccess(strSqliteSql, strTargetAccessTable, blnOverwriteTable, varPassed)
End Function

Public Function ImportSQLiteToAccess( _
    ByVal strSqliteTableOrSql As String, _
    ByVal strTargetAccessTable As String, _
    Optional ByVal blnOverwriteTable As Boolean = False, _
    Optional ByVal varParamList As Variant) As Long

    Dim strSql As String
    Dim varPassed As Variant
    Dim objStmt As clsSQLiteStatement
    Dim lngColCount As Long
    Dim lngCol As Long
    Dim astrColNames() As String
    Dim astrDeclTypes() As String
    Dim alngDaoTypes() As Long
    Dim colBufferedRows As Collection
    Dim avntRow() As Variant
    Dim lngRc As Long
    Dim varVal As Variant
    Dim tdfNew As DAO.TableDef
    Dim fldNew As DAO.Field
    Dim rstDst As DAO.Recordset
    Dim blnDaoTx As Boolean
    Dim lngRowIdx As Long
    Dim vntBuffered As Variant

    EnsureReady
    If Len(Trim$(strSqliteTableOrSql)) = 0 Or Len(Trim$(strTargetAccessTable)) = 0 Then
        RaiseSQLiteError sqErrInvalidArgument, "clsSQLiteAccessBridge.ImportSQLiteToAccess", _
            "Source SQLite table/SQL and target Access table name are required."
    End If

    If IsSelectSql(strSqliteTableOrSql) Then
        strSql = strSqliteTableOrSql
    Else
        strSql = "SELECT * FROM " & QuoteSQLiteIdent(strSqliteTableOrSql) & ";"
    End If

    If IsMissing(varParamList) Then
        varPassed = Empty
    Else
        varPassed = varParamList
    End If
    Set objStmt = m_Conn.PrepareStatement(strSql, False)
    On Error GoTo ImportCleanFail

    objStmt.BindFlexibleParams varPassed
    lngColCount = objStmt.ColumnCount()
    If lngColCount <= 0 Then
        objStmt.FinalizeStatement
        RaiseSQLiteError sqErrBridge, "clsSQLiteAccessBridge.ImportSQLiteToAccess", _
            "SQLite query returned 0 result columns."
    End If

    ReDim astrColNames(0 To lngColCount - 1)
    ReDim astrDeclTypes(0 To lngColCount - 1)
    ReDim alngDaoTypes(0 To lngColCount - 1)

    For lngCol = 0 To lngColCount - 1
        astrColNames(lngCol) = objStmt.ColumnName(lngCol)
        If Len(astrColNames(lngCol)) = 0 Then astrColNames(lngCol) = "Field" & (lngCol + 1)
        astrDeclTypes(lngCol) = UCase$(Trim$(objStmt.ColumnDeclType(lngCol)))
        alngDaoTypes(lngCol) = MapDeclTypeToDaoType(astrDeclTypes(lngCol))
    Next lngCol

    ' Buffer rows so we can upgrade >255-char strings to dbMemo and refine untyped expression columns.
    Set colBufferedRows = New Collection
    Do
        lngRc = objStmt.StepOnce()
        If lngRc = sqDone Then Exit Do
        ReDim avntRow(0 To lngColCount - 1)
        For lngCol = 0 To lngColCount - 1
            varVal = objStmt.ReadColumnValue(lngCol, True, astrDeclTypes(lngCol))
            avntRow(lngCol) = varVal
            If Not IsNull(varVal) Then
                UpgradeDaoTypeFromSample alngDaoTypes(lngCol), astrDeclTypes(lngCol), varVal
            End If
        Next lngCol
        colBufferedRows.Add avntRow
    Loop
    objStmt.FinalizeStatement
    Set objStmt = Nothing

    For lngCol = 0 To lngColCount - 1
        If alngDaoTypes(lngCol) = 0 Then alngDaoTypes(lngCol) = dbText
    Next lngCol

    If AccessTableExists(strTargetAccessTable) Then
        If blnOverwriteTable Then
            m_DaoDb.TableDefs.Delete strTargetAccessTable
            m_DaoDb.TableDefs.Refresh
        End If
    End If

    If Not AccessTableExists(strTargetAccessTable) Then
        Set tdfNew = m_DaoDb.CreateTableDef(strTargetAccessTable)
        For lngCol = 0 To lngColCount - 1
            If alngDaoTypes(lngCol) = dbText Then
                Set fldNew = tdfNew.CreateField(astrColNames(lngCol), dbText, 255)
                fldNew.AllowZeroLength = True
            ElseIf alngDaoTypes(lngCol) = dbMemo Then
                Set fldNew = tdfNew.CreateField(astrColNames(lngCol), dbMemo)
                fldNew.AllowZeroLength = True
            Else
                Set fldNew = tdfNew.CreateField(astrColNames(lngCol), alngDaoTypes(lngCol))
            End If
            tdfNew.Fields.Append fldNew
        Next lngCol
        m_DaoDb.TableDefs.Append tdfNew
        m_DaoDb.TableDefs.Refresh
    End If

    If colBufferedRows.Count > 0 Then
        DBEngine.Workspaces(0).BeginTrans
        blnDaoTx = True
        Set rstDst = m_DaoDb.OpenRecordset(strTargetAccessTable, dbOpenDynaset, dbAppendOnly)
        For lngRowIdx = 1 To colBufferedRows.Count
            vntBuffered = colBufferedRows(lngRowIdx)
            rstDst.AddNew
            For lngCol = 0 To lngColCount - 1
                If Not IsNull(vntBuffered(lngCol)) Then
                    AssignValueToDaoField rstDst.Fields(lngCol), vntBuffered(lngCol)
                End If
            Next lngCol
            rstDst.Update
        Next lngRowIdx
        rstDst.Close
        Set rstDst = Nothing
        DBEngine.Workspaces(0).CommitTrans
        blnDaoTx = False
    End If

    ImportSQLiteToAccess = colBufferedRows.Count
    Exit Function

ImportCleanFail:
    Dim lngErrNum As Long
    Dim strErrSrc As String
    Dim strErrDesc As String
    lngErrNum = Err.Number
    strErrSrc = Err.Source
    strErrDesc = Err.Description
    On Error Resume Next
    If Not objStmt Is Nothing Then objStmt.FinalizeStatement
    If Not rstDst Is Nothing Then rstDst.Close
    If blnDaoTx Then DBEngine.Workspaces(0).Rollback
    On Error GoTo 0
    RaiseSQLiteError sqErrBridge, "clsSQLiteAccessBridge.ImportSQLiteToAccess", _
        "Import failed: " & strErrDesc & " (" & strErrSrc & " #" & lngErrNum & ")"
End Function

' =============================================================================
' Internal Type Mapping Helpers
' =============================================================================

Private Sub BindDaoFieldToSqlite(ByVal objStmt As clsSQLiteStatement, ByVal lngOneBasedIdx As Long, ByVal objFld As DAO.Field)
    Dim varVal As Variant
    Dim abytBlob() As Byte

    varVal = objFld.Value
    If IsNull(varVal) Or IsEmpty(varVal) Then
        objStmt.BindByIndex lngOneBasedIdx, Null
        Exit Sub
    End If

    Select Case objFld.Type
        Case dbBoolean
            objStmt.BindByIndex lngOneBasedIdx, CBool(varVal)
        Case dbByte, dbInteger, dbLong
            objStmt.BindByIndex lngOneBasedIdx, CLng(varVal)
        Case dbBigInt
            objStmt.BindByIndex lngOneBasedIdx, CCur(varVal)
        Case dbCurrency, dbDecimal, dbNumeric
            objStmt.BindByIndex lngOneBasedIdx, CDbl(varVal)
        Case dbSingle, dbDouble, dbFloat
            objStmt.BindByIndex lngOneBasedIdx, CDbl(varVal)
        Case dbDate, dbTime, dbTimeStamp
            objStmt.BindByIndex lngOneBasedIdx, FormatIsoDateValue(CDate(varVal), (objFld.Type <> dbDate Or CDate(varVal) <> Int(CDate(varVal))))
        Case dbLongBinary, dbVarBinary
            abytBlob = varVal
            objStmt.BindByIndex lngOneBasedIdx, abytBlob
        Case Else
            objStmt.BindByIndex lngOneBasedIdx, CStr(varVal)
    End Select
End Sub

Private Sub AssignValueToDaoField(ByVal objFld As DAO.Field, ByVal varVal As Variant)
    Dim dtmParsed As Date
    Dim abytBlob() As Byte

    Select Case objFld.Type
        Case dbBoolean
            objFld.Value = CBool(varVal)
        Case dbByte
            objFld.Value = CByte(varVal)
        Case dbInteger
            objFld.Value = CInt(varVal)
        Case dbLong
            objFld.Value = CLng(varVal)
        Case dbCurrency
            objFld.Value = CCur(varVal)
        Case dbSingle
            objFld.Value = CSng(varVal)
        Case dbDouble
            objFld.Value = CDbl(varVal)
        Case dbDate
            If VarType(varVal) = vbDate Then
                objFld.Value = CDate(varVal)
            ElseIf TryParseIsoDate(CStr(varVal), dtmParsed) Then
                objFld.Value = dtmParsed
            Else
                objFld.Value = CDate(varVal)
            End If
        Case dbLongBinary, dbVarBinary
            If VarType(varVal) = (vbArray Or vbByte) Then
                abytBlob = varVal
                If GetByteArrayLength(abytBlob) > 0 Then
                    objFld.AppendChunk abytBlob
                End If
            End If
        Case Else
            objFld.Value = CStr(varVal)
    End Select
End Sub

Private Function MapDaoFieldToSqliteDdl(ByVal objFld As DAO.Field, ByVal strSinglePkName As String) As String
    Dim blnIsPk As Boolean
    blnIsPk = (Len(strSinglePkName) > 0 And StrComp(objFld.Name, strSinglePkName, vbTextCompare) = 0)

    If (objFld.Attributes And dbAutoIncrField) = dbAutoIncrField And blnIsPk Then
        MapDaoFieldToSqliteDdl = "INTEGER PRIMARY KEY AUTOINCREMENT"
        Exit Function
    End If

    Select Case objFld.Type
        Case dbBoolean
            MapDaoFieldToSqliteDdl = "BOOLEAN"
        Case dbByte, dbInteger, dbLong, dbBigInt
            MapDaoFieldToSqliteDdl = "INTEGER"
        Case dbCurrency
            MapDaoFieldToSqliteDdl = "CURRENCY"
        Case dbDecimal, dbNumeric, dbDouble, dbFloat, dbSingle
            MapDaoFieldToSqliteDdl = "REAL"
        Case dbDate, dbTimeStamp
            MapDaoFieldToSqliteDdl = "DATETIME"
        Case dbTime
            MapDaoFieldToSqliteDdl = "TEXT"
        Case dbLongBinary, dbVarBinary
            MapDaoFieldToSqliteDdl = "BLOB"
        Case dbMemo
            MapDaoFieldToSqliteDdl = "MEMO"
        Case dbGUID
            MapDaoFieldToSqliteDdl = "GUID"
        Case Else
            MapDaoFieldToSqliteDdl = "TEXT"
    End Select

    If blnIsPk Then
        MapDaoFieldToSqliteDdl = MapDaoFieldToSqliteDdl & " PRIMARY KEY"
    ElseIf objFld.Required Then
        MapDaoFieldToSqliteDdl = MapDaoFieldToSqliteDdl & " NOT NULL"
    End If
End Function

Private Function MapDeclTypeToDaoType(ByVal strUpperDecl As String) As Long
    If Len(strUpperDecl) = 0 Then
        MapDeclTypeToDaoType = 0 ' Infer from row values
    ElseIf InStr(1, strUpperDecl, "BOOL", vbBinaryCompare) > 0 Or strUpperDecl = "BIT" Then
        MapDeclTypeToDaoType = dbBoolean
    ElseIf InStr(1, strUpperDecl, "INT", vbBinaryCompare) > 0 Then
        MapDeclTypeToDaoType = dbLong
    ElseIf InStr(1, strUpperDecl, "CURR", vbBinaryCompare) > 0 Or InStr(1, strUpperDecl, "MONEY", vbBinaryCompare) > 0 Then
        MapDeclTypeToDaoType = dbCurrency
    ElseIf InStr(1, strUpperDecl, "REAL", vbBinaryCompare) > 0 Or InStr(1, strUpperDecl, "DOUB", vbBinaryCompare) > 0 Or InStr(1, strUpperDecl, "FLOA", vbBinaryCompare) > 0 Or InStr(1, strUpperDecl, "NUM", vbBinaryCompare) > 0 Or InStr(1, strUpperDecl, "DEC", vbBinaryCompare) > 0 Then
        MapDeclTypeToDaoType = dbDouble
    ElseIf InStr(1, strUpperDecl, "DATE", vbBinaryCompare) > 0 Or InStr(1, strUpperDecl, "TIME", vbBinaryCompare) > 0 Then
        MapDeclTypeToDaoType = dbDate
    ElseIf InStr(1, strUpperDecl, "BLOB", vbBinaryCompare) > 0 Or InStr(1, strUpperDecl, "BIN", vbBinaryCompare) > 0 Then
        MapDeclTypeToDaoType = dbLongBinary
    ElseIf InStr(1, strUpperDecl, "MEMO", vbBinaryCompare) > 0 Or InStr(1, strUpperDecl, "JSON", vbBinaryCompare) > 0 Or InStr(1, strUpperDecl, "CLOB", vbBinaryCompare) > 0 Then
        MapDeclTypeToDaoType = dbMemo
    Else
        MapDeclTypeToDaoType = dbText
    End If
End Function

Private Sub UpgradeDaoTypeFromSample(ByRef lngDaoType As Long, ByVal strUpperDecl As String, ByVal varSample As Variant)
    Dim intVt As Integer
    intVt = VarType(varSample)

    If lngDaoType = 0 Then
        Select Case intVt
            Case vbBoolean
                lngDaoType = dbBoolean
            Case vbByte, vbInteger, vbLong
                lngDaoType = dbLong
            Case vbCurrency
                ' Large 64-bit integers beyond Long map safely to Double/Currency in Access
                lngDaoType = dbDouble
            Case vbSingle, vbDouble, vbDecimal
                lngDaoType = dbDouble
            Case vbDate
                lngDaoType = dbDate
            Case (vbArray Or vbByte)
                lngDaoType = dbLongBinary
            Case Else
                If Len(CStr(varSample)) > 255 Then
                    lngDaoType = dbMemo
                Else
                    lngDaoType = dbText
                End If
        End Select
    ElseIf lngDaoType = dbLong And (intVt = vbCurrency Or intVt = vbDouble) Then
        If CDbl(varSample) > 2147483647# Or CDbl(varSample) < -2147483648# Then
            lngDaoType = dbDouble
        End If
    ElseIf lngDaoType = dbText And intVt = vbString Then
        If Len(CStr(varSample)) > 255 Then
            lngDaoType = dbMemo
        End If
    End If
End Sub

Private Function FindAccessSinglePkField(ByVal strTableName As String) As String
    On Error GoTo NoPk
    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index
    Set tdf = m_DaoDb.TableDefs(strTableName)
    For Each idx In tdf.Indexes
        If idx.Primary And idx.Fields.Count = 1 Then
            FindAccessSinglePkField = idx.Fields(0).Name
            Exit Function
        End If
    Next idx
NoPk:
    On Error GoTo 0
End Function

Private Function AccessTableExists(ByVal strTableName As String) As Boolean
    Dim tdf As DAO.TableDef
    For Each tdf In m_DaoDb.TableDefs
        If StrComp(tdf.Name, strTableName, vbTextCompare) = 0 Then
            AccessTableExists = True
            Exit Function
        End If
    Next tdf
End Function

Private Function IsSelectSql(ByVal strInput As String) As Boolean
    Dim strHead As String
    strHead = UCase$(Left$(LTrim$(strInput), 7))
    IsSelectSql = (Left$(strHead, 7) = "SELECT " Or Left$(strHead, 5) = "WITH " Or Left$(strHead, 7) = "PRAGMA ")
End Function

Private Sub EnsureReady()
    If m_Conn Is Nothing Or Not m_Conn.IsOpen Then
        RaiseSQLiteError sqErrConnectionClosed, "clsSQLiteAccessBridge", "An open clsSQLiteConnection is required."
    End If
    If m_DaoDb Is Nothing Then Set m_DaoDb = CurrentDb
End Sub

