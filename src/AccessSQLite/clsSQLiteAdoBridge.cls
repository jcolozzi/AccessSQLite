Attribute VB_Name = "clsSQLiteAdoBridge"
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
' clsSQLiteAdoBridge
' Disconnected in-memory ADODB.Recordset factory & two-way SQLite bridge.
' Enables zero-temp-table Access Continuous Forms (Set Me.Recordset = rs),
' virtual multi-select checkbox columns (e.g. IsSelected), inline row writeback
' to SQLite via clsSQLiteRepository, and external ADODB -> SQLite bulk imports.
' Uses late-bound CreateObject("ADODB.Recordset") so consuming .accdb projects
' need zero extra Tools -> References (while still compatible with early-bound
' ADODB.Recordset variables via COM QueryInterface).
' =============================================================================

' ADODB CursorLocationEnum / CursorTypeEnum / LockTypeEnum
Private Const ADO_USE_CLIENT As Long = 3
Private Const ADO_OPEN_STATIC As Long = 3
Private Const ADO_LOCK_OPTIMISTIC As Long = 3

' ADODB DataTypeEnum (Access form-safe types)
Private Const ADO_INTEGER As Long = 3
Private Const ADO_SINGLE As Long = 4
Private Const ADO_DOUBLE As Long = 5
Private Const ADO_CURRENCY As Long = 6
Private Const ADO_DATE As Long = 7
Private Const ADO_BOOLEAN As Long = 11
Private Const ADO_DECIMAL As Long = 14
Private Const ADO_TINYINT As Long = 16
Private Const ADO_UNSIGNED_TINYINT As Long = 17
Private Const ADO_SMALLINT As Long = 2
Private Const ADO_BIGINT As Long = 20
Private Const ADO_GUID As Long = 72
Private Const ADO_BINARY As Long = 128
Private Const ADO_CHAR As Long = 129
Private Const ADO_WCHAR As Long = 130
Private Const ADO_NUMERIC As Long = 131
Private Const ADO_DBDATE As Long = 133
Private Const ADO_DBTIME As Long = 134
Private Const ADO_DBTIMESTAMP As Long = 135
Private Const ADO_VARCHAR As Long = 200
Private Const ADO_LONGVARCHAR As Long = 201
Private Const ADO_VARWCHAR As Long = 202
Private Const ADO_LONGVARWCHAR As Long = 203
Private Const ADO_VARBINARY As Long = 204
Private Const ADO_LONGVARBINARY As Long = 205

' ADODB FieldAttributeEnum
Private Const ADO_FLD_UPDATABLE As Long = &H4
Private Const ADO_FLD_UNKNOWN_UPDATABLE As Long = &H8
Private Const ADO_FLD_IS_NULLABLE As Long = &H20
Private Const ADO_FLD_ATTR_NULLABLE As Long = &H2C   ' 4 + 8 + 32
Private Const ADO_FLD_ATTR_REQUIRED As Long = &HC    ' 4 + 8
Private Const ADO_MEMO_SIZE As Long = 1073741823

Private m_Conn As clsSQLiteConnection

Public Sub InitAdoBridge(ByVal objConn As clsSQLiteConnection)
    Set m_Conn = objConn
End Sub

' Executes a parameterized SQLite SELECT query and materializes a disconnected,
' updatable client-side ADODB.Recordset ready for `Set Me.Recordset = rs`.
Public Function QueryToAdoRecordset( _
    ByVal strSql As String, _
    Optional ByVal varParamList As Variant, _
    Optional ByVal objTableMap As clsSQLiteTableMap = Nothing, _
    Optional ByVal blnIncludeSelectColumn As Boolean = False, _
    Optional ByVal strSelectColumnName As String = "IsSelected") As Object

    Dim varPassed As Variant
    Dim objStmt As clsSQLiteStatement
    Dim lngColCount As Long
    Dim lngCol As Long
    Dim astrNames() As String
    Dim astrDeclTypes() As String
    Dim alngAdoTypes() As Long
    Dim alngSizes() As Long
    Dim dictUsedNames As Scripting.Dictionary
    Dim colBuffered As Collection
    Dim avntRow() As Variant
    Dim lngRc As Long
    Dim varVal As Variant
    Dim objRs As Object
    Dim strSelCol As String
    Dim lngRowIdx As Long
    Dim vntBufRow As Variant
    Dim objColDef As clsSQLiteColumnDef

    EnsureConn
    If IsObject(varParamList) Then
        Set varPassed = varParamList
    ElseIf IsMissing(varParamList) Then
        varPassed = Empty
    Else
        varPassed = varParamList
    End If

    Set objStmt = m_Conn.PrepareStatement(strSql, True)
    On Error GoTo CleanFail

    objStmt.BindFlexibleParams varPassed
    lngColCount = objStmt.ColumnCount()
    If lngColCount <= 0 Then
        objStmt.ResetForReuse
        If Not objStmt.IsCached Then objStmt.FinalizeStatement
        RaiseSQLiteError sqErrBridge, "clsSQLiteAdoBridge.QueryToAdoRecordset", _
            "SQLite query did not return any result columns."
    End If

    ReDim astrNames(0 To lngColCount - 1)
    ReDim astrDeclTypes(0 To lngColCount - 1)
    ReDim alngAdoTypes(0 To lngColCount - 1)
    ReDim alngSizes(0 To lngColCount - 1)
    Set dictUsedNames = NewTextDict()

    strSelCol = Trim$(strSelectColumnName)
    If Len(strSelCol) = 0 Then strSelCol = "IsSelected"
    If blnIncludeSelectColumn Then dictUsedNames(strSelCol) = True

    For lngCol = 0 To lngColCount - 1
        astrNames(lngCol) = MakeUniqueFieldName(objStmt.ColumnName(lngCol), lngCol, dictUsedNames)
        astrDeclTypes(lngCol) = UCase$(Trim$(objStmt.ColumnDeclType(lngCol)))

        Set objColDef = Nothing
        If Not objTableMap Is Nothing Then
            Set objColDef = objTableMap.GetColumn(astrNames(lngCol))
        End If

        If Not objColDef Is Nothing Then
            MapColumnDefToAdo objColDef, alngAdoTypes(lngCol), alngSizes(lngCol)
        Else
            MapDeclTypeToAdo astrDeclTypes(lngCol), alngAdoTypes(lngCol), alngSizes(lngCol)
        End If
    Next lngCol

    ' Buffer rows first so untyped expressions and >255-char strings get accurate ADO field definitions.
    Set colBuffered = New Collection
    Do
        lngRc = objStmt.StepOnce()
        If lngRc = sqDone Then Exit Do
        ReDim avntRow(0 To lngColCount - 1)
        For lngCol = 0 To lngColCount - 1
            varVal = objStmt.ReadColumnValue(lngCol, True, astrDeclTypes(lngCol))
            If Not objTableMap Is Nothing And Not IsNull(varVal) Then
                Set objColDef = objTableMap.GetColumn(astrNames(lngCol))
                If Not objColDef Is Nothing Then
                    varVal = CoerceForAdoField(varVal, objColDef.ColumnKind)
                End If
            End If
            avntRow(lngCol) = varVal
            If Not IsNull(varVal) Then
                RefineAdoFieldFromSample alngAdoTypes(lngCol), alngSizes(lngCol), varVal
            End If
        Next lngCol
        colBuffered.Add avntRow
    Loop

    objStmt.ResetForReuse
    If Not objStmt.IsCached Then objStmt.FinalizeStatement
    Set objStmt = Nothing

    ' Default any remaining untyped all-NULL expression columns to VarWChar(255).
    For lngCol = 0 To lngColCount - 1
        If alngAdoTypes(lngCol) = 0 Then
            alngAdoTypes(lngCol) = ADO_VARWCHAR
            alngSizes(lngCol) = 255
        End If
    Next lngCol

    Set objRs = CreateDisconnectedRecordset()
    If blnIncludeSelectColumn Then
        objRs.Fields.Append strSelCol, ADO_BOOLEAN, 0, ADO_FLD_ATTR_REQUIRED
    End If

    For lngCol = 0 To lngColCount - 1
        AppendAdoField objRs, astrNames(lngCol), alngAdoTypes(lngCol), alngSizes(lngCol)
    Next lngCol

    objRs.Open

    For lngRowIdx = 1 To colBuffered.Count
        vntBufRow = colBuffered(lngRowIdx)
        objRs.AddNew
        If blnIncludeSelectColumn Then
            objRs.Fields.Item(strSelCol).Value = False
        End If
        For lngCol = 0 To lngColCount - 1
            AssignAdoFieldValue objRs.Fields.Item(astrNames(lngCol)), alngAdoTypes(lngCol), vntBufRow(lngCol)
        Next lngCol
        objRs.Update
    Next lngRowIdx

    If objRs.RecordCount > 0 Then objRs.MoveFirst
    Set QueryToAdoRecordset = objRs
    Exit Function

CleanFail:
    Dim lngErrNum As Long
    Dim strErrSrc As String
    Dim strErrDesc As String
    lngErrNum = Err.Number
    strErrSrc = Err.Source
    strErrDesc = Err.Description
    On Error Resume Next
    If Not objStmt Is Nothing Then
        objStmt.ResetForReuse
        If Not objStmt.IsCached Then objStmt.FinalizeStatement
    End If
    If Not objRs Is Nothing Then
        If objRs.State <> 0 Then objRs.Close
    End If
    On Error GoTo 0
    RaiseSQLiteError sqErrBridge, "clsSQLiteAdoBridge.QueryToAdoRecordset", _
        "ADO Recordset build failed: " & strErrDesc & " (" & strErrSrc & " #" & lngErrNum & ")"
End Function

' Creates an empty, open, updatable disconnected ADODB.Recordset from a clsSQLiteTableMap.
Public Function CreateEmptyRecordsetFromMap( _
    ByVal objTableMap As clsSQLiteTableMap, _
    Optional ByVal blnIncludeSelectColumn As Boolean = False, _
    Optional ByVal strSelectColumnName As String = "IsSelected") As Object

    Dim objRs As Object
    Dim objCol As clsSQLiteColumnDef
    Dim lngAdoType As Long
    Dim lngSize As Long
    Dim strSelCol As String

    If objTableMap Is Nothing Or objTableMap.Columns.Count = 0 Then
        RaiseSQLiteError sqErrInvalidArgument, "clsSQLiteAdoBridge.CreateEmptyRecordsetFromMap", _
            "A non-empty clsSQLiteTableMap is required."
    End If

    Set objRs = CreateDisconnectedRecordset()
    strSelCol = IIf(Len(Trim$(strSelectColumnName)) > 0, Trim$(strSelectColumnName), "IsSelected")
    If blnIncludeSelectColumn Then
        objRs.Fields.Append strSelCol, ADO_BOOLEAN, 0, ADO_FLD_ATTR_REQUIRED
    End If

    For Each objCol In objTableMap.Columns
        MapColumnDefToAdo objCol, lngAdoType, lngSize
        AppendAdoField objRs, objCol.ColumnName, lngAdoType, lngSize
    Next objCol

    objRs.Open
    Set CreateEmptyRecordsetFromMap = objRs
End Function

' Converts the current row of an ADODB.Recordset into a case-insensitive Scripting.Dictionary.
Public Function AdoRowToDictionary( _
    ByVal objAdoRs As Object, _
    Optional ByVal strExcludeColumn As String = "IsSelected") As Scripting.Dictionary

    Dim dictRow As Scripting.Dictionary
    Dim objFld As Object
    Dim strName As String

    If objAdoRs Is Nothing Then
        RaiseSQLiteError sqErrInvalidArgument, "clsSQLiteAdoBridge.AdoRowToDictionary", "ADODB.Recordset cannot be Nothing."
    End If
    If objAdoRs.BOF Or objAdoRs.EOF Then
        Set AdoRowToDictionary = NewTextDict()
        Exit Function
    End If

    Set dictRow = NewTextDict()
    For Each objFld In objAdoRs.Fields
        strName = CStr(objFld.Name)
        If Len(strExcludeColumn) = 0 Or StrComp(strName, strExcludeColumn, vbTextCompare) <> 0 Then
            dictRow(strName) = objFld.Value
        End If
    Next objFld
    Set AdoRowToDictionary = dictRow
End Function

' Persists the current row of a form-bound in-memory ADODB.Recordset into SQLite via clsSQLiteRepository,
' hydrating any generated primary key and timestamps back onto the in-memory ADO row.
Public Function SaveAdoRowToRepository( _
    ByVal objAdoRs As Object, _
    ByVal objRepo As clsSQLiteRepository, _
    Optional ByVal blnDirtyOnly As Boolean = True, _
    Optional ByVal strExcludeColumn As String = "IsSelected") As Variant

    Dim dictRow As Scripting.Dictionary
    Dim dictMapped As Scripting.Dictionary
    Dim objMap As clsSQLiteTableMap
    Dim strPkCol As String
    Dim objPkDef As clsSQLiteColumnDef
    Dim varPkVal As Variant
    Dim blnIsInsert As Boolean
    Dim varSavedId As Variant
    Dim lngUpdated As Long
    Dim dictReloaded As Scripting.Dictionary
    Dim varKey As Variant

    If objAdoRs Is Nothing Or objRepo Is Nothing Then
        RaiseSQLiteError sqErrInvalidArgument, "clsSQLiteAdoBridge.SaveAdoRowToRepository", _
            "Both ADODB.Recordset and clsSQLiteRepository are required."
    End If
    If objAdoRs.BOF Or objAdoRs.EOF Then Exit Function

    Set objMap = objRepo.TableMap
    strPkCol = objMap.PrimaryKeyColumn
    Set objPkDef = objMap.GetColumn(strPkCol)

    Set dictRow = AdoRowToDictionary(objAdoRs, strExcludeColumn)
    Set dictMapped = NewTextDict()
    For Each varKey In dictRow.Keys
        If objMap.HasColumn(CStr(varKey)) Then
            dictMapped(CStr(varKey)) = dictRow(varKey)
        End If
    Next varKey

    If dictMapped.Exists(strPkCol) Then
        varPkVal = dictMapped(strPkCol)
    Else
        varPkVal = Empty
    End If

    If Not objPkDef Is Nothing And objPkDef.IsAutoIncrement Then
        blnIsInsert = IsUnsetKey(varPkVal)
    ElseIf IsUnsetKey(varPkVal) Then
        blnIsInsert = True
    Else
        blnIsInsert = Not objRepo.ExistsById(varPkVal, True)
    End If

    If blnIsInsert Then
        varSavedId = objRepo.InsertRecord(dictMapped, False)
        lngUpdated = 1
    Else
        lngUpdated = objRepo.UpdateRecord(varPkVal, dictMapped, blnDirtyOnly, False)
        varSavedId = varPkVal
    End If

    ' Sync generated PK and UpdatedAt/CreatedAt back onto the live in-memory ADO row
    If lngUpdated > 0 Then
        Set dictReloaded = objRepo.FindByIdRaw(varSavedId, True)
        If Not dictReloaded Is Nothing Then
            SyncDictIntoCurrentAdoRow objAdoRs, dictReloaded
        End If
    End If

    SaveAdoRowToRepository = varSavedId
End Function

' Returns a Collection of primary key values where the virtual selection checkbox column is True.
' Uses Recordset.Clone so a bound Continuous Form's active row and scroll position never jump.
Public Function GetSelectedKeys( _
    ByVal objAdoRs As Object, _
    ByVal strPrimaryKeyCol As String, _
    Optional ByVal strSelectColumnName As String = "IsSelected") As Collection

    Dim colKeys As Collection
    Dim objClone As Object
    Dim strSelCol As String

    Set colKeys = New Collection
    If objAdoRs Is Nothing Then
        Set GetSelectedKeys = colKeys
        Exit Function
    End If
    If objAdoRs.RecordCount = 0 Then
        Set GetSelectedKeys = colKeys
        Exit Function
    End If

    strSelCol = IIf(Len(Trim$(strSelectColumnName)) > 0, Trim$(strSelectColumnName), "IsSelected")
    Set objClone = objAdoRs.Clone
    objClone.MoveFirst
    Do Until objClone.EOF
        If CBool(Nz(objClone.Fields.Item(strSelCol).Value, False)) Then
            If Not IsNull(objClone.Fields.Item(strPrimaryKeyCol).Value) Then
                colKeys.Add objClone.Fields.Item(strPrimaryKeyCol).Value
            End If
        End If
        objClone.MoveNext
    Loop
    objClone.Close
    Set GetSelectedKeys = colKeys
End Function

' Checks or unchecks all rows on the virtual selection checkbox column in memory via Recordset.Clone.
Public Function SetAllSelected( _
    ByVal objAdoRs As Object, _
    ByVal blnSelected As Boolean, _
    Optional ByVal strSelectColumnName As String = "IsSelected") As Long

    Dim objClone As Object
    Dim strSelCol As String
    Dim lngTouched As Long

    If objAdoRs Is Nothing Then Exit Function
    If objAdoRs.RecordCount = 0 Then Exit Function

    strSelCol = IIf(Len(Trim$(strSelectColumnName)) > 0, Trim$(strSelectColumnName), "IsSelected")
    Set objClone = objAdoRs.Clone
    objClone.MoveFirst
    Do Until objClone.EOF
        If CBool(Nz(objClone.Fields.Item(strSelCol).Value, False)) <> blnSelected Then
            objClone.Fields.Item(strSelCol).Value = blnSelected
            objClone.Update
        End If
        lngTouched = lngTouched + 1
        objClone.MoveNext
    Loop
    objClone.Close
    SetAllSelected = lngTouched
End Function

' Imports any open ADODB.Recordset (in-memory or external OLEDB/ODBC) into a SQLite table inside one transaction.
Public Function ImportAdoRecordsetToSQLite( _
    ByVal objAdoRs As Object, _
    ByVal strTargetSqliteTable As String, _
    Optional ByVal blnOverwriteTable As Boolean = False, _
    Optional ByVal strExcludeColumn As String = "IsSelected") As Long

    Dim colIncludedIdx As Collection
    Dim lngFldIdx As Long
    Dim objFld As Object
    Dim strColsDdl As String
    Dim strInsertCols As String
    Dim strPlaceholders As String
    Dim lngIncludedCount As Long
    Dim lngPos As Long
    Dim objStmt As clsSQLiteStatement
    Dim objReader As Object
    Dim blnUsedClone As Boolean
    Dim lngImported As Long

    EnsureConn
    If objAdoRs Is Nothing Or Len(Trim$(strTargetSqliteTable)) = 0 Then
        RaiseSQLiteError sqErrInvalidArgument, "clsSQLiteAdoBridge.ImportAdoRecordsetToSQLite", _
            "Open ADODB.Recordset and target SQLite table name are required."
    End If

    Set colIncludedIdx = New Collection
    For lngFldIdx = 0 To objAdoRs.Fields.Count - 1
        Set objFld = objAdoRs.Fields.Item(lngFldIdx)
        If Len(strExcludeColumn) = 0 Or StrComp(CStr(objFld.Name), strExcludeColumn, vbTextCompare) <> 0 Then
            colIncludedIdx.Add lngFldIdx
            If lngIncludedCount > 0 Then
                strColsDdl = strColsDdl & ", "
                strInsertCols = strInsertCols & ", "
                strPlaceholders = strPlaceholders & ", "
            End If
            strColsDdl = strColsDdl & QuoteSQLiteIdent(CStr(objFld.Name)) & " " & MapAdoTypeToSqliteDdl(CLng(objFld.Type))
            strInsertCols = strInsertCols & QuoteSQLiteIdent(CStr(objFld.Name))
            strPlaceholders = strPlaceholders & "?"
            lngIncludedCount = lngIncludedCount + 1
        End If
    Next lngFldIdx

    If lngIncludedCount = 0 Then
        RaiseSQLiteError sqErrBridge, "clsSQLiteAdoBridge.ImportAdoRecordsetToSQLite", _
            "ADODB.Recordset has no exportable columns."
    End If

    On Error Resume Next
    Set objReader = objAdoRs.Clone
    If Err.Number = 0 And Not objReader Is Nothing Then
        blnUsedClone = True
    Else
        Err.Clear
        Set objReader = objAdoRs
    End If
    On Error GoTo ImportFail

    m_Conn.BeginTrans
    If blnOverwriteTable Then
        m_Conn.SchemaManager().DropTableIfExists strTargetSqliteTable
    End If
    m_Conn.ExecuteParams _
        "CREATE TABLE IF NOT EXISTS " & QuoteSQLiteIdent(strTargetSqliteTable) & " (" & strColsDdl & ");", Empty

    If Not (objReader.BOF And objReader.EOF) Then
        objReader.MoveFirst
        Set objStmt = m_Conn.PrepareStatement( _
            "INSERT INTO " & QuoteSQLiteIdent(strTargetSqliteTable) & _
            " (" & strInsertCols & ") VALUES (" & strPlaceholders & ");", False)

        Do Until objReader.EOF
            objStmt.ResetForReuse
            For lngPos = 1 To colIncludedIdx.Count
                objStmt.BindByIndex lngPos, objReader.Fields.Item(CLng(colIncludedIdx(lngPos))).Value
            Next lngPos
            objStmt.StepOnce
            lngImported = lngImported + 1
            objReader.MoveNext
        Loop
        objStmt.FinalizeStatement
        Set objStmt = Nothing
    End If

    m_Conn.CommitTrans
    If blnUsedClone Then objReader.Close
    ImportAdoRecordsetToSQLite = lngImported
    Exit Function

ImportFail:
    Dim lngErrNum As Long
    Dim strErrSrc As String
    Dim strErrDesc As String
    lngErrNum = Err.Number
    strErrSrc = Err.Source
    strErrDesc = Err.Description
    On Error Resume Next
    If Not objStmt Is Nothing Then objStmt.FinalizeStatement
    If m_Conn.InTransaction Then m_Conn.RollbackTrans
    If blnUsedClone And Not objReader Is Nothing Then objReader.Close
    On Error GoTo 0
    RaiseSQLiteError sqErrBridge, "clsSQLiteAdoBridge.ImportAdoRecordsetToSQLite", _
        "ADO import to SQLite failed: " & strErrDesc & " (" & strErrSrc & " #" & lngErrNum & ")"
End Function

' =============================================================================
' Internal Helpers
' =============================================================================

Private Function CreateDisconnectedRecordset() As Object
    Dim objRs As Object
    Set objRs = CreateObject("ADODB.Recordset")
    objRs.CursorLocation = ADO_USE_CLIENT
    objRs.CursorType = ADO_OPEN_STATIC
    objRs.LockType = ADO_LOCK_OPTIMISTIC
    Set CreateDisconnectedRecordset = objRs
End Function

Private Sub AppendAdoField(ByVal objRs As Object, ByVal strName As String, ByVal lngAdoType As Long, ByVal lngSize As Long)
    If lngAdoType = ADO_BOOLEAN Then
        objRs.Fields.Append strName, ADO_BOOLEAN, 0, ADO_FLD_ATTR_REQUIRED
    ElseIf lngAdoType = ADO_VARWCHAR Then
        objRs.Fields.Append strName, ADO_VARWCHAR, IIf(lngSize > 0, lngSize, 255), ADO_FLD_ATTR_NULLABLE
    ElseIf lngAdoType = ADO_LONGVARWCHAR Then
        objRs.Fields.Append strName, ADO_LONGVARWCHAR, ADO_MEMO_SIZE, ADO_FLD_ATTR_NULLABLE
    ElseIf lngAdoType = ADO_LONGVARBINARY Then
        objRs.Fields.Append strName, ADO_LONGVARBINARY, ADO_MEMO_SIZE, ADO_FLD_ATTR_NULLABLE
    Else
        objRs.Fields.Append strName, lngAdoType, 0, ADO_FLD_ATTR_NULLABLE
    End If
End Sub

Private Sub AssignAdoFieldValue(ByVal objFld As Object, ByVal lngAdoType As Long, ByVal varVal As Variant)
    Dim dtmParsed As Date
    Dim abytBlob() As Byte

    If IsNull(varVal) Or IsEmpty(varVal) Then
        If lngAdoType = ADO_BOOLEAN Then
            objFld.Value = False
        Else
            objFld.Value = Null
        End If
        Exit Sub
    End If

    Select Case lngAdoType
        Case ADO_BOOLEAN
            objFld.Value = CBool(varVal)
        Case ADO_INTEGER
            objFld.Value = CLng(varVal)
        Case ADO_DOUBLE
            objFld.Value = CDbl(varVal)
        Case ADO_CURRENCY
            objFld.Value = CCur(varVal)
        Case ADO_DATE
            If VarType(varVal) = vbDate Then
                objFld.Value = CDate(varVal)
            ElseIf TryParseIsoDate(CStr(varVal), dtmParsed) Then
                objFld.Value = dtmParsed
            Else
                objFld.Value = CDate(varVal)
            End If
        Case ADO_LONGVARBINARY
            If VarType(varVal) = (vbArray Or vbByte) Then
                abytBlob = varVal
                If GetByteArrayLength(abytBlob) > 0 Then
                    objFld.Value = abytBlob
                Else
                    objFld.Value = Null
                End If
            End If
        Case Else
            objFld.Value = CStr(varVal)
    End Select
End Sub

Private Sub MapColumnDefToAdo(ByVal objCol As clsSQLiteColumnDef, ByRef lngAdoType As Long, ByRef lngSize As Long)
    lngSize = 0
    Select Case objCol.ColumnKind
        Case sqColBoolean
            lngAdoType = ADO_BOOLEAN
        Case sqColLong, sqColInt64
            lngAdoType = ADO_INTEGER
        Case sqColDouble
            lngAdoType = ADO_DOUBLE
        Case sqColCurrency
            lngAdoType = ADO_CURRENCY
        Case sqColDate, sqColDateTime
            lngAdoType = ADO_DATE
        Case sqColBlob
            lngAdoType = ADO_LONGVARBINARY
            lngSize = ADO_MEMO_SIZE
        Case sqColMemo, sqColJson
            lngAdoType = ADO_LONGVARWCHAR
            lngSize = ADO_MEMO_SIZE
        Case Else
            If objCol.MaxLength > 255 Then
                lngAdoType = ADO_LONGVARWCHAR
                lngSize = ADO_MEMO_SIZE
            ElseIf objCol.MaxLength > 0 Then
                lngAdoType = ADO_VARWCHAR
                lngSize = objCol.MaxLength
            Else
                lngAdoType = ADO_VARWCHAR
                lngSize = 255
            End If
    End Select
End Sub

Private Sub MapDeclTypeToAdo(ByVal strUpperDecl As String, ByRef lngAdoType As Long, ByRef lngSize As Long)
    lngSize = 0
    If Len(strUpperDecl) = 0 Then
        lngAdoType = 0
    ElseIf InStr(1, strUpperDecl, "BOOL", vbBinaryCompare) > 0 Or strUpperDecl = "BIT" Then
        lngAdoType = ADO_BOOLEAN
    ElseIf InStr(1, strUpperDecl, "INT", vbBinaryCompare) > 0 Then
        lngAdoType = ADO_INTEGER
    ElseIf InStr(1, strUpperDecl, "CURR", vbBinaryCompare) > 0 Or InStr(1, strUpperDecl, "MONEY", vbBinaryCompare) > 0 Then
        lngAdoType = ADO_CURRENCY
    ElseIf InStr(1, strUpperDecl, "REAL", vbBinaryCompare) > 0 Or InStr(1, strUpperDecl, "DOUB", vbBinaryCompare) > 0 Or InStr(1, strUpperDecl, "FLOA", vbBinaryCompare) > 0 Or InStr(1, strUpperDecl, "NUM", vbBinaryCompare) > 0 Or InStr(1, strUpperDecl, "DEC", vbBinaryCompare) > 0 Then
        lngAdoType = ADO_DOUBLE
    ElseIf InStr(1, strUpperDecl, "DATE", vbBinaryCompare) > 0 Or InStr(1, strUpperDecl, "TIME", vbBinaryCompare) > 0 Then
        lngAdoType = ADO_DATE
    ElseIf InStr(1, strUpperDecl, "BLOB", vbBinaryCompare) > 0 Or InStr(1, strUpperDecl, "BIN", vbBinaryCompare) > 0 Then
        lngAdoType = ADO_LONGVARBINARY
        lngSize = ADO_MEMO_SIZE
    ElseIf InStr(1, strUpperDecl, "MEMO", vbBinaryCompare) > 0 Or InStr(1, strUpperDecl, "JSON", vbBinaryCompare) > 0 Or InStr(1, strUpperDecl, "CLOB", vbBinaryCompare) > 0 Then
        lngAdoType = ADO_LONGVARWCHAR
        lngSize = ADO_MEMO_SIZE
    Else
        lngAdoType = ADO_VARWCHAR
        lngSize = 255
    End If
End Sub

Private Function CoerceForAdoField(ByVal varVal As Variant, ByVal enmKind As SQLiteColumnKind) As Variant
    Dim dtmParsed As Date
    If IsNull(varVal) Or IsEmpty(varVal) Then
        CoerceForAdoField = Null
        Exit Function
    End If

    Select Case enmKind
        Case sqColBoolean
            CoerceForAdoField = (CDbl(varVal) <> 0#)
        Case sqColCurrency
            CoerceForAdoField = CCur(varVal)
        Case sqColDate, sqColDateTime
            If VarType(varVal) = vbDate Then
                CoerceForAdoField = CDate(varVal)
            ElseIf TryParseIsoDate(CStr(varVal), dtmParsed) Then
                CoerceForAdoField = dtmParsed
            ElseIf IsNumeric(varVal) Then
                CoerceForAdoField = JulianDayToDate(CDbl(varVal))
            Else
                CoerceForAdoField = varVal
            End If
        Case Else
            CoerceForAdoField = varVal
    End Select
End Function

Private Sub RefineAdoFieldFromSample(ByRef lngAdoType As Long, ByRef lngSize As Long, ByVal varSample As Variant)
    Dim intVt As Integer
    intVt = VarType(varSample)

    If lngAdoType = 0 Then
        Select Case intVt
            Case vbBoolean
                lngAdoType = ADO_BOOLEAN
            Case vbByte, vbInteger, vbLong
                lngAdoType = ADO_INTEGER
            Case vbCurrency, vbSingle, vbDouble, vbDecimal
                lngAdoType = ADO_DOUBLE
            Case vbDate
                lngAdoType = ADO_DATE
            Case (vbArray Or vbByte)
                lngAdoType = ADO_LONGVARBINARY
                lngSize = ADO_MEMO_SIZE
            Case Else
                If Len(CStr(varSample)) > 255 Then
                    lngAdoType = ADO_LONGVARWCHAR
                    lngSize = ADO_MEMO_SIZE
                Else
                    lngAdoType = ADO_VARWCHAR
                    lngSize = 255
                End If
        End Select
    ElseIf lngAdoType = ADO_INTEGER And (intVt = vbCurrency Or intVt = vbDouble) Then
        If CDbl(varSample) > 2147483647# Or CDbl(varSample) < -2147483648# Then
            lngAdoType = ADO_DOUBLE
        End If
    ElseIf lngAdoType = ADO_VARWCHAR And intVt = vbString Then
        If Len(CStr(varSample)) > IIf(lngSize > 0, lngSize, 255) Then
            lngAdoType = ADO_LONGVARWCHAR
            lngSize = ADO_MEMO_SIZE
        End If
    End If
End Sub

Private Function MapAdoTypeToSqliteDdl(ByVal lngAdoType As Long) As String
    Select Case lngAdoType
        Case ADO_BOOLEAN
            MapAdoTypeToSqliteDdl = "BOOLEAN"
        Case ADO_TINYINT, ADO_UNSIGNED_TINYINT, ADO_SMALLINT, ADO_INTEGER, ADO_BIGINT
            MapAdoTypeToSqliteDdl = "INTEGER"
        Case ADO_CURRENCY
            MapAdoTypeToSqliteDdl = "CURRENCY"
        Case ADO_SINGLE, ADO_DOUBLE, ADO_DECIMAL, ADO_NUMERIC
            MapAdoTypeToSqliteDdl = "REAL"
        Case ADO_DATE, ADO_DBDATE, ADO_DBTIME, ADO_DBTIMESTAMP
            MapAdoTypeToSqliteDdl = "DATETIME"
        Case ADO_BINARY, ADO_VARBINARY, ADO_LONGVARBINARY
            MapAdoTypeToSqliteDdl = "BLOB"
        Case ADO_LONGVARCHAR, ADO_LONGVARWCHAR
            MapAdoTypeToSqliteDdl = "MEMO"
        Case ADO_GUID
            MapAdoTypeToSqliteDdl = "GUID"
        Case Else
            MapAdoTypeToSqliteDdl = "TEXT"
    End Select
End Function

Private Sub SyncDictIntoCurrentAdoRow(ByVal objAdoRs As Object, ByVal dictReloaded As Scripting.Dictionary)
    Dim objFld As Object
    Dim strFldName As String
    Dim blnChanged As Boolean

    On Error Resume Next
    For Each objFld In objAdoRs.Fields
        strFldName = CStr(objFld.Name)
        If dictReloaded.Exists(strFldName) Then
            AssignAdoFieldValue objFld, CLng(objFld.Type), dictReloaded(strFldName)
            blnChanged = True
        End If
    Next objFld
    If blnChanged Then objAdoRs.Update
    On Error GoTo 0
End Sub

Private Function MakeUniqueFieldName( _
    ByVal strRawName As String, _
    ByVal lngColIdx As Long, _
    ByVal dictUsed As Scripting.Dictionary) As String

    Dim strBase As String
    Dim strCandidate As String
    Dim lngSuffix As Long

    strBase = Trim$(strRawName)
    If Len(strBase) = 0 Then strBase = "Col" & (lngColIdx + 1)
    strCandidate = strBase
    lngSuffix = 2
    Do While dictUsed.Exists(strCandidate)
        strCandidate = strBase & "_" & lngSuffix
        lngSuffix = lngSuffix + 1
    Loop
    dictUsed(strCandidate) = True
    MakeUniqueFieldName = strCandidate
End Function

Private Function IsUnsetKey(ByVal varKeyVal As Variant) As Boolean
    If IsNull(varKeyVal) Or IsEmpty(varKeyVal) Then
        IsUnsetKey = True
    ElseIf IsNumeric(varKeyVal) Then
        IsUnsetKey = (CDbl(varKeyVal) <= 0#)
    ElseIf VarType(varKeyVal) = vbString Then
        IsUnsetKey = (Len(Trim$(CStr(varKeyVal))) = 0)
    End If
End Function

Private Sub EnsureConn()
    If m_Conn Is Nothing Or Not m_Conn.IsOpen Then
        RaiseSQLiteError sqErrConnectionClosed, "clsSQLiteAdoBridge", "An open clsSQLiteConnection is required."
    End If
End Sub

