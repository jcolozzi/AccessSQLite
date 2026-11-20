Attribute VB_Name = "clsSQLiteRepository"
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
' clsSQLiteRepository
' Dual-mode ORM Repository supporting both Scripting.Dictionary rows and
' strongly-typed ISQLiteEntity classes, dirty-field updates, Upsert,
' auto CreatedAt/UpdatedAt timestamps, soft deletes, and N+1-free eager loading.
' =============================================================================

Private Const EAGER_IN_CHUNK_SIZE As Long = 900

Private m_Conn As clsSQLiteConnection
Private m_Map As clsSQLiteTableMap

Public Sub InitRepo(ByVal objConn As clsSQLiteConnection, ByVal objMap As clsSQLiteTableMap)
    If objConn Is Nothing Or objMap Is Nothing Then
        RaiseSQLiteError sqErrInvalidArgument, "clsSQLiteRepository.InitRepo", "Connection and TableMap are required."
    End If
    Set m_Conn = objConn
    Set m_Map = objMap
End Sub

Public Property Get Connection() As clsSQLiteConnection
    Set Connection = m_Conn
End Property

Public Property Get TableMap() As clsSQLiteTableMap
    Set TableMap = m_Map
End Property

Public Property Get TableName() As String
    TableName = m_Map.TableName
End Property

Public Property Get PrimaryKeyColumn() As String
    PrimaryKeyColumn = m_Map.PrimaryKeyColumn
End Property

Public Function NewQuery(Optional ByVal strAlias As String = vbNullString) As clsSQLiteQueryBuilder
    Dim objQb As clsSQLiteQueryBuilder
    Set objQb = New clsSQLiteQueryBuilder
    objQb.InitBuilder m_Conn, m_Map.TableName, strAlias, m_Map
    Set NewQuery = objQb
End Function

' =============================================================================
' Dictionary CRUD Operations
' =============================================================================

Public Function FindById( _
    ByVal varId As Variant, _
    Optional ByVal blnIncludeTrashed As Boolean = False) As Scripting.Dictionary

    Dim objQb As clsSQLiteQueryBuilder
    Set objQb = NewQuery().WhereEq(m_Map.PrimaryKeyColumn, varId)
    If blnIncludeTrashed Then objQb.WithTrashed
    Set FindById = objQb.FirstRow()
End Function

Public Function FindByIdRaw( _
    ByVal varId As Variant, _
    Optional ByVal blnIncludeTrashed As Boolean = False) As Scripting.Dictionary

    Dim objQb As clsSQLiteQueryBuilder
    Set objQb = NewQuery().WhereEq(m_Map.PrimaryKeyColumn, varId)
    If blnIncludeTrashed Then objQb.WithTrashed
    Set FindByIdRaw = objQb.FirstRowRaw()
End Function

Public Function GetOrFail(ByVal varId As Variant) As Scripting.Dictionary
    Dim dictRow As Scripting.Dictionary
    Set dictRow = FindByIdRaw(varId, False)
    If dictRow Is Nothing Then
        RaiseSQLiteError sqErrNotFound, "clsSQLiteRepository.GetOrFail", _
            m_Map.TableName & " record with " & m_Map.PrimaryKeyColumn & " = " & CStr(Nz(varId, "NULL")) & " was not found."
    End If
    Set GetOrFail = dictRow
End Function

Public Function ExistsById( _
    ByVal varId As Variant, _
    Optional ByVal blnIncludeTrashed As Boolean = False) As Boolean

    Dim objQb As clsSQLiteQueryBuilder
    If IsNull(varId) Or IsEmpty(varId) Then Exit Function
    Set objQb = NewQuery().WhereEq(m_Map.PrimaryKeyColumn, varId)
    If blnIncludeTrashed Then objQb.WithTrashed
    ExistsById = objQb.ExistsAny()
End Function

Public Function CountAll(Optional ByVal blnIncludeTrashed As Boolean = False) As Long
    Dim objQb As clsSQLiteQueryBuilder
    Set objQb = NewQuery()
    If blnIncludeTrashed Then objQb.WithTrashed
    CountAll = objQb.CountRows()
End Function

Public Function ListAll( _
    Optional ByVal strOrderByCol As String = vbNullString, _
    Optional ByVal blnDescending As Boolean = False) As Collection

    Dim objQb As clsSQLiteQueryBuilder
    Set objQb = NewQuery()
    If Len(Trim$(strOrderByCol)) > 0 Then
        objQb.OrderByCol strOrderByCol, blnDescending
    ElseIf Len(m_Map.PrimaryKeyColumn) > 0 And m_Map.HasColumn(m_Map.PrimaryKeyColumn) Then
        objQb.OrderByCol m_Map.PrimaryKeyColumn, False
    End If
    Set ListAll = objQb.ToRows()
End Function

Public Function ListAllRaw( _
    Optional ByVal strOrderByCol As String = vbNullString, _
    Optional ByVal blnDescending As Boolean = False) As Collection

    Dim objQb As clsSQLiteQueryBuilder
    Set objQb = NewQuery()
    If Len(Trim$(strOrderByCol)) > 0 Then
        objQb.OrderByCol strOrderByCol, blnDescending
    ElseIf Len(m_Map.PrimaryKeyColumn) > 0 And m_Map.HasColumn(m_Map.PrimaryKeyColumn) Then
        objQb.OrderByCol m_Map.PrimaryKeyColumn, False
    End If
    Set ListAllRaw = objQb.ToRowsRaw()
End Function

' Inserts a new row and returns its primary key (generated ROWID/AutoIncrement or explicit PK).
Public Function InsertRecord( _
    ByVal dictValues As Scripting.Dictionary, _
    Optional ByVal blnEnforceClientEditable As Boolean = False) As Variant

    Dim dictWork As Scripting.Dictionary
    Dim objPkCol As clsSQLiteColumnDef
    Dim dictClean As Scripting.Dictionary
    Dim strNowIso As String
    Dim strColsSql As String
    Dim strPlaceholders As String
    Dim colParams As Collection
    Dim varKey As Variant
    Dim blnExplicitPk As Boolean

    Set dictWork = CopyDict(dictValues)
    Set objPkCol = m_Map.GetColumn(m_Map.PrimaryKeyColumn)

    ' Strip empty/zero auto-increment PK before validation.
    If Not objPkCol Is Nothing Then
        If objPkCol.IsAutoIncrement And dictWork.Exists(objPkCol.ColumnName) Then
            If IsUnsetAutoKey(dictWork(objPkCol.ColumnName)) Then
                dictWork.Remove objPkCol.ColumnName
            Else
                blnExplicitPk = True
            End If
        ElseIf dictWork.Exists(objPkCol.ColumnName) Then
            If Not IsNull(dictWork(objPkCol.ColumnName)) And Not IsEmpty(dictWork(objPkCol.ColumnName)) Then
                blnExplicitPk = True
            End If
        End If
    End If

    Set dictClean = m_Map.ValidateAndCoerce(dictWork, True, blnEnforceClientEditable)

    strNowIso = UtcNowIsoString()
    If Len(m_Map.CreatedAtColumn) > 0 Then
        If Not dictClean.Exists(m_Map.CreatedAtColumn) Or IsNull(dictClean(m_Map.CreatedAtColumn)) Then
            dictClean(m_Map.CreatedAtColumn) = m_Map.CoerceColumnValue(m_Map.CreatedAtColumn, strNowIso)
        End If
    End If
    If Len(m_Map.UpdatedAtColumn) > 0 Then
        If Not dictClean.Exists(m_Map.UpdatedAtColumn) Or IsNull(dictClean(m_Map.UpdatedAtColumn)) Then
            dictClean(m_Map.UpdatedAtColumn) = m_Map.CoerceColumnValue(m_Map.UpdatedAtColumn, strNowIso)
        End If
    End If

    If dictClean.Count = 0 Then
        m_Conn.ExecuteParams "INSERT INTO " & QuoteSQLiteIdent(m_Map.TableName) & " DEFAULT VALUES;", Empty
    Else
        Set colParams = New Collection
        For Each varKey In dictClean.Keys
            If Len(strColsSql) > 0 Then
                strColsSql = strColsSql & ", "
                strPlaceholders = strPlaceholders & ", "
            End If
            strColsSql = strColsSql & QuoteSQLiteIdent(CStr(varKey))
            strPlaceholders = strPlaceholders & "?"
            colParams.Add dictClean(varKey)
        Next varKey

        m_Conn.ExecuteParams _
            "INSERT INTO " & QuoteSQLiteIdent(m_Map.TableName) & _
            " (" & strColsSql & ") VALUES (" & strPlaceholders & ");", _
            colParams
    End If

    If blnExplicitPk And dictClean.Exists(m_Map.PrimaryKeyColumn) Then
        InsertRecord = dictClean(m_Map.PrimaryKeyColumn)
    Else
        InsertRecord = m_Conn.LastInsertRowId
    End If
End Function

' Bulk-inserts a Collection of Scripting.Dictionary records atomically inside a transaction/savepoint.
Public Function InsertMany(ByVal colRecords As Collection) As Long
    Dim dictItem As Scripting.Dictionary
    Dim lngCount As Long

    If colRecords Is Nothing Or colRecords.Count = 0 Then Exit Function

    m_Conn.BeginTrans
    On Error GoTo BatchFail
    For Each dictItem In colRecords
        InsertRecord dictItem, False
        lngCount = lngCount + 1
    Next dictItem
    m_Conn.CommitTrans
    InsertMany = lngCount
    Exit Function

BatchFail:
    Dim lngErrNum As Long
    Dim strErrSrc As String
    Dim strErrDesc As String
    lngErrNum = Err.Number
    strErrSrc = Err.Source
    strErrDesc = Err.Description
    m_Conn.RollbackTrans
    Err.Raise lngErrNum, strErrSrc, strErrDesc
End Function

' Updates a record by primary key; when blnDirtyOnly=True, diffs against stored values first.
Public Function UpdateRecord( _
    ByVal varId As Variant, _
    ByVal dictValues As Scripting.Dictionary, _
    Optional ByVal blnDirtyOnly As Boolean = False, _
    Optional ByVal blnEnforceClientEditable As Boolean = False) As Long

    Dim dictWork As Scripting.Dictionary
    Dim dictClean As Scripting.Dictionary
    Dim dictStored As Scripting.Dictionary
    Dim dictEffective As Scripting.Dictionary
    Dim strSetSql As String
    Dim colParams As Collection
    Dim varKey As Variant

    If dictValues Is Nothing Or dictValues.Count = 0 Then Exit Function

    Set dictWork = CopyDict(dictValues)
    If dictWork.Exists(m_Map.PrimaryKeyColumn) Then
        dictWork.Remove m_Map.PrimaryKeyColumn
    End If
    If dictWork.Count = 0 Then Exit Function

    Set dictClean = m_Map.ValidateAndCoerce(dictWork, False, blnEnforceClientEditable)

    If blnDirtyOnly Then
        Set dictStored = FindByIdRaw(varId, True)
        If dictStored Is Nothing Then
            RaiseSQLiteError sqErrNotFound, "clsSQLiteRepository.UpdateRecord", _
                m_Map.TableName & " record with " & m_Map.PrimaryKeyColumn & " = " & CStr(Nz(varId, "NULL")) & " was not found."
        End If
        Set dictEffective = m_Map.DiffChangedColumns(dictClean, dictStored)
        If dictEffective.Count = 0 Then Exit Function
    Else
        Set dictEffective = dictClean
    End If

    If Len(m_Map.UpdatedAtColumn) > 0 And Not dictEffective.Exists(m_Map.UpdatedAtColumn) Then
        dictEffective(m_Map.UpdatedAtColumn) = m_Map.CoerceColumnValue(m_Map.UpdatedAtColumn, UtcNowIsoString())
    End If

    Set colParams = New Collection
    For Each varKey In dictEffective.Keys
        If Len(strSetSql) > 0 Then strSetSql = strSetSql & ", "
        strSetSql = strSetSql & QuoteSQLiteIdent(CStr(varKey)) & " = ?"
        colParams.Add dictEffective(varKey)
    Next varKey
    colParams.Add m_Map.CoerceColumnValue(m_Map.PrimaryKeyColumn, varId)

    UpdateRecord = m_Conn.ExecuteParams( _
        "UPDATE " & QuoteSQLiteIdent(m_Map.TableName) & _
        " SET " & strSetSql & _
        " WHERE " & QuoteSQLiteIdent(m_Map.PrimaryKeyColumn) & " = ?;", _
        colParams)
End Function

' Atomic SQLite UPSERT: INSERT ... ON CONFLICT (cols) DO UPDATE SET ...
Public Function UpsertRecord( _
    ByVal strConflictColumnsCsv As String, _
    ByVal dictValues As Scripting.Dictionary) As Variant

    Dim dictWork As Scripting.Dictionary
    Dim objPkCol As clsSQLiteColumnDef
    Dim dictClean As Scripting.Dictionary
    Dim astrConflict() As String
    Dim dictConflictSet As Scripting.Dictionary
    Dim strConflictSql As String
    Dim strColsSql As String
    Dim strPlaceholders As String
    Dim strUpdateSql As String
    Dim colParams As Collection
    Dim lngIdx As Long
    Dim strCCol As String
    Dim varKey As Variant
    Dim strNowIso As String

    Set dictWork = CopyDict(dictValues)
    Set objPkCol = m_Map.GetColumn(m_Map.PrimaryKeyColumn)
    If Not objPkCol Is Nothing Then
        If objPkCol.IsAutoIncrement And dictWork.Exists(objPkCol.ColumnName) Then
            If IsUnsetAutoKey(dictWork(objPkCol.ColumnName)) Then
                dictWork.Remove objPkCol.ColumnName
            End If
        End If
    End If

    Set dictClean = m_Map.ValidateAndCoerce(dictWork, True, False)

    strNowIso = UtcNowIsoString()
    If Len(m_Map.CreatedAtColumn) > 0 And Not dictClean.Exists(m_Map.CreatedAtColumn) Then
        dictClean(m_Map.CreatedAtColumn) = m_Map.CoerceColumnValue(m_Map.CreatedAtColumn, strNowIso)
    End If
    If Len(m_Map.UpdatedAtColumn) > 0 And Not dictClean.Exists(m_Map.UpdatedAtColumn) Then
        dictClean(m_Map.UpdatedAtColumn) = m_Map.CoerceColumnValue(m_Map.UpdatedAtColumn, strNowIso)
    End If

    Set dictConflictSet = NewTextDict()
    astrConflict = Split(strConflictColumnsCsv, ",")
    For lngIdx = LBound(astrConflict) To UBound(astrConflict)
        strCCol = m_Map.CanonicalColumn(Trim$(astrConflict(lngIdx)))
        dictConflictSet(strCCol) = True
        If Len(strConflictSql) > 0 Then strConflictSql = strConflictSql & ", "
        strConflictSql = strConflictSql & QuoteSQLiteIdent(strCCol)
    Next lngIdx

    Set colParams = New Collection
    For Each varKey In dictClean.Keys
        If Len(strColsSql) > 0 Then
            strColsSql = strColsSql & ", "
            strPlaceholders = strPlaceholders & ", "
        End If
        strColsSql = strColsSql & QuoteSQLiteIdent(CStr(varKey))
        strPlaceholders = strPlaceholders & "?"
        colParams.Add dictClean(varKey)

        If Not dictConflictSet.Exists(CStr(varKey)) And StrComp(CStr(varKey), m_Map.CreatedAtColumn, vbTextCompare) <> 0 Then
            If Len(strUpdateSql) > 0 Then strUpdateSql = strUpdateSql & ", "
            strUpdateSql = strUpdateSql & QuoteSQLiteIdent(CStr(varKey)) & " = excluded." & QuoteSQLiteIdent(CStr(varKey))
        End If
    Next varKey

    If Len(strUpdateSql) = 0 Then
        m_Conn.ExecuteParams _
            "INSERT INTO " & QuoteSQLiteIdent(m_Map.TableName) & _
            " (" & strColsSql & ") VALUES (" & strPlaceholders & ") " & _
            "ON CONFLICT (" & strConflictSql & ") DO NOTHING;", _
            colParams
    Else
        m_Conn.ExecuteParams _
            "INSERT INTO " & QuoteSQLiteIdent(m_Map.TableName) & _
            " (" & strColsSql & ") VALUES (" & strPlaceholders & ") " & _
            "ON CONFLICT (" & strConflictSql & ") DO UPDATE SET " & strUpdateSql & ";", _
            colParams
    End If

    If dictClean.Exists(m_Map.PrimaryKeyColumn) Then
        UpsertRecord = dictClean(m_Map.PrimaryKeyColumn)
    Else
        UpsertRecord = m_Conn.LastInsertRowId
    End If
End Function

Public Function DeleteRecord( _
    ByVal varId As Variant, _
    Optional ByVal blnForceHardDelete As Boolean = False) As Long

    DeleteRecord = NewQuery().WithTrashed().WhereEq(m_Map.PrimaryKeyColumn, varId).DeleteMatching(blnForceHardDelete)
End Function

Public Function RestoreRecord(ByVal varId As Variant) As Long
    Dim dictRestore As Scripting.Dictionary
    If Not m_Map.HasSoftDelete Then Exit Function
    Set dictRestore = NewTextDict()
    dictRestore(m_Map.SoftDeleteColumn) = Null
    RestoreRecord = UpdateRecord(varId, dictRestore, False, False)
End Function

' =============================================================================
' Strongly-Typed Entity CRUD (ISQLiteEntity)
' =============================================================================

Public Function FindEntity( _
    ByVal varId As Variant, _
    ByVal objPrototype As ISQLiteEntity, _
    Optional ByVal blnIncludeTrashed As Boolean = False) As ISQLiteEntity

    Dim objQb As clsSQLiteQueryBuilder
    Set objQb = NewQuery().WhereEq(m_Map.PrimaryKeyColumn, varId)
    If blnIncludeTrashed Then objQb.WithTrashed
    Set FindEntity = objQb.FirstEntity(objPrototype)
End Function

Public Function ListEntities( _
    ByVal objPrototype As ISQLiteEntity, _
    Optional ByVal strOrderByCol As String = vbNullString, _
    Optional ByVal blnDescending As Boolean = False) As Collection

    Dim objQb As clsSQLiteQueryBuilder
    Set objQb = NewQuery()
    If Len(Trim$(strOrderByCol)) > 0 Then
        objQb.OrderByCol strOrderByCol, blnDescending
    ElseIf Len(m_Map.PrimaryKeyColumn) > 0 And m_Map.HasColumn(m_Map.PrimaryKeyColumn) Then
        objQb.OrderByCol m_Map.PrimaryKeyColumn, False
    End If
    Set ListEntities = objQb.ToEntities(objPrototype)
End Function

Public Function SaveEntity( _
    ByVal objEntity As ISQLiteEntity, _
    Optional ByVal blnDirtyOnly As Boolean = True) As Variant

    Dim dictPayload As Scripting.Dictionary
    Dim objPkCol As clsSQLiteColumnDef
    Dim varPk As Variant
    Dim blnIsInsert As Boolean
    Dim varSavedId As Variant
    Dim lngRowsChanged As Long
    Dim dictReloaded As Scripting.Dictionary

    If objEntity Is Nothing Then
        RaiseSQLiteError sqErrInvalidArgument, "clsSQLiteRepository.SaveEntity", "Entity instance cannot be Nothing."
    End If

    Set dictPayload = NewTextDict()
    objEntity.WriteToRow dictPayload

    Set objPkCol = m_Map.GetColumn(m_Map.PrimaryKeyColumn)
    If dictPayload.Exists(m_Map.PrimaryKeyColumn) Then
        varPk = dictPayload(m_Map.PrimaryKeyColumn)
    Else
        varPk = Empty
    End If

    If Not objPkCol Is Nothing And objPkCol.IsAutoIncrement Then
        blnIsInsert = IsUnsetAutoKey(varPk)
    ElseIf IsNull(varPk) Or IsEmpty(varPk) Or Len(Trim$(CStr(varPk))) = 0 Then
        blnIsInsert = True
    Else
        blnIsInsert = Not ExistsById(varPk, True)
    End If

    If blnIsInsert Then
        varSavedId = InsertRecord(dictPayload, False)
        lngRowsChanged = 1
    Else
        lngRowsChanged = UpdateRecord(varPk, dictPayload, blnDirtyOnly, False)
        varSavedId = varPk
    End If

    If lngRowsChanged > 0 Then
        Set dictReloaded = FindByIdRaw(varSavedId, True)
        If Not dictReloaded Is Nothing Then
            objEntity.ReadFromRow dictReloaded
        End If
    End If
    SaveEntity = varSavedId
End Function

Public Function DeleteEntity( _
    ByVal objEntity As ISQLiteEntity, _
    Optional ByVal blnForceHardDelete As Boolean = False) As Long

    Dim dictPayload As Scripting.Dictionary
    If objEntity Is Nothing Then Exit Function
    Set dictPayload = NewTextDict()
    objEntity.WriteToRow dictPayload
    If Not dictPayload.Exists(m_Map.PrimaryKeyColumn) Then
        RaiseSQLiteError sqErrInvalidArgument, "clsSQLiteRepository.DeleteEntity", _
            "Entity did not write primary key '" & m_Map.PrimaryKeyColumn & "'."
    End If
    DeleteEntity = DeleteRecord(dictPayload(m_Map.PrimaryKeyColumn), blnForceHardDelete)
End Function

' =============================================================================
' Batch Eager Relationship Loading (Eliminates N+1 Queries)
' =============================================================================

Public Sub IncludeHasMany( _
    ByVal colParentRows As Collection, _
    ByVal strChildTable As String, _
    ByVal strForeignKeyCol As String, _
    Optional ByVal strAttachAs As String = vbNullString, _
    Optional ByVal strLocalKeyCol As String = vbNullString, _
    Optional ByVal strOrderByChildCol As String = vbNullString)

    Dim strLocalCol As String
    Dim strPropKey As String
    Dim dictBuckets As Scripting.Dictionary
    Dim colDistinctIds As Collection
    Dim dictSeen As Scripting.Dictionary
    Dim dictParent As Scripting.Dictionary
    Dim varIdVal As Variant
    Dim strIdKey As String
    Dim colChunk As Collection
    Dim lngIdx As Long
    Dim objChildQb As clsSQLiteQueryBuilder
    Dim colChildren As Collection
    Dim dictChild As Scripting.Dictionary
    Dim colBucket As Collection

    If colParentRows Is Nothing Or colParentRows.Count = 0 Then Exit Sub

    strLocalCol = IIf(Len(Trim$(strLocalKeyCol)) > 0, Trim$(strLocalKeyCol), m_Map.PrimaryKeyColumn)
    strPropKey = IIf(Len(Trim$(strAttachAs)) > 0, Trim$(strAttachAs), Trim$(strChildTable))

    Set dictBuckets = NewTextDict()
    Set dictSeen = NewTextDict()
    Set colDistinctIds = New Collection

    For Each dictParent In colParentRows
        Set dictParent(strPropKey) = New Collection
        If dictParent.Exists(strLocalCol) Then
            varIdVal = dictParent(strLocalCol)
            If Not IsNull(varIdVal) And Not IsEmpty(varIdVal) Then
                strIdKey = CStr(varIdVal)
                If Not dictSeen.Exists(strIdKey) Then
                    dictSeen(strIdKey) = True
                    colDistinctIds.Add varIdVal
                    Set dictBuckets(strIdKey) = New Collection
                End If
            End If
        End If
    Next dictParent

    If colDistinctIds.Count = 0 Then Exit Sub

    Set colChunk = New Collection
    For lngIdx = 1 To colDistinctIds.Count
        colChunk.Add colDistinctIds(lngIdx)
        If colChunk.Count >= EAGER_IN_CHUNK_SIZE Or lngIdx = colDistinctIds.Count Then
            Set objChildQb = m_Conn.FromTable(strChildTable).WhereIn(strForeignKeyCol, colChunk)
            If Len(Trim$(strOrderByChildCol)) > 0 Then
                objChildQb.OrderByCol strOrderByChildCol, False
            End If
            Set colChildren = objChildQb.ToRowsRaw()
            For Each dictChild In colChildren
                If dictChild.Exists(strForeignKeyCol) And Not IsNull(dictChild(strForeignKeyCol)) Then
                    strIdKey = CStr(dictChild(strForeignKeyCol))
                    If dictBuckets.Exists(strIdKey) Then
                        Set colBucket = dictBuckets(strIdKey)
                        colBucket.Add dictChild
                    End If
                End If
            Next dictChild
            Set colChunk = New Collection
        End If
    Next lngIdx

    For Each dictParent In colParentRows
        If dictParent.Exists(strLocalCol) And Not IsNull(dictParent(strLocalCol)) Then
            strIdKey = CStr(dictParent(strLocalCol))
            If dictBuckets.Exists(strIdKey) Then
                Set dictParent(strPropKey) = dictBuckets(strIdKey)
            End If
        End If
    Next dictParent
End Sub

Public Sub IncludeBelongsTo( _
    ByVal colChildRows As Collection, _
    ByVal strParentTable As String, _
    ByVal strForeignKeyCol As String, _
    Optional ByVal strAttachAs As String = vbNullString, _
    Optional ByVal strParentKeyCol As String = vbNullString)

    Dim objParentRepo As clsSQLiteRepository
    Dim strTargetPk As String
    Dim strPropKey As String
    Dim dictSeen As Scripting.Dictionary
    Dim colDistinctIds As Collection
    Dim dictLookup As Scripting.Dictionary
    Dim dictChild As Scripting.Dictionary
    Dim varFk As Variant
    Dim strFkKey As String
    Dim colChunk As Collection
    Dim lngIdx As Long
    Dim colParents As Collection
    Dim dictParent As Scripting.Dictionary

    If colChildRows Is Nothing Or colChildRows.Count = 0 Then Exit Sub

    Set objParentRepo = m_Conn.Repository(strParentTable)
    strTargetPk = IIf(Len(Trim$(strParentKeyCol)) > 0, Trim$(strParentKeyCol), objParentRepo.PrimaryKeyColumn)
    strPropKey = IIf(Len(Trim$(strAttachAs)) > 0, Trim$(strAttachAs), Trim$(strParentTable))

    Set dictSeen = NewTextDict()
    Set colDistinctIds = New Collection
    Set dictLookup = NewTextDict()

    For Each dictChild In colChildRows
        Set dictChild(strPropKey) = Nothing
        If dictChild.Exists(strForeignKeyCol) Then
            varFk = dictChild(strForeignKeyCol)
            If Not IsNull(varFk) And Not IsEmpty(varFk) Then
                strFkKey = CStr(varFk)
                If Not dictSeen.Exists(strFkKey) Then
                    dictSeen(strFkKey) = True
                    colDistinctIds.Add varFk
                End If
            End If
        End If
    Next dictChild

    If colDistinctIds.Count = 0 Then Exit Sub

    Set colChunk = New Collection
    For lngIdx = 1 To colDistinctIds.Count
        colChunk.Add colDistinctIds(lngIdx)
        If colChunk.Count >= EAGER_IN_CHUNK_SIZE Or lngIdx = colDistinctIds.Count Then
            Set colParents = objParentRepo.NewQuery().WhereIn(strTargetPk, colChunk).ToRowsRaw()
            For Each dictParent In colParents
                If dictParent.Exists(strTargetPk) And Not IsNull(dictParent(strTargetPk)) Then
                    Set dictLookup(CStr(dictParent(strTargetPk))) = dictParent
                End If
            Next dictParent
            Set colChunk = New Collection
        End If
    Next lngIdx

    For Each dictChild In colChildRows
        If dictChild.Exists(strForeignKeyCol) And Not IsNull(dictChild(strForeignKeyCol)) Then
            strFkKey = CStr(dictChild(strForeignKeyCol))
            If dictLookup.Exists(strFkKey) Then
                Set dictChild(strPropKey) = dictLookup(strFkKey)
            End If
        End If
    Next dictChild
End Sub

' =============================================================================
' Internal Helpers
' =============================================================================

Private Function CopyDict(ByVal dictIn As Scripting.Dictionary) As Scripting.Dictionary
    Dim dictOut As Scripting.Dictionary
    Dim varKey As Variant
    Set dictOut = NewTextDict()
    If Not dictIn Is Nothing Then
        For Each varKey In dictIn.Keys
            dictOut(CStr(varKey)) = dictIn(varKey)
        Next varKey
    End If
    Set CopyDict = dictOut
End Function

Private Function IsUnsetAutoKey(ByVal varKeyVal As Variant) As Boolean
    If IsNull(varKeyVal) Or IsEmpty(varKeyVal) Then
        IsUnsetAutoKey = True
    ElseIf IsNumeric(varKeyVal) Then
        IsUnsetAutoKey = (CDbl(varKeyVal) <= 0#)
    ElseIf VarType(varKeyVal) = vbString Then
        IsUnsetAutoKey = (Len(Trim$(CStr(varKeyVal))) = 0)
    End If
End Function

