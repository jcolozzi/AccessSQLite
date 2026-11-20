Attribute VB_Name = "clsSQLiteQueryBuilder"
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
' clsSQLiteQueryBuilder
' Fluent, injection-safe parameterized SQL query builder for SQLite.
' Supports SELECT, JOIN, WHERE (Eq/Op/In/Between/Like/Contains/Null/Raw),
' soft-delete scoping, GROUP BY/HAVING, ORDER BY, LIMIT/OFFSET/Page,
' aggregates, bulk UPDATE/DELETE, and Dictionary / ISQLiteEntity hydration.
' =============================================================================

Private m_Conn As clsSQLiteConnection
Private m_TableMap As clsSQLiteTableMap
Private m_TableName As String
Private m_TableAlias As String
Private m_Distinct As Boolean

Private m_SelectItems As Collection
Private m_JoinClauses As Collection
Private m_WhereClauses As Collection
Private m_WhereParams As Collection
Private m_GroupByItems As Collection
Private m_HavingClauses As Collection
Private m_HavingParams As Collection
Private m_OrderByItems As Collection

Private m_LimitCount As Long
Private m_OffsetCount As Long
Private m_SoftDeleteMode As Long ' 0 = exclude deleted, 1 = with trashed, 2 = only trashed

Private Sub Class_Initialize()
    ResetClauses
End Sub

Public Sub InitBuilder( _
    ByVal objConn As clsSQLiteConnection, _
    ByVal strTableName As String, _
    Optional ByVal strAlias As String = vbNullString, _
    Optional ByVal objMap As clsSQLiteTableMap = Nothing)

    Set m_Conn = objConn
    m_TableName = Trim$(strTableName)
    m_TableAlias = Trim$(strAlias)
    Set m_TableMap = objMap
    ResetClauses
End Sub

Private Sub ResetClauses()
    m_Distinct = False
    Set m_SelectItems = New Collection
    Set m_JoinClauses = New Collection
    Set m_WhereClauses = New Collection
    Set m_WhereParams = New Collection
    Set m_GroupByItems = New Collection
    Set m_HavingClauses = New Collection
    Set m_HavingParams = New Collection
    Set m_OrderByItems = New Collection
    m_LimitCount = -1
    m_OffsetCount = -1
    m_SoftDeleteMode = 0
End Sub

' =============================================================================
' Projections & Joins
' =============================================================================

Public Function SelectCols(ParamArray varCols() As Variant) As clsSQLiteQueryBuilder
    Dim lngIdx As Long
    Dim astrParts() As String
    Dim lngSub As Long
    Dim strItem As String

    For lngIdx = LBound(varCols) To UBound(varCols)
        astrParts = Split(CStr(varCols(lngIdx)), ",")
        For lngSub = LBound(astrParts) To UBound(astrParts)
            strItem = Trim$(astrParts(lngSub))
            If Len(strItem) > 0 Then
                m_SelectItems.Add QuoteSQLiteIdent(strItem)
            End If
        Next lngSub
    Next lngIdx
    Set SelectCols = Me
End Function

Public Function SelectAliased(ByVal strCol As String, ByVal strAsAlias As String) As clsSQLiteQueryBuilder
    m_SelectItems.Add QuoteSQLiteIdent(strCol) & " AS " & QuoteSQLiteIdent(strAsAlias)
    Set SelectAliased = Me
End Function

Public Function SelectRaw(ByVal strSqlExpr As String) As clsSQLiteQueryBuilder
    m_SelectItems.Add Trim$(strSqlExpr)
    Set SelectRaw = Me
End Function

Public Function DistinctRows(Optional ByVal blnDistinct As Boolean = True) As clsSQLiteQueryBuilder
    m_Distinct = blnDistinct
    Set DistinctRows = Me
End Function

Public Function InnerJoinTable( _
    ByVal strJoinTable As String, _
    ByVal strLeftCol As String, _
    ByVal strRightCol As String, _
    Optional ByVal strJoinAlias As String = vbNullString) As clsSQLiteQueryBuilder

    AddTypedJoin "INNER JOIN", strJoinTable, strLeftCol, strRightCol, strJoinAlias
    Set InnerJoinTable = Me
End Function

Public Function LeftJoinTable( _
    ByVal strJoinTable As String, _
    ByVal strLeftCol As String, _
    ByVal strRightCol As String, _
    Optional ByVal strJoinAlias As String = vbNullString) As clsSQLiteQueryBuilder

    AddTypedJoin "LEFT JOIN", strJoinTable, strLeftCol, strRightCol, strJoinAlias
    Set LeftJoinTable = Me
End Function

Public Function JoinRaw(ByVal strJoinClause As String) As clsSQLiteQueryBuilder
    m_JoinClauses.Add Trim$(strJoinClause)
    Set JoinRaw = Me
End Function

Private Sub AddTypedJoin( _
    ByVal strJoinKeyword As String, _
    ByVal strJoinTable As String, _
    ByVal strLeftCol As String, _
    ByVal strRightCol As String, _
    ByVal strJoinAlias As String)

    Dim strClause As String
    strClause = strJoinKeyword & " " & QuoteSQLiteIdent(strJoinTable)
    If Len(Trim$(strJoinAlias)) > 0 Then
        strClause = strClause & " AS " & QuoteSQLiteIdent(strJoinAlias)
    End If
    strClause = strClause & " ON " & QuoteSQLiteIdent(strLeftCol) & " = " & QuoteSQLiteIdent(strRightCol)
    m_JoinClauses.Add strClause
End Sub

' =============================================================================
' WHERE Filtering (100% Bound Parameters)
' =============================================================================

Public Function WhereEq(ByVal strCol As String, ByVal varVal As Variant) As clsSQLiteQueryBuilder
    If IsNull(varVal) Or IsEmpty(varVal) Then
        m_WhereClauses.Add QuoteSQLiteIdent(strCol) & " IS NULL"
    Else
        m_WhereClauses.Add QuoteSQLiteIdent(strCol) & " = ?"
        m_WhereParams.Add varVal
    End If
    Set WhereEq = Me
End Function

Public Function WhereNotEq(ByVal strCol As String, ByVal varVal As Variant) As clsSQLiteQueryBuilder
    If IsNull(varVal) Or IsEmpty(varVal) Then
        m_WhereClauses.Add QuoteSQLiteIdent(strCol) & " IS NOT NULL"
    Else
        m_WhereClauses.Add QuoteSQLiteIdent(strCol) & " <> ?"
        m_WhereParams.Add varVal
    End If
    Set WhereNotEq = Me
End Function

Public Function WhereOp( _
    ByVal strCol As String, _
    ByVal strOperator As String, _
    ByVal varVal As Variant) As clsSQLiteQueryBuilder

    Dim strOp As String
    strOp = UCase$(Trim$(strOperator))
    Select Case strOp
        Case "=", "<>", "!=", "<", "<=", ">", ">=", "LIKE", "NOT LIKE", "GLOB"
            m_WhereClauses.Add QuoteSQLiteIdent(strCol) & " " & strOp & " ?"
            m_WhereParams.Add varVal
        Case Else
            RaiseSQLiteError sqErrInvalidArgument, "clsSQLiteQueryBuilder.WhereOp", _
                "Unsupported SQL comparison operator: " & strOperator
    End Select
    Set WhereOp = Me
End Function

Public Function WhereIn(ByVal strCol As String, ByVal varValues As Variant) As clsSQLiteQueryBuilder
    AddInPredicate strCol, varValues, False
    Set WhereIn = Me
End Function

Public Function WhereNotIn(ByVal strCol As String, ByVal varValues As Variant) As clsSQLiteQueryBuilder
    AddInPredicate strCol, varValues, True
    Set WhereNotIn = Me
End Function

Private Sub AddInPredicate(ByVal strCol As String, ByVal varValues As Variant, ByVal blnNegate As Boolean)
    Dim colItems As Collection
    Dim strPlaceholders As String
    Dim lngIdx As Long
    Dim lngLow As Long
    Dim lngHigh As Long

    Set colItems = NormalizeListValues(varValues)
    If colItems.Count = 0 Then
        m_WhereClauses.Add IIf(blnNegate, "1 = 1", "1 = 0")
        Exit Sub
    End If

    For lngIdx = 1 To colItems.Count
        If lngIdx > 1 Then strPlaceholders = strPlaceholders & ", "
        strPlaceholders = strPlaceholders & "?"
        m_WhereParams.Add colItems(lngIdx)
    Next lngIdx

    m_WhereClauses.Add QuoteSQLiteIdent(strCol) & IIf(blnNegate, " NOT IN (", " IN (") & strPlaceholders & ")"
    Exit Sub
    lngLow = lngHigh
End Sub

Public Function WhereBetween( _
    ByVal strCol As String, _
    ByVal varLow As Variant, _
    ByVal varHigh As Variant) As clsSQLiteQueryBuilder

    m_WhereClauses.Add QuoteSQLiteIdent(strCol) & " BETWEEN ? AND ?"
    m_WhereParams.Add varLow
    m_WhereParams.Add varHigh
    Set WhereBetween = Me
End Function

Public Function WhereLike(ByVal strCol As String, ByVal strPattern As String) As clsSQLiteQueryBuilder
    m_WhereClauses.Add QuoteSQLiteIdent(strCol) & " LIKE ?"
    m_WhereParams.Add strPattern
    Set WhereLike = Me
End Function

' Escapes %, _, and \ in user search text and matches substring case-insensitively.
Public Function WhereContains(ByVal strCol As String, ByVal strSearchText As String) As clsSQLiteQueryBuilder
    Dim strEscaped As String
    strEscaped = Replace(strSearchText, "\", "\\")
    strEscaped = Replace(strEscaped, "%", "\%")
    strEscaped = Replace(strEscaped, "_", "\_")
    m_WhereClauses.Add QuoteSQLiteIdent(strCol) & " LIKE ? ESCAPE '\'"
    m_WhereParams.Add "%" & strEscaped & "%"
    Set WhereContains = Me
End Function

Public Function WhereNull(ByVal strCol As String) As clsSQLiteQueryBuilder
    m_WhereClauses.Add QuoteSQLiteIdent(strCol) & " IS NULL"
    Set WhereNull = Me
End Function

Public Function WhereNotNull(ByVal strCol As String) As clsSQLiteQueryBuilder
    m_WhereClauses.Add QuoteSQLiteIdent(strCol) & " IS NOT NULL"
    Set WhereNotNull = Me
End Function

Public Function WhereRaw(ByVal strPredicateSql As String, ParamArray varArgs() As Variant) As clsSQLiteQueryBuilder
    Dim lngIdx As Long
    m_WhereClauses.Add "(" & Trim$(strPredicateSql) & ")"
    If Not IsMissing(varArgs) Then
        For lngIdx = LBound(varArgs) To UBound(varArgs)
            m_WhereParams.Add varArgs(lngIdx)
        Next lngIdx
    End If
    Set WhereRaw = Me
End Function

Public Function WithTrashed() As clsSQLiteQueryBuilder
    m_SoftDeleteMode = 1
    Set WithTrashed = Me
End Function

Public Function OnlyTrashed() As clsSQLiteQueryBuilder
    m_SoftDeleteMode = 2
    Set OnlyTrashed = Me
End Function

' =============================================================================
' Grouping, Sorting & Pagination
' =============================================================================

Public Function GroupByCols(ParamArray varCols() As Variant) As clsSQLiteQueryBuilder
    Dim lngIdx As Long
    Dim astrParts() As String
    Dim lngSub As Long
    Dim strItem As String

    For lngIdx = LBound(varCols) To UBound(varCols)
        astrParts = Split(CStr(varCols(lngIdx)), ",")
        For lngSub = LBound(astrParts) To UBound(astrParts)
            strItem = Trim$(astrParts(lngSub))
            If Len(strItem) > 0 Then m_GroupByItems.Add QuoteSQLiteIdent(strItem)
        Next lngSub
    Next lngIdx
    Set GroupByCols = Me
End Function

Public Function HavingRaw(ByVal strHavingSql As String, ParamArray varArgs() As Variant) As clsSQLiteQueryBuilder
    Dim lngIdx As Long
    m_HavingClauses.Add "(" & Trim$(strHavingSql) & ")"
    If Not IsMissing(varArgs) Then
        For lngIdx = LBound(varArgs) To UBound(varArgs)
            m_HavingParams.Add varArgs(lngIdx)
        Next lngIdx
    End If
    Set HavingRaw = Me
End Function

Public Function OrderByCol( _
    ByVal strCol As String, _
    Optional ByVal blnDescending As Boolean = False) As clsSQLiteQueryBuilder

    m_OrderByItems.Add QuoteSQLiteIdent(strCol) & IIf(blnDescending, " DESC", " ASC")
    Set OrderByCol = Me
End Function

Public Function OrderByDesc(ByVal strCol As String) As clsSQLiteQueryBuilder
    Set OrderByDesc = OrderByCol(strCol, True)
End Function

Public Function LimitRows(ByVal lngMaxRows As Long) As clsSQLiteQueryBuilder
    m_LimitCount = lngMaxRows
    Set LimitRows = Me
End Function

Public Function OffsetRows(ByVal lngSkipRows As Long) As clsSQLiteQueryBuilder
    m_OffsetCount = lngSkipRows
    Set OffsetRows = Me
End Function

Public Function Page(ByVal lngPageNumber As Long, ByVal lngPageSize As Long) As clsSQLiteQueryBuilder
    Dim lngSafePage As Long
    Dim lngSafeSize As Long
    lngSafePage = IIf(lngPageNumber < 1, 1, lngPageNumber)
    lngSafeSize = IIf(lngPageSize < 1, 1, lngPageSize)
    m_LimitCount = lngSafeSize
    m_OffsetCount = (lngSafePage - 1) * lngSafeSize
    Set Page = Me
End Function

' =============================================================================
' SQL Generation & Terminal Execution Methods
' =============================================================================

Public Function ToSql(Optional ByVal strOverrideSelect As String = vbNullString) As String
    Dim strSql As String
    Dim strProj As String
    Dim strWhere As String
    Dim lngIdx As Long

    If Len(strOverrideSelect) > 0 Then
        strProj = strOverrideSelect
    ElseIf m_SelectItems.Count = 0 Then
        If Len(m_TableAlias) > 0 And m_JoinClauses.Count = 0 Then
            strProj = QuoteSQLiteIdent(m_TableAlias) & ".*"
        Else
            strProj = "*"
        End If
    Else
        For lngIdx = 1 To m_SelectItems.Count
            If lngIdx > 1 Then strProj = strProj & ", "
            strProj = strProj & CStr(m_SelectItems(lngIdx))
        Next lngIdx
    End If

    strSql = "SELECT "
    If m_Distinct And Len(strOverrideSelect) = 0 Then strSql = strSql & "DISTINCT "
    strSql = strSql & strProj & " FROM " & QuoteSQLiteIdent(m_TableName)
    If Len(m_TableAlias) > 0 Then
        strSql = strSql & " AS " & QuoteSQLiteIdent(m_TableAlias)
    End If

    For lngIdx = 1 To m_JoinClauses.Count
        strSql = strSql & " " & CStr(m_JoinClauses(lngIdx))
    Next lngIdx

    strWhere = BuildWhereSql(True)
    If Len(strWhere) > 0 Then
        strSql = strSql & " WHERE " & strWhere
    End If

    If m_GroupByItems.Count > 0 And Len(strOverrideSelect) = 0 Then
        strSql = strSql & " GROUP BY "
        For lngIdx = 1 To m_GroupByItems.Count
            If lngIdx > 1 Then strSql = strSql & ", "
            strSql = strSql & CStr(m_GroupByItems(lngIdx))
        Next lngIdx
    End If

    If m_HavingClauses.Count > 0 And Len(strOverrideSelect) = 0 Then
        strSql = strSql & " HAVING "
        For lngIdx = 1 To m_HavingClauses.Count
            If lngIdx > 1 Then strSql = strSql & " AND "
            strSql = strSql & CStr(m_HavingClauses(lngIdx))
        Next lngIdx
    End If

    If Len(strOverrideSelect) = 0 Then
        If m_OrderByItems.Count > 0 Then
            strSql = strSql & " ORDER BY "
            For lngIdx = 1 To m_OrderByItems.Count
                If lngIdx > 1 Then strSql = strSql & ", "
                strSql = strSql & CStr(m_OrderByItems(lngIdx))
            Next lngIdx
        End If

        If m_LimitCount >= 0 Then
            strSql = strSql & " LIMIT " & m_LimitCount
            If m_OffsetCount > 0 Then
                strSql = strSql & " OFFSET " & m_OffsetCount
            End If
        ElseIf m_OffsetCount > 0 Then
            strSql = strSql & " LIMIT -1 OFFSET " & m_OffsetCount
        End If
    End If

    ToSql = strSql & ";"
End Function

Public Function BoundParameters(Optional ByVal blnIncludeHaving As Boolean = True) As Collection
    Dim colAll As Collection
    Dim varItem As Variant
    Set colAll = New Collection
    For Each varItem In m_WhereParams
        colAll.Add varItem
    Next varItem
    If blnIncludeHaving Then
        For Each varItem In m_HavingParams
            colAll.Add varItem
        Next varItem
    End If
    Set BoundParameters = colAll
End Function

Public Function ToRows() As Collection
    EnsureConn
    Set ToRows = m_Conn.QueryParams(ToSql(), BoundParameters(True), False, Nothing)
End Function

Public Function ToRowsRaw() As Collection
    EnsureConn
    Set ToRowsRaw = m_Conn.QueryParams(ToSql(), BoundParameters(True), True, m_TableMap)
End Function

Public Function FirstRow() As Scripting.Dictionary
    Dim lngSavedLimit As Long
    Dim colRows As Collection
    EnsureConn
    lngSavedLimit = m_LimitCount
    m_LimitCount = 1
    Set colRows = m_Conn.QueryParams(ToSql(), BoundParameters(True), False, Nothing)
    m_LimitCount = lngSavedLimit
    If colRows.Count > 0 Then Set FirstRow = colRows(1)
End Function

Public Function FirstRowRaw() As Scripting.Dictionary
    Dim lngSavedLimit As Long
    Dim colRows As Collection
    EnsureConn
    lngSavedLimit = m_LimitCount
    m_LimitCount = 1
    Set colRows = m_Conn.QueryParams(ToSql(), BoundParameters(True), True, m_TableMap)
    m_LimitCount = lngSavedLimit
    If colRows.Count > 0 Then Set FirstRowRaw = colRows(1)
End Function

Public Function ToEntities(ByVal objPrototype As ISQLiteEntity) As Collection
    Dim colRows As Collection
    Dim dictRow As Scripting.Dictionary
    Dim colEntities As Collection
    Dim objInst As ISQLiteEntity

    If objPrototype Is Nothing Then
        RaiseSQLiteError sqErrInvalidArgument, "clsSQLiteQueryBuilder.ToEntities", "Prototype ISQLiteEntity cannot be Nothing."
    End If

    Set colRows = ToRowsRaw()
    Set colEntities = New Collection
    For Each dictRow In colRows
        Set objInst = objPrototype.CreateBlank()
        objInst.ReadFromRow dictRow
        colEntities.Add objInst
    Next dictRow
    Set ToEntities = colEntities
End Function

Public Function FirstEntity(ByVal objPrototype As ISQLiteEntity) As ISQLiteEntity
    Dim dictRow As Scripting.Dictionary
    Dim objInst As ISQLiteEntity
    Set dictRow = FirstRowRaw()
    If Not dictRow Is Nothing Then
        Set objInst = objPrototype.CreateBlank()
        objInst.ReadFromRow dictRow
        Set FirstEntity = objInst
    End If
End Function

Public Function CountRows() As Long
    Dim varVal As Variant
    EnsureConn
    varVal = m_Conn.QueryScalarParams(ToSql("COUNT(*)"), BoundParameters(False))
    If Not IsNull(varVal) Then CountRows = CLng(varVal)
End Function

Public Function ExistsAny() As Boolean
    Dim lngSavedLimit As Long
    Dim strSubSql As String
    Dim varVal As Variant

    EnsureConn
    lngSavedLimit = m_LimitCount
    m_LimitCount = 1
    strSubSql = Left$(ToSql("1"), Len(ToSql("1")) - 1)
    m_LimitCount = lngSavedLimit

    varVal = m_Conn.QueryScalarParams("SELECT EXISTS(" & strSubSql & ");", BoundParameters(True))
    If Not IsNull(varVal) Then ExistsAny = (CLng(varVal) <> 0)
End Function

Public Function ScalarVal(Optional ByVal strSelectExpr As String = vbNullString) As Variant
    EnsureConn
    ScalarVal = m_Conn.QueryScalarParams(ToSql(strSelectExpr), BoundParameters(True))
End Function

Public Function PluckCol(ByVal strCol As String) As Collection
    Dim colRows As Collection
    Dim dictRow As Scripting.Dictionary
    Dim colVals As Collection
    Dim varFirstKey As Variant

    EnsureConn
    If m_SelectItems.Count = 0 Then SelectCols strCol
    Set colRows = ToRowsRaw()
    Set colVals = New Collection
    For Each dictRow In colRows
        If dictRow.Exists(strCol) Then
            colVals.Add dictRow(strCol)
        Else
            For Each varFirstKey In dictRow.Keys
                colVals.Add dictRow(varFirstKey)
                Exit For
            Next varFirstKey
        End If
    Next dictRow
    Set PluckCol = colVals
End Function

' Parameterized bulk UPDATE matching current WHERE clauses.
Public Function UpdateValues(ByVal dictSetValues As Scripting.Dictionary) As Long
    Dim dictWork As Scripting.Dictionary
    Dim colParams As Collection
    Dim strSetSql As String
    Dim strWhere As String
    Dim varKey As Variant
    Dim strCanon As String
    Dim varWhereArg As Variant

    EnsureConn
    If dictSetValues Is Nothing Or dictSetValues.Count = 0 Then Exit Function

    If Not m_TableMap Is Nothing Then
        Set dictWork = m_TableMap.ValidateAndCoerce(dictSetValues, False, False)
        If Len(m_TableMap.UpdatedAtColumn) > 0 And Not dictWork.Exists(m_TableMap.UpdatedAtColumn) Then
            dictWork(m_TableMap.UpdatedAtColumn) = UtcNowIsoString()
        End If
    Else
        Set dictWork = dictSetValues
    End If

    Set colParams = New Collection
    For Each varKey In dictWork.Keys
        strCanon = CStr(varKey)
        If Len(strSetSql) > 0 Then strSetSql = strSetSql & ", "
        strSetSql = strSetSql & QuoteSQLiteIdent(strCanon) & " = ?"
        colParams.Add dictWork(varKey)
    Next varKey

    strWhere = BuildWhereSql(False)
    For Each varWhereArg In m_WhereParams
        colParams.Add varWhereArg
    Next varWhereArg

    If Len(strWhere) > 0 Then
        UpdateValues = m_Conn.ExecuteParams( _
            "UPDATE " & QuoteSQLiteIdent(m_TableName) & " SET " & strSetSql & " WHERE " & strWhere & ";", colParams)
    Else
        UpdateValues = m_Conn.ExecuteParams( _
            "UPDATE " & QuoteSQLiteIdent(m_TableName) & " SET " & strSetSql & ";", colParams)
    End If
End Function

' Bulk DELETE matching current WHERE clauses (uses soft-delete timestamp when configured unless forced).
Public Function DeleteMatching(Optional ByVal blnForceHardDelete As Boolean = False) As Long
    Dim dictSoft As Scripting.Dictionary
    Dim strWhere As String

    EnsureConn
    If Not blnForceHardDelete And Not m_TableMap Is Nothing Then
        If m_TableMap.HasSoftDelete Then
            Set dictSoft = NewTextDict()
            dictSoft(m_TableMap.SoftDeleteColumn) = UtcNowIsoString()
            DeleteMatching = UpdateValues(dictSoft)
            Exit Function
        End If
    End If

    strWhere = BuildWhereSql(False)
    If Len(strWhere) > 0 Then
        DeleteMatching = m_Conn.ExecuteParams( _
            "DELETE FROM " & QuoteSQLiteIdent(m_TableName) & " WHERE " & strWhere & ";", BoundParameters(False))
    Else
        DeleteMatching = m_Conn.ExecuteParams( _
            "DELETE FROM " & QuoteSQLiteIdent(m_TableName) & ";", Empty)
    End If
End Function

' =============================================================================
' Internal Helpers
' =============================================================================

Private Function BuildWhereSql(ByVal blnQualifyAlias As Boolean) As String
    Dim colParts As Collection
    Dim strSoftCol As String
    Dim lngIdx As Long
    Dim strOut As String

    Set colParts = New Collection
    If Not m_TableMap Is Nothing Then
        If m_TableMap.HasSoftDelete Then
            If blnQualifyAlias And Len(m_TableAlias) > 0 Then
                strSoftCol = QuoteSQLiteIdent(m_TableAlias & "." & m_TableMap.SoftDeleteColumn)
            Else
                strSoftCol = QuoteSQLiteIdent(m_TableMap.SoftDeleteColumn)
            End If
            If m_SoftDeleteMode = 0 Then
                colParts.Add strSoftCol & " IS NULL"
            ElseIf m_SoftDeleteMode = 2 Then
                colParts.Add strSoftCol & " IS NOT NULL"
            End If
        End If
    End If

    For lngIdx = 1 To m_WhereClauses.Count
        colParts.Add CStr(m_WhereClauses(lngIdx))
    Next lngIdx

    For lngIdx = 1 To colParts.Count
        If lngIdx > 1 Then strOut = strOut & " AND "
        strOut = strOut & CStr(colParts(lngIdx))
    Next lngIdx
    BuildWhereSql = strOut
End Function

Private Function NormalizeListValues(ByVal varValues As Variant) As Collection
    Dim colOut As Collection
    Dim colIn As Collection
    Dim lngIdx As Long

    Set colOut = New Collection
    If IsNull(varValues) Or IsEmpty(varValues) Then
        Set NormalizeListValues = colOut
        Exit Function
    End If

    If IsObject(varValues) Then
        If varValues Is Nothing Then
            Set NormalizeListValues = colOut
            Exit Function
        ElseIf TypeOf varValues Is Collection Then
            Set colIn = varValues
            For lngIdx = 1 To colIn.Count
                colOut.Add colIn(lngIdx)
            Next lngIdx
            Set NormalizeListValues = colOut
            Exit Function
        End If
    End If

    If IsArray(varValues) Then
        On Error GoTo EmptyArr
        For lngIdx = LBound(varValues) To UBound(varValues)
            colOut.Add varValues(lngIdx)
        Next lngIdx
        On Error GoTo 0
    Else
        colOut.Add varValues
    End If
    Set NormalizeListValues = colOut
    Exit Function

EmptyArr:
    On Error GoTo 0
    Set NormalizeListValues = colOut
End Function

Private Sub EnsureConn()
    If m_Conn Is Nothing Then
        RaiseSQLiteError sqErrConnectionClosed, "clsSQLiteQueryBuilder", "QueryBuilder is not attached to a clsSQLiteConnection."
    End If
End Sub

