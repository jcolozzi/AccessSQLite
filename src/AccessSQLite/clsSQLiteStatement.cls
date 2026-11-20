Attribute VB_Name = "clsSQLiteStatement"
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
' clsSQLiteStatement
' Prepared SQLite statement wrapper supporting ?, @name, :name, $name bindings,
' Dictionary/Collection/Array parameter lists, exact type marshaling, and
' reusable cached statement reset.
' =============================================================================

Private m_Driver As ISQLiteDriver
Private m_DbHandle As LongPtr
Private m_StmtHandle As LongPtr
Private m_SqlText As String
Private m_DateStorage As SQLiteDateStorage
Private m_IsCached As Boolean

Private Sub Class_Terminate()
    If Not m_IsCached Then
        FinalizeStatement
    End If
End Sub

Public Sub InitStatement( _
    ByVal objDriver As ISQLiteDriver, _
    ByVal ptrDb As LongPtr, _
    ByVal ptrStmt As LongPtr, _
    ByVal strSql As String, _
    ByVal enmDateStorage As SQLiteDateStorage, _
    ByVal blnCached As Boolean)

    Set m_Driver = objDriver
    m_DbHandle = ptrDb
    m_StmtHandle = ptrStmt
    m_SqlText = strSql
    m_DateStorage = enmDateStorage
    m_IsCached = blnCached
End Sub

Public Property Get StmtHandle() As LongPtr
    StmtHandle = m_StmtHandle
End Property

Public Property Get SqlText() As String
    SqlText = m_SqlText
End Property

Public Property Get IsCached() As Boolean
    IsCached = m_IsCached
End Property

Public Property Get IsValid() As Boolean
    IsValid = (m_StmtHandle <> 0 And Not m_Driver Is Nothing)
End Property

Public Property Get ParameterCount() As Long
    EnsureValid
    ParameterCount = m_Driver.GetParamCount(m_StmtHandle)
End Property

Public Property Get ExpandedSql() As String
    If Not IsValid Then Exit Property
    ExpandedSql = m_Driver.GetExpandedSql(m_StmtHandle)
    If Len(ExpandedSql) = 0 Then ExpandedSql = m_SqlText
End Property

Public Sub ResetForReuse()
    If m_StmtHandle = 0 Or m_Driver Is Nothing Then Exit Sub
    m_Driver.ResetStmt m_StmtHandle
    m_Driver.ClearStmtBindings m_StmtHandle
End Sub

Public Sub FinalizeStatement()
    If m_StmtHandle <> 0 And Not m_Driver Is Nothing Then
        m_Driver.FinalizeStmt m_StmtHandle
    End If
    m_StmtHandle = 0
    m_IsCached = False
End Sub

' Binds a single value at a 1-based parameter position.
Public Sub BindByIndex(ByVal lngOneBasedIndex As Long, ByVal varVal As Variant)
    Dim lngRc As Long
    Dim intVt As Integer
    Dim abytBlob() As Byte

    EnsureValid

    If IsNull(varVal) Or IsEmpty(varVal) Then
        lngRc = m_Driver.BindNull(m_StmtHandle, lngOneBasedIndex)
        CheckBindRc lngRc, "#" & lngOneBasedIndex
        Exit Sub
    End If

    If IsObject(varVal) Then
        If varVal Is Nothing Then
            lngRc = m_Driver.BindNull(m_StmtHandle, lngOneBasedIndex)
            CheckBindRc lngRc, "#" & lngOneBasedIndex
            Exit Sub
        End If
        RaiseSQLiteError sqErrInvalidArgument, "clsSQLiteStatement.BindByIndex", _
            "Cannot bind unsupported Object parameter at index " & lngOneBasedIndex & "."
    End If

    intVt = VarType(varVal)
    If intVt = (vbArray Or vbByte) Then
        abytBlob = varVal
        lngRc = m_Driver.BindBlob(m_StmtHandle, lngOneBasedIndex, abytBlob)
        CheckBindRc lngRc, "#" & lngOneBasedIndex
        Exit Sub
    End If

    Select Case intVt
        Case vbBoolean
            lngRc = m_Driver.BindInt32(m_StmtHandle, lngOneBasedIndex, IIf(CBool(varVal), 1&, 0&))

        Case vbByte, vbInteger, vbLong
            lngRc = m_Driver.BindInt32(m_StmtHandle, lngOneBasedIndex, CLng(varVal))

        #If Win64 Then
        Case vbLongLong
            lngRc = m_Driver.BindInt64(m_StmtHandle, lngOneBasedIndex, varVal)
        #End If

        Case vbCurrency, vbDecimal
            If varVal = Int(varVal) Then
                If varVal >= -2147483648@ And varVal <= 2147483647@ Then
                    lngRc = m_Driver.BindInt32(m_StmtHandle, lngOneBasedIndex, CLng(varVal))
                Else
                    lngRc = m_Driver.BindInt64(m_StmtHandle, lngOneBasedIndex, CCur(varVal))
                End If
            Else
                lngRc = m_Driver.BindDouble(m_StmtHandle, lngOneBasedIndex, CDbl(varVal))
            End If

        Case vbSingle, vbDouble
            lngRc = m_Driver.BindDouble(m_StmtHandle, lngOneBasedIndex, CDbl(varVal))

        Case vbDate
            Select Case m_DateStorage
                Case sqDateJulianReal
                    lngRc = m_Driver.BindDouble(m_StmtHandle, lngOneBasedIndex, DateToJulianDay(CDate(varVal)))
                Case sqDateUnixSeconds
                    lngRc = m_Driver.BindInt64(m_StmtHandle, lngOneBasedIndex, DateToUnixSeconds(CDate(varVal)))
                Case Else
                    lngRc = m_Driver.BindText(m_StmtHandle, lngOneBasedIndex, FormatIsoDateValue(CDate(varVal), False))
            End Select

        Case Else
            lngRc = m_Driver.BindText(m_StmtHandle, lngOneBasedIndex, CStr(varVal))
    End Select

    CheckBindRc lngRc, "#" & lngOneBasedIndex
End Sub

' Binds a value to a named parameter (@name, :name, or $name).
Public Sub BindByName(ByVal strParamName As String, ByVal varVal As Variant)
    Dim lngIdx As Long
    EnsureValid
    lngIdx = ResolveNamedParamIndex(strParamName)
    If lngIdx <= 0 Then
        RaiseSQLiteError sqErrInvalidArgument, "clsSQLiteStatement.BindByName", _
            "Named parameter '" & strParamName & "' was not found in SQL: " & Left$(m_SqlText, 160)
    End If
    BindByIndex lngIdx, varVal
End Sub

' Binds all entries of a Scripting.Dictionary by parameter name.
Public Sub BindDict(ByVal dictParams As Scripting.Dictionary)
    Dim varKey As Variant
    If dictParams Is Nothing Then Exit Sub
    For Each varKey In dictParams.Keys
        BindByName CStr(varKey), dictParams(varKey)
    Next varKey
End Sub

' Binds flexible parameter inputs: Empty, Dictionary, Collection, positional array, or (@name, val) pairs.
Public Sub BindFlexibleParams(ByVal varInput As Variant)
    Dim colItems As Collection
    Dim lngIdx As Long
    Dim lngLower As Long
    Dim lngUpper As Long
    Dim lngElemCount As Long
    Dim blnTreatAsNamedPairs As Boolean

    EnsureValid
    If IsEmpty(varInput) Or IsNull(varInput) Then Exit Sub

    If IsObject(varInput) Then
        If varInput Is Nothing Then Exit Sub
        If TypeOf varInput Is Scripting.Dictionary Then
            BindDict varInput
            Exit Sub
        ElseIf TypeOf varInput Is Collection Then
            Set colItems = varInput
            For lngIdx = 1 To colItems.Count
                BindByIndex lngIdx, colItems(lngIdx)
            Next lngIdx
            Exit Sub
        End If
    End If

    If Not IsArray(varInput) Then
        BindByIndex 1, varInput
        Exit Sub
    End If

    On Error GoTo EmptyParamArray
    lngLower = LBound(varInput)
    lngUpper = UBound(varInput)
    On Error GoTo 0
    If lngUpper < lngLower Then Exit Sub

    lngElemCount = (lngUpper - lngLower) + 1

    ' Unwrap single-element container when caller passed a Dictionary or Collection inside ParamArray.
    If lngElemCount = 1 Then
        If IsObject(varInput(lngLower)) Then
            If varInput(lngLower) Is Nothing Then
                BindByIndex 1, Null
            ElseIf TypeOf varInput(lngLower) Is Scripting.Dictionary Then
                BindDict varInput(lngLower)
            ElseIf TypeOf varInput(lngLower) Is Collection Then
                BindFlexibleParams varInput(lngLower)
            Else
                BindByIndex 1, varInput(lngLower)
            End If
            Exit Sub
        ElseIf IsArray(varInput(lngLower)) And VarType(varInput(lngLower)) <> (vbArray Or vbByte) Then
            BindFlexibleParams varInput(lngLower)
            Exit Sub
        End If
    End If

    ' Check if arguments are ("@name", value, "@name2", value2) pairs matching the SQL statement.
    If (lngElemCount >= 2) And ((lngElemCount Mod 2) = 0) Then
        If VarType(varInput(lngLower)) = vbString Then
            If ResolveNamedParamIndex(CStr(varInput(lngLower))) > 0 Then
                blnTreatAsNamedPairs = True
            End If
        End If
    End If

    If blnTreatAsNamedPairs Then
        For lngIdx = lngLower To lngUpper - 1 Step 2
            BindByName CStr(varInput(lngIdx)), varInput(lngIdx + 1)
        Next lngIdx
    Else
        For lngIdx = lngLower To lngUpper
            BindByIndex (lngIdx - lngLower) + 1, varInput(lngIdx)
        Next lngIdx
    End If
    Exit Sub

EmptyParamArray:
    On Error GoTo 0
End Sub

Public Function ResolveNamedParamIndex(ByVal strParamName As String) As Long
    Dim strClean As String
    Dim strBare As String
    Dim lngIdx As Long

    strClean = Trim$(strParamName)
    If Len(strClean) = 0 Then Exit Function

    lngIdx = m_Driver.GetParamIndex(m_StmtHandle, strClean)
    If lngIdx > 0 Then
        ResolveNamedParamIndex = lngIdx
        Exit Function
    End If

    If Left$(strClean, 1) = "@" Or Left$(strClean, 1) = ":" Or Left$(strClean, 1) = "$" Then
        strBare = Mid$(strClean, 2)
    Else
        strBare = strClean
    End If

    lngIdx = m_Driver.GetParamIndex(m_StmtHandle, "@" & strBare)
    If lngIdx = 0 Then lngIdx = m_Driver.GetParamIndex(m_StmtHandle, ":" & strBare)
    If lngIdx = 0 Then lngIdx = m_Driver.GetParamIndex(m_StmtHandle, "$" & strBare)
    ResolveNamedParamIndex = lngIdx
End Function

Public Function StepOnce() As Long
    Dim lngRc As Long
    EnsureValid
    lngRc = m_Driver.StepStmt(m_StmtHandle)
    If lngRc <> sqRow And lngRc <> sqDone And lngRc <> sqOk Then
        RaiseDriverSqlError "StepOnce", lngRc
    End If
    StepOnce = lngRc
End Function

Public Function ColumnCount() As Long
    EnsureValid
    ColumnCount = m_Driver.GetColumnCount(m_StmtHandle)
End Function

Public Function ColumnName(ByVal lngZeroBasedCol As Long) As String
    EnsureValid
    ColumnName = m_Driver.GetColumnName(m_StmtHandle, lngZeroBasedCol)
End Function

Public Function ColumnDeclType(ByVal lngZeroBasedCol As Long) As String
    EnsureValid
    ColumnDeclType = m_Driver.GetColumnDeclType(m_StmtHandle, lngZeroBasedCol)
End Function

Public Function ColumnAffinity(ByVal lngZeroBasedCol As Long) As SQLiteNativeAffinity
    EnsureValid
    ColumnAffinity = m_Driver.GetColumnType(m_StmtHandle, lngZeroBasedCol)
End Function

Public Function ReadColumnValue( _
    ByVal lngZeroBasedCol As Long, _
    Optional ByVal blnRawTypes As Boolean = False, _
    Optional ByVal strUpperDeclType As String = vbNullString) As Variant

    Dim enmType As SQLiteNativeAffinity
    Dim varInt As Variant
    Dim strText As String
    Dim dtmParsed As Date

    enmType = m_Driver.GetColumnType(m_StmtHandle, lngZeroBasedCol)
    Select Case enmType
        Case sqAffinityNull
            ReadColumnValue = Null

        Case sqAffinityInteger
            varInt = m_Driver.GetColumnInt64(m_StmtHandle, lngZeroBasedCol)
            If blnRawTypes And Len(strUpperDeclType) > 0 Then
                If InStr(1, strUpperDeclType, "BOOL", vbBinaryCompare) > 0 Or strUpperDeclType = "BIT" Then
                    ReadColumnValue = (varInt <> 0)
                    Exit Function
                End If
            End If
            ReadColumnValue = varInt

        Case sqAffinityFloat
            If blnRawTypes And (InStr(1, strUpperDeclType, "DATE", vbBinaryCompare) > 0 Or InStr(1, strUpperDeclType, "TIME", vbBinaryCompare) > 0) Then
                ReadColumnValue = JulianDayToDate(m_Driver.GetColumnDouble(m_StmtHandle, lngZeroBasedCol))
            ElseIf blnRawTypes And (InStr(1, strUpperDeclType, "CURR", vbBinaryCompare) > 0 Or InStr(1, strUpperDeclType, "MONEY", vbBinaryCompare) > 0) Then
                ReadColumnValue = CCur(m_Driver.GetColumnDouble(m_StmtHandle, lngZeroBasedCol))
            Else
                ReadColumnValue = m_Driver.GetColumnDouble(m_StmtHandle, lngZeroBasedCol)
            End If

        Case sqAffinityText
            strText = m_Driver.GetColumnText(m_StmtHandle, lngZeroBasedCol)
            If blnRawTypes And Len(strUpperDeclType) > 0 Then
                If InStr(1, strUpperDeclType, "DATE", vbBinaryCompare) > 0 Or InStr(1, strUpperDeclType, "TIME", vbBinaryCompare) > 0 Then
                    If TryParseIsoDate(strText, dtmParsed) Then
                        ReadColumnValue = dtmParsed
                        Exit Function
                    End If
                End If
            End If
            ReadColumnValue = strText

        Case sqAffinityBlob
            ReadColumnValue = m_Driver.GetColumnBlob(m_StmtHandle, lngZeroBasedCol)

        Case Else
            ReadColumnValue = Null
    End Select
End Function

' Steps through all result rows and returns a Collection of Scripting.Dictionary (CompareMode = TextCompare).
Public Function ReadAllRows( _
    Optional ByVal blnRawTypes As Boolean = False, _
    Optional ByVal objTableMap As clsSQLiteTableMap = Nothing) As Collection

    Dim colRows As Collection
    Dim dictRow As Scripting.Dictionary
    Dim lngColCount As Long
    Dim lngCol As Long
    Dim astrNames() As String
    Dim astrDeclTypes() As String
    Dim lngRc As Long
    Dim blnMetaLoaded As Boolean

    EnsureValid
    Set colRows = New Collection

    On Error GoTo CleanFail
    Do
        lngRc = m_Driver.StepStmt(m_StmtHandle)
        If lngRc = sqDone Then Exit Do
        If lngRc <> sqRow Then
            RaiseDriverSqlError "ReadAllRows", lngRc
        End If

        If Not blnMetaLoaded Then
            lngColCount = m_Driver.GetColumnCount(m_StmtHandle)
            If lngColCount > 0 Then
                ReDim astrNames(0 To lngColCount - 1)
                ReDim astrDeclTypes(0 To lngColCount - 1)
                For lngCol = 0 To lngColCount - 1
                    astrNames(lngCol) = m_Driver.GetColumnName(m_StmtHandle, lngCol)
                    If Len(astrNames(lngCol)) = 0 Then astrNames(lngCol) = "Col" & lngCol
                    astrDeclTypes(lngCol) = UCase$(Trim$(m_Driver.GetColumnDeclType(m_StmtHandle, lngCol)))
                Next lngCol
            End If
            blnMetaLoaded = True
        End If

        Set dictRow = NewTextDict()
        For lngCol = 0 To lngColCount - 1
            dictRow(astrNames(lngCol)) = ReadColumnValue(lngCol, blnRawTypes, astrDeclTypes(lngCol))
        Next lngCol

        If blnRawTypes And Not objTableMap Is Nothing Then
            objTableMap.HydrateRowInPlace dictRow
        End If

        colRows.Add dictRow
    Loop

    ResetForReuse
    Set ReadAllRows = colRows
    Exit Function

CleanFail:
    Dim lngErrNum As Long
    Dim strErrSrc As String
    Dim strErrDesc As String
    lngErrNum = Err.Number
    strErrSrc = Err.Source
    strErrDesc = Err.Description
    ResetForReuse
    Err.Raise lngErrNum, strErrSrc, strErrDesc
End Function

Private Sub EnsureValid()
    If m_StmtHandle = 0 Or m_Driver Is Nothing Then
        RaiseSQLiteError sqErrConnectionClosed, "clsSQLiteStatement", "Prepared statement is closed or uninitialized."
    End If
End Sub

Private Sub CheckBindRc(ByVal lngRc As Long, ByVal strParamLabel As String)
    If lngRc <> sqOk Then
        RaiseDriverSqlError "Bind(" & strParamLabel & ")", lngRc
    End If
End Sub

Private Sub RaiseDriverSqlError(ByVal strMethod As String, ByVal lngRc As Long)
    Dim lngExt As Long
    Dim strMsg As String
    Dim enmErr As SQLiteErrorNumber

    lngExt = m_Driver.GetExtendedErrCode(m_DbHandle)
    strMsg = m_Driver.GetErrMsg(m_DbHandle)
    If Len(strMsg) = 0 Then strMsg = "SQLite step/bind failed."
    strMsg = strMsg & " | SQL: " & Left$(m_SqlText, 200)

    If lngRc = sqConstraint Or (lngExt And &HFF&) = sqConstraint Then
        enmErr = sqErrConstraint
    Else
        enmErr = sqErrSqlFailed
    End If
    RaiseSQLiteError enmErr, "clsSQLiteStatement." & strMethod, strMsg, lngRc, lngExt
End Sub

