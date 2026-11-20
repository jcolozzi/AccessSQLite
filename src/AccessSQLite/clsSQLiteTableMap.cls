Attribute VB_Name = "clsSQLiteTableMap"
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
' clsSQLiteTableMap
' Whitelisted table & column schema map with fluent column/index builders,
' strict type coercion, required/max-length validation, dirty-field diffing,
' timestamps, and soft-delete configuration.
' =============================================================================

Private m_TableName As String
Private m_PrimaryKeyCol As String
Private m_CreatedAtCol As String
Private m_UpdatedAtCol As String
Private m_SoftDeleteCol As String
Private m_UseStrict As Boolean

Private m_ColumnOrder As Collection
Private m_ColumnsByName As Scripting.Dictionary
Private m_Indexes As Collection

Private Sub Class_Initialize()
    Set m_ColumnOrder = New Collection
    Set m_ColumnsByName = NewTextDict()
    Set m_Indexes = New Collection
End Sub

Public Sub InitMap(ByVal strTableName As String, Optional ByVal strPrimaryKey As String = "ID")
    m_TableName = Trim$(strTableName)
    If Len(m_TableName) = 0 Then
        RaiseSQLiteError sqErrInvalidArgument, "clsSQLiteTableMap.InitMap", "Table name cannot be empty."
    End If
    m_PrimaryKeyCol = Trim$(strPrimaryKey)
    Set m_ColumnOrder = New Collection
    Set m_ColumnsByName = NewTextDict()
    Set m_Indexes = New Collection
End Sub

Public Property Get TableName() As String
    TableName = m_TableName
End Property

Public Property Get PrimaryKeyColumn() As String
    PrimaryKeyColumn = m_PrimaryKeyCol
End Property
Public Property Let PrimaryKeyColumn(ByVal strCol As String)
    m_PrimaryKeyCol = Trim$(strCol)
End Property

Public Property Get CreatedAtColumn() As String
    CreatedAtColumn = m_CreatedAtCol
End Property

Public Property Get UpdatedAtColumn() As String
    UpdatedAtColumn = m_UpdatedAtCol
End Property

Public Property Get SoftDeleteColumn() As String
    SoftDeleteColumn = m_SoftDeleteCol
End Property

Public Property Get HasSoftDelete() As Boolean
    HasSoftDelete = (Len(m_SoftDeleteCol) > 0)
End Property

Public Property Get UseStrictTable() As Boolean
    UseStrictTable = m_UseStrict
End Property

Public Property Get Columns() As Collection
    Set Columns = m_ColumnOrder
End Property

Public Property Get Indexes() As Collection
    Set Indexes = m_Indexes
End Property

Public Function SetStrict(Optional ByVal blnStrict As Boolean = True) As clsSQLiteTableMap
    m_UseStrict = blnStrict
    Set SetStrict = Me
End Function

Public Function AddPrimaryKey( _
    ByVal strColName As String, _
    Optional ByVal enmKind As SQLiteColumnKind = sqColInt64, _
    Optional ByVal blnAutoIncrement As Boolean = True) As clsSQLiteTableMap

    m_PrimaryKeyCol = Trim$(strColName)
    Set AddPrimaryKey = AddColumn( _
        strColName:=m_PrimaryKeyCol, _
        enmKind:=enmKind, _
        strLabel:=m_PrimaryKeyCol, _
        blnIsPrimaryKey:=True, _
        blnAutoIncrement:=blnAutoIncrement, _
        blnRequired:=Not blnAutoIncrement, _
        blnClientEditable:=Not blnAutoIncrement)
End Function

Public Function AddColumn( _
    ByVal strColName As String, _
    ByVal enmKind As SQLiteColumnKind, _
    Optional ByVal strLabel As String = vbNullString, _
    Optional ByVal blnIsPrimaryKey As Boolean = False, _
    Optional ByVal blnAutoIncrement As Boolean = False, _
    Optional ByVal blnRequired As Boolean = False, _
    Optional ByVal lngMaxLength As Long = 0, _
    Optional ByVal blnUnique As Boolean = False, _
    Optional ByVal varDefault As Variant, _
    Optional ByVal strFkTable As String = vbNullString, _
    Optional ByVal strFkColumn As String = vbNullString, _
    Optional ByVal strOnDelete As String = vbNullString, _
    Optional ByVal blnClientEditable As Boolean = True, _
    Optional ByVal blnAudited As Boolean = True) As clsSQLiteTableMap

    Dim objCol As clsSQLiteColumnDef
    Dim strClean As String

    strClean = Trim$(strColName)
    If Len(strClean) = 0 Then
        RaiseSQLiteError sqErrInvalidArgument, "clsSQLiteTableMap.AddColumn", "Column name cannot be empty."
    End If

    If m_ColumnsByName.Exists(strClean) Then
        Set objCol = m_ColumnsByName(strClean)
    Else
        Set objCol = New clsSQLiteColumnDef
        objCol.ColumnName = strClean
        m_ColumnOrder.Add objCol
        Set m_ColumnsByName(strClean) = objCol
    End If

    objCol.ColumnKind = enmKind
    If Len(strLabel) > 0 Then objCol.LabelText = strLabel
    objCol.IsPrimaryKey = blnIsPrimaryKey
    objCol.IsAutoIncrement = blnAutoIncrement
    objCol.IsRequired = blnRequired
    objCol.MaxLength = lngMaxLength
    objCol.IsUnique = blnUnique
    If Not IsMissing(varDefault) Then
        objCol.DefaultValue = varDefault
    End If
    objCol.ForeignKeyTable = strFkTable
    objCol.ForeignKeyColumn = strFkColumn
    objCol.OnDeleteAction = strOnDelete
    objCol.IsClientEditable = blnClientEditable
    objCol.IsAudited = blnAudited

    If blnIsPrimaryKey Then m_PrimaryKeyCol = strClean
    Set AddColumn = Me
End Function

Public Function AddTimestamps( _
    Optional ByVal strCreatedCol As String = "CreatedAt", _
    Optional ByVal strUpdatedCol As String = "UpdatedAt") As clsSQLiteTableMap

    If Len(Trim$(strCreatedCol)) > 0 Then
        m_CreatedAtCol = Trim$(strCreatedCol)
        If Not HasColumn(m_CreatedAtCol) Then
            AddColumn m_CreatedAtCol, sqColDateTime, "Created at", False, False, False, 0, False, , , , , False, False
        End If
    End If
    If Len(Trim$(strUpdatedCol)) > 0 Then
        m_UpdatedAtCol = Trim$(strUpdatedCol)
        If Not HasColumn(m_UpdatedAtCol) Then
            AddColumn m_UpdatedAtCol, sqColDateTime, "Updated at", False, False, False, 0, False, , , , , False, False
        End If
    End If
    Set AddTimestamps = Me
End Function

Public Function AddSoftDelete(Optional ByVal strDeletedCol As String = "DeletedAt") As clsSQLiteTableMap
    m_SoftDeleteCol = Trim$(strDeletedCol)
    If Len(m_SoftDeleteCol) > 0 And Not HasColumn(m_SoftDeleteCol) Then
        AddColumn m_SoftDeleteCol, sqColDateTime, "Deleted at", False, False, False, 0, False, , , , , False, False
    End If
    Set AddSoftDelete = Me
End Function

Public Function AddIndex( _
    ByVal strIndexName As String, _
    ByVal strColumnsCsv As String, _
    Optional ByVal blnUnique As Boolean = False, _
    Optional ByVal strWhereSql As String = vbNullString) As clsSQLiteTableMap

    Dim dictIdx As Scripting.Dictionary
    Set dictIdx = NewTextDict()
    dictIdx("IndexName") = Trim$(strIndexName)
    dictIdx("ColumnsCsv") = Trim$(strColumnsCsv)
    dictIdx("IsUnique") = blnUnique
    dictIdx("WhereSql") = Trim$(strWhereSql)
    m_Indexes.Add dictIdx
    Set AddIndex = Me
End Function

Public Function HasColumn(ByVal strColName As String) As Boolean
    HasColumn = m_ColumnsByName.Exists(Trim$(strColName))
End Function

Public Function GetColumn(ByVal strColName As String) As clsSQLiteColumnDef
    Dim strKey As String
    strKey = Trim$(strColName)
    If m_ColumnsByName.Exists(strKey) Then
        Set GetColumn = m_ColumnsByName(strKey)
    End If
End Function

Public Function CanonicalColumn(ByVal strColName As String) As String
    Dim objCol As clsSQLiteColumnDef
    Set objCol = GetColumn(strColName)
    If objCol Is Nothing Then
        RaiseSQLiteError sqErrValidation, "clsSQLiteTableMap.CanonicalColumn", _
            "Unknown column '" & strColName & "' on table '" & m_TableName & "'."
    End If
    CanonicalColumn = objCol.ColumnName
End Function

' Coerces an untyped input value into the target column's strongly-typed VBA representation.
Public Function CoerceColumnValue( _
    ByVal strColName As String, _
    ByVal varValue As Variant, _
    Optional ByVal blnBlankTextToNull As Boolean = False) As Variant

    Dim objCol As clsSQLiteColumnDef
    Dim strLabel As String
    Dim strText As String
    Dim dtmParsed As Date
    Dim abytBlob() As Byte

    Set objCol = GetColumn(strColName)
    If objCol Is Nothing Then
        RaiseSQLiteError sqErrValidation, "clsSQLiteTableMap.CoerceColumnValue", _
            "Unknown column '" & strColName & "' on table '" & m_TableName & "'."
    End If

    strLabel = objCol.LabelText
    If IsNull(varValue) Or IsEmpty(varValue) Then
        CoerceColumnValue = Null
        Exit Function
    End If

    If VarType(varValue) = vbString Then
        If blnBlankTextToNull And Len(Trim$(CStr(varValue))) = 0 Then
            CoerceColumnValue = Null
            Exit Function
        End If
        If (objCol.ColumnKind <> sqColText And objCol.ColumnKind <> sqColMemo And _
            objCol.ColumnKind <> sqColJson And objCol.ColumnKind <> sqColGuid) And Len(Trim$(CStr(varValue))) = 0 Then
            CoerceColumnValue = Null
            Exit Function
        End If
    End If

    Select Case objCol.ColumnKind
        Case sqColText, sqColGuid
            strText = CStr(varValue)
            If objCol.MaxLength > 0 And Len(strText) > objCol.MaxLength Then
                RaiseSQLiteError sqErrValidation, "clsSQLiteTableMap.CoerceColumnValue", _
                    strLabel & " exceeds maximum length of " & objCol.MaxLength & " characters."
            End If
            CoerceColumnValue = strText

        Case sqColMemo, sqColJson
            strText = CStr(varValue)
            If objCol.MaxLength > 0 And Len(strText) > objCol.MaxLength Then
                RaiseSQLiteError sqErrValidation, "clsSQLiteTableMap.CoerceColumnValue", _
                    strLabel & " exceeds maximum length of " & objCol.MaxLength & " characters."
            End If
            CoerceColumnValue = strText

        Case sqColLong
            If Not IsNumeric(varValue) Then
                RaiseSQLiteError sqErrValidation, "clsSQLiteTableMap.CoerceColumnValue", strLabel & " must be a whole number."
            End If
            If CDbl(varValue) <> Fix(CDbl(varValue)) Or CDbl(varValue) < -2147483648# Or CDbl(varValue) > 2147483647# Then
                RaiseSQLiteError sqErrValidation, "clsSQLiteTableMap.CoerceColumnValue", strLabel & " must be a 32-bit integer."
            End If
            CoerceColumnValue = CLng(varValue)

        Case sqColInt64
            If Not IsNumeric(varValue) Then
                RaiseSQLiteError sqErrValidation, "clsSQLiteTableMap.CoerceColumnValue", strLabel & " must be a whole number."
            End If
            If CDec(varValue) <> Fix(CDec(varValue)) Then
                RaiseSQLiteError sqErrValidation, "clsSQLiteTableMap.CoerceColumnValue", strLabel & " must be an integer."
            End If
            If CDec(varValue) >= -2147483648# And CDec(varValue) <= 2147483647# Then
                CoerceColumnValue = CLng(varValue)
            Else
                CoerceColumnValue = CCur(varValue)
            End If

        Case sqColDouble
            If Not IsNumeric(varValue) Then
                RaiseSQLiteError sqErrValidation, "clsSQLiteTableMap.CoerceColumnValue", strLabel & " must be a numeric value."
            End If
            CoerceColumnValue = CDbl(varValue)

        Case sqColCurrency
            If Not IsNumeric(varValue) Then
                RaiseSQLiteError sqErrValidation, "clsSQLiteTableMap.CoerceColumnValue", strLabel & " must be a currency amount."
            End If
            CoerceColumnValue = CCur(varValue)

        Case sqColBoolean
            If VarType(varValue) = vbBoolean Then
                CoerceColumnValue = CBool(varValue)
            ElseIf IsNumeric(varValue) Then
                CoerceColumnValue = (CDbl(varValue) <> 0#)
            ElseIf StrComp(Trim$(CStr(varValue)), "true", vbTextCompare) = 0 Or Trim$(CStr(varValue)) = "1" Then
                CoerceColumnValue = True
            ElseIf StrComp(Trim$(CStr(varValue)), "false", vbTextCompare) = 0 Or Trim$(CStr(varValue)) = "0" Then
                CoerceColumnValue = False
            Else
                RaiseSQLiteError sqErrValidation, "clsSQLiteTableMap.CoerceColumnValue", strLabel & " must be True or False."
            End If

        Case sqColDate, sqColDateTime
            If VarType(varValue) = vbDate Then
                If objCol.ColumnKind = sqColDate Then
                    CoerceColumnValue = Int(CDate(varValue))
                Else
                    CoerceColumnValue = CDate(varValue)
                End If
            ElseIf TryParseIsoDate(CStr(varValue), dtmParsed) Then
                If objCol.ColumnKind = sqColDate Then
                    CoerceColumnValue = Int(dtmParsed)
                Else
                    CoerceColumnValue = dtmParsed
                End If
            ElseIf IsDate(varValue) Then
                If objCol.ColumnKind = sqColDate Then
                    CoerceColumnValue = Int(CDate(varValue))
                Else
                    CoerceColumnValue = CDate(varValue)
                End If
            Else
                RaiseSQLiteError sqErrValidation, "clsSQLiteTableMap.CoerceColumnValue", strLabel & " must be a valid date."
            End If

        Case sqColBlob
            If VarType(varValue) = (vbArray Or vbByte) Then
                abytBlob = varValue
                CoerceColumnValue = abytBlob
            ElseIf VarType(varValue) = vbString Then
                abytBlob = StrConv(CStr(varValue), vbFromUnicode)
                CoerceColumnValue = abytBlob
            Else
                RaiseSQLiteError sqErrValidation, "clsSQLiteTableMap.CoerceColumnValue", strLabel & " must be a Byte() BLOB."
            End If

        Case Else
            CoerceColumnValue = varValue
    End Select
End Function

' Validates and coerces an input dictionary for INSERT (blnIsNew=True) or UPDATE (blnIsNew=False).
Public Function ValidateAndCoerce( _
    ByVal dictInput As Scripting.Dictionary, _
    ByVal blnIsNew As Boolean, _
    Optional ByVal blnEnforceClientEditable As Boolean = False) As Scripting.Dictionary

    Dim dictClean As Scripting.Dictionary
    Dim varKey As Variant
    Dim strCanon As String
    Dim objCol As clsSQLiteColumnDef
    Dim colMissing As Collection
    Dim strMissingList As String
    Dim lngIdx As Long

    Set dictClean = NewTextDict()
    If Not dictInput Is Nothing Then
        For Each varKey In dictInput.Keys
            strCanon = CanonicalColumn(CStr(varKey))
            Set objCol = m_ColumnsByName(strCanon)
            If blnEnforceClientEditable And Not objCol.IsClientEditable Then
                RaiseSQLiteError sqErrValidation, "clsSQLiteTableMap.ValidateAndCoerce", _
                    "Column '" & strCanon & "' on '" & m_TableName & "' is not client-editable."
            End If
            dictClean(strCanon) = CoerceColumnValue(strCanon, dictInput(varKey))
        Next varKey
    End If

    ' Apply declared defaults on new records when omitted.
    If blnIsNew Then
        For Each objCol In m_ColumnOrder
            If Not dictClean.Exists(objCol.ColumnName) Then
                If objCol.HasDefault And Not objCol.IsAutoIncrement Then
                    dictClean(objCol.ColumnName) = objCol.DefaultValue
                End If
            End If
        Next objCol
    End If

    ' Check required columns.
    Set colMissing = New Collection
    For Each objCol In m_ColumnOrder
        If objCol.IsRequired And Not objCol.IsAutoIncrement Then
            If dictClean.Exists(objCol.ColumnName) Then
                If IsNull(dictClean(objCol.ColumnName)) Then
                    colMissing.Add objCol.LabelText
                ElseIf VarType(dictClean(objCol.ColumnName)) = vbString Then
                    If Len(Trim$(CStr(dictClean(objCol.ColumnName)))) = 0 Then
                        colMissing.Add objCol.LabelText
                    End If
                End If
            ElseIf blnIsNew And Not objCol.HasDefault And Len(objCol.DefaultRawSql) = 0 Then
                colMissing.Add objCol.LabelText
            End If
        End If
    Next objCol

    If colMissing.Count > 0 Then
        For lngIdx = 1 To colMissing.Count
            If lngIdx > 1 Then strMissingList = strMissingList & ", "
            strMissingList = strMissingList & CStr(colMissing(lngIdx))
        Next lngIdx
        RaiseSQLiteError sqErrValidation, "clsSQLiteTableMap.ValidateAndCoerce", _
            "Missing required field(s) on '" & m_TableName & "': " & strMissingList
    End If

    Set ValidateAndCoerce = dictClean
End Function

' Returns only the columns whose coerced values differ from dictStored (dirty-field detection).
Public Function DiffChangedColumns( _
    ByVal dictIncoming As Scripting.Dictionary, _
    ByVal dictStored As Scripting.Dictionary) As Scripting.Dictionary

    Dim dictDiff As Scripting.Dictionary
    Dim varKey As Variant
    Dim strCanon As String
    Dim varNewVal As Variant

    Set dictDiff = NewTextDict()
    If dictIncoming Is Nothing Then
        Set DiffChangedColumns = dictDiff
        Exit Function
    End If

    For Each varKey In dictIncoming.Keys
        strCanon = CanonicalColumn(CStr(varKey))
        If StrComp(strCanon, m_PrimaryKeyCol, vbTextCompare) <> 0 Then
            varNewVal = CoerceColumnValue(strCanon, dictIncoming(varKey))
            If dictStored Is Nothing Then
                dictDiff(strCanon) = varNewVal
            ElseIf Not dictStored.Exists(strCanon) Then
                dictDiff(strCanon) = varNewVal
            ElseIf Not ValuesEqual(varNewVal, CoerceColumnValue(strCanon, dictStored(strCanon))) Then
                dictDiff(strCanon) = varNewVal
            End If
        End If
    Next varKey

    Set DiffChangedColumns = dictDiff
End Function

' Hydrates raw SQLite row values into exact VBA types declared by this TableMap.
Public Sub HydrateRowInPlace(ByVal dictRow As Scripting.Dictionary)
    Dim varKey As Variant
    Dim objCol As clsSQLiteColumnDef
    Dim varVal As Variant
    Dim dtmParsed As Date

    If dictRow Is Nothing Then Exit Sub
    For Each varKey In dictRow.Keys
        If m_ColumnsByName.Exists(CStr(varKey)) Then
            Set objCol = m_ColumnsByName(CStr(varKey))
            varVal = dictRow(varKey)
            If Not IsNull(varVal) And Not IsEmpty(varVal) Then
                Select Case objCol.ColumnKind
                    Case sqColBoolean
                        If VarType(varVal) <> vbBoolean Then dictRow(varKey) = (CDbl(varVal) <> 0#)
                    Case sqColCurrency
                        If VarType(varVal) <> vbCurrency Then dictRow(varKey) = CCur(varVal)
                    Case sqColLong
                        If VarType(varVal) <> vbLong Then dictRow(varKey) = CLng(varVal)
                    Case sqColDate, sqColDateTime
                        If VarType(varVal) = vbString Then
                            If TryParseIsoDate(CStr(varVal), dtmParsed) Then dictRow(varKey) = dtmParsed
                        ElseIf IsNumeric(varVal) And VarType(varVal) <> vbDate Then
                            dictRow(varKey) = JulianDayToDate(CDbl(varVal))
                        End If
                End Select
            End If
        End If
    Next varKey
End Sub

Private Function ValuesEqual(ByVal varA As Variant, ByVal varB As Variant) As Boolean
    Dim abytA() As Byte
    Dim abytB() As Byte
    Dim lngLenA As Long
    Dim lngLenB As Long
    Dim lngIdx As Long

    If IsNull(varA) And IsNull(varB) Then
        ValuesEqual = True
        Exit Function
    End If
    If IsNull(varA) Or IsNull(varB) Then Exit Function

    If VarType(varA) = (vbArray Or vbByte) And VarType(varB) = (vbArray Or vbByte) Then
        abytA = varA
        abytB = varB
        lngLenA = GetByteArrayLength(abytA)
        lngLenB = GetByteArrayLength(abytB)
        If lngLenA <> lngLenB Then Exit Function
        For lngIdx = 0 To lngLenA - 1
            If abytA(LBound(abytA) + lngIdx) <> abytB(LBound(abytB) + lngIdx) Then Exit Function
        Next lngIdx
        ValuesEqual = True
        Exit Function
    End If

    If VarType(varA) = vbString And VarType(varB) = vbString Then
        ValuesEqual = (StrComp(CStr(varA), CStr(varB), vbBinaryCompare) = 0)
    Else
        ValuesEqual = (varA = varB)
    End If
End Function

