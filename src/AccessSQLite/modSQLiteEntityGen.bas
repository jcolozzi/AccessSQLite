Attribute VB_Name = "modSQLiteEntityGen"
Option Compare Database
Option Explicit

' =============================================================================
' modSQLiteEntityGen
' Generates strongly-typed VBA ORM entity class code (Implements ISQLiteEntity)
' from any mapped or reflected SQLite table, automatically avoiding Access/VBA
' reserved words in generated property and member identifiers.
' =============================================================================

Private Const RESERVED_IDENTIFIERS As String = _
    "|DATE|TIME|NOW|YEAR|MONTH|DAY|HOUR|MINUTE|SECOND|NAME|VALUE|TEXT|TYPE|COUNT|INDEX|" & _
    "SELECT|FROM|WHERE|JOIN|ORDER|GROUP|UNION|TABLE|FIELD|FIELDS|RECORDSET|DATABASE|" & _
    "FORM|FORMS|REPORT|REPORTS|SECTION|CONTROLS|MODULES|PAGES|LEFT|RIGHT|MID|LEN|INSTR|" & _
    "OPEN|CLOSE|STEP|RESET|ERROR|STRING|LONG|INTEGER|DOUBLE|BOOLEAN|CURRENCY|VARIANT|OBJECT|"

Public Function GenerateEntityClassCode( _
    ByVal objConn As clsSQLiteConnection, _
    ByVal strTableName As String, _
    Optional ByVal strClassName As String = vbNullString) As String

    Dim objMap As clsSQLiteTableMap
    Dim objCol As clsSQLiteColumnDef
    Dim strCls As String
    Dim strProp As String
    Dim strVbaType As String
    Dim strFieldsBlock As String
    Dim strPropsBlock As String
    Dim strReadBlock As String
    Dim strWriteBlock As String
    Dim strOut As String

    If objConn Is Nothing Then
        RaiseSQLiteError sqErrConnectionClosed, "modSQLiteEntityGen.GenerateEntityClassCode", "Connection is Nothing."
    End If

    Set objMap = objConn.GetTableMap(strTableName)
    If objMap Is Nothing Then
        Set objMap = objConn.SchemaManager().ReflectTableMap(strTableName)
    End If

    strCls = Trim$(strClassName)
    If Len(strCls) = 0 Then
        strCls = "cls" & SanitizePascalIdentifier(objMap.TableName, False)
    End If

    For Each objCol In objMap.Columns
        strProp = SanitizePascalIdentifier(objCol.ColumnName, True)
        strVbaType = MapColumnToVbaPropertyType(objCol)

        strFieldsBlock = strFieldsBlock & "Private m_" & strProp & " As " & strVbaType & vbCrLf

        strPropsBlock = strPropsBlock & _
            "Public Property Get " & strProp & "() As " & strVbaType & vbCrLf & _
            "    " & strProp & " = m_" & strProp & vbCrLf & _
            "End Property" & vbCrLf & _
            "Public Property Let " & strProp & "(ByVal valInput As " & strVbaType & ")" & vbCrLf & _
            "    m_" & strProp & " = valInput" & vbCrLf & _
            "End Property" & vbCrLf & vbCrLf

        If strVbaType = "Variant" Then
            strReadBlock = strReadBlock & _
                "    If dictRow.Exists(""" & objCol.ColumnName & """) Then m_" & strProp & " = dictRow(""" & objCol.ColumnName & """)" & vbCrLf
        ElseIf strVbaType = "String" Then
            strReadBlock = strReadBlock & _
                "    If dictRow.Exists(""" & objCol.ColumnName & """) Then m_" & strProp & " = CStr(Nz(dictRow(""" & objCol.ColumnName & """), vbNullString))" & vbCrLf
        ElseIf strVbaType = "Boolean" Then
            strReadBlock = strReadBlock & _
                "    If dictRow.Exists(""" & objCol.ColumnName & """) Then m_" & strProp & " = CBool(Nz(dictRow(""" & objCol.ColumnName & """), False))" & vbCrLf
        ElseIf strVbaType = "Long" Then
            strReadBlock = strReadBlock & _
                "    If dictRow.Exists(""" & objCol.ColumnName & """) Then m_" & strProp & " = CLng(Nz(dictRow(""" & objCol.ColumnName & """), 0))" & vbCrLf
        ElseIf strVbaType = "Double" Then
            strReadBlock = strReadBlock & _
                "    If dictRow.Exists(""" & objCol.ColumnName & """) Then m_" & strProp & " = CDbl(Nz(dictRow(""" & objCol.ColumnName & """), 0#))" & vbCrLf
        ElseIf strVbaType = "Currency" Then
            strReadBlock = strReadBlock & _
                "    If dictRow.Exists(""" & objCol.ColumnName & """) Then m_" & strProp & " = CCur(Nz(dictRow(""" & objCol.ColumnName & """), 0@))" & vbCrLf
        End If

        strWriteBlock = strWriteBlock & _
            "    dictRow(""" & objCol.ColumnName & """) = m_" & strProp & vbCrLf
    Next objCol

    strOut = _
        "Option Compare Database" & vbCrLf & _
        "Option Explicit" & vbCrLf & vbCrLf & _
        "' Auto-generated ORM entity class for SQLite table [" & objMap.TableName & "]" & vbCrLf & _
        "Implements ISQLiteEntity" & vbCrLf & vbCrLf & _
        strFieldsBlock & vbCrLf & _
        strPropsBlock & _
        "Private Property Get ISQLiteEntity_EntityTableName() As String" & vbCrLf & _
        "    ISQLiteEntity_EntityTableName = """ & objMap.TableName & """" & vbCrLf & _
        "End Property" & vbCrLf & vbCrLf & _
        "Private Function ISQLiteEntity_CreateBlank() As ISQLiteEntity" & vbCrLf & _
        "    Set ISQLiteEntity_CreateBlank = New " & strCls & vbCrLf & _
        "End Function" & vbCrLf & vbCrLf & _
        "Private Sub ISQLiteEntity_ReadFromRow(ByVal dictRow As Scripting.Dictionary)" & vbCrLf & _
        "    If dictRow Is Nothing Then Exit Sub" & vbCrLf & _
        strReadBlock & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub ISQLiteEntity_WriteToRow(ByVal dictRow As Scripting.Dictionary)" & vbCrLf & _
        "    If dictRow Is Nothing Then Exit Sub" & vbCrLf & _
        strWriteBlock & _
        "End Sub"

    GenerateEntityClassCode = strOut
End Function

Private Function MapColumnToVbaPropertyType(ByVal objCol As clsSQLiteColumnDef) As String
    ' Nullable non-string fields use Variant so SQL NULL round-trips faithfully.
    If Not objCol.IsPrimaryKey And Not objCol.IsRequired Then
        Select Case objCol.ColumnKind
            Case sqColText, sqColMemo, sqColJson, sqColGuid
                MapColumnToVbaPropertyType = "Variant"
            Case Else
                MapColumnToVbaPropertyType = "Variant"
        End Select
        Exit Function
    End If

    Select Case objCol.ColumnKind
        Case sqColLong
            MapColumnToVbaPropertyType = "Long"
        Case sqColInt64
            If objCol.IsPrimaryKey Then
                MapColumnToVbaPropertyType = "Long"
            Else
                MapColumnToVbaPropertyType = "Variant"
            End If
        Case sqColBoolean
            MapColumnToVbaPropertyType = "Boolean"
        Case sqColDouble
            MapColumnToVbaPropertyType = "Double"
        Case sqColCurrency
            MapColumnToVbaPropertyType = "Currency"
        Case sqColText, sqColMemo, sqColJson, sqColGuid
            MapColumnToVbaPropertyType = "String"
        Case Else
            MapColumnToVbaPropertyType = "Variant"
    End Select
End Function

Public Function SanitizePascalIdentifier(ByVal strRaw As String, ByVal blnCheckReserved As Boolean) As String
    Dim lngIdx As Long
    Dim strCh As String
    Dim intAsc As Integer
    Dim blnCapNext As Boolean
    Dim strOut As String

    blnCapNext = True
    For lngIdx = 1 To Len(strRaw)
        strCh = Mid$(strRaw, lngIdx, 1)
        intAsc = AscW(strCh)
        If (intAsc >= 65 And intAsc <= 90) Or (intAsc >= 97 And intAsc <= 122) Or (intAsc >= 48 And intAsc <= 57) Then
            If Len(strOut) = 0 And (intAsc >= 48 And intAsc <= 57) Then
                strOut = "Col" & strCh
                blnCapNext = False
            ElseIf blnCapNext Then
                strOut = strOut & UCase$(strCh)
                blnCapNext = False
            Else
                strOut = strOut & strCh
            End If
        Else
            blnCapNext = True
        End If
    Next lngIdx

    If Len(strOut) = 0 Then strOut = "FieldVal"
    If blnCheckReserved Then
        If InStr(1, RESERVED_IDENTIFIERS, "|" & UCase$(strOut) & "|", vbBinaryCompare) > 0 Then
            strOut = strOut & "Value"
        End If
    End If
    SanitizePascalIdentifier = strOut
End Function

