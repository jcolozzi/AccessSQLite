Attribute VB_Name = "clsSQLiteSchema"
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
' clsSQLiteSchema
' DDL generator (CREATE TABLE [STRICT], CREATE INDEX, additive ALTER TABLE)
' and PRAGMA reverse-reflection into clsSQLiteTableMap.
' =============================================================================

Private m_Conn As clsSQLiteConnection

Public Sub InitSchema(ByVal objConn As clsSQLiteConnection)
    Set m_Conn = objConn
End Sub

Public Function BuildCreateTableSql( _
    ByVal objMap As clsSQLiteTableMap, _
    Optional ByVal blnIfNotExists As Boolean = True) As String

    Dim strSql As String
    Dim colParts As Collection
    Dim objCol As clsSQLiteColumnDef
    Dim strPart As String
    Dim strDefLit As String
    Dim lngIdx As Long

    If objMap Is Nothing Or objMap.Columns.Count = 0 Then
        RaiseSQLiteError sqErrInvalidArgument, "clsSQLiteSchema.BuildCreateTableSql", "TableMap must have at least one column."
    End If

    Set colParts = New Collection

    For Each objCol In objMap.Columns
        strPart = "    " & QuoteSQLiteIdent(objCol.ColumnName) & " " & objCol.SqliteTypeDeclaration(objMap.UseStrictTable)
        If objCol.IsPrimaryKey Then
            strPart = strPart & " PRIMARY KEY"
            If objCol.IsAutoIncrement And objCol.SqliteTypeDeclaration(objMap.UseStrictTable) = "INTEGER" Then
                strPart = strPart & " AUTOINCREMENT"
            End If
        Else
            If objCol.IsRequired Then strPart = strPart & " NOT NULL"
            If objCol.IsUnique Then strPart = strPart & " UNIQUE"
            strDefLit = objCol.FormatDefaultSqlLiteral()
            If Len(strDefLit) > 0 Then
                strPart = strPart & " DEFAULT " & strDefLit
            End If
        End If

        If Len(objCol.ForeignKeyTable) > 0 Then
            strPart = strPart & " REFERENCES " & QuoteSQLiteIdent(objCol.ForeignKeyTable)
            If Len(objCol.ForeignKeyColumn) > 0 Then
                strPart = strPart & "(" & QuoteSQLiteIdent(objCol.ForeignKeyColumn) & ")"
            End If
            If Len(objCol.OnDeleteAction) > 0 Then
                strPart = strPart & " ON DELETE " & objCol.OnDeleteAction
            End If
        End If

        colParts.Add strPart
    Next objCol

    strSql = "CREATE TABLE "
    If blnIfNotExists Then strSql = strSql & "IF NOT EXISTS "
    strSql = strSql & QuoteSQLiteIdent(objMap.TableName) & " (" & vbCrLf

    For lngIdx = 1 To colParts.Count
        strSql = strSql & CStr(colParts(lngIdx))
        If lngIdx < colParts.Count Then strSql = strSql & ","
        strSql = strSql & vbCrLf
    Next lngIdx

    strSql = strSql & ")"
    If objMap.UseStrictTable Then strSql = strSql & " STRICT"
    strSql = strSql & ";"
    BuildCreateTableSql = strSql
End Function

Public Function BuildCreateIndexesSql(ByVal objMap As clsSQLiteTableMap) As Collection
    Dim colSql As Collection
    Dim dictIdx As Scripting.Dictionary
    Dim astrCols() As String
    Dim strColList As String
    Dim lngPos As Long
    Dim strStmt As String
    Dim strColExpr As String
    Dim strOrderSuffix As String
    Dim lngSpace As Long

    Set colSql = New Collection
    If objMap Is Nothing Then
        Set BuildCreateIndexesSql = colSql
        Exit Function
    End If

    For Each dictIdx In objMap.Indexes
        astrCols = Split(CStr(dictIdx("ColumnsCsv")), ",")
        strColList = vbNullString
        For lngPos = LBound(astrCols) To UBound(astrCols)
            strColExpr = Trim$(astrCols(lngPos))
            strOrderSuffix = vbNullString
            lngSpace = InStr(1, strColExpr, " ", vbBinaryCompare)
            If lngSpace > 0 Then
                strOrderSuffix = " " & UCase$(Trim$(Mid$(strColExpr, lngSpace + 1)))
                strColExpr = Left$(strColExpr, lngSpace - 1)
            End If
            If Len(strColList) > 0 Then strColList = strColList & ", "
            strColList = strColList & QuoteSQLiteIdent(strColExpr) & strOrderSuffix
        Next lngPos

        strStmt = "CREATE "
        If CBool(dictIdx("IsUnique")) Then strStmt = strStmt & "UNIQUE "
        strStmt = strStmt & "INDEX IF NOT EXISTS " & QuoteSQLiteIdent(CStr(dictIdx("IndexName"))) & _
                  " ON " & QuoteSQLiteIdent(objMap.TableName) & " (" & strColList & ")"
        If Len(CStr(dictIdx("WhereSql"))) > 0 Then
            strStmt = strStmt & " WHERE " & CStr(dictIdx("WhereSql"))
        End If
        strStmt = strStmt & ";"
        colSql.Add strStmt
    Next dictIdx

    Set BuildCreateIndexesSql = colSql
End Function

' Creates the table and indexes if absent, and optionally adds any newly mapped columns.
Public Sub EnsureTable( _
    ByVal objMap As clsSQLiteTableMap, _
    Optional ByVal blnSyncAdditiveColumns As Boolean = True)

    Dim colIdxSql As Collection
    Dim varStmt As Variant

    EnsureConn
    If Not m_Conn.TableExists(objMap.TableName) Then
        m_Conn.ExecuteParams BuildCreateTableSql(objMap, True), Empty
    ElseIf blnSyncAdditiveColumns Then
        SyncAdditiveColumns objMap
    End If

    Set colIdxSql = BuildCreateIndexesSql(objMap)
    For Each varStmt In colIdxSql
        m_Conn.ExecuteParams CStr(varStmt), Empty
    Next varStmt

    m_Conn.RegisterTableMap objMap
End Sub

' Inspects PRAGMA table_xinfo and adds any columns present in objMap but missing in SQLite.
Public Function SyncAdditiveColumns(ByVal objMap As clsSQLiteTableMap) As Long
    Dim colInfo As Collection
    Dim dictExisting As Scripting.Dictionary
    Dim dictColRow As Scripting.Dictionary
    Dim objCol As clsSQLiteColumnDef
    Dim strAlter As String
    Dim strDef As String
    Dim lngAdded As Long

    EnsureConn
    Set colInfo = m_Conn.QueryRows("PRAGMA table_xinfo(" & QuoteSQLiteIdent(objMap.TableName) & ");")
    Set dictExisting = NewTextDict()
    For Each dictColRow In colInfo
        dictExisting(CStr(dictColRow("name"))) = True
    Next dictColRow

    For Each objCol In objMap.Columns
        If Not dictExisting.Exists(objCol.ColumnName) Then
            strAlter = "ALTER TABLE " & QuoteSQLiteIdent(objMap.TableName) & _
                       " ADD COLUMN " & QuoteSQLiteIdent(objCol.ColumnName) & " " & _
                       objCol.SqliteTypeDeclaration(objMap.UseStrictTable)
            strDef = objCol.FormatDefaultSqlLiteral()
            If objCol.IsRequired Then
                If Len(strDef) = 0 Then strDef = "''"
                strAlter = strAlter & " NOT NULL DEFAULT " & strDef
            ElseIf Len(strDef) > 0 Then
                strAlter = strAlter & " DEFAULT " & strDef
            End If
            strAlter = strAlter & ";"
            m_Conn.ExecuteParams strAlter, Empty
            lngAdded = lngAdded + 1
        End If
    Next objCol

    SyncAdditiveColumns = lngAdded
End Function

Public Sub DropTableIfExists(ByVal strTableName As String)
    EnsureConn
    m_Conn.ExecuteParams "DROP TABLE IF EXISTS " & QuoteSQLiteIdent(strTableName) & ";", Empty
End Sub

Public Function ListTables(Optional ByVal blnIncludeViews As Boolean = False) As Collection
    Dim strFilter As String
    Dim colRows As Collection
    Dim dictRow As Scripting.Dictionary
    Dim colNames As Collection

    EnsureConn
    If blnIncludeViews Then
        strFilter = "type IN ('table','view')"
    Else
        strFilter = "type = 'table'"
    End If

    Set colRows = m_Conn.QueryRows( _
        "SELECT name FROM sqlite_master WHERE " & strFilter & " AND name NOT LIKE 'sqlite_%' ORDER BY name;")
    Set colNames = New Collection
    For Each dictRow In colRows
        colNames.Add CStr(dictRow("name"))
    Next dictRow
    Set ListTables = colNames
End Function

' Reverse-engineers an existing SQLite table into a clsSQLiteTableMap using PRAGMA metadata.
Public Function ReflectTableMap(ByVal strTableName As String) As clsSQLiteTableMap
    Dim colInfo As Collection
    Dim colFks As Collection
    Dim dictFks As Scripting.Dictionary
    Dim dictRow As Scripting.Dictionary
    Dim objMap As clsSQLiteTableMap
    Dim strColName As String
    Dim strDeclType As String
    Dim blnPk As Boolean
    Dim blnNotNull As Boolean
    Dim blnAutoInc As Boolean
    Dim strCreateSql As String
    Dim dictFkInfo As Scripting.Dictionary

    EnsureConn
    If Not m_Conn.TableExists(strTableName) Then
        RaiseSQLiteError sqErrNotFound, "clsSQLiteSchema.ReflectTableMap", _
            "SQLite table or view '" & strTableName & "' does not exist."
    End If

    strCreateSql = CStr(Nz(m_Conn.QueryScalar( _
        "SELECT sql FROM sqlite_master WHERE LOWER(name) = LOWER(?);", Trim$(strTableName)), vbNullString))

    Set colFks = m_Conn.QueryRows("PRAGMA foreign_key_list(" & QuoteSQLiteIdent(strTableName) & ");")
    Set dictFks = NewTextDict()
    For Each dictRow In colFks
        Set dictFkInfo = NewTextDict()
        dictFkInfo("table") = CStr(Nz(dictRow("table"), vbNullString))
        dictFkInfo("to") = CStr(Nz(dictRow("to"), vbNullString))
        dictFkInfo("on_delete") = CStr(Nz(dictRow("on_delete"), vbNullString))
        Set dictFks(CStr(dictRow("from"))) = dictFkInfo
    Next dictRow

    Set colInfo = m_Conn.QueryRows("PRAGMA table_xinfo(" & QuoteSQLiteIdent(strTableName) & ");")
    Set objMap = New clsSQLiteTableMap
    objMap.InitMap Trim$(strTableName), "rowid"

    If InStr(1, UCase$(strCreateSql), ") STRICT", vbBinaryCompare) > 0 Then
        objMap.SetStrict True
    End If

    For Each dictRow In colInfo
        If CLng(Nz(dictRow("hidden"), 0)) = 0 Then
            strColName = CStr(dictRow("name"))
            strDeclType = UCase$(Trim$(CStr(Nz(dictRow("type"), vbNullString))))
            blnPk = (CLng(Nz(dictRow("pk"), 0)) > 0)
            blnNotNull = (CLng(Nz(dictRow("notnull"), 0)) <> 0)
            blnAutoInc = (blnPk And InStr(1, strDeclType, "INT", vbBinaryCompare) > 0)

            If dictFks.Exists(strColName) Then
                Set dictFkInfo = dictFks(strColName)
                objMap.AddColumn strColName, InferColumnKind(strDeclType), strColName, _
                    blnPk, blnAutoInc, blnNotNull, 0, False, , _
                    CStr(dictFkInfo("table")), CStr(dictFkInfo("to")), CStr(dictFkInfo("on_delete")), _
                    Not blnAutoInc
            Else
                objMap.AddColumn strColName, InferColumnKind(strDeclType), strColName, _
                    blnPk, blnAutoInc, blnNotNull, 0, False, , , , , Not blnAutoInc
            End If

            If StrComp(strColName, "CreatedAt", vbTextCompare) = 0 Then
                objMap.AddTimestamps "CreatedAt", vbNullString
            ElseIf StrComp(strColName, "UpdatedAt", vbTextCompare) = 0 Then
                objMap.AddTimestamps vbNullString, "UpdatedAt"
            ElseIf StrComp(strColName, "DeletedAt", vbTextCompare) = 0 Then
                objMap.AddSoftDelete "DeletedAt"
            End If
        End If
    Next dictRow

    Set ReflectTableMap = objMap
End Function

Private Function InferColumnKind(ByVal strUpperDeclType As String) As SQLiteColumnKind
    If InStr(1, strUpperDeclType, "BOOL", vbBinaryCompare) > 0 Or strUpperDeclType = "BIT" Then
        InferColumnKind = sqColBoolean
    ElseIf InStr(1, strUpperDeclType, "BIGINT", vbBinaryCompare) > 0 Or strUpperDeclType = "INT64" Then
        InferColumnKind = sqColInt64
    ElseIf InStr(1, strUpperDeclType, "INT", vbBinaryCompare) > 0 Then
        InferColumnKind = sqColInt64
    ElseIf InStr(1, strUpperDeclType, "CURR", vbBinaryCompare) > 0 Or InStr(1, strUpperDeclType, "MONEY", vbBinaryCompare) > 0 Or InStr(1, strUpperDeclType, "DECIMAL", vbBinaryCompare) > 0 Or InStr(1, strUpperDeclType, "NUMERIC", vbBinaryCompare) > 0 Then
        InferColumnKind = sqColCurrency
    ElseIf InStr(1, strUpperDeclType, "REAL", vbBinaryCompare) > 0 Or InStr(1, strUpperDeclType, "FLOA", vbBinaryCompare) > 0 Or InStr(1, strUpperDeclType, "DOUB", vbBinaryCompare) > 0 Then
        InferColumnKind = sqColDouble
    ElseIf InStr(1, strUpperDeclType, "DATETIME", vbBinaryCompare) > 0 Or InStr(1, strUpperDeclType, "TIMESTAMP", vbBinaryCompare) > 0 Then
        InferColumnKind = sqColDateTime
    ElseIf InStr(1, strUpperDeclType, "DATE", vbBinaryCompare) > 0 Then
        InferColumnKind = sqColDate
    ElseIf InStr(1, strUpperDeclType, "BLOB", vbBinaryCompare) > 0 Or InStr(1, strUpperDeclType, "BIN", vbBinaryCompare) > 0 Then
        InferColumnKind = sqColBlob
    ElseIf InStr(1, strUpperDeclType, "JSON", vbBinaryCompare) > 0 Then
        InferColumnKind = sqColJson
    ElseIf InStr(1, strUpperDeclType, "GUID", vbBinaryCompare) > 0 Or InStr(1, strUpperDeclType, "UUID", vbBinaryCompare) > 0 Then
        InferColumnKind = sqColGuid
    ElseIf InStr(1, strUpperDeclType, "MEMO", vbBinaryCompare) > 0 Or InStr(1, strUpperDeclType, "CLOB", vbBinaryCompare) > 0 Then
        InferColumnKind = sqColMemo
    Else
        InferColumnKind = sqColText
    End If
End Function

Private Sub EnsureConn()
    If m_Conn Is Nothing Then
        RaiseSQLiteError sqErrConnectionClosed, "clsSQLiteSchema", "Schema manager is not attached to a clsSQLiteConnection."
    End If
End Sub

