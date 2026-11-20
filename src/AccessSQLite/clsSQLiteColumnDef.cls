Attribute VB_Name = "clsSQLiteColumnDef"
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
' clsSQLiteColumnDef
' Metadata definition for a single column in a clsSQLiteTableMap.
' =============================================================================

Private m_ColumnName As String
Private m_ColumnKind As SQLiteColumnKind
Private m_LabelText As String
Private m_IsPrimaryKey As Boolean
Private m_IsAutoIncrement As Boolean
Private m_IsRequired As Boolean
Private m_IsUnique As Boolean
Private m_MaxLength As Long
Private m_HasDefault As Boolean
Private m_DefaultValue As Variant
Private m_DefaultRawSql As String
Private m_FkTable As String
Private m_FkColumn As String
Private m_OnDeleteAction As String
Private m_IsClientEditable As Boolean
Private m_IsAudited As Boolean

Private Sub Class_Initialize()
    m_ColumnKind = sqColText
    m_IsClientEditable = True
    m_IsAudited = True
End Sub

Public Property Get ColumnName() As String
    ColumnName = m_ColumnName
End Property
Public Property Let ColumnName(ByVal strVal As String)
    m_ColumnName = Trim$(strVal)
    If Len(m_LabelText) = 0 Then m_LabelText = m_ColumnName
End Property

Public Property Get ColumnKind() As SQLiteColumnKind
    ColumnKind = m_ColumnKind
End Property
Public Property Let ColumnKind(ByVal enmVal As SQLiteColumnKind)
    m_ColumnKind = enmVal
End Property

Public Property Get LabelText() As String
    LabelText = m_LabelText
End Property
Public Property Let LabelText(ByVal strVal As String)
    m_LabelText = strVal
End Property

Public Property Get IsPrimaryKey() As Boolean
    IsPrimaryKey = m_IsPrimaryKey
End Property
Public Property Let IsPrimaryKey(ByVal blnVal As Boolean)
    m_IsPrimaryKey = blnVal
End Property

Public Property Get IsAutoIncrement() As Boolean
    IsAutoIncrement = m_IsAutoIncrement
End Property
Public Property Let IsAutoIncrement(ByVal blnVal As Boolean)
    m_IsAutoIncrement = blnVal
End Property

Public Property Get IsRequired() As Boolean
    IsRequired = m_IsRequired
End Property
Public Property Let IsRequired(ByVal blnVal As Boolean)
    m_IsRequired = blnVal
End Property

Public Property Get IsUnique() As Boolean
    IsUnique = m_IsUnique
End Property
Public Property Let IsUnique(ByVal blnVal As Boolean)
    m_IsUnique = blnVal
End Property

Public Property Get MaxLength() As Long
    MaxLength = m_MaxLength
End Property
Public Property Let MaxLength(ByVal lngVal As Long)
    m_MaxLength = lngVal
End Property

Public Property Get HasDefault() As Boolean
    HasDefault = m_HasDefault
End Property
Public Property Let HasDefault(ByVal blnVal As Boolean)
    m_HasDefault = blnVal
End Property

Public Property Get DefaultValue() As Variant
    DefaultValue = m_DefaultValue
End Property
Public Property Let DefaultValue(ByVal varVal As Variant)
    m_DefaultValue = varVal
    m_HasDefault = True
End Property

Public Property Get DefaultRawSql() As String
    DefaultRawSql = m_DefaultRawSql
End Property
Public Property Let DefaultRawSql(ByVal strVal As String)
    m_DefaultRawSql = Trim$(strVal)
End Property

Public Property Get ForeignKeyTable() As String
    ForeignKeyTable = m_FkTable
End Property
Public Property Let ForeignKeyTable(ByVal strVal As String)
    m_FkTable = Trim$(strVal)
End Property

Public Property Get ForeignKeyColumn() As String
    ForeignKeyColumn = m_FkColumn
End Property
Public Property Let ForeignKeyColumn(ByVal strVal As String)
    m_FkColumn = Trim$(strVal)
End Property

Public Property Get OnDeleteAction() As String
    OnDeleteAction = m_OnDeleteAction
End Property
Public Property Let OnDeleteAction(ByVal strVal As String)
    m_OnDeleteAction = UCase$(Trim$(strVal))
End Property

Public Property Get IsClientEditable() As Boolean
    IsClientEditable = m_IsClientEditable
End Property
Public Property Let IsClientEditable(ByVal blnVal As Boolean)
    m_IsClientEditable = blnVal
End Property

Public Property Get IsAudited() As Boolean
    IsAudited = m_IsAudited
End Property
Public Property Let IsAudited(ByVal blnVal As Boolean)
    m_IsAudited = blnVal
End Property

' Maps logical ORM column kind to SQLite DDL type declaration (supports STRICT mode).
Public Function SqliteTypeDeclaration(Optional ByVal blnStrictMode As Boolean = False) As String
    If m_IsPrimaryKey And m_IsAutoIncrement Then
        SqliteTypeDeclaration = "INTEGER"
        Exit Function
    End If

    If blnStrictMode Then
        Select Case m_ColumnKind
            Case sqColLong, sqColInt64, sqColBoolean
                SqliteTypeDeclaration = "INTEGER"
            Case sqColDouble, sqColCurrency
                SqliteTypeDeclaration = "REAL"
            Case sqColBlob
                SqliteTypeDeclaration = "BLOB"
            Case Else
                SqliteTypeDeclaration = "TEXT"
        End Select
    Else
        Select Case m_ColumnKind
            Case sqColLong, sqColInt64
                SqliteTypeDeclaration = "INTEGER"
            Case sqColBoolean
                SqliteTypeDeclaration = "BOOLEAN"
            Case sqColDouble
                SqliteTypeDeclaration = "REAL"
            Case sqColCurrency
                SqliteTypeDeclaration = "CURRENCY"
            Case sqColDate
                SqliteTypeDeclaration = "DATE"
            Case sqColDateTime
                SqliteTypeDeclaration = "DATETIME"
            Case sqColBlob
                SqliteTypeDeclaration = "BLOB"
            Case sqColJson
                SqliteTypeDeclaration = "JSON"
            Case sqColGuid
                SqliteTypeDeclaration = "GUID"
            Case Else
                SqliteTypeDeclaration = "TEXT"
        End Select
    End If
End Function

' Formats the SQL DEFAULT literal for DDL generation.
Public Function FormatDefaultSqlLiteral() As String
    If Len(m_DefaultRawSql) > 0 Then
        FormatDefaultSqlLiteral = m_DefaultRawSql
        Exit Function
    End If
    If Not m_HasDefault Then Exit Function
    If IsNull(m_DefaultValue) Or IsEmpty(m_DefaultValue) Then
        FormatDefaultSqlLiteral = "NULL"
        Exit Function
    End If

    Select Case m_ColumnKind
        Case sqColBoolean
            FormatDefaultSqlLiteral = IIf(CBool(m_DefaultValue), "1", "0")
        Case sqColLong, sqColInt64
            FormatDefaultSqlLiteral = CStr(m_DefaultValue)
        Case sqColDouble, sqColCurrency
            FormatDefaultSqlLiteral = Replace(CStr(CDbl(m_DefaultValue)), ",", ".")
        Case sqColDate, sqColDateTime
            If VarType(m_DefaultValue) = vbDate Then
                FormatDefaultSqlLiteral = "'" & Replace(FormatIsoDateValue(CDate(m_DefaultValue), (m_ColumnKind = sqColDateTime)), "'", "''") & "'"
            Else
                FormatDefaultSqlLiteral = "'" & Replace(CStr(m_DefaultValue), "'", "''") & "'"
            End If
        Case Else
            FormatDefaultSqlLiteral = "'" & Replace(CStr(m_DefaultValue), "'", "''") & "'"
    End Select
End Function
