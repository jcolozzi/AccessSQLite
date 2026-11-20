Attribute VB_Name = "clsTaskItem"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

' =============================================================================
' clsTaskItem
' Strongly-typed ORM entity in AccessSQLiteDemo implementing ISQLiteEntity
' from the referenced AccessSQLite library. Used by frmTaskDetail.
' =============================================================================

Implements AccessSQLite.ISQLiteEntity

Private m_TaskID As Long
Private m_TaskTitle As String
Private m_CategoryName As String
Private m_PriorityLevel As String
Private m_EstimateCost As Currency
Private m_DueDate As Variant
Private m_IsCompleted As Boolean
Private m_TaskNotes As String
Private m_CreatedAt As Variant
Private m_UpdatedAt As Variant
Private m_DeletedAt As Variant

Private Sub Class_Initialize()
    m_CategoryName = "General"
    m_PriorityLevel = "Medium"
    m_EstimateCost = 0@
    m_DueDate = Null
    m_CreatedAt = Null
    m_UpdatedAt = Null
    m_DeletedAt = Null
End Sub

Public Property Get TaskID() As Long
    TaskID = m_TaskID
End Property
Public Property Let TaskID(ByVal lngVal As Long)
    m_TaskID = lngVal
End Property

Public Property Get TaskTitle() As String
    TaskTitle = m_TaskTitle
End Property
Public Property Let TaskTitle(ByVal strVal As String)
    m_TaskTitle = strVal
End Property

Public Property Get CategoryName() As String
    CategoryName = m_CategoryName
End Property
Public Property Let CategoryName(ByVal strVal As String)
    m_CategoryName = strVal
End Property

Public Property Get PriorityLevel() As String
    PriorityLevel = m_PriorityLevel
End Property
Public Property Let PriorityLevel(ByVal strVal As String)
    m_PriorityLevel = strVal
End Property

Public Property Get EstimateCost() As Currency
    EstimateCost = m_EstimateCost
End Property
Public Property Let EstimateCost(ByVal curVal As Currency)
    m_EstimateCost = curVal
End Property

Public Property Get DueDate() As Variant
    DueDate = m_DueDate
End Property
Public Property Let DueDate(ByVal varVal As Variant)
    m_DueDate = varVal
End Property

Public Property Get IsCompleted() As Boolean
    IsCompleted = m_IsCompleted
End Property
Public Property Let IsCompleted(ByVal blnVal As Boolean)
    m_IsCompleted = blnVal
End Property

Public Property Get TaskNotes() As String
    TaskNotes = m_TaskNotes
End Property
Public Property Let TaskNotes(ByVal strVal As String)
    m_TaskNotes = strVal
End Property

Public Property Get CreatedAt() As Variant
    CreatedAt = m_CreatedAt
End Property

Public Property Get UpdatedAt() As Variant
    UpdatedAt = m_UpdatedAt
End Property

Public Property Get DeletedAt() As Variant
    DeletedAt = m_DeletedAt
End Property

Private Property Get ISQLiteEntity_EntityTableName() As String
    ISQLiteEntity_EntityTableName = "Tasks"
End Property

Private Function ISQLiteEntity_CreateBlank() As AccessSQLite.ISQLiteEntity
    Set ISQLiteEntity_CreateBlank = New clsTaskItem
End Function

Private Sub ISQLiteEntity_ReadFromRow(ByVal dictRow As Scripting.Dictionary)
    If dictRow Is Nothing Then Exit Sub
    If dictRow.Exists("TaskID") Then m_TaskID = CLng(Nz(dictRow("TaskID"), 0))
    If dictRow.Exists("TaskTitle") Then m_TaskTitle = CStr(Nz(dictRow("TaskTitle"), vbNullString))
    If dictRow.Exists("CategoryName") Then m_CategoryName = CStr(Nz(dictRow("CategoryName"), "General"))
    If dictRow.Exists("PriorityLevel") Then m_PriorityLevel = CStr(Nz(dictRow("PriorityLevel"), "Medium"))
    If dictRow.Exists("EstimateCost") Then m_EstimateCost = CCur(Nz(dictRow("EstimateCost"), 0@))
    If dictRow.Exists("DueDate") Then m_DueDate = dictRow("DueDate")
    If dictRow.Exists("IsCompleted") Then m_IsCompleted = CBool(Nz(dictRow("IsCompleted"), False))
    If dictRow.Exists("TaskNotes") Then m_TaskNotes = CStr(Nz(dictRow("TaskNotes"), vbNullString))
    If dictRow.Exists("CreatedAt") Then m_CreatedAt = dictRow("CreatedAt")
    If dictRow.Exists("UpdatedAt") Then m_UpdatedAt = dictRow("UpdatedAt")
    If dictRow.Exists("DeletedAt") Then m_DeletedAt = dictRow("DeletedAt")
End Sub

Private Sub ISQLiteEntity_WriteToRow(ByVal dictRow As Scripting.Dictionary)
    If dictRow Is Nothing Then Exit Sub
    dictRow("TaskID") = m_TaskID
    dictRow("TaskTitle") = m_TaskTitle
    dictRow("CategoryName") = m_CategoryName
    dictRow("PriorityLevel") = m_PriorityLevel
    dictRow("EstimateCost") = m_EstimateCost
    dictRow("DueDate") = m_DueDate
    dictRow("IsCompleted") = m_IsCompleted
    dictRow("TaskNotes") = m_TaskNotes
End Sub
