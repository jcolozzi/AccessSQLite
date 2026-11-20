Attribute VB_Name = "clsDemoTask"
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
' clsDemoTask
' Strongly-typed example ORM entity implementing ISQLiteEntity for DemoTasks.
' =============================================================================

Implements ISQLiteEntity

Private m_TaskID As Long
Private m_ProjectID As Long
Private m_TitleText As String
Private m_PriorityRank As Long
Private m_EstimateCost As Currency
Private m_IsCompleted As Boolean
Private m_DueDate As Variant
Private m_CreatedAt As Variant
Private m_UpdatedAt As Variant
Private m_DeletedAt As Variant

Private Sub Class_Initialize()
    m_PriorityRank = 3
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

Public Property Get ProjectID() As Long
    ProjectID = m_ProjectID
End Property
Public Property Let ProjectID(ByVal lngVal As Long)
    m_ProjectID = lngVal
End Property

Public Property Get TitleText() As String
    TitleText = m_TitleText
End Property
Public Property Let TitleText(ByVal strVal As String)
    m_TitleText = strVal
End Property

Public Property Get PriorityRank() As Long
    PriorityRank = m_PriorityRank
End Property
Public Property Let PriorityRank(ByVal lngVal As Long)
    m_PriorityRank = lngVal
End Property

Public Property Get EstimateCost() As Currency
    EstimateCost = m_EstimateCost
End Property
Public Property Let EstimateCost(ByVal curVal As Currency)
    m_EstimateCost = curVal
End Property

Public Property Get IsCompleted() As Boolean
    IsCompleted = m_IsCompleted
End Property
Public Property Let IsCompleted(ByVal blnVal As Boolean)
    m_IsCompleted = blnVal
End Property

Public Property Get DueDate() As Variant
    DueDate = m_DueDate
End Property
Public Property Let DueDate(ByVal varVal As Variant)
    m_DueDate = varVal
End Property

Public Property Get CreatedAt() As Variant
    CreatedAt = m_CreatedAt
End Property
Public Property Let CreatedAt(ByVal varVal As Variant)
    m_CreatedAt = varVal
End Property

Public Property Get UpdatedAt() As Variant
    UpdatedAt = m_UpdatedAt
End Property
Public Property Let UpdatedAt(ByVal varVal As Variant)
    m_UpdatedAt = varVal
End Property

Public Property Get DeletedAt() As Variant
    DeletedAt = m_DeletedAt
End Property
Public Property Let DeletedAt(ByVal varVal As Variant)
    m_DeletedAt = varVal
End Property

Private Property Get ISQLiteEntity_EntityTableName() As String
    ISQLiteEntity_EntityTableName = "DemoTasks"
End Property

Private Function ISQLiteEntity_CreateBlank() As ISQLiteEntity
    Set ISQLiteEntity_CreateBlank = New clsDemoTask
End Function

Private Sub ISQLiteEntity_ReadFromRow(ByVal dictRow As Scripting.Dictionary)
    If dictRow Is Nothing Then Exit Sub
    If dictRow.Exists("TaskID") Then m_TaskID = CLng(Nz(dictRow("TaskID"), 0))
    If dictRow.Exists("ProjectID") Then m_ProjectID = CLng(Nz(dictRow("ProjectID"), 0))
    If dictRow.Exists("TitleText") Then m_TitleText = CStr(Nz(dictRow("TitleText"), vbNullString))
    If dictRow.Exists("PriorityRank") Then m_PriorityRank = CLng(Nz(dictRow("PriorityRank"), 3))
    If dictRow.Exists("EstimateCost") Then m_EstimateCost = CCur(Nz(dictRow("EstimateCost"), 0@))
    If dictRow.Exists("IsCompleted") Then m_IsCompleted = CBool(Nz(dictRow("IsCompleted"), False))
    If dictRow.Exists("DueDate") Then m_DueDate = dictRow("DueDate")
    If dictRow.Exists("CreatedAt") Then m_CreatedAt = dictRow("CreatedAt")
    If dictRow.Exists("UpdatedAt") Then m_UpdatedAt = dictRow("UpdatedAt")
    If dictRow.Exists("DeletedAt") Then m_DeletedAt = dictRow("DeletedAt")
End Sub

Private Sub ISQLiteEntity_WriteToRow(ByVal dictRow As Scripting.Dictionary)
    If dictRow Is Nothing Then Exit Sub
    dictRow("TaskID") = m_TaskID
    dictRow("ProjectID") = m_ProjectID
    dictRow("TitleText") = m_TitleText
    dictRow("PriorityRank") = m_PriorityRank
    dictRow("EstimateCost") = m_EstimateCost
    dictRow("IsCompleted") = m_IsCompleted
    dictRow("DueDate") = m_DueDate
End Sub
