Attribute VB_Name = "ISQLiteEntity"
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
' ISQLiteEntity
' Strongly-typed ORM entity contract with prototype factory (CreateBlank)
' enabling generic collection materialization in VBA without COM registration.
' =============================================================================

Public Property Get EntityTableName() As String
End Property

Public Function CreateBlank() As ISQLiteEntity
End Function

Public Sub ReadFromRow(ByVal dictRow As Scripting.Dictionary)
End Sub

Public Sub WriteToRow(ByVal dictRow As Scripting.Dictionary)
End Sub
