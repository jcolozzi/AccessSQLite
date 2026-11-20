Attribute VB_Name = "clsSQLiteUnitOfWork"
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
' clsSQLiteUnitOfWork
' Atomic multi-repository batch unit-of-work backed by SAVEPOINT / transaction.
' =============================================================================

Private Enum UowOpKind
    uowInsertDict = 1
    uowUpdateDict = 2
    uowDeleteById = 3
    uowSaveEntity = 4
    uowDeleteEntity = 5
    uowExecSql = 6
End Enum

Private m_Conn As clsSQLiteConnection
Private m_Operations As Collection

Private Sub Class_Initialize()
    Set m_Operations = New Collection
End Sub

Public Sub InitUnitOfWork(ByVal objConn As clsSQLiteConnection)
    Set m_Conn = objConn
    Set m_Operations = New Collection
End Sub

Public Property Get PendingCount() As Long
    PendingCount = m_Operations.Count
End Property

Public Sub ClearPending()
    Set m_Operations = New Collection
End Sub

Public Sub RegisterInsert(ByVal strTableName As String, ByVal dictValues As Scripting.Dictionary)
    Dim dictOp As Scripting.Dictionary
    Set dictOp = NewTextDict()
    dictOp("OpKind") = uowInsertDict
    dictOp("TableName") = Trim$(strTableName)
    Set dictOp("Payload") = dictValues
    m_Operations.Add dictOp
End Sub

Public Sub RegisterUpdate( _
    ByVal strTableName As String, _
    ByVal varId As Variant, _
    ByVal dictValues As Scripting.Dictionary, _
    Optional ByVal blnDirtyOnly As Boolean = False)

    Dim dictOp As Scripting.Dictionary
    Set dictOp = NewTextDict()
    dictOp("OpKind") = uowUpdateDict
    dictOp("TableName") = Trim$(strTableName)
    dictOp("RecordId") = varId
    dictOp("DirtyOnly") = blnDirtyOnly
    Set dictOp("Payload") = dictValues
    m_Operations.Add dictOp
End Sub

Public Sub RegisterDelete( _
    ByVal strTableName As String, _
    ByVal varId As Variant, _
    Optional ByVal blnForceHardDelete As Boolean = False)

    Dim dictOp As Scripting.Dictionary
    Set dictOp = NewTextDict()
    dictOp("OpKind") = uowDeleteById
    dictOp("TableName") = Trim$(strTableName)
    dictOp("RecordId") = varId
    dictOp("ForceHard") = blnForceHardDelete
    m_Operations.Add dictOp
End Sub

Public Sub RegisterEntitySave(ByVal objEntity As ISQLiteEntity, Optional ByVal blnDirtyOnly As Boolean = True)
    Dim dictOp As Scripting.Dictionary
    Set dictOp = NewTextDict()
    dictOp("OpKind") = uowSaveEntity
    dictOp("DirtyOnly") = blnDirtyOnly
    Set dictOp("Entity") = objEntity
    m_Operations.Add dictOp
End Sub

Public Sub RegisterEntityDelete(ByVal objEntity As ISQLiteEntity, Optional ByVal blnForceHardDelete As Boolean = False)
    Dim dictOp As Scripting.Dictionary
    Set dictOp = NewTextDict()
    dictOp("OpKind") = uowDeleteEntity
    dictOp("ForceHard") = blnForceHardDelete
    Set dictOp("Entity") = objEntity
    m_Operations.Add dictOp
End Sub

Public Sub RegisterSql(ByVal strSql As String, ParamArray varParams() As Variant)
    Dim dictOp As Scripting.Dictionary
    Dim varPassed As Variant
    varPassed = varParams
    Set dictOp = NewTextDict()
    dictOp("OpKind") = uowExecSql
    dictOp("SqlText") = strSql
    dictOp("Params") = varPassed
    m_Operations.Add dictOp
End Sub

' Executes all queued operations atomically inside a transaction/savepoint and clears queue on success.
Public Function CommitWork() As Long
    Dim dictOp As Scripting.Dictionary
    Dim objEntity As ISQLiteEntity
    Dim objRepo As clsSQLiteRepository
    Dim lngExecuted As Long

    If m_Conn Is Nothing Then
        RaiseSQLiteError sqErrConnectionClosed, "clsSQLiteUnitOfWork.CommitWork", "UnitOfWork is not attached to a clsSQLiteConnection."
    End If
    If m_Operations.Count = 0 Then Exit Function

    m_Conn.BeginTrans
    On Error GoTo WorkFailed

    For Each dictOp In m_Operations
        Select Case CLng(dictOp("OpKind"))
            Case uowInsertDict
                Set objRepo = m_Conn.Repository(CStr(dictOp("TableName")))
                objRepo.InsertRecord dictOp("Payload"), False

            Case uowUpdateDict
                Set objRepo = m_Conn.Repository(CStr(dictOp("TableName")))
                objRepo.UpdateRecord dictOp("RecordId"), dictOp("Payload"), CBool(dictOp("DirtyOnly")), False

            Case uowDeleteById
                Set objRepo = m_Conn.Repository(CStr(dictOp("TableName")))
                objRepo.DeleteRecord dictOp("RecordId"), CBool(dictOp("ForceHard"))

            Case uowSaveEntity
                Set objEntity = dictOp("Entity")
                Set objRepo = m_Conn.Repository(objEntity.EntityTableName)
                objRepo.SaveEntity objEntity, CBool(dictOp("DirtyOnly"))

            Case uowDeleteEntity
                Set objEntity = dictOp("Entity")
                Set objRepo = m_Conn.Repository(objEntity.EntityTableName)
                objRepo.DeleteEntity objEntity, CBool(dictOp("ForceHard"))

            Case uowExecSql
                m_Conn.ExecuteParams CStr(dictOp("SqlText")), dictOp("Params")
        End Select
        lngExecuted = lngExecuted + 1
    Next dictOp

    m_Conn.CommitTrans
    Set m_Operations = New Collection
    CommitWork = lngExecuted
    Exit Function

WorkFailed:
    Dim lngErrNum As Long
    Dim strErrSrc As String
    Dim strErrDesc As String
    lngErrNum = Err.Number
    strErrSrc = Err.Source
    strErrDesc = Err.Description
    m_Conn.RollbackTrans
    Err.Raise lngErrNum, strErrSrc, strErrDesc
End Function

Public Sub RollbackWork()
    Set m_Operations = New Collection
End Sub

