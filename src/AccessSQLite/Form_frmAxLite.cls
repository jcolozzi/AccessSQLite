Attribute VB_Name = "Form_frmAxLite"
Attribute VB_Base = "0{D3141D46-F0EC-4EF0-AE72-E5CB6355FB12}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

' Thin Edge Browser Control host form forwarding events to clsAxLiteHost.

Private m_objHost As clsAxLiteHost
Private m_strHostKey As String

Public Sub StartAxLiteWindow(ByVal strHostKey As String, Optional ByVal strSqlitePath As String = vbNullString)
    If Not m_objHost Is Nothing Then Exit Sub
    m_strHostKey = strHostKey
    Set m_objHost = New clsAxLiteHost
    m_objHost.StartHost Me, Me.ctlWebHost, strSqlitePath
End Sub

Private Sub Form_Load()
    If g_OpeningAxLiteViaFactory Then Exit Sub
    If Not m_objHost Is Nothing Then Exit Sub
    If Not IsNull(Me.OpenArgs) Then
        StartAxLiteWindow vbNullString, CStr(Me.OpenArgs)
    Else
        StartAxLiteWindow vbNullString, vbNullString
    End If
End Sub

Private Sub Form_Resize()
    Dim lngWidth As Long
    Dim lngHeight As Long
    On Error Resume Next
    lngWidth = Me.InsideWidth
    lngHeight = Me.InsideHeight
    If lngWidth <= 0 Or lngHeight <= 0 Then Exit Sub
    If lngHeight > Me.Section(acDetail).Height Then Me.Section(acDetail).Height = lngHeight
    Me.ctlWebHost.Top = 0
    Me.ctlWebHost.Left = 0
    Me.ctlWebHost.Width = lngWidth
    Me.ctlWebHost.Height = lngHeight
    Me.Width = lngWidth
    Me.Section(acDetail).Height = lngHeight
    On Error GoTo 0
End Sub

Private Sub Form_Timer()
    If Not m_objHost Is Nothing Then m_objHost.OnTimer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not m_objHost Is Nothing Then
        m_objHost.ShutdownHost
        Set m_objHost = Nothing
    End If
End Sub

Private Sub Form_Close()
    If Len(m_strHostKey) > 0 Then ReleaseAxLiteWindow m_strHostKey
End Sub

Private Sub ctlWebHost_BeforeNavigate(Cancel As Integer, URL As String)
    If Not m_objHost Is Nothing Then m_objHost.OnBeforeNavigate Cancel, URL
End Sub

Private Sub ctlWebHost_DocumentComplete(URL As Variant)
    If Not m_objHost Is Nothing Then m_objHost.OnDocumentComplete
End Sub
