Attribute VB_Name = "clsSQLiteWebBridge"
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
' clsSQLiteWebBridge
' Access Edge Browser Control <-> JavaScript JSON bridge for AccessSQLite Studio.
' Adapted from Access Designer's clsAasWebBridge:
'   - Polls window.SQLiteStudio.bridge.peek() via RetrieveJavascriptValue (no '?' in script!)
'   - Acknowledges processed sequence numbers via ExecuteJavascript
'   - Splits replies > 32,000 chars into receivePart(...) chunks so oversized payloads never drop
' =============================================================================

Private Const HOST_PREFIX As String = "https://msaccess/"
Private Const PEEK_SCRIPT As String = "(window.SQLiteStudio && window.SQLiteStudio.bridge && window.SQLiteStudio.bridge.peek()) || ''"
Private Const TIMEOUT_TEXT As String = "RetrieveJavascriptValue timed out*"
Private Const CHUNK_CHARS As Long = 32000

Private m_objEdge As Object
Private m_blnReady As Boolean
Private m_blnNavigating As Boolean
Private m_dblLastSeq As Double
Private m_lngChunkKey As Long

Public Sub InitBridge(ByVal objEdge As Object)
    Set m_objEdge = objEdge
End Sub

Public Property Get IsReady() As Boolean
    IsReady = m_blnReady
End Property

Public Sub MarkReady()
    m_blnReady = True
End Sub

Public Sub MarkNavigating()
    m_blnNavigating = True
    m_dblLastSeq = 0
End Sub

Public Function IsAppUrl(ByVal strUrl As String) As Boolean
    IsAppUrl = (StrComp(Left$(strUrl, Len(HOST_PREFIX)), HOST_PREFIX, vbTextCompare) = 0) Or (strUrl = "about:blank")
End Function

Public Sub LoadPage(ByVal strLocalHtmlPath As String)
    Dim strMapped As String
    m_blnReady = False
    m_blnNavigating = False
    m_dblLastSeq = 0
    strMapped = ToMappedDrivePath(strLocalHtmlPath)
    m_objEdge.Navigate HOST_PREFIX & strMapped
End Sub

Public Function PollOutbox() As Collection
    Dim strRaw As String
    Dim udtParsed As ParseResult
    Dim varItem As Variant
    Dim dblSeq As Double
    Dim colFresh As New Collection

    Set PollOutbox = colFresh
    If Not (m_blnReady Or m_blnNavigating) Then Exit Function

    On Error Resume Next
    strRaw = Nz(m_objEdge.RetrieveJavascriptValue(PEEK_SCRIPT), vbNullString)
    If Err.Number <> 0 Then strRaw = vbNullString
    On Error GoTo 0

    If strRaw Like TIMEOUT_TEXT And Not m_blnReady Then m_blnNavigating = False
    If Len(strRaw) = 0 Or strRaw Like TIMEOUT_TEXT Then Exit Function
    If Left$(strRaw, 3) <> "%5B" Then Exit Function

    udtParsed = LibJSON.Parse(PercentDecodeToUtf8(strRaw), jpCodeUTF8)
    If Not udtParsed.IsValid Then
        ExecuteJs "SQLiteStudio.bridge.ack(Number.MAX_SAFE_INTEGER);"
        Exit Function
    End If
    If Not IsObject(udtParsed.Value) Then Exit Function
    If Not TypeOf udtParsed.Value Is Collection Then Exit Function

    For Each varItem In udtParsed.Value
        If IsObject(varItem) Then
            dblSeq = Val(varItem("seq"))
            If dblSeq > m_dblLastSeq Then
                If IsObject(varItem("env")) Then colFresh.Add varItem("env")
                m_dblLastSeq = dblSeq
            End If
        End If
    Next varItem
End Function

Public Sub AckOutbox()
    ExecuteJs "SQLiteStudio.bridge.ack(" & m_dblLastSeq & ");"
End Sub

Public Sub SendReply(ByVal dicReply As Scripting.Dictionary)
    SendJsonString LibJSON.Serialize(dicReply)
End Sub

Public Sub EmitEvent(ByVal strEventName As String, ByVal varData As Variant)
    Dim dicMsg As New Scripting.Dictionary
    dicMsg("event") = strEventName
    If IsObject(varData) Then Set dicMsg("data") = varData Else dicMsg("data") = varData
    SendReply dicMsg
End Sub

Private Sub SendJsonString(ByVal strJson As String)
    Dim lngParts As Long
    Dim lngPart As Long
    Dim strKey As String

    If Len(strJson) <= CHUNK_CHARS Then
        ExecuteJs "SQLiteStudio.bridge.receive(" & strJson & ");"
        Exit Sub
    End If

    m_lngChunkKey = m_lngChunkKey + 1
    strKey = CStr(m_lngChunkKey)
    lngParts = (Len(strJson) + CHUNK_CHARS - 1) \ CHUNK_CHARS
    For lngPart = 0 To lngParts - 1
        ExecuteJs "SQLiteStudio.bridge.receivePart(" & strKey & "," & lngPart & "," & lngParts & "," & _
            LibJSON.Serialize(Mid$(strJson, lngPart * CHUNK_CHARS + 1, CHUNK_CHARS)) & ");"
    Next lngPart
End Sub

Private Sub ExecuteJs(ByVal strScript As String)
    On Error Resume Next
    m_objEdge.ExecuteJavascript strScript
    On Error GoTo 0
End Sub

Private Function PercentDecodeToUtf8(ByVal strText As String) As Byte()
    Dim arrOut() As Byte
    Dim lngCount As Long
    Dim lngPos As Long
    Dim lngCode As Long
    Dim strHexPair As String

    ReDim arrOut(0 To Len(strText) * 4)
    lngPos = 1
    Do While lngPos <= Len(strText)
        lngCode = AscW(Mid$(strText, lngPos, 1)) And &HFFFF&
        If lngCode = 37 And lngPos + 2 <= Len(strText) Then
            strHexPair = Mid$(strText, lngPos + 1, 2)
            If Not strHexPair Like "[0-9A-Fa-f][0-9A-Fa-f]" Then
                RaiseSQLiteError sqErrInvalidArgument, "clsSQLiteWebBridge", "Malformed percent-escape in bridge outbox."
            End If
            arrOut(lngCount) = CByte("&H" & strHexPair)
            lngCount = lngCount + 1
            lngPos = lngPos + 3
        Else
            If lngCode >= &HD800& And lngCode <= &HDBFF& And lngPos < Len(strText) Then
                lngCode = &H10000 + (lngCode - &HD800&) * &H400& + ((AscW(Mid$(strText, lngPos + 1, 1)) And &HFFFF&) - &HDC00&)
                lngPos = lngPos + 1
            End If
            AppendUtf8ByteSeq arrOut, lngCount, lngCode
            lngPos = lngPos + 1
        End If
    Loop
    If lngCount = 0 Then
        ReDim arrOut(0 To 0)
        PercentDecodeToUtf8 = arrOut
        Exit Function
    End If
    ReDim Preserve arrOut(0 To lngCount - 1)
    PercentDecodeToUtf8 = arrOut
End Function

Private Sub AppendUtf8ByteSeq(ByRef arrOut() As Byte, ByRef lngCount As Long, ByVal lngCode As Long)
    Select Case lngCode
        Case Is < &H80&
            arrOut(lngCount) = lngCode
            lngCount = lngCount + 1
        Case Is < &H800&
            arrOut(lngCount) = &HC0& Or (lngCode \ &H40&)
            arrOut(lngCount + 1) = &H80& Or (lngCode And &H3F&)
            lngCount = lngCount + 2
        Case Is < &H10000
            arrOut(lngCount) = &HE0& Or (lngCode \ &H1000&)
            arrOut(lngCount + 1) = &H80& Or ((lngCode \ &H40&) And &H3F&)
            arrOut(lngCount + 2) = &H80& Or (lngCode And &H3F&)
            lngCount = lngCount + 3
        Case Else
            arrOut(lngCount) = &HF0& Or (lngCode \ &H40000)
            arrOut(lngCount + 1) = &H80& Or ((lngCode \ &H1000&) And &H3F&)
            arrOut(lngCount + 2) = &H80& Or ((lngCode \ &H40&) And &H3F&)
            arrOut(lngCount + 3) = &H80& Or (lngCode And &H3F&)
            lngCount = lngCount + 4
    End Select
End Sub

