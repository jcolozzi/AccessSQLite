Attribute VB_Name = "modAxLite"
Option Compare Database
Option Explicit

' =============================================================================
' modAxLite
' Public entry points, window manager, Ribbon callbacks, and USysRibbons
' installer for the Monaco-powered AxLite SQLite front-end client.
' Ribbon callbacks use `As Object` so no Office type library reference is needed.
' =============================================================================

Public g_OpeningAxLiteViaFactory As Boolean
Private m_colAxLiteWindows As Collection
Private m_lngNextWindowSeq As Long
Private m_objRibbonUi As Object
Private m_objHeadlessAxLite As clsAxLiteHost

Public Function OpenAxLite(Optional ByVal strSqlitePath As String = vbNullString) As String
    Dim frmEditor As Form_frmAxLite
    Dim strKey As String

    On Error GoTo OpenErr
    If m_colAxLiteWindows Is Nothing Then Set m_colAxLiteWindows = New Collection
    m_lngNextWindowSeq = m_lngNextWindowSeq + 1
    strKey = "axlite_" & m_lngNextWindowSeq

    g_OpeningAxLiteViaFactory = True
    Set frmEditor = New Form_frmAxLite
    g_OpeningAxLiteViaFactory = False

    m_colAxLiteWindows.Add frmEditor, strKey
    frmEditor.StartAxLiteWindow strKey, strSqlitePath
    frmEditor.Visible = True
    On Error Resume Next
    frmEditor.SetFocus
    On Error GoTo 0
    OpenAxLite = strKey
    Exit Function

OpenErr:
    g_OpeningAxLiteViaFactory = False
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Prompts for an existing .db / .sqlite file via native Access FileDialog and opens it in AxLite.
Public Function OpenAxLitePicker() As String
    Const MSO_FILE_DIALOG_OPEN As Long = 3
    Dim objDlg As Object
    Dim strPicked As String

    On Error GoTo PickerErr
    Set objDlg = Application.FileDialog(MSO_FILE_DIALOG_OPEN)
    objDlg.Title = "Open SQLite Database in AxLite"
    objDlg.AllowMultiSelect = False
    objDlg.InitialFileName = CurrentProject.Path & "\"
    On Error Resume Next
    objDlg.Filters.Clear
    objDlg.Filters.Add "SQLite Databases (*.db; *.sqlite; *.sqlite3)", "*.db;*.sqlite;*.sqlite3"
    objDlg.Filters.Add "All Files (*.*)", "*.*"
    On Error GoTo PickerErr

    If Not objDlg.Show Then Exit Function
    strPicked = CStr(objDlg.SelectedItems(1))
    OpenAxLitePicker = OpenAxLite(strPicked)
    Exit Function

PickerErr:
    MsgBox "Could not open file picker: " & Err.Description, vbExclamation, "AxLite"
End Function

' Opens an isolated in-memory (:memory:) SQLite scratchpad in AxLite.
Public Function OpenAxLiteMemory() As String
    OpenAxLiteMemory = OpenAxLite(":memory:")
End Function

Public Sub CloseAllAxLiteWindows()
    Set m_colAxLiteWindows = Nothing
End Sub

Public Function OpenAxLiteCount() As Long
    If Not m_colAxLiteWindows Is Nothing Then
        OpenAxLiteCount = m_colAxLiteWindows.Count
    End If
End Function

Public Sub ReleaseAxLiteWindow(ByVal strHostKey As String)
    If m_colAxLiteWindows Is Nothing Then Exit Sub
    On Error Resume Next
    m_colAxLiteWindows.Remove strHostKey
    On Error GoTo 0
End Sub

' Headless JSON dispatch helper for automated AxLite bridge verification.
Public Function AxLiteDispatchHeadless( _
    ByVal strDbPath As String, _
    ByVal strEnvelopeJson As String) As String

    If m_objHeadlessAxLite Is Nothing Then
        Set m_objHeadlessAxLite = New clsAxLiteHost
        m_objHeadlessAxLite.StartHeadless strDbPath
    ElseIf Len(Trim$(strDbPath)) > 0 And StrComp(m_objHeadlessAxLite.CurrentDatabasePath, Trim$(strDbPath), vbTextCompare) <> 0 Then
        m_objHeadlessAxLite.StartHeadless strDbPath
    End If
    AxLiteDispatchHeadless = m_objHeadlessAxLite.DispatchJson(strEnvelopeJson)
End Function

Public Sub ResetHeadlessAxLite()
    If Not m_objHeadlessAxLite Is Nothing Then
        m_objHeadlessAxLite.ShutdownHost
        Set m_objHeadlessAxLite = Nothing
    End If
End Sub

' =============================================================================
' AccessSQLite Library Ribbon Callbacks (USysRibbons: "AxLite")
' =============================================================================

Public Sub AxLiteRibbonLoad(objRibbon As Object)
    Set m_objRibbonUi = objRibbon
End Sub

Public Sub AxLiteRibbonAction(objControl As Object)
    On Error GoTo RibbonErr
    Select Case CStr(objControl.Id)
        Case "btnOpenAxLite"
            OpenAxLite vbNullString
        Case "btnOpenDbPicker"
            OpenAxLitePicker
        Case "btnOpenMemoryAxLite"
            OpenAxLiteMemory
        Case "btnRunSelfTests"
            MsgBox RunAllSQLiteTests(), vbInformation, "AccessSQLite Automated Test Suite"
        Case "btnCloseAxLite"
            CloseAllAxLiteWindows
        Case "btnAxLiteAbout"
            MsgBox "AccessSQLite Class Library, ORM & AxLite SQL Editor" & vbCrLf & vbCrLf & _
                   "Active Default Driver: " & GetSQLiteEngineVersion(sqDriverAuto) & vbCrLf & _
                   "Bundled Fallback Driver: " & GetSQLiteEngineVersion(sqDriverBundled) & vbCrLf & _
                   "Library Folder: " & GetLibraryDirectory(), _
                   vbInformation, "About AxLite & AccessSQLite"
    End Select
    Exit Sub

RibbonErr:
    MsgBox Err.Description, vbExclamation, "AxLite"
End Sub

' Installs/updates a CustomUI XML definition into CurrentDb's USysRibbons table and sets CustomRibbonID.
Public Function InstallCustomRibbonXml(ByVal strRibbonName As String, ByVal strRibbonXml As String) As String
    Dim dbCur As DAO.Database
    Dim tdfItem As DAO.TableDef
    Dim blnHasTable As Boolean
    Dim rstRib As DAO.Recordset
    Dim prpNew As DAO.Property

    Set dbCur = CurrentDb
    For Each tdfItem In dbCur.TableDefs
        If StrComp(tdfItem.Name, "USysRibbons", vbTextCompare) = 0 Then
            blnHasTable = True
            Exit For
        End If
    Next tdfItem

    If Not blnHasTable Then
        dbCur.Execute _
            "CREATE TABLE USysRibbons (" & _
            "  ID AUTOINCREMENT CONSTRAINT pkUSysRibbons PRIMARY KEY," & _
            "  RibbonName TEXT(255) NOT NULL," & _
            "  RibbonXML MEMO" & _
            ");", dbFailOnError
        dbCur.TableDefs.Refresh
    End If

    dbCur.Execute "DELETE FROM USysRibbons WHERE RibbonName = '" & Replace(strRibbonName, "'", "''") & "';", dbFailOnError
    Set rstRib = dbCur.OpenRecordset("USysRibbons", dbOpenDynaset)
    rstRib.AddNew
    rstRib!RibbonName = strRibbonName
    rstRib!RibbonXML = strRibbonXml
    rstRib.Update
    rstRib.Close

    On Error Resume Next
    Application.LoadCustomUI strRibbonName, strRibbonXml
    Err.Clear
    dbCur.Properties("CustomRibbonID").Value = strRibbonName
    If Err.Number <> 0 Then
        Err.Clear
        Set prpNew = dbCur.CreateProperty("CustomRibbonID", dbText, strRibbonName)
        dbCur.Properties.Append prpNew
    End If
    On Error GoTo 0

    InstallCustomRibbonXml = "Installed ribbon '" & strRibbonName & "'"
End Function

Public Function InstallAxLiteRibbon() As String
    Dim strXml As String
    strXml = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
    strXml = strXml & "<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""AxLiteRibbonLoad"">" & vbCrLf
    strXml = strXml & "  <ribbon startFromScratch=""false"">" & vbCrLf
    strXml = strXml & "    <tabs>" & vbCrLf
    strXml = strXml & "      <tab id=""tabAxLite"" label=""AxLite"" insertBeforeMso=""TabHomeAccess"">" & vbCrLf
    strXml = strXml & "        <group id=""grpAxLiteWorkbench"" label=""AxLite SQL Editor (Monaco)"">" & vbCrLf
    strXml = strXml & "          <button id=""btnOpenAxLite"" label=""AxLite Editor"" size=""large"" imageMso=""SqlView"" onAction=""AxLiteRibbonAction"" screentip=""Open AxLite SQL Editor"" supertip=""Opens the Monaco-powered AxLite SQLite front-end client to browse tables/columns, manage stored queries, and run any SQL operation."" />" & vbCrLf
    strXml = strXml & "          <button id=""btnOpenDbPicker"" label=""Open .db File..."" size=""large"" imageMso=""FileOpen"" onAction=""AxLiteRibbonAction"" screentip=""Open SQLite File"" supertip=""Select any .db, .sqlite, or .sqlite3 file on disk and open it in AxLite."" />" & vbCrLf
    strXml = strXml & "          <button id=""btnOpenMemoryAxLite"" label=""In-Memory Scratchpad"" size=""large"" imageMso=""CreateQueryInDesignView"" onAction=""AxLiteRibbonAction"" screentip=""Open :memory: in AxLite"" supertip=""Opens an isolated in-memory SQLite database in the AxLite SQL workbench."" />" & vbCrLf
    strXml = strXml & "        </group>" & vbCrLf
    strXml = strXml & "        <group id=""grpAxLiteTools"" label=""Library &amp; Diagnostics"">" & vbCrLf
    strXml = strXml & "          <button id=""btnRunSelfTests"" label=""Run Self-Tests"" size=""large"" imageMso=""MacroPlay"" onAction=""AxLiteRibbonAction"" screentip=""Run AccessSQLite Tests"" supertip=""Executes the automated VBA unit &amp; integration test suite."" />" & vbCrLf
    strXml = strXml & "          <button id=""btnCloseAxLite"" label=""Close AxLite Windows"" imageMso=""FileClose"" onAction=""AxLiteRibbonAction"" />" & vbCrLf
    strXml = strXml & "          <button id=""btnAxLiteAbout"" label=""About AxLite"" imageMso=""Info"" onAction=""AxLiteRibbonAction"" />" & vbCrLf
    strXml = strXml & "        </group>" & vbCrLf
    strXml = strXml & "      </tab>" & vbCrLf
    strXml = strXml & "    </tabs>" & vbCrLf
    strXml = strXml & "  </ribbon>" & vbCrLf
    strXml = strXml & "</customUI>"

    InstallAxLiteRibbon = InstallCustomRibbonXml("AxLite", strXml)
End Function

' Evaluates JavaScript inside an open AxLite window (used by automated UI verification).
Public Function AxLiteEvalJs(ByVal strHostKey As String, ByVal strScript As String) As String
    Dim objEdge As Object
    If Len(strHostKey) > 0 Then
        Set objEdge = m_colAxLiteWindows(strHostKey).ctlWebHost
    Else
        Set objEdge = Forms("frmAxLite").ctlWebHost
    End If
    objEdge.ExecuteJavascript "try { window.__axEval = String((0, eval)(" & LibJSON.Serialize(strScript) & ")) } " & _
        "catch (e) { window.__axEval = 'Error: ' + e.message }"
    AxLiteEvalJs = Nz(objEdge.RetrieveJavascriptValue("window.__axEval"), vbNullString)
End Function
