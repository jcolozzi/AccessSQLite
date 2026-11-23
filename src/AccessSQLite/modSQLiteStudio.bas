Attribute VB_Name = "modSQLiteStudio"
Option Compare Database
Option Explicit

' =============================================================================
' modSQLiteStudio
' Public entry points, window manager, Ribbon callbacks, and USysRibbons
' installer for the Monaco-powered AccessSQLite Studio front-end client.
' Ribbon callbacks use `As Object` so no Office type library reference is needed.
' =============================================================================

Public g_OpeningStudioViaFactory As Boolean
Private m_colStudioWindows As Collection
Private m_lngNextWindowSeq As Long
Private m_objRibbonUi As Object
Private m_objHeadlessStudio As clsSQLiteStudioHost

Public Function OpenSQLiteStudio(Optional ByVal strSqlitePath As String = vbNullString) As String
    Dim frmStudio As Form_frmSQLiteStudio
    Dim strKey As String

    On Error GoTo OpenErr
    If m_colStudioWindows Is Nothing Then Set m_colStudioWindows = New Collection
    m_lngNextWindowSeq = m_lngNextWindowSeq + 1
    strKey = "studio_" & m_lngNextWindowSeq

    g_OpeningStudioViaFactory = True
    Set frmStudio = New Form_frmSQLiteStudio
    g_OpeningStudioViaFactory = False

    m_colStudioWindows.Add frmStudio, strKey
    frmStudio.StartStudioWindow strKey, strSqlitePath
    frmStudio.Visible = True
    On Error Resume Next
    frmStudio.SetFocus
    On Error GoTo 0
    OpenSQLiteStudio = strKey
    Exit Function

OpenErr:
    g_OpeningStudioViaFactory = False
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Prompts for an existing .db / .sqlite file via native Access FileDialog and opens it in SQLite Studio.
Public Function OpenSQLiteStudioPicker() As String
    Const MSO_FILE_DIALOG_OPEN As Long = 3
    Dim objDlg As Object
    Dim strPicked As String

    On Error GoTo PickerErr
    Set objDlg = Application.FileDialog(MSO_FILE_DIALOG_OPEN)
    objDlg.Title = "Open SQLite Database in AccessSQLite Studio"
    objDlg.AllowMultiSelect = False
    objDlg.InitialFileName = CurrentProject.Path & "\"
    On Error Resume Next
    objDlg.Filters.Clear
    objDlg.Filters.Add "SQLite Databases (*.db; *.sqlite; *.sqlite3)", "*.db;*.sqlite;*.sqlite3"
    objDlg.Filters.Add "All Files (*.*)", "*.*"
    On Error GoTo PickerErr

    If Not objDlg.Show Then Exit Function
    strPicked = CStr(objDlg.SelectedItems(1))
    OpenSQLiteStudioPicker = OpenSQLiteStudio(strPicked)
    Exit Function

PickerErr:
    MsgBox "Could not open file picker: " & Err.Description, vbExclamation, "AccessSQLite Studio"
End Function

' Opens an isolated in-memory (:memory:) SQLite scratchpad in AccessSQLite Studio.
Public Function OpenSQLiteStudioMemory() As String
    OpenSQLiteStudioMemory = OpenSQLiteStudio(":memory:")
End Function

Public Sub CloseAllSQLiteStudios()
    Set m_colStudioWindows = Nothing
End Sub

Public Function OpenSQLiteStudioCount() As Long
    If Not m_colStudioWindows Is Nothing Then
        OpenSQLiteStudioCount = m_colStudioWindows.Count
    End If
End Function

Public Sub ReleaseSQLiteStudioWindow(ByVal strHostKey As String)
    If m_colStudioWindows Is Nothing Then Exit Sub
    On Error Resume Next
    m_colStudioWindows.Remove strHostKey
    On Error GoTo 0
End Sub

' Headless JSON dispatch helper for automated Studio bridge verification.
Public Function SQLiteStudioDispatchHeadless( _
    ByVal strDbPath As String, _
    ByVal strEnvelopeJson As String) As String

    If m_objHeadlessStudio Is Nothing Then
        Set m_objHeadlessStudio = New clsSQLiteStudioHost
        m_objHeadlessStudio.StartHeadless strDbPath
    ElseIf Len(Trim$(strDbPath)) > 0 And StrComp(m_objHeadlessStudio.CurrentDatabasePath, Trim$(strDbPath), vbTextCompare) <> 0 Then
        m_objHeadlessStudio.StartHeadless strDbPath
    End If
    SQLiteStudioDispatchHeadless = m_objHeadlessStudio.DispatchJson(strEnvelopeJson)
End Function

Public Sub ResetHeadlessSQLiteStudio()
    If Not m_objHeadlessStudio Is Nothing Then
        m_objHeadlessStudio.ShutdownHost
        Set m_objHeadlessStudio = Nothing
    End If
End Sub

' =============================================================================
' AccessSQLite Library Ribbon Callbacks (USysRibbons: "SQLiteStudio")
' =============================================================================

Public Sub SQLiteRibbonLoad(objRibbon As Object)
    Set m_objRibbonUi = objRibbon
End Sub

Public Sub SQLiteRibbonAction(objControl As Object)
    On Error GoTo RibbonErr
    Select Case CStr(objControl.Id)
        Case "btnOpenStudio"
            OpenSQLiteStudio vbNullString
        Case "btnOpenDbPicker"
            OpenSQLiteStudioPicker
        Case "btnOpenMemoryStudio"
            OpenSQLiteStudioMemory
        Case "btnRunSelfTests"
            MsgBox RunAllSQLiteTests(), vbInformation, "AccessSQLite Automated Test Suite"
        Case "btnCloseStudios"
            CloseAllSQLiteStudios
        Case "btnStudioAbout"
            MsgBox "AccessSQLite Class Library, ORM & Monaco SQLite Studio" & vbCrLf & vbCrLf & _
                   "Active Default Driver: " & GetSQLiteEngineVersion(sqDriverAuto) & vbCrLf & _
                   "Bundled Fallback Driver: " & GetSQLiteEngineVersion(sqDriverBundled) & vbCrLf & _
                   "Library Folder: " & GetLibraryDirectory(), _
                   vbInformation, "About AccessSQLite"
    End Select
    Exit Sub

RibbonErr:
    MsgBox Err.Description, vbExclamation, "AccessSQLite Studio"
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

Public Function InstallLibraryStudioRibbon() As String
    Dim strXml As String
    strXml = _
        "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf & _
        "<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""SQLiteRibbonLoad"">" & vbCrLf & _
        "  <ribbon startFromScratch=""false"">" & vbCrLf & _
        "    <tabs>" & vbCrLf & _
        "      <tab id=""tabSQLiteStudio"" label=""SQLite Studio"" insertBeforeMso=""TabHomeAccess"">" & vbCrLf & _
        "        <group id=""grpSQLiteWorkbench"" label=""SQLite Workbench (Monaco)"">" & vbCrLf & _
        "          <button id=""btnOpenStudio"" label=""SQLite Studio"" size=""large"" imageMso=""SqlView"" onAction=""SQLiteRibbonAction""" & _
        "                  screentip=""Open AccessSQLite Studio"" supertip=""Opens the Monaco SQLite front-end client to browse tables/columns, manage stored queries, and run any SQL operation."" />" & vbCrLf & _
        "          <button id=""btnOpenDbPicker"" label=""Open .db File..."" size=""large"" imageMso=""FileOpen"" onAction=""SQLiteRibbonAction""" & _
        "                  screentip=""Open SQLite File"" supertip=""Select any .db, .sqlite, or .sqlite3 file on disk and open it in SQLite Studio."" />" & vbCrLf & _
        "          <button id=""btnOpenMemoryStudio"" label=""In-Memory Scratchpad"" size=""large"" imageMso=""CreateQueryInDesignView"" onAction=""SQLiteRibbonAction""" & _
        "                  screentip=""Open :memory: SQLite Studio"" supertip=""Opens an isolated in-memory SQLite database in the Monaco SQL workbench."" />" & vbCrLf & _
        "        </group>" & vbCrLf & _
        "        <group id=""grpSQLiteTools"" label=""Library &amp; Diagnostics"">" & vbCrLf & _
        "          <button id=""btnRunSelfTests"" label=""Run Self-Tests"" size=""large"" imageMso=""MacroPlay"" onAction=""SQLiteRibbonAction""" & _
        "                  screentip=""Run AccessSQLite Tests"" supertip=""Executes the automated VBA unit &amp; integration test suite."" />" & vbCrLf & _
        "          <button id=""btnCloseStudios"" label=""Close Studio Windows"" imageMso=""FileClose"" onAction=""SQLiteRibbonAction"" />" & vbCrLf & _
        "          <button id=""btnStudioAbout"" label=""About AccessSQLite"" imageMso=""Info"" onAction=""SQLiteRibbonAction"" />" & vbCrLf & _
        "        </group>" & vbCrLf & _
        "      </tab>" & vbCrLf & _
        "    </tabs>" & vbCrLf & _
        "  </ribbon>" & vbCrLf & _
        "</customUI>"

    InstallLibraryStudioRibbon = InstallCustomRibbonXml("SQLiteStudio", strXml)
End Function


' Evaluates JavaScript inside an open Studio window (used by automated UI verification).
Public Function SQLiteStudioEvalJs(ByVal strHostKey As String, ByVal strScript As String) As String
    Dim objEdge As Object
    If Len(strHostKey) > 0 Then
        Set objEdge = m_colStudioWindows(strHostKey).ctlWebHost
    Else
        Set objEdge = Forms("frmSQLiteStudio").ctlWebHost
    End If
    objEdge.ExecuteJavascript "try { window.__ssEval = String((0, eval)(" & LibJSON.Serialize(strScript) & ")) } " & _
        "catch (e) { window.__ssEval = 'Error: ' + e.message }"
    SQLiteStudioEvalJs = Nz(objEdge.RetrieveJavascriptValue("window.__ssEval"), vbNullString)
End Function


