Attribute VB_Name = "modDemoRibbon"
Option Compare Database
Option Explicit

' =============================================================================
' modDemoRibbon
' CustomUI Ribbon callbacks & installer for AccessSQLiteDemo.accdb.
' =============================================================================

Private m_DemoRibbon As Object ' IRibbonUI

Public Sub DemoRibbonLoad(ByVal objRibbon As Object)
    Set m_DemoRibbon = objRibbon
End Sub

Public Sub DemoRibbonAction(ByVal objControl As Object)
    On Error GoTo ActionErr
    Select Case CStr(objControl.ID)
        Case "btnDemoStudio"
            GetDemoConnection
            AccessSQLite.OpenSQLiteStudio DemoSqliteFilePath()
        Case "btnDemoOpenDbPicker"
            AccessSQLite.OpenSQLiteStudioPicker
        Case "btnDemoMemoryStudio"
            AccessSQLite.OpenSQLiteStudioMemory
        Case "btnOpenAdoContinuous"
            DoCmd.OpenForm "frmTaskListAdo", acNormal
        Case "btnOpenDaoContinuous"
            DoCmd.OpenForm "frmTaskList", acNormal
        Case "btnOpenSingleDetail"
            DoCmd.OpenForm "frmTaskDetail", acNormal
        Case "btnRunDemoVerify"
            MsgBox RunDemoFormsVerification(), vbInformation, "AccessSQLite Demo Verification"
        Case "btnDemoAbout"
            MsgBox "AccessSQLite Reference Front-End" & vbCrLf & vbCrLf & _
                   "Demonstrates 3 form patterns + Monaco SQLite Studio:" & vbCrLf & _
                   "  1. frmTaskListAdo - In-Memory ADODB.Recordset (zero temp table + virtual checkboxes)" & vbCrLf & _
                   "  2. frmTaskList    - DAO Staging Table (tblTaskBuffer)" & vbCrLf & _
                   "  3. frmTaskDetail  - Unbound ISQLiteEntity Single-Page Form" & vbCrLf & _
                   "  4. SQLite Studio  - Monaco Editor SQL Workbench", _
                   vbInformation, "About AccessSQLite Demo"
    End Select
    Exit Sub

ActionErr:
    MsgBox Err.Description, vbExclamation, "AccessSQLite Demo"
End Sub

Public Function InstallDemoRibbon() As String
    Dim strXml As String
    strXml = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
    strXml = strXml & "<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""DemoRibbonLoad"">" & vbCrLf
    strXml = strXml & "  <ribbon startFromScratch=""false"">" & vbCrLf
    strXml = strXml & "    <tabs>" & vbCrLf
    strXml = strXml & "      <tab id=""tabAccessSQLiteDemo"" label=""AccessSQLite"" insertBeforeMso=""TabHomeAccess"">" & vbCrLf
    strXml = strXml & "        <group id=""grpMonacoStudio"" label=""SQLite Studio (Monaco)"">" & vbCrLf
    strXml = strXml & "          <button id=""btnDemoStudio"" label=""SQLite Studio"" size=""large"" imageMso=""SqlView"" onAction=""DemoRibbonAction"" screentip=""Open AccessSQLite Studio"" supertip=""Opens the Monaco SQLite front-end client connected to AccessSQLiteDemo.db."" />" & vbCrLf
    strXml = strXml & "          <button id=""btnDemoOpenDbPicker"" label=""Open .db File..."" size=""large"" imageMso=""FileOpen"" onAction=""DemoRibbonAction"" screentip=""Open Any SQLite File"" supertip=""Browse for any .db, .sqlite, or .sqlite3 file and open it in SQLite Studio."" />" & vbCrLf
    strXml = strXml & "          <button id=""btnDemoMemoryStudio"" label=""In-Memory DB"" size=""large"" imageMso=""CreateQueryInDesignView"" onAction=""DemoRibbonAction"" screentip=""Open :memory: Scratchpad"" supertip=""Opens an isolated in-memory SQLite database in Monaco Studio."" />" & vbCrLf
    strXml = strXml & "        </group>" & vbCrLf
    strXml = strXml & "        <group id=""grpDemoForms"" label=""Access Form Approaches (ADO vs DAO vs ORM)"">" & vbCrLf
    strXml = strXml & "          <button id=""btnOpenAdoContinuous"" label=""Tasks (ADO In-Memory)"" size=""large"" imageMso=""TableInsert"" onAction=""DemoRibbonAction"" screentip=""In-Memory ADODB.Recordset Continuous Form"" supertip=""Continuous Form bound to a disconnected client-side ADODB.Recordset with virtual per-row selection checkboxes and zero temp tables."" />" & vbCrLf
    strXml = strXml & "          <button id=""btnOpenDaoContinuous"" label=""Tasks (DAO Buffer)"" size=""large"" imageMso=""ViewsDatasheetView"" onAction=""DemoRibbonAction"" screentip=""DAO Staging Table Continuous Form"" supertip=""Continuous Form backed by local tblTaskBuffer with inline ORM writeback."" />" & vbCrLf
    strXml = strXml & "          <button id=""btnOpenSingleDetail"" label=""Task Detail (ORM)"" size=""large"" imageMso=""ViewsFormView"" onAction=""DemoRibbonAction"" screentip=""Single-Page ORM Form"" supertip=""Unbound Single-Page Form backed directly by clsTaskItem (ISQLiteEntity)."" />" & vbCrLf
    strXml = strXml & "        </group>" & vbCrLf
    strXml = strXml & "        <group id=""grpDemoDiag"" label=""Verification"">" & vbCrLf
    strXml = strXml & "          <button id=""btnRunDemoVerify"" label=""Verify All Forms"" size=""large"" imageMso=""MacroPlay"" onAction=""DemoRibbonAction"" screentip=""Run End-to-End Tests"" supertip=""Verifies DAO Continuous Form, In-Memory ADO Continuous Form, and Single-Page ORM Form round-trips."" />" & vbCrLf
    strXml = strXml & "          <button id=""btnDemoAbout"" label=""About"" imageMso=""Info"" onAction=""DemoRibbonAction"" />" & vbCrLf
    strXml = strXml & "        </group>" & vbCrLf
    strXml = strXml & "      </tab>" & vbCrLf
    strXml = strXml & "    </tabs>" & vbCrLf
    strXml = strXml & "  </ribbon>" & vbCrLf
    strXml = strXml & "</customUI>"

    InstallDemoRibbon = AccessSQLite.InstallCustomRibbonXml("AccessSQLiteDemo", strXml)
End Function

