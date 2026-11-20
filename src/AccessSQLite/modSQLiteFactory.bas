Attribute VB_Name = "modSQLiteFactory"
Option Compare Database
Option Explicit

' =============================================================================
' modSQLiteFactory
' Public constructor & helper entry points for AccessSQLite.accdb.
' Enables external .accdb front-ends referencing AccessSQLite (Tools -> References)
' to instantiate connections, table maps, migrators, bridges, and units of work
' with full early-bound IntelliSense.
' =============================================================================

Public Function CreateSQLiteConnection() As clsSQLiteConnection
    Set CreateSQLiteConnection = New clsSQLiteConnection
End Function

Public Function OpenSQLite( _
    ByVal strFilePath As String, _
    Optional ByVal enmProfile As SQLitePragmaProfile = sqProfileAuto, _
    Optional ByVal enmDriver As SQLiteDriverKind = sqDriverAuto, _
    Optional ByVal lngOpenFlags As Long = 0, _
    Optional ByVal strDllDirectory As String = vbNullString) As clsSQLiteConnection

    Dim objConn As clsSQLiteConnection
    Set objConn = New clsSQLiteConnection
    objConn.OpenConnection strFilePath, enmProfile, enmDriver, lngOpenFlags, strDllDirectory
    Set OpenSQLite = objConn
End Function

Public Function OpenSQLiteMemory( _
    Optional ByVal enmDriver As SQLiteDriverKind = sqDriverAuto, _
    Optional ByVal strDllDirectory As String = vbNullString) As clsSQLiteConnection

    Dim objConn As clsSQLiteConnection
    Set objConn = New clsSQLiteConnection
    objConn.OpenInMemory enmDriver, strDllDirectory
    Set OpenSQLiteMemory = objConn
End Function

Public Function NewSQLiteTableMap( _
    ByVal strTableName As String, _
    Optional ByVal strPrimaryKey As String = "ID") As clsSQLiteTableMap

    Dim objMap As clsSQLiteTableMap
    Set objMap = New clsSQLiteTableMap
    objMap.InitMap strTableName, strPrimaryKey
    Set NewSQLiteTableMap = objMap
End Function

Public Function NewSQLiteMigrator(ByVal objConn As clsSQLiteConnection) As clsSQLiteMigrator
    Set NewSQLiteMigrator = objConn.NewMigrator()
End Function

Public Function NewSQLiteAccessBridge( _
    ByVal objConn As clsSQLiteConnection, _
    Optional ByVal objDaoDb As DAO.Database = Nothing) As clsSQLiteAccessBridge

    Set NewSQLiteAccessBridge = objConn.NewAccessBridge(objDaoDb)
End Function

Public Function NewSQLiteUnitOfWork(ByVal objConn As clsSQLiteConnection) As clsSQLiteUnitOfWork
    Set NewSQLiteUnitOfWork = objConn.NewUnitOfWork()
End Function

Public Function GetSQLiteEngineVersion( _
    Optional ByVal enmDriver As SQLiteDriverKind = sqDriverAuto, _
    Optional ByVal strDllDirectory As String = vbNullString) As String

    Dim objConn As clsSQLiteConnection
    Set objConn = OpenSQLiteMemory(enmDriver, strDllDirectory)
    GetSQLiteEngineVersion = objConn.Driver.DriverTitle & " v" & objConn.Driver.LibraryVersion
    objConn.CloseConnection
End Function

