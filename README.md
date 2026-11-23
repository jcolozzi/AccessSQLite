# AccessSQLite — VBA Class Library & ORM for Microsoft Access

`AccessSQLite` is a self-contained, 32-bit and 64-bit Microsoft Access VBA class library, lightweight Object-Relational Mapper (ORM), disconnected **ADO in-memory recordset** bridge, and **Monaco Editor SQLite front-end client** (`AccessSQLite Studio`) for SQLite 3. It supports zero-deployment execution via Windows' built-in `winsqlite3.dll` as well as a portable bundled `sqlite3.dll` fallback, automatic SMB/UNC vs. local-disk pragma profiles, prepared statement caching, fluent query building, dual-mode `Dictionary` and strongly-typed `ISQLiteEntity` repositories, zero-temp-table Continuous Forms via disconnected `ADODB.Recordset`s, schema migrations, bidirectional Access DAO & ADO data transfer, and a Ribbon-launchable SQL workbench.

---

## Table of Contents

1. [Directory Layout](#1-directory-layout)
2. [Installation & Referencing Guide](#2-installation--referencing-guide)
3. [Quick-Start Examples](#3-quick-start-examples)
4. [Architecture Overview](#4-architecture-overview)
5. [Driver Layer & Pragma Profiles](#5-driver-layer--pragma-profiles)
6. [Core SQL Engine (`clsSQLiteConnection` & `clsSQLiteStatement`)](#6-core-sql-engine-clssqliteconnection--clssqlitestatement)
7. [Schema Mapping, Validation & Reflection](#7-schema-mapping-validation--reflection)
8. [Fluent Query Builder (`clsSQLiteQueryBuilder`)](#8-fluent-query-builder-clssqlitequerybuilder)
9. [ORM Repositories, Typed Entities, Eager Loading & Unit of Work](#9-orm-repositories-typed-entities-eager-loading--unit-of-work)
10. [Entity Class Generator (`modSQLiteEntityGen`)](#10-entity-class-generator-modsqliteentitygen)
11. [Versioned Schema Migrations (`clsSQLiteMigrator`)](#11-versioned-schema-migrations-clssqlitemigrator)
12. [Access DAO <-> SQLite Bridge (`clsSQLiteAccessBridge`)](#12-access-dao---sqlite-bridge-clssqliteaccessbridge)
13. [In-Memory ADO Recordsets for Continuous Forms (`clsSQLiteAdoBridge`)](#13-in-memory-ado-recordsets-for-continuous-forms-clssqliteadobridge)
14. [AccessSQLite Studio — Monaco SQL Front-End Client & Ribbon](#14-accesssqlite-studio--monaco-sql-front-end-client--ribbon)
15. [Data Type Marshaling & Edge-Case Guarantees](#15-data-type-marshaling--edge-case-guarantees)
16. [Error Handling & Diagnostics](#16-error-handling--diagnostics)
17. [Developer Internals, Testing & Contribution Notes](#17-developer-internals-testing--contribution-notes)
18. [Credits, Acknowledgments & Third-Party Licenses](#18-credits-acknowledgments--third-party-licenses)

---

## 1. Directory Layout

| File / Folder | Purpose |
|---|---|
| [`./AccessSQLite.accdb`](./AccessSQLite.accdb) | Host Access class-library database (VBA Project Name: `AccessSQLite`; public classes set `Attribute VB_Exposed = True`; hosts `frmSQLiteStudio` and `SQLiteStudio` Ribbon) |
| [`./src/AccessSQLite/`](./src/AccessSQLite) | Exported VBA source modules (25 `.bas` and `.cls` modules) from `AccessSQLite.accdb` |
| [`./src/AccessSQLiteDemo/`](./src/AccessSQLiteDemo) | Exported VBA source & form code-behind modules (7 `.bas` and `.cls` modules) from `AccessSQLiteDemo.accdb` |
| [`./web/studio/`](./web/studio) | `AccessSQLite Studio` web frontend (`index.html`, `studio.css`, `bridge.js`, `studio.js`) loaded inside Access's Edge Browser Control (`https://msaccess/`) |
| [`./web/vendor/monaco/`](./web/vendor/monaco) | Bundled local `monaco-editor` (`0.52.2` minified AMD build) — runs 100% offline with SQL syntax highlighting & schema IntelliSense |
| [`./lib/sqlite3.dll`](./lib/sqlite3.dll) | Bundled **32-bit (`x86`)** SQLite `3.45.3` engine used by `clsSQLiteBundledDriver` under 32-bit Microsoft Access |
| [`./lib/SQLite3_StdCall.dll`](./lib/SQLite3_StdCall.dll) | Bundled **32-bit (`x86`)** `__stdcall` thunk wrapper used alongside 32-bit `lib\sqlite3.dll` |
| [`./lib/x64/sqlite3.dll`](./lib/x64/sqlite3.dll) | Bundled **64-bit (`x64`)** SQLite engine isolated in `lib\x64\` for 64-bit Microsoft Access |
| [`./AccessSQLite_Test.db`](./AccessSQLite_Test.db) | Sample/test SQLite database configured with SMB-safe `journal_mode = DELETE` |
| [`./AccessSQLiteDemo.accdb`](./AccessSQLiteDemo.accdb) | Reference demo application showcasing **both** a DAO Staging Continuous Form (`frmTaskList`) and an In-Memory ADO Continuous Form (`frmTaskListAdo`), an unbound ORM Single-Page Form (`frmTaskDetail`), and the `AccessSQLite` CustomUI Ribbon |
| [`./AccessSQLiteDemo.db`](./AccessSQLiteDemo.db) | Live SQLite database backing `AccessSQLiteDemo.accdb` (migrated and seeded via `clsSQLiteMigrator` + `clsSQLiteRepository`) |
| [`./Demo-README.md`](./Demo-README.md) | Architectural walkthrough comparing DAO staging tables vs. In-Memory ADO recordsets vs. Unbound ORM forms + Monaco SQLite Studio |
| [`./README.md`](./README.md) | Developer and usage documentation |

---

## 2. Installation & Referencing Guide

You can use `AccessSQLite` in two ways:

### Option A — Reference `AccessSQLite.accdb` as an External Class Library (Recommended)

1. Open your front-end `.accdb` in Microsoft Access and press **Alt+F11** to open the VBA Editor.
2. Select **Tools -> References...**
3. Ensure **`Microsoft Scripting Runtime`** (`scrrun.dll`, for `Scripting.Dictionary`) and **`Microsoft Office 16.0 Access database engine Object Library`** (`DAO`) are checked.
4. Click **Browse...**, change the file filter to **Microsoft Access Databases (`*.accdb`)**, and select `AccessSQLite.accdb`.
5. Verify **`AccessSQLite`** appears checked in your References list, then click **Debug -> Compile**.

> **Why `modSQLiteFactory` exists:** Office VBA class modules in referenced libraries expose `PublicNotCreatable` (`Attribute VB_Exposed = True`, `Attribute VB_Creatable = False`). External `.accdb` projects declare variables with full early-bound IntelliSense (`Dim db As AccessSQLite.clsSQLiteConnection`) and instantiate them through `AccessSQLite.OpenSQLite(...)`, `AccessSQLite.OpenSQLiteMemory(...)`, or `AccessSQLite.NewSQLiteTableMap(...)`.

### Option B — Import Modules Directly into a Single-File `.accdb`

Copy or import the standard modules (`modSQLiteTypes`, `modSQLiteFactory`, `modSQLiteEntityGen`) and class modules (`ISQLiteDriver`, `clsSQLiteWinDriver`, `clsSQLiteBundledDriver`, `clsSQLiteStatement`, `clsSQLiteConnection`, `clsSQLiteColumnDef`, `clsSQLiteTableMap`, `clsSQLiteSchema`, `clsSQLiteQueryBuilder`, `ISQLiteEntity`, `clsSQLiteRepository`, `clsSQLiteUnitOfWork`, `clsSQLiteMigrator`, `clsSQLiteAccessBridge`) directly into your database. If you only target Windows 10/11+, `winsqlite3.dll` requires **zero external DLL files**.

---

## 3. Quick-Start Examples

### Raw Parameterized SQL (`clsSQLiteConnection`)

```vba
Public Sub QuickRawSqlDemo()
    Dim db As AccessSQLite.clsSQLiteConnection
    Dim rows As Collection
    Dim row As Scripting.Dictionary

    ' Auto-selects winsqlite3.dll and chooses WAL (local disk) or DELETE (UNC/SMB share)
    Set db = AccessSQLite.OpenSQLite(CurrentProject.Path & "\AppData.db")

    db.ExecuteScript _
        "CREATE TABLE IF NOT EXISTS Customers (" & _
        "  CustomerID INTEGER PRIMARY KEY AUTOINCREMENT," & _
        "  CustomerName TEXT NOT NULL," & _
        "  CreditLimit CURRENCY NOT NULL DEFAULT 0," & _
        "  IsActive BOOLEAN NOT NULL DEFAULT 1" & _
        ");"

    ' Positional (?) parameters
    db.ExecuteSql "INSERT INTO Customers (CustomerName, CreditLimit, IsActive) VALUES (?, ?, ?);", _
                  "Northwind Traders", 25000.5@, True

    ' Named (@key) parameter pairs — familiar to DAO/clsDb users
    Set rows = db.QueryRowsRaw( _
        "SELECT * FROM Customers WHERE IsActive = @active AND CreditLimit >= @minLimit ORDER BY CustomerName;", _
        "@active", True, "@minLimit", 10000@)

    For Each row In rows
        Debug.Print row("CustomerID"), row("CustomerName"), row("CreditLimit"), row("IsActive")
    Next row

    db.CloseConnection
End Sub
```

### Schema Definition, Repository CRUD & Fluent Query Builder

```vba
Public Sub QuickOrmDemo()
    Dim db As AccessSQLite.clsSQLiteConnection
    Dim mapTasks As AccessSQLite.clsSQLiteTableMap
    Dim repoTasks As AccessSQLite.clsSQLiteRepository
    Dim payload As Scripting.Dictionary
    Dim newTaskId As Long
    Dim dueSoon As Collection

    Set db = AccessSQLite.OpenSQLiteMemory()

    ' 1. Define schema with validation rules, auto-timestamps, and soft-delete support
    Set mapTasks = db.DefineTable("Tasks", "TaskID") _
        .AddPrimaryKey("TaskID", sqColInt64, True) _
        .AddColumn("TitleText", sqColText, "Task title", blnRequired:=True, lngMaxLength:=120) _
        .AddColumn("PriorityRank", sqColLong, "Priority", blnRequired:=True, varDefault:=3) _
        .AddColumn("EstimateCost", sqColCurrency, "Estimated cost", varDefault:=0@) _
        .AddColumn("IsCompleted", sqColBoolean, "Completed", varDefault:=False) _
        .AddColumn("DueDate", sqColDate, "Due date") _
        .AddTimestamps("CreatedAt", "UpdatedAt") _
        .AddSoftDelete("DeletedAt") _
        .AddIndex("idx_Tasks_Priority", "PriorityRank ASC")

    db.SchemaManager().EnsureTable mapTasks
    Set repoTasks = db.Repository("Tasks")

    ' 2. Insert via Dictionary (auto-validates, coerces types, and stamps CreatedAt/UpdatedAt)
    Set payload = AccessSQLite.NewTextDict()
    payload("TitleText") = "Prepare Q4 financial export"
    payload("PriorityRank") = 1
    payload("EstimateCost") = 850.75@
    payload("DueDate") = DateSerial(2026, 12, 15)
    newTaskId = CLng(repoTasks.InsertRecord(payload))

    ' 3. Query with the fluent parameterized QueryBuilder
    Set dueSoon = repoTasks.NewQuery() _
        .WhereEq("IsCompleted", False) _
        .WhereOp("PriorityRank", "<=", 2) _
        .OrderByCol("PriorityRank") _
        .Page(1, 20) _
        .ToRowsRaw()

    Debug.Print "Inserted TaskID=" & newTaskId & ", Matching Active Tasks=" & dueSoon.Count
    db.CloseConnection
End Sub
```

---

## 4. Architecture Overview

`AccessSQLite` is organized into four clean layers plus a public class-library entry module:

```text
+-----------------------------------------------------------------------------------+
| Public Library Surface (Tools -> References -> AccessSQLite.accdb)                |
|   modSQLiteFactory  |  modSQLiteTypes                                             |
+-----------------------------------------------------------------------------------+
| Layer 4: ORM, Query Builder, Migrations & Access DAO Bridge                       |
|   ISQLiteEntity        clsSQLiteRepository    clsSQLiteUnitOfWork                 |
|   clsSQLiteQueryBuilder clsSQLiteMigrator     clsSQLiteAccessBridge               |
|   modSQLiteEntityGen                                                              |
+-----------------------------------------------------------------------------------+
| Layer 3: Schema Metadata, Validation & Reflection                                 |
|   clsSQLiteColumnDef   clsSQLiteTableMap      clsSQLiteSchema                     |
+-----------------------------------------------------------------------------------+
| Layer 2: Connection Management, Prepared Statement Cache & Savepoints             |
|   clsSQLiteConnection  clsSQLiteStatement                                         |
+-----------------------------------------------------------------------------------+
| Layer 1: Pluggable Native C-API Drivers (32-bit & 64-bit VBA7)                    |
|   ISQLiteDriver        clsSQLiteWinDriver     clsSQLiteBundledDriver              |
+-----------------------------------------------------------------------------------+
```

---

## 5. Driver Layer & Pragma Profiles

### Driver Selection (`SQLiteDriverKind`)

| Enum Value | Driver Class | How It Works |
|---|---|---|
| `sqDriverAuto` *(default)* | `clsSQLiteWinDriver` -> `clsSQLiteBundledDriver` | Tries `winsqlite3.dll` first; automatically falls back to the bundled DLLs if `winsqlite3.dll` is unavailable. |
| `sqDriverWinSqlite` | `clsSQLiteWinDriver` | Uses `C:\Windows\SysWOW64\winsqlite3.dll` (32-bit Access) or `C:\Windows\System32\winsqlite3.dll` (64-bit Access). Microsoft compiles `winsqlite3.dll` with `__stdcall` and undecorated exports on both architectures — zero external DLL deployment needed. |
| `sqDriverBundled` | `clsSQLiteBundledDriver` | Resolves DLLs relative to `GetBundledDllDirectory()` (`<CodeProject.Path>\lib`, or a custom `strDllDirectory`): loads `lib\sqlite3.dll` + `lib\SQLite3_StdCall.dll` on 32-bit Access (bridging newer C exports like `sqlite3_column_decltype16` and `sqlite3_expanded_sql` via `DispCallFunc`) or `lib\x64\sqlite3.dll` on 64-bit Access. |

> **Important `winsqlite3.dll` Note (`OMIT_LOCALTIME`):** Microsoft builds `winsqlite3.dll` with `OMIT_LOCALTIME`, so SQL expressions must not call `datetime('now', 'localtime')`. Use `UtcNowIsoString()` / VBA `Now` for local timestamps or SQLite `datetime('now')` for UTC.

### Automatic Pragma Profiles (`SQLitePragmaProfile`)

SQLite's `WAL` mode relies on memory-mapped `-shm` shared memory and can corrupt databases placed on multi-user SMB/UNC network shares (`K:\...` or `\\server\share\...`). `OpenSQLite` inspects the target path via Win32 `GetDriveTypeW` and applies safe defaults automatically:

| Enum Value | Pragmas Applied | When Selected by `sqProfileAuto` |
|---|---|---|
| `sqProfileLocalWal` | `journal_mode=WAL`, `synchronous=NORMAL`, `foreign_keys=ON`, `temp_store=MEMORY`, `busy_timeout=5000` | Local fixed/SSD drives (`C:\...`) |
| `sqProfileNetworkSafe` | `journal_mode=DELETE`, `synchronous=FULL`, `foreign_keys=ON`, `temp_store=MEMORY`, `busy_timeout=5000` | UNC paths (`\\...`) and mapped remote drives (`DRIVE_REMOTE`, e.g., `K:\...`) |
| `sqProfileReadOnly` | `query_only=ON`, `temp_store=MEMORY`, `busy_timeout=5000` | When `sqOpenReadOnly` flag or `sqProfileReadOnly` is passed |
| `sqProfileInMemory` | `foreign_keys=ON`, `temp_store=MEMORY`, `busy_timeout=5000` | `:memory:` or `sqOpenMemory` connections |

---

## 6. Core SQL Engine (`clsSQLiteConnection` & `clsSQLiteStatement`)

### Opening & Closing Connections

```vba
' File-backed connection (creates file if absent unless sqProfileReadOnly)
Dim db As AccessSQLite.clsSQLiteConnection
Set db = AccessSQLite.OpenSQLite("C:\Data\Inventory.db")

' Explicit driver or read-only mode
Set db = AccessSQLite.OpenSQLite("K:\Shared\Archive.db", sqProfileReadOnly, sqDriverBundled)

' Fast isolated in-memory database
Set db = AccessSQLite.OpenSQLiteMemory()

' Always close when finished (also auto-closed in Class_Terminate)
db.CloseConnection
```

### Parameter Binding Styles

Every query and execution method on `clsSQLiteConnection` routes through `clsSQLiteStatement.BindFlexibleParams`, supporting four parameter styles interchangeably:

1. **Positional `?` parameters:**
   ```vba
   db.ExecuteSql "UPDATE Products SET UnitPrice = ? WHERE ProductID = ?;", 19.95@, 42
   ```
2. **Named `@key` / `:key` / `$key` pairs in `ParamArray`:**
   ```vba
   Set row = db.QueryOneRow("SELECT * FROM Products WHERE ProductCode = @code;", "@code", "NW-100")
   ```
3. **Named `Scripting.Dictionary` (`ExecuteNamed` or passed directly):**
   ```vba
   Dim p As Scripting.Dictionary
   Set p = AccessSQLite.NewTextDict()
   p("code") = "NW-100"   ' Matches @code, :code, or $code automatically
   Set row = db.QueryOneRow("SELECT * FROM Products WHERE ProductCode = @code;", p)
   ```
4. **Positional `Collection` or `Variant()` array:**
   ```vba
   Dim args As New Collection
   args.Add 10: args.Add 20
   Set rows = db.QueryRows("SELECT * FROM Products WHERE ProductID IN (?, ?);", args)
   ```

> **Positional `@` Disambiguation:** If a SQL statement uses positional `?` placeholders and your data value starts with `"@"` (e.g., `"@johndoe"`), `ResolveNamedParamIndex` verifies against the compiled statement's parameter metadata first so `"@johndoe"` is bound safely as a positional string value rather than misread as a parameter name.

### Query Methods Summary (`clsSQLiteConnection`)

| Method | Return Type | Description |
|---|---|---|
| `ExecuteSql(strSql, ParamArray varParams())` | `Long` | Runs `INSERT`/`UPDATE`/`DELETE`/DDL and returns `ChangesCount` (rows affected by this statement). |
| `ExecuteNamed(strSql, dictParams)` | `Long` | Runs a statement with parameters from a `Scripting.Dictionary`. |
| `ExecuteScript(strMultiSql)` | `Long` | Executes multi-statement SQL scripts (`;`-delimited DDL/seeds) via SQLite's native UTF-16 tail pointer and returns the number of statements executed. |
| `QueryRows(strSql, ParamArray varParams())` | `Collection` | Returns a `Collection` of case-insensitive `Scripting.Dictionary` rows; dates remain JSON-ready ISO-8601 strings. |
| `QueryRowsRaw(strSql, ParamArray varParams())` | `Collection` | Returns `Scripting.Dictionary` rows hydrated to native VBA `Boolean`, `Currency`, and `Date` types based on declared column types. |
| `QueryOneRow` / `QueryOneRowRaw` | `Scripting.Dictionary` | Returns the first matching row, or `Nothing` when 0 rows match. |
| `QueryScalar(strSql, ParamArray varParams())` | `Variant` | Returns column `0` of row `1`, or `Null` when 0 rows match. |
| `LastInsertRowId` | `Variant` | Returns `Long` (or `Currency`/`LongLong` when `> 2,147,483,647`) of the most recent `INSERT`. |

### Prepared Statement Cache

- Enabled by default (`db.StatementCacheEnabled = True`, capacity `128` statements).
- Single-statement `SELECT`, `WITH`, `INSERT`, `UPDATE`, `DELETE`, and `REPLACE` strings are prepared with `SQLITE_PREPARE_PERSISTENT` and reused across calls, automatically calling `sqlite3_reset` and `sqlite3_clear_bindings` after every run (including on error).
- Calling `db.VacuumDb()`, changing `db.DateStorage`, or calling `db.ClearStatementCache()` finalizes and clears cached handles cleanly.

### Transactions & Nested Savepoints

`BeginTrans`, `CommitTrans`, and `RollbackTrans` support arbitrary nesting:
- **Depth `0 -> 1`:** Issues `BEGIN IMMEDIATE TRANSACTION;` (acquires the reserved write lock upfront instead of failing mid-transaction with `SQLITE_BUSY` lock-upgrade deadlocks).
- **Depth `1 -> N`:** Issues `SAVEPOINT sp_orm_<N>;`.
- **`RollbackTrans` at depth `> 1`:** Rolls back only the inner savepoint (`ROLLBACK TO SAVEPOINT sp_orm_<N>; RELEASE SAVEPOINT sp_orm_<N>;`), leaving outer work intact.

```vba
db.BeginTrans
On Error GoTo TxError
    db.ExecuteSql "INSERT INTO Orders (CustomerName) VALUES (?);", "Alpine Ski House"
    
    db.BeginTrans ' Creates SAVEPOINT sp_orm_2
        db.ExecuteSql "INSERT INTO AuditNotes (NoteText) VALUES (?);", "Optional step"
    db.RollbackTrans ' Discards only sp_orm_2; Alpine Ski House insert stays staged!

db.CommitTrans
Exit Sub

TxError:
    db.RollbackTrans
    Err.Raise Err.Number, Err.Source, Err.Description
```

### Maintenance & SQL Timing Trace

- **`db.StartTrace` / `db.StopTrace() As String`:** Records sub-millisecond `QueryPerformanceCounter` durations alongside `sqlite3_expanded_sql` (with bound values interpolated for debugging).
- **`db.BackupToFile "C:\Backups\AppData_Backup.db"`:** Uses the native SQLite Online Backup C API (`sqlite3_backup_init` / `step` / `finish`) to snapshot a live open database safely.
- **`db.IntegrityCheck() As Boolean`:** Runs `PRAGMA integrity_check;` and returns `True` when `"ok"`.
- **`db.VacuumDb`:** Clears cached statements and runs `VACUUM;` (must be called outside a transaction).

---

## 7. Schema Mapping, Validation & Reflection

### Logical Column Types (`SQLiteColumnKind`)

| `SQLiteColumnKind` | Standard SQLite DDL | `STRICT` Mode DDL | Coerced VBA Type |
|---|---|---|---|
| `sqColText` | `TEXT` | `TEXT` | `String` (enforces `lngMaxLength` when `> 0`) |
| `sqColMemo` | `TEXT` | `TEXT` | `String` (unbounded length) |
| `sqColLong` | `INTEGER` | `INTEGER` | `Long` (`-2,147,483,648 .. 2,147,483,647`) |
| `sqColInt64` | `INTEGER` | `INTEGER` | `Long` within 32-bit range; `Currency` / `LongLong` above 32-bit |
| `sqColDouble` | `REAL` | `REAL` | `Double` |
| `sqColCurrency` | `CURRENCY` | `REAL` | `Currency` (4-decimal fixed-point in VBA) |
| `sqColBoolean` | `BOOLEAN` | `INTEGER` | `Boolean` (stored as `1` / `0` in SQLite) |
| `sqColDate` | `DATE` | `TEXT` | `Date` (date-only `Int(d)`, default ISO `"yyyy-mm-dd"`) |
| `sqColDateTime` | `DATETIME` | `TEXT` | `Date` (date+time, default ISO `"yyyy-mm-ddThh:nn:ss"`) |
| `sqColBlob` | `BLOB` | `BLOB` | `Byte()` |
| `sqColJson` | `JSON` | `TEXT` | `String` |
| `sqColGuid` | `GUID` | `TEXT` | `String` |

### Building a `clsSQLiteTableMap`

```vba
Dim mapEmp As AccessSQLite.clsSQLiteTableMap
Set mapEmp = db.DefineTable("Employees", "EmployeeID") _
    .SetStrict(True) _
    .AddPrimaryKey("EmployeeID", sqColInt64, blnAutoIncrement:=True) _
    .AddColumn("BadgeCode", sqColText, "Badge code", blnRequired:=True, lngMaxLength:=16, blnUnique:=True) _
    .AddColumn("FullName", sqColText, "Full name", blnRequired:=True, lngMaxLength:=100) _
    .AddColumn("DepartmentID", sqColLong, "Department", blnRequired:=True, _
               strFkTable:="Departments", strFkColumn:="DepartmentID", strOnDelete:="RESTRICT") _
    .AddColumn("HourlyRate", sqColCurrency, "Hourly rate", varDefault:=0@) _
    .AddColumn("HireDate", sqColDate, "Hire date") _
    .AddTimestamps("CreatedAt", "UpdatedAt") _
    .AddSoftDelete("DeletedAt") _
    .AddIndex("idx_Employees_Dept", "DepartmentID")

' Creates table + indexes if missing, or runs additive ALTER TABLE ADD COLUMN for new mapped columns
db.SchemaManager().EnsureTable mapEmp, blnSyncAdditiveColumns:=True
```

### Reverse-Engineering Existing SQLite Tables (`ReflectTableMap`)

If you open a pre-existing SQLite database without calling `DefineTable`, calling `db.Repository("ExistingTable")` or `db.SchemaManager().ReflectTableMap("ExistingTable")` automatically inspects `PRAGMA table_xinfo` and `PRAGMA foreign_key_list` to construct a `clsSQLiteTableMap` (including `STRICT` detection and conventional `CreatedAt` / `UpdatedAt` / `DeletedAt` columns).

---

## 8. Fluent Query Builder (`clsSQLiteQueryBuilder`)

Create a builder via `db.FromTable("TableName", Optional "alias")` or `repo.NewQuery(Optional "alias")` (the repository version automatically attaches the `clsSQLiteTableMap` for native type hydration and soft-delete filtering).

### Clause Reference

- **Projections:**
  - `.SelectCols("EmployeeID, FullName", "HourlyRate")` — quotes every column/qualified `alias.col` safely
  - `.SelectAliased("e.FullName", "EmployeeName")`
  - `.SelectRaw("COUNT(*) AS TotalStaff")` *(only pass trusted SQL expressions, never user input)*
  - `.DistinctRows()`
- **Joins:**
  - `.InnerJoinTable("Departments", "e.DepartmentID", "d.DepartmentID", "d")`
  - `.LeftJoinTable("Departments", "e.DepartmentID", "d.DepartmentID", "d")`
  - `.JoinRaw(...)`
- **Filtering (all values are bound via `?` parameters):**
  - `.WhereEq(col, val)` / `.WhereNotEq(col, val)` *(automatically emits `IS NULL` / `IS NOT NULL` when `val` is `Null` or `Empty`)*
  - `.WhereOp(col, op, val)` — supports `=`, `<>`, `!=`, `<`, `<=`, `>`, `>=`, `LIKE`, `NOT LIKE`, `GLOB`
  - `.WhereIn(col, arrayOrCollection)` / `.WhereNotIn(col, arrayOrCollection)` *(safely emits `1 = 0` / `1 = 1` when passed an empty collection)*
  - `.WhereBetween(col, lowVal, highVal)`
  - `.WhereLike(col, rawPattern)`
  - `.WhereContains(col, userSearchText)` *(escapes `\`, `%`, and `_` with `ESCAPE '\'` so literal `%` or `_` characters in user search boxes never act as wildcards)*
  - `.WhereNull(col)` / `.WhereNotNull(col)`
  - `.WhereRaw("e.HourlyRate * 40 > ?", 2000@)`
  - `.WithTrashed()` / `.OnlyTrashed()` *(controls soft-delete filtering when `AddSoftDelete` is configured)*
- **Grouping, Ordering & Pagination:**
  - `.GroupByCols("DepartmentID")` / `.HavingRaw("COUNT(*) >= ?", 5)`
  - `.OrderByCol("FullName")` / `.OrderByDesc("HourlyRate")`
  - `.LimitRows(50)` / `.OffsetRows(100)` / `.Page(lngPageNumber, lngPageSize)` *(1-based page number)*
- **Terminal Execution Methods:**
  - `.ToRows()` / `.ToRowsRaw()` -> `Collection` of `Scripting.Dictionary`
  - `.FirstRow()` / `.FirstRowRaw()` -> `Scripting.Dictionary` (or `Nothing`)
  - `.ToEntities(New clsMyEntity)` / `.FirstEntity(New clsMyEntity)` -> strongly-typed `ISQLiteEntity` instances
  - `.CountRows()` -> `Long`
  - `.ExistsAny()` -> `Boolean` (`SELECT EXISTS(... LIMIT 1)`)
  - `.PluckCol("BadgeCode")` -> `Collection` of single-column scalar values
  - `.ScalarVal(Optional strSelectExpr)` -> `Variant`
  - `.UpdateValues(dictSetValues)` -> `Long` (parameterized bulk `UPDATE` + auto `UpdatedAt`)
  - `.DeleteMatching(Optional blnForceHardDelete = False)` -> `Long` (soft-deletes when configured unless `blnForceHardDelete:=True`)
  - `.ToSql()` / `.BoundParameters()` -> inspect generated SQL and bound parameter list

---

## 9. ORM Repositories, Typed Entities, Eager Loading & Unit of Work

`clsSQLiteRepository` (`Set repo = db.Repository("TableName")`) supports **two working styles** side by side:

### Mode 1 — `Scripting.Dictionary` Payloads (Zero Boilerplate)

- `FindById(varId, Optional blnIncludeTrashed = False) As Scripting.Dictionary`
- `FindByIdRaw(varId, Optional blnIncludeTrashed = False) As Scripting.Dictionary`
- `GetOrFail(varId) As Scripting.Dictionary` *(raises `sqErrNotFound` if missing)*
- `ExistsById(varId, Optional blnIncludeTrashed = False) As Boolean`
- `CountAll(Optional blnIncludeTrashed = False) As Long`
- `ListAll` / `ListAllRaw`
- `InsertRecord(dictValues, Optional blnEnforceClientEditable = False) As Variant`
- `InsertMany(colRecords As Collection) As Long` *(transactional batch insert)*
- `UpdateRecord(varId, dictValues, Optional blnDirtyOnly = False, Optional blnEnforceClientEditable = False) As Long`
  - When `blnDirtyOnly:=True`, fetches the stored row, diffs coerced values via `DiffChangedColumns`, and **skips the `UPDATE` completely** if nothing changed.
  - When `blnEnforceClientEditable:=True` (useful for web/JSON form payloads), rejects writes to server-managed columns (`IsClientEditable = False`, such as PKs and audit stamps).
- `UpsertRecord("UniqueCol1, UniqueCol2", dictValues) As Variant` *(emits `INSERT ... ON CONFLICT (...) DO UPDATE SET ...`, preserving `CreatedAt`)*
- `DeleteRecord(varId, Optional blnForceHardDelete = False) As Long`
- `RestoreRecord(varId) As Long` *(clears `DeletedAt` back to `NULL`)*

### Mode 2 — Strongly-Typed VBA Classes (`Implements ISQLiteEntity`)

Because VBA lacks runtime reflection (`Activator.CreateInstance`) for private class modules, `ISQLiteEntity` uses the lightweight **Prototype Factory pattern** (`CreateBlank`):

```vba
' Class Module: clsDemoTask (see AccessSQLite.accdb for the complete working class)
Option Compare Database
Option Explicit

Implements ISQLiteEntity

Public TaskID As Long
Public ProjectID As Long
Public TitleText As String
Public EstimateCost As Currency
Public IsCompleted As Boolean

Private Property Get ISQLiteEntity_EntityTableName() As String
    ISQLiteEntity_EntityTableName = "DemoTasks"
End Property

Private Function ISQLiteEntity_CreateBlank() As ISQLiteEntity
    Set ISQLiteEntity_CreateBlank = New clsDemoTask
End Function

Private Sub ISQLiteEntity_ReadFromRow(ByVal dictRow As Scripting.Dictionary)
    TaskID = CLng(Nz(dictRow("TaskID"), 0))
    ProjectID = CLng(Nz(dictRow("ProjectID"), 0))
    TitleText = CStr(Nz(dictRow("TitleText"), vbNullString))
    EstimateCost = CCur(Nz(dictRow("EstimateCost"), 0@))
    IsCompleted = CBool(Nz(dictRow("IsCompleted"), False))
End Sub

Private Sub ISQLiteEntity_WriteToRow(ByVal dictRow As Scripting.Dictionary)
    dictRow("TaskID") = TaskID
    dictRow("ProjectID") = ProjectID
    dictRow("TitleText") = TitleText
    dictRow("EstimateCost") = EstimateCost
    dictRow("IsCompleted") = IsCompleted
End Sub
```

Using typed entities with `clsSQLiteRepository`:

```vba
Dim task As New clsDemoTask
task.ProjectID = 1
task.TitleText = "Implement Invoice Export"
task.EstimateCost = 600@

' Inserts new entity (TaskID = 0) and hydrates generated TaskID + CreatedAt back onto `task`
repoTasks.SaveEntity task

' Later mutation: diffs against stored row and updates ONLY changed columns + UpdatedAt
task.IsCompleted = True
repoTasks.SaveEntity task, blnDirtyOnly:=True

' Hydrate a single entity or a filtered Collection of clsDemoTask instances
Dim loaded As clsDemoTask
Set loaded = repoTasks.FindEntity(task.TaskID, New clsDemoTask)

Dim openTasks As Collection
Set openTasks = repoTasks.NewQuery().WhereEq("IsCompleted", False).ToEntities(New clsDemoTask)
```

### Batch Eager Relationship Loading (Zero `N+1` Queries)

Instead of running a query inside a `For Each` loop over parent rows, `IncludeHasMany` and `IncludeBelongsTo` collect distinct foreign keys, chunk `WHERE ... IN (...)` queries in batches of `900` parameters (well within SQLite's parameter limit), and attach related rows in memory in **2 total SQL queries**:

```vba
Dim colProjects As Collection
Dim proj As Scripting.Dictionary
Dim taskRow As Scripting.Dictionary

Set colProjects = db.Repository("Projects").ListAllRaw("ProjectName")

' Attaches a Collection of child task Dictionaries at proj("Tasks") in 1 batched SELECT
db.Repository("Projects").IncludeHasMany _
    colParentRows:=colProjects, _
    strChildTable:="DemoTasks", _
    strForeignKeyCol:="ProjectID", _
    strAttachAs:="Tasks", _
    strLocalKeyCol:="ProjectID", _
    strOrderByChildCol:="PriorityRank"

For Each proj In colProjects
    Debug.Print proj("ProjectTitle") & " (" & proj("Tasks").Count & " tasks)"
    For Each taskRow In proj("Tasks")
        Debug.Print "  - " & taskRow("TitleText")
    Next taskRow
Next proj
```

Similarly, `IncludeBelongsTo` attaches each child row's parent `Scripting.Dictionary` (e.g., `taskRow("Project")`):

```vba
db.Repository("DemoTasks").IncludeBelongsTo colTasks, "Projects", "ProjectID", "Project"
```

### Atomic Multi-Repository Batches (`clsSQLiteUnitOfWork`)

`clsSQLiteUnitOfWork` queues inserts, updates, deletes, entity saves, and custom SQL statements across multiple tables and commits them inside a single transaction/savepoint — rolling back every change automatically if any validation rule or constraint fails:

```vba
Dim uow As AccessSQLite.clsSQLiteUnitOfWork
Set uow = db.NewUnitOfWork()

uow.RegisterInsert "Projects", dictNewProject
uow.RegisterEntitySave objFirstTask
uow.RegisterUpdate "Counters", 1, dictCounterUpdate
uow.RegisterDelete "StagingImports", 99, blnForceHardDelete:=True

uow.CommitWork ' Atomic: all 4 succeed or all 4 roll back
```

---

## 10. Entity Class Generator (`modSQLiteEntityGen`)

You can auto-generate complete, ready-to-save VBA class module code (`Implements ISQLiteEntity`) from any mapped or existing SQLite table. Columns named after Access/VBA/SQL reserved words (such as `Date`, `Value`, `Name`, `Select`, `Count`, `Order`, `Index`, `Type`) are automatically suffixed with `Value` (`DateValue`, `SelectValue`) in VBA identifiers while preserving exact column names in `ReadFromRow` / `WriteToRow`:

```vba
Dim db As AccessSQLite.clsSQLiteConnection
Set db = AccessSQLite.OpenSQLite(CurrentProject.Path & "\AppData.db")
Debug.Print AccessSQLite.GenerateEntityClassCode(db, "Orders", "clsOrder")
db.CloseConnection
```

---

## 11. Versioned Schema Migrations (`clsSQLiteMigrator`)

`clsSQLiteMigrator` tracks schema evolution in `_schema_migrations` (`VersionNum`, `MigrationTitle`, `ChecksumHex`, `AppliedAtIso`, `DurationMs`). Before applying new migrations, `VerifyChecksums` computes a deterministic 32-bit FNV-1a hash of each registered `UpSql` script and raises `sqErrMigration` if an already-applied migration script was altered after deployment.

```vba
Dim mig As AccessSQLite.clsSQLiteMigrator
Set mig = db.NewMigrator()

mig.AddMigration 1, "Create Categories", _
    "CREATE TABLE Categories (CategoryID INTEGER PRIMARY KEY, CategoryName TEXT NOT NULL UNIQUE);", _
    "DROP TABLE Categories;"

mig.AddMigration 2, "Create Items", _
    "CREATE TABLE Items (" & _
    "  ItemID INTEGER PRIMARY KEY AUTOINCREMENT," & _
    "  CategoryID INTEGER NOT NULL REFERENCES Categories(CategoryID)," & _
    "  SKU TEXT NOT NULL UNIQUE," & _
    "  UnitCost CURRENCY NOT NULL DEFAULT 0" & _
    ");", _
    "DROP TABLE Items;"

mig.MigrateUp                  ' Applies pending migrations in ascending order (each in a savepoint)
Debug.Print mig.CurrentVersion ' -> 2

' Optional rollback during development/testing:
' mig.RollbackLast 1           ' Runs v2 DownSql and removes v2 from _schema_migrations
```

---

## 12. Access DAO <-> SQLite Bridge (`clsSQLiteAccessBridge`)

`clsSQLiteAccessBridge` moves tables and queries between Access (`DAO.Database`) and SQLite at prepared-statement transaction speeds while fixing common pitfalls in legacy ports:

- **No `Null`-binding bleed:** Resets and clears statement bindings before every row so `Null` columns never inherit the previous row's bound value.
- **`> 255`-character Memo preservation:** Inspects incoming string lengths and `DECLTYPE` metadata, automatically creating `dbMemo` instead of truncating at `dbText(255)`.
- **`0`-row schema fidelity:** `ImportSQLiteToAccess` / `ImportQueryToAccess` creates the destination Access `DAO.TableDef` from SQLite column metadata even when `0` rows match the `SELECT` filter.

```vba
Dim bridge As AccessSQLite.clsSQLiteAccessBridge
Set bridge = db.NewAccessBridge() ' Defaults to CurrentDb; or pass an open back-end DAO.Database

' 1. Export an Access table or SELECT query into SQLite
Dim exportedCount As Long
exportedCount = bridge.ExportAccessToSQLite("tblOrders", "OrdersArchive", blnOverwriteTable:=True)

' 2. Import an entire SQLite table into a local Access table
Dim importedCount As Long
importedCount = bridge.ImportSQLiteToAccess("OrdersArchive", "tblOrdersLocal", blnOverwriteTable:=True)

' 3. Import a parameterized SQLite SELECT query into Access
importedCount = bridge.ImportQueryToAccess( _
    "SELECT * FROM OrdersArchive WHERE OrderTotal >= ? AND RegionCode = ?;", _
    "tblHighValueOrders", True, 5000@, "MIDWEST")
```

---

## 13. In-Memory ADO Recordsets for Continuous Forms (`clsSQLiteAdoBridge`)

In addition to local DAO staging tables (`clsSQLiteAccessBridge`), `AccessSQLite` provides **`clsSQLiteAdoBridge`** to materialize any SQLite query, `clsSQLiteQueryBuilder`, or `clsSQLiteRepository` into a **disconnected client-side in-memory `ADODB.Recordset`** (`adUseClient` + `adOpenStatic` + `adLockOptimistic`, `ActiveConnection = Nothing`).

Binding an Access Continuous Form via `Set Me.Recordset = rs` eliminates local staging tables (`DELETE`/`INSERT` front-end `.accdb` bloat and multi-user lock contention) and enables **virtual client-only selection checkboxes** per row in RAM (such as `[IsSelected]`, where nothing is persisted to the front-end `.accdb`).

> **Zero extra VBA Reference required:** `clsSQLiteAdoBridge` creates `ADODB.Recordset` instances via late-bound COM (`CreateObject("ADODB.Recordset")`), so referencing front-end `.accdb` projects do not even need to add a reference to `Microsoft ActiveX Data Objects` unless they prefer early-bound `ADODB.Recordset` declarations.

### Querying SQLite into an In-Memory `ADODB.Recordset`

```vba
' 1. Directly from clsSQLiteConnection (positional or @named parameters):
Dim rs As Object
Set rs = db.QueryAdoRecordset( _
    "SELECT TaskID, TaskTitle, EstimateCost, IsCompleted FROM Tasks WHERE CategoryName = ?;", _
    "Engineering")

' 2. From clsSQLiteQueryBuilder WITH a virtual in-memory [IsSelected] checkbox column:
Dim rsTasks As Object
Set rsTasks = db.Repository("Tasks").NewQuery() _
    .WhereEq("IsCompleted", False) _
    .OrderByCol("TaskID") _
    .ToAdoRecordset(blnIncludeSelectColumn:=True, strSelectColumnName:="IsSelected")

' Bind directly to an Access Continuous Form (Form.RecordSource = "")
Set Me.Recordset = rsTasks
```

### Inline Writeback & Virtual Multi-Select Checkboxes

```vba
' Inside Form_AfterUpdate on an ADO-bound Continuous Form:
' Validates and saves dirty columns (or inserts a new row) straight back to SQLite!
db.Repository("Tasks").SaveFromAdoRow Me.Recordset, blnDirtyOnly:=True

' Read all checked primary keys from the virtual [IsSelected] column (clone-safe — does not move active row):
Dim adoBridge As AccessSQLite.clsSQLiteAdoBridge
Dim selectedIds As Collection
Set adoBridge = db.NewAdoBridge()
Set selectedIds = adoBridge.GetSelectedKeys(Me.Recordset, "TaskID", "IsSelected")

' Check or uncheck every row in RAM without re-querying SQLite:
adoBridge.SetAllSelected Me.Recordset, True, "IsSelected"
Me.Refresh

' Bulk-import any external ADODB.Recordset (SQL Server, Excel, CSV, in-memory) into SQLite:
adoBridge.ImportAdoRecordsetToSQLite rsExternal, "ImportedRows", blnOverwriteTable:=True
```

---

## 14. AccessSQLite Studio — Monaco SQL Front-End Client & Ribbon

`AccessSQLite` includes a built-in, Ribbon-launchable SQLite front-end workbench (**`AccessSQLite Studio`**) hosted inside Access's native Edge Browser Control (`ctlWebHost`, `https://msaccess/` virtual origin) and powered offline by **Monaco Editor `0.52.2`**.

### Key Studio Capabilities
- **Schema Explorer:** Browse all **Tables** (columns, `PK` / `FK` / `NN` / `STRICT` badges, row counts, indexes), **Views**, and **Triggers**. Click any column to insert its quoted identifier at the cursor, or click **`Rows`**, **`Ins`**, or **`DDL`** to inspect data or generate SQL templates in a new tab.
- **Stored Queries (`_access_saved_queries`):** Save reusable SQL scripts with category and description metadata (**Ctrl+S** or **Save Query**), filter them in the sidebar, and open or execute them in one click.
- **Monaco SQL Workbench:** Multi-tab SQL editor with live schema-aware IntelliSense (tables, views, columns, `PRAGMA`s, and built-in SQLite functions), **Ctrl+Enter / F5** execution (executes highlighted selection or full active tab), dark/light theme toggle, and persistent transaction indicator (`BEGIN` / `COMMIT` / `ROLLBACK`).
- **Results, Messages, Query Plan & History:** Sortable data grid with `NULL` badges and row-limit selector (`200` / `1000` / `5000` / `All`), **Export CSV...**, **Copy JSON**, **`EXPLAIN QUERY PLAN`** tree viewer, and clickable session execution history.
- **Database Switching & Maintenance:** Switch live between any `.db` / `.sqlite` file on disk (**Open DB...** / **New DB...**) or an isolated **`:memory:`** scratchpad, and run **`PRAGMA integrity_check`**, **`VACUUM`**, or online **`BackupToFile`**.

### Launching SQLite Studio from Ribbon or VBA

In `AccessSQLite.accdb`, open the **`SQLite Studio`** Ribbon tab (or in `AccessSQLiteDemo.accdb`, open the **`AccessSQLite`** Ribbon tab) and click **SQLite Studio**. You can also launch Studio windows programmatically from any referencing front-end `.accdb`:

```vba
' Open Studio connected to a specific SQLite file (multiple Studio windows can be open at once):
AccessSQLite.OpenSQLiteStudio CurrentProject.Path & "\AppData.db"

' Prompt user with native FileDialog to pick a .db / .sqlite / .sqlite3 file:
AccessSQLite.OpenSQLiteStudioPicker

' Open an isolated in-memory (:memory:) SQLite scratchpad:
AccessSQLite.OpenSQLiteStudioMemory

' Close all open Studio windows:
AccessSQLite.CloseAllSQLiteStudios
```

---

## 15. Data Type Marshaling & Edge-Case Guarantees

| Scenario | Legacy VBA Wrapper Behavior | `AccessSQLite` Guarantee |
|---|---|---|
| **Empty String `""` vs `Null`** | `StrPtr("")` can be `0`, causing `sqlite3_bind_text16` to bind SQL `NULL` and fail `NOT NULL` constraints | Binds a valid 2-byte zero pointer with `0` byte length for `""` (`''` in SQL) and calls `sqlite3_bind_null` only for `Null` / `Empty` / `Nothing`. |
| **64-bit Integers (`> 2,147,483,647`)** | Overflowed 32-bit `Long` or truncated to 16-bit `Integer` | Uses `LongLong` on 64-bit VBA and scaled `Currency` (`* 10000@` / `/ 10000@`) on 32-bit VBA, preserving exact integers up to `±922,337,203,685,477`. |
| **`0`-Byte BLOBs (`X''`)** | Crashed with `Subscript out of range` on `ReDim buf(length - 1)` | Binds via `sqlite3_bind_zeroblob(..., 0)` and returns an unallocated 0-length `Byte()` (`GetByteArrayLength(b) = 0`) cleanly. |
| **Unicode Paths & Text** | ANSI `String` Declares corrupted non-ACP characters (`äöü`, CJK, symbols) | Uses UTF-16 `StrPtr` + `sqlite3_column_bytes16` + `RtlMoveMemory` end to end (zero off-by-one truncation). |
| **DLL Unload & `LastDllError`** | `Debug.Print Err.LastDllError` wiped Win32 error to `0` (`Err.Raise 0` -> Error 5); `FreeLibrary` left stale handles | Captures `Err.LastDllError` immediately into a local `Long`, formats via `FormatMessageW`, frees `SQLite3_StdCall.dll` before `sqlite3.dll`, and zeroes all module handles. |

---

## 16. Error Handling & Diagnostics

All library errors raise structured `SQLiteErrorNumber` codes (`Err.Source = "AccessSQLite.<Class>.<Method>"`), including SQLite primary and extended result codes (`[sqlite_rc=19, ext_rc=2067]`):

| Enum Member | Numeric Offset | Raised When |
|---|---|---|
| `sqErrInitFailed` | `vbObjectError + 2101` | Driver DLL fails to load or `OpenConnection` cannot open the database file |
| `sqErrConnectionClosed` | `vbObjectError + 2102` | An operation is invoked on a closed `clsSQLiteConnection` or finalized `clsSQLiteStatement` |
| `sqErrSqlFailed` | `vbObjectError + 2103` | SQL syntax error, busy/locked timeout, I/O error, or invalid transaction state |
| `sqErrConstraint` | `vbObjectError + 2104` | `UNIQUE`, `NOT NULL`, `FOREIGN KEY`, or `CHECK` / `STRICT` constraint violation |
| `sqErrValidation` | `vbObjectError + 2105` | `clsSQLiteTableMap` validation failure (unknown column, non-editable field, max length, bad type, missing required field) |
| `sqErrNotFound` | `vbObjectError + 2106` | `GetOrFail`, dirty `UpdateRecord`, or `ReflectTableMap` targets a non-existent record/table |
| `sqErrMigration` | `vbObjectError + 2107` | Duplicate migration version, altered migration checksum, or failed up/down SQL script |
| `sqErrInvalidArgument` | `vbObjectError + 2108` | Empty identifier, invalid operator, missing named parameter, or `Nothing` object argument |
| `sqErrBridge` | `vbObjectError + 2109` | Access DAO <-> SQLite export/import error |

---

## 17. Developer Internals, Testing & Contribution Notes

### Running the Automated Test Suite

Open `AccessSQLite.accdb` and run in the VBA Immediate Window (**Ctrl+G**):

```vba
? RunAllSQLiteTests()
```

Or run via Access SQL:

```sql
SELECT RunAllSQLiteTests() AS TestSuiteResult FROM tblDual;
```

Expected output:

```text
PASS | Assertions=97 | Failed=0 | Elapsed=... ms
```

`modSQLiteTests` verifies:
1. Both `clsSQLiteWinDriver` (`winsqlite3.dll`) and `clsSQLiteBundledDriver` (`lib\sqlite3.dll` + `lib\SQLite3_StdCall.dll`)
2. Automatic SMB `DELETE` journal mode and `PRAGMA foreign_keys = ON` on `./AccessSQLite_Test.db`
3. Empty string `""` vs `Null`, Unicode UTF-16 strings, `12 KB` Memo strings, 64-bit integers (`5,000,000,007`), 4-decimal `Currency`, `Boolean`, `DateTime`, `0`-byte and binary `Byte()` BLOBs, and `@literal` positional disambiguation
4. Outer transactions and nested `SAVEPOINT` partial rollbacks
5. `STRICT` tables, `SyncAdditiveColumns`, `WhereContains` wildcard escaping, `WhereIn`, `WhereBetween`, `Page`, and `PluckCol`
6. `UpsertRecord`, `sqErrValidation`, `clsDemoTask` entity CRUD, skipped no-op dirty updates, soft deletes + `OnlyTrashed` + `RestoreRecord`, 2-query eager `IncludeHasMany`, `clsSQLiteUnitOfWork` rollback/commit, and reserved-word sanitization in `modSQLiteEntityGen`
7. Multi-step `clsSQLiteMigrator` up/down rollbacks, `clsSQLiteAccessBridge` DAO <-> SQLite round-trips (including `Null` binding reset, `400`-char Memo fields, and `0`-row table creation), and online `BackupToFile`
8. `clsSQLiteAdoBridge` disconnected client-side `ADODB.Recordset` hydration (`QueryAdoRecordset`, `ToAdoRecordset`), virtual `[IsSelected]` checkboxes (`GetSelectedKeys` / `SetAllSelected`), two-way `SaveFromAdoRow` insert & dirty update, and `ImportAdoRecordsetToSQLite`
9. Headless `clsSQLiteStudioHost` JSON protocol verification (`studio.init`, `studio.saveQuery`, `studio.exec`, `studio.explain`, `studio.deleteQuery`, and `PRAGMA integrity_check`)

### Coding & Naming Conventions

When extending `AccessSQLite`:
- Keep `Option Compare Database` and `Option Explicit` at the top of every module and class.
- Avoid Access, VBA, DAO, and ACE/Jet reserved words (`Date`, `Time`, `Name`, `Value`, `Text`, `Count`, `Index`, `Select`, `Order`, `Group`, `Table`, `Field`, `Section`, `Form`, `Report`) as variable, property, parameter, or procedure names.
- Keep `Attribute VB_Exposed = True` on all public classes and add factory functions to `modSQLiteFactory` for any new creatable public class.
- Preserve the VBA project name `AccessSQLite` (`EnsureLibraryProjectName "AccessSQLite"`) before compiling and distributing updates.

---

## 18. Credits, Acknowledgments & Third-Party Licenses

`AccessSQLite` builds upon foundational ideas, C/VBA interop techniques, and components from the SQLite and Office/VBA open-source community:

- **[SQLite](https://www.sqlite.org/) — D. Richard Hipp & the SQLite Consortium** *(Public Domain)*
  - The embedded SQL database engine powering both Windows' built-in `winsqlite3.dll` and the bundled `lib/sqlite3.dll` (`x86`) / `lib/x64/sqlite3.dll` (`x64`) binaries.
- **[`govert/SQLiteForExcel`](https://github.com/govert/SQLiteForExcel) — Govert van Drimmelen** *(MIT License)*
  - Author of the `SQLite3_StdCall.dll` 32-bit `__stdcall` C wrapper (`lib/SQLite3_StdCall.dll`) and the original VBA UTF-16 `StrPtr`, Win32 `MultiByteToWideChar` / `WideCharToMultiByte`, Julian-day date conversion, and online backup (`sqlite3_backup_*`) patterns adapted in `clsSQLiteBundledDriver` and `modSQLiteTypes`.
- **[`ReneNyffenegger/winsqlite3.dll-4-VBA`](https://github.com/ReneNyffenegger/winsqlite3.dll-4-VBA) — René Nyffenegger**
  - Demonstrated calling Windows 10/11's system `winsqlite3.dll` directly from VBA without external wrapper DLLs, inspiring `clsSQLiteWinDriver`'s zero-deployment 32-bit/64-bit `PtrSafe` driver architecture.
- ***SQLite for Access (`Beispiel.accdb`)* — Thomas Gewinnus** *(linked via [`govert/SQLiteForExcel`](https://github.com/govert/SQLiteForExcel))*
  - Ported the `SQLiteForExcel` interface to Microsoft Access VBA with an initial DAO-style `SQLiteDatabase` class and Access <-> SQLite import/export routines, inspiring `clsSQLiteAccessBridge`.
- **[`b-gonzalez/SQLiteForExcel`](https://github.com/b-gonzalez/SQLiteForExcel) — Brian Gonzalez** *(MIT License)*
  - Refactored `SQLiteForExcel` into class/interface modules (`ISqlite3` / `cSqlite3`) using unified `LongPtr` signatures across 32-bit and 64-bit VBA7, informing `ISQLiteDriver`.
- **[`cristianbuse/VBA-FastJSON`](https://github.com/cristianbuse/VBA-FastJSON) — Cristian Buse** *(MIT License)*
  - Fast, pure-VBA JSON parser and serializer (`LibJSON.bas`) used by `clsSQLiteWebBridge` and `clsSQLiteStudioHost` for JSON RPC marshaling between VBA and Monaco Studio.
- **[`microsoft/monaco-editor`](https://github.com/microsoft/monaco-editor) — Microsoft Corporation** *(MIT License)*
  - Browser-based code editor (`0.52.2` AMD build in `web/vendor/monaco/`) powering `AccessSQLite Studio` inside Access's Edge Browser Control (`https://msaccess/`).

### Third-Party License Notice (`SQLiteForExcel` / `SQLite3_StdCall.dll`)

```text
The MIT License (MIT)

Copyright (c) Govert van Drimmelen

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

