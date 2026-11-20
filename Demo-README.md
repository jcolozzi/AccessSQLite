# AccessSQLiteDemo — How the Continuous & Single-Page Forms Work Together

[`./AccessSQLiteDemo.accdb`](./AccessSQLiteDemo.accdb) is a reference Microsoft Access front-end application that consumes [`./AccessSQLite.accdb`](./AccessSQLite.accdb) as an external **VBA Class Library** (`Tools -> References`) and connects live SQLite data in [`./AccessSQLiteDemo.db`](./AccessSQLiteDemo.db) to both a **Continuous Form** (`frmTaskList`) and a **Single-Page Form** (`frmTaskDetail`).

For full API reference documentation on the underlying class library and ORM, see [`./README.md`](./README.md).

---

## Table of Contents

- [AccessSQLiteDemo — How the Continuous \& Single-Page Forms Work Together](#accesssqlitedemo--how-the-continuous--single-page-forms-work-together)
  - [Table of Contents](#table-of-contents)
  - [1. Files Involved \& Their Roles](#1-files-involved--their-roles)
  - [2. End-to-End Architecture](#2-end-to-end-architecture)
  - [3. How the VBA Reference Connects the Two `.accdb` Files](#3-how-the-vba-reference-connects-the-two-accdb-files)
  - [4. Startup \& Schema Bootstrap Flow (`modDemoDataService`)](#4-startup--schema-bootstrap-flow-moddemodataservice)
  - [5. Pattern 1: Continuous Form (`frmTaskList` + `tblTaskBuffer`)](#5-pattern-1-continuous-form-frmtasklist--tbltaskbuffer)
    - [Why a Local Buffer Table Is Used for Continuous Forms](#why-a-local-buffer-table-is-used-for-continuous-forms)
    - [Read Path (`RefreshFromSQLite` -\> `SyncBufferFromSQLite`)](#read-path-refreshfromsqlite---syncbufferfromsqlite)
    - [Write Path (Two-Way Inline Editing)](#write-path-two-way-inline-editing)
  - [6. Pattern 2: Single-Page Form (`frmTaskDetail` + `clsTaskItem`)](#6-pattern-2-single-page-form-frmtaskdetail--clstaskitem)
    - [Direct Unbound Entity Architecture](#direct-unbound-entity-architecture)
  - [7. How Both Forms Stay Synchronized](#7-how-both-forms-stay-synchronized)
  - [8. Interactive Walkthrough — Things to Try](#8-interactive-walkthrough--things-to-try)
  - [9. Object-by-Object Inventory in `AccessSQLiteDemo.accdb`](#9-object-by-object-inventory-in-accesssqlitedemoaccdb)
  - [10. Key Access + SQLite Gotchas Solved in This Demo](#10-key-access--sqlite-gotchas-solved-in-this-demo)
  - [11. Automated End-to-End Verification (`modDemoTests`)](#11-automated-end-to-end-verification-moddemotests)

---

## 1. Files Involved & Their Roles

| File | Role |
|---|---|
| [`./AccessSQLiteDemo.accdb`](./AccessSQLiteDemo.accdb) | **Front-end UI application** (VBA Project: `AccessSQLiteDemo`). Contains `frmTaskList`, `frmTaskDetail`, `clsTaskItem`, `modDemoDataService`, `modDemoTests`, and the local UI staging table `tblTaskBuffer`. |
| [`./AccessSQLite.accdb`](./AccessSQLite.accdb) | **Reusable VBA Class Library & ORM** (VBA Project: `AccessSQLite`). Contains all SQLite C-API drivers, `clsSQLiteConnection`, `clsSQLiteQueryBuilder`, `clsSQLiteTableMap`, `clsSQLiteRepository`, `ISQLiteEntity`, `clsSQLiteMigrator`, and `modSQLiteFactory`. |
| [`./AccessSQLiteDemo.db`](./AccessSQLiteDemo.db) | **Single source of truth for application data.** Standard SQLite 3 database holding the `Tasks` table and `_schema_migrations` history. |
| `winsqlite3.dll` *(Windows System)* / [`./lib/sqlite3.dll`](./lib/sqlite3.dll) | **Native SQLite C engine.** By default, `AccessSQLite` uses Windows' built-in `winsqlite3.dll` (zero external DLL deployment) and falls back automatically to the bundled `lib\sqlite3.dll` + `lib\SQLite3_StdCall.dll` (x86) or `lib\x64\sqlite3.dll` (x64). |

---

## 2. End-to-End Architecture

Microsoft Access has two distinct form architectures (`Continuous Forms` vs. `Single Form`), and each pairs best with a specific SQLite data-binding strategy. `AccessSQLiteDemo.accdb` demonstrates **both** patterns working side by side against the same `Tasks` table in `AccessSQLiteDemo.db`:

```text
+---------------------------------------------------------------------------------------+
|                            AccessSQLiteDemo.accdb (Front-End)                         |
|                                                                                       |
|  +---------------------------------------+   +-------------------------------------+  |
|  | frmTaskList (Continuous Form)         |   | frmTaskDetail (Single-Page Form)    |  |
|  |  - Bound to local tblTaskBuffer       |   |  - 100% Unbound controls            |  |
|  |  - Filter bar -> SyncBufferFromSQLite |   |  - Hydrates/saves clsTaskItem       |  |
|  |  - Form_AfterUpdate -> PushBufferRow  |   |    (Implements ISQLiteEntity)       |  |
|  +-------------------+-------------------+   +------------------+------------------+  |
|                      |                                          |                     |
|                      |    NotifyTaskListChanged(lngTaskId)      |                     |
|                      +<-----------------------------------------+                     |
|                      |                                          |                     |
|                      v                                          v                     |
|  +---------------------------------------------------------------------------------+  |
|  | modDemoDataService (Shared Connection, Schema Map, Migrations & Sync Helpers)   |  |
|  +----------------------------------------+----------------------------------------+  |
+-------------------------------------------|-------------------------------------------+
                                            | Early-bound Tools -> References
                                            v
+---------------------------------------------------------------------------------------+
|                        AccessSQLite.accdb (VBA Class Library & ORM)                   |
|   modSQLiteFactory -> clsSQLiteConnection -> clsSQLiteRepository / QueryBuilder       |
|                    -> clsSQLiteTableMap   -> clsSQLiteMigrator                        |
|                    -> ISQLiteDriver (clsSQLiteWinDriver / clsSQLiteBundledDriver)     |
+-------------------------------------------|-------------------------------------------+
                                            | Native UTF-16 SQLite3 C API
                                            v
+---------------------------------------------------------------------------------------+
|                           AccessSQLiteDemo.db (SQLite 3 File)                         |
|   Tables: Tasks (10 seeded rows + timestamps + soft-delete), _schema_migrations       |
+---------------------------------------------------------------------------------------+
```

---

## 3. How the VBA Reference Connects the Two `.accdb` Files

1. **Distinct VBA Project Names:**
   - `AccessSQLite.accdb` has its VBA project named **`AccessSQLite`**.
   - `AccessSQLiteDemo.accdb` has its VBA project named **`AccessSQLiteDemo`**.
   - Keeping unique project names avoids Access's *"Name conflicts with existing module, project, or object library"* error when linking `.accdb` files under **Tools -> References**.
2. **Exposed Public Classes (`Attribute VB_Exposed = True`):**
   - Every public class and interface in `AccessSQLite.accdb` (`clsSQLiteConnection`, `clsSQLiteRepository`, `clsSQLiteQueryBuilder`, `clsSQLiteTableMap`, `ISQLiteEntity`, etc.) is marked `PublicNotCreatable`.
   - Code inside `AccessSQLiteDemo.accdb` gets full compile-time type checking and IntelliSense (`Dim objQb As AccessSQLite.clsSQLiteQueryBuilder`).
3. **Factory Instantiation (`modSQLiteFactory`):**
   - Because VBA does not allow `New ExternalProject.clsName` across `.accdb` project boundaries, `AccessSQLiteDemo` opens connections and helper dictionaries through public factory functions (`AccessSQLite.OpenSQLite(...)`, `AccessSQLite.NewTextDict()`), and implements `AccessSQLite.ISQLiteEntity` directly on `clsTaskItem`.

---

## 4. Startup & Schema Bootstrap Flow (`modDemoDataService`)

`AccessSQLiteDemo.accdb` sets `StartupForm = "frmTaskList"`. On first data access, `modDemoDataService.GetDemoConnection()` runs a lazy one-time bootstrap (`ConfigureSchemaAndSeed`):

1. **Opens `AccessSQLiteDemo.db` with `sqProfileAuto`:**
   - Automatically detects whether `CurrentProject.Path` is on a UNC/mapped network share (`K:\...` -> applies SMB-safe `PRAGMA journal_mode = DELETE; PRAGMA synchronous = FULL; PRAGMA foreign_keys = ON;`) or a local SSD (`C:\...` -> applies `PRAGMA journal_mode = WAL;`).
2. **Applies Schema Migration v1 (`clsSQLiteMigrator`):**
   - Creates `_schema_migrations` (if absent), verifies the FNV-1a SQL checksum, and creates the `Tasks` table (`TaskID`, `TaskTitle`, `CategoryName`, `PriorityLevel`, `EstimateCost`, `DueDate`, `IsCompleted`, `TaskNotes`, `CreatedAt`, `UpdatedAt`, `DeletedAt`) plus indexes on `CategoryName` and `IsCompleted`.
3. **Registers the ORM `clsSQLiteTableMap`:**
   - Configures required-field rules (`TaskTitle` required, max 150 chars), defaults (`CategoryName = 'General'`, `PriorityLevel = 'Medium'`, `EstimateCost = 0`), automatic `CreatedAt` / `UpdatedAt` ISO timestamps, and soft-delete tracking on `DeletedAt`.
4. **Seeds Initial Demo Data:**
   - If `Tasks` has `0` rows, `SeedInitialTasks` batch-inserts 10 sample tasks inside a single SQLite transaction via `objRepo.InsertMany`.

---

## 5. Pattern 1: Continuous Form (`frmTaskList` + `tblTaskBuffer`)

### Why a Local Buffer Table Is Used for Continuous Forms
Native Access Continuous Forms (`DefaultView = 1`) require an Access/DAO `RecordSource` to paint repeating detail rows, scrollbars, and per-row bound editors without third-party ODBC drivers or ActiveX grids.

In `AccessSQLiteDemo`, **`tblTaskBuffer` is strictly a disposable UI viewport** — **`AccessSQLiteDemo.db` remains the single source of truth**:

### Read Path (`RefreshFromSQLite` -> `SyncBufferFromSQLite`)
1. Reads the user's current header controls (`txtSearch`, `cboFilterCategory`, `cboFilterStatus`).
2. Builds a 100% parameterized SQLite query via `TasksRepo().NewQuery()`:
   - **View scope:**
     - `"Active"` -> default ORM scope (`WHERE "DeletedAt" IS NULL`)
     - `"Open Only"` -> `.WhereEq "IsCompleted", False`
     - `"Completed Only"` -> `.WhereEq "IsCompleted", True`
     - `"Trash (Soft-Deleted)"` -> `.OnlyTrashed` (`WHERE "DeletedAt" IS NOT NULL`)
   - **Category filter:** `.WhereEq "CategoryName", strCleanCat`
   - **Substring search:** `.WhereContains "TaskTitle", strCleanSearch` (escapes `%`, `_`, and `\` automatically)
3. Executes `.OrderByCol("IsCompleted", False).OrderByCol("TaskID", False).ToRowsRaw()`, clears `tblTaskBuffer`, appends the matching rows, calls `Me.Requery`, updates the live SQLite KPI banner (`GetSqliteSummaryBanner()`), and restores the selection bookmark if a `lngSelectTaskId` was supplied.

### Write Path (Two-Way Inline Editing)
- When you check/uncheck **`Done`** (`chkRowDone_AfterUpdate`), the form commits the current row immediately (`CommitPendingEditQuiet`), firing **`Form_AfterUpdate`**.
- When you edit **`Task Title`**, **`Category`**, **`Priority`**, **`Est. Cost`**, or **`Due Date`** directly in any row of the Continuous Form, **`Form_AfterUpdate`** calls `PushBufferRowToSQLite`.
- `PushBufferRowToSQLite` passes the row's values to **`TasksRepo().UpdateRecord(lngTaskId, dictChanges, blnDirtyOnly:=True)`**, which:
  1. Validates & coerces every value against `clsSQLiteTableMap` (raising `sqErrValidation` if `TaskTitle` was cleared),
  2. Diffs against the stored SQLite row so unchanged columns are ignored,
  3. Stamps a fresh `UpdatedAt` timestamp in `AccessSQLiteDemo.db`, and
  4. Refreshes the header KPI summary (`Open` count and `$ Open Est.` update immediately).

---

## 6. Pattern 2: Single-Page Form (`frmTaskDetail` + `clsTaskItem`)

Unlike Continuous Forms, an Access **Single Form** (`DefaultView = 0`) displays one record at a time and **needs no local Access staging table at all**.

### Direct Unbound Entity Architecture
1. **Unbound Controls (`ControlSource = ""`):**
   - Every input control on `frmTaskDetail` (`txtTaskID`, `txtTaskTitle`, `cboCategoryName`, `cboPriorityLevel`, `txtEstimateCost`, `txtDueDate`, `chkIsCompleted`, `txtTaskNotes`, `txtAuditStamp`) is unbound.
2. **Strongly-Typed Entity (`clsTaskItem`):**
   - Implements `AccessSQLite.ISQLiteEntity` (`EntityTableName`, `CreateBlank`, `ReadFromRow`, `WriteToRow`), mapping SQLite columns cleanly to typed VBA properties (`Long`, `String`, `Currency`, `Variant` nullable `Date`, `Boolean`).
3. **Loading & Navigating Directly Against SQLite:**
   - `LoadTaskFromSQLite(lngTaskId)` calls:
     ```vba
     Set objLoaded = TasksRepo().FindEntity(lngTaskId, New clsTaskItem, True)
     ```
     and populates the form controls plus the read-only `SQLite Timestamps` audit bar (`CreatedAt` / `UpdatedAt` or `TRASHED` timestamp).
   - Clicking **`< Prev`** or **`Next >`** calls `FindTargetSqliteTaskId("PREV" / "NEXT", lngCur)`, running `SELECT MAX(TaskID) ...` / `SELECT MIN(TaskID) ...` directly on `AccessSQLiteDemo.db` without touching `tblTaskBuffer`.
4. **Saving & Validation (`SaveCurrentTaskToSQLite`):**
   - Copies control values onto `m_CurrentTask` (`clsTaskItem`) and calls:
     ```vba
     TasksRepo().SaveEntity m_CurrentTask, blnDirtyOnly:=True
     ```
   - For a new task (`TaskID = 0`), `SaveEntity` inserts into SQLite and automatically hydrates the newly generated `TaskID` and `CreatedAt` timestamp back onto `m_CurrentTask`.
   - For an existing task, `SaveEntity` diffs against the stored row in SQLite — skipping the `UPDATE` if nothing changed, or updating only modified columns + `UpdatedAt`.
   - If validation fails (for example, leaving `Task Title *` blank), `clsSQLiteTableMap` raises `sqErrValidation`, which `frmTaskDetail` catches and displays in a clean validation prompt.

---

## 7. How Both Forms Stay Synchronized

Both forms can stay open at the same time in Access's tabbed document interface and coordinate seamlessly:

- **Continuous Form -> Single-Page Form:**
  - Clicking **`Edit...`** on a row (or **`Open Detail Form`** in the footer) checks `CurrentProject.AllForms("frmTaskDetail").IsLoaded`. If `frmTaskDetail` is already open, it calls `Forms("frmTaskDetail").LoadTaskFromSQLite lngId` and focuses it; otherwise it opens `frmTaskDetail` passing `OpenArgs:=CStr(lngId)`.
  - Clicking **`+ New Task`** opens `frmTaskDetail` with `OpenArgs:="NEW"` (or calls `BeginNewTask` if already open).
- **Single-Page Form -> Continuous Form:**
  - Whenever `frmTaskDetail` saves, soft-deletes, or restores a task, it calls `modDemoDataService.NotifyTaskListChanged(lngTaskId)`.
  - If `frmTaskList` is open, `NotifyTaskListChanged` calls `Forms("frmTaskList").RefreshFromSQLite lngTaskId` — re-querying SQLite, updating the header KPI banner, and moving the Continuous Form's active row highlight directly onto the saved record.

---

## 8. Interactive Walkthrough — Things to Try

1. **Launch the Demo:** Open [`./AccessSQLiteDemo.accdb`](./AccessSQLiteDemo.accdb). `frmTaskList` opens automatically showing 10 seeded tasks from `AccessSQLiteDemo.db` (7 Open at the top, 3 Completed below).
2. **Inline Toggle & Cost Rollup:** Click the **`Done`** checkbox on any open task or change its **`Est. Cost`**. Watch the top slate banner immediately recalculate the `Open` count and `$ Open Est.` directly from SQLite.
3. **Filter & Wildcard-Safe Search:** Select `Category: Engineering` or type a keyword in `Search:` and press **Enter**. Click **Reset** to clear all filters.
4. **Add a New Task via the ORM Single-Page Form:** Click **`+ New Task`**, leave **`Task Title *`** blank, and click **`Save to SQLite`** to see `clsSQLiteTableMap` validation in action. Then enter a title, pick a category/priority/cost/date, and click **`Save to SQLite`** — watch `SQLite Task ID` and `SQLite Timestamps` populate and switch back to the `frmTaskList` tab to see the new row already highlighted.
5. **Soft-Delete Trash & Restore:** Select any task and click **`Soft Delete`**. Switch the **`View:`** dropdown on `frmTaskList` to **`Trash (Soft-Deleted)`** — notice the footer button automatically changes to **`Restore Task`**. Click **`Restore Task`** to undelete it (`DeletedAt -> NULL`).

---

## 9. Object-by-Object Inventory in `AccessSQLiteDemo.accdb`

| Object Type | Name | Description |
|---|---|---|
| **Form** | `frmTaskList` | Continuous Form dashboard (`DefaultView = 1`) bound to `tblTaskBuffer` with live SQLite filter bar, KPI summary banner, and two-way inline editing. |
| **Form** | `frmTaskDetail` | Unbound Single-Page Form (`DefaultView = 0`) bound in VBA to `clsTaskItem` & `clsSQLiteRepository` with direct SQLite `< Prev` / `Next >` navigation. |
| **Class Module** | `clsTaskItem` | Domain entity implementing `AccessSQLite.ISQLiteEntity` for the `Tasks` table. |
| **Standard Module** | `modDemoDataService` | Manages the `clsSQLiteConnection`, runs `clsSQLiteMigrator` v1, defines `clsSQLiteTableMap`, seeds initial rows, and syncs `frmTaskList` <-> `frmTaskDetail`. |
| **Standard Module** | `modDemoTests` | Contains `RunDemoFormsVerification()` for automated end-to-end testing of both forms and their SQLite round-trips. |
| **Table** | `tblTaskBuffer` | Local UI viewport table populated from `AccessSQLiteDemo.db` for `frmTaskList`. |
| **Table** | `tblDual` | Single-row helper table enabling SQL-based test/diagnostic execution (`SELECT RunDemoFormsVerification() FROM tblDual`). |

---

## 10. Key Access + SQLite Gotchas Solved in This Demo

If you use `AccessSQLiteDemo.accdb` as a template for your own Access + SQLite apps, note four subtle Access behaviors already handled in the code:

1. **Access Boolean Sort Polarity (`True = -1`, `False = 0`):**
   - In SQLite, `IsCompleted = 0` (`False`) sorts before `1` (`True`) with `ASC`. In Access/ACE `Yes/No` columns (`tblTaskBuffer.IsCompleted`), `True` is `-1` and `False` is `0`, so `frmTaskList.RecordSource` uses `ORDER BY [IsCompleted] DESC, [TaskID] ASC` to keep open (`0`) tasks at the top.
2. **Guarding `Me.Dirty` on Empty Filter Results (`Error 2455`):**
   - On a Continuous Form with `AllowAdditions = False`, reading `Me.Dirty` when a search returns `0` rows raises Access Run-time Error `2455`. `frmTaskList.CommitPendingEditQuiet()` checks `Not (Me.RecordsetClone.BOF And Me.RecordsetClone.EOF)` before touching `Me.Dirty`.
3. **Avoiding `Workspaces(0)` Cache Delay on Local Staging Tables:**
   - Wrapping `DELETE FROM [tblTaskBuffer]` in `DBEngine.Workspaces(0).BeginTrans` on a separate `CurrentDb` handle can cause `Me.Requery` to read stale cached pages before ACE's page cache expires. Executing directly against `CurrentDb` without an outer DAO workspace transaction keeps `Me.Requery` instantaneous and consistent.
4. **Unbound Form Controls Must Have Empty `ControlSource` (`Error 2448`):**
   - Controls on unbound ORM forms like `frmTaskDetail` must leave `ControlSource = ""` so Access does not treat them as broken bound expressions (`#Name?`) when VBA assigns `.Value`.

---

## 11. Automated End-to-End Verification (`modDemoTests`)

You can run the automated UI + SQLite integration check at any time from the VBA Immediate Window (**Ctrl+G**) inside `AccessSQLiteDemo.accdb`:

```vba
? RunDemoFormsVerification()
```

`RunDemoFormsVerification()` opens `frmTaskList`, verifies row counts against SQLite, tests category filtering and 0-row search/reset handling, opens `frmTaskDetail` in `"NEW"` mode, inserts a task via `clsTaskItem` + `SaveEntity`, verifies `frmTaskList` auto-refreshes, updates the row inline, verifies `frmTaskDetail` reads back the updated values from `AccessSQLiteDemo.db`, cleans up its temporary row, and returns:

```text
PASS | Initial=10 | FilteredEng=3 | RoundTripID=...
```
