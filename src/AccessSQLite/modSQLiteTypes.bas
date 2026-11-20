Attribute VB_Name = "modSQLiteTypes"
Option Compare Database
Option Explicit

' =============================================================================
' modSQLiteTypes
' Core enums, error numbers, Win32 helpers, UTF-8/UTF-16 marshaling,
' ISO-8601/Julian/Unix date conversions, and high-resolution timing.
' =============================================================================

Public Const SQLITE_JULIAN_EPOCH_OFFSET As Double = 2415018.5
Public Const SQLITE_UNIX_EPOCH_DATE As Double = 25569# ' 1970-01-01 in OLE Automation Date

Public Enum SQLiteDriverKind
    sqDriverAuto = 0
    sqDriverWinSqlite = 1
    sqDriverBundled = 2
End Enum

Public Enum SQLiteResultCode
    sqOk = 0
    sqError = 1
    sqInternal = 2
    sqPerm = 3
    sqAbort = 4
    sqBusy = 5
    sqLocked = 6
    sqNoMem = 7
    sqReadOnly = 8
    sqInterrupt = 9
    sqIoErr = 10
    sqCorrupt = 11
    sqNotFound = 12
    sqFull = 13
    sqCantOpen = 14
    sqProtocol = 15
    sqEmpty = 16
    sqSchema = 17
    sqTooBig = 18
    sqConstraint = 19
    sqMismatch = 20
    sqMisuse = 21
    sqNoLfs = 22
    sqAuth = 23
    sqFormat = 24
    sqRange = 25
    sqNotADb = 26
    sqRow = 100
    sqDone = 101
End Enum

Public Enum SQLiteOpenFlags
    sqOpenReadOnly = &H1
    sqOpenReadWrite = &H2
    sqOpenCreate = &H4
    sqOpenUri = &H40
    sqOpenMemory = &H80
    sqOpenNoMutex = &H8000&
    sqOpenFullMutex = &H10000
    sqOpenSharedCache = &H20000
    sqOpenPrivateCache = &H40000
End Enum

Public Enum SQLitePrepareFlags
    sqPrepareNormal = 0
    sqPreparePersistent = &H1
End Enum

Public Enum SQLiteNativeAffinity
    sqAffinityInteger = 1
    sqAffinityFloat = 2
    sqAffinityText = 3
    sqAffinityBlob = 4
    sqAffinityNull = 5
End Enum

Public Enum SQLiteColumnKind
    sqColText = 1
    sqColMemo = 2
    sqColLong = 3
    sqColInt64 = 4
    sqColDouble = 5
    sqColCurrency = 6
    sqColBoolean = 7
    sqColDate = 8
    sqColDateTime = 9
    sqColBlob = 10
    sqColJson = 11
    sqColGuid = 12
End Enum

Public Enum SQLiteDateStorage
    sqDateIso8601 = 1
    sqDateJulianReal = 2
    sqDateUnixSeconds = 3
End Enum

Public Enum SQLitePragmaProfile
    sqProfileAuto = 0
    sqProfileLocalWal = 1
    sqProfileNetworkSafe = 2
    sqProfileReadOnly = 3
    sqProfileInMemory = 4
End Enum

Public Enum SQLiteErrorNumber
    sqErrInitFailed = vbObjectError + 2101
    sqErrConnectionClosed = vbObjectError + 2102
    sqErrSqlFailed = vbObjectError + 2103
    sqErrConstraint = vbObjectError + 2104
    sqErrValidation = vbObjectError + 2105
    sqErrNotFound = vbObjectError + 2106
    sqErrMigration = vbObjectError + 2107
    sqErrInvalidArgument = vbObjectError + 2108
    sqErrBridge = vbObjectError + 2109
End Enum

Private Const CP_UTF8 As Long = 65001
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
Private Const DRIVE_REMOTE As Long = 4

Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal lngCodePage As Long, _
    ByVal lngFlags As Long, _
    ByVal ptrMultiByteStr As LongPtr, _
    ByVal lngMultiByteLen As Long, _
    ByVal ptrWideCharStr As LongPtr, _
    ByVal lngWideCharCount As Long) As Long

Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal lngCodePage As Long, _
    ByVal lngFlags As Long, _
    ByVal ptrWideCharStr As LongPtr, _
    ByVal lngWideCharCount As Long, _
    ByVal ptrMultiByteStr As LongPtr, _
    ByVal lngMultiByteLen As Long, _
    ByVal ptrDefaultChar As LongPtr, _
    ByVal ptrUsedDefaultChar As LongPtr) As Long

Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" ( _
    ByVal ptrDest As LongPtr, _
    ByVal ptrSource As LongPtr, _
    ByVal lngByteLen As Long)

Private Declare PtrSafe Function lstrlenW Lib "kernel32" ( _
    ByVal ptrWideString As LongPtr) As Long

Private Declare PtrSafe Function FormatMessageW Lib "kernel32" ( _
    ByVal lngFlags As Long, _
    ByVal ptrSource As LongPtr, _
    ByVal lngMessageId As Long, _
    ByVal lngLanguageId As Long, _
    ByVal ptrBuffer As LongPtr, _
    ByVal lngBufferSize As Long, _
    ByVal ptrArguments As LongPtr) As Long

Private Declare PtrSafe Function GetDriveTypeW Lib "kernel32" ( _
    ByVal ptrRootPathName As LongPtr) As Long

Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" ( _
    ByRef curCounter As Currency) As Long

Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" ( _
    ByRef curFrequency As Currency) As Long

Private m_PerfFreq As Currency

' Creates a case-insensitive Scripting.Dictionary (CompareMode = TextCompare).
Public Function NewTextDict() As Scripting.Dictionary
    Dim dictResult As Scripting.Dictionary
    Set dictResult = New Scripting.Dictionary
    dictResult.CompareMode = TextCompare
    Set NewTextDict = dictResult
End Function

' Captures a high-resolution QPC tick counter.
Public Function StartPerfTimer() As Currency
    Dim curNow As Currency
    QueryPerformanceCounter curNow
    StartPerfTimer = curNow
End Function

' Computes elapsed milliseconds since StartPerfTimer().
Public Function ElapsedPerfMs(ByVal curStarted As Currency) As Double
    Dim curNow As Currency
    If m_PerfFreq = 0@ Then
        QueryPerformanceFrequency m_PerfFreq
    End If
    QueryPerformanceCounter curNow
    If m_PerfFreq > 0@ Then
        ElapsedPerfMs = ((curNow - curStarted) / m_PerfFreq) * 1000#
    End If
End Function

' Formats a Win32 GetLastError / Err.LastDllError integer into human-readable system text.
Public Function FormatWin32Error(ByVal lngErrCode As Long) As String
    Dim strBuf As String
    Dim lngChars As Long
    If lngErrCode = 0 Then
        FormatWin32Error = "Success (0)"
        Exit Function
    End If
    strBuf = String$(512, vbNullChar)
    lngChars = FormatMessageW( _
        FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
        0, lngErrCode, 0, StrPtr(strBuf), Len(strBuf), 0)
    If lngChars > 0 Then
        strBuf = Left$(strBuf, lngChars)
        strBuf = Replace(Replace(strBuf, vbCr, vbNullString), vbLf, vbNullString)
        FormatWin32Error = Trim$(strBuf) & " (Win32=" & lngErrCode & ")"
    Else
        FormatWin32Error = "Win32 Error " & lngErrCode
    End If
End Function

' Detects whether a database path resides on a UNC share or mapped SMB network drive.
Public Function IsNetworkOrUncPath(ByVal strFilePath As String) As Boolean
    Dim strTrimmed As String
    Dim strRoot As String
    strTrimmed = Trim$(strFilePath)
    If Len(strTrimmed) = 0 Then Exit Function
    If StrComp(strTrimmed, ":memory:", vbTextCompare) = 0 Then Exit Function
    If Left$(strTrimmed, 2) = "\\" Then
        IsNetworkOrUncPath = True
        Exit Function
    End If
    If Len(strTrimmed) >= 2 And Mid$(strTrimmed, 2, 1) = ":" Then
        strRoot = Left$(strTrimmed, 2) & "\"
        IsNetworkOrUncPath = (GetDriveTypeW(StrPtr(strRoot)) = DRIVE_REMOTE)
    End If
End Function

' Resolves the folder containing AccessSQLite.accdb itself (even when referenced from an external .accdb).
Public Function GetLibraryDirectory() As String
    On Error Resume Next
    GetLibraryDirectory = CodeProject.Path
    If Len(GetLibraryDirectory) = 0 Then
        GetLibraryDirectory = CurrentProject.Path
    End If
    On Error GoTo 0
End Function

' Resolves the default bundled DLL directory (<CodeProject.Path>\lib, with fallback to <CodeProject.Path>).
Public Function GetBundledDllDirectory() As String
    Dim strRoot As String
    Dim strLibSub As String
    strRoot = GetLibraryDirectory()
    If Right$(strRoot, 1) = "\" Then strRoot = Left$(strRoot, Len(strRoot) - 1)
    strLibSub = strRoot & "\lib"
    On Error Resume Next
    If (GetAttr(strLibSub) And vbDirectory) = vbDirectory Then
        If Err.Number = 0 Then
            GetBundledDllDirectory = strLibSub
            On Error GoTo 0
            Exit Function
        End If
    End If
    Err.Clear
    On Error GoTo 0
    GetBundledDllDirectory = strRoot
End Function

' Quotes an identifier (or alias.column / table.*) safely for SQLite SQL using double quotes.
Public Function QuoteSQLiteIdent(ByVal strIdent As String) As String
    Dim strClean As String
    Dim lngDotPos As Long
    Dim strLeftPart As String
    Dim strRightPart As String

    strClean = Trim$(strIdent)
    If Len(strClean) = 0 Then
        Err.Raise sqErrInvalidArgument, "modSQLiteTypes.QuoteSQLiteIdent", "SQLite identifier cannot be empty."
    End If
    If strClean = "*" Then
        QuoteSQLiteIdent = "*"
        Exit Function
    End If

    lngDotPos = InStr(1, strClean, ".", vbBinaryCompare)
    If lngDotPos > 0 Then
        strLeftPart = Left$(strClean, lngDotPos - 1)
        strRightPart = Mid$(strClean, lngDotPos + 1)
        If strRightPart = "*" Then
            QuoteSQLiteIdent = """" & Replace(strLeftPart, """", """""") & """.*"
        Else
            QuoteSQLiteIdent = """" & Replace(strLeftPart, """", """""") & """.""" & Replace(strRightPart, """", """""") & """"
        End If
    Else
        QuoteSQLiteIdent = """" & Replace(strClean, """", """""") & """"
    End If
End Function

' Quotes an identifier safely for Access ACE/Jet SQL using square brackets.
Public Function QuoteAccessIdent(ByVal strIdent As String) As String
    Dim strClean As String
    strClean = Trim$(strIdent)
    If Len(strClean) = 0 Or InStr(1, strClean, "]", vbBinaryCompare) > 0 Or InStr(1, strClean, "!", vbBinaryCompare) > 0 Then
        Err.Raise sqErrInvalidArgument, "modSQLiteTypes.QuoteAccessIdent", "Invalid Access identifier: " & strIdent
    End If
    QuoteAccessIdent = "[" & strClean & "]"
End Function

' Converts a null-terminated UTF-8 pointer from SQLite into a VBA UTF-16 String.
Public Function Utf8PtrToVbaString(ByVal ptrUtf8 As LongPtr) As String
    Dim lngWideChars As Long
    Dim lngWritten As Long
    If ptrUtf8 = 0 Then Exit Function
    lngWideChars = MultiByteToWideChar(CP_UTF8, 0, ptrUtf8, -1, 0, 0)
    If lngWideChars <= 1 Then Exit Function
    Utf8PtrToVbaString = String$(lngWideChars - 1, vbNullChar)
    lngWritten = MultiByteToWideChar(CP_UTF8, 0, ptrUtf8, -1, StrPtr(Utf8PtrToVbaString), lngWideChars)
    If lngWritten = 0 Then Utf8PtrToVbaString = vbNullString
End Function

' Converts a UTF-16LE pointer from SQLite into a VBA String without off-by-one truncation.
Public Function Utf16PtrToVbaString(ByVal ptrUtf16 As LongPtr, Optional ByVal lngByteCount As Long = -1) As String
    Dim lngCharCount As Long
    Dim lngBytesToCopy As Long
    If ptrUtf16 = 0 Then Exit Function
    If lngByteCount >= 0 Then
        If lngByteCount = 0 Then Exit Function
        lngCharCount = lngByteCount \ 2
        lngBytesToCopy = lngCharCount * 2
    Else
        lngCharCount = lstrlenW(ptrUtf16)
        If lngCharCount <= 0 Then Exit Function
        lngBytesToCopy = lngCharCount * 2
    End If
    Utf16PtrToVbaString = String$(lngCharCount, vbNullChar)
    RtlMoveMemory StrPtr(Utf16PtrToVbaString), ptrUtf16, lngBytesToCopy
End Function

' Converts a VBA String to a null-terminated UTF-8 Byte array (always 0-based with trailing 0 byte).
Public Function VbaStringToUtf8Bytes(ByVal strInput As String) As Byte()
    Dim abytOut() As Byte
    Dim lngByteSize As Long
    If Len(strInput) = 0 Then
        ReDim abytOut(0 To 0)
        abytOut(0) = 0
        VbaStringToUtf8Bytes = abytOut
        Exit Function
    End If
    lngByteSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strInput), -1, 0, 0, 0, 0)
    If lngByteSize <= 0 Then
        ReDim abytOut(0 To 0)
        abytOut(0) = 0
        VbaStringToUtf8Bytes = abytOut
        Exit Function
    End If
    ReDim abytOut(0 To lngByteSize - 1)
    WideCharToMultiByte CP_UTF8, 0, StrPtr(strInput), -1, VarPtr(abytOut(0)), lngByteSize, 0, 0
    VbaStringToUtf8Bytes = abytOut
End Function

' Copies raw unmanaged bytes from a SQLite BLOB pointer safely (handles 0-byte BLOBs).
Public Function PtrToByteArray(ByVal ptrData As LongPtr, ByVal lngByteCount As Long) As Byte()
    Dim abytEmpty() As Byte
    Dim abytResult() As Byte
    If lngByteCount <= 0 Or ptrData = 0 Then
        abytEmpty = StrConv(vbNullString, vbFromUnicode)
        PtrToByteArray = abytEmpty
        Exit Function
    End If
    ReDim abytResult(0 To lngByteCount - 1)
    RtlMoveMemory VarPtr(abytResult(0)), ptrData, lngByteCount
    PtrToByteArray = abytResult
End Function

' Returns True if a Byte() dynamic array is unallocated or has 0 length, and outputs element count.
Public Function GetByteArrayLength(ByRef abytData() As Byte) As Long
    On Error GoTo EmptyArr
    Dim lngLower As Long
    Dim lngUpper As Long
    lngLower = LBound(abytData)
    lngUpper = UBound(abytData)
    If lngUpper < lngLower Then
        GetByteArrayLength = 0
    Else
        GetByteArrayLength = (lngUpper - lngLower) + 1
    End If
    Exit Function
EmptyArr:
    GetByteArrayLength = 0
End Function

' Date conversion helpers
Public Function DateToJulianDay(ByVal dtmValue As Date) As Double
    DateToJulianDay = CDbl(dtmValue) + SQLITE_JULIAN_EPOCH_OFFSET
End Function

Public Function JulianDayToDate(ByVal dblJulian As Double) As Date
    JulianDayToDate = CDate(dblJulian - SQLITE_JULIAN_EPOCH_OFFSET)
End Function

Public Function DateToUnixSeconds(ByVal dtmValue As Date) As Currency
    DateToUnixSeconds = CCur(Fix((CDbl(dtmValue) - SQLITE_UNIX_EPOCH_DATE) * 86400# + 0.5))
End Function

Public Function UnixSecondsToDate(ByVal curSeconds As Currency) As Date
    UnixSecondsToDate = CDate((CDbl(curSeconds) / 86400#) + SQLITE_UNIX_EPOCH_DATE)
End Function

Public Function FormatIsoDateValue(ByVal dtmValue As Date, Optional ByVal blnAlwaysIncludeTime As Boolean = False) As String
    If Not blnAlwaysIncludeTime And dtmValue = Int(dtmValue) Then
        FormatIsoDateValue = Format$(dtmValue, "yyyy-mm-dd")
    Else
        FormatIsoDateValue = Format$(dtmValue, "yyyy-mm-dd\Thh:nn:ss")
    End If
End Function

Public Function UtcNowIsoString() As String
    ' Safe local ISO timestamp generated in VBA (avoids winsqlite3 OMIT_LOCALTIME compile option).
    UtcNowIsoString = Format$(Now, "yyyy-mm-dd\Thh:nn:ss")
End Function

' Strict ISO-8601 date/datetime parser ("yyyy-mm-dd", "yyyy-mm-ddThh:nn:ss", "yyyy-mm-dd hh:nn:ss[.fff][Z]").
Public Function TryParseIsoDate(ByVal strInput As String, ByRef dtmParsed As Date) As Boolean
    On Error GoTo ParseFail
    Dim strClean As String
    Dim lngYear As Long
    Dim lngMonth As Long
    Dim lngDay As Long
    Dim lngHour As Long
    Dim lngMin As Long
    Dim lngSec As Long
    Dim strSep As String

    strClean = Trim$(strInput)
    If Len(strClean) < 10 Then Exit Function
    If Mid$(strClean, 5, 1) <> "-" Or Mid$(strClean, 8, 1) <> "-" Then Exit Function
    If Not (IsDigitsOnly(Left$(strClean, 4)) And IsDigitsOnly(Mid$(strClean, 6, 2)) And IsDigitsOnly(Mid$(strClean, 9, 2))) Then Exit Function

    lngYear = CLng(Left$(strClean, 4))
    lngMonth = CLng(Mid$(strClean, 6, 2))
    lngDay = CLng(Mid$(strClean, 9, 2))
    If lngMonth < 1 Or lngMonth > 12 Or lngDay < 1 Or lngDay > 31 Then Exit Function

    dtmParsed = DateSerial(lngYear, lngMonth, lngDay)
    If Year(dtmParsed) <> lngYear Or Month(dtmParsed) <> lngMonth Or Day(dtmParsed) <> lngDay Then Exit Function

    If Len(strClean) = 10 Then
        TryParseIsoDate = True
        Exit Function
    End If

    strSep = Mid$(strClean, 11, 1)
    If strSep <> "T" And strSep <> "t" And strSep <> " " Then Exit Function
    If Len(strClean) < 19 Then Exit Function
    If Mid$(strClean, 14, 1) <> ":" Or Mid$(strClean, 17, 1) <> ":" Then Exit Function
    If Not (IsDigitsOnly(Mid$(strClean, 12, 2)) And IsDigitsOnly(Mid$(strClean, 15, 2)) And IsDigitsOnly(Mid$(strClean, 18, 2))) Then Exit Function

    lngHour = CLng(Mid$(strClean, 12, 2))
    lngMin = CLng(Mid$(strClean, 15, 2))
    lngSec = CLng(Mid$(strClean, 18, 2))
    If lngHour > 23 Or lngMin > 59 Or lngSec > 59 Then Exit Function

    dtmParsed = dtmParsed + TimeSerial(lngHour, lngMin, lngSec)
    TryParseIsoDate = True
    Exit Function
ParseFail:
    TryParseIsoDate = False
End Function

Private Function IsDigitsOnly(ByVal strPart As String) As Boolean
    Dim lngIdx As Long
    Dim intAsc As Integer
    If Len(strPart) = 0 Then Exit Function
    For lngIdx = 1 To Len(strPart)
        intAsc = AscW(Mid$(strPart, lngIdx, 1))
        If intAsc < 48 Or intAsc > 57 Then Exit Function
    Next lngIdx
    IsDigitsOnly = True
End Function

' Deterministic 32-bit FNV-1a checksum in 8-char uppercase hex (used by clsSQLiteMigrator).
Public Function ComputeFnv1aHex(ByVal strText As String) As String
    Dim curHash As Currency
    Dim curPrime As Currency
    Dim curMod As Currency
    Dim lngIdx As Long
    Dim lngCode As Long

    curHash = 2166136261@
    curPrime = 16777619@
    curMod = 4294967296@

    For lngIdx = 1 To Len(strText)
        lngCode = AscW(Mid$(strText, lngIdx, 1)) And &HFFFF&
        curHash = XorUint32(curHash, CCur(lngCode And &HFF&))
        curHash = MulModUint32(curHash, curPrime, curMod)
        If lngCode > &HFF& Then
            curHash = XorUint32(curHash, CCur((lngCode \ 256) And &HFF&))
            curHash = MulModUint32(curHash, curPrime, curMod)
        End If
    Next lngIdx

    ComputeFnv1aHex = Uint32ToHex8(curHash)
End Function

Private Function XorUint32(ByVal curA As Currency, ByVal curByte As Currency) As Currency
    Dim curHigh As Currency
    Dim lngLow As Long
    curHigh = Int(curA / 256@) * 256@
    lngLow = CLng(curA - curHigh) Xor CLng(curByte)
    XorUint32 = curHigh + CCur(lngLow)
End Function

Private Function MulModUint32(ByVal curVal As Currency, ByVal curPrime As Currency, ByVal curMod As Currency) As Currency
    Dim curLo As Currency
    Dim curHi As Currency
    Dim curProd As Currency
    curLo = curVal - (Int(curVal / 65536@) * 65536@)
    curHi = Int(curVal / 65536@)
    curProd = (curLo * curPrime) + (((curHi * curPrime) - (Int((curHi * curPrime) / 65536@) * 65536@)) * 65536@)
    MulModUint32 = curProd - (Int(curProd / curMod) * curMod)
End Function

Private Function Uint32ToHex8(ByVal curVal As Currency) As String
    Dim lngHi As Long
    Dim lngLo As Long
    lngHi = CLng(Int(curVal / 65536@))
    lngLo = CLng(curVal - (CCur(lngHi) * 65536@))
    Uint32ToHex8 = Right$("0000" & Hex$(lngHi), 4) & Right$("0000" & Hex$(lngLo), 4)
End Function

' Raises a structured SQLite VBA runtime error.
Public Sub RaiseSQLiteError( _
    ByVal enmErrNum As SQLiteErrorNumber, _
    ByVal strSource As String, _
    ByVal strMessage As String, _
    Optional ByVal lngSqliteCode As Long = 0, _
    Optional ByVal lngExtendedCode As Long = 0)

    Dim strFullMsg As String
    strFullMsg = strMessage
    If lngSqliteCode <> 0 Or lngExtendedCode <> 0 Then
        strFullMsg = strFullMsg & " [sqlite_rc=" & lngSqliteCode
        If lngExtendedCode <> 0 And lngExtendedCode <> lngSqliteCode Then
            strFullMsg = strFullMsg & ", ext_rc=" & lngExtendedCode
        End If
        strFullMsg = strFullMsg & "]"
    End If
    Err.Raise enmErrNum, "AccessSQLite." & strSource, strFullMsg
End Sub

' Renames this library's own VBProject (matched by CodeProject.Name, never ActiveVBProject) to "AccessSQLite".
Public Function EnsureLibraryProjectName(Optional ByVal strProjectName As String = "AccessSQLite") As String
    Dim objProj As Object
    Dim strTargetFile As String
    On Error Resume Next
    strTargetFile = LCase$(CodeProject.Name)
    For Each objProj In Application.VBE.VBProjects
        If Right$(LCase$(CStr(objProj.FileName)), Len(strTargetFile)) = strTargetFile Then
            If objProj.Name <> strProjectName Then
                objProj.Name = strProjectName
            End If
            EnsureLibraryProjectName = objProj.Name
            Exit For
        End If
    Next objProj
    On Error GoTo 0
End Function

' Saves all VBA modules and compiles/saves the VBProject name permanently to disk.
Public Function SaveAllVbaModules() As Long
    Dim objProj As Object
    Dim objComp As Object
    Dim strTargetFile As String
    Dim lngSaved As Long
    On Error Resume Next
    EnsureLibraryProjectName "AccessSQLite"
    strTargetFile = LCase$(CodeProject.Name)
    For Each objProj In Application.VBE.VBProjects
        If Right$(LCase$(CStr(objProj.FileName)), Len(strTargetFile)) = strTargetFile Then
            For Each objComp In objProj.VBComponents
                DoCmd.Save acModule, objComp.Name
                lngSaved = lngSaved + 1
            Next objComp
            Exit For
        End If
    Next objProj
    DoCmd.RunCommand 126 ' acCmdCompileAndSaveAllModules
    On Error GoTo 0
    SaveAllVbaModules = lngSaved
End Function
