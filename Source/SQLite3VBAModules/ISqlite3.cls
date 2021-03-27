VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISqlite3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Property Get SqliteInitOk() As Long

End Property
Public Property Get SqliteInitError() As Long

End Property
Public Property Get SqliteInteger() As Long

End Property
Public Property Get SqliteFloat() As Long

End Property
Public Property Get SqliteText() As Long

End Property
Public Property Get SqliteBlob() As Long

End Property
Public Property Get SqliteNull() As Long

End Property
Public Property Get SqliteOk() As Long

End Property
Public Property Get SqliteError() As Long

End Property
Public Property Get SqliteInternal() As Long

End Property
Public Property Get SqlitePerm() As Long

End Property
Public Property Get SqliteAbort() As Long

End Property
Public Property Get SqliteBusy() As Long

End Property
Public Property Get SqliteLocked() As Long

End Property
Public Property Get SqliteNomem() As Long

End Property
Public Property Get SqliteReadonly() As Long

End Property
Public Property Get SqliteInterrupt() As Long

End Property
Public Property Get SqliteIoerr() As Long

End Property
Public Property Get SqliteCorrupt() As Long

End Property
Public Property Get SqliteNotfound() As Long

End Property
Public Property Get SqliteFull() As Long

End Property
Public Property Get SqliteCantopen() As Long

End Property
Public Property Get SqliteProtocol() As Long

End Property
Public Property Get SqliteEmpty() As Long

End Property
Public Property Get SqliteSchema() As Long

End Property
Public Property Get SqliteToobig() As Long

End Property
Public Property Get SqliteConstraint() As Long

End Property
Public Property Get SqliteMismatch() As Long

End Property
Public Property Get SqliteMisuse() As Long

End Property
Public Property Get SqliteNolfs() As Long

End Property
Public Property Get SqliteAuth() As Long

End Property
Public Property Get SqliteFormat() As Long

End Property
Public Property Get SqliteRange() As Long

End Property
Public Property Get SqliteNotadb() As Long

End Property
Public Property Get SqliteRow() As Long

End Property
Public Property Get SqliteDone() As Long

End Property
Public Property Get SqliteIoerrRead() As Long

End Property
Public Property Get SqliteIoerrShortRead() As Long

End Property
Public Property Get SqliteIoerrWrite() As Long

End Property
Public Property Get SqliteIoerrFsync() As Long

End Property
Public Property Get SqliteIoerrDirFsync() As Long

End Property
Public Property Get SqliteIoerrTruncate() As Long

End Property
Public Property Get SqliteIoerrFstat() As Long

End Property
Public Property Get SqliteIoerrUnlock() As Long

End Property
Public Property Get SqliteIoerrRdlock() As Long

End Property
Public Property Get SqliteIoerrDelete() As Long

End Property
Public Property Get SqliteIoerrBlocked() As Long

End Property
Public Property Get SqliteIoerrNomem() As Long

End Property
Public Property Get SqliteIoerrAccess() As Long

End Property
Public Property Get SqliteIoerrCheckreservedlock() As Long

End Property
Public Property Get SqliteIoerrLock() As Long

End Property
Public Property Get SqliteIoerrClose() As Long

End Property
Public Property Get SqliteIoerrDirClose() As Long

End Property
Public Property Get SqliteLockedSharedcache() As Long

End Property
Public Property Get SqliteOpenReadonly() As Long

End Property
Public Property Get SqliteOpenReadwrite() As Long

End Property
Public Property Get SqliteOpenCreate() As Long

End Property
Public Property Get SqliteOpenDeleteonclose() As Long

End Property
Public Property Get SqliteOpenExclusive() As Long

End Property
Public Property Get SqliteOpenAutoproxy() As Long

End Property
Public Property Get SqliteOpenUri() As Long

End Property
Public Property Get SqliteOpenMemory() As Long

End Property
Public Property Get SqliteOpenMainDb() As Long

End Property
Public Property Get SqliteOpenTempDb() As Long

End Property
Public Property Get SqliteOpenTransientDb() As Long

End Property
Public Property Get SqliteOpenMainJournal() As Long

End Property
Public Property Get SqliteOpenTempJournal() As Long

End Property
Public Property Get SqliteOpenSubjournal() As Long

End Property
Public Property Get SqliteOpenMasterJournal() As Long

End Property
Public Property Get SqliteOpenNomutex() As Long

End Property
Public Property Get SqliteOpenFullmutex() As Long

End Property
Public Property Get SqliteOpenSharedcache() As Long

End Property
Public Property Get SqliteOpenPrivatecache() As Long

End Property
Public Property Get SqliteOpenWal() As Long

End Property
Private Sub CheckOrChangeDir(Optional MasterDirectory As String = "")

End Sub

Public Property Let DatabaseDirectory(value As String)

End Property

Public Property Get DatabaseDirectory() As String

End Property

Public Function SQLite3Initialize(Optional ByVal libDir As String) As Long

End Function

Public Sub SQLite3Free()

End Sub


'=====================================================================================
' SQLite library version

Public Function SQLite3LibVersion() As String

End Function

'=====================================================================================
' Database connections
Public Function SQLite3Open(ByVal fileName As String, ByRef dbHandle As LongPtr) As Long

End Function

Public Function SQLite3OpenV2(ByVal fileName As String, ByRef dbHandle As LongPtr, ByVal flags As Long, ByVal vfsName As String) As Long

End Function

Public Function SQLite3Close(ByVal dbHandle As LongPtr) As Long

End Function

'=====================================================================================
' Error information

Public Function SQLite3ErrMsg(ByVal dbHandle As LongPtr) As String

End Function

Public Function SQLite3ErrCode(ByVal dbHandle As LongPtr) As Long

End Function

Public Function SQLite3ExtendedErrCode(ByVal dbHandle As LongPtr) As Long

End Function

'=====================================================================================
' Change Counts

Public Function SQLite3Changes(ByVal dbHandle As LongPtr) As Long

End Function

Public Function SQLite3TotalChanges(ByVal dbHandle As LongPtr) As Long

End Function

'=====================================================================================
' Statements

Public Function SQLite3PrepareV2(ByVal dbHandle As LongPtr, ByVal sql As String, ByRef stmtHandle As LongPtr) As Long

End Function

Public Function SQLite3Step(ByVal stmtHandle As LongPtr) As Long

End Function

Public Function SQLite3Reset(ByVal stmtHandle As LongPtr) As Long

End Function

Public Function SQLite3Finalize(ByVal stmtHandle As LongPtr) As Long

End Function

'=====================================================================================
' Statement column access (0-based indices)

Public Function SQLite3ColumnCount(ByVal stmtHandle As LongPtr) As Long

End Function

Public Function SQLite3ColumnType(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Long

End Function

Public Function SQLite3ColumnName(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As String

End Function

Public Function SQLite3ColumnDouble(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Double

End Function

Public Function SQLite3ColumnInt32(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Long

End Function

Public Function SQLite3ColumnText(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As String

End Function

Public Function SQLite3ColumnDate(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Date

End Function

Public Function SQLite3ColumnBlob(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Byte()

End Function
'=====================================================================================
' Statement bindings

Public Function SQLite3BindText(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal value As String) As Long

End Function

Public Function SQLite3BindDouble(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal value As Double) As Long

End Function

Public Function SQLite3BindInt32(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal value As Long) As Long

End Function

Public Function SQLite3BindDate(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal value As Date) As Long

End Function

Public Function SQLite3BindBlob(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByRef value() As Byte) As Long

End Function

Public Function SQLite3BindNull(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long) As Long

End Function

Public Function SQLite3BindParameterCount(ByVal stmtHandle As LongPtr) As Long

End Function

Public Function SQLite3BindParameterName(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long) As String

End Function

Public Function SQLite3BindParameterIndex(ByVal stmtHandle As LongPtr, ByVal paramName As String) As Long

End Function

Public Function SQLite3ClearBindings(ByVal stmtHandle As LongPtr) As Long

End Function

'=====================================================================================
' Backup
Public Function SQLite3Sleep(ByVal timeToSleepInMs As Long) As Long
End Function

Public Function SQLite3BackupInit(ByVal dbHandleDestination As LongPtr, ByVal destinationName As String, ByVal dbHandleSource As LongPtr, ByVal sourceName As String) As LongPtr

End Function

Public Function SQLite3BackupFinish(ByVal backupHandle As LongPtr) As Long

End Function

Public Function SQLite3BackupStep(ByVal backupHandle As LongPtr, ByVal numberOfPages) As Long

End Function

Public Function SQLite3BackupPageCount(ByVal backupHandle As LongPtr) As Long

End Function

Public Function SQLite3BackupRemaining(ByVal backupHandle As LongPtr) As Long

End Function

' String Helpers
Function Utf8PtrToString(ByVal pUtf8String As LongPtr) As String

End Function

Function StringToUtf8Bytes(ByVal str As String) As Variant

End Function

Function Utf16PtrToString(ByVal pUtf16String As LongPtr) As String

End Function

' Date Helpers
Public Function ToJulianDay(oleDate As Date) As Double

End Function

Public Function FromJulianDay(julianDay As Double) As Date

End Function
