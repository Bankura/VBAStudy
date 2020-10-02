VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiKernel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WindowAPI kernel クラス
'* [詳  細] WindowAPIのkernel32.dllの機能を提供する。
'*
'* [参  考] Win32API_PtrSafe.TXT
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Any, ByVal source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Sub CopyMemoryByRef Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef source As Any, ByVal Length As LongPtr)

Private Declare PtrSafe Function InterlockedIncrement64 Lib "kernel32" (lpAddend As LongPtr) As LongPtr
Private Declare PtrSafe Function InterlockedDecrement64 Lib "kernel32" (lpAddend As LongPtr) As LongPtr
Private Declare PtrSafe Function InterlockedExchange64 Lib "kernel32" (Target As LongPtr, ByVal Value As LongPtr) As LongPtr
Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
Private Declare PtrSafe Function InterlockedExchange Lib "kernel32" (Target As Long, ByVal Value As Long) As Long

' Loader Routines
Private Declare PtrSafe Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As LongPtr, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr

Private Declare PtrSafe Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare PtrSafe Function SetProcessShutdownParameters Lib "kernel32" (ByVal dwLevel As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function GetProcessShutdownParameters Lib "kernel32" (lpdwLevel As Long, lpdwFlags As Long) As Long

Private Declare PtrSafe Sub FatalAppExit Lib "kernel32" Alias "FatalAppExitA" (ByVal uAction As Long, ByVal lpMessageText As String)
Private Declare PtrSafe Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)
Private Declare PtrSafe Function GetCommandLine Lib "kernel32" Alias "GetCommandLineA" () As String
Private Declare PtrSafe Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare PtrSafe Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Private Declare PtrSafe Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long

Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As LongPtr, ByVal dwFlags As Long) As LongPtr

Private Declare PtrSafe Function LoadModule Lib "kernel32" (ByVal lpModuleName As String, lpParameterBlock As Any) As Long
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Private Declare PtrSafe Sub DebugBreak Lib "kernel32" ()
Private Declare PtrSafe Function ContinueDebugEvent Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwThreadId As Long, ByVal dwContinueStatus As Long) As Long
Private Declare PtrSafe Function DebugActiveProcess Lib "kernel32" (ByVal dwProcessId As Long) As Long

Private Declare PtrSafe Sub InitializeCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Private Declare PtrSafe Sub EnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Private Declare PtrSafe Sub LeaveCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Private Declare PtrSafe Sub DeleteCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Private Declare PtrSafe Function SetEvent Lib "kernel32" (ByVal hEvent As LongPtr) As Long
Private Declare PtrSafe Function ResetEvent Lib "kernel32" (ByVal hEvent As LongPtr) As Long
Private Declare PtrSafe Function PulseEvent Lib "kernel32" (ByVal hEvent As LongPtr) As Long
Private Declare PtrSafe Function ReleaseSemaphore Lib "kernel32" (ByVal hSemaphore As LongPtr, ByVal lReleaseCount As Long, lpPreviousCount As Long) As Long
Private Declare PtrSafe Function ReleaseMutex Lib "kernel32" (ByVal hMutex As LongPtr) As Long
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
Private Declare PtrSafe Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As LongPtr, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
Private Declare PtrSafe Function GetVersion Lib "kernel32" () As Long

Private Declare PtrSafe Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

Private Declare PtrSafe Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare PtrSafe Function SetHandleCount Lib "kernel32" (ByVal wNumber As Long) As Long
Private Declare PtrSafe Function GetLogicalDrives Lib "kernel32" () As Long
Private Declare PtrSafe Function LockFile Lib "kernel32" (ByVal hFile As LongPtr, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long) As Long
Private Declare PtrSafe Function UnlockFile Lib "kernel32" (ByVal hFile As LongPtr, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long) As Long
Private Declare PtrSafe Function LockFileEx Lib "kernel32" (ByVal hFile As LongPtr, ByVal dwFlags As Long, ByVal dwReserved As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long, lpOverlapped As OVERLAPPED) As Long

Private Declare PtrSafe Function UnlockFileEx Lib "kernel32" (ByVal hFile As LongPtr, ByVal dwReserved As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long, lpOverlapped As OVERLAPPED) As Long

Private Declare PtrSafe Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As LongPtr, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Private Declare PtrSafe Function GetFileType Lib "kernel32" (ByVal hFile As LongPtr) As Long
Private Declare PtrSafe Function GetFileSize Lib "kernel32" (ByVal hFile As LongPtr, lpFileSizeHigh As Long) As Long
Private Declare PtrSafe Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As LongPtr
Private Declare PtrSafe Function SetStdHandle Lib "kernel32" (ByVal nStdHandle As Long, ByVal nHandle As LongPtr) As Long
Private Declare PtrSafe Function WriteFile Lib "kernel32" (ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare PtrSafe Function FlushFileBuffers Lib "kernel32" (ByVal hFile As LongPtr) As Long
Private Declare PtrSafe Function DeviceIoControl Lib "kernel32" (ByVal hDevice As LongPtr, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare PtrSafe Function SetEndOfFile Lib "kernel32" (ByVal hFile As LongPtr) As Long
Private Declare PtrSafe Function SetFilePointer Lib "kernel32" (ByVal hFile As LongPtr, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare PtrSafe Function FindClose Lib "kernel32" (ByVal hFindFile As LongPtr) As Long
Private Declare PtrSafe Function GetFileTime Lib "kernel32" (ByVal hFile As LongPtr, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare PtrSafe Function SetFileTime Lib "kernel32" (ByVal hFile As LongPtr, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As LongPtr, ByVal hSourceHandle As LongPtr, ByVal hTargetProcessHandle As LongPtr, lpTargetHandle As LongPtr, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long

Private Declare PtrSafe Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalHandle Lib "kernel32" (wMem As Any) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalReAlloc Lib "kernel32" (ByVal hMem As LongPtr, ByVal dwBytes As LongPtr, ByVal wFlags As Long) As LongPtr

Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalFlags Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Declare PtrSafe Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function LocalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function LocalHandle Lib "kernel32" (wMem As Any) As LongPtr
Private Declare PtrSafe Function LocalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function LocalReAlloc Lib "kernel32" (ByVal hMem As LongPtr, ByVal wBytes As LongPtr, ByVal wFlags As Long) As LongPtr

Private Declare PtrSafe Function LocalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function LocalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function LocalFlags Lib "kernel32" (ByVal hMem As LongPtr) As Long

Private Declare PtrSafe Function FlushInstructionCache Lib "kernel32" (ByVal hProcess As LongPtr, lpBaseAddress As Any, ByVal dwSize As LongPtr) As Long
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" (lpAddress As Any, ByVal dwSize As LongPtr, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
Private Declare PtrSafe Function VirtualFree Lib "kernel32" (lpAddress As Any, ByVal dwSize As LongPtr, ByVal dwFreeType As Long) As Long
Private Declare PtrSafe Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As LongPtr, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare PtrSafe Function VirtualQuery Lib "kernel32" (lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As LongPtr) As LongPtr
Private Declare PtrSafe Function VirtualProtectEx Lib "kernel32" (ByVal hProcess As LongPtr, lpAddress As Any, ByVal dwSize As LongPtr, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare PtrSafe Function VirtualQueryEx Lib "kernel32" (ByVal hProcess As LongPtr, lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As LongPtr) As LongPtr
Private Declare PtrSafe Function HeapCreate Lib "kernel32" (ByVal flOptions As Long, ByVal dwInitialSize As LongPtr, ByVal dwMaximumSize As LongPtr) As LongPtr
Private Declare PtrSafe Function HeapDestroy Lib "kernel32" (ByVal hHeap As LongPtr) As Long
Private Declare PtrSafe Function HeapAlloc Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function HeapReAlloc Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, lpMem As Any, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function HeapFree Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare PtrSafe Function HeapSize Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, lpMem As Any) As LongPtr
Private Declare PtrSafe Function GetProcessHeap Lib "kernel32" () As LongPtr
Private Declare PtrSafe Function GetProcessTimes Lib "kernel32" (ByVal hProcess As LongPtr, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare PtrSafe Function TerminateProcess Lib "kernel32" (ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Long
Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As LongPtr, lpExitCode As Long) As Long

Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
Private Declare PtrSafe Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)

Private Declare PtrSafe Sub SetLastErrorEx Lib "user32" (ByVal dwErrCode As Long, ByVal dwType As Long)
Private Declare PtrSafe Function GetOverlappedResult Lib "kernel32" (ByVal hFile As LongPtr, lpOverlapped As OVERLAPPED, lpNumberOfBytesTransferred As Long, ByVal bWait As Long) As Long

Private Declare PtrSafe Sub SetDebugErrorLevel Lib "user32" (ByVal dwLevel As Long)
Private Declare PtrSafe Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long

Private Declare PtrSafe Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As LongPtr, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As LongPtr, lpNumberOfBytesWritten As LongPtr) As Long
Private Declare PtrSafe Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As LongPtr, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As LongPtr, lpNumberOfBytesWritten As LongPtr) As Long
Private Declare PtrSafe Function GetThreadContext Lib "kernel32" (ByVal hThread As LongPtr, lpContext As CONTEXT) As Long
Private Declare PtrSafe Function SetThreadContext Lib "kernel32" (ByVal hThread As LongPtr, lpContext As CONTEXT) As Long
Private Declare PtrSafe Function SuspendThread Lib "kernel32" (ByVal hThread As LongPtr) As Long
Private Declare PtrSafe Function ResumeThread Lib "kernel32" (ByVal hThread As LongPtr) As Long

Private Declare PtrSafe Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As LongPtr, ByVal lpName As String, ByVal lpType As String) As LongPtr
Private Declare PtrSafe Function FindResourceEx Lib "kernel32" Alias "FindResourceExA" (ByVal hModule As LongPtr, ByVal lpType As String, ByVal lpName As String, ByVal wLanguage As Long) As LongPtr
Private Declare PtrSafe Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceA" (ByVal pFileName As String, ByVal bDeleteExistingResources As Long) As LongPtr
Private Declare PtrSafe Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" (ByVal hUpdate As LongPtr, ByVal lpType As String, ByVal lpName As String, ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare PtrSafe Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceA" (ByVal hUpdate As LongPtr, ByVal fDiscard As Long) As Long
Private Declare PtrSafe Function LoadResource Lib "kernel32" (ByVal hInstance As LongPtr, ByVal hResInfo As LongPtr) As LongPtr
Private Declare PtrSafe Function LockResource Lib "kernel32" (ByVal hResData As LongPtr) As LongPtr
Private Declare PtrSafe Function SizeofResource Lib "kernel32" (ByVal hInstance As LongPtr, ByVal hResInfo As LongPtr) As Long

Private Declare PtrSafe Function InitAtomTable Lib "kernel32" (ByVal nSize As Long) As Long
Private Declare PtrSafe Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Private Declare PtrSafe Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Private Declare PtrSafe Function GlobalFindAtom Lib "kernel32" Alias "GlobalFindAtomA" (ByVal lpString As String) As Integer
Private Declare PtrSafe Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

' User Profile Routines
' NOTE: The lpKeyName argument for GetProfileString, WriteProfileString,
'       GetPrivateProfileString, and WritePrivateProfileString can be either
'       a string or NULL.  This is why the argument is defined as "As Any".
'          For example, to pass a string specify   ByVal "wallpaper"
'          To pass NULL specify                    ByVal 0&
'       You can also pass NULL for the lpString argument for WriteProfileString
'       and WritePrivateProfileString
Private Declare PtrSafe Function GetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long

Private Declare PtrSafe Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Private Declare PtrSafe Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Private Declare PtrSafe Function GetProfileSection Lib "kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare PtrSafe Function WriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
Private Declare PtrSafe Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare PtrSafe Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Private Declare PtrSafe Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare PtrSafe Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare PtrSafe Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Private Declare PtrSafe Function GetCurrentDirectory Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare PtrSafe Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare PtrSafe Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare PtrSafe Function CreateDirectoryEx Lib "kernel32" Alias "CreateDirectoryExA" (ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare PtrSafe Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Private Declare PtrSafe Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long

Private Declare PtrSafe Function DefineDosDevice Lib "kernel32" Alias "DefineDosDeviceA" (ByVal dwFlags As Long, ByVal lpDeviceName As String, ByVal lpTargetPath As String) As Long
Private Declare PtrSafe Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" (ByVal lpDeviceName As String, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long
Private Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
Private Declare PtrSafe Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare PtrSafe Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare PtrSafe Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare PtrSafe Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As LongPtr
Private Declare PtrSafe Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As LongPtr, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare PtrSafe Function SearchPath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare PtrSafe Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare PtrSafe Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare PtrSafe Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function CreateNamedPipe Lib "kernel32" Alias "CreateNamedPipeA" (ByVal lpName As String, ByVal dwOpenMode As Long, ByVal dwPipeMode As Long, ByVal nMaxInstances As Long, ByVal nOutBufferSize As Long, ByVal nInBufferSize As Long, ByVal nDefaultTimeOut As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) As LongPtr
Private Declare PtrSafe Function GetNamedPipeHandleState Lib "kernel32" Alias "GetNamedPipeHandleStateA" (ByVal hNamedPipe As LongPtr, lpState As Long, lpCurInstances As Long, lpMaxCollectionCount As Long, lpCollectDataTimeout As Long, ByVal lpUserName As String, ByVal nMaxUserNameSize As Long) As Long
Private Declare PtrSafe Function CallNamedPipe Lib "kernel32" Alias "CallNamedPipeA" (ByVal lpNamedPipeName As String, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesRead As Long, ByVal nTimeOut As Long) As Long
Private Declare PtrSafe Function WaitNamedPipe Lib "kernel32" Alias "WaitNamedPipeA" (ByVal lpNamedPipeName As String, ByVal nTimeOut As Long) As Long
Private Declare PtrSafe Function SetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
Private Declare PtrSafe Sub SetFileApisToOEM Lib "kernel32" ()
Private Declare PtrSafe Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare PtrSafe Function ClearEventLog Lib "advapi32.dll" Alias "ClearEventLogA" (ByVal hEventLog As LongPtr, ByVal lpBackupFileName As String) As Long
Private Declare PtrSafe Function BackupEventLog Lib "advapi32.dll" Alias "BackupEventLogA" (ByVal hEventLog As LongPtr, ByVal lpBackupFileName As String) As Long
Private Declare PtrSafe Function CloseEventLog Lib "advapi32.dll" (ByVal hEventLog As LongPtr) As Long
Private Declare PtrSafe Function DeregisterEventSource Lib "advapi32.dll" (ByVal hEventLog As LongPtr) As Long
Private Declare PtrSafe Function GetNumberOfEventLogRecords Lib "advapi32.dll" (ByVal hEventLog As LongPtr, NumberOfRecords As Long) As Long
Private Declare PtrSafe Function GetOldestEventLogRecord Lib "advapi32.dll" (ByVal hEventLog As LongPtr, OldestRecord As Long) As Long
Private Declare PtrSafe Function OpenEventLog Lib "advapi32.dll" Alias "OpenEventLogA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As LongPtr
Private Declare PtrSafe Function RegisterEventSource Lib "advapi32.dll" Alias "RegisterEventSourceA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As LongPtr
Private Declare PtrSafe Function OpenBackupEventLog Lib "advapi32.dll" Alias "OpenBackupEventLogA" (ByVal lpUNCServerName As String, ByVal lpFileName As String) As LongPtr
Private Declare PtrSafe Function ReadEventLog Lib "advapi32.dll" Alias "ReadEventLogA" (ByVal hEventLog As LongPtr, ByVal dwReadFlags As Long, ByVal dwRecordOffset As Long, lpBuffer As EVENTLOGRECORD, ByVal nNumberOfBytesToRead As Long, pnBytesRead As Long, pnMinNumberOfBytesNeeded As Long) As Long
Private Declare PtrSafe Function ReportEvent Lib "advapi32.dll" Alias "ReportEventA" (ByVal hEventLog As LongPtr, ByVal wType As Long, ByVal wCategory As Long, ByVal dwEventID As Long, lpUserSid As Any, ByVal wNumStrings As Long, ByVal dwDataSize As Long, ByVal lpStrings As LongPtr, lpRawData As Any) As Long

Private Declare PtrSafe Function DuplicateToken Lib "advapi32.dll" (ByVal ExistingTokenHandle As LongPtr, ImpersonationLevel As Integer, DuplicateTokenHandle As LongPtr) As Long
Private Declare PtrSafe Function GetKernelObjectSecurity Lib "advapi32.dll" (ByVal handle As LongPtr, ByVal RequestedInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Private Declare PtrSafe Function ImpersonateNamedPipeClient Lib "advapi32.dll" (ByVal hNamedPipe As LongPtr) As Long
Private Declare PtrSafe Function ImpersonateSelf Lib "advapi32.dll" (ImpersonationLevel As Integer) As Long
Private Declare PtrSafe Function RevertToSelf Lib "advapi32.dll" () As Long
Private Declare PtrSafe Function AccessCheck Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal ClientToken As LongPtr, ByVal DesiredAccess As Long, GenericMapping As GENERIC_MAPPING, PrivilegeSet As PRIVILEGE_SET, PrivilegeSetLength As Long, GrantedAccess As Long, ByVal Status As LongPtr) As Long

Private Declare PtrSafe Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As LongPtr, ByVal DesiredAccess As Long, TokenHandle As LongPtr) As Long
Private Declare PtrSafe Function OpenThreadToken Lib "advapi32.dll" (ByVal ThreadHandle As LongPtr, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As LongPtr) As Long

Private Declare PtrSafe Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As LongPtr, TokenInformationClass As Integer, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare PtrSafe Function SetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As LongPtr, TokenInformationClass As Integer, TokenInformation As Any, ByVal TokenInformationLength As Long) As Long
Private Declare PtrSafe Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As LongPtr, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare PtrSafe Function AdjustTokenGroups Lib "advapi32.dll" (ByVal TokenHandle As LongPtr, ByVal ResetToDefault As Long, NewState As TOKEN_GROUPS, ByVal BufferLength As Long, PreviousState As TOKEN_GROUPS, ReturnLength As Long) As Long
Private Declare PtrSafe Function PrivilegeCheck Lib "advapi32.dll" (ByVal ClientToken As LongPtr, RequiredPrivileges As PRIVILEGE_SET, ByVal pfResult As LongPtr) As Long
Private Declare PtrSafe Function AccessCheckAndAuditAlarm Lib "advapi32.dll" Alias "AccessCheckAndAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal ObjectTypeName As String, ByVal ObjectName As String, SecurityDescriptor As SECURITY_DESCRIPTOR, ByVal DesiredAccess As Long, GenericMapping As GENERIC_MAPPING, ByVal ObjectCreation As Long, GrantedAccess As Long, ByVal AccessStatus As LongPtr, ByVal pfGenerateOnClose As LongPtr) As Long
Private Declare PtrSafe Function ObjectOpenAuditAlarm Lib "kernel32" Alias "ObjectOpenAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal ObjectTypeName As String, ByVal ObjectName As String, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal ClientToken As LongPtr, ByVal DesiredAccess As Long, ByVal GrantedAccess As Long, Privileges As PRIVILEGE_SET, ByVal ObjectCreation As Long, ByVal AccessGranted As Long, ByVal GenerateOnClose As LongPtr) As Long
Private Declare PtrSafe Function ObjectPrivilegeAuditAlarm Lib "advapi32.dll" Alias "ObjectPrivilegeAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal ClientToken As LongPtr, ByVal DesiredAccess As Long, Privileges As PRIVILEGE_SET, ByVal AccessGranted As Long) As Long
Private Declare PtrSafe Function ObjectCloseAuditAlarm Lib "advapi32.dll" Alias "ObjectCloseAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal GenerateOnClose As Long) As Long
Private Declare PtrSafe Function PrivilegedServiceAuditAlarm Lib "advapi32.dll" Alias "PrivilegedServiceAuditAlarmA" (ByVal SubsystemName As String, ByVal serviceName As String, ByVal ClientToken As LongPtr, Privileges As PRIVILEGE_SET, ByVal AccessGranted As Long) As Long
Private Declare PtrSafe Function IsValidSid Lib "advapi32.dll" (pSid As Any) As Long
Private Declare PtrSafe Function EqualSid Lib "advapi32.dll" (pSid1 As Any, pSid2 As Any) As Long
Private Declare PtrSafe Function EqualPrefixSid Lib "advapi32.dll" (pSid1 As Any, pSid2 As Any) As Long
Private Declare PtrSafe Function GetSidLengthRequired Lib "advapi32.dll" (ByVal nSubAuthorityCount As Byte) As Long
Private Declare PtrSafe Function AllocateAndInitializeSid Lib "advapi32.dll" (pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, lpPSid As LongPtr) As Long
Private Declare PtrSafe Sub FreeSid Lib "advapi32.dll" (pSid As Any)
Private Declare PtrSafe Function InitializeSid Lib "advapi32.dll" (Sid As Any, pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte) As Long
Private Declare PtrSafe Function GetSidIdentifierAuthority Lib "advapi32.dll" (pSid As Any) As SID_IDENTIFIER_AUTHORITY
Private Declare PtrSafe Function GetSidSubAuthority Lib "advapi32.dll" (pSid As Any, ByVal nSubAuthority As Long) As LongPtr
Private Declare PtrSafe Function GetSidSubAuthorityCount Lib "advapi32.dll" (pSid As Any) As LongPtr
Private Declare PtrSafe Function GetLengthSid Lib "advapi32.dll" (pSid As Any) As Long
Private Declare PtrSafe Function CopySid Lib "advapi32.dll" (ByVal nDestinationSidLength As Long, pDestinationSid As Any, pSourceSid As Any) As Long
Private Declare PtrSafe Function AreAllAccessesGranted Lib "advapi32.dll" (ByVal GrantedAccess As Long, ByVal DesiredAccess As Long) As Long
Private Declare PtrSafe Function AreAnyAccessesGranted Lib "advapi32.dll" (ByVal GrantedAccess As Long, ByVal DesiredAccess As Long) As Long
Private Declare PtrSafe Sub MapGenericMask Lib "advapi32.dll" (AccessMask As Long, GenericMapping As GENERIC_MAPPING)
Private Declare PtrSafe Function IsValidAcl Lib "advapi32.dll" (pAcl As ACL) As Long
Private Declare PtrSafe Function InitializeAcl Lib "advapi32.dll" (pAcl As ACL, ByVal nAclLength As Long, ByVal dwAclRevision As Long) As Long
Private Declare PtrSafe Function GetAclInformation Lib "advapi32.dll" (pAcl As ACL, pAclInformation As Any, ByVal nAclInformationLength As Long, ByVal dwAclInformationClass As Integer) As Long
Private Declare PtrSafe Function SetAclInformation Lib "advapi32.dll" (pAcl As ACL, pAclInformation As Any, ByVal nAclInformationLength As Long, ByVal dwAclInformationClass As Integer) As Long
Private Declare PtrSafe Function AddAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal dwStartingAceIndex As Long, pAceList As Any, ByVal nAceListLength As Long) As Long
Private Declare PtrSafe Function DeleteAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceIndex As Long) As Long
Private Declare PtrSafe Function GetAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceIndex As Long, pAce As Any) As Long
Private Declare PtrSafe Function AddAccessAllowedAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal AccessMask As Long, pSid As Any) As Long
Private Declare PtrSafe Function AddAccessDeniedAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal AccessMask As Long, pSid As Any) As Long
Private Declare PtrSafe Function AddAuditAccessAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal dwAccessMask As Long, pSid As Any, ByVal bAuditSuccess As Long, ByVal bAuditFailure As Long) As Long
Private Declare PtrSafe Function FindFirstFreeAce Lib "advapi32.dll" (pAcl As ACL, pAce As LongPtr) As Long
Private Declare PtrSafe Function InitializeSecurityDescriptor Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal dwRevision As Long) As Long
Private Declare PtrSafe Function IsValidSecurityDescriptor Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Private Declare PtrSafe Function GetSecurityDescriptorLength Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long

Private Declare PtrSafe Function GetSecurityDescriptorControl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pControl As Integer, lpdwRevision As Long) As Long
Private Declare PtrSafe Function SetSecurityDescriptorDacl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal bDaclPresent As Long, pDacl As ACL, ByVal bDaclDefaulted As Long) As Long
Private Declare PtrSafe Function GetSecurityDescriptorDacl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, lpbDaclPresent As Long, pDacl As ACL, lpbDaclDefaulted As Long) As Long
Private Declare PtrSafe Function SetSecurityDescriptorSacl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal bSaclPresent As Long, pSacl As ACL, ByVal bSaclDefaulted As Long) As Long
Private Declare PtrSafe Function GetSecurityDescriptorSacl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal lpbSaclPresent As LongPtr, pSacl As ACL, ByVal lpbSaclDefaulted As LongPtr) As Long
Private Declare PtrSafe Function SetSecurityDescriptorOwner Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pOwner As Any, ByVal bOwnerDefaulted As Long) As Long
Private Declare PtrSafe Function GetSecurityDescriptorOwner Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pOwner As Any, ByVal lpbOwnerDefaulted As LongPtr) As Long
Private Declare PtrSafe Function SetSecurityDescriptorGroup Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pGroup As Any, ByVal bGroupDefaulted As Long) As Long
Private Declare PtrSafe Function GetSecurityDescriptorGroup Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pGroup As Any, ByVal lpbGroupDefaulted As LongPtr) As Long
Private Declare PtrSafe Function CreatePrivateObjectSecurity Lib "advapi32.dll" (ParentDescriptor As SECURITY_DESCRIPTOR, CreatorDescriptor As SECURITY_DESCRIPTOR, NewDescriptor As SECURITY_DESCRIPTOR, ByVal IsDirectoryObject As Long, ByVal Token As LongPtr, GenericMapping As GENERIC_MAPPING) As Long
Private Declare PtrSafe Function SetPrivateObjectSecurity Lib "advapi32.dll" (ByVal SecurityInformation As Long, ModificationDescriptor As SECURITY_DESCRIPTOR, ObjectsSecurityDescriptor As SECURITY_DESCRIPTOR, GenericMapping As GENERIC_MAPPING, ByVal Token As LongPtr) As Long
Private Declare PtrSafe Function GetPrivateObjectSecurity Lib "advapi32.dll" (ObjectDescriptor As SECURITY_DESCRIPTOR, ByVal SecurityInformation As Long, ResultantDescriptor As SECURITY_DESCRIPTOR, ByVal DescriptorLength As Long, ReturnLength As Long) As Long
Private Declare PtrSafe Function DestroyPrivateObjectSecurity Lib "advapi32.dll" (ObjectDescriptor As SECURITY_DESCRIPTOR) As Long
Private Declare PtrSafe Function MakeSelfRelativeSD Lib "advapi32.dll" (pAbsoluteSecurityDescriptor As SECURITY_DESCRIPTOR, pSelfRelativeSecurityDescriptor As SECURITY_DESCRIPTOR, lpdwBufferLength As Long) As Long
Private Declare PtrSafe Function MakeAbsoluteSD Lib "advapi32.dll" (pSelfRelativeSecurityDescriptor As SECURITY_DESCRIPTOR, pAbsoluteSecurityDescriptor As SECURITY_DESCRIPTOR, lpdwAbsoluteSecurityDescriptorSize As Long, pDacl As ACL, lpdwDaclSize As Long, pSacl As ACL, lpdwSaclSize As Long, pOwner As Any, lpdwOwnerSize As Long, pPrimaryGroup As Any, lpdwPrimaryGroupSize As Long) As Long
Private Declare PtrSafe Function SetFileSecurity Lib "advapi32.dll" Alias "SetFileSecurityA" (ByVal lpFileName As String, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Private Declare PtrSafe Function GetFileSecurity Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Private Declare PtrSafe Function SetKernelObjectSecurity Lib "advapi32.dll" (ByVal handle As LongPtr, ByVal SecurityInformation As Long, SecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Private Declare PtrSafe Function FindFirstChangeNotification Lib "kernel32" Alias "FindFirstChangeNotificationA" (ByVal lpPathName As String, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long) As LongPtr
Private Declare PtrSafe Function FindNextChangeNotification Lib "kernel32" (ByVal hChangeHandle As LongPtr) As Long
Private Declare PtrSafe Function FindCloseChangeNotification Lib "kernel32" (ByVal hChangeHandle As LongPtr) As Long
Private Declare PtrSafe Function VirtualLock Lib "kernel32" (lpAddress As Any, ByVal dwSize As LongPtr) As Long
Private Declare PtrSafe Function VirtualUnlock Lib "kernel32" (lpAddress As Any, ByVal dwSize As LongPtr) As Long
Private Declare PtrSafe Function MapViewOfFileEx Lib "kernel32" (ByVal hFileMappingObject As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As LongPtr, lpBaseAddress As Any) As LongPtr
Private Declare PtrSafe Function SetPriorityClass Lib "kernel32" (ByVal hProcess As LongPtr, ByVal dwPriorityClass As Long) As Long
Private Declare PtrSafe Function GetPriorityClass Lib "kernel32" (ByVal hProcess As LongPtr) As Long

Private Declare PtrSafe Sub FatalExit Lib "kernel32" (ByVal code As Long)
Private Declare PtrSafe Function GetEnvironmentStrings Lib "kernel32" Alias "GetEnvironmentStringsA" () As String
Private Declare PtrSafe Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As LongPtr)
Private Declare PtrSafe Function UnhandledExceptionFilter Lib "kernel32" (ExceptionInfo As EXCEPTION_POINTERS) As Long

Private Declare PtrSafe Function CreateThread Lib "kernel32" (lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As LongPtr, lpStartAddress As LongPtr, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As LongPtr
Private Declare PtrSafe Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As LongPtr, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As LongPtr, lpStartAddress As LongPtr, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As LongPtr
Private Declare PtrSafe Function GetCurrentThread Lib "kernel32" () As LongPtr
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare PtrSafe Function SetThreadPriority Lib "kernel32" (ByVal hThread As LongPtr, ByVal nPriority As Long) As Long
Private Declare PtrSafe Function GetThreadPriority Lib "kernel32" (ByVal hThread As LongPtr) As Long
Private Declare PtrSafe Function GetThreadTimes Lib "kernel32" (ByVal hThread As LongPtr, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Private Declare PtrSafe Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Private Declare PtrSafe Function TerminateThread Lib "kernel32" (ByVal hThread As LongPtr, ByVal dwExitCode As Long) As Long
Private Declare PtrSafe Function GetExitCodeThread Lib "kernel32" (ByVal hThread As LongPtr, lpExitCode As Long) As Long
Private Declare PtrSafe Function GetThreadSelectorEntry Lib "kernel32" (ByVal hThread As LongPtr, ByVal dwSelector As Long, lpSelectorEntry As LDT_ENTRY) As Long

' COMM declarations
Private Declare PtrSafe Function SetCommState Lib "kernel32" (ByVal hCommDev As LongPtr, lpDCB As DCB) As Long
Private Declare PtrSafe Function SetCommTimeouts Lib "kernel32" (ByVal hFile As LongPtr, lpCommTimeouts As COMMTIMEOUTS) As Long
Private Declare PtrSafe Function GetCommState Lib "kernel32" (ByVal nCid As LongPtr, lpDCB As DCB) As Long
Private Declare PtrSafe Function GetCommTimeouts Lib "kernel32" (ByVal hFile As LongPtr, lpCommTimeouts As COMMTIMEOUTS) As Long
Private Declare PtrSafe Function PurgeComm Lib "kernel32" (ByVal hFile As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BuildCommDCB Lib "kernel32" Alias "BuildCommDCBA" (ByVal lpDef As String, lpDCB As DCB) As Long
Private Declare PtrSafe Function BuildCommDCBAndTimeouts Lib "kernel32" Alias "BuildCommDCBAndTimeoutsA" (ByVal lpDef As String, lpDCB As DCB, lpCommTimeouts As COMMTIMEOUTS) As Long
Private Declare PtrSafe Function TransmitCommChar Lib "kernel32" (ByVal nCid As LongPtr, ByVal cChar As Byte) As Long
Private Declare PtrSafe Function SetCommBreak Lib "kernel32" (ByVal nCid As LongPtr) As Long
Private Declare PtrSafe Function SetCommMask Lib "kernel32" (ByVal hFile As LongPtr, ByVal dwEvtMask As Long) As Long
Private Declare PtrSafe Function ClearCommBreak Lib "kernel32" (ByVal nCid As LongPtr) As Long
Private Declare PtrSafe Function ClearCommError Lib "kernel32" (ByVal hFile As LongPtr, lpErrors As Long, lpStat As COMSTAT) As Long
Private Declare PtrSafe Function SetupComm Lib "kernel32" (ByVal hFile As Long, ByVal dwInQueue As Long, ByVal dwOutQueue As Long) As Long
Private Declare PtrSafe Function EscapeCommFunction Lib "kernel32" (ByVal nCid As LongPtr, ByVal nFunc As Long) As Long
Private Declare PtrSafe Function GetCommMask Lib "kernel32" (ByVal hFile As LongPtr, lpEvtMask As Long) As Long
Private Declare PtrSafe Function GetCommProperties Lib "kernel32" (ByVal hFile As LongPtr, lpCommProp As COMMPROP) As Long
Private Declare PtrSafe Function GetCommModemStatus Lib "kernel32" (ByVal hFile As LongPtr, lpModemStat As Long) As Long
Private Declare PtrSafe Function WaitCommEvent Lib "kernel32" (ByVal hFile As LongPtr, lpEvtMask As Long, lpOverlapped As OVERLAPPED) As Long

Private Declare PtrSafe Function SetTapePosition Lib "kernel32" (ByVal hDevice As LongPtr, ByVal dwPositionMethod As Long, ByVal dwPartition As Long, ByVal dwOffsetLow As Long, ByVal dwOffsetHigh As Long, ByVal bimmediate As Long) As Long
Private Declare PtrSafe Function GetTapePosition Lib "kernel32" (ByVal hDevice As LongPtr, ByVal dwPositionType As Long, lpdwPartition As Long, lpdwOffsetLow As Long, lpdwOffsetHigh As Long) As Long
Private Declare PtrSafe Function PrepareTape Lib "kernel32" (ByVal hDevice As LongPtr, ByVal dwOperation As Long, ByVal bimmediate As Long) As Long
Private Declare PtrSafe Function EraseTape Lib "kernel32" (ByVal hDevice As LongPtr, ByVal dwEraseType As Long, ByVal bimmediate As Long) As Long
Private Declare PtrSafe Function CreateTapePartition Lib "kernel32" (ByVal hDevice As LongPtr, ByVal dwPartitionMethod As Long, ByVal dwCount As Long, ByVal dwSize As Long) As Long
Private Declare PtrSafe Function WriteTapemark Lib "kernel32" (ByVal hDevice As LongPtr, ByVal dwTapemarkType As Long, ByVal dwTapemarkCount As Long, ByVal bimmediate As Long) As Long
Private Declare PtrSafe Function GetTapeStatus Lib "kernel32" (ByVal hDevice As LongPtr) As Long
Private Declare PtrSafe Function GetTapeParameters Lib "kernel32" (ByVal hDevice As LongPtr, ByVal dwOperation As Long, lpdwSize As Long, lpTapeInformation As Any) As Long

Private Declare PtrSafe Function SetTapeParameters Lib "kernel32" (ByVal hDevice As LongPtr, ByVal dwOperation As Long, lpTapeInformation As Any) As Long

Private Declare PtrSafe Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Private Declare PtrSafe Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (lpSystemTime As SystemTime)
Private Declare PtrSafe Function SetSystemTime Lib "kernel32" (lpSystemTime As SystemTime) As Long
Private Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystemTime As SystemTime)
Private Declare PtrSafe Function SetLocalTime Lib "kernel32" (lpSystemTime As SystemTime) As Long
Private Declare PtrSafe Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Private Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare PtrSafe Function SetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

' Routines to convert back and forth
' between system time and file time

Private Declare PtrSafe Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SystemTime, lpFileTime As FILETIME) As Long
Private Declare PtrSafe Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare PtrSafe Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare PtrSafe Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SystemTime) As Long
Private Declare PtrSafe Function CompareFileTime Lib "kernel32" (lpFileTime1 As FILETIME, lpFileTime2 As FILETIME) As Long
Private Declare PtrSafe Function FileTimeToDosDateTime Lib "kernel32" (lpFileTime As FILETIME, ByVal lpFatDate As LongPtr, ByVal lpFatTime As LongPtr) As Long
Private Declare PtrSafe Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FILETIME) As Long

Private Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongPtr
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As LongPtr) As Long

Private Declare PtrSafe Function CreatePipe Lib "kernel32" (phReadPipe As LongPtr, phWritePipe As LongPtr, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare PtrSafe Function ConnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As LongPtr, lpOverlapped As OVERLAPPED) As Long
Private Declare PtrSafe Function DisconnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As LongPtr) As Long
Private Declare PtrSafe Function SetNamedPipeHandleState Lib "kernel32" (ByVal hNamedPipe As LongPtr, lpMode As Long, lpMaxCollectionCount As Long, lpCollectDataTimeout As Long) As Long
Private Declare PtrSafe Function GetNamedPipeInfo Lib "kernel32" (ByVal hNamedPipe As LongPtr, lpFlags As Long, lpOutBufferSize As Long, lpInBufferSize As Long, lpMaxInstances As Long) As Long
Private Declare PtrSafe Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As LongPtr, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
Private Declare PtrSafe Function TransactNamedPipe Lib "kernel32" (ByVal hNamedPipe As LongPtr, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesRead As Long, lpOverlapped As OVERLAPPED) As Long

Private Declare PtrSafe Function CreateMailslot Lib "kernel32" Alias "CreateMailslotA" (ByVal lpName As String, ByVal nMaxMessageSize As Long, ByVal lReadTimeout As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) As LongPtr
Private Declare PtrSafe Function GetMailslotInfo Lib "kernel32" (ByVal hMailslot As LongPtr, lpMaxMessageSize As Long, lpNextSize As Long, lpMessageCount As Long, lpReadTimeout As Long) As Long
Private Declare PtrSafe Function SetMailslotInfo Lib "kernel32" (ByVal hMailslot As LongPtr, ByVal lReadTimeout As Long) As Long
Private Declare PtrSafe Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As LongPtr) As LongPtr

Private Declare PtrSafe Function FlushViewOfFile Lib "kernel32" (lpBaseAddress As Any, ByVal dwNumberOfBytesToFlush As LongPtr) As Long
Private Declare PtrSafe Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long

Private Declare PtrSafe Function lstrcmp Lib "kernel32" Alias "lstrcmpA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare PtrSafe Function lstrcmpi Lib "kernel32" Alias "lstrcmpiA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare PtrSafe Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Private Declare PtrSafe Function lopen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Private Declare PtrSafe Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Private Declare PtrSafe Function lcreat Lib "kernel32" Alias "_lcreat" (ByVal lpPathName As String, ByVal iAttribute As Long) As Long
Private Declare PtrSafe Function llseek Lib "kernel32" Alias "_llseek" (ByVal hFile As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Private Declare PtrSafe Function lread Lib "kernel32" Alias "_lread" (ByVal hFile As Long, lpBuffer As Any, ByVal wBytes As Long) As Long
Private Declare PtrSafe Function lwrite Lib "kernel32" Alias "_lwrite" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal wBytes As Long) As Long

Private Declare PtrSafe Function hread Lib "kernel32" Alias "_hread" (ByVal hFile As Long, lpBuffer As Any, ByVal lBytes As Long) As Long
Private Declare PtrSafe Function hwrite Lib "kernel32" Alias "_hwrite" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal lBytes As Long) As Long

Private Declare PtrSafe Function TlsAlloc Lib "kernel32" () As Long

Private Declare PtrSafe Function TlsGetValue Lib "kernel32" (ByVal dwTlsIndex As Long) As LongPtr
Private Declare PtrSafe Function TlsSetValue Lib "kernel32" (ByVal dwTlsIndex As Long, lpTlsValue As Any) As Long
Private Declare PtrSafe Function TlsFree Lib "kernel32" (ByVal dwTlsIndex As Long) As Long
Private Declare PtrSafe Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Private Declare PtrSafe Function WaitForSingleObjectEx Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Private Declare PtrSafe Function WaitForMultipleObjectsEx Lib "kernel32" (ByVal nCount As Long, lpHandles As LongPtr, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Private Declare PtrSafe Function BackupRead Lib "kernel32" (ByVal hFile As LongPtr, lpBuffer As Byte, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, lpContext As Any) As Long
Private Declare PtrSafe Function BackupSeek Lib "kernel32" (ByVal hFile As LongPtr, ByVal dwLowBytesToSeek As Long, ByVal dwHighBytesToSeek As Long, lpdwLowByteSeeked As Long, lpdwHighByteSeeked As Long, lpContext As LongPtr) As Long
Private Declare PtrSafe Function BackupWrite Lib "kernel32" (ByVal hFile As LongPtr, lpBuffer As Byte, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, lpContext As LongPtr) As Long

Private Declare PtrSafe Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As LongPtr
Private Declare PtrSafe Function OpenMutex Lib "kernel32" Alias "OpenMutexA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As LongPtr
Private Declare PtrSafe Function CreateEvent Lib "kernel32" Alias "CreateEventA" (lpEventAttributes As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As LongPtr
Private Declare PtrSafe Function OpenEvent Lib "kernel32" Alias "OpenEventA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As LongPtr
Private Declare PtrSafe Function CreateSemaphore Lib "kernel32" Alias "CreateSemaphoreA" (lpSemaphoreAttributes As SECURITY_ATTRIBUTES, ByVal lInitialCount As Long, ByVal lMaximumCount As Long, ByVal lpName As String) As LongPtr
Private Declare PtrSafe Function OpenSemaphore Lib "kernel32" Alias "OpenSemaphoreA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As LongPtr
Private Declare PtrSafe Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As LongPtr, lpFileMappigAttributes As SECURITY_ATTRIBUTES, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As LongPtr
Private Declare PtrSafe Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As LongPtr
Private Declare PtrSafe Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare PtrSafe Function IsBadReadPtr Lib "kernel32" (lp As Any, ByVal ucb As LongPtr) As Long
Private Declare PtrSafe Function IsBadWritePtr Lib "kernel32" (lp As Any, ByVal ucb As LongPtr) As Long
Private Declare PtrSafe Function IsBadStringPtr Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As String, ByVal ucchMax As LongPtr) As Long
Private Declare PtrSafe Function IsBadHugeReadPtr Lib "kernel32" (lp As Any, ByVal ucb As LongPtr) As Long
Private Declare PtrSafe Function IsBadHugeWritePtr Lib "kernel32" (lp As Any, ByVal ucb As LongPtr) As Long

Private Declare PtrSafe Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, Sid As Any, ByVal Name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long

Private Declare PtrSafe Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameA" (ByVal lpSystemName As String, ByVal lpAccountName As String, ByVal Sid As LongPtr, cbSid As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long
Private Declare PtrSafe Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
Private Declare PtrSafe Function LookupPrivilegeName Lib "advapi32.dll" Alias "LookupPrivilegeNameA" (ByVal lpSystemName As String, lpLuid As LARGE_INTEGER, ByVal lpName As String, cbName As Long) As Long
Private Declare PtrSafe Function LookupPrivilegeDisplayName Lib "advapi32.dll" Alias "LookupPrivilegeDisplayNameA" (ByVal lpSystemName As String, ByVal lpName As String, ByVal lpDisplayName As String, cbDisplayName As Long, lpLanguageID As Long) As Long
Private Declare PtrSafe Function AllocateLocallyUniqueId Lib "advapi32.dll" (Luid As LARGE_INTEGER) As Long

Private Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare PtrSafe Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

' Performance counter API's
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long

' Code Page Dependent APIs

Private Declare PtrSafe Function IsValidCodePage Lib "kernel32" (ByVal CodePage As Long) As Long
Private Declare PtrSafe Function GetACP Lib "kernel32" () As Long
Private Declare PtrSafe Function GetOEMCP Lib "kernel32" () As Long
Private Declare PtrSafe Function GetCPInfo Lib "kernel32" (ByVal CodePage As Long, lpCPInfo As CPINFO) As Long
Private Declare PtrSafe Function IsDBCSLeadByte Lib "kernel32" (ByVal bTestChar As Byte) As Long
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As LongPtr) As Long

' Locale Dependent APIs

Private Declare PtrSafe Function CompareString Lib "kernel32" Alias "CompareStringA" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As String, ByVal cchCount1 As Long, ByVal lpString2 As String, ByVal cchCount2 As Long) As Long
Private Declare PtrSafe Function LCMapString Lib "kernel32" Alias "LCMapStringA" (ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long
Private Declare PtrSafe Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare PtrSafe Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SystemTime, ByVal lpFormat As String, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long
Private Declare PtrSafe Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SystemTime, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
Private Declare PtrSafe Function SetThreadLocale Lib "kernel32" (ByVal Locale As Long) As Long
Private Declare PtrSafe Function GetSystemDefaultLangID Lib "kernel32" () As Integer
Private Declare PtrSafe Function GetUserDefaultLangID Lib "kernel32" () As Integer
Private Declare PtrSafe Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare PtrSafe Function GetUserDefaultLCID Lib "kernel32" () As Long

' Locale Independent APIs

Private Declare PtrSafe Function GetStringTypeA Lib "kernel32" (ByVal lcid As Long, ByVal dwInfoType As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, lpCharType As Long) As Long
Private Declare PtrSafe Function FoldString Lib "kernel32" Alias "FoldStringA" (ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long

' * winnls.h -- NLS procedure declarations
Private Declare PtrSafe Function EnumTimeFormats Lib "kernel32" (ByVal lpTimeFmtEnumProc As LongPtr, ByVal Locale As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function EnumDateFormats Lib "kernel32" (ByVal lpDateFmtEnumProc As LongPtr, ByVal Locale As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function IsValidLocale Lib "kernel32" (ByVal Locale As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function ConvertDefaultLocale Lib "kernel32" (ByVal Locale As Long) As Long
Private Declare PtrSafe Function GetThreadLocale Lib "kernel32" () As Long
Private Declare PtrSafe Function EnumSystemLocales Lib "kernel32" (ByVal lpLocaleEnumProc As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function EnumSystemCodePages Lib "kernel32" (ByVal lpCodePageEnumProc As LongPtr, ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function ReadConsoleOutput Lib "kernel32" Alias "ReadConsoleOutputA" (ByVal hConsoleOutput As LongPtr, lpBuffer As CHAR_INFO, dwBufferSize As COORD, dwBufferCoord As COORD, lpReadRegion As SMALL_RECT) As Long
Private Declare PtrSafe Function WriteConsoleOutput Lib "kernel32" Alias "WriteConsoleOutputA" (ByVal hConsoleOutput As LongPtr, lpBuffer As CHAR_INFO, dwBufferSize As COORD, dwBufferCoord As COORD, lpWriteRegion As SMALL_RECT) As Long
Private Declare PtrSafe Function ReadConsoleOutputCharacter Lib "kernel32" Alias "ReadConsoleOutputCharacterA" (ByVal hConsoleOutput As LongPtr, ByVal lpCharacter As String, ByVal nLength As Long, dwReadCoord As COORD, lpNumberOfCharsRead As Long) As Long
Private Declare PtrSafe Function ReadConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As LongPtr, lpAttribute As Long, ByVal nLength As Long, dwReadCoord As COORD, lpNumberOfAttrsRead As Long) As Long
Private Declare PtrSafe Function WriteConsoleOutputCharacter Lib "kernel32" Alias "WriteConsoleOutputCharacterA" (ByVal hConsoleOutput As LongPtr, ByVal lpCharacter As String, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfCharsWritten As Long) As Long

Private Declare PtrSafe Function WriteConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As LongPtr, lpAttribute As Integer, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfAttrsWritten As Long) As Long
Private Declare PtrSafe Function FillConsoleOutputCharacter Lib "kernel32" Alias "FillConsoleOutputCharacterA" (ByVal hConsoleOutput As LongPtr, ByVal cCharacter As Byte, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfCharsWritten As Long) As Long
Private Declare PtrSafe Function FillConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As LongPtr, ByVal wAttribute As Long, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfAttrsWritten As Long) As Long
Private Declare PtrSafe Function GetConsoleMode Lib "kernel32" (ByVal hConsoleHandle As LongPtr, lpMode As Long) As Long
Private Declare PtrSafe Function GetNumberOfConsoleInputEvents Lib "kernel32" (ByVal hConsoleInput As LongPtr, lpNumberOfEvents As Long) As Long
Private Declare PtrSafe Function GetConsoleScreenBufferInfo Lib "kernel32" (ByVal hConsoleOutput As LongPtr, lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
Private Declare PtrSafe Function GetLargestConsoleWindowSize Lib "kernel32" (ByVal hConsoleOutput As LongPtr) As COORD
Private Declare PtrSafe Function GetConsoleCursorInfo Lib "kernel32" (ByVal hConsoleOutput As LongPtr, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
Private Declare PtrSafe Function GetNumberOfConsoleMouseButtons Lib "kernel32" (lpNumberOfMouseButtons As Long) As Long
Private Declare PtrSafe Function SetConsoleMode Lib "kernel32" (ByVal hConsoleHandle As LongPtr, ByVal dwMode As Long) As Long
Private Declare PtrSafe Function SetConsoleActiveScreenBuffer Lib "kernel32" (ByVal hConsoleOutput As LongPtr) As Long
Private Declare PtrSafe Function FlushConsoleInputBuffer Lib "kernel32" (ByVal hConsoleInput As LongPtr) As Long
Private Declare PtrSafe Function SetConsoleScreenBufferSize Lib "kernel32" (ByVal hConsoleOutput As LongPtr, dwSize As COORD) As Long
Private Declare PtrSafe Function SetConsoleCursorPosition Lib "kernel32" (ByVal hConsoleOutput As LongPtr, dwCursorPosition As COORD) As Long
Private Declare PtrSafe Function SetConsoleCursorInfo Lib "kernel32" (ByVal hConsoleOutput As LongPtr, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
Private Declare PtrSafe Function ScrollConsoleScreenBuffer Lib "kernel32" Alias "ScrollConsoleScreenBufferA" (ByVal hConsoleOutput As LongPtr, lpScrollRectangle As SMALL_RECT, lpClipRectangle As SMALL_RECT, dwDestinationOrigin As COORD, lpFill As CHAR_INFO) As Long
Private Declare PtrSafe Function SetConsoleWindowInfo Lib "kernel32" (ByVal hConsoleOutput As LongPtr, ByVal bAbsolute As Long, lpConsoleWindow As SMALL_RECT) As Long
Private Declare PtrSafe Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As LongPtr, ByVal wAttributes As Long) As Long
Private Declare PtrSafe Function SetConsoleCtrlHandler Lib "kernel32" (ByVal HandlerRoutine As LongPtr, ByVal Add As Long) As Long
Private Declare PtrSafe Function GenerateConsoleCtrlEvent Lib "kernel32" (ByVal dwCtrlEvent As Long, ByVal dwProcessGroupId As Long) As Long
Private Declare PtrSafe Function AllocConsole Lib "kernel32" () As Long
Private Declare PtrSafe Function FreeConsole Lib "kernel32" () As Long
Private Declare PtrSafe Function GetConsoleTitle Lib "kernel32" Alias "GetConsoleTitleA" (ByVal lpConsoleTitle As String, ByVal nSize As Long) As Long
Private Declare PtrSafe Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Private Declare PtrSafe Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As LongPtr, lpBuffer As Any, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any) As Long
Private Declare PtrSafe Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As LongPtr, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, ByVal lpReserved As LongPtr) As Long

Private Declare PtrSafe Function CreateConsoleScreenBuffer Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwFlags As Long, lpScreenBufferData As Any) As LongPtr
Private Declare PtrSafe Function GetConsoleCP Lib "kernel32" () As Long
Private Declare PtrSafe Function SetConsoleCP Lib "kernel32" (ByVal wCodePageID As Long) As Long
Private Declare PtrSafe Function GetConsoleOutputCP Lib "kernel32" () As Long
Private Declare PtrSafe Function SetConsoleOutputCP Lib "kernel32" (ByVal wCodePageID As Long) As Long

'******************************************************************************
'* [概  要] WinApi_Sleep
'* [詳  細] WindowsAPI「Sleep」ラッパーメソッド。
'*
'* @param ms 待機時間（ミリ秒）
'*
'******************************************************************************
Friend Sub WinApi_Sleep(ByVal ms As Long)
    Call Sleep(ms)
End Sub

'******************************************************************************
'* [概  要] WinApi_CopyMemory
'* [詳  細] WindowsAPI「 CopyMemory 」ラッパーメソッド。
'*          Lib "kernel32"
'*
'******************************************************************************
Friend Sub WinApi_CopyMemory(ByVal Destination As LongPtr, ByVal source As LongPtr, ByVal Length As LongPtr)
    Call CopyMemory(Destination, source, Length)
End Sub

'******************************************************************************
'* [概  要] WinApi_CopyMemoryByRef
'* [詳  細] WindowsAPI「 CopyMemory 」ラッパーメソッド。
'*          Lib "kernel32"
'*
'******************************************************************************
Friend Sub WinApi_CopyMemoryByRef(ByRef Destination As LongPtr, ByRef source As LongPtr, ByVal Length As LongPtr, _
                                Optional destByVal As Boolean = False, Optional sourceByVal As Boolean = False)
    If destByVal And sourceByVal Then
        Call CopyMemoryByRef(ByVal Destination, ByVal source, Length)
    ElseIf destByVal Then
        Call CopyMemoryByRef(ByVal Destination, source, Length)
    ElseIf sourceByVal Then
        Call CopyMemoryByRef(Destination, ByVal source, Length)
    Else
        Call CopyMemoryByRef(Destination, source, Length)
    End If
End Sub

'******************************************************************************
'* [概  要] WinApi_GetPrivateProfileString
'* [詳  細] WindowsAPI「 GetPrivateProfileString 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_GetPrivateProfileString(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    WinApi_GetPrivateProfileString = GetPrivateProfileString(lpApplicationName, lpKeyName, lpDefault, lpReturnedString, nSize, lpFileName)
End Function

'******************************************************************************
'* [概  要] WinApi_CreateFile
'* [詳  細] WindowsAPI「 CreateFile 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_CreateFile(ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
    WinApi_CreateFile = CreateFile(lpFileName, dwDesiredAccess, dwShareMode, lpSecurityAttributes, dwCreationDisposition, dwFlagsAndAttributes, hTemplateFile)
End Function

'******************************************************************************
'* [概  要] WinApi_CloseHandle
'* [詳  細] WindowsAPI「 CloseHandle 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_CloseHandle(ByVal hObject As LongPtr) As Long
    WinApi_CloseHandle = CloseHandle(hObject)
End Function

'******************************************************************************
'* [概  要] WinApi_LocalFileTimeToFileTime
'* [詳  細] WindowsAPI「 LocalFileTimeToFileTime 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_LocalFileTimeToFileTime(lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
    WinApi_LocalFileTimeToFileTime = LocalFileTimeToFileTime(lpLocalFileTime, lpFileTime)
End Function

'******************************************************************************
'* [概  要] WinApi_SystemTimeToFileTime
'* [詳  細] WindowsAPI「 SystemTimeToFileTime 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_SystemTimeToFileTime(lpSystemTime As SystemTime, lpFileTime As FILETIME) As Long
    WinApi_SystemTimeToFileTime = SystemTimeToFileTime(lpSystemTime, lpFileTime)
End Function

'******************************************************************************
'* [概  要] WinApi_SetFileTime
'* [詳  細] WindowsAPI「 SetFileTime 」ラッパーメソッド。
'*          Lib "kernel32"
'*
'******************************************************************************
Friend Function WinApi_SetFileTime(ByVal hFile As LongPtr, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
    WinApi_SetFileTime = SetFileTime(hFile, lpCreationTime, lpLastAccessTime, lpLastWriteTime)
End Function

'******************************************************************************
'* [概  要] WinApi_GetFileTime
'* [詳  細] WindowsAPI「 GetFileTime 」ラッパーメソッド。
'*          Lib "kernel32"
'*
'******************************************************************************
Friend Function WinApi_GetFileTime(ByVal hFile As LongPtr, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
    WinApi_GetFileTime = GetFileTime(hFile, lpCreationTime, lpLastAccessTime, lpLastWriteTime)
End Function

