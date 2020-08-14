VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiKernelWinbase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WindowAPI kernel WinBase クラス
'* [詳  細] WindowAPIのkernel32.dll内winbase.h定義の機能を提供する。
'*
'* [参  考] Win32API_PtrSafe.TXT
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
' winbase.h -- This module defines the 32-Bit Windows Base APIs
'  Define API decoration for direct importing of DLL references.
Private Declare PtrSafe Function HeapValidate Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare PtrSafe Function HeapCompact Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long) As LongPtr

Private Declare PtrSafe Function HeapLock Lib "kernel32" (ByVal hHeap As LongPtr) As Long
Private Declare PtrSafe Function HeapUnlock Lib "kernel32" (ByVal hHeap As LongPtr) As Long

Private Declare PtrSafe Function GetBinaryType Lib "kernel32" Alias "GetBinaryTypeA" (ByVal lpApplicationName As String, lpBinaryType As Long) As Long

Private Declare PtrSafe Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare PtrSafe Function GetProcessAffinityMask Lib "kernel32" (ByVal hProcess As LongPtr, lpProcessAffinityMask As LongPtr, SystemAffinityMask As LongPtr) As Long

Private Declare PtrSafe Function LogonUser Lib "kernel32" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As LongPtr) As Long

Private Declare PtrSafe Function ImpersonateLoggedOnUser Lib "kernel32" (ByVal hToken As LongPtr) As Long
Private Declare PtrSafe Function CreateProcessAsUser Lib "advapi32.dll" Alias "CreateProcessAsUserA" (ByVal hToken As LongPtr, ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As String, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare PtrSafe Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare PtrSafe Function GetSystemPowerStatus Lib "kernel32" (lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long
Private Declare PtrSafe Function SetSystemPowerState Lib "kernel32" (ByVal fSuspend As Long, ByVal fForce As Long) As Long



' winbaseかは定かではない
Private Declare PtrSafe Function CommConfigDialog Lib "kernel32" Alias "CommConfigDialogA" (ByVal lpszName As String, ByVal hWnd As LongPtr, lpCC As COMMCONFIG) As Long
Private Declare PtrSafe Function CreateIoCompletionPort Lib "kernel32" (ByVal FileHandle As LongPtr, ByVal ExistingCompletionPort As LongPtr, ByVal CompletionKey As LongPtr, ByVal NumberOfConcurrentThreads As Long) As LongPtr
Private Declare PtrSafe Function DisableThreadLibraryCalls Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function EnumResourceLanguages Lib "kernel32" Alias "EnumResourceLanguagesA" (ByVal hModule As LongPtr, ByVal lpType As String, ByVal lpName As String, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function EnumResourceNames Lib "kernel32" Alias "EnumResourceNamesA" (ByVal hModule As LongPtr, ByVal lpType As String, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function EnumResourceTypes Lib "kernel32" Alias "EnumResourceTypesA" (ByVal hModule As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function FreeEnvironmentStrings Lib "kernel32" Alias "FreeEnvironmentStringsA" (ByVal lpsz As String) As Long
Private Declare PtrSafe Sub FreeLibraryAndExitThread Lib "kernel32" (ByVal hLibModule As LongPtr, ByVal dwExitCode As Long)
Private Declare PtrSafe Function FreeResource Lib "kernel32" (ByVal hResData As LongPtr) As Long
Private Declare PtrSafe Function GetCommConfig Lib "kernel32" (ByVal hCommDev As LongPtr, lpCC As COMMCONFIG, lpdwSize As Long) As Long
Private Declare PtrSafe Function GetCompressedFileSize Lib "kernel32" Alias "GetCompressedFileSizeA" (ByVal lpFileName As String, lpFileSizeHigh As Long) As Long
Private Declare PtrSafe Function GetDefaultCommConfig Lib "kernel32" Alias "GetDefaultCommConfigA" (ByVal lpszName As String, lpCC As COMMCONFIG, lpdwSize As Long) As Long
Private Declare PtrSafe Function GetHandleInformation Lib "kernel32" (ByVal hObject As LongPtr, lpdwFlags As Long) As Long
Private Declare PtrSafe Function GetProcessHeaps Lib "kernel32" (ByVal NumberOfHeaps As Long, ProcessHeaps As LongPtr) As Long
Private Declare PtrSafe Function GetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As LongPtr, lpMinimumWorkingSetSize As LongPtr, lpMaximumWorkingSetSize As LongPtr) As Long
Private Declare PtrSafe Function GetQueuedCompletionStatus Lib "kernel32" (ByVal CompletionPort As LongPtr, lpNumberOfBytesTransferred As Long, lpCompletionKey As LongPtr, lpOverlapped As LongPtr, ByVal dwMilliseconds As Long) As Long
Private Declare PtrSafe Function GetSystemTimeAdjustment Lib "kernel32" (lpTimeAdjustment As Long, lpTimeIncrement As Long, lpTimeAdjustmentDisabled As Long) As Long
Private Declare PtrSafe Function GlobalCompact Lib "kernel32" (ByVal dwMinFree As Long) As LongPtr
Private Declare PtrSafe Sub GlobalFix Lib "kernel32" (ByVal hMem As LongPtr)
Private Declare PtrSafe Sub GlobalUnfix Lib "kernel32" (ByVal hMem As LongPtr)
Private Declare PtrSafe Function GlobalWire Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnWire Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As LongPtr) As Long
Private Declare PtrSafe Function LocalCompact Lib "kernel32" (ByVal uMinFree As Long) As LongPtr
Private Declare PtrSafe Function LocalShrink Lib "kernel32" (ByVal hMem As LongPtr, ByVal cbNewSize As Long) As LongPtr
Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As LongPtr)
Private Declare PtrSafe Function ReadFileEx Lib "kernel32" (ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As LongPtr) As Long
Private Declare PtrSafe Function SetCommConfig Lib "kernel32" (ByVal hCommDev As LongPtr, lpCC As COMMCONFIG, ByVal dwSize As Long) As Long
Private Declare PtrSafe Function SetDefaultCommConfig Lib "kernel32" Alias "SetDefaultCommConfigA" (ByVal lpszName As String, lpCC As COMMCONFIG, ByVal dwSize As Long) As Long
Private Declare PtrSafe Sub SetFileApisToANSI Lib "kernel32" ()
Private Declare PtrSafe Function SetHandleInformation Lib "kernel32" (ByVal hObject As LongPtr, ByVal dwMask As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As LongPtr, ByVal dwMinimumWorkingSetSize As LongPtr, ByVal dwMaximumWorkingSetSize As LongPtr) As Long
Private Declare PtrSafe Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As LongPtr
Private Declare PtrSafe Function lstrcpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As String, ByVal lpString2 As String, ByVal iMaxLength As Long) As LongPtr
Private Declare PtrSafe Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As LongPtr
Private Declare PtrSafe Function SetSystemTimeAdjustment Lib "kernel32" (ByVal dwTimeAdjustment As Long, ByVal bTimeAdjustmentDisabled As Long) As Long
Private Declare PtrSafe Function SetThreadAffinityMask Lib "kernel32" (ByVal hThread As LongPtr, ByVal dwThreadAffinityMask As LongPtr) As LongPtr
Private Declare PtrSafe Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As LongPtr) As LongPtr
Private Declare PtrSafe Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpUniversalTime As SystemTime, lpLocalTime As SystemTime) As Long
Private Declare PtrSafe Function WriteFileEx Lib "kernel32" (ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As LongPtr) As Long
Private Declare PtrSafe Function EnumCalendarInfo Lib "kernel32" Alias "EnumCalendarInfoA" (ByVal lpCalInfoEnumProc As LongPtr, ByVal Locale As Long, ByVal Calendar As Long, ByVal CalType As Long) As Long
Private Declare PtrSafe Function GetCurrencyFormat Lib "kernel32" Alias "GetCurrencyFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, ByVal lpValue As String, lpFormat As CURRENCYFMT, ByVal lpCurrencyStr As String, ByVal cchCurrency As Long) As Long
Private Declare PtrSafe Function GetNumberFormat Lib "kernel32" Alias "GetNumberFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, ByVal lpValue As String, lpFormat As NUMBERFMT, ByVal lpNumberStr As String, ByVal cchNumber As Long) As Long
Private Declare PtrSafe Function GetStringTypeEx Lib "kernel32" Alias "GetStringTypeExA" (ByVal Locale As Long, ByVal dwInfoType As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, lpCharType As Integer) As Long
Private Declare PtrSafe Function GetStringTypeW Lib "kernel32" (ByVal dwInfoType As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, lpCharType As Integer) As Long
Private Declare PtrSafe Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long


