VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiAdvapi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WindowAPI advapi クラス
'* [詳  細] WindowAPIのadvapi32.dllの機能を提供する。
'*
'* [参  考] Win32API_PtrSafe.TXT
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
' Registry API prototypes
Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As LongPtr) As Long
Private Declare PtrSafe Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As LongPtr, phkResult As LongPtr) As Long
Private Declare PtrSafe Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, phkResult As LongPtr) As Long
Private Declare PtrSafe Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As LongPtr, lpdwDisposition As Long) As Long
Private Declare PtrSafe Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As LongPtr, ByVal lpSubKey As String) As Long
Private Declare PtrSafe Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As LongPtr, ByVal lpValueName As String) As Long
Private Declare PtrSafe Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As LongPtr, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare PtrSafe Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As LongPtr, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As LongPtr, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare PtrSafe Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As LongPtr, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As LongPtr, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare PtrSafe Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As LongPtr) As Long
Private Declare PtrSafe Function RegGetKeySecurity Lib "advapi32.dll" (ByVal hKey As LongPtr, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, lpcbSecurityDescriptor As Long) As Long
Private Declare PtrSafe Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal lpFile As String) As Long
Private Declare PtrSafe Function RegNotifyChangeKeyValue Lib "advapi32.dll" (ByVal hKey As LongPtr, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, ByVal hEvent As LongPtr, ByVal fAsynchronus As Long) As Long
Private Declare PtrSafe Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, phkResult As LongPtr) As Long
Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As LongPtr) As Long
Private Declare PtrSafe Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As LongPtr, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As LongPtr, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare PtrSafe Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Private Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As LongPtr, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare PtrSafe Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
Private Declare PtrSafe Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As LongPtr, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As LongPtr, ByVal lpFile As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare PtrSafe Function RegSetKeySecurity Lib "advapi32.dll" (ByVal hKey As LongPtr, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Private Declare PtrSafe Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare PtrSafe Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare PtrSafe Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKey As LongPtr, ByVal lpSubKey As String) As Long
Private Declare PtrSafe Function InitiateSystemShutdown Lib "advapi32.dll" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long
Private Declare PtrSafe Function AbortSystemShutdown Lib "advapi32.dll" Alias "AbortSystemShutdownA" (ByVal lpMachineName As String) As Long

'  API Function Prototypes
Private Declare PtrSafe Function ChangeServiceConfig Lib "advapi32.dll" Alias "ChangeServiceConfigA" (ByVal hService As LongPtr, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, lpdwTagId As Long, ByVal lpDependencies As String, ByVal lpServiceStartName As String, ByVal lpPassword As String, ByVal lpDisplayName As String) As Long
Private Declare PtrSafe Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As LongPtr) As Long
Private Declare PtrSafe Function ControlService Lib "advapi32.dll" (ByVal hService As LongPtr, ByVal dwControl As Long, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare PtrSafe Function CreateService Lib "advapi32.dll" Alias "CreateServiceA" (ByVal hSCManager As LongPtr, ByVal lpServiceName As String, ByVal lpDisplayName As String, ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, lpdwTagId As Long, ByVal lpDependencies As String, ByVal lp As String, ByVal lpPassword As String) As LongPtr
Private Declare PtrSafe Function DeleteService Lib "advapi32.dll" (ByVal hService As LongPtr) As Long
Private Declare PtrSafe Function EnumDependentServices Lib "advapi32.dll" Alias "EnumDependentServicesA" (ByVal hService As LongPtr, ByVal dwServiceState As Long, lpServices As ENUM_SERVICE_STATUS, ByVal cbBufSize As Long, pcbBytesNeeded As Long, lpServicesReturned As Long) As Long
Private Declare PtrSafe Function EnumServicesStatus Lib "advapi32.dll" Alias "EnumServicesStatusA" (ByVal hSCManager As LongPtr, ByVal dwServiceType As Long, ByVal dwServiceState As Long, lpServices As ENUM_SERVICE_STATUS, ByVal cbBufSize As Long, pcbBytesNeeded As Long, lpServicesReturned As Long, lpResumeHandle As Long) As Long
Private Declare PtrSafe Function GetServiceKeyName Lib "advapi32.dll" Alias "GetServiceKeyNameA" (ByVal hSCManager As LongPtr, ByVal lpDisplayName As String, ByVal lpServiceName As String, lpcchBuffer As Long) As Long
Private Declare PtrSafe Function GetServiceDisplayName Lib "advapi32.dll" Alias "GetServiceDisplayNameA" (ByVal hSCManager As LongPtr, ByVal lpServiceName As String, ByVal lpDisplayName As String, lpcchBuffer As Long) As Long
Private Declare PtrSafe Function LockServiceDatabase Lib "advapi32.dll" (ByVal hSCManager As LongPtr) As LongPtr
Private Declare PtrSafe Function NotifyBootConfigStatus Lib "advapi32.dll" (ByVal BootAcceptable As Long) As Long
Private Declare PtrSafe Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As LongPtr
Private Declare PtrSafe Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As LongPtr, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As LongPtr
Private Declare PtrSafe Function QueryServiceConfig Lib "advapi32.dll" Alias "QueryServiceConfigA" (ByVal hService As LongPtr, lpServiceConfig As QUERY_SERVICE_CONFIG, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Private Declare PtrSafe Function QueryServiceLockStatus Lib "advapi32.dll" Alias "QueryServiceLockStatusA" (ByVal hSCManager As LongPtr, lpLockStatus As QUERY_SERVICE_LOCK_STATUS, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Private Declare PtrSafe Function QueryServiceObjectSecurity Lib "advapi32.dll" (ByVal hService As LongPtr, ByVal dwSecurityInformation As Long, lpSecurityDescriptor As Any, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Private Declare PtrSafe Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As LongPtr, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare PtrSafe Function RegisterServiceCtrlHandler Lib "advapi32.dll" Alias "RegisterServiceCtrlHandlerA" (ByVal lpServiceName As String, ByVal lpHandlerProc As LongPtr) As LongPtr
Private Declare PtrSafe Function SetServiceObjectSecurity Lib "advapi32.dll" (ByVal hService As LongPtr, ByVal dwSecurityInformation As Long, lpSecurityDescriptor As Any) As Long
Private Declare PtrSafe Function SetServiceStatus Lib "advapi32.dll" (ByVal hServiceStatus As LongPtr, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare PtrSafe Function StartServiceCtrlDispatcher Lib "advapi32.dll" Alias "StartServiceCtrlDispatcherA" (lpServiceStartTable As SERVICE_TABLE_ENTRY) As Long
Private Declare PtrSafe Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As LongPtr, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As LongPtr) As Long
Private Declare PtrSafe Function UnlockServiceDatabase Lib "advapi32.dll" (ScLock As Any) As Long


Private Declare PtrSafe Function SetServiceBits Lib "advapi32" (ByVal hServiceStatus As LongPtr, ByVal dwServiceBits As Long, ByVal bSetBitsOn As Long, ByVal bUpdateImmediately As Long) As Long
Private Declare PtrSafe Function IsTextUnicode Lib "advapi32" (lpBuffer As Any, ByVal cb As Long, lpi As Long) As Long
Private Declare PtrSafe Function NotifyChangeEventLog Lib "advapi32" (ByVal hEventLog As LongPtr, ByVal hEvent As LongPtr) As Long
Private Declare PtrSafe Function SetThreadToken Lib "advapi32" (Thread As LongPtr, ByVal Token As LongPtr) As Long
