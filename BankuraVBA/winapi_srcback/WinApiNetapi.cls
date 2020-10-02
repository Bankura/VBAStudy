VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiNetapi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WindowAPI Netapi クラス
'* [詳  細] WindowAPIのNetapi32.dllの機能を提供する。
'*
'* [参  考] Win32API_PtrSafe.TXT
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
'  Windows Network support
Private Declare PtrSafe Function NetUserChangePassword Lib "netapi32.dll" (Domain As Any, User As Any, OldPass As Byte, NewPass As Byte) As Long
Private Declare PtrSafe Function NetUserGetInfo Lib "netapi32.dll" (lpServer As Any, UserName As Byte, ByVal Level As Long, lpBuffer As LongPtr) As Long
Private Declare PtrSafe Function NetUserGetGroups Lib "netapi32.dll" (lpServer As Any, UserName As Byte, ByVal Level As Long, lpBuffer As LongPtr, ByVal PrefMaxLen As Long, lpEntriesRead As Long, lpTotalEntries As Long) As Long
Private Declare PtrSafe Function NetUserGetLocalGroups Lib "netapi32.dll" (lpServer As Any, UserName As Byte, ByVal Level As Long, ByVal Flags As Long, lpBuffer As LongPtr, ByVal MaxLen As Long, lpEntriesRead As Long, lpTotalEntries As Long) As Long
Private Declare PtrSafe Function NetUserAdd Lib "netapi32" (lpServer As Any, ByVal Level As Long, lpUser As USER_INFO_3, lpError As Long) As Long
Private Declare PtrSafe Function NetWkstaGetInfo Lib "netapi32.dll" (lpServer As Any, ByVal Level As Long, lpBuffer As Any) As Long
Private Declare PtrSafe Function NetWkstaUserGetInfo Lib "netapi32.dll" (ByVal Reserved As Any, ByVal Level As Long, lpBuffer As Any) As Long
Private Declare PtrSafe Function NetApiBufferFree Lib "netapi32.dll" (ByVal lpBuffer As LongPtr) As Long
Private Declare PtrSafe Function NetRemoteTOD Lib "netapi32.dll" (yServer As Any, pBuffer As LongPtr) As Long
Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare PtrSafe Function GetUserNameW Lib "advapi32.dll" (lpBuffer As Byte, nSize As Long) As Long
Private Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare PtrSafe Function GetComputerNameW Lib "kernel32" (lpBuffer As Any, nSize As Long) As Long
Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr
Private Declare PtrSafe Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As LongPtr, ByVal DesiredAccess As Long, TokenHandle As LongPtr) As Long
Private Declare PtrSafe Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As LongPtr, ByVal TokenInformationClass As Long, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare PtrSafe Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidW" (ByVal lpSystemName As Any, Sid As Any, Name As Any, cbName As Long, ReferencedDomainName As Any, cbReferencedDomainName As Long, peUse As Integer) As Long
Private Declare PtrSafe Function NetLocalGroupDelMembers Lib "netapi32.dll" (ByVal psServer As LongPtr, ByVal psLocalGroup As LongPtr, ByVal lLevel As Long, uMember As LOCALGROUP_MEMBERS_INFO_0, ByVal lMemberCount As Long) As Long
Private Declare PtrSafe Function NetLocalGroupGetMembers Lib "netapi32.dll" (ByVal psServer As LongPtr, ByVal psLocalGroup As LongPtr, ByVal lLevel As Long, pBuffer As LongPtr, ByVal lMaxLength As Long, plEntriesRead As Long, plTotalEntries As Long, phResume As LongPtr) As Long
Private Declare PtrSafe Function Netbios Lib "netapi32.dll" (pncb As NCB) As Byte
