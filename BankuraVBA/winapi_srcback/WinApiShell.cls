VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WindowAPI shell クラス
'* [詳  細] WindowAPIのshell32.dllの機能を提供する。
'*
'* [参  考] Win32API_PtrSafe.TXT
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
' shellapi.h -  SHELL.DLL functions
Private Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As LongPtr, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Private Declare PtrSafe Function DragQueryPoint Lib "shell32.dll" (ByVal hDrop As LongPtr, lpPoint As POINTAPI) As Long
Private Declare PtrSafe Sub DragFinish Lib "shell32.dll" (ByVal hDrop As LongPtr)
Private Declare PtrSafe Sub DragAcceptFiles Lib "shell32.dll" (ByVal hWnd As LongPtr, ByVal fAccept As Long)
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
Private Declare PtrSafe Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As LongPtr
Private Declare PtrSafe Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As LongPtr, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As LongPtr) As Long
Private Declare PtrSafe Function DuplicateIcon Lib "shell32.dll" (ByVal hInst As LongPtr, ByVal hIcon As LongPtr) As LongPtr
Private Declare PtrSafe Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As LongPtr, ByVal lpIconPath As String, lpiIcon As Long) As LongPtr
Private Declare PtrSafe Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As LongPtr, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As LongPtr
Private Declare PtrSafe Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As LongPtr
Private Declare PtrSafe Function DoEnvironmentSubst Lib "shell32.dll" Alias "DoEnvironmentSubstA" (ByVal szString As String, ByVal cbString As Long) As Long
Private Declare PtrSafe Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As LongPtr, phiconSmall As LongPtr, ByVal nIcons As Long) As Long
Private Declare PtrSafe Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare PtrSafe Sub SHFreeNameMappings Lib "shell32.dll" (ByVal hNameMappings As LongPtr)
Private Declare PtrSafe Function Shell_NotifyIcon Lib "shell32.dll" Alias " Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare PtrSafe Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As LongPtr
Private Declare PtrSafe Function SHGetNewLinkInfo Lib "shell32.dll" Alias "SHGetNewLinkInfoA" (ByVal pszLinkto As String, ByVal pszDir As String, ByVal pszName As String, pfMustCopy As Long, ByVal uFlags As Long) As Long

Private Declare PtrSafe Function CommandLineToArgv Lib "shell32" Alias "CommandLineToArgvW" (ByVal lpCmdLine As String, pNumArgs As Long) As LongPtr

Private Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" (ByVal hWnd As LongPtr, ByVal pszPath As String, ByVal psa As Any) As LongPtr

'******************************************************************************
'* [概  要] WinApi_SHCreateDirectoryEx
'* [詳  細] WindowsAPI「 SHCreateDirectoryEx 」ラッパーメソッド。
'*          Lib "shell32"
'*
'******************************************************************************
Friend Function WinApi_SHCreateDirectoryEx(ByVal hWnd As LongPtr, ByVal pszPath As String, ByVal psa As Long) As LongPtr
    WinApi_SHCreateDirectoryEx = SHCreateDirectoryEx(hWnd, pszPath, psa)
End Function
