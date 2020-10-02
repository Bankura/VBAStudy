VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiLz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WindowAPI lz クラス
'* [詳  細] WindowAPIのlz32.dllの機能を提供する。
'*
'* [参  考] Win32API_PtrSafe.TXT
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Function LZCopy Lib "lz32.dll" (ByVal hfSource As Long, ByVal hfDest As Long) As Long
Private Declare PtrSafe Function LZInit Lib "lz32.dll" (ByVal hfSrc As Long) As Long
Private Declare PtrSafe Function GetExpandedName Lib "lz32.dll" Alias "GetExpandedNameA" (ByVal lpszSource As String, ByVal lpszBuffer As String) As Long
Private Declare PtrSafe Function LZOpenFile Lib "lz32.dll" Alias "LZOpenFileA" (ByVal lpszFile As String, lpOf As OFSTRUCT, ByVal style As Long) As Long
Private Declare PtrSafe Function LZSeek Lib "lz32.dll" (ByVal hfFile As Long, ByVal lOffset As Long, ByVal nOrigin As Long) As Long
Private Declare PtrSafe Function LZRead Lib "lz32.dll" (ByVal hfFile As Long, ByVal lpvBuf As String, ByVal cbread As Long) As Long
Private Declare PtrSafe Sub LZClose Lib "lz32.dll" (ByVal hfFile As Long)

Private Declare PtrSafe Function CopyLZFile Lib "lz32" (ByVal n1 As Long, ByVal n2 As Long) As Long
Private Declare PtrSafe Function LZStart Lib "lz32" () As Long
Private Declare PtrSafe Sub LZDone Lib "lz32" ()
