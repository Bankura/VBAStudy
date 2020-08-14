VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiWininet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WindowAPI wininet クラス
'* [詳  細] WindowAPIの「wininet.dll」の機能を提供する。
'*
'* [参  考]
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long



'******************************************************************************
'* [概  要] WinApi_DeleteUrlCacheEntry
'* [詳  細] WindowsAPI「 DeleteUrlCacheEntry 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_DeleteUrlCacheEntry(ByVal lpszUrlName As String) As Long
    WinApi_DeleteUrlCacheEntry = DeleteUrlCacheEntry(lpszUrlName)
End Function

