VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiOle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WindowAPI ole クラス
'* [詳  細] WindowAPIの「ole32.dll」の機能を提供する。
'*
'* [参  考]
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Function IIDFromString Lib "ole32.dll" (lpsz As Any, lpiid As Any) As Long

'******************************************************************************
'* [概  要] WinApi_IIDFromString
'* [詳  細] WindowsAPI「 IIDFromString 」ラッパーメソッド。
'*          Lib "ole32"
'*
'******************************************************************************
Friend Function WinApi_IIDFromString(lpsz As Byte, lpiid As LongPtr) As Long
    WinApi_IIDFromString = IIDFromString(lpsz, lpiid)
End Function

