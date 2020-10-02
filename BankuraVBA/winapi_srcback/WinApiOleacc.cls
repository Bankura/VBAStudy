VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiOleacc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WindowAPI oleacc クラス
'* [詳  細] WindowAPIの「oleacc.dll」の機能を提供する。
'*
'* [参  考]
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Function ObjectFromLresult Lib "oleacc" (ByVal lResult As LongPtr, riid As Any, ByVal wParam As LongPtr, ppvObject As Any) As LongPtr


'******************************************************************************
'* [概  要] WinApi_ObjectFromLresult
'* [詳  細] WindowsAPI「 ObjectFromLresult 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_ObjectFromLresult(ByVal lResult As LongPtr, riid As LongPtr, ByVal wParam As LongPtr, ppvObject As Object) As LongPtr
    WinApi_ObjectFromLresult = ObjectFromLresult(lResult, riid, wParam, ppvObject)
End Function
