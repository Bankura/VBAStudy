VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiOleaut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WindowAPI oleaut クラス
'* [詳  細] WindowAPIの「oleaut.dll」の機能を提供する。
'*
'* [参  考]
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc_ As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByRef prgvt As Integer, ByRef prgpvarg As LongPtr, ByRef pvargResult As Variant) As Long
Private Declare PtrSafe Function SafeArrayAllocDescriptor Lib "oleaut32" (ByVal cDims As Long, ByRef ppsaOut() As Any) As Long

'******************************************************************************
'* [概  要] WinApi_DispCallFunc
'* [詳  細] WindowsAPI「 DispCallFunc 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_DispCallFunc(ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc_ As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByRef prgvt As Integer, ByRef prgpvarg As LongPtr, ByRef pvargResult As Variant) As Long
    WinApi_DispCallFunc = DispCallFunc(pvInstance, oVft, cc_, vtReturn, cActuals, prgvt, prgpvarg, pvargResult)
End Function

'******************************************************************************
'* [概  要] WinApi_SafeArrayAllocDescriptor
'* [詳  細] WindowsAPI「 SafeArrayAllocDescriptor 」ラッパーメソッド。
'*
'******************************************************************************
Friend Function WinApi_SafeArrayAllocDescriptor(ByVal cDims As Long, ByRef ppsaOut() As Variant) As Long
    WinApi_SafeArrayAllocDescriptor = SafeArrayAllocDescriptor(cDims, ppsaOut)
End Function


