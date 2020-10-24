VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] CommandResultクラス
'* [詳  細] コマンドのResult情報データクラス。
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mExitCd As String
Private mStdOutList As Variant
Private mStdErrTxt As String
Private mRowMax As Long
Private mColMax As Long
    
'******************************************************************************
'* プロパティ定義
'******************************************************************************
'Status Code
Public Property Get ExitCd() As String
    ExitCd = mExitCd
End Property
Public Property Let ExitCd(arg As String)
    mExitCd = arg
End Property

'StdOutList
Public Property Get StdOutList() As Variant
    StdOutList = mStdOutList
End Property
Public Property Let StdOutList(arg As Variant)
    mStdOutList = arg
End Property

'StdError Text
Public Property Get StdErrTxt() As String
    StdErrTxt = mStdErrTxt
End Property
Public Property Let StdErrTxt(arg As String)
    mStdErrTxt = arg
End Property

'RowMax
Public Property Get RowMax() As Long
    RowMax = mRowMax
End Property
Public Property Let RowMax(arg As Long)
    mRowMax = arg
End Property

'ColMax
Public Property Get ColMax() As Long
    ColMax = mColMax
End Property
Public Property Let ColMax(arg As Long)
    mColMax = arg
End Property

