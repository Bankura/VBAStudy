VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemNamedValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WbemScripting.SWbemNamedValueクラスのラップ・拡張クラス
'* [詳  細] WbemScripting.SWbemNamedValueクラスのWrapperとしての機能を提供する他、
'*          WbemScripting.SWbemNamedValueクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WbemScripting.SWbemNamedValueクラスについて。
' <https://msdn.microsoft.com/ja-jp/windows/aa393731(v=vs.80)>
' A named value
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemNamedValue
Private mOrigin As Object
    
'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* オリジナルObject プロパティ
'*-----------------------------------------------------------------------------
Property Get Origin() As Object
    Set Origin = mOrigin
End Property
Property Set Origin(arg As Object)
    Set mOrigin = arg
End Property

'*-----------------------------------------------------------------------------
'* Name プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The Name of this Value.
'*-----------------------------------------------------------------------------
Property Get Name() As String
    Name = mOrigin.Name
End Property

'*-----------------------------------------------------------------------------
'* Value プロパティ（既定メンバー）
'*
'* [補  足]
'* ・The Value of this Named element.
'*-----------------------------------------------------------------------------
Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    If IsObject(mOrigin.Value) Then
        Set Value = mOrigin.Value
    Else
        Value = mOrigin.Value
    End If
End Property
Property Set Value(arg As Variant)
    If IsObject(arg) Then
        Set mOrigin.Value = arg
    Else
        mOrigin.Value = arg
    End If
End Property


'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WbemScripting.SWbemNamedValue")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************
'なし

'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------
