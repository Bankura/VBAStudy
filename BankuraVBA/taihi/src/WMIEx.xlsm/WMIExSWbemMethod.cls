VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WbemScripting.SWbemMethodクラスのラップ・拡張クラス
'* [詳  細] WbemScripting.SWbemMethodクラスのWrapperとしての機能を提供する他、
'*          WbemScripting.SWbemMethodクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WbemScripting.SWbemMethodクラスについて。
' <https://msdn.microsoft.com/ja-jp/windows/aa393722(v=vs.80)>
' A Method
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemMethod
Private mOrigin As Object
    
'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* オリジナルObject プロパティ
'*-----------------------------------------------------------------------------
Property Get Original() As Object
    Set Original = mOrigin
End Property
Property Set Original(arg As Object)
    Set mOrigin = arg
End Property

'*-----------------------------------------------------------------------------
'* InParameters プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The in parameters for this Method.
'*-----------------------------------------------------------------------------
Property Get InParameters() As WMIExSWbemObject
    Dim wrap As WMIExSWbemObject: Set wrap = New WMIExSWbemObject
    Set wrap.Origin = mOrigin.InParameters
    Set InParameters = wrap
End Property

'*-----------------------------------------------------------------------------
'* Name プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The name of this Method.
'*-----------------------------------------------------------------------------
Property Get Name() As String
    Name = mOrigin.Name
End Property

'*-----------------------------------------------------------------------------
'* Origin プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The originating class of this Method.
'*-----------------------------------------------------------------------------
Property Get Origin() As String
    Origin = mOrigin.Origin
End Property

'*-----------------------------------------------------------------------------
'* OutParameters プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The out parameters for this Method.
'*-----------------------------------------------------------------------------
Property Get OutParameters() As WMIExSWbemObject
    Dim wrap As WMIExSWbemObject: Set wrap = New WMIExSWbemObject
    Set wrap.Origin = mOrigin.OutParameters
    Set OutParameters = wrap
End Property

'*-----------------------------------------------------------------------------
'* Qualifiers_ プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The collection of Qualifiers of this Method.
'*-----------------------------------------------------------------------------
Property Get Qualifiers_() As WMIExSWbemQualifierSet
    Dim wrap As WMIExSWbemQualifierSet: Set wrap = New WMIExSWbemQualifierSet
    Set wrap.Origin = mOrigin.Qualifiers_
    Set Qualifiers_ = wrap
End Property
    
'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WbemScripting.SWbemMethod")
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
