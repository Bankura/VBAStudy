VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WbemScripting.SWbemPropertyクラスのラップ・拡張クラス
'* [詳  細] WbemScripting.SWbemPropertyクラスのWrapperとしての機能を提供する他、
'*          WbemScripting.SWbemPropertyクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WbemScripting.SWbemPropertyクラスについて。
' <https://msdn.microsoft.com/ja-jp/windows/aa393804(v=vs.80)>
' A Property
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemProperty
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
'* CIMType プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The CIM Type of this Property.
'*-----------------------------------------------------------------------------
Property Get CIMType() As WbemCimtypeEnum
    CIMType = mOrigin.CIMType
End Property

'*-----------------------------------------------------------------------------
'* IsArray プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Indicates whether this Property is an array type.
'*-----------------------------------------------------------------------------
Property Get IsArray() As Boolean
    IsArray = mOrigin.IsArray
End Property

'*-----------------------------------------------------------------------------
'* IsLocal プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Indicates whether this Property is local or propagated.
'*-----------------------------------------------------------------------------
Property Get IsLocal() As Boolean
    IsLocal = mOrigin.IsLocal
End Property
    
'*-----------------------------------------------------------------------------
'* Name プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The name of this Property.
'*-----------------------------------------------------------------------------
Property Get Name() As String
    Name = mOrigin.Name
End Property
    
'*-----------------------------------------------------------------------------
'* Origin プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The originating class of this Property.
'*-----------------------------------------------------------------------------
Property Get Origin() As String
    Origin = mOrigin.Origin
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
    
'*-----------------------------------------------------------------------------
'* Value プロパティ（既定メンバー）
'*
'* [補  足]
'* ・The value of this Property.
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
    'Set mOrigin = CreateObject("WbemScripting.SWbemProperty")
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


