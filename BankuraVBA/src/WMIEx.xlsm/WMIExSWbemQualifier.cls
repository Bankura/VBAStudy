VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemQualifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WbemScripting.SWbemQualifierクラスのラップ・拡張クラス
'* [詳  細] WbemScripting.SWbemQualifierクラスのWrapperとしての機能を提供する他、
'*          WbemScripting.SWbemQualifierクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WbemScripting.SWbemQualifierクラスについて。
' <https://msdn.microsoft.com/ja-jp/windows/aa393817(v=vs.80)>
' A Qualifier
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemQualifier
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
'* IsAmended プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Determines whether the value of this Qualifier has been amended.
'*-----------------------------------------------------------------------------
Property Get IsAmended() As Boolean
    IsAmended = mOrigin.IsAmended
End Property

'*-----------------------------------------------------------------------------
'* IsLocal プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Indicates whether this Qualifier is local or propagated.
'*-----------------------------------------------------------------------------
Property Get IsLocal() As Boolean
    IsLocal = mOrigin.IsLocal
End Property

'*-----------------------------------------------------------------------------
'* IsOverridable プロパティ
'*
'* [補  足]
'* ・Determines whether this Qualifier can be overridden where propagated.
'*-----------------------------------------------------------------------------
Property Get IsOverridable() As Boolean
    IsOverridable = mOrigin.IsOverridable
End Property
Property Let IsOverridable(arg As Boolean)
    mOrigin.IsOverridable = arg
End Property

'*-----------------------------------------------------------------------------
'* Name プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The name of this Qualifier.
'*-----------------------------------------------------------------------------
Property Get Name() As String
    Name = mOrigin.Name
End Property
    
'*-----------------------------------------------------------------------------
'* PropagatesToInstance プロパティ
'*
'* [補  足]
'* ・Determines whether this Qualifier can propagate to instances.
'*-----------------------------------------------------------------------------
Property Get PropagatesToInstance() As Boolean
    PropagatesToInstance = mOrigin.PropagatesToInstance
End Property
Property Let PropagatesToInstance(arg As Boolean)
    mOrigin.PropagatesToInstance = arg
End Property
    
'*-----------------------------------------------------------------------------
'* PropagatesToSubclass プロパティ
'*
'* [補  足]
'* ・Determines whether this Qualifier can propagate to subclasses.
'*-----------------------------------------------------------------------------
Property Get PropagatesToSubclass() As Boolean
    PropagatesToSubclass = mOrigin.PropagatesToSubclass
End Property
Property Let PropagatesToSubclass(arg As Boolean)
    mOrigin.PropagatesToSubclass = arg
End Property

'*-----------------------------------------------------------------------------
'* Value プロパティ（既定メンバー）
'*
'* [補  足]
'* ・The value of this Qualifier.
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
    'Set mOrigin = CreateObject("WbemScripting.SWbemQualifier")
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


