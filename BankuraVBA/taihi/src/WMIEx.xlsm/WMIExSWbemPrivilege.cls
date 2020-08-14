VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemPrivilege"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WbemScripting.SWbemPrivilegeクラスのラップ・拡張クラス
'* [詳  細] WbemScripting.SWbemPrivilegeクラスのWrapperとしての機能を提供する他、
'*          WbemScripting.SWbemPrivilegeクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WbemScripting.SWbemPrivilegeクラスについて。
' <https://msdn.microsoft.com/ja-jp/windows/aa393792(v=vs.80)>
' A Privilege Override
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemPrivilege
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
'* DisplayName プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The display name of the Privilege.
'*-----------------------------------------------------------------------------
Property Get DisplayName() As String
    DisplayName = mOrigin.DisplayName
End Property

'*-----------------------------------------------------------------------------
'* Identifier プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The Privilege identifier.
'*-----------------------------------------------------------------------------
Property Get Identifier() As WbemPrivilegeEnum
    Identifier = mOrigin.Identifier
End Property
    
'*-----------------------------------------------------------------------------
'* IsEnabled プロパティ（既定メンバー）
'*
'* [補  足]
'* ・Whether the Privilege is to be enabled or disabled.
'*-----------------------------------------------------------------------------
Property Get IsEnabled() As Boolean
Attribute IsEnabled.VB_UserMemId = 0
    IsEnabled = mOrigin.IsEnabled
End Property
Property Let IsEnabled(arg As Boolean)
    mOrigin.IsEnabled = arg
End Property

'*-----------------------------------------------------------------------------
'* Name プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The name of the Privilege.
'*-----------------------------------------------------------------------------
Property Get Name() As String
    Name = mOrigin.Name
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WbemScripting.SWbemPrivilege")
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

