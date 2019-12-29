VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WbemScripting.SWbemSecurityクラスのラップ・拡張クラス
'* [詳  細] WbemScripting.SWbemSecurityクラスのWrapperとしての機能を提供する他、
'*          WbemScripting.SWbemSecurityクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WbemScripting.SWbemSecurityクラスについて。
' <https://msdn.microsoft.com/ja-jp/windows/aa393850(v=vs.80)>
' A Security Configurator
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemSecurity
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
'* AuthenticationLevel プロパティ
'*
'* [補  足]
'* ・The security authentication level.
'*-----------------------------------------------------------------------------
Property Get AuthenticationLevel() As WbemAuthenticationLevelEnum
    AuthenticationLevel = mOrigin.AuthenticationLevel
End Property
Property Let AuthenticationLevel(arg As WbemAuthenticationLevelEnum)
    mOrigin.AuthenticationLevel = arg
End Property

'*-----------------------------------------------------------------------------
'* ImpersonationLevel プロパティ
'*
'* [補  足]
'* ・The security impersonation level.
'*-----------------------------------------------------------------------------
Property Get ImpersonationLevel() As WbemImpersonationLevelEnum
    ImpersonationLevel = mOrigin.ImpersonationLevel
End Property
Property Let ImpersonationLevel(arg As WbemImpersonationLevelEnum)
    mOrigin.ImpersonationLevel = arg
End Property

'*-----------------------------------------------------------------------------
'* Privileges プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The collection of privileges for this object.
'*-----------------------------------------------------------------------------
Property Get Privileges() As WMIExSWbemPrivilegeSet
    Dim wrap As WMIExSWbemPrivilegeSet: Set wrap = New WMIExSWbemPrivilegeSet
    Set wrap.Origin = mOrigin.Privileges
    Set Privileges = wrap
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WbemScripting.SWbemSecurity")
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



