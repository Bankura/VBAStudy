VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemLocator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WbemScripting.SWbemLocatorクラスのラップ・拡張クラス
'* [詳  細] WbemScripting.SWbemLocatorクラスのWrapperとしての機能を提供する他、
'*          WbemScripting.SWbemLocatorクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WbemScripting.SWbemLocatorクラスについて。
' <https://msdn.microsoft.com/ja-jp/windows/aa393719(v=vs.80)>
' Used to obtain Namespace connections
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemLocator
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
'* Security_ プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The Security Configurator for this Object.
'*-----------------------------------------------------------------------------
Property Get Security_() As WMIExSWbemSecurity
    Dim wrap As WMIExSWbemSecurity: Set wrap = New WMIExSWbemSecurity
    Set wrap.Origin = mOrigin.Security_
    Set Security_ = wrap
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("WbemScripting.SWbemLocator")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] ConnectServer メソッド
'* [詳  細] ConnectServer のラッパーメソッド。
'*          Connect to a Namespace.
'*
'* @param strServer xx
'* @param strNamespace xx
'* @param strUser xx
'* @param strPassword xx
'* @param strLocale xx
'* @param strAuthority xx
'* @param iSecurityFlags xx
'* @param objWbemNamedValueSet xx
'* @return WMIExSWbemServicesEx
'*
'******************************************************************************
Function ConnectServer(Optional strServer As String = ".", _
                       Optional strNamespace As String, _
                       Optional strUser As String, _
                       Optional strPassword As String, _
                       Optional strLocale As String, _
                       Optional strAuthority As String, _
                       Optional iSecurityFlags As Long, _
                       Optional objWbemNamedValueSet As Object) As WMIExSWbemServicesEx
    Dim wrap As WMIExSWbemServicesEx: Set wrap = New WMIExSWbemServicesEx
    Set wrap.Origin = mOrigin.ConnectServer(strServer, strNamespace, strUser, strPassword, strLocale, strAuthority, iSecurityFlags, objWbemNamedValueSet)
    Set ConnectServer = wrap
End Function

'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------

'******************************************************************************
'* [概  要] ConnectServerNoEx メソッド
'* [詳  細] ConnectServer のラッパーメソッド。
'           戻り値をWMIExSWbemServicesとして取得する。
'*          Connect to a Namespace.
'*
'* @param strServer xx
'* @param strNamespace xx
'* @param strUser xx
'* @param strPassword xx
'* @param strLocale xx
'* @param strAuthority xx
'* @param iSecurityFlags xx
'* @param objWbemNamedValueSet xx
'* @return WMIExSWbemServices
'*
'******************************************************************************
Function ConnectServerNoEx(Optional strServer As String = ".", _
                           Optional strNamespace As String, _
                           Optional strUser As String, _
                           Optional strPassword As String, _
                           Optional strLocale As String, _
                           Optional strAuthority As String, _
                           Optional iSecurityFlags As Long, _
                           Optional objWbemNamedValueSet As Object) As WMIExSWbemServices
    Dim wrap As WMIExSWbemServices: Set wrap = New WMIExSWbemServices
    Set wrap.Origin = mOrigin.ConnectServer(strServer, strNamespace, strUser, strPassword, strLocale, strAuthority, iSecurityFlags, objWbemNamedValueSet)
    Set ConnectServerNoEx = wrap
End Function

