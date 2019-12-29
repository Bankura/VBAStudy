VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemObjectPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WbemScripting.SWbemObjectPathクラスのラップ・拡張クラス
'* [詳  細] WbemScripting.SWbemObjectPathクラスのWrapperとしての機能を提供する他、
'*          WbemScripting.SWbemObjectPathクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WbemScripting.SWbemObjectPathクラスについて。
' <https://msdn.microsoft.com/ja-jp/windows/aa393746(v=vs.80)>
' Object Path
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemObjectPath
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
'* Authority プロパティ
'*
'* [補  足]
'* ・Defines authentication authority component of this path.
'*-----------------------------------------------------------------------------
Property Get Authority() As String
    Authority = mOrigin.Authority
End Property
Property Let Authority(arg As String)
    mOrigin.Authority = arg
End Property

'*-----------------------------------------------------------------------------
'* Class プロパティ
'*
'* [補  足]
'* ・The Class name.
'*-----------------------------------------------------------------------------
Property Get Class() As String
    Class = mOrigin.Class
End Property
Property Let Class(arg As String)
    mOrigin.Class = arg
End Property

'*-----------------------------------------------------------------------------
'* DisplayName プロパティ
'*
'* [補  足]
'* ・The Display Name for this path.
'*-----------------------------------------------------------------------------
Property Get DisplayName() As String
    DisplayName = mOrigin.DisplayName
End Property
Property Let DisplayName(arg As String)
    mOrigin.DisplayName = arg
End Property

'*-----------------------------------------------------------------------------
'* IsClass プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Indicates whether this path addresses a Class.
'*-----------------------------------------------------------------------------
Property Get IsClass() As Boolean
    IsClass = mOrigin.IsClass
End Property

'*-----------------------------------------------------------------------------
'* IsSingleton プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Indicates whether this path addresses a Singleton Instance.
'*-----------------------------------------------------------------------------
Property Get IsSingleton() As Boolean
    IsSingleton = mOrigin.IsSingleton
End Property

'*-----------------------------------------------------------------------------
'* Keys プロパティ（読み取り専用）
'*
'* [補  足]
'* The collection of Key value bindings for this path.
'*-----------------------------------------------------------------------------
Property Get Keys() As WMIExSWbemNamedValueSet
    Dim wrap As WMIExSWbemNamedValueSet: Set wrap = New WMIExSWbemNamedValueSet
    Set wrap.Origin = mOrigin.Keys
    Set Keys = wrap
End Property

'*-----------------------------------------------------------------------------
'* Locale プロパティ
'*
'* [補  足]
'* ・Defines locale component of this path.
'*-----------------------------------------------------------------------------
Property Get Locale() As String
    Locale = mOrigin.Locale
End Property
Property Let Locale(arg As String)
    mOrigin.Locale = arg
End Property

'*-----------------------------------------------------------------------------
'* Namespace プロパティ
'*
'* [補  足]
'* ・The Namespace path.
'*-----------------------------------------------------------------------------
Property Get Namespace() As String
    Namespace = mOrigin.Namespace
End Property
Property Let Namespace(arg As String)
    mOrigin.Namespace = arg
End Property

'*-----------------------------------------------------------------------------
'* ParentNamespace プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The parent Namespace path.
'*-----------------------------------------------------------------------------
Property Get ParentNamespace() As String
    ParentNamespace = mOrigin.ParentNamespace
End Property

'*-----------------------------------------------------------------------------
'* Path プロパティ（既定メンバー）
'*
'* [補  足]
'* ・The full path.
'*-----------------------------------------------------------------------------
Property Get Path() As String
Attribute Path.VB_UserMemId = 0
    Path = mOrigin.Path
End Property
Property Let Path(arg As String)
    mOrigin.Path = arg
End Property

'*-----------------------------------------------------------------------------
'* RelPath プロパティ
'*
'* [補  足]
'* ・The relative path.
'*-----------------------------------------------------------------------------
Property Get RelPath() As String
    RelPath = mOrigin.RelPath
End Property
Property Let RelPath(arg As String)
    mOrigin.RelPath = arg
End Property

'*-----------------------------------------------------------------------------
'* Security_ プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Defines the security components of this path.
'*-----------------------------------------------------------------------------
Property Get Security_() As WMIExSWbemSecurity
    Dim wrap As WMIExSWbemSecurity: Set wrap = New WMIExSWbemSecurity
    Set wrap.Origin = mOrigin.Security_
    Set Security_ = wrap
End Property

'*-----------------------------------------------------------------------------
'* Server プロパティ
'*
'* [補  足]
'* ・The name of the Server.
'*-----------------------------------------------------------------------------
Property Get Server() As String
    Server = mOrigin.Server
End Property
Property Let Server(arg As String)
    mOrigin.Server = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WbemScripting.SWbemObjectPath")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] SetAsClass メソッド
'* [詳  細] SetAsClass のラッパーメソッド。
'*          Coerce this path to address a Class.
'*
'******************************************************************************
Sub SetAsClass()
    Call mOrigin.SetAsClass
End Sub

'******************************************************************************
'* [概  要] SetAsSingleton メソッド
'* [詳  細] SetAsSingleton のラッパーメソッド。
'*          Coerce this path to address a Singleton Instance.
'*
'******************************************************************************
Sub SetAsSingleton()
    Call mOrigin.SetAsSingleton
End Sub

'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------
