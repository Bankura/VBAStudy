VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDOExConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] CDO.Configurationクラスのラップ・拡張クラス
'* [詳  細] CDO.ConfigurationクラスのWrapperとしての機能を提供する他、
'*          CDO.Configurationクラスを使用したユーティリティを提供する。
'*
'* [参  考] <xx>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' CDO.Configurationクラスについて。
' <xx>
' Defines an object used to store configuration information for CDO objects.
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As CDO.Configuration
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
'* Fields プロパティ（読み取り専用）
'*
'* [補  足]
'* ・the Object 's Fields collection.
'*-----------------------------------------------------------------------------
Property Get Fields() As ADODBExFields
Attribute Fields.VB_UserMemId = 0
    Dim wrap As ADODBExFields: Set wrap = New ADODBExFields
    Set wrap.Origin = mOrigin.Fields()
    Set Fields = wrap
End Property
    
'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("CDO.Configuration")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] GetInterface メソッド
'* [詳  細] GetInterface のラッパーメソッド。
'*          Returns a specified interface on this object; provided for script languages.
'*
'* @param Interface  xx
'* @return Object
'******************************************************************************
Function GetInterface(Interface As String) As Object
    Set GetInterface = mOrigin.GetInterface(Interface)
End Function

'******************************************************************************
'* [概  要] Load メソッド
'* [詳  細] Load のラッパーメソッド。
'*          Loads the specified configuration.
'*
'* @param LoadFrom xx
'* @param URL xx
'******************************************************************************
Sub Load(LoadFrom As CdoConfigSource, Optional URL As String)
    Call mOrigin.Load(LoadFrom, URL)
End Sub

'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------
