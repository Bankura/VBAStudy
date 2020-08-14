VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDOExIDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] CDO.IDataSourceクラスのラップ・拡張クラス
'* [詳  細] CDO.IDataSourceクラスのWrapperとしての機能を提供する他、
'*          CDO.IDataSourceクラスを使用したユーティリティを提供する。
'*
'* [参  考] <xx>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' CDO.IDataSourceクラスについて。
' <xx>
' Defines methods, properties, and collections used to extract messages from or embed messages into other CDO message body parts.
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As CDO.IDataSource
Private mOrigin As Object
Private mKeys As Collection
 
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
'* ActiveConnection プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Not implemented. Reserved for future use.
'*-----------------------------------------------------------------------------
'Property Get ActiveConnection() As Connection
'    Set ActiveConnection = mOrigin.ActiveConnection
'End Property

'*-----------------------------------------------------------------------------
'* IsDirty プロパティ
'*
'* [補  足]
'* ・Indicates whether the local data has been changed since the last save or bind operation.
'*-----------------------------------------------------------------------------
Property Get IsDirty() As Boolean
    IsDirty = mOrigin.IsDirty
End Property
Property Let IsDirty(arg As Boolean)
    mOrigin.IsDirty = arg
End Property
   
'*-----------------------------------------------------------------------------
'* Source プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Returns the currently bound object. When you bind resources by URL, an ADO _Record interface is returned on an open Record object.
'*-----------------------------------------------------------------------------
Property Get Source() As IUnknown
    Set Source = mOrigin.Source
End Property
    
'*-----------------------------------------------------------------------------
'* SourceClass プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The interface name (type) of the currently bound object. When you bind resources by URL, the value _Record is returned.
'*-----------------------------------------------------------------------------
Property Get SourceClass() As String
    SourceClass = mOrigin.SourceClass
End Property

'*-----------------------------------------------------------------------------
'* SourceURL プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Not Implemented. Reserved for future use.
'*-----------------------------------------------------------------------------
'Property Get SourceURL() As String
'    SourceURL = mOrigin.SourceURL
'End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("CDO.IDataSource")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] OpenDs メソッド
'* [詳  細] Open のラッパーメソッド。
'*          Not implemented. Reserved for future use.
'*
'* @param SourceURL
'* @param ActiveConnection  省略可能。
'* @param Mode  省略可能。
'* @param CreateOptions  省略可能。
'* @param Options  省略可能。
'* @param UserName  省略可能。
'* @param Password  省略可能。
'*
'******************************************************************************
Sub OpenDs(SourceURL As String, Optional ActiveConnection As Object, _
                                Optional Mode As ConnectModeEnum, _
                                Optional CreateOptions As RecordCreateOptionsEnum = adFailIfNotExists, _
                                Optional Options As RecordOpenOptionsEnum, _
                                Optional UserName As String, _
                                Optional Password As String)
    Call mOrigin.Open(SourceURL, ActiveConnection, Mode, CreateOptions, Options, UserName, Password)
End Sub

'******************************************************************************
'* [概  要] OpenObject メソッド
'* [詳  細] OpenObject のラッパーメソッド。
'*          Binds to and opens data from the specified object.
'*
'* @param Source
'* @param InterfaceName
'*
'******************************************************************************
Sub OpenObject(Source As IUnknown, InterfaceName As String)
    Call mOrigin.OpenObject(Source, InterfaceName)
End Sub

'******************************************************************************
'* [概  要] Save メソッド
'* [詳  細] Save のラッパーメソッド。
'*          Saves data into the currently bound object.
'*
'******************************************************************************
Sub Save()
    Call mOrigin.Save
End Sub

'******************************************************************************
'* [概  要] SaveTo メソッド
'* [詳  細] SaveTo のラッパーメソッド。
'*          Not implemented. Reserved for future use.
'*
'* @param SourceURL
'* @param ActiveConnection  省略可能。
'* @param Mode  省略可能。
'* @param CreateOptions  省略可能。
'* @param Options  省略可能。
'* @param UserName  省略可能。
'* @param Password  省略可能。
'*
'******************************************************************************
Sub SaveTo(SourceURL As String, Optional ActiveConnection As Object, _
                                Optional Mode As ConnectModeEnum, _
                                Optional CreateOptions As RecordCreateOptionsEnum = adFailIfNotExists, _
                                Optional Options As RecordOpenOptionsEnum, _
                                Optional UserName As String, _
                                Optional Password As String)
    Call mOrigin.SaveTo(SourceURL, ActiveConnection, Mode, CreateOptions, Options, UserName, Password)
End Sub

'******************************************************************************
'* [概  要] SaveToContainer メソッド
'* [詳  細] SaveToContainer のラッパーメソッド。
'*          Not implemented. Reserved for future use.
'*
'* @param ContainerURL
'* @param ActiveConnection  省略可能。
'* @param Mode  省略可能。
'* @param CreateOptions  省略可能。
'* @param Options  省略可能。
'* @param UserName  省略可能。
'* @param Password  省略可能。
'*
'******************************************************************************
Sub SaveToContainer(ContainerURL As String, Optional ActiveConnection As Object, _
                                Optional Mode As ConnectModeEnum, _
                                Optional CreateOptions As RecordCreateOptionsEnum = adFailIfNotExists, _
                                Optional Options As RecordOpenOptionsEnum, _
                                Optional UserName As String, _
                                Optional Password As String)
    Call mOrigin.SaveToContainer(ContainerURL, ActiveConnection, Mode, CreateOptions, Options, UserName, Password)
End Sub

'******************************************************************************
'* [概  要] SaveToObject メソッド
'* [詳  細] SaveToObject のラッパーメソッド。
'*           Binds to and saves data into the specified object.
'*
'* @param Source
'* @param InterfaceName
'*
'******************************************************************************
Sub SaveToObject(Source As IUnknown, InterfaceName As String)
    Call mOrigin.SaveToObject(Source, InterfaceName)
End Sub
