VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADODBExProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] ADODB.Propertyクラスのラップ・拡張クラス
'* [詳  細] ADODB.PropertyクラスのWrapperとしての機能を提供する他、
'*          ADODB.Propertyクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' ADODB.Propertyクラスについて。
' <https://docs.microsoft.com/en-us/previous-versions/ms805115%28v%3dmsdn.10%29>
' <https://docs.microsoft.com/ja-jp/sql/ado/reference/ado-api/connection-object-ado?view=sql-server-2017>
' <https://docs.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/ado-api-reference>
'
' Represents a characteristic of an ADO object that is defined by the provider.
' ADO objects have two types of properties: built-in and dynamic.
' Built-in properties are those properties implemented in ADO and immediately
' available to any new object.
' The Property object is a container for dynamic properties, defined by the
' underlying provider.
'
' プロバイダーによって定義されるADOオブジェクトの特性を表す。
' ADOオブジェクトには、組み込みと動的の2種類のプロパティがある。 組み込みプロ
' パティは、ADOで実装され、新しいオブジェクトですぐに使用できるプロパティ。
' Propertyオブジェクトは、基になるプロバイダーによって定義される動的プロパティ
' のコンテナーである。
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As ADODB.Property
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
'* Attributes プロパティ
'*
'* [補  足]
'* ・オブジェクトの 1 つまたは複数の属性を示す。
'* ・Attributes プロパティは値の取得のみ可能で、その値は 1 つまたは複数の
'*   PropertyAttributesEnum 値の合計になる｡
'* ・複数の属性を設定する場合は、該当する定数の合計を使用できる。
'*   プロパティの値を､互換性のない定数を含む合計に設定すると､エラーが発生する｡
'* ・このプロパティは､クライアント側のConnectionオブジェクトでは使用できない｡
'*-----------------------------------------------------------------------------
Property Get Attributes() As Long
    Attributes = mOrigin.Attributes
End Property
Property Let Attributes(arg As Long)
    mOrigin.Attributes = arg
End Property

'*-----------------------------------------------------------------------------
'* Name プロパティ
'*
'* [補  足]
'* ・オブジェクトの名前を示す値を設定または取得する。
'* ・値は、  Property オブジクトでは読み取り専用になる。
'*-----------------------------------------------------------------------------
Property Get Name() As String
    Name = mOrigin.Name
End Property
Property Let Name(arg As String)
    mOrigin.Name = arg
End Property

'*-----------------------------------------------------------------------------
'* Type プロパティ（Typeが予約語のためType_に変更）
'*
'* [補  足]
'* ・操作の種類またはデータ型を示す。
'*-----------------------------------------------------------------------------
Property Get Type_() As DataTypeEnum
    Type_ = mOrigin.Type
End Property
Property Let Type_(arg As DataTypeEnum)
    mOrigin.Type = arg
End Property

'*-----------------------------------------------------------------------------
'* Value プロパティ（既定メンバー）
'*
'* [補  足]
'* ・オブジェクトに割り当てられた値を示す。
'* ・Value プロパティはさまざまな要因によって値の取得と設定が可能な場合と、
'*   値の取得のみが可能な場合がある。
'* ・ADO では、 Value プロパティを使用してロングバイナリデータを設定および
'*   取得可能。
'*-----------------------------------------------------------------------------
Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    If IsObject(mOrigin.Value) Then
        Set Value = mOrigin.Value
    Else
        Value = mOrigin.Value
    End If
End Property
Property Let Value(arg As Variant)
    If IsObject(arg) Then
        Set mOrigin.Value = arg
    Else
        mOrigin.Value = arg
    End If
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Public Sub Class_Initialize()
    'Set mOrigin = CreateObject("ADODB.Property")
End Sub
Public Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
