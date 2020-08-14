VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADODBExProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] ADODB.Propertiesクラスのラップ・拡張クラス
'* [詳  細] ADODB.PropertiesクラスのWrapperとしての機能を提供する他、
'*          ADODB.Propertiesクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' ADODB.Propertiesクラスについて。
' <https://docs.microsoft.com/en-us/previous-versions/ms805115%28v%3dmsdn.10%29>
' <https://docs.microsoft.com/ja-jp/sql/ado/reference/ado-api/connection-object-ado?view=sql-server-2017>
' <https://docs.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/ado-api-reference>
'
' Contains all the Property objects for a specific instance of an object.
'
' オブジェクトの特定のインスタンスのすべてのPropertyオブジェクトが含まれる。
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As ADODB.Properties
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
    Set mKeys = New Collection
    Dim myProp As Object
    For Each myProp In mOrigin
        Dim wrap As ADODBExProperty: Set wrap = New ADODBExProperty
        Set wrap.Origin = myProp
        mKeys.Add wrap, myProp.Name
    Next myProp
End Property

'*-----------------------------------------------------------------------------
'* Count プロパティ（読み取り専用）
'*
'* [補  足]
'* ・コレクション内のオブジェクト数を示す。
'* ・Count プロパティは､特定のコレクション内のオブジェクトの数を調べるために
'*   使う｡
'* ・コレクションのメンバーは 0 から順に番号が割り当てられるため、ループを使う
'*   場合は常に 0 から始めて、 Count プロパティより 1 小さい値で終わらせる必要
'*   がある。
'*   Microsoft Visual Basic で Count プロパティをチェックせずにコレクションの
'*   メンバーをループ処理するには、 For Each...Next コマンドを使う。
'* ・Count が 0 の場合、コレクションにはオブジェクトが含まれていないことを意味
'*   する。
'*-----------------------------------------------------------------------------
Property Get Count() As Long
    Count = mOrigin.Count
End Property

'*-----------------------------------------------------------------------------
'* Item プロパティ（読み取り専用:既定メンバー）
'*
'* [補  足]
'* ・コレクションの特定のメンバーをその名前または序数で示す。
'* ・Indexには、コレクション内のオブジェクトの名前または序数に評価される
'*   バリアント型 ( Variant ) の式を指定。
'* ・Item プロパティは、コレクション内の特定のオブジェクトを返すために使う。
'*   コレクション内で Item が Index 引数に対応するオブジェクトを見つけられない
'*   場合は、エラーが発生する。
'*   また、コレクションの中には名前付きオブジェクトをサポートしていないものも
'*   ある。このようなコレクションでは、序数参照を使う必要がある。
'* ・Item プロパティはすべてのコレクションの既定プロパティなので､次のいずれの
'*   構文形式でも同じ結果が得られる｡
'*     collection.Item (Index)
'*     collection (Index)
'*-----------------------------------------------------------------------------
Property Get Item(Index) As ADODBExProperty
Attribute Item.VB_UserMemId = 0
    Dim wrap As ADODBExProperty: Set wrap = New ADODBExProperty
    Set wrap.Origin = mOrigin.Item(Index)
    Set Item = wrap
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Public Sub Class_Initialize()
    'Set mOrigin = CreateObject("ADODB.Properties")
    'Set mKeys = New Collection
End Sub
Public Sub Class_Terminate()
    Set mOrigin = Nothing
    Set mKeys = Nothing
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************
'なし

'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------

'******************************************************************************
'* [概  要] NewEnum メソッド
'* [詳  細] For Each 構文でイテレーターを返すためのメソッド。
'*
'* @return イテレーター
'******************************************************************************
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mKeys.[_NewEnum]
End Function
