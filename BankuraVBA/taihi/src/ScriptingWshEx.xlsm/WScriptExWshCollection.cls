VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WScriptExWshCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WScript.WshCollectionクラスのラップ・拡張クラス
'* [詳  細] WScript.WshCollectionクラスのWrapperとしての機能を提供する他、
'*          WScript.WshCollectionクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WScript.WshCollectionクラスについて。
' Windows の特殊フォルダのコレクションWshSpecialFolders オブジェクトとして
' 使用します。
' <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364437(v=msdn.10)>。
'
' また、WScript.NetworkのEnumNetworkDrives,EnumPrinterConnectionsメソッドの
' 戻り値としても使用します。
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As IWshRuntimeLibrary.WshCollection
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
    Dim v
    For Each v In mOrigin
        mKeys.Add v
    Next
End Property

'*-----------------------------------------------------------------------------
'* length プロパティ（読み取り専用）
'*
'* [補  足]
'* ・コレクション内のオブジェクト数を示す。
'*-----------------------------------------------------------------------------
Property Get length() As Long
    length = mOrigin.length
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("WScript.WshCollection")
    'Set mKeys = New Collection
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
    Set mKeys = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Countメソッド
'* [詳  細] Countのラッパーメソッド。
'*          コレクションの要素数を返却。
'*
'* @return コレクションの要素数
'*
'******************************************************************************
Function Count() As Long
    Count = mOrigin.Count
End Function

'******************************************************************************
'* [概  要] Itemメソッド（既定メンバー）
'* [詳  細] Itemのラッパーメソッド。
'*          コレクションの要素を返却。
'*
'* @param Index コレクション内でのオブジェクトの位置 (インデックス)
'* @return コレクションの要素
'*
'******************************************************************************
Function Item(Index)
Attribute Item.VB_UserMemId = 0
    Item = mOrigin.Item(Index)
End Function

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
    'Set NewEnum = mOrigin.[_NewEnum]
    Set NewEnum = mKeys.[_NewEnum]
End Function
