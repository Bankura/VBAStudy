VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScriptingExDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Scripting.Dictionaryクラスのラップ・拡張クラス
'* [詳  細] Scripting.DictionaryクラスのWrapperとしての機能を提供する他、
'*          Scripting.Dictionaryクラスを使用したユーティリティを提供する。
'*
'* [参  考] <https://miau.hatenablog.com/entry/20110203/1296706824>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' Scripting.Dictionaryクラスについて。
' <https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/dictionary-object>
' データ キー/アイテムのペアを保存するオブジェクト。
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As Scripting.Dictionary
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
    Dim k
    For Each k In mOrigin.Keys
        Call mKeys.Add(k)
    Next
End Property

'*-----------------------------------------------------------------------------
'* CompareMode プロパティ
'*
'* [補  足]
'* ・文字列を比較するメソッドを設定する。値の取得も可能。
'*-----------------------------------------------------------------------------
Property Get CompareMode() As CompareMethod
    CompareMode = mOrigin.CompareMode
End Property
Property Let CompareMode(arg As CompareMethod)
    mOrigin.CompareMode = arg
End Property

'*-----------------------------------------------------------------------------
'* Count プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ディクショナリ内の項目数を取得する。
'*-----------------------------------------------------------------------------
Property Get Count() As Long
    Count = mOrigin.Count
End Property


'*-----------------------------------------------------------------------------
'* Item プロパティ（既定メンバー）
'*
'* [補  足]
'* ・指定したキーに対する項目を設定します。値の取得も可能。
'*-----------------------------------------------------------------------------
Property Get Item(Key)
Attribute Item.VB_UserMemId = 0
    If IsObject(mOrigin(Key)) Then
        Set Item = mOrigin(Key)
    Else
        Item = mOrigin(Key)
    End If
End Property
Property Set Item(Key, arg)
    If IsObject(arg) Then
        Set mOrigin.Item(Key) = arg
    Else
        mOrigin.Item(Key) = arg
    End If
End Property

'*-----------------------------------------------------------------------------
'* Key プロパティ
'*
'* [補  足]
'* ・キーを別のキーに変更する。
'*-----------------------------------------------------------------------------
'Property Get Key(KeyValue)
'    Key = mOrigin.Key(KeyValue)
'End Property
Property Let Key(KeyValue, arg)
    If mOrigin.Exists(KeyValue) Then
        mOrigin.Key(KeyValue) = arg
        
        Dim i As Long
        For i = 1 To mKeys.Count
            If mKeys(i) = KeyValue Then
                mKeys.Add arg, After:=i
                mKeys.Remove i
                Exit Property
            End If
        Next
    End If
End Property

'*-----------------------------------------------------------------------------
'* HashVal プロパティ（非表示：読み取り専用）
'*
'* [補  足]
'* ・ディクショナリ内の項目数を取得する。
'*-----------------------------------------------------------------------------
Property Get HashVal(KeyValue)
    HashVal = mOrigin.HashVal(KeyValue)
End Property
    
'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("Scripting.Dictionary")
    Set mKeys = New Collection
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
    Set mKeys = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Addメソッド
'* [詳  細] Addのラッパーメソッド。
'*          ディクショナリに新しいキーおよび項目を追加する｡
'*
'* @param Key  キー
'* @param Item 項目値
'******************************************************************************
Sub Add(Key, Item)
    Call mOrigin.Add(Key, Item)
    Call mKeys.Add(Key)
End Sub

'******************************************************************************
'* [概  要] Existsメソッド
'* [詳  細] Existsのラッパーメソッド。
'*          指定したキーがディクショナリに含まれているかどうかを示す。
'*
'* @param Key キー
'* @return チェック結果（True:存在する、False:存在しない）
'******************************************************************************
Function Exists(Key) As Boolean
    Exists = mOrigin.Exists(Key)
End Function

'******************************************************************************
'* [概  要] Itemsメソッド
'* [詳  細] Itemsのラッパーメソッド。
'*          ディクショナリ内のすべての項目を含む配列を取得する｡
'*
'* @return ディクショナリ内のすべての項目を含む配列
'******************************************************************************
Function Items()
    Items = mOrigin.Items()
End Function

'******************************************************************************
'* [概  要] Keysメソッド
'* [詳  細] Keysのラッパーメソッド。
'*          ディクショナリ内のすべてのキーを含む配列を取得する｡
'*
'* @return ディクショナリ内のすべてのキーを含む配列
'******************************************************************************
Function Keys()
    Keys = mOrigin.Keys()
End Function

'******************************************************************************
'* [概  要] Removeメソッド
'* [詳  細] Removeのラッパーメソッド。
'*          指定したキーをディクショナリから削除する｡
'*
'* @param Key キー
'******************************************************************************
Sub Remove(Key)
    If TypeName(Key) = "String" Then
        Call mOrigin.Remove(Key)
        Call mKeys.Remove(Key)
    Else
        Dim k
        For Each k In mOrigin.Keys
            If mKeys(Key) = k Then
                Call mOrigin.Remove(k)
                mKeys.Remove Key
                Exit Sub
            End If
        Next
    End If
    Exit Sub
End Sub

'******************************************************************************
'* [概  要] RemoveAllメソッド
'* [詳  細] RemoveAllのラッパーメソッド。
'*          ディクショナリからすべての情報を削除する｡
'*
'******************************************************************************
Sub RemoveAll()
    Call mOrigin.RemoveAll
    Set mKeys = Nothing
    Set mKeys = New Collection
End Sub


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
