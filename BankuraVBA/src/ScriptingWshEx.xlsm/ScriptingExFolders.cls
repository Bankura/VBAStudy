VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScriptingExFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Scripting.Foldersクラスのラップ・拡張クラス
'* [詳  細] Scripting.FoldersクラスのWrapperとしての機能を提供する他、
'*          Scripting.Foldersクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' Scripting.Foldersクラスについて。
' <https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/folders-collection>
' Folder オブジェクトに含まれるすべての Folder オブジェクトのコレクションです。
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As Scripting.Folders
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
    Dim myFolder As Object
    For Each myFolder In mOrigin
        Dim wrap As ScriptingExFolder: Set wrap = New ScriptingExFolder
        Set wrap.Origin = myFolder
        mKeys.Add wrap, myFolder.Name
    Next myFolder
End Property

'*-----------------------------------------------------------------------------
'* Count プロパティ（読み取り専用）
'*
'* [補  足]
'* ・フォルダの数｡
'*-----------------------------------------------------------------------------
Property Get Count() As Long
    Count = mOrigin.Count
End Property

'*-----------------------------------------------------------------------------
'* Item プロパティ（読み取り専用:既定メンバー）
'*
'* [補  足]
'* ・フォルダを取得します。
'*-----------------------------------------------------------------------------
Property Get Item(Key) As ScriptingExFolder
Attribute Item.VB_UserMemId = 0
    Dim wrap As ScriptingExFolder: Set wrap = New ScriptingExFolder
    Set wrap.Origin = mOrigin.Item(Key)
    Set Item = wrap
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("Scripting.Folders")
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
'* [概  要] Add メソッド
'* [詳  細] Add のラッパーメソッド。
'*          新規フォルダを作成します｡
'*
'* @param Name
'* @return ScriptingExFolder 新規フォルダ
'*
'******************************************************************************
Function Add(Name As String) As ScriptingExFolder
    Dim wrap As ScriptingExFolder: Set wrap = New ScriptingExFolder
    Set wrap.Origin = mOrigin.Add(Name)
    Set Add = wrap
    mKeys.Add wrap, Name
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
    Set NewEnum = mKeys.[_NewEnum]
End Function

