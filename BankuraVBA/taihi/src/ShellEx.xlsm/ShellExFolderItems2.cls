VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShellExFolderItems2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Shell32.FolderItems2クラスのラップ・拡張クラス
'* [詳  細] Shell32.FolderItems2クラスのWrapperとしての機能を提供する他、
'*          Shell32.FolderItems2クラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' Shell32.FolderItems2クラスについて。
' <xxxxxxxxxxxxxx>
'
' Definition of interface FolderItems Version 2
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mOrigin As shell32.FolderItems2
'Private mOrigin As Object
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
    Dim myItem As Object
    For Each myItem In mOrigin
        Dim wrap As ShellExFolderItem: Set wrap = New ShellExFolderItem
        Set wrap.Origin = myItem
        mKeys.Add wrap, myItem.Name
    Next myItem
End Property

'*-----------------------------------------------------------------------------
'* Application プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Get Application object
'*-----------------------------------------------------------------------------
Property Get Application() As Object
    Set Application = mOrigin.Application
End Property

'*-----------------------------------------------------------------------------
'* Count プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Get count of items in the folder
'*-----------------------------------------------------------------------------
Property Get Count() As Long
    Count = mOrigin.Count
End Property

'*-----------------------------------------------------------------------------
'* Parent プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Get Parent object
'*-----------------------------------------------------------------------------
Property Get Parent() As Object
    Set Parent = mOrigin.Parent
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("Shell32.FolderItems2")
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
'* [概  要] InvokeVerbEx メソッド
'* [詳  細] InvokeVerbEx のラッパーメソッド。
'*          Extended version of InvokeVerb for a collection of Folder Items
'*
'* @param vVerb
'* @param vArgs
'*
'******************************************************************************
Sub InvokeVerbEx(Optional vVerb, Optional vArgs)
    If IsMissing(vVerb) Then
        If IsMissing(vArgs) Then
            Call mOrigin.InvokeVerbEx
        Else
            Call mOrigin.InvokeVerbEx(, vArgs)
        End If
    ElseIf IsMissing(vArgs) Then
        Call mOrigin.InvokeVerbEx(vVerb)
    Else
        Call mOrigin.InvokeVerbEx(vVerb, vArgs)
    End If
End Sub

'******************************************************************************
'* [概  要] Item メソッド
'* [詳  細] Item のラッパーメソッド。
'*          Return the figure for the given index
'*
'* @param index
'* @return ShellExFolderItem
'******************************************************************************
Function Item(Optional index) As ShellExFolderItem
    Dim wrap As ShellExFolderItem: Set wrap = New ShellExFolderItem
    If IsMissing(index) Then
       Set wrap.Origin = mOrigin.Item
    Else
       Set wrap.Origin = mOrigin.Item(index)
    End If
    Set Item = wrap
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
