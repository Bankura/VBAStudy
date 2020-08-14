VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShellExFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Shell32.Folderクラスのラップ・拡張クラス
'* [詳  細] Shell32.FolderクラスのWrapperとしての機能を提供する他、
'*          Shell32.Folderクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' Shell32.Folderクラスについて。
' <xxxxxxxxxxxxxx>
'
' Definition of interface Folder.
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mOrigin As shell32.Folder
'Private mOrigin As Object

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
'* Application プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Get Application object
'*-----------------------------------------------------------------------------
Property Get Application() As Object
    Set Application = mOrigin.Application
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

'*-----------------------------------------------------------------------------
'* ParentFolder プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Get Parent object
'*-----------------------------------------------------------------------------
Property Get ParentFolder() As ShellExFolder
    Dim wrap As ShellExFolder: Set wrap = New ShellExFolder
    Set wrap.Origin = mOrigin.ParentFolder
    Set ParentFolder = wrap
End Property
    
'*-----------------------------------------------------------------------------
'* Title プロパティ（読み取り専用:既定メンバー）
'*
'* [補  足]
'* ・Get the display name for the window
'*-----------------------------------------------------------------------------
Property Get Title() As String
Attribute Title.VB_UserMemId = 0
    Title = mOrigin.Title
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("Shell32.Folder")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] CopyHere メソッド
'* [詳  細] CopyHere のラッパーメソッド。
'*          Copy Items to this folder.
'*
'* @param vItem
'* @param vOptions
'*
'******************************************************************************
Sub CopyHere(vItem, Optional vOptions)
    If IsMissing(vOptions) Then
       Call mOrigin.CopyHere(vItem)
    Else
       Call mOrigin.CopyHere(vItem, vOptions)
    End If
End Sub
 
'******************************************************************************
'* [概  要] GetDetailsOf メソッド
'* [詳  細] GetDetailsOf のラッパーメソッド。
'*          Get the details about an item.
'*
'* @param vItem
'* @param iColumn
'* @return String
'******************************************************************************
Function GetDetailsOf(vItem, iColumn As Long) As String
    GetDetailsOf = mOrigin.GetDetailsOf(vItem, iColumn)
End Function

'******************************************************************************
'* [概  要] Items メソッド
'* [詳  細] Items のラッパーメソッド。
'*          The collection of Items in folder.
'*
'* @return FolderItems
'******************************************************************************
Function Items() As ShellExFolderItems
    Dim wrap As ShellExFolderItems: Set wrap = New ShellExFolderItems
    Set wrap.Origin = mOrigin.Items()
    Set Items = wrap
End Function

'******************************************************************************
'* [概  要] MoveHere メソッド
'* [詳  細] MoveHere のラッパーメソッド。
'*          Create a new sub folder in this folder.
'*
'* @param vItem
'* @param vOptions
'*
'******************************************************************************
Sub MoveHere(vItem, Optional vOptions)
    If IsMissing(vOptions) Then
       Call mOrigin.MoveHere(vItem)
    Else
       Call mOrigin.MoveHere(vItem, vOptions)
    End If
End Sub
    
'******************************************************************************
'* [概  要] NewFolder メソッド
'* [詳  細] NewFolder のラッパーメソッド。
'*          Create a new sub folder in this folder.
'*
'* @param bName
'* @param vOptions
'******************************************************************************
Sub NewFolder(bName As String, Optional vOptions)
    If IsMissing(vOptions) Then
        Call mOrigin.NewFolder(bName)
    Else
       Call mOrigin.NewFolder(bName, vOptions)
    End If
End Sub

'******************************************************************************
'* [概  要] ParseName メソッド
'* [詳  細] ParseName のラッパーメソッド。
'*          Parse the name to get an item.
'*
'* @param bName
'* @return FolderItem
'******************************************************************************
Function ParseName(bName As String) As ShellExFolderItem
    Dim wrap As ShellExFolderItem: Set wrap = New ShellExFolderItem
    Set wrap.Origin = mOrigin.ParseName(bName)
    Set ParseName = wrap
End Function
