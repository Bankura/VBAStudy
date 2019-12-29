VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShellExFolderItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Shell32.FolderItemクラスのラップ・拡張クラス
'* [詳  細] Shell32.FolderItemクラスのWrapperとしての機能を提供する他、
'*          Shell32.FolderItemクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' Shell32.FolderItemクラスについて。
' <xxxxxxxxxxxxxx>
'
' Definition of interface FolderItem
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mOrigin As shell32.FolderItem
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
'* GetFolder プロパティ（読み取り専用）
'*
'* [補  足]
'* ・If item is a folder return folder object
'*-----------------------------------------------------------------------------
Property Get GetFolder() As Object
    Set GetFolder = mOrigin.GetFolder
End Property

'*-----------------------------------------------------------------------------
'* GetLink プロパティ（読み取り専用）
'*
'* [補  足]
'* ・If item is link return link object
'*-----------------------------------------------------------------------------
Property Get GetLink() As Object
    Set GetLink = mOrigin.GetLink
End Property

'*-----------------------------------------------------------------------------
'* IsBrowsable プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Is the item browsable?
'*-----------------------------------------------------------------------------
Property Get IsBrowsable() As Boolean
    IsBrowsable = mOrigin.IsBrowsable
End Property

'*-----------------------------------------------------------------------------
'* IsFileSystem プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Is the item a file system object?
'*-----------------------------------------------------------------------------
Property Get IsFileSystem() As Boolean
    IsFileSystem = mOrigin.IsFileSystem
End Property

'*-----------------------------------------------------------------------------
'* IsFolder プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Is the item a Folder?
'*-----------------------------------------------------------------------------
Property Get IsFolder() As Boolean
    IsFolder = mOrigin.IsFolder
End Property

'*-----------------------------------------------------------------------------
'* IsLink プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Is the item a link?
'*-----------------------------------------------------------------------------
Property Get IsLink() As Boolean
    IsLink = mOrigin.IsLink
End Property
    
'*-----------------------------------------------------------------------------
'* ModifyDate プロパティ
'*
'* [補  足]
'* ・Modification Date?
'*-----------------------------------------------------------------------------
Property Get ModifyDate() As Date
    ModifyDate = mOrigin.ModifyDate
End Property
Property Let ModifyDate(arg As Date)
    mOrigin.ModifyDate = arg
End Property

'*-----------------------------------------------------------------------------
'* Name プロパティ（既定メンバー）
'*
'* [補  足]
'* ・Get display name for item
'*-----------------------------------------------------------------------------
Property Get Name() As String
Attribute Name.VB_UserMemId = 0
    Name = mOrigin.Name
End Property
Property Let Name(arg As String)
    mOrigin.Name = arg
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
'* Path プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Get the pathname to the item
'*-----------------------------------------------------------------------------
Property Get Path() As String
    Path = mOrigin.Path
End Property

'*-----------------------------------------------------------------------------
'* Size プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Size
'*-----------------------------------------------------------------------------
Property Get Size() As Long
    Size = mOrigin.Size
End Property

'*-----------------------------------------------------------------------------
'* Type_ プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Type
'*-----------------------------------------------------------------------------
Property Get Type_() As String
    Type_ = mOrigin.Type
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("Shell32.FolderItem")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] InvokeVerb メソッド
'* [詳  細] InvokeVerb のラッパーメソッド。
'*          Execute a command on the item.
'*
'* @param vVerb
'*
'******************************************************************************
Sub InvokeVerb(Optional vVerb)
    If IsMissing(vVerb) Then
       Call mOrigin.InvokeVerb
    Else
       Call mOrigin.InvokeVerb(vVerb)
    End If
End Sub

'******************************************************************************
'* [概  要] Verbs メソッド
'* [詳  細] Verbs のラッパーメソッド。
'*          Get the list of verbs for the object.
'*
'* @param vItem
'* @param iColumn
'* @return String
'******************************************************************************
Function Verbs() As ShellExFolderItemVerbs
    Dim wrap As ShellExFolderItemVerbs: Set wrap = New ShellExFolderItemVerbs
    Set wrap.Origin = mOrigin.Verbs()
    Set Verbs = wrap
End Function
