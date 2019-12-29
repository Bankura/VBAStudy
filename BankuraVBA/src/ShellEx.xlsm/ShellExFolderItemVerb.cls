VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShellExFolderItemVerb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Shell32.FolderItemVerbクラスのラップ・拡張クラス
'* [詳  細] Shell32.FolderItemVerbクラスのWrapperとしての機能を提供する他、
'*          Shell32.FolderItemVerbクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' Shell32.FolderItemVerbクラスについて。
' <xxxxxxxxxxxxxx>
'
' Definition of interface FolderItemVerb
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mOrigin As shell32.FolderItemVerb
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
'* Name プロパティ（読み取り専用：既定メンバー）
'*
'* [補  足]
'* ・Get display name for item
'*-----------------------------------------------------------------------------
Property Get Name() As String
Attribute Name.VB_UserMemId = 0
    Name = mOrigin.Name
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
    'Set mOrigin = CreateObject("Shell32.FolderItemVerb")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************
    
'******************************************************************************
'* [概  要] DoIt メソッド
'* [詳  細] DoIt のラッパーメソッド。
'*          Execute the verb.
'*
'******************************************************************************
Sub DoIt()
    Call mOrigin.DoIt
End Sub
