VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShellExShellLinkObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Shell32.ShellLinkObjectクラスのラップ・拡張クラス
'* [詳  細] Shell32.ShellLinkObjectクラスのWrapperとしての機能を提供する他、
'*          Shell32.ShellLinkObjectクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' Shell32.ShellLinkObjectクラスについて。
' <xxxxxxxxxxxxxx>
'
' Shell Link object
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mOrigin As shell32.ShellLinkObject
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
'* Arguments プロパティ
'*
'* [補  足]
'* ・Get the arguments for the link
'*-----------------------------------------------------------------------------
Property Get Arguments() As String
    Arguments = mOrigin.Arguments
End Property
Property Let Arguments(arg As String)
    mOrigin.Arguments = arg
End Property

'*-----------------------------------------------------------------------------
'* Description プロパティ
'*
'* [補  足]
'* ・Get the description for the link
'*-----------------------------------------------------------------------------
Property Get Description() As String
    Description = mOrigin.Description
End Property
Property Let Description(arg As String)
    mOrigin.Description = arg
End Property

'*-----------------------------------------------------------------------------
'* Hotkey プロパティ
'*
'* [補  足]
'* ・Get the Hotkey for the link
'*-----------------------------------------------------------------------------
Property Get Hotkey() As Long
    Hotkey = mOrigin.Hotkey
End Property
Property Let Hotkey(arg As Long)
    mOrigin.Hotkey = arg
End Property

'*-----------------------------------------------------------------------------
'* Path プロパティ
'*
'* [補  足]
'* ・Get the path of the link
'*-----------------------------------------------------------------------------
Property Get Path() As String
    Path = mOrigin.Path
End Property
Property Let Path(arg As String)
    mOrigin.Path = arg
End Property

'*-----------------------------------------------------------------------------
'* ShowCommand プロパティ
'*
'* [補  足]
'* ・Get the Show Command for the link
'*-----------------------------------------------------------------------------
Property Get ShowCommand() As Long
    ShowCommand = mOrigin.ShowCommand
End Property
Property Let ShowCommand(arg As Long)
    mOrigin.ShowCommand = arg
End Property

'*-----------------------------------------------------------------------------
'* Target プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Get the target of a link object
'*-----------------------------------------------------------------------------
Property Get Target() As ShellExFolderItem
    Dim wrap As ShellExFolderItem: Set wrap = New ShellExFolderItem
    Set wrap.Origin = mOrigin.Target
    Set Target = wrap
End Property

'*-----------------------------------------------------------------------------
'* WorkingDirectory プロパティ
'*
'* [補  足]
'* ・Get the working directory for the link
'*-----------------------------------------------------------------------------
Property Get WorkingDirectory() As String
    WorkingDirectory = mOrigin.WorkingDirectory
End Property
Property Let WorkingDirectory(arg As String)
    mOrigin.WorkingDirectory = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("Shell32.ShellLinkObject")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] GetIconLocation メソッド
'* [詳  細] GetIconLocation のラッパーメソッド。
'*          Get the IconLocation for the link.
'*
'* @param pbs
'* @return Long
'*
'******************************************************************************
Function GetIconLocation(pbs As String) As Long
    GetIconLocation = GetIconLocation(pbs)
End Function

'******************************************************************************
'* [概  要] FilterView メソッド
'* [詳  細] FilterView のラッパーメソッド。
'*          Tell the link to resolve itself.
'*
'* @param fFlags
'*
'******************************************************************************
Sub Resolve(fFlags As Long)
    Call Resolve(fFlags)
End Sub

'******************************************************************************
'* [概  要] FilterView メソッド
'* [詳  細] FilterView のラッパーメソッド。
'*          Tell the link to save the changes.
'*
'* @param vWhere
'*
'******************************************************************************
Sub Save(Optional vWhere)
    If IsMissing(vWhere) Then
        Call Save
    Else
        Call Save(vWhere)
    End If
End Sub

'******************************************************************************
'* [概  要] FilterView メソッド
'* [詳  細] FilterView のラッパーメソッド。
'*          Set the IconLocation for the link.
'*
'* @param bstrFilterText
'*
'******************************************************************************
Sub SetIconLocation(bs As String, iIcon As Long)
    Call mOrigin.SetIconLocation(bs, iIcon)
End Sub
