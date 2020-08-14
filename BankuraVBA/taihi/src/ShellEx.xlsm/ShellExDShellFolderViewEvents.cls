VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShellExDShellFolderViewEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Shell32.DShellFolderViewEventsクラスのラップ・拡張クラス
'* [詳  細] Shell32.DShellFolderViewEventsクラスのWrapperとしての機能を提供する他、
'*          Shell32.DShellFolderViewEventsクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' Shell32.DShellFolderViewEventsクラスについて。
' <xxx>
'
' Event interface for ShellFolderView.
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mOrigin As shell32.DShellFolderViewEvents
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

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("Shell32.DShellFolderViewEvents")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] BeginDrag メソッド
'* [詳  細] BeginDrag のラッパーメソッド。
'*          user started to drag an item (return false to cancel).
'*
'******************************************************************************
Function BeginDrag() As Boolean
    BeginDrag = mOrigin.BeginDrag
End Function

'******************************************************************************
'* [概  要] DefaultVerbInvoked メソッド
'* [詳  細] DefaultVerbInvoked のラッパーメソッド。
'*          the default verb (double click) was invoked on an items in the view (return false to cancel).
'*
'******************************************************************************
Function DefaultVerbInvoked() As Boolean
    DefaultVerbInvoked = mOrigin.DefaultVerbInvoked
End Function

'******************************************************************************
'* [概  要] EnumDone メソッド
'* [詳  細] EnumDone のラッパーメソッド。
'*          The folder has finished enumerating (flashlight is gone).
'*
'******************************************************************************
Sub EnumDone()
    Call mOrigin.EnumDone
End Sub

'******************************************************************************
'* [概  要] SelectionChanged メソッド
'* [詳  細] SelectionChanged のラッパーメソッド。
'*          The Selection in the view changed.
'*
'******************************************************************************
Sub SelectionChanged()
    Call mOrigin.SelectionChanged
End Sub

'******************************************************************************
'* [概  要] VerbInvoked メソッド
'* [詳  細] VerbInvoked のラッパーメソッド。
'*          A verb was invoked on an items in the view (return false to cancel).
'*
'******************************************************************************
Function VerbInvoked() As Boolean
    VerbInvoked = mOrigin.VerbInvoked
End Function
