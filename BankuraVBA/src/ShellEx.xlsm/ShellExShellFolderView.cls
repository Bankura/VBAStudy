VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShellExShellFolderView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Shell32.ShellFolderViewクラスのラップ・拡張クラス
'* [詳  細] Shell32.ShellFolderViewクラスのWrapperとしての機能を提供する他、
'*          Shell32.ShellFolderViewクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' Shell32.ShellFolderViewクラスについて。
' <xxxxxxxxxxxxxx>
'
' Shell Folder View Object
'
' [イベント情報]
'Event BeginDrag()
'    user started to drag an item (return false to cancel).
'Event DefaultVerbInvoked()
'    the default verb (double click) was invoked on an items in the view (return false to cancel).
'Event EnumDone()
'    The folder has finished enumerating (flashlight is gone).
'Event SelectionChanged()
'    The Selection in the view changed.
'Event VerbInvoked()
'    A verb was invoked on an items in the view (return false to cancel).
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As shell32.ShellFolderView
Private mOrigin As Object

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
'* CurrentViewMode プロパティ
'*
'* [補  足]
'* ・サポートされていないバリアント型(unsigned int)
'* ・Get Current View Mode
'*-----------------------------------------------------------------------------
Property Get CurrentViewMode() As Long
    CurrentViewMode = mOrigin.CurrentViewMode
End Property
Property Let CurrentViewMode(arg As Long)
    mOrigin.CurrentViewMode = arg
End Property

'*-----------------------------------------------------------------------------
'* FocusedItem プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The currently focused item in the folder
'*-----------------------------------------------------------------------------
Property Get FocusedItem() As ShellExFolderItem
    Dim wrap As ShellExFolderItem: Set wrap = New ShellExFolderItem
    Set wrap.Origin = mOrigin.FocusedItem
    Set FocusedItem = wrap
End Property

'*-----------------------------------------------------------------------------
'* Folder プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Get the folder being viewed
'*-----------------------------------------------------------------------------
Property Get Folder() As ShellExFolder
    Dim wrap As ShellExFolder: Set wrap = New ShellExFolder
    Set wrap.Origin = mOrigin.Folder
    Set Folder = wrap
End Property

'*-----------------------------------------------------------------------------
'* FolderFlags プロパティ
'*
'* [補  足]
'* ・サポートされていないバリアント型
'* ・Get Folder Flags
'*-----------------------------------------------------------------------------
Property Get FolderFlags() As Long
    FolderFlags = mOrigin.FolderFlags
End Property
Property Let FolderFlags(arg As Long)
    mOrigin.FolderFlags = arg
End Property

'*-----------------------------------------------------------------------------
'* GetFolder プロパティ
'*
'* [補  足]
'* ・Get Group By Column
'*-----------------------------------------------------------------------------
Property Get GroupBy() As String
    GroupBy = mOrigin.GroupBy
End Property
Property Let GroupBy(arg As String)
    mOrigin.GroupBy = arg
End Property

'*-----------------------------------------------------------------------------
'* IconSize プロパティ
'*
'* [補  足]
'* ・Set Icon Size
'*-----------------------------------------------------------------------------
Property Get IconSize() As Long
    IconSize = mOrigin.IconSize
End Property
Property Let IconSize(arg As Long)
    mOrigin.IconSize = arg
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
'* Script プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Returns the scripting automation model.
'*-----------------------------------------------------------------------------
Property Get Script() As Object
    Set Script = mOrigin.Script
End Property

'*-----------------------------------------------------------------------------
'* SortColumns プロパティ
'*
'* [補  足]
'* ・Get Sorting Columns
'*-----------------------------------------------------------------------------
Property Get SortColumns() As String
    SortColumns = mOrigin.SortColumns
End Property
Property Let SortColumns(arg As String)
    mOrigin.SortColumns = arg
End Property

'*-----------------------------------------------------------------------------
'* ViewOptions プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Returns the view options for showing a folder.
'*-----------------------------------------------------------------------------
Property Get ViewOptions() As Long
    ViewOptions = mOrigin.ViewOptions
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("Shell32.ShellFolderView")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] FilterView メソッド
'* [詳  細] FilterView のラッパーメソッド。
'*          Filter View.
'*
'* @param bstrFilterText
'*
'******************************************************************************
Sub FilterView(bstrFilterText As String)
    Call mOrigin.FilterView(bstrFilterText)
End Sub

    
'******************************************************************************
'* [概  要] PopupItemMenu メソッド
'* [詳  細] PopupItemMenu のラッパーメソッド。
'*          Show items menu and return command selected.
'*
'* @param pfi
'* @param vx
'* @param vy
'* @return String
'*
'******************************************************************************
Function PopupItemMenu(pfi As ShellExFolderItem, Optional vx, Optional vy) As String
    If IsMissing(vx) Then
        If IsMissing(vy) Then
            PopupItemMenu = mOrigin.PopupItemMenu(pfi.Origin)
        Else
            PopupItemMenu = mOrigin.PopupItemMenu(pfi.Origin, , vy)
        End If
    ElseIf IsMissing(vy) Then
        PopupItemMenu = mOrigin.PopupItemMenu(pfi.Origin, vx)
    Else
        PopupItemMenu = mOrigin.PopupItemMenu(pfi.Origin, vx, vy)
    End If
End Function

    
'******************************************************************************
'* [概  要] SelectedItems メソッド
'* [詳  細] SelectedItems のラッパーメソッド。
'*          The collection of Selected Items in folder.
'*
'* @return ShellExFolderItems
'*
'******************************************************************************
Function SelectedItems() As ShellExFolderItems
    Dim wrap As ShellExFolderItems: Set wrap = New ShellExFolderItems
    Set wrap.Origin = mOrigin.SelectedItems()
    Set SelectedItems = wrap
End Function

'******************************************************************************
'* [概  要] SelectItem メソッド
'* [詳  細] SelectItem のラッパーメソッド。
'*          Select the item.
'*
'* @param pvfi
'* @param dwFlags
'*
'******************************************************************************
Sub SelectItem(pvfi, dwFlags As Long)
    Call mOrigin.SelectItem(pvfi, dwFlags)
End Sub

'******************************************************************************
'* [概  要] SelectItemRelative メソッド
'* [詳  細] SelectItemRelative のラッパーメソッド。
'*          Select Item relative to the Current Item
'*
'* @param iRelative
'*
'******************************************************************************
Sub SelectItemRelative(iRelative As Long)
    Call mOrigin.SelectItemRelative(iRelative)
End Sub

