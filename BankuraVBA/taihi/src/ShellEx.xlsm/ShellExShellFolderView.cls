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
'* [�@�\��] Shell32.ShellFolderView�N���X�̃��b�v�E�g���N���X
'* [��  ��] Shell32.ShellFolderView�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          Shell32.ShellFolderView�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' Shell32.ShellFolderView�N���X�ɂ��āB
' <xxxxxxxxxxxxxx>
'
' Shell Folder View Object
'
' [�C�x���g���]
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
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As shell32.ShellFolderView
Private mOrigin As Object

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************
'*-----------------------------------------------------------------------------
'* �I���W�i��Object �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get Origin() As Object
    Set Origin = mOrigin
End Property
Property Set Origin(arg As Object)
    Set mOrigin = arg
End Property

'*-----------------------------------------------------------------------------
'* Application �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EGet Application object
'*-----------------------------------------------------------------------------
Property Get Application() As Object
    Set Application = mOrigin.Application
End Property

'*-----------------------------------------------------------------------------
'* CurrentViewMode �v���p�e�B
'*
'* [��  ��]
'* �E�T�|�[�g����Ă��Ȃ��o���A���g�^(unsigned int)
'* �EGet Current View Mode
'*-----------------------------------------------------------------------------
Property Get CurrentViewMode() As Long
    CurrentViewMode = mOrigin.CurrentViewMode
End Property
Property Let CurrentViewMode(arg As Long)
    mOrigin.CurrentViewMode = arg
End Property

'*-----------------------------------------------------------------------------
'* FocusedItem �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EThe currently focused item in the folder
'*-----------------------------------------------------------------------------
Property Get FocusedItem() As ShellExFolderItem
    Dim wrap As ShellExFolderItem: Set wrap = New ShellExFolderItem
    Set wrap.Origin = mOrigin.FocusedItem
    Set FocusedItem = wrap
End Property

'*-----------------------------------------------------------------------------
'* Folder �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EGet the folder being viewed
'*-----------------------------------------------------------------------------
Property Get Folder() As ShellExFolder
    Dim wrap As ShellExFolder: Set wrap = New ShellExFolder
    Set wrap.Origin = mOrigin.Folder
    Set Folder = wrap
End Property

'*-----------------------------------------------------------------------------
'* FolderFlags �v���p�e�B
'*
'* [��  ��]
'* �E�T�|�[�g����Ă��Ȃ��o���A���g�^
'* �EGet Folder Flags
'*-----------------------------------------------------------------------------
Property Get FolderFlags() As Long
    FolderFlags = mOrigin.FolderFlags
End Property
Property Let FolderFlags(arg As Long)
    mOrigin.FolderFlags = arg
End Property

'*-----------------------------------------------------------------------------
'* GetFolder �v���p�e�B
'*
'* [��  ��]
'* �EGet Group By Column
'*-----------------------------------------------------------------------------
Property Get GroupBy() As String
    GroupBy = mOrigin.GroupBy
End Property
Property Let GroupBy(arg As String)
    mOrigin.GroupBy = arg
End Property

'*-----------------------------------------------------------------------------
'* IconSize �v���p�e�B
'*
'* [��  ��]
'* �ESet Icon Size
'*-----------------------------------------------------------------------------
Property Get IconSize() As Long
    IconSize = mOrigin.IconSize
End Property
Property Let IconSize(arg As Long)
    mOrigin.IconSize = arg
End Property
    
'*-----------------------------------------------------------------------------
'* Parent �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EGet Parent object
'*-----------------------------------------------------------------------------
Property Get Parent() As Object
    Set Parent = mOrigin.Parent
End Property

'*-----------------------------------------------------------------------------
'* Script �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EReturns the scripting automation model.
'*-----------------------------------------------------------------------------
Property Get Script() As Object
    Set Script = mOrigin.Script
End Property

'*-----------------------------------------------------------------------------
'* SortColumns �v���p�e�B
'*
'* [��  ��]
'* �EGet Sorting Columns
'*-----------------------------------------------------------------------------
Property Get SortColumns() As String
    SortColumns = mOrigin.SortColumns
End Property
Property Let SortColumns(arg As String)
    mOrigin.SortColumns = arg
End Property

'*-----------------------------------------------------------------------------
'* ViewOptions �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EReturns the view options for showing a folder.
'*-----------------------------------------------------------------------------
Property Get ViewOptions() As Long
    ViewOptions = mOrigin.ViewOptions
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("Shell32.ShellFolderView")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] FilterView ���\�b�h
'* [��  ��] FilterView �̃��b�p�[���\�b�h�B
'*          Filter View.
'*
'* @param bstrFilterText
'*
'******************************************************************************
Sub FilterView(bstrFilterText As String)
    Call mOrigin.FilterView(bstrFilterText)
End Sub

    
'******************************************************************************
'* [�T  �v] PopupItemMenu ���\�b�h
'* [��  ��] PopupItemMenu �̃��b�p�[���\�b�h�B
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
'* [�T  �v] SelectedItems ���\�b�h
'* [��  ��] SelectedItems �̃��b�p�[���\�b�h�B
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
'* [�T  �v] SelectItem ���\�b�h
'* [��  ��] SelectItem �̃��b�p�[���\�b�h�B
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
'* [�T  �v] SelectItemRelative ���\�b�h
'* [��  ��] SelectItemRelative �̃��b�p�[���\�b�h�B
'*          Select Item relative to the Current Item
'*
'* @param iRelative
'*
'******************************************************************************
Sub SelectItemRelative(iRelative As Long)
    Call mOrigin.SelectItemRelative(iRelative)
End Sub
