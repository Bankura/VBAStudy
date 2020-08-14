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
'* [�@�\��] Shell32.Folder�N���X�̃��b�v�E�g���N���X
'* [��  ��] Shell32.Folder�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          Shell32.Folder�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' Shell32.Folder�N���X�ɂ��āB
' <xxxxxxxxxxxxxx>
'
' Definition of interface Folder.
'
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
Private mOrigin As shell32.Folder
'Private mOrigin As Object

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
'* Parent �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EGet Parent object
'*-----------------------------------------------------------------------------
Property Get Parent() As Object
    Set Parent = mOrigin.Parent
End Property

'*-----------------------------------------------------------------------------
'* ParentFolder �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EGet Parent object
'*-----------------------------------------------------------------------------
Property Get ParentFolder() As ShellExFolder
    Dim wrap As ShellExFolder: Set wrap = New ShellExFolder
    Set wrap.Origin = mOrigin.ParentFolder
    Set ParentFolder = wrap
End Property
    
'*-----------------------------------------------------------------------------
'* Title �v���p�e�B�i�ǂݎ���p:���胁���o�[�j
'*
'* [��  ��]
'* �EGet the display name for the window
'*-----------------------------------------------------------------------------
Property Get Title() As String
Attribute Title.VB_UserMemId = 0
    Title = mOrigin.Title
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("Shell32.Folder")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] CopyHere ���\�b�h
'* [��  ��] CopyHere �̃��b�p�[���\�b�h�B
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
'* [�T  �v] GetDetailsOf ���\�b�h
'* [��  ��] GetDetailsOf �̃��b�p�[���\�b�h�B
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
'* [�T  �v] Items ���\�b�h
'* [��  ��] Items �̃��b�p�[���\�b�h�B
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
'* [�T  �v] MoveHere ���\�b�h
'* [��  ��] MoveHere �̃��b�p�[���\�b�h�B
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
'* [�T  �v] NewFolder ���\�b�h
'* [��  ��] NewFolder �̃��b�p�[���\�b�h�B
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
'* [�T  �v] ParseName ���\�b�h
'* [��  ��] ParseName �̃��b�p�[���\�b�h�B
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