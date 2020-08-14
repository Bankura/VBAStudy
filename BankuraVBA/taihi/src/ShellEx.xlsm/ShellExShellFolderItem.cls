VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShellExShellFolderItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] Shell32.ShellFolderItem�N���X�̃��b�v�E�g���N���X
'* [��  ��] Shell32.ShellFolderItem�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          Shell32.ShellFolderItem�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' Shell32.ShellFolderItem�N���X�ɂ��āB
' <xxxxxxxxxxxxxx>
'
' Shell Folder Item
'
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
Private mOrigin As shell32.ShellFolderItem
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
'* GetFolder �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EIf item is a folder return folder object
'*-----------------------------------------------------------------------------
Property Get GetFolder() As Object
    Set GetFolder = mOrigin.GetFolder
End Property

'*-----------------------------------------------------------------------------
'* GetLink �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EIf item is link return link object
'*-----------------------------------------------------------------------------
Property Get GetLink() As Object
    Set GetLink = mOrigin.GetLink
End Property

'*-----------------------------------------------------------------------------
'* IsBrowsable �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EIs the item browsable?
'*-----------------------------------------------------------------------------
Property Get IsBrowsable() As Boolean
    IsBrowsable = mOrigin.IsBrowsable
End Property

'*-----------------------------------------------------------------------------
'* IsFileSystem �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EIs the item a file system object?
'*-----------------------------------------------------------------------------
Property Get IsFileSystem() As Boolean
    IsFileSystem = mOrigin.IsFileSystem
End Property

'*-----------------------------------------------------------------------------
'* IsFolder �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EIs the item a Folder?
'*-----------------------------------------------------------------------------
Property Get IsFolder() As Boolean
    IsFolder = mOrigin.IsFolder
End Property

'*-----------------------------------------------------------------------------
'* IsLink �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EIs the item a link?
'*-----------------------------------------------------------------------------
Property Get IsLink() As Boolean
    IsLink = mOrigin.IsLink
End Property
    
'*-----------------------------------------------------------------------------
'* ModifyDate �v���p�e�B
'*
'* [��  ��]
'* �EModification Date?
'*-----------------------------------------------------------------------------
Property Get ModifyDate() As Date
    ModifyDate = mOrigin.ModifyDate
End Property
Property Let ModifyDate(arg As Date)
    mOrigin.ModifyDate = arg
End Property

'*-----------------------------------------------------------------------------
'* Name �v���p�e�B�i���胁���o�[�j
'*
'* [��  ��]
'* �EGet display name for item
'*-----------------------------------------------------------------------------
Property Get Name() As String
Attribute Name.VB_UserMemId = 0
    Name = mOrigin.Name
End Property
Property Let Name(arg As String)
    mOrigin.Name = arg
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
'* Path �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EGet the pathname to the item
'*-----------------------------------------------------------------------------
Property Get Path() As String
    Path = mOrigin.Path
End Property

'*-----------------------------------------------------------------------------
'* Size �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �ESize
'*-----------------------------------------------------------------------------
Property Get Size() As Long
    Size = mOrigin.Size
End Property

'*-----------------------------------------------------------------------------
'* Type_ �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �EType
'*-----------------------------------------------------------------------------
Property Get Type_() As String
    Type_ = mOrigin.Type
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("Shell32.ShellFolderItem")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] ExtendedProperty ���\�b�h
'* [��  ��] ExtendedProperty �̃��b�p�[���\�b�h�B
'*          Access an extended property.
'*
'* @param bstrPropName
'* @return Variant
'*
'******************************************************************************
Function ExtendedProperty(bstrPropName As String)
    ExtendedProperty = mOrigin.ExtendedProperty(bstrPropName)
End Function

'******************************************************************************
'* [�T  �v] InvokeVerb ���\�b�h
'* [��  ��] InvokeVerb �̃��b�p�[���\�b�h�B
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
'* [�T  �v] InvokeVerbEx ���\�b�h
'* [��  ��] InvokeVerbEx �̃��b�p�[���\�b�h�B
'*          Extended version of InvokeVerb
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
'* [�T  �v] Verbs ���\�b�h
'* [��  ��] Verbs �̃��b�p�[���\�b�h�B
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
