Attribute VB_Name = "ScriptingEx"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] Scripting���b�v�E�g�����W���[��
'* [��  ��] Scripting��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�AScripting���g�p����
'*          ���[�e�B���e�B��񋟂���B
'*          ���b�v����Scripting���C�u�����͈ȉ��̂��̂Ƃ���B
'*              [name] Microsoft Scripting Runtime
'*              [dll] C:\Windows\System32\scrrun.dll
'* [�Q  �l]
'*  <https://docs.microsoft.com/ja-jp/office/vba/language/reference/objects-visual-basic-for-applications>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* Enum��`
'******************************************************************************

'*-----------------------------------------------------------------------------
'* CompareMethod
'*
'*-----------------------------------------------------------------------------
Public Enum CompareMethod
    BinaryCompare = 0
    DatabaseCompare = 2
    TextCompare = 1
End Enum

'*-----------------------------------------------------------------------------
'* DriveTypeConst
'* �h���C�u�̎�ނ�\��
'*-----------------------------------------------------------------------------
Public Enum DriveTypeConst
    CDRom = 4       'CD-ROM �h���C�u
    Fixed = 2       '�n�[�h�f�B�X�N
    RamDisk = 5     'RAM �f�B�X�N
    Remote = 3      '�l�b�g���[�N�h���C�u
    Removable = 1   '�����[�o�u���f�B�X�N
    UnknownType = 0 '�s��
End Enum

'*-----------------------------------------------------------------------------
'* FileAttribute
'*
'*-----------------------------------------------------------------------------
Public Enum FileAttribute
    Alias = 1024
    Archive = 32
    Compressed = 2048
    Directory = 16
    Hidden = 2
    Normal = 0
    ReadOnly = 1
    System = 4
    Volume = 8
End Enum

'*-----------------------------------------------------------------------------
'* IOMode
'*
'*-----------------------------------------------------------------------------
Public Enum IOMode
    ForAppending = 8
    ForReading = 1
    ForWriting = 2
End Enum

'*-----------------------------------------------------------------------------
'* SpecialFolderConst
'*
'*-----------------------------------------------------------------------------
Public Enum SpecialFolderConst
    SystemFolder = 1
    TemporaryFolder = 2
    WindowsFolder = 0
End Enum

'*-----------------------------------------------------------------------------
'* StandardStreamTypes
'*
'*-----------------------------------------------------------------------------
Public Enum StandardStreamTypes
    StdErr = 2
    StdIn = 0
    StdOut = 1
End Enum

'*-----------------------------------------------------------------------------
'* Tristate
'*
'*-----------------------------------------------------------------------------
Public Enum Tristate
    TristateFalse = 0
    TristateMixed = -2
    TristateTrue = -1
    TristateUseDefault = -2
End Enum

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************


