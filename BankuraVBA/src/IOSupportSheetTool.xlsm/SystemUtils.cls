VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SystemUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] System���[�e�B���e�B�N���X
'* [��  ��] �V�X�e���S�ʂɊւ��郆�[�e�B���e�B��񋟂���B
'*          Static�N���X�iAttribute VB_PredeclaredId = True�j�Ƃ���B
'* [�Q  �l]
'*
'* @author Bankura
'* Copyright (c) 2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI��`
'******************************************************************************
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

'******************************************************************************
'* Enum��`
'******************************************************************************

'******************************************************************************
'* �萔��`
'******************************************************************************

'******************************************************************************
'* �����ϐ���`
'******************************************************************************

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************


'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
End Sub
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] GetScreenResolution
'* [��  ��] ��ʉ𑜓x���擾����
'*
'* @return Variant ��ʉ𑜓x�i1�����z��j0: ���� 1:����
'*
'******************************************************************************
Public Function GetScreenResolution() As Variant
    Dim x As Long, y As Long
    x = GetSystemMetrics(0) ' ��ʂ̕����擾
    y = GetSystemMetrics(1) ' ��ʂ̍������擾
    GetScreenResolution = Array(x, y)
End Function

'******************************************************************************
'* [�T  �v] DisableF1Key
'* [��  ��] F1�̃w���v�N���𖳌�������
'*
'* @param willUndo �����������Ƃɖ߂���
'*
'******************************************************************************
Public Sub DisableF1Key(Optional willUndo As Boolean = False)
    If willUndo Then
        Application.OnKey "{F1}"
    Else
        Application.OnKey "{F1}", ""
    End If
End Sub

'******************************************************************************
'* [�T  �v] DisableEscKey
'* [��  ��] ESC�L�[�ɂ��v���O�������s���f�𖳌�������
'*
'* @param willUndo �����������Ƃɖ߂���
'*
'******************************************************************************
Public Sub DisableEscKey(Optional willUndo As Boolean = False)
    If willUndo Then
        Application.EnableCancelKey = xlInterrupt
    Else
        Application.EnableCancelKey = xlDisabled
    End If
End Sub

'******************************************************************************
'* [�T  �v] GetShellFolderItem
'* [��  ��] �w�肵���t�H���_/�t�@�C���p�X��ShellFolderItem�I�u�W�F�N�g���擾����
'*
'* @param fPath �t�H���_/�t�@�C���p�X
'* @return ShellFolderItem�I�u�W�F�N�g
'*
'******************************************************************************
Public Function GetShellFolderItem(fPath As String) As Object
    If IO.FolderExists(fPath) Then
         Set GetShellFolderItem = Base.GetShell.NameSpace(CVar(fPath)).Self
    ElseIf IO.FileExists(fPath) Then
         Set GetShellFolderItem = Base.GetShell.NameSpace(CVar(IO.GetParentFolderName(fPath))).ParseName(IO.GetFileName(fPath))
    Else
        Err.Raise 9999, "SystemtUtils#GetShellFolderItem", "Folder/File not found. [" & fPath & "]"
    End If
End Function

'******************************************************************************
'* [�T  �v] DoFolderItemVerb
'* [��  ��] �w�肵���t�H���_/�t�@�C���p�X��ShellFolderItem�I�u�W�F�N�g���擾
'*          ���A�w�肵�������iVerb�j�����s����
'*
'* @param fPath �t�H���_/�t�@�C���p�X
'* @return ShellFolderItem�I�u�W�F�N�g
'*
'******************************************************************************
Public Sub DoFolderItemVerb(fPath As String, verbName As String)
    Dim folderItemObj As Object
    Set folderItemObj = GetShellFolderItem(fPath)
    
    If Not folderItemObj Is Nothing Then
        Dim v
        For Each v In folderItemObj.Verbs
            If v.Name = verbName Then
                v.doit
                Exit Sub
            End If
        Next
    End If
End Sub

'******************************************************************************
'* [�T  �v] CopyFromFolderItem
'* [��  ��] �w�肵���t�@�C��/�t�H���_���R�s�[����
'*          �i�G�N�X�v���[���̃R�s�[�Ɠ����������s���A�R�s�[�o�b�t�@�[�ɓ����j
'*
'* @param folderPath �t�H���_�p�X
'******************************************************************************
Public Sub CopyFromFolderItem(fPath As String)
    Call DoFolderItemVerb(fPath, "�R�s�[(&C)")
End Sub

'******************************************************************************
'* [�T  �v] CutFromFolderItem
'* [��  ��] �w�肵���t�@�C��/�t�H���_��؂��肷��
'*          �i�G�N�X�v���[���̐؂���Ɠ����������s���A�R�s�[�o�b�t�@�[�ɓ����j
'*
'* @param folderPath �t�H���_�p�X
'******************************************************************************
Public Sub CutFromFolderItem(fPath As String)
    Call DoFolderItemVerb(fPath, "�؂���(&T)")
End Sub

'******************************************************************************
'* [�T  �v] PasteInFolder
'* [��  ��] �w�肵���t�H���_�ɃR�s�[���̃t�@�C��/�t�H���_��\��t����
'*
'* @param folderPath �t�H���_�p�X
'******************************************************************************
Public Sub PasteInFolder(folderPath As String)
    If Not IO.FolderExists(folderPath) Then
        Err.Raise 9999, "SystemtUtils#PasteInFolder", "Folder not found. [" & folderPath & "]"
    End If

    Call DoFolderItemVerb(folderPath, "�\��t��(&P)")
End Sub

'******************************************************************************
'* [�T  �v] ShowFolderItemVerbs
'* [��  ��] FolderItem�I�u�W�F�N�g���������iVerb�j�̈ꗗ���C�~�f�B�G�C�g
'*          �E�B���h�E�ɕ\������
'*
'* @param folderItemObj Shell32.FolderItem�I�u�W�F�N�g
'******************************************************************************
Public Sub ShowFolderItemVerbs(folderItemObj As Object)
    If Not folderItemObj Is Nothing And ValidateUtils.IsFolderItem(folderItemObj) Then
        Debug.Print folderItemObj.Type
        Dim v
        For Each v In folderItemObj.Verbs
            Debug.Print v.Name
        Next
    End If
End Sub

'******************************************************************************
'* [�T  �v] AddToRecent ���\�b�h
'* [��  ��] Shell.Application��AddToRecent �̃��b�p�[���\�b�h�B
'*          �ŋߎg�p�����iMRU�j���X�g�Ƀt�@�C����ǉ�����B
'*
'* @param varFile
'* @param bstrCategory
'******************************************************************************
Public Sub AddToRecent(varFile, Optional bstrCategory As String)
    Call Base.GetShell.AddToRecent(varFile, bstrCategory)
End Sub

'******************************************************************************
'* [�T  �v] BrowseForFolder ���\�b�h
'* [��  ��] Shell.Application��BrowseForFolder �̃��b�p�[���\�b�h�B
'*          ���[�U���I�������t�H���_��Folder�I�u�W�F�N�g��Ԃ��_�C�A���O
'*          �{�b�N�X���쐬����B
'* [�Q  �l] http://www.roy.hi-ho.ne.jp/mutaguchi/wsh/object/shell.htm
'*
'* @param Hwnd       �E�B���h�E�̃n���h���B�ʏ��0�ɂ��Ă����Ė��Ȃ�
'* @param Title      �_�C�A���O�ɕ\�������������
'* @param Options    �_�C�A���O�̃^�C�v�B�����̃I�v�V�������w�肷�鎞�́A�萔�̘a���w��
'* @param RootFolder ���[�g�t�H���_�̃p�X�A�������͓���t�H���_�������l
'* @return Object(of Folder)
'******************************************************************************
Public Function BrowseForFolder(hWnd As Long, Title As String, options As Long, Optional RootFolder) As Object
    Set BrowseForFolder = Base.GetShell.BrowseForFolder(hWnd, Title, options, RootFolder)
End Function

'******************************************************************************
'* [�T  �v] CanStartStopService ���\�b�h
'* [��  ��] Shell.Application��CanStartStopService �̃��b�p�[���\�b�h�B
'*          ���݂̃��[�U�[�����O�t���T�[�r�X���J�n����ђ�~�ł��邩�ǂ�����
'*          �m�F����B
'*
'* @param ServiceName
'* @return Variant
'******************************************************************************
Public Function CanStartStopService(ServiceName As String)
    CanStartStopService = Base.GetShell.CanStartStopService(ServiceName)
End Function

'******************************************************************************
'* [�T  �v] CascadeWindows ���\�b�h
'* [��  ��] Shell.Application��CascadeWindows �̃��b�p�[���\�b�h�B
'*          �f�X�N�g�b�v��̂��ׂẴE�B���h�E���J�X�P�[�h����B
'*          �i�����̃E�C���h�E�����������炵�Ȃ���d�˂ĕ\������B�j
'*
'******************************************************************************
Public Sub CascadeWindows()
    Call Base.GetShell.CascadeWindows
End Sub

'******************************************************************************
'* [�T  �v] ControlPanelItem ���\�b�h
'* [��  ��] Shell.Application��ControlPanelItem �̃��b�p�[���\�b�h�B
'*          �R���g���[���p�l���̃A�v���P�[�V���������s����B
'* [�Q  �l] https://atmarkit.itmedia.co.jp/ait/articles/0507/02/news016.html
'*
'* @param bstrDir �R���g���[���p�l���̎�ނ�\���t�@�C�����i�g���q.cpl�j
'*                APPWIZ.CPL  �v���O�����Ƌ@�\
'*                DESK.CPL    ��ʂ̉𑜓x
'*                INETCPL.CPL �C���^�[�l�b�g�I�v�V����
'*                SYSDM.CPL   �V�X�e���̃v���p�e�B
'*                MAIN.CPL    �}�E�X�̃v���p�e�B
'*                  ��C:\Windows\System32 �z����.cpl�t�@�C���͔z�u
'******************************************************************************
Public Sub ControlPanelItem(bstrDir As String)
    Call Base.GetShell.ControlPanelItem(bstrDir)
End Sub

'******************************************************************************
'* [�T  �v] Explore ���\�b�h
'* [��  ��] Shell.Application��Explore �̃��b�p�[���\�b�h�B
'*          �f�B���N�g�������w�肵�ăG�N�X�v���[�����N��
'*
'* @param vDir �f�B���N�g���p�X�A�������͓���t�H���_�������l
'******************************************************************************
Public Sub Explore(vDir)
    Call Base.GetShell.Explore(vDir)
End Sub

'******************************************************************************
'* [�T  �v] ExplorerPolicy ���\�b�h
'* [��  ��] Shell.Application��ExplorerPolicy �̃��b�p�[���\�b�h�B
'*          �w�肵��Internet Explorer�|���V�[�̒l���擾����B
'*
'* @param bstrPolicyName
'* @return Variant
'******************************************************************************
Public Function ExplorerPolicy(bstrPolicyName As String)
    ExplorerPolicy = Base.GetShell.ExplorerPolicy(bstrPolicyName)
End Function

'******************************************************************************
'* [�T  �v] FileRun ���\�b�h
'* [��  ��] Shell.Application��FileRun �̃��b�p�[���\�b�h�B
'*          �u�t�@�C�������w�肵�Ď��s�v��\���B
'*
'******************************************************************************
Public Sub FileRun()
    Call Base.GetShell.FileRun
End Sub

'******************************************************************************
'* [�T  �v] FindComputer ���\�b�h
'* [��  ��] Shell.Application��FindComputer �̃��b�p�[���\�b�h�B
'*          �u�R���s���[�^�̌����v�_�C�A���O�{�b�N�X��\���B
'*
'******************************************************************************
Public Sub FindComputer()
    Call Base.GetShell.FindComputer
End Sub

'******************************************************************************
'* [�T  �v] FindFiles ���\�b�h
'* [��  ��] Shell.Application��FindFiles �̃��b�p�[���\�b�h�B
'*          �u�t�@�C���̌����v�_�C�A���O�{�b�N�X��\���B
'*
'******************************************************************************
Public Sub FindFiles()
    Call Base.GetShell.FindFiles
End Sub

'******************************************************************************
'* [�T  �v] FindPrinter ���\�b�h
'* [��  ��] Shell.Application��FindPrinter �̃��b�p�[���\�b�h�B
'*          �u�v�����^�̌����v�_�C�A���O�{�b�N�X��\���B
'*
'* @param Name
'* @param location
'* @param model
'******************************************************************************
Public Sub FindPrinter(Optional Name As String, Optional location As String, Optional model As String)
     Call Base.GetShell.FindPrinter(Name, location, model)
End Sub

'******************************************************************************
'* [�T  �v] GetSetting ���\�b�h
'* [��  ��] Shell.Application��GetSetting �̃��b�p�[���\�b�h�B
'*          �O���[�o���V�F���ݒ���擾�B
'* [�Q  �l] https://docs.microsoft.com/ja-jp/windows/win32/shell/shell-getsetting
'*
'* @param lSetting
'* @return Boolean �ݒ肪���݂���ꍇ�� True
'******************************************************************************
Public Function GetSetting(lSetting As Long) As Boolean
    GetSetting = Base.GetShell.GetSetting(lSetting)
End Function

'******************************************************************************
'* [�T  �v] GetSystemInformation ���\�b�h
'* [��  ��] Shell.Application��GetSystemInformation �̃��b�p�[���\�b�h�B
'*          �V�X�e�������擾�B
'* [�Q  �l] https://docs.microsoft.com/ja-jp/windows/win32/shell/shell-getsysteminformation
'*
'* @param Name �V�X�e�����
'*               DirectoryServiceAvailable �f�B���N�g���T�[�r�X���g�p�\�ȏꍇ��True
'*               DoubleClickTime           �_�u���N���b�N���� (�~���b�P��)
'*               ProcessorLevel            �v���Z�b�T���x��
'*               ProcessorSpeed            �v���Z�b�T���x (���K�w���c (MHz)
'*               ProcessorArchitecture     �v���Z�b�T�A�[�L�e�N�`��
'*               PhysicalMemoryInstalled   �C���X�g�[������Ă��镨���������̗� (�o�C�g�P��)
'* @return Variant �V�X�e�����̒l
'******************************************************************************
Public Function GetSystemInformation(Name As String)
    GetSystemInformation = Base.GetShell.GetSystemInformation(Name)
End Function

'******************************************************************************
'* [�T  �v] Help ���\�b�h
'* [��  ��] Shell.Application��Help �̃��b�p�[���\�b�h�B
'*          Windows�̃w���v�\���B
'*
'******************************************************************************
Public Sub Help()
    Call Base.GetShell.Help
End Sub

'******************************************************************************
'* [�T  �v] IsRestricted ���\�b�h
'* [��  ��] Shell.Application��IsRestricted �̃��b�p�[���\�b�h�B
'*          ���W�X�g������O���[�v�̐����ݒ���擾�B
'*
'* @param Group       �O���[�v�����܂ޕ�����B�������m�F���郌�W�X�g���T�u�L�[�̖��O�B
'* @param Restriction �l���擾����鐧�����i�[���Ă��镶����B
'* @return Long       �����̒l�B �w�肳�ꂽ������������Ȃ��ꍇ�A0
'******************************************************************************
Public Function IsRestricted(Group As String, Restriction As String) As Long
    IsRestricted = Base.GetShell.IsRestricted(Group, Restriction)
End Function

'******************************************************************************
'* [�T  �v] IsServiceRunning ���\�b�h
'* [��  ��] Shell.Application��IsServiceRunning �̃��b�p�[���\�b�h�B
'*          ����̃T�[�r�X�����s����Ă��邩�ǂ����������l��Ԃ��B
'*
'* @param ServiceName �T�[�r�X�̖��O���܂ޕ�����
'* @return Variant(of Boolean) �w�肳�ꂽ�T�[�r�X�����s����Ă���ꍇ�ATrue
'******************************************************************************
Public Function IsServiceRunning(ServiceName As String)
    IsServiceRunning = Base.GetShell.IsServiceRunning(ServiceName)
End Function

'******************************************************************************
'* [�T  �v] MinimizeAll ���\�b�h
'* [��  ��] Shell.Application��MinimizeAll �̃��b�p�[���\�b�h�B
'*         �E�B���h�E�����ׂčŏ�������B
'*
'******************************************************************************
Public Sub MinimizeAll()
    Call Base.GetShell.MinimizeAll
End Sub

'******************************************************************************
'* [�T  �v] NameSpace ���\�b�h
'* [��  ��] Shell.Application��NameSpace �̃��b�p�[���\�b�h�B
'*          �w�肵���p�X��Folder�I�u�W�F�N�g��Ԃ��B
'*
'* @param vDir �f�B���N�g���p�X�A�������͓���t�H���_�������l
'* @return Object(of Folder)
'******************************************************************************
Public Function NameSpace(vDir) As Object
    Set NameSpace = Base.GetShell.NameSpace(vDir)
End Function

'******************************************************************************
'* [�T  �v] Open ���\�b�h
'* [��  ��] Shell.Application��Open �̃��b�p�[���\�b�h�B
'*          �C�ӂ̃t�H���_���}�C�R���s���[�^�`���ŊJ���B
'*
'* @param vDir �f�B���N�g���p�X�A�������͓���t�H���_�������l
'******************************************************************************
Public Sub OpenFolder(vDir)
    Call Base.GetShell.Open(vDir)
End Sub

'******************************************************************************
'* [�T  �v] RefreshMenu ���\�b�h
'* [��  ��] Shell.Application��RefreshMenu �̃��b�p�[���\�b�h�B
'*          �X�^�[�g���j���[���X�V����B
'*
'******************************************************************************
Public Sub RefreshMenu()
    Call Base.GetShell.RefreshMenu
End Sub

'******************************************************************************
'* [�T  �v] SearchCommand ���\�b�h
'* [��  ��] Shell.Application��SearchCommand �̃��b�p�[���\�b�h�B
'*          �u�A�v�������v�E�B���h�E��\���B
'*
'******************************************************************************
Public Sub SearchCommand()
    Call Base.GetShell.SearchCommand
End Sub

'******************************************************************************
'* [�T  �v] ServiceStart ���\�b�h
'* [��  ��] Shell.Application��ServiceStart �̃��b�p�[���\�b�h�B
'*          ���O�t���T�[�r�X���J�n�B
'*
'* @param ServiceName �T�[�r�X�̖��O���܂ޕ�����
'* @param Persistent  �T�[�r�X�������I�ɊJ�n�����悤�ɂ���ꍇ�Atrue
'* @return Variant    ���������ꍇ�� true
'******************************************************************************
Public Function ServiceStart(ServiceName As String, Persistent)
    ServiceStart = Base.GetShell.ServiceStart(ServiceName, Persistent)
End Function

'******************************************************************************
'* [�T  �v] ServiceStop ���\�b�h
'* [��  ��] Shell.Application��ServiceStop �̃��b�p�[���\�b�h�B
'*          ���O�t���T�[�r�X���~�B
'*
'* @param ServiceName �T�[�r�X�̖��O���܂ޕ�����
'* @param Persistent  �T�[�r�X�������I�ɊJ�n�����悤�ɂ���ꍇ�Atrue
'* @return Variant    ���������ꍇ�� true
'******************************************************************************
Public Function ServiceStop(ServiceName As String, Persistent)
    ServiceStop = Base.GetShell.ServiceStop(ServiceName, Persistent)
End Function

'******************************************************************************
'* [�T  �v] SetTime ���\�b�h
'* [��  ��] Shell.Application��SetTime �̃��b�p�[���\�b�h�B
'*          ���t�Ǝ����̃v���p�e�B��\������B
'*
'******************************************************************************
Public Sub SetTime()
    Call Base.GetShell.SetTime
End Sub

'******************************************************************************
'* [�T  �v] ShellExecute ���\�b�h
'* [��  ��] Shell.Application��ShellExecute �̃��b�p�[���\�b�h�B
'*          �w�肳�ꂽ�t�@�C���ɑ΂��Ďw�肳�ꂽ��������s����B
'* [�Q  �l] https://docs.microsoft.com/ja-jp/windows/win32/shell/shell-shellexecute
'*
'* @param sFile      ���s����t�@�C����
'* @param vArgs      ����̃p�����[�^�l
'* @param vDir       ���s����t�@�C���̂���f�B���N�g���̐�΃p�X�B�ȗ����͌��݂̍�ƃf�B���N�g���B
'* @param vOperation ���s���鑀��
'* @param vShow      �A�v���P�[�V�����E�B���h�E���ŏ��ɂǂ̂悤�ɕ\�����邩�w��
'* @return Variant
'******************************************************************************
Public Sub ShellExecute(sFile As String, Optional vArgs, Optional vDir, Optional vOperation, Optional vShow)
    Call Base.GetShell.ShellExecute(sFile, vArgs, vDir, vOperation, vShow)
End Sub

'******************************************************************************
'* [�T  �v] ShutdownWindows ���\�b�h
'* [��  ��] Shell.Application��ShutdownWindows �̃��b�p�[���\�b�h�B
'*          �uWindows�̏I���v�_�C�A���O��\���B
'*
'******************************************************************************
Public Sub ShutdownWindows()
    Call Base.GetShell.ShutdownWindows
End Sub

'******************************************************************************
'* [�T  �v] Suspend ���\�b�h
'* [��  ��] Shell.Application��Suspend �̃��b�p�[���\�b�h�B
'*          �T�X�y���h�i�V�X�e�����ȓd�̓��[�h�őҋ@�j������B
'*
'******************************************************************************
Public Sub Suspend()
    Call Base.GetShell.Suspend
End Sub

'******************************************************************************
'* [�T  �v] TileHorizontally ���\�b�h
'* [��  ��] Shell.Application��TileHorizontally �̃��b�p�[���\�b�h�B
'*          �E�B���h�E�����E�ɕ��ׂĕ\���B
'*
'******************************************************************************
Public Sub TileHorizontally()
    Call Base.GetShell.TileHorizontally
End Sub

'******************************************************************************
'* [�T  �v] TileVertically ���\�b�h
'* [��  ��] Shell.Application��TileVertically �̃��b�p�[���\�b�h�B
'*          �E�B���h�E���㉺�ɕ��ׂĕ\���B
'*
'******************************************************************************
Public Sub TileVertically()
    Call Base.GetShell.TileVertically
End Sub

'******************************************************************************
'* [�T  �v] ToggleDesktop ���\�b�h
'* [��  ��] Shell.Application��ToggleDesktop �̃��b�p�[���\�b�h�B
'*          �f�X�N�g�b�v��\���܂��͔�\���ɂ���B
'*
'******************************************************************************
Public Sub ToggleDesktop()
    Call Base.GetShell.ToggleDesktop
End Sub

'******************************************************************************
'* [�T  �v] TrayProperties ���\�b�h
'* [��  ��] Shell.Application��TrayProperties �̃��b�p�[���\�b�h�B
'*          �u�^�X�N�o�[�� [�X�^�[�g] ���j���[�̃v���p�e�B�v�_�C�A���O�{�b�N�X�\���B
'*
'******************************************************************************
Public Sub TrayProperties()
    Call Base.GetShell.TrayProperties
End Sub

'******************************************************************************
'* [�T  �v] UndoMinimizeALL ���\�b�h
'* [��  ��] Shell.Application��UndoMinimizeALL �̃��b�p�[���\�b�h�B
'*          ���ׂẴf�X�N�g�b�v�E�B���h�E���A�Ō�� MinimizeAll �R�}���h�̑O��
'*          ������Ԃɕ����B
'*
'******************************************************************************
Public Sub UndoMinimizeALL()
    Call Base.GetShell.UndoMinimizeALL
End Sub

'******************************************************************************
'* [�T  �v] Windows ���\�b�h
'* [��  ��] Shell.Application��Windows �̃��b�p�[���\�b�h�B
'*          ���݊J����Ă��邷�ׂẴt�H���_�̃R���N�V����(Windows�I�u�W�F�N�g)
'*          ��Ԃ��B
'*          ���̃I�u�W�F�N�g��Item�v���p�e�B���Ԃ��I�u�W�F�N�g�́A
'*          InternetExplorer�Ɠ������\�b�h�E�v���p�e�B�����B
'*
'* @return Object(of Windows)
'******************************************************************************
Public Function Windows() As Object
    Call Base.GetShell.Windows
End Function

'******************************************************************************
'* [�T  �v] WindowsSecurity ���\�b�h
'* [��  ��] Shell.Application��WindowsSecurity �̃��b�p�[���\�b�h�B
'*          �u�V���� Windows �Z�L�����e�B�v�_�C�A���O�{�b�N�X��\���B
'*
'******************************************************************************
Public Sub WindowsSecurity()
    Call Base.GetShell.WindowsSecurity
End Sub
    
'******************************************************************************
'* [�T  �v] WindowSwitcher ���\�b�h
'* [��  ��] Shell.Application��WindowSwitcher �̃��b�p�[���\�b�h�B
'*          �J���Ă���E�B���h�E��3D�X�^�b�N�ŕ\���B
'*
'******************************************************************************
Public Sub WindowSwitcher()
    Call Base.GetShell.WindowSwitcher
End Sub

'******************************************************************************
'* [�T  �v] ExecWmiQuery ���\�b�h
'* [��  ��] WQL�N�G�������s���AWMI�ň�����I�u�W�F�N�g���擾����B
'*
'* @param wqlQuery WQL�N�G��
'* @return Object(of SWbemObjectSet) WMI�̃I�u�W�F�N�g
'******************************************************************************
Public Function ExecWmiQuery(wqlQuery As String) As Object
    Set ExecWmiQuery = Base.CreateSWbemServices().ExecQuery(wqlQuery)
End Function

'******************************************************************************
'* [�T  �v] GetWmiSWbemObject ���\�b�h
'* [��  ��] WMI�ň�����I�u�W�F�N�g���擾����B
'*
'* @param objectName WMI�̃I�u�W�F�N�g��
'* @return Object(of SWbemObjectSet) WMI�̃I�u�W�F�N�g
'******************************************************************************
Public Function GetWmiSWbemObject(objectName As String) As Object
    Set GetWmiSWbemObject = Base.CreateSWbemServices().InstancesOf(objectName)
End Function

'******************************************************************************
'* [�T  �v] GetWin32Process ���\�b�h
'* [��  ��] WMI�I�u�W�F�N�g�i�v���Z�X�̈ꗗ���j���擾����B
'*
'* @return Object WMI�I�u�W�F�N�g�i�v���Z�X�̈ꗗ���j
'******************************************************************************
Public Function GetWin32Process() As Object
    Set GetWin32Process = GetWmiSWbemObject("Win32_Process")
End Function
'******************************************************************************
'* [�T  �v] GetWin32NetworkAdapterConfiguration ���\�b�h
'* [��  ��] WMI�I�u�W�F�N�g�i�l�b�g���[�N�A�_�v�^�ݒ�̈ꗗ���j���擾����B
'*
'* @return Object WMI�I�u�W�F�N�g�i�l�b�g���[�N�A�_�v�^�ݒ�̈ꗗ���j
'******************************************************************************
Public Function GetWin32NetworkAdapterConfiguration() As Object
    Set GetWin32NetworkAdapterConfiguration = GetWmiSWbemObject("Win32_NetworkAdapterConfiguration")
End Function
'******************************************************************************
'* [�T  �v] GetWin32OperatingSystem ���\�b�h
'* [��  ��] WMI�I�u�W�F�N�g�iOS���j���擾����B
'*
'* @return Object WMI�I�u�W�F�N�g�iOS���j
'******************************************************************************
Public Function GetWin32OperatingSystem() As Object
    Set GetWin32OperatingSystem = GetWmiSWbemObject("Win32_OperatingSystem")
End Function
'******************************************************************************
'* [�T  �v] GetWin32Printer ���\�b�h
'* [��  ��] WMI�I�u�W�F�N�g�i�v�����^�̈ꗗ���j���擾����B
'*
'* @return Object WMI�I�u�W�F�N�g�i�v�����^�̈ꗗ���j
'******************************************************************************
Public Function GetWin32Printer() As Object
    Set GetWin32Printer = GetWmiSWbemObject("Win32_Printer")
End Function
'******************************************************************************
'* [�T  �v] GetWin32ComputerSystem ���\�b�h
'* [��  ��] WMI�I�u�W�F�N�g�iPC�̏��j���擾����B
'*
'* @return Object WMI�I�u�W�F�N�g�iPC�̏��j
'******************************************************************************
Public Function GetWin32ComputerSystem() As Object
    Set GetWin32ComputerSystem = GetWmiSWbemObject("Win32_ComputerSystem")
End Function
'******************************************************************************
'* [�T  �v] GetWin32Processor ���\�b�h
'* [��  ��] WMI�I�u�W�F�N�g�iCPU�̏��j���擾����B
'*
'* @return Object WMI�I�u�W�F�N�g�iCPU�̏��j
'******************************************************************************
Public Function GetWin32Processor() As Object
    Set GetWin32Processor = GetWmiSWbemObject("Win32_Processor")
End Function
'******************************************************************************
'* [�T  �v] GetWin32LogicalDisk ���\�b�h
'* [��  ��] WMI�I�u�W�F�N�g�i�f�B�X�N�̏��j���擾����B
'*
'* @return Object WMI�I�u�W�F�N�g�i�f�B�X�N�̏��j
'******************************************************************************
Public Function GetWin32LogicalDisk() As Object
    Set GetWin32LogicalDisk = GetWmiSWbemObject("Win32_LogicalDisk")
End Function
'******************************************************************************
'* [�T  �v] GetWin32NetworkAdapter ���\�b�h
'* [��  ��] WMI�I�u�W�F�N�g�i�l�b�g���[�N�A�_�v�^�̈ꗗ���j���擾����B
'*
'* @return Object WMI�I�u�W�F�N�g�i�l�b�g���[�N�A�_�v�^�̈ꗗ���j
'******************************************************************************
Public Function GetWin32NetworkAdapter() As Object
    Set GetWin32NetworkAdapter = GetWmiSWbemObject("Win32_NetworkAdapter")
End Function
'******************************************************************************
'* [�T  �v] GetWin32PnPEntity ���\�b�h
'* [��  ��] WMI�I�u�W�F�N�g�i�v���O�A���h�v���C�f�o�C�X�̈ꗗ���j���擾����B
'*
'* @return Object WMI�I�u�W�F�N�g�i�v���O�A���h�v���C�f�o�C�X�̈ꗗ���j
'******************************************************************************
Public Function GetWin32PnPEntity() As Object
    Set GetWin32PnPEntity = GetWmiSWbemObject("Win32_PnPEntity")
End Function
'******************************************************************************
'* [�T  �v] GetWin32Service ���\�b�h
'* [��  ��] WMI�I�u�W�F�N�g�i�T�[�r�X�̈ꗗ���j���擾����B
'*
'* @return Object WMI�I�u�W�F�N�g�i�T�[�r�X�̈ꗗ���j
'******************************************************************************
Public Function GetWin32Service() As Object
    Set GetWin32Service = GetWmiSWbemObject("Win32_Service")
End Function
'******************************************************************************
'* [�T  �v] GetWin32UserAccount ���\�b�h
'* [��  ��] WMI�I�u�W�F�N�g�i���[�U�A�J�E���g�̈ꗗ���j���擾����B
'*
'* @return Object WMI�I�u�W�F�N�g�i���[�U�A�J�E���g�̈ꗗ���j
'******************************************************************************
Public Function GetWin32UserAccount() As Object
    Set GetWin32UserAccount = GetWmiSWbemObject("Win32_UserAccount")
End Function
'******************************************************************************
'* [�T  �v] GetWin32Product ���\�b�h
'* [��  ��] WMI�I�u�W�F�N�g�iMSI�ŃC���X�g�[�����ꂽ���i�ꗗ���j���擾����B
'*
'* @return Object WMI�I�u�W�F�N�g�iMSI�ŃC���X�g�[�����ꂽ���i�ꗗ���j
'******************************************************************************
Public Function GetWin32Product() As Object
    Set GetWin32Product = GetWmiSWbemObject("Win32_Product")
End Function
'******************************************************************************
'* [�T  �v] GetWin32ScheduledJob ���\�b�h
'* [��  ��] WMI�I�u�W�F�N�g�i�X�P�W���[�����ꂽ�W���u�̈ꗗ���j���擾����B
'*
'* @return Object WMI�I�u�W�F�N�g�i�X�P�W���[�����ꂽ�W���u�̈ꗗ���j
'******************************************************************************
Public Function GetWin32ScheduledJob() As Object
    Set GetWin32ScheduledJob = GetWmiSWbemObject("Win32_ScheduledJob")
End Function
'******************************************************************************
'* [�T  �v] GetWin32BIOS ���\�b�h
'* [��  ��] WMI�I�u�W�F�N�g�iBIOS�̏��j���擾����B
'*
'* @return Object WMI�I�u�W�F�N�g�iBIOS�̏��j
'******************************************************************************
Public Function GetWin32BIOS() As Object
    Set GetWin32BIOS = GetWmiSWbemObject("Win32_BIOS")
End Function