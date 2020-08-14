VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShellExShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] Shell32.Shell�N���X�̃��b�v�E�g���N���X
'* [��  ��] Shell32.Shell�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          Shell32.Shell�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' Shell32.Shell�N���X�ɂ��āB
' <xxxxxxxxxxxxxx>
'
' Definition of interface FolderItem
'
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
Private mOrigin As shell32.Shell
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


'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("Shell.Application")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] AddToRecent ���\�b�h
'* [��  ��] AddToRecent �̃��b�p�[���\�b�h�B
'*          Add an object to the Recent Docuements
'*
'* @param varFile
'* @param bstrCategory
'******************************************************************************
Sub AddToRecent(varFile, Optional bstrCategory As String)
    If IsMissing(bstrCategory) Then
        Call mOrigin.AddToRecent(varFile)
    Else
        Call mOrigin.AddToRecent(varFile, bstrCategory)
    End If
End Sub

'******************************************************************************
'* [�T  �v] BrowseForFolder ���\�b�h
'* [��  ��] BrowseForFolder �̃��b�p�[���\�b�h�B
'*          Browse the name space for a Folder
'*
'* @param Hwnd
'* @param Title
'* @param Options
'* @param RootFolder
'* @return ShellExFolder
'******************************************************************************
Function BrowseForFolder(Hwnd As Long, Title As String, Options As Long, Optional RootFolder) As ShellExFolder
    Dim wrap As ShellExFolder: Set wrap = New ShellExFolder
    If IsMissing(RootFolder) Then
       Set wrap.Origin = mOrigin.BrowseForFolder(Hwnd, Title, Options)
    Else
       Set wrap.Origin = mOrigin.BrowseForFolder(Hwnd, Title, Options, RootFolder)
    End If
    Set BrowseForFolder = wrap
End Function

'******************************************************************************
'* [�T  �v] CanStartStopService ���\�b�h
'* [��  ��] CanStartStopService �̃��b�p�[���\�b�h�B
'*          Determine if the current user can start/stop the named service.
'*
'* @param ServiceName
'* @return Variant
'******************************************************************************
Function CanStartStopService(ServiceName As String)
    CanStartStopService = mOrigin.CanStartStopService(ServiceName)
End Function

'******************************************************************************
'* [�T  �v] CascadeWindows ���\�b�h
'* [��  ��] CascadeWindows �̃��b�p�[���\�b�h�B
'*          Cascade Windows
'*
'******************************************************************************
Sub CascadeWindows()
    Call mOrigin.CascadeWindows
End Sub

'******************************************************************************
'* [�T  �v] ControlPanelItem ���\�b�h
'* [��  ��] ControlPanelItem �̃��b�p�[���\�b�h�B
'*          Run a Control Panel Item
'*
'* @param bstrDir
'******************************************************************************
Sub ControlPanelItem(bstrDir As String)
    Call mOrigin.ControlPanelItem(bstrDir)
End Sub

'******************************************************************************
'* [�T  �v] EjectPC ���\�b�h
'* [��  ��] EjectPC �̃��b�p�[���\�b�h�B
'*          Eject the pc
'*
'******************************************************************************
Sub EjectPC()
    Call mOrigin.EjectPC
End Sub

'******************************************************************************
'* [�T  �v] Explore ���\�b�h
'* [��  ��] Explore �̃��b�p�[���\�b�h�B
'*          Explore a folder
'*
'* @param vDir
'******************************************************************************
Sub Explore(vDir)
    Call mOrigin.Explore(vDir)
End Sub

'******************************************************************************
'* [�T  �v] ExplorerPolicy ���\�b�h
'* [��  ��] ExplorerPolicy �̃��b�p�[���\�b�h�B
'*          Return explorer policy value
'*
'* @param bstrPolicyName
'* @return Variant
'******************************************************************************
Function ExplorerPolicy(bstrPolicyName As String)
    ExplorerPolicy = mOrigin.ExplorerPolicy(bstrPolicyName)
End Function

'******************************************************************************
'* [�T  �v] FileRun ���\�b�h
'* [��  ��] FileRun �̃��b�p�[���\�b�h�B
'*          Bring up the file run
'*
'******************************************************************************
Sub FileRun()
    Call mOrigin.FileRun
End Sub

'******************************************************************************
'* [�T  �v] FindComputer ���\�b�h
'* [��  ��] FindComputer �̃��b�p�[���\�b�h�B
'*          Find a computer
'*
'******************************************************************************
Sub FindComputer()
    Call mOrigin.FindComputer
End Sub

'******************************************************************************
'* [�T  �v] FindFiles ���\�b�h
'* [��  ��] FindFiles �̃��b�p�[���\�b�h�B
'*          Find Files
'*
'******************************************************************************
Sub FindFiles()
    Call mOrigin.FindFiles
End Sub

'******************************************************************************
'* [�T  �v] FindPrinter ���\�b�h
'* [��  ��] FindPrinter �̃��b�p�[���\�b�h�B
'*          Find a Printer in the Directory Service
'*
'* @param Name
'* @param location
'* @param model
'******************************************************************************
Sub FindPrinter(Optional Name As String, Optional location As String, Optional model As String)
    If IsMissing(Name) Then
        If IsMissing(location) Then
            If IsMissing(model) Then
                Call mOrigin.FindPrinter
            Else
                Call mOrigin.FindPrinter(, , model)
            End If
        ElseIf IsMissing(model) Then
            Call mOrigin.FindPrinter(, location)
        Else
            Call mOrigin.FindPrinter(, location, model)
        End If
    Else
        If IsMissing(location) Then
             If IsMissing(model) Then
                Call mOrigin.FindPrinter(Name)
            Else
                Call mOrigin.FindPrinter(Name, , model)
            End If
        ElseIf IsMissing(model) Then
            Call mOrigin.FindPrinter(Name, location)
        Else
            Call mOrigin.FindPrinter(Name, location, model)
        End If
    End If
End Sub

'******************************************************************************
'* [�T  �v] GetSetting ���\�b�h
'* [��  ��] GetSetting �̃��b�p�[���\�b�h�B
'*          Return shell global setting
'*
'* @param lSetting
'* @return Boolean
'******************************************************************************
Function GetSetting(lSetting As Long) As Boolean
    GetSetting = mOrigin.GetSetting(lSetting)
End Function

'******************************************************************************
'* [�T  �v] GetSystemInformation ���\�b�h
'* [��  ��] GetSystemInformation �̃��b�p�[���\�b�h�B
'*          Retrieve info about the user's system
'*
'* @param Name
'* @return Variant
'******************************************************************************
Function GetSystemInformation(Name As String)
    GetSystemInformation = mOrigin.GetSystemInformation(Name)
End Function

'******************************************************************************
'* [�T  �v] Help ���\�b�h
'* [��  ��] Help �̃��b�p�[���\�b�h�B
'*          Display shell help
'*
'******************************************************************************
Sub Help()
    Call mOrigin.Help
End Sub

'******************************************************************************
'* [�T  �v] IsRestricted ���\�b�h
'* [��  ��] IsRestricted �̃��b�p�[���\�b�h�B
'*          get restriction settings
'*
'* @param Group
'* @param Restriction
'* @return Long
'******************************************************************************
Function IsRestricted(Group As String, Restriction As String) As Long
    IsRestricted = mOrigin.IsRestricted(Group, Restriction)
End Function

'******************************************************************************
'* [�T  �v] IsServiceRunning ���\�b�h
'* [��  ��] IsServiceRunning �̃��b�p�[���\�b�h�B
'*          Determine if a service is running by name.
'*
'* @param ServiceName
'* @return Variant
'******************************************************************************
Function IsServiceRunning(ServiceName As String)
    IsServiceRunning = mOrigin.IsServiceRunning(ServiceName)
End Function

'******************************************************************************
'* [�T  �v] MinimizeAll ���\�b�h
'* [��  ��] MinimizeAll �̃��b�p�[���\�b�h�B
'*          Minimize all windows
'*
'******************************************************************************
Sub MinimizeAll()
    Call mOrigin.MinimizeAll
End Sub

'******************************************************************************
'* [�T  �v] NameSpace ���\�b�h
'* [��  ��] NameSpace �̃��b�p�[���\�b�h�B
'*          Get special folder from ShellSpecialFolderConstants
'*
'* @param vDir
'* @return ShellExFolder
'******************************************************************************
Function NameSpace(vDir) As ShellExFolder
    Dim wrap As ShellExFolder: Set wrap = New ShellExFolder
    Set NameSpace = mOrigin.NameSpace(vDir)
End Function

'******************************************************************************
'* [�T  �v] Open ���\�b�h
'* [��  ��] Open �̃��b�p�[���\�b�h�B
'*          Open a folder
'*
'* @param vDir
'******************************************************************************
Sub OpenFolder(vDir)
    Call mOrigin.OpenFolder(vDir)
End Sub

'******************************************************************************
'* [�T  �v] RefreshMenu ���\�b�h
'* [��  ��] RefreshMenu �̃��b�p�[���\�b�h�B
'*          Refresh the menu
'*
'******************************************************************************
Sub RefreshMenu()
    Call mOrigin.RefreshMenu
End Sub

'******************************************************************************
'* [�T  �v] SearchCommand ���\�b�h
'* [��  ��] SearchCommand �̃��b�p�[���\�b�h�B
'*          Immersive Search
'*
'******************************************************************************
Sub SearchCommand()
    Call mOrigin.SearchCommand
End Sub

'******************************************************************************
'* [�T  �v] ServiceStart ���\�b�h
'* [��  ��] ServiceStart �̃��b�p�[���\�b�h�B
'*          Start a service by name, and optionally set it to autostart.
'*
'* @param ServiceName
'* @param Persistent
'* @return Variant
'******************************************************************************
Function ServiceStart(ServiceName As String, Persistent)
    ServiceStart = mOrigin.ServiceStart(ServiceName, Persistent)
End Function

'******************************************************************************
'* [�T  �v] ServiceStop ���\�b�h
'* [��  ��] ServiceStop �̃��b�p�[���\�b�h�B
'*          Stop a service by name, and optionally disable autostart.
'*
'* @param ServiceName
'* @param Persistent
'* @return Variant
'******************************************************************************
Function ServiceStop(ServiceName As String, Persistent)
    ServiceStop = mOrigin.ServiceStop(ServiceName, Persistent)
End Function

'******************************************************************************
'* [�T  �v] SetTime ���\�b�h
'* [��  ��] SetTime �̃��b�p�[���\�b�h�B
'*          Bring up the Set time dialog
'*
'******************************************************************************
Sub SetTime()
    Call mOrigin.SetTime
End Sub

'******************************************************************************
'* [�T  �v] ShellExecute ���\�b�h
'* [��  ��] ShellExecute �̃��b�p�[���\�b�h�B
'*          Execute generic command
'*
'* @param File
'* @param vArgs
'* @param vDir
'* @param vOperation
'* @param vShow
'* @return Variant
'******************************************************************************
Sub ShellExecute(File As String, Optional vArgs, Optional vDir, Optional vOperation, Optional vShow)
    If IsMissing(vArgs) Then
        If IsMissing(vDir) Then
            If IsMissing(vOperation) Then
                If IsMissing(vShow) Then
                    Call mOrigin.ShellExecute(File)
                Else
                    Call mOrigin.ShellExecute(File, , , , vShow)
                End If
            ElseIf IsMissing(vShow) Then
                Call mOrigin.ShellExecute(File, , , vOperation)
            Else
                Call mOrigin.ShellExecute(File, , , vOperation, vShow)
            End If
        Else
            If IsMissing(vOperation) Then
                 If IsMissing(vShow) Then
                    Call mOrigin.ShellExecute(File, , vDir)
                Else
                    Call mOrigin.ShellExecute(File, , vDir, , vShow)
                End If
            ElseIf IsMissing(vShow) Then
                Call mOrigin.ShellExecute(File, , vDir, vOperation)
            Else
                Call mOrigin.ShellExecute(File, , vDir, vOperation, vShow)
            End If
        End If
    Else
        If IsMissing(vDir) Then
            If IsMissing(vOperation) Then
                If IsMissing(vShow) Then
                    Call mOrigin.ShellExecute(File, vArgs)
                Else
                    Call mOrigin.ShellExecute(File, vArgs, , , vShow)
                End If
            ElseIf IsMissing(vShow) Then
                Call mOrigin.ShellExecute(File, vArgs, , vOperation)
            Else
                Call mOrigin.ShellExecute(File, vArgs, , vOperation, vShow)
            End If
        Else
            If IsMissing(vOperation) Then
                 If IsMissing(vShow) Then
                    Call mOrigin.ShellExecute(File, vArgs, vDir)
                Else
                    Call mOrigin.ShellExecute(File, vArgs, vDir, , vShow)
                End If
            ElseIf IsMissing(vShow) Then
                Call mOrigin.ShellExecute(File, vArgs, vDir, vOperation)
            Else
                Call mOrigin.ShellExecute(File, vArgs, vDir, vOperation, vShow)
            End If
        End If
    End If
End Sub

'******************************************************************************
'* [�T  �v] ShowBrowserBar ���\�b�h
'* [��  ��] ShowBrowserBar �̃��b�p�[���\�b�h�B
'*          Show/Hide browser bar.
'*
'* @param bstrClsid
'* @param bShow
'* @return Variant
'******************************************************************************
Function ShowBrowserBar(bstrClsid As String, bShow)
    ShowBrowserBar = mOrigin.ShowBrowserBar(bstrClsid, bShow)
End Function

'******************************************************************************
'* [�T  �v] ShutdownWindows ���\�b�h
'* [��  ��] ShutdownWindows �̃��b�p�[���\�b�h�B
'*          Exit Windows
'*
'******************************************************************************
Sub ShutdownWindows()
    Call mOrigin.ShutdownWindows
End Sub

'******************************************************************************
'* [�T  �v] Suspend ���\�b�h
'* [��  ��] Suspend �̃��b�p�[���\�b�h�B
'*          Suspend the pc
'*
'******************************************************************************
Sub Suspend()
    Call mOrigin.Suspend
End Sub

'******************************************************************************
'* [�T  �v] TileHorizontally ���\�b�h
'* [��  ��] TileHorizontally �̃��b�p�[���\�b�h�B
'*          Tile windows horizontally
'*
'******************************************************************************
Sub TileHorizontally()
    Call mOrigin.TileHorizontally
End Sub

'******************************************************************************
'* [�T  �v] TileVertically ���\�b�h
'* [��  ��] TileVertically �̃��b�p�[���\�b�h�B
'*          Tile windows vertically
'*
'******************************************************************************
Sub TileVertically()
    Call mOrigin.TileVertically
End Sub

'******************************************************************************
'* [�T  �v] ToggleDesktop ���\�b�h
'* [��  ��] ToggleDesktop �̃��b�p�[���\�b�h�B
'*          Raise/lower the desktop
'*
'******************************************************************************
Sub ToggleDesktop()
    Call mOrigin.ToggleDesktop
End Sub

'******************************************************************************
'* [�T  �v] TrayProperties ���\�b�h
'* [��  ��] TrayProperties �̃��b�p�[���\�b�h�B
'*          Handle Tray properties
'*
'******************************************************************************
Sub TrayProperties()
    Call mOrigin.TrayProperties
End Sub

'******************************************************************************
'* [�T  �v] UndoMinimizeALL ���\�b�h
'* [��  ��] UndoMinimizeALL �̃��b�p�[���\�b�h�B
'*          Undo Minimize All
'*
'******************************************************************************
Sub UndoMinimizeALL()
    Call mOrigin.UndoMinimizeALL
End Sub

'******************************************************************************
'* [�T  �v] Windows ���\�b�h
'* [��  ��] Windows �̃��b�p�[���\�b�h�B
'*          The collection of open folder windows
'*
'******************************************************************************
Function Windows() As Object
    Call mOrigin.Windows
End Function


'******************************************************************************
'* [�T  �v] WindowsSecurity ���\�b�h
'* [��  ��] WindowsSecurity �̃��b�p�[���\�b�h�B
'*          Windows security
'*
'******************************************************************************
Sub WindowsSecurity()
    Call mOrigin.WindowsSecurity
End Sub
    
'******************************************************************************
'* [�T  �v] WindowSwitcher ���\�b�h
'* [��  ��] WindowSwitcher �̃��b�p�[���\�b�h�B
'*          Window Switcher
'*
'******************************************************************************
Sub WindowSwitcher()
    Call mOrigin.WindowSwitcher
End Sub