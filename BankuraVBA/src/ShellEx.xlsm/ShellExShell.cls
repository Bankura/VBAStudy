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
'* [機能名] Shell32.Shellクラスのラップ・拡張クラス
'* [詳  細] Shell32.ShellクラスのWrapperとしての機能を提供する他、
'*          Shell32.Shellクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' Shell32.Shellクラスについて。
' <xxxxxxxxxxxxxx>
'
' Definition of interface FolderItem
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mOrigin As shell32.Shell
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
    Set mOrigin = CreateObject("Shell.Application")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] AddToRecent メソッド
'* [詳  細] AddToRecent のラッパーメソッド。
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
'* [概  要] BrowseForFolder メソッド
'* [詳  細] BrowseForFolder のラッパーメソッド。
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
'* [概  要] CanStartStopService メソッド
'* [詳  細] CanStartStopService のラッパーメソッド。
'*          Determine if the current user can start/stop the named service.
'*
'* @param ServiceName
'* @return Variant
'******************************************************************************
Function CanStartStopService(ServiceName As String)
    CanStartStopService = mOrigin.CanStartStopService(ServiceName)
End Function

'******************************************************************************
'* [概  要] CascadeWindows メソッド
'* [詳  細] CascadeWindows のラッパーメソッド。
'*          Cascade Windows
'*
'******************************************************************************
Sub CascadeWindows()
    Call mOrigin.CascadeWindows
End Sub

'******************************************************************************
'* [概  要] ControlPanelItem メソッド
'* [詳  細] ControlPanelItem のラッパーメソッド。
'*          Run a Control Panel Item
'*
'* @param bstrDir
'******************************************************************************
Sub ControlPanelItem(bstrDir As String)
    Call mOrigin.ControlPanelItem(bstrDir)
End Sub

'******************************************************************************
'* [概  要] EjectPC メソッド
'* [詳  細] EjectPC のラッパーメソッド。
'*          Eject the pc
'*
'******************************************************************************
Sub EjectPC()
    Call mOrigin.EjectPC
End Sub

'******************************************************************************
'* [概  要] Explore メソッド
'* [詳  細] Explore のラッパーメソッド。
'*          Explore a folder
'*
'* @param vDir
'******************************************************************************
Sub Explore(vDir)
    Call mOrigin.Explore(vDir)
End Sub

'******************************************************************************
'* [概  要] ExplorerPolicy メソッド
'* [詳  細] ExplorerPolicy のラッパーメソッド。
'*          Return explorer policy value
'*
'* @param bstrPolicyName
'* @return Variant
'******************************************************************************
Function ExplorerPolicy(bstrPolicyName As String)
    ExplorerPolicy = mOrigin.ExplorerPolicy(bstrPolicyName)
End Function

'******************************************************************************
'* [概  要] FileRun メソッド
'* [詳  細] FileRun のラッパーメソッド。
'*          Bring up the file run
'*
'******************************************************************************
Sub FileRun()
    Call mOrigin.FileRun
End Sub

'******************************************************************************
'* [概  要] FindComputer メソッド
'* [詳  細] FindComputer のラッパーメソッド。
'*          Find a computer
'*
'******************************************************************************
Sub FindComputer()
    Call mOrigin.FindComputer
End Sub

'******************************************************************************
'* [概  要] FindFiles メソッド
'* [詳  細] FindFiles のラッパーメソッド。
'*          Find Files
'*
'******************************************************************************
Sub FindFiles()
    Call mOrigin.FindFiles
End Sub

'******************************************************************************
'* [概  要] FindPrinter メソッド
'* [詳  細] FindPrinter のラッパーメソッド。
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
'* [概  要] GetSetting メソッド
'* [詳  細] GetSetting のラッパーメソッド。
'*          Return shell global setting
'*
'* @param lSetting
'* @return Boolean
'******************************************************************************
Function GetSetting(lSetting As Long) As Boolean
    GetSetting = mOrigin.GetSetting(lSetting)
End Function

'******************************************************************************
'* [概  要] GetSystemInformation メソッド
'* [詳  細] GetSystemInformation のラッパーメソッド。
'*          Retrieve info about the user's system
'*
'* @param Name
'* @return Variant
'******************************************************************************
Function GetSystemInformation(Name As String)
    GetSystemInformation = mOrigin.GetSystemInformation(Name)
End Function

'******************************************************************************
'* [概  要] Help メソッド
'* [詳  細] Help のラッパーメソッド。
'*          Display shell help
'*
'******************************************************************************
Sub Help()
    Call mOrigin.Help
End Sub

'******************************************************************************
'* [概  要] IsRestricted メソッド
'* [詳  細] IsRestricted のラッパーメソッド。
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
'* [概  要] IsServiceRunning メソッド
'* [詳  細] IsServiceRunning のラッパーメソッド。
'*          Determine if a service is running by name.
'*
'* @param ServiceName
'* @return Variant
'******************************************************************************
Function IsServiceRunning(ServiceName As String)
    IsServiceRunning = mOrigin.IsServiceRunning(ServiceName)
End Function

'******************************************************************************
'* [概  要] MinimizeAll メソッド
'* [詳  細] MinimizeAll のラッパーメソッド。
'*          Minimize all windows
'*
'******************************************************************************
Sub MinimizeAll()
    Call mOrigin.MinimizeAll
End Sub

'******************************************************************************
'* [概  要] NameSpace メソッド
'* [詳  細] NameSpace のラッパーメソッド。
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
'* [概  要] Open メソッド
'* [詳  細] Open のラッパーメソッド。
'*          Open a folder
'*
'* @param vDir
'******************************************************************************
Sub OpenFolder(vDir)
    Call mOrigin.Open(vDir)
End Sub

'******************************************************************************
'* [概  要] RefreshMenu メソッド
'* [詳  細] RefreshMenu のラッパーメソッド。
'*          Refresh the menu
'*
'******************************************************************************
Sub RefreshMenu()
    Call mOrigin.RefreshMenu
End Sub

'******************************************************************************
'* [概  要] SearchCommand メソッド
'* [詳  細] SearchCommand のラッパーメソッド。
'*          Immersive Search
'*
'******************************************************************************
Sub SearchCommand()
    Call mOrigin.SearchCommand
End Sub

'******************************************************************************
'* [概  要] ServiceStart メソッド
'* [詳  細] ServiceStart のラッパーメソッド。
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
'* [概  要] ServiceStop メソッド
'* [詳  細] ServiceStop のラッパーメソッド。
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
'* [概  要] SetTime メソッド
'* [詳  細] SetTime のラッパーメソッド。
'*          Bring up the Set time dialog
'*
'******************************************************************************
Sub SetTime()
    Call mOrigin.SetTime
End Sub

'******************************************************************************
'* [概  要] ShellExecute メソッド
'* [詳  細] ShellExecute のラッパーメソッド。
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
'* [概  要] ShowBrowserBar メソッド
'* [詳  細] ShowBrowserBar のラッパーメソッド。
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
'* [概  要] ShutdownWindows メソッド
'* [詳  細] ShutdownWindows のラッパーメソッド。
'*          Exit Windows
'*
'******************************************************************************
Sub ShutdownWindows()
    Call mOrigin.ShutdownWindows
End Sub

'******************************************************************************
'* [概  要] Suspend メソッド
'* [詳  細] Suspend のラッパーメソッド。
'*          Suspend the pc
'*
'******************************************************************************
Sub Suspend()
    Call mOrigin.Suspend
End Sub

'******************************************************************************
'* [概  要] TileHorizontally メソッド
'* [詳  細] TileHorizontally のラッパーメソッド。
'*          Tile windows horizontally
'*
'******************************************************************************
Sub TileHorizontally()
    Call mOrigin.TileHorizontally
End Sub

'******************************************************************************
'* [概  要] TileVertically メソッド
'* [詳  細] TileVertically のラッパーメソッド。
'*          Tile windows vertically
'*
'******************************************************************************
Sub TileVertically()
    Call mOrigin.TileVertically
End Sub

'******************************************************************************
'* [概  要] ToggleDesktop メソッド
'* [詳  細] ToggleDesktop のラッパーメソッド。
'*          Raise/lower the desktop
'*
'******************************************************************************
Sub ToggleDesktop()
    Call mOrigin.ToggleDesktop
End Sub

'******************************************************************************
'* [概  要] TrayProperties メソッド
'* [詳  細] TrayProperties のラッパーメソッド。
'*          Handle Tray properties
'*
'******************************************************************************
Sub TrayProperties()
    Call mOrigin.TrayProperties
End Sub

'******************************************************************************
'* [概  要] UndoMinimizeALL メソッド
'* [詳  細] UndoMinimizeALL のラッパーメソッド。
'*          Undo Minimize All
'*
'******************************************************************************
Sub UndoMinimizeALL()
    Call mOrigin.UndoMinimizeALL
End Sub

'******************************************************************************
'* [概  要] Windows メソッド
'* [詳  細] Windows のラッパーメソッド。
'*          The collection of open folder windows
'*
'******************************************************************************
Function Windows() As Object
    Call mOrigin.Windows
End Function


'******************************************************************************
'* [概  要] WindowsSecurity メソッド
'* [詳  細] WindowsSecurity のラッパーメソッド。
'*          Windows security
'*
'******************************************************************************
Sub WindowsSecurity()
    Call mOrigin.WindowsSecurity
End Sub
    
'******************************************************************************
'* [概  要] WindowSwitcher メソッド
'* [詳  細] WindowSwitcher のラッパーメソッド。
'*          Window Switcher
'*
'******************************************************************************
Sub WindowSwitcher()
    Call mOrigin.WindowSwitcher
End Sub
