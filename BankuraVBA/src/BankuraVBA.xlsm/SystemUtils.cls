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
'* [機能名] Systemユーティリティクラス
'* [詳  細] システム全般に関するユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'* [参  考]
'*
'* @author Bankura
'* Copyright (c) 2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

'******************************************************************************
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 定数定義
'******************************************************************************

'******************************************************************************
'* 内部変数定義
'******************************************************************************

'******************************************************************************
'* プロパティ定義
'******************************************************************************


'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
End Sub
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] GetScreenResolution
'* [詳  細] 画面解像度を取得する｡
'*
'* @return Variant 画面解像度（1次元配列）0: 横幅 1:高さ
'*
'******************************************************************************
Public Function GetScreenResolution() As Variant
    Dim x As Long, y As Long
    x = GetSystemMetrics(0) ' 画面の幅を取得
    y = GetSystemMetrics(1) ' 画面の高さを取得
    GetScreenResolution = Array(x, y)
End Function

'******************************************************************************
'* [概  要] DisableF1Key
'* [詳  細] F1のヘルプ起動を無効化する｡
'*
'* @param willUndo 無効化をもとに戻すか
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
'* [概  要] DisableEscKey
'* [詳  細] ESCキーによるプログラム実行中断を無効化する｡
'*
'* @param willUndo 無効化をもとに戻すか
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
'* [概  要] GetShellFolderItem
'* [詳  細] 指定したフォルダ/ファイルパスのShellFolderItemオブジェクトを取得する｡
'*
'* @param fPath フォルダ/ファイルパス
'* @return ShellFolderItemオブジェクト
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
'* [概  要] DoFolderItemVerb
'* [詳  細] 指定したフォルダ/ファイルパスのShellFolderItemオブジェクトを取得
'*          し、指定した動詞（Verb）を実行する｡
'*
'* @param fPath フォルダ/ファイルパス
'* @return ShellFolderItemオブジェクト
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
'* [概  要] CopyFromFolderItem
'* [詳  細] 指定したファイル/フォルダをコピーする｡
'*          （エクスプローラのコピーと同じ処理を行い、コピーバッファーに入れる）
'*
'* @param folderPath フォルダパス
'******************************************************************************
Public Sub CopyFromFolderItem(fPath As String)
    Call DoFolderItemVerb(fPath, "コピー(&C)")
End Sub

'******************************************************************************
'* [概  要] CutFromFolderItem
'* [詳  細] 指定したファイル/フォルダを切り取りする｡
'*          （エクスプローラの切り取りと同じ処理を行い、コピーバッファーに入れる）
'*
'* @param folderPath フォルダパス
'******************************************************************************
Public Sub CutFromFolderItem(fPath As String)
    Call DoFolderItemVerb(fPath, "切り取り(&T)")
End Sub

'******************************************************************************
'* [概  要] PasteInFolder
'* [詳  細] 指定したフォルダにコピー中のファイル/フォルダを貼り付ける｡
'*
'* @param folderPath フォルダパス
'******************************************************************************
Public Sub PasteInFolder(folderPath As String)
    If Not IO.FolderExists(folderPath) Then
        Err.Raise 9999, "SystemtUtils#PasteInFolder", "Folder not found. [" & folderPath & "]"
    End If

    Call DoFolderItemVerb(folderPath, "貼り付け(&P)")
End Sub

'******************************************************************************
'* [概  要] ShowFolderItemVerbs
'* [詳  細] FolderItemオブジェクトが持つ動詞（Verb）の一覧をイミディエイト
'*          ウィンドウに表示する｡
'*
'* @param folderItemObj Shell32.FolderItemオブジェクト
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
'* [概  要] AddToRecent メソッド
'* [詳  細] Shell.ApplicationのAddToRecent のラッパーメソッド。
'*          最近使用した（MRU）リストにファイルを追加する。
'*
'* @param varFile
'* @param bstrCategory
'******************************************************************************
Public Sub AddToRecent(varFile, Optional bstrCategory As String)
    Call Base.GetShell.AddToRecent(varFile, bstrCategory)
End Sub

'******************************************************************************
'* [概  要] BrowseForFolder メソッド
'* [詳  細] Shell.ApplicationのBrowseForFolder のラッパーメソッド。
'*          ユーザが選択したフォルダのFolderオブジェクトを返すダイアログ
'*          ボックスを作成する。
'* [参  考] http://www.roy.hi-ho.ne.jp/mutaguchi/wsh/object/shell.htm
'*
'* @param Hwnd       ウィンドウのハンドル。通常は0にしておいて問題なし
'* @param Title      ダイアログに表示させる説明文
'* @param Options    ダイアログのタイプ。複数のオプションを指定する時は、定数の和を指定
'* @param RootFolder ルートフォルダのパス、もしくは特殊フォルダを示す値
'* @return Object(of Folder)
'******************************************************************************
Public Function BrowseForFolder(hWnd As Long, Title As String, options As Long, Optional rootFolder) As Object
    Set BrowseForFolder = Base.GetShell.BrowseForFolder(hWnd, Title, options, rootFolder)
End Function

'******************************************************************************
'* [概  要] CanStartStopService メソッド
'* [詳  細] Shell.ApplicationのCanStartStopService のラッパーメソッド。
'*          現在のユーザーが名前付きサービスを開始および停止できるかどうかを
'*          確認する。
'*
'* @param ServiceName
'* @return Variant
'******************************************************************************
Public Function CanStartStopService(ServiceName As String)
    CanStartStopService = Base.GetShell.CanStartStopService(ServiceName)
End Function

'******************************************************************************
'* [概  要] CascadeWindows メソッド
'* [詳  細] Shell.ApplicationのCascadeWindows のラッパーメソッド。
'*          デスクトップ上のすべてのウィンドウをカスケードする。
'*          （複数のウインドウを少しずつずらしながら重ねて表示する。）
'*
'******************************************************************************
Public Sub CascadeWindows()
    Call Base.GetShell.CascadeWindows
End Sub

'******************************************************************************
'* [概  要] ControlPanelItem メソッド
'* [詳  細] Shell.ApplicationのControlPanelItem のラッパーメソッド。
'*          コントロールパネルのアプリケーションを実行する。
'* [参  考] https://atmarkit.itmedia.co.jp/ait/articles/0507/02/news016.html
'*
'* @param bstrDir コントロールパネルの種類を表すファイル名（拡張子.cpl）
'*                APPWIZ.CPL  プログラムと機能
'*                DESK.CPL    画面の解像度
'*                INETCPL.CPL インターネットオプション
'*                SYSDM.CPL   システムのプロパティ
'*                MAIN.CPL    マウスのプロパティ
'*                  ※C:\Windows\System32 配下に.cplファイルは配置
'******************************************************************************
Public Sub ControlPanelItem(bstrDir As String)
    Call Base.GetShell.ControlPanelItem(bstrDir)
End Sub

'******************************************************************************
'* [概  要] Explore メソッド
'* [詳  細] Shell.ApplicationのExplore のラッパーメソッド。
'*          ディレクトリ名を指定してエクスプローラを起動
'*
'* @param vDir ディレクトリパス、もしくは特殊フォルダを示す値
'******************************************************************************
Public Sub Explore(vDir)
    Call Base.GetShell.Explore(vDir)
End Sub

'******************************************************************************
'* [概  要] ExplorerPolicy メソッド
'* [詳  細] Shell.ApplicationのExplorerPolicy のラッパーメソッド。
'*          指定したInternet Explorerポリシーの値を取得する。
'*
'* @param bstrPolicyName
'* @return Variant
'******************************************************************************
Public Function ExplorerPolicy(bstrPolicyName As String)
    ExplorerPolicy = Base.GetShell.ExplorerPolicy(bstrPolicyName)
End Function

'******************************************************************************
'* [概  要] FileRun メソッド
'* [詳  細] Shell.ApplicationのFileRun のラッパーメソッド。
'*          「ファイル名を指定して実行」を表示。
'*
'******************************************************************************
Public Sub FileRun()
    Call Base.GetShell.FileRun
End Sub

'******************************************************************************
'* [概  要] FindComputer メソッド
'* [詳  細] Shell.ApplicationのFindComputer のラッパーメソッド。
'*          「コンピュータの検索」ダイアログボックスを表示。
'*
'******************************************************************************
Public Sub FindComputer()
    Call Base.GetShell.FindComputer
End Sub

'******************************************************************************
'* [概  要] FindFiles メソッド
'* [詳  細] Shell.ApplicationのFindFiles のラッパーメソッド。
'*          「ファイルの検索」ダイアログボックスを表示。
'*
'******************************************************************************
Public Sub FindFiles()
    Call Base.GetShell.FindFiles
End Sub

'******************************************************************************
'* [概  要] FindPrinter メソッド
'* [詳  細] Shell.ApplicationのFindPrinter のラッパーメソッド。
'*          「プリンタの検索」ダイアログボックスを表示。
'*
'* @param Name
'* @param location
'* @param model
'******************************************************************************
Public Sub FindPrinter(Optional Name As String, Optional location As String, Optional model As String)
     Call Base.GetShell.FindPrinter(Name, location, model)
End Sub

'******************************************************************************
'* [概  要] GetSetting メソッド
'* [詳  細] Shell.ApplicationのGetSetting のラッパーメソッド。
'*          グローバルシェル設定を取得。
'* [参  考] https://docs.microsoft.com/ja-jp/windows/win32/shell/shell-getsetting
'*
'* @param lSetting
'* @return Boolean 設定が存在する場合は True
'******************************************************************************
Public Function GetSetting(lSetting As Long) As Boolean
    GetSetting = Base.GetShell.GetSetting(lSetting)
End Function

'******************************************************************************
'* [概  要] GetSystemInformation メソッド
'* [詳  細] Shell.ApplicationのGetSystemInformation のラッパーメソッド。
'*          システム情報を取得。
'* [参  考] https://docs.microsoft.com/ja-jp/windows/win32/shell/shell-getsysteminformation
'*
'* @param Name システム情報名
'*               DirectoryServiceAvailable ディレクトリサービスが使用可能な場合はTrue
'*               DoubleClickTime           ダブルクリック時間 (ミリ秒単位)
'*               ProcessorLevel            プロセッサレベル
'*               ProcessorSpeed            プロセッサ速度 (メガヘルツ (MHz)
'*               ProcessorArchitecture     プロセッサアーキテクチャ
'*               PhysicalMemoryInstalled   インストールされている物理メモリの量 (バイト単位)
'* @return Variant システム情報の値
'******************************************************************************
Public Function GetSystemInformation(Name As String)
    GetSystemInformation = Base.GetShell.GetSystemInformation(Name)
End Function

'******************************************************************************
'* [概  要] Help メソッド
'* [詳  細] Shell.ApplicationのHelp のラッパーメソッド。
'*          Windowsのヘルプ表示。
'*
'******************************************************************************
Public Sub Help()
    Call Base.GetShell.Help
End Sub

'******************************************************************************
'* [概  要] IsRestricted メソッド
'* [詳  細] Shell.ApplicationのIsRestricted のラッパーメソッド。
'*          レジストリからグループの制限設定を取得。
'*
'* @param Group       グループ名を含む文字列。制限を確認するレジストリサブキーの名前。
'* @param Restriction 値が取得される制限を格納している文字列。
'* @return Long       制限の値。 指定された制限が見つからない場合、0
'******************************************************************************
Public Function IsRestricted(Group As String, Restriction As String) As Long
    IsRestricted = Base.GetShell.IsRestricted(Group, Restriction)
End Function

'******************************************************************************
'* [概  要] IsServiceRunning メソッド
'* [詳  細] Shell.ApplicationのIsServiceRunning のラッパーメソッド。
'*          特定のサービスが実行されているかどうかを示す値を返す。
'*
'* @param ServiceName サービスの名前を含む文字列
'* @return Variant(of Boolean) 指定されたサービスが実行されている場合、True
'******************************************************************************
Public Function IsServiceRunning(ServiceName As String)
    IsServiceRunning = Base.GetShell.IsServiceRunning(ServiceName)
End Function

'******************************************************************************
'* [概  要] MinimizeAll メソッド
'* [詳  細] Shell.ApplicationのMinimizeAll のラッパーメソッド。
'*         ウィンドウをすべて最小化する。
'*
'******************************************************************************
Public Sub MinimizeAll()
    Call Base.GetShell.MinimizeAll
End Sub

'******************************************************************************
'* [概  要] NameSpace メソッド
'* [詳  細] Shell.ApplicationのNameSpace のラッパーメソッド。
'*          指定したパスのFolderオブジェクトを返す。
'*
'* @param vDir ディレクトリパス、もしくは特殊フォルダを示す値
'* @return Object(of Folder)
'******************************************************************************
Public Function NameSpace(vDir) As Object
    Set NameSpace = Base.GetShell.NameSpace(vDir)
End Function

'******************************************************************************
'* [概  要] Open メソッド
'* [詳  細] Shell.ApplicationのOpen のラッパーメソッド。
'*          任意のフォルダをマイコンピュータ形式で開く。
'*
'* @param vDir ディレクトリパス、もしくは特殊フォルダを示す値
'******************************************************************************
Public Sub OpenFolder(vDir)
    Call Base.GetShell.Open(vDir)
End Sub

'******************************************************************************
'* [概  要] RefreshMenu メソッド
'* [詳  細] Shell.ApplicationのRefreshMenu のラッパーメソッド。
'*          スタートメニューを更新する。
'*
'******************************************************************************
Public Sub RefreshMenu()
    Call Base.GetShell.RefreshMenu
End Sub

'******************************************************************************
'* [概  要] SearchCommand メソッド
'* [詳  細] Shell.ApplicationのSearchCommand のラッパーメソッド。
'*          「アプリ検索」ウィンドウを表示。
'*
'******************************************************************************
Public Sub SearchCommand()
    Call Base.GetShell.SearchCommand
End Sub

'******************************************************************************
'* [概  要] ServiceStart メソッド
'* [詳  細] Shell.ApplicationのServiceStart のラッパーメソッド。
'*          名前付きサービスを開始。
'*
'* @param ServiceName サービスの名前を含む文字列
'* @param Persistent  サービスが自動的に開始されるようにする場合、true
'* @return Variant    成功した場合は true
'******************************************************************************
Public Function ServiceStart(ServiceName As String, Persistent)
    ServiceStart = Base.GetShell.ServiceStart(ServiceName, Persistent)
End Function

'******************************************************************************
'* [概  要] ServiceStop メソッド
'* [詳  細] Shell.ApplicationのServiceStop のラッパーメソッド。
'*          名前付きサービスを停止。
'*
'* @param ServiceName サービスの名前を含む文字列
'* @param Persistent  サービスが自動的に開始されるようにする場合、true
'* @return Variant    成功した場合は true
'******************************************************************************
Public Function ServiceStop(ServiceName As String, Persistent)
    ServiceStop = Base.GetShell.ServiceStop(ServiceName, Persistent)
End Function

'******************************************************************************
'* [概  要] SetTime メソッド
'* [詳  細] Shell.ApplicationのSetTime のラッパーメソッド。
'*          日付と時刻のプロパティを表示する。
'*
'******************************************************************************
Public Sub SetTime()
    Call Base.GetShell.SetTime
End Sub

'******************************************************************************
'* [概  要] ShellExecute メソッド
'* [詳  細] Shell.ApplicationのShellExecute のラッパーメソッド。
'*          指定されたファイルに対して指定された操作を実行する。
'* [参  考] https://docs.microsoft.com/ja-jp/windows/win32/shell/shell-shellexecute
'*
'* @param sFile      実行するファイル名
'* @param vArgs      操作のパラメータ値
'* @param vDir       実行するファイルのあるディレクトリの絶対パス。省略時は現在の作業ディレクトリ。
'* @param vOperation 実行する操作
'* @param vShow      アプリケーションウィンドウを最初にどのように表示するか指定
'* @return Variant
'******************************************************************************
Public Sub ShellExecute(sFile As String, Optional vArgs, Optional vDir, Optional vOperation, Optional vShow)
    Call Base.GetShell.ShellExecute(sFile, vArgs, vDir, vOperation, vShow)
End Sub

'******************************************************************************
'* [概  要] ShutdownWindows メソッド
'* [詳  細] Shell.ApplicationのShutdownWindows のラッパーメソッド。
'*          「Windowsの終了」ダイアログを表示。
'*
'******************************************************************************
Public Sub ShutdownWindows()
    Call Base.GetShell.ShutdownWindows
End Sub

'******************************************************************************
'* [概  要] Suspend メソッド
'* [詳  細] Shell.ApplicationのSuspend のラッパーメソッド。
'*          サスペンド（システムを省電力モードで待機）させる。
'*
'******************************************************************************
Public Sub Suspend()
    Call Base.GetShell.Suspend
End Sub

'******************************************************************************
'* [概  要] TileHorizontally メソッド
'* [詳  細] Shell.ApplicationのTileHorizontally のラッパーメソッド。
'*          ウィンドウを左右に並べて表示。
'*
'******************************************************************************
Public Sub TileHorizontally()
    Call Base.GetShell.TileHorizontally
End Sub

'******************************************************************************
'* [概  要] TileVertically メソッド
'* [詳  細] Shell.ApplicationのTileVertically のラッパーメソッド。
'*          ウィンドウを上下に並べて表示。
'*
'******************************************************************************
Public Sub TileVertically()
    Call Base.GetShell.TileVertically
End Sub

'******************************************************************************
'* [概  要] ToggleDesktop メソッド
'* [詳  細] Shell.ApplicationのToggleDesktop のラッパーメソッド。
'*          デスクトップを表示または非表示にする。
'*
'******************************************************************************
Public Sub ToggleDesktop()
    Call Base.GetShell.ToggleDesktop
End Sub

'******************************************************************************
'* [概  要] TrayProperties メソッド
'* [詳  細] Shell.ApplicationのTrayProperties のラッパーメソッド。
'*          「タスクバーと [スタート] メニューのプロパティ」ダイアログボックス表示。
'*
'******************************************************************************
Public Sub TrayProperties()
    Call Base.GetShell.TrayProperties
End Sub

'******************************************************************************
'* [概  要] UndoMinimizeALL メソッド
'* [詳  細] Shell.ApplicationのUndoMinimizeALL のラッパーメソッド。
'*          すべてのデスクトップウィンドウを、最後の MinimizeAll コマンドの前と
'*          同じ状態に復元。
'*
'******************************************************************************
Public Sub UndoMinimizeALL()
    Call Base.GetShell.UndoMinimizeALL
End Sub

'******************************************************************************
'* [概  要] Windows メソッド
'* [詳  細] Shell.ApplicationのWindows のラッパーメソッド。
'*          現在開かれているすべてのフォルダのコレクション(Windowsオブジェクト)
'*          を返す。
'*          このオブジェクトのItemプロパティが返すオブジェクトは、
'*          InternetExplorerと同じメソッド・プロパティを持つ。
'*
'* @return Object(of Windows)
'******************************************************************************
Public Function Windows() As Object
    Call Base.GetShell.Windows
End Function

'******************************************************************************
'* [概  要] WindowsSecurity メソッド
'* [詳  細] Shell.ApplicationのWindowsSecurity のラッパーメソッド。
'*          「新しい Windows セキュリティ」ダイアログボックスを表示。
'*
'******************************************************************************
Public Sub WindowsSecurity()
    Call Base.GetShell.WindowsSecurity
End Sub
    
'******************************************************************************
'* [概  要] WindowSwitcher メソッド
'* [詳  細] Shell.ApplicationのWindowSwitcher のラッパーメソッド。
'*          開いているウィンドウを3Dスタックで表示。
'*
'******************************************************************************
Public Sub WindowSwitcher()
    Call Base.GetShell.WindowSwitcher
End Sub

'******************************************************************************
'* [概  要] ExecWmiQuery メソッド
'* [詳  細] WQLクエリを実行し、WMIで扱えるオブジェクトを取得する。
'*
'* @param wqlQuery WQLクエリ
'* @return Object(of SWbemObjectSet) WMIのオブジェクト
'******************************************************************************
Public Function ExecWmiQuery(wqlQuery As String) As Object
    Set ExecWmiQuery = Base.CreateSWbemServices().ExecQuery(wqlQuery)
End Function

'******************************************************************************
'* [概  要] GetWmiSWbemObject メソッド
'* [詳  細] WMIで扱えるオブジェクトを取得する。
'*
'* @param objectName WMIのオブジェクト名
'* @return Object(of SWbemObjectSet) WMIのオブジェクト
'******************************************************************************
Public Function GetWmiSWbemObject(objectName As String) As Object
    Set GetWmiSWbemObject = Base.CreateSWbemServices().InstancesOf(objectName)
End Function

'******************************************************************************
'* [概  要] GetWin32Process メソッド
'* [詳  細] WMIオブジェクト（プロセスの一覧情報）を取得する。
'*
'* @return Object WMIオブジェクト（プロセスの一覧情報）
'******************************************************************************
Public Function GetWin32Process() As Object
    Set GetWin32Process = GetWmiSWbemObject("Win32_Process")
End Function
'******************************************************************************
'* [概  要] GetWin32NetworkAdapterConfiguration メソッド
'* [詳  細] WMIオブジェクト（ネットワークアダプタ設定の一覧情報）を取得する。
'*
'* @return Object WMIオブジェクト（ネットワークアダプタ設定の一覧情報）
'******************************************************************************
Public Function GetWin32NetworkAdapterConfiguration() As Object
    Set GetWin32NetworkAdapterConfiguration = GetWmiSWbemObject("Win32_NetworkAdapterConfiguration")
End Function
'******************************************************************************
'* [概  要] GetWin32OperatingSystem メソッド
'* [詳  細] WMIオブジェクト（OS情報）を取得する。
'*
'* @return Object WMIオブジェクト（OS情報）
'******************************************************************************
Public Function GetWin32OperatingSystem() As Object
    Set GetWin32OperatingSystem = GetWmiSWbemObject("Win32_OperatingSystem")
End Function
'******************************************************************************
'* [概  要] GetWin32Printer メソッド
'* [詳  細] WMIオブジェクト（プリンタの一覧情報）を取得する。
'*
'* @return Object WMIオブジェクト（プリンタの一覧情報）
'******************************************************************************
Public Function GetWin32Printer() As Object
    Set GetWin32Printer = GetWmiSWbemObject("Win32_Printer")
End Function
'******************************************************************************
'* [概  要] GetWin32ComputerSystem メソッド
'* [詳  細] WMIオブジェクト（PCの情報）を取得する。
'*
'* @return Object WMIオブジェクト（PCの情報）
'******************************************************************************
Public Function GetWin32ComputerSystem() As Object
    Set GetWin32ComputerSystem = GetWmiSWbemObject("Win32_ComputerSystem")
End Function
'******************************************************************************
'* [概  要] GetWin32Processor メソッド
'* [詳  細] WMIオブジェクト（CPUの情報）を取得する。
'*
'* @return Object WMIオブジェクト（CPUの情報）
'******************************************************************************
Public Function GetWin32Processor() As Object
    Set GetWin32Processor = GetWmiSWbemObject("Win32_Processor")
End Function
'******************************************************************************
'* [概  要] GetWin32LogicalDisk メソッド
'* [詳  細] WMIオブジェクト（ディスクの情報）を取得する。
'*
'* @return Object WMIオブジェクト（ディスクの情報）
'******************************************************************************
Public Function GetWin32LogicalDisk() As Object
    Set GetWin32LogicalDisk = GetWmiSWbemObject("Win32_LogicalDisk")
End Function
'******************************************************************************
'* [概  要] GetWin32NetworkAdapter メソッド
'* [詳  細] WMIオブジェクト（ネットワークアダプタの一覧情報）を取得する。
'*
'* @return Object WMIオブジェクト（ネットワークアダプタの一覧情報）
'******************************************************************************
Public Function GetWin32NetworkAdapter() As Object
    Set GetWin32NetworkAdapter = GetWmiSWbemObject("Win32_NetworkAdapter")
End Function
'******************************************************************************
'* [概  要] GetWin32PnPEntity メソッド
'* [詳  細] WMIオブジェクト（プラグアンドプレイデバイスの一覧情報）を取得する。
'*
'* @return Object WMIオブジェクト（プラグアンドプレイデバイスの一覧情報）
'******************************************************************************
Public Function GetWin32PnPEntity() As Object
    Set GetWin32PnPEntity = GetWmiSWbemObject("Win32_PnPEntity")
End Function
'******************************************************************************
'* [概  要] GetWin32Service メソッド
'* [詳  細] WMIオブジェクト（サービスの一覧情報）を取得する。
'*
'* @return Object WMIオブジェクト（サービスの一覧情報）
'******************************************************************************
Public Function GetWin32Service() As Object
    Set GetWin32Service = GetWmiSWbemObject("Win32_Service")
End Function
'******************************************************************************
'* [概  要] GetWin32UserAccount メソッド
'* [詳  細] WMIオブジェクト（ユーザアカウントの一覧情報）を取得する。
'*
'* @return Object WMIオブジェクト（ユーザアカウントの一覧情報）
'******************************************************************************
Public Function GetWin32UserAccount() As Object
    Set GetWin32UserAccount = GetWmiSWbemObject("Win32_UserAccount")
End Function
'******************************************************************************
'* [概  要] GetWin32Product メソッド
'* [詳  細] WMIオブジェクト（MSIでインストールされた製品一覧情報）を取得する。
'*
'* @return Object WMIオブジェクト（MSIでインストールされた製品一覧情報）
'******************************************************************************
Public Function GetWin32Product() As Object
    Set GetWin32Product = GetWmiSWbemObject("Win32_Product")
End Function
'******************************************************************************
'* [概  要] GetWin32ScheduledJob メソッド
'* [詳  細] WMIオブジェクト（スケジュールされたジョブの一覧情報）を取得する。
'*
'* @return Object WMIオブジェクト（スケジュールされたジョブの一覧情報）
'******************************************************************************
Public Function GetWin32ScheduledJob() As Object
    Set GetWin32ScheduledJob = GetWmiSWbemObject("Win32_ScheduledJob")
End Function
'******************************************************************************
'* [概  要] GetWin32BIOS メソッド
'* [詳  細] WMIオブジェクト（BIOSの情報）を取得する。
'*
'* @return Object WMIオブジェクト（BIOSの情報）
'******************************************************************************
Public Function GetWin32BIOS() As Object
    Set GetWin32BIOS = GetWmiSWbemObject("Win32_BIOS")
End Function
