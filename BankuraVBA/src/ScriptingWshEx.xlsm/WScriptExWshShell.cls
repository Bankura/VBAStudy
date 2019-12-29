VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WScriptExWshShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WScript.WshShellクラスのラップ・拡張クラス
'* [詳  細] WScript.WshShellクラスのWrapperとしての機能を提供する他、
'*          WScript.WshShellクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WScript.WshShellクラスについて。
' WindowsScriptHostのシェルオブジェクト。Dosコマンドの実行、アプリケーションの起動などに使用する。
' <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364436(v=msdn.10)>
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As IWshRuntimeLibrary.WshShell
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
'* CurrentDirectory プロパティ
'*
'* [補  足]
'* ・現在アクティブになっているディレクトリを取得または変更します。
'*-----------------------------------------------------------------------------
Property Get CurrentDirectory() As String
    CurrentDirectory = mOrigin.CurrentDirectory
End Property
Property Let CurrentDirectory(arg As String)
    mOrigin.CurrentDirectory = arg
End Property

'*-----------------------------------------------------------------------------
'* Environment プロパティ（読み取り専用）
'*
'* [補  足]
'* ・WshEnvironment オブジェクト (環境変数のコレクション) を返します。
'* ・<https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364502%28v%3dmsdn.10%29>
'*
'* @param Type_ 省略可能。環境変数の種類を指定します。
'*              環境変数の存在場所を示す値として System、User、Volatile、
'*              または Process のいずれかを使用可能。指定しない場合はSystem。
'*
'*-----------------------------------------------------------------------------
Property Get Environment(Optional Type_) As WScriptExWshEnvironment
    Dim wrap As WScriptExWshEnvironment: Set wrap = New WScriptExWshEnvironment
    If IsMissing(Type_) Then
        Set wrap.Origin = mOrigin.Environment
    Else
        Set wrap.Origin = mOrigin.Environment(Type_)
    End If
    Set Environment = wrap
End Property

'*-----------------------------------------------------------------------------
'* SpecialFolders プロパティ（読み取り専用）
'*
'* [補  足]
'* ・SpecialFolders オブジェクト (特定フォルダのコレクション) を返します。
'* ・<https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364490%28v%3dmsdn.10%29>
'*-----------------------------------------------------------------------------
Property Get SpecialFolders() As WScriptExWshCollection
    Dim wrap As WScriptExWshCollection: Set wrap = New WScriptExWshCollection
    Set wrap.Origin = mOrigin.SpecialFolders
    Set SpecialFolders = wrap
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("WScript.Shell")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] AppActivate メソッド
'* [詳  細] AppActivate のラッパーメソッド。
'*          アプリケーション ウィンドウをアクティブにします。
'* [参  考] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364396%28v%3dmsdn.10%29>
'*
'* @param App アクティブにするアプリケーションを指定します。
'*            この引数には、アプリケーションのタイトル バーに表示される
'*            タイトル文字列か、アプリケーションのプロセス ID を指定できます。
'* @param Wait 省略可能。
'* @return Boolean
'*
'******************************************************************************
Public Function AppActivate(App, Optional Wait) As Boolean
    If IsMissing(Wait) Then
        AppActivate = mOrigin.AppActivate(App)
    Else
        AppActivate = mOrigin.AppActivate(App, Wait)
    End If
End Function

'******************************************************************************
'* [概  要] CreateShortcut メソッド
'* [詳  細] CreateShortcut のラッパーメソッド。
'*          ショートカットまたは URL ショートカットへのオブジェクト参照を作成します。
'* [参  考] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364400%28v%3dmsdn.10%29>
'*
'* @param PathLink 作成するショートカットのパスとファイル名
'* @return WScriptExWshShortcutまたはWScriptExWshURLShortcutオブジェクト
'*
'******************************************************************************
Public Function CreateShortcut(PathLink As String) As Object
    Dim obj As Object
    Set obj = mOrigin.CreateShortcut(PathLink)
    If TypeName(obj) = "WshShortcut" Then
        Dim wrap As WScriptExWshShortcut: Set wrap = New WScriptExWshShortcut
        Set wrap.Origin = obj
        Set CreateShortcut = wrap
    ElseIf TypeName(obj) = "WshURLShortcut" Then
        Dim wrap2 As WScriptExWshURLShortcut: Set wrap2 = New WScriptExWshURLShortcut
        Set wrap2.Origin = obj
        Set CreateShortcut = wrap2
    Else
        Set CreateShortcut = obj
    End If
End Function

'******************************************************************************
'* [概  要] Exec メソッド
'* [詳  細] Exec のラッパーメソッド。
'*          子コマンドシェルでアプリケーションを実行します。
'*          アプリケーションから StdIn/StdOut/StdErr ストリームにアクセス
'*          できます。
'* [参  考] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364356%28v%3dmsdn.10%29>
'*
'* @param Command スクリプトの実行に使用するコマンド ラインを示す文字列値です。
'*                コマンド プロンプトから入力する場合と全く同じコマンド ライン
'*                を指定します。
'* @return WScriptExWshExec
'*
'******************************************************************************
Public Function Exec(Command As String) As WScriptExWshExec
    Dim wrap As WScriptExWshExec: Set wrap = New WScriptExWshExec
    Set wrap.Origin = mOrigin.Exec(Command)
    Set Exec = wrap
End Function

'******************************************************************************
'* [概  要] ExpandEnvironmentStrings メソッド
'* [詳  細] ExpandEnvironmentStrings のラッパーメソッド。
'*          環境変数を展開した値を返します。
'*          ExpandEnvironmentStrings メソッドは、PROCESS 環境領域に定義された
'*          環境変数のみを展開します。環境変数名は "%" で囲む必要があります。
'*          また、環境変数名の大文字と小文字は区別されません。
'* [参  考] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364405%28v%3dmsdn.10%29>
'*
'* @param Src 展開する環境変数の名前
'* @return 環境変数を展開した値
'*
'******************************************************************************
Public Function ExpandEnvironmentStrings(Src As String) As String
    ExpandEnvironmentStrings = mOrigin.ExpandEnvironmentStrings(Src)
End Function

'******************************************************************************
'* [概  要] LogEvent メソッド
'* [詳  細] LogEvent のラッパーメソッド。
'*          イベント エントリをログ ファイルに追加します。
'* [参  考] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364408%28v%3dmsdn.10%29>
'*
'* @param Type_ イベントの種類を表現する数値
'* @param Message ログ エントリのテキスト
'* @param Target 省略可能。イベント ログを格納するコンピュータシステムの名前
'*               を示す文字列値。既定値はローカルのコンピュータ システム。
'* @return イベントが正常に記録された場合は TRUE、それ以外の場合は FALSE
'*
'******************************************************************************
Public Function LogEvent(Type_, Message As String, Optional Target As String) As Boolean
    LogEvent = mOrigin.LogEvent(Type_, Message, Target)
End Function

'******************************************************************************
'* [概  要] Popup メソッド
'* [詳  細] Popup のラッパーメソッド。
'*          ポップアップ メッセージ ボックスにテキストを表示します。
'* [参  考] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364428%28v%3dmsdn.10%29>
'*
'* @param Text ポップアップ ウィンドウに表示するテキスト
'* @param SecondsToWait 省略可能。ポップアップ ウィンドウを閉じるまで待機する
'*                      秒数。ゼロ (既定値) を指定した場合、ポップアップ
'*                      メッセージボックスはユーザーがウィンドウを閉じるまで
'*                      ずっと表示されます。
'* @param Title 省略可能。ポップアップウィンドウのタイトルに表示するテキスト。
'*              省略すると、"Windows Script Host" となります。
'* @param Type_ 省略可能。ポップアップメッセージ ボックスに表示するボタンと
'*              アイコンの種類を示す数値。選択する種類によってメッセージ
'*              ボックスの使い方が決まります。
'* @return メッセージ ボックス終了時にクリックするボタンの番号を示す整数値
'*
'******************************************************************************
Public Function Popup(text As String, Optional SecondsToWait, Optional Title, Optional Type_) As Long
    If IsMissing(SecondsToWait) Then
        If IsMissing(Title) Then
             If IsMissing(Type_) Then
                Popup = mOrigin.Popup(text)
            Else
                Popup = mOrigin.Popup(text, , , Type_)
            End If
        ElseIf IsMissing(Type_) Then
            Popup = mOrigin.Popup(text, , Title)
        Else
            Popup = mOrigin.Popup(text, , Title, Type_)
        End If
    Else
        If IsMissing(Title) Then
             If IsMissing(Type_) Then
                Popup = mOrigin.Popup(text, SecondsToWait)
            Else
                Popup = mOrigin.Popup(text, SecondsToWait, , Type_)
            End If
        ElseIf IsMissing(Type_) Then
            Popup = mOrigin.Popup(text, SecondsToWait, Title)
        Else
            Popup = mOrigin.Popup(text, SecondsToWait, Title, Type_)
        End If
    End If
End Function

'******************************************************************************
'* [概  要] RegDelete メソッド
'* [詳  細] RegDelete のラッパーメソッド。
'*          レジストリから指定されたキーまたは値を削除します。
'* [参  考] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364415%28v%3dmsdn.10%29>
'*
'* @param Name レジストリ内で削除するキーまたは値の名前を示す文字列値
'*
'******************************************************************************
Public Sub RegDelete(Name As String)
    Call mOrigin.RegDelete(Name)
End Sub

'******************************************************************************
'* [概  要] RegRead メソッド
'* [詳  細] RegRead のラッパーメソッド。
'*          レジストリ内のキー名または値名の値を返します。
'* [参  考] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364416%28v%3dmsdn.10%29>
'*
'* @param Name 読み取るキーまたは値の名前
'* @return 次の 5 種類の値を返します。
'*           型            説明                        形式
'*           -----------------------------------------------------------
'*           REG_SZ        文字列です。                文字列
'*           REG_DWORD     数値です。                  整数
'*           REG_BINARY    2進数の値です。             VBArray (整数)
'*           REG_EXPAND_SZ 展開可能な文字列です｡
'*                         (例 : "%windir%\\calc.exe") 文字列
'*           REG_MULTI_SZ  文字列の配列です。          VBArray (文字列)
'*
'******************************************************************************
Public Function RegRead(Name As String)
    RegRead = mOrigin.RegRead(Name)
End Function

'******************************************************************************
'* [概  要] RegWrite メソッド
'* [詳  細] RegWrite のラッパーメソッド。
'*          新しいキーの作成、新しい値名の既存キーへの追加 (および値の設定)、
'*          既存の値名の値変更などを行います。
'* [参  考] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364417(v=msdn.10)>
'*
'* @param Name 作成、追加、変更するキー名、値名、または値を示す文字列値
'* @param Value 作成するキーの名前、既存のキーに追加する値の名前、または既存
'*              の値名に設定する値
'* @param Type_ 省略可能。レジストリに保存する値のデータ型。
'*
'******************************************************************************
Public Sub RegWrite(Name As String, Value, Optional Type_)
    If IsMissing(Type_) Then
        Call mOrigin.RegWrite(Name, Value)
    Else
        Call mOrigin.RegWrite(Name, Value, Type_)
    End If
End Sub

'******************************************************************************
'* [概  要] Run メソッド
'* [詳  細] Run のラッパーメソッド。
'*          新しいプロセス内でプログラムを実行します。
'* [参  考] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364421(v=msdn.10)>
'*
'* @param Command 実行するコマンド ラインを示す文字列値。
'*                この引数には、実行可能ファイルに渡すべきパラメータをすべて
'*                含める必要があります。
'* @param WindowStyle 省略可能。プログラムのウィンドウの外観を示す整数値。
'* @param WaitOnReturn 省略可能。スクリプト内の次のステートメントに進まずに
'*                     プログラムの実行が終了するまでスクリプトを待機させるか
'*                     どうかを示すブール値。
'*                     WaitOnReturn に TRUE を指定すると、プログラムの実行が
'*                     終了するまでスクリプトの実行は中断され、Run メソッドは
'*                     アプリケーションから返される任意のエラーコードを返し
'*                     ます。FALSE を指定すると、プログラムが開始すると Run
'*                     メソッドは即座に復帰して自動的に 0 を返します
'* @return 実行結果を示す整数値（エラーコード等）
'*
'******************************************************************************
Public Function Run(Command As String, Optional WindowStyle, Optional WaitOnReturn) As Long
    If IsMissing(WindowStyle) Then
        If IsMissing(WaitOnReturn) Then
            Run = mOrigin.Run(Command)
        Else
            Run = mOrigin.Run(Command, , WaitOnReturn)
        End If
    ElseIf IsMissing(WaitOnReturn) Then
        Run = mOrigin.Run(Command, WindowStyle)
    Else
        Run = mOrigin.Run(Command, WindowStyle, WaitOnReturn)
    End If
End Function

'******************************************************************************
'* [概  要] SendKeys メソッド
'* [詳  細] SendKeys のラッパーメソッド。
'*          キーボードから入力したときのように、1 つ以上のキー ストロークを
'*          アクティブなウィンドウに送ります。
'* [参  考] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364423(v=msdn.10)>
'*
'* @param Keys 送信するキーストロークを文字列式で指定
'* @param Wait 省略可能。
'*
'******************************************************************************
Public Sub SendKeys(Keys As String, Optional Wait)
    If IsMissing(Wait) Then
        Call mOrigin.SendKeys(Keys)
    Else
        Call mOrigin.SendKeys(Keys, Wait)
    End If
End Sub

'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------

'******************************************************************************
'* [概  要] SendKeysCtlF メソッド
'* [詳  細] SendKeys メソッドに Ctl + F キーを送信する。
'*
'* @param Wait 省略可能。
'*
'******************************************************************************
Public Sub SendKeysCtlF(Optional Wait As Long = 100)
    Call mOrigin.SendKeys("^f")
    Sleep Wait
End Sub

'******************************************************************************
'* [概  要] SendKeysESC メソッド
'* [詳  細] SendKeys メソッドに ESC キーを送信する。
'*
'* @param Wait 省略可能。
'*
'******************************************************************************
Public Sub SendKeysESC(Optional Wait As Long = 100)
    Call mOrigin.SendKeys("{ESC}")
    Sleep Wait
End Sub

'******************************************************************************
'* [概  要] SendKeysHOME メソッド
'* [詳  細] SendKeys メソッドに HOME キーを送信する。
'*
'* @param Wait 省略可能。
'*
'******************************************************************************
Public Sub SendKeysHOME(Optional Wait As Long = 100)
    Call mOrigin.SendKeys("{HOME}")
    Sleep Wait
End Sub

'******************************************************************************
'* [概  要] SendKeysShiftEND メソッド
'* [詳  細] SendKeys メソッドに Shift + END キーを送信する。
'*
'* @param Wait 省略可能。
'*
'******************************************************************************
Public Sub SendKeysShiftEND(Optional Wait As Long = 100)
    Call mOrigin.SendKeys("+{END}")
    Sleep Wait
End Sub

'******************************************************************************
'* [概  要] SendKeysCtlC メソッド
'* [詳  細] SendKeys メソッドに Ctl + C キーを送信する。
'*
'* @param Wait 省略可能。
'*
'******************************************************************************
Public Sub SendKeysCtlC(Optional Wait As Long = 100)
    Call mOrigin.SendKeys("^c")
    Sleep Wait
End Sub

'******************************************************************************
'* [概  要] SendKeysAltF4 メソッド
'* [詳  細] SendKeys メソッドに Alt + F4 キーを送信する。
'*
'* @param Wait 省略可能。
'*
'******************************************************************************
Public Sub SendKeysAltF4(Optional Wait As Long = 100)
    Call mOrigin.SendKeys("%{F4}")
    Sleep Wait
End Sub

'******************************************************************************
'* [概  要] Is32BitProcessorForApp メソッド
'* [詳  細] 使用するアプリケーションが32ビットかをチェックします。
'*
'* @return チェック結果（True: 32Bit、False: 64bit）
'*
'******************************************************************************
Public Function Is32BitProcessorForApp() As Boolean
    Dim proc As String: proc = Me.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
    If proc = "x86" Then
       Is32BitProcessorForApp = True
    Else
       Is32BitProcessorForApp = False
    End If
End Function

'******************************************************************************
'* [概  要] Is32BitProcessor メソッド
'* [詳  細] 使用する端末のプロセッサが32ビットかをチェックします。
'*
'* @return チェック結果（True: 32Bit、False: 64bit）
'*
'******************************************************************************
Public Function Is32BitProcessor() As Boolean
    Dim proc As String: proc = Me.ExpandEnvironmentStrings("%PROCESSOR_ARCHITEW6432%")
    If proc = "x86" Then
       Is32BitProcessor = True
    Else
       Is32BitProcessor = False
    End If
End Function

'******************************************************************************
'* [概  要] CreateObject32bit メソッド
'* [詳  細] 32ビット環境のObjectを生成します。
'* [参  考] <https://github.com/vocho/vbs/blob/a5c3ee608103638678c983da00ec290c4b8ab90c/CreateObject32bit.vbs>
'*
'* @param strClassName 生成対象のクラス名。"Shell.Application"等。
'* @return 32ビット環境Object
'*
'******************************************************************************
Public Function CreateObject32bit(ByVal strClassName As String) As Variant
    If Is32BitProcessorForApp Then
     Set CreateObject32bit = CreateObject(strClassName)
     Exit Function
    End If

    Dim shell As Object: Set shell = CreateObject("Shell.Application")
    shell.Windows().Item(0).PutProperty strClassName, Nothing
    
    ' 一時スクリプトコマンドテキスト生成
    Dim strScriptCodes As String
    strScriptCodes = "CreateObject(""Shell.Application"").Windows().Item(0).PutProperty """ & strClassName & """, CreateObject(""" & strClassName & """)" & vbNewLine & _
                     "Set objExec = CreateObject(""WScript.Shell"").Exec(""MSHTA.EXE -"")" & vbNewLine & _
                     "Set objWMIService = GetObject(""winmgmts:"")" & vbNewLine & _
                     "lngCurrentPID = objWMIService.Get(""Win32_Process.Handle="" & objExec.ProcessID).ParentProcessID" & vbNewLine & _
                     "objExec.Terminate" & vbNewLine & _
                     "lngParentPID = objWMIService.Get(""Win32_Process.Handle="" & lngCurrentPID).ParentProcessID" & vbNewLine & _
                     "Do While objWMIService.ExecQuery(""SELECT * FROM Win32_Process WHERE ProcessID="" & lngParentPID).Count<>0" & vbNewLine & _
                     "    WScript.Sleep 1000" & vbNewLine & _
                     "Loop" & vbNewLine & _
                     "Set objFSO = CreateObject(""Scripting.FileSystemObject"")" & vbNewLine & _
                     "If objFSO.FileExists(WScript.ScriptFullName) Then objFSO.DeleteFile WScript.ScriptFullName" & vbNewLine & _
                     ""

    ' 一時スクリプトファイル作成
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim strTempFile As String
    With fso
        Do
            strTempFile = .BuildPath(.GetSpecialFolder(2), .GetTempName() & ".vbs")
        Loop While .FileExists(strTempFile)
        With .OpenTextFile(strTempFile, 2, True)
            .WriteLine strScriptCodes
            .Close
        End With
    End With
    
    ' 一時スクリプトファイル実行(32bit)
    With Me.Environment("Process")
        .Item("SysWOW64") = fso.BuildPath(.Item("SystemRoot"), "SysWOW64")
        .Item("WScriptName") = fso.GetFileName("C:\WINDOWS\SysWOW64\cscript.exe")
        .Item("WScriptWOW64") = fso.BuildPath(.Item("SysWOW64"), .Item("WScriptName"))
        .Item("Run") = .Item("WScriptWOW64") & " """ & strTempFile & """"
         Me.Run .Item("Run"), True
    End With
    
    ' オブジェクト受け取り
    Do
        Set CreateObject32bit = shell.Windows().Item(0).GetProperty(strClassName)
    Loop While CreateObject32bit Is Nothing
End Function

'******************************************************************************
'* [概  要] CreateObjectFromWsh メソッド
'* [詳  細] Wsh経由でObjectを生成します。
'*
'* @param strClassName 生成対象のクラス名。"Shell.Application"等。
'* @param strTypeName  タイプ名。
'* @return Object
'*
'******************************************************************************
Public Function CreateObjectFromWsh(ByVal strClassName As String, ByVal strTypeName As String) As Variant
    Dim shell As Object: Set shell = CreateObject("Shell.Application")
    shell.Windows().Item(0).PutProperty strClassName, Nothing

    ' 一時スクリプトコマンドテキスト生成
    Dim strScriptCodes As String
    strScriptCodes = "CreateObject(""Shell.Application"").Windows().Item(0).PutProperty """ & strClassName & """, WScript.CreateObject(""" & strClassName & """, """ & strTypeName & """)" & vbNewLine & _
                     "Set objExec = CreateObject(""WScript.Shell"").Exec(""MSHTA.EXE -"")" & vbNewLine & _
                     "Set objWMIService = GetObject(""winmgmts:"")" & vbNewLine & _
                     "lngCurrentPID = objWMIService.Get(""Win32_Process.Handle="" & objExec.ProcessID).ParentProcessID" & vbNewLine & _
                     "objExec.Terminate" & vbNewLine & _
                     "lngParentPID = objWMIService.Get(""Win32_Process.Handle="" & lngCurrentPID).ParentProcessID" & vbNewLine & _
                     "Do While objWMIService.ExecQuery(""SELECT * FROM Win32_Process WHERE ProcessID="" & lngParentPID).Count<>0" & vbNewLine & _
                     "    WScript.Sleep 1000" & vbNewLine & _
                     "Loop" & vbNewLine & _
                     "Set objFSO = CreateObject(""Scripting.FileSystemObject"")" & vbNewLine & _
                     "If objFSO.FileExists(WScript.ScriptFullName) Then objFSO.DeleteFile WScript.ScriptFullName" & vbNewLine & _
                     ""
    ' 一時スクリプトファイル作成
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim strTempFile
    With fso
        Do
            strTempFile = .BuildPath(.GetSpecialFolder(2), .GetTempName() & ".vbs")
        Loop While .FileExists(strTempFile)
        With .OpenTextFile(strTempFile, 2, True)
            .WriteLine strScriptCodes
            .Close
        End With
    End With

    ' 一時スクリプトファイル実行
    Me.Run "cscript.exe """ & strTempFile & """", True
    
    ' オブジェクト受け取り
    Do
        Set CreateObjectFromWsh = shell.Windows().Item(0).GetProperty(strClassName)
    Loop While CreateObjectFromWsh Is Nothing
End Function

