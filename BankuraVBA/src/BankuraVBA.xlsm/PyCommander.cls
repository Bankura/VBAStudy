VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PyCommander"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Pythonコマンド実行クラス
'* [詳  細] Pythonの実行に関する処理を提供する。
'*
'* @author Bankura
'* Copyright (c) 2020-2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************

'******************************************************************************
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 定数定義
'******************************************************************************

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mCommander As Commander
Private mPythonInstallPath As String

'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* PythonInstallPath プロパティ
'*
'* Pythonのインストール先パス
'*-----------------------------------------------------------------------------
Public Property Get PythonInstallPath() As String
    PythonInstallPath = mPythonInstallPath
End Property
Public Property Let PythonInstallPath(arg As String)
    mPythonInstallPath = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Call SetPythonInstallPathFromRegistry
    Call SetPythonEnviron

    Set mCommander = New Commander
    Call CommandMode
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] CommandMode
'* [詳  細] コマンドを実行するモードに変更する。
'*
'******************************************************************************
Public Sub CommandMode()
    mCommander.BaseCommand = "cmd.exe /c " & mPythonInstallPath & "\python.exe -c ""{0}"" > {1}"
End Sub

'******************************************************************************
'* [概  要] ScriptMode
'* [詳  細] スクリプトを実行するモードに変更する。
'*
'******************************************************************************
Public Sub ScriptMode()
    mCommander.BaseCommand = "cmd.exe /c " & mPythonInstallPath & "\python.exe {0} > {1}"
End Sub

'******************************************************************************
'* [概  要] SetPythonEnviron
'* [詳  細] Pythonの環境変数を設定する。
'*          環境変数「Path」にPythonのパスを追加する。
'*
'******************************************************************************
Public Sub SetPythonEnviron()
    If mPythonInstallPath <> "" Then
        Call Base.AppendEnvItem("Path", mPythonInstallPath & "\Scripts")
        Call Base.AppendEnvItem("Path", mPythonInstallPath & "\Library\bin")
        Call Base.AppendEnvItem("Path", mPythonInstallPath)
    End If
End Sub

'******************************************************************************
'* [概  要] WriteScriptAndRun
'* [詳  細] スクリプトを作成し、実行した結果を取得する。
'*
'* @param scriptString Script文字列、または1次元配列
'* @param recieveResultString 結果を文字列で受け取るか
'*                            （True:String, False:配列）
'* @return Variant コマンド実行結果
'******************************************************************************
Public Function WriteScriptAndRun(scriptString, Optional recieveResultString As Boolean = True) As Variant
    Dim tempScrPath As String: tempScrPath = FileUtils.GetTempFilePath(, ".py")
    Call FileUtils.WriteUTF8TextFile(tempScrPath, scriptString)
    
    WriteScriptAndRun = ExecScript(tempScrPath, recieveResultString)

    Call IO.fso.DeleteFile(tempScrPath, True)
End Function

'******************************************************************************
'* [概  要] ExecCommand
'* [詳  細] コマンドを実行した結果を取得する。
'*
'* @param scriptString Script文字列、または1次元配列
'* @param recieveResultString 結果を文字列で受け取るか
'*                            （True:String, False:配列）
'* @return Variant コマンド実行結果
'******************************************************************************
Public Function ExecCommand(commandString As String, Optional recieveResultString As Boolean = True) As Variant
    Call CommandMode
    ExecCommand = Exec(commandString, recieveResultString)
End Function

'******************************************************************************
'* [概  要] ExecScript
'* [詳  細] スクリプトを実行した結果を取得する。
'*
'* @param scriptPath スクリプトのパス
'* @param recieveResultString 結果を文字列で受け取るか
'*                            （True:String, False:配列）
'* @return Variant スクリプト実行結果
'******************************************************************************
Public Function ExecScript(scriptPath As String, Optional recieveResultString As Boolean = True) As Variant
    Call ScriptMode
    ExecScript = Exec(scriptPath, recieveResultString)
End Function

'******************************************************************************
'* [概  要] Exec
'* [詳  細] 指定したコマンドまたはスクリプトを実行した結果を取得する。
'* [備  考] 事前に CommandMode / ScriptMode を呼び出してモードを指定する想定。
'*
'* @param execString 実行文字列（コマンドまたはスクリプト）
'* @param recieveResultString 結果を文字列で受け取るか
'*                            （True:String, False:配列）
'* @return Variant スクリプト実行結果
'******************************************************************************
Public Function Exec(execString As String, Optional recieveResultString As Boolean = True) As Variant
    If recieveResultString Then
        Exec = mCommander.GetCommandResult(execString)
    Else
        Exec = mCommander.GetCommandResultAsArray(execString)
    End If
End Function

'******************************************************************************
'* [概  要] SetPythonInstallPathFromRegistry
'* [詳  細] レジストリからPythonのインストールパスを取得し、設定する。
'* [参  考] https://www.python.org/dev/peps/pep-0514/
'*
'******************************************************************************
Private Sub SetPythonInstallPathFromRegistry()
    Const PYTHON_REG_PATH As String = "SOFTWARE\Python\PythonCore"
    Const PYTHON_REG_WOW_PATH As String = "SOFTWARE\Wow6432Node\Python\PythonCore"
    
    Dim stdRegProv As Object: Set stdRegProv = CreateStdRegProv()
    Dim regPythonInstallPathString As String
    Dim regRoot As Long

    Dim vArr
    stdRegProv.EnumKey HKEY_CURRENT_USER, PYTHON_REG_PATH, vArr
    If IsNull(vArr) Then
        stdRegProv.EnumKey HKEY_LOCAL_MACHINE, PYTHON_REG_PATH, vArr
        If IsNull(vArr) Then
            stdRegProv.EnumKey HKEY_LOCAL_MACHINE, PYTHON_REG_WOW_PATH, vArr
            regPythonInstallPathString = PYTHON_REG_WOW_PATH & "\{0}\InstallPath"
            regRoot = HKEY_LOCAL_MACHINE
        Else
            regPythonInstallPathString = PYTHON_REG_PATH & "\{0}\InstallPath"
            regRoot = HKEY_LOCAL_MACHINE
        End If
    Else
        regPythonInstallPathString = PYTHON_REG_PATH & "\{0}\InstallPath"
        regRoot = HKEY_CURRENT_USER
    End If
    
    If Not IsNull(vArr) Then
        ArrayUtils.Sort vArr, False
        stdRegProv.GetStringValue regRoot, StringUtils.PlaceHolder(regPythonInstallPathString, CStr(vArr(0))), , mPythonInstallPath
    End If
End Sub
