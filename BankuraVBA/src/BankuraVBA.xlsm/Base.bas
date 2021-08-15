Attribute VB_Name = "Base"
Option Explicit
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] BankuraVBA共通基盤モジュール
'* [詳  細] 共通で使用するプロシージャを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019-2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************

'******************************************************************************
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 構造体定義
'******************************************************************************
Public Type Array2DIndex
    x As Long
    y As Long
End Type

'******************************************************************************
'* 定数定義
'******************************************************************************
Public Const MAX_INT As Integer = 32767
Public Const MAX_LONG As Long = 2147483647
Public Const TWIP As Long = 567

#If Win64 Then
    Public Const LONGPTR_SIZE = 8
#Else
    Public Const LONGPTR_SIZE = 4
#End If

' For Wsh
Public Const WshHide = 0
Public Const ForReading = 1
    
'******************************************************************************
'* 変数定義
'******************************************************************************
Private mDisplayAlerts As Boolean
Private mScreenUpdating As Boolean
Private mCalculation As Long
Private mEnableEvents As Boolean
Private mRegExp As Object
Private mShell As Object
Private mWshNetwork As Object
Private mSc As Object
Private mSettingInfo As SettingInfo
Private mSettingSheetName As String
Private mSettingSheetStartRow As Long
Private mSettingSheetStartCol As Long
Private mLogger As Logger

'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* Logger プロパティ
'*-----------------------------------------------------------------------------
Public Property Get Logger() As Logger
    If mLogger Is Nothing Then
        Set mLogger = New Logger
        Call mLogger.Init(LogLevelEnum.lvTrace, True, IO.ExecPath, "bankuravba.log")
    End If
    Set Logger = mLogger
End Property

'*-----------------------------------------------------------------------------
'* SettingInfo プロパティ
'*-----------------------------------------------------------------------------
Public Property Get SettingInfo() As SettingInfo
    Set SettingInfo = mSettingInfo
End Property
Public Property Set SettingInfo(ByVal arg As SettingInfo)
    Set mSettingInfo = arg
End Property

'*-----------------------------------------------------------------------------
'* SettingSheetName プロパティ
'*-----------------------------------------------------------------------------
Public Property Get SettingSheetName() As String
    SettingSheetName = mSettingSheetName
End Property
Public Property Let SettingSheetName(ByVal arg As String)
    mSettingSheetName = arg
End Property

'*-----------------------------------------------------------------------------
'* SettingSheetStartRow プロパティ
'*-----------------------------------------------------------------------------
Public Property Get SettingSheetStartRow() As Long
    SettingSheetStartRow = mSettingSheetStartRow
End Property
Public Property Let SettingSheetStartRow(ByVal arg As Long)
    mSettingSheetStartRow = arg
End Property

'*-----------------------------------------------------------------------------
'* SettingSheetStartCol プロパティ
'*-----------------------------------------------------------------------------
Public Property Get SettingSheetStartCol() As Long
    SettingSheetStartCol = mSettingSheetStartCol
End Property
Public Property Let SettingSheetStartCol(ByVal arg As Long)
    mSettingSheetStartCol = arg
End Property

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] ChangeDisplayWorkbookTabs
'* [詳  細] シート見出しの表示・非表示を切り替える。
'*
'******************************************************************************
Public Sub ChangeDisplayWorkbookTabs()
    With ActiveWindow
        If .DisplayWorkbookTabs Then
            'シート見出しを非表示
            .DisplayWorkbookTabs = False
        Else
            'シート見出しを表示
            .DisplayWorkbookTabs = True
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] ChangeDisplayGridlines
'* [詳  細] 罫線の表示・非表示を切り替える。
'*
'******************************************************************************
Public Sub ChangeDisplayGridlines()
    With ActiveWindow
        If .DisplayGridlines Then
            '罫線を非表示
            .DisplayGridlines = False
        Else
            '罫線を表示
            .DisplayGridlines = True
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] ChangeDisplayHeadings
'* [詳  細] 行列番号の表示・非表示を切り替える。
'*
'******************************************************************************
Public Sub ChangeDisplayHeadings()
    With ActiveWindow
        If .DisplayHeadings Then
            '行列番号を非表示
            .DisplayHeadings = False
        Else
            '行列番号を表示
            .DisplayHeadings = True
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] ChangeDisplayHeadings
'* [詳  細] スクロールバーの表示・非表示を切り替える。
'*
'*******************************************************************************
Public Sub ChangeDisplayScrollBar()
    With ActiveWindow
        If .DisplayHorizontalScrollBar Then
            'スクロールバーを非表示
            .DisplayHorizontalScrollBar = False
            .DisplayVerticalScrollBar = False
        Else
            'スクロールバーを表示
            .DisplayHorizontalScrollBar = True
            .DisplayVerticalScrollBar = True
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] ChangeReferenceStyle
'* [詳  細] Excel参照形式を切り替える。
'*
'******************************************************************************
Public Sub ChangeReferenceStyle()
    With Application
        If .ReferenceStyle = xlA1 Then
            .ReferenceStyle = xlR1C1
        Else
            .ReferenceStyle = xlA1
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] エラー処理。
'* [詳  細] エラー発生時の処理を行う。
'*
'******************************************************************************
Public Sub ErrorProcess()
    Debug.Print "エラー発生 Number: " & Err.Number & " Source: " & Err.Source & " Description: " & Err.Description
End Sub

'******************************************************************************
'* [概  要] 開始処理。
'* [詳  細] 処理のスピード向上のため、Excelの設定を変更する。
'*
'******************************************************************************
Public Sub StartProcess()
    Call SaveApplicationProperties
    With Application
        .Cursor = xlWait
        .DisplayAlerts = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
End Sub

'******************************************************************************
'* [概  要] 終了処理。
'* [詳  細] 処理のスピード向上のため変更したExcelの設定を元に戻す。
'*
'******************************************************************************
Public Sub EndProcess()
    With Application
        .Cursor = xlDefault
        .DisplayAlerts = mDisplayAlerts
        .ScreenUpdating = mScreenUpdating
        .Calculation = mCalculation
        .EnableEvents = mEnableEvents
        .StatusBar = False
    End With
End Sub

'******************************************************************************
'* [概  要] Application設定退避処理。
'* [詳  細] Applicationの設定をメンバ変数に退避する。
'*
'******************************************************************************
Public Sub SaveApplicationProperties()
    With Application
        mDisplayAlerts = .DisplayAlerts
        mScreenUpdating = .ScreenUpdating
        mCalculation = .Calculation
        mEnableEvents = .EnableEvents
    End With
End Sub

'******************************************************************************
'* [概  要] 正規表現オブジェクト取得処理。
'* [詳  細] 正規表現オブジェクトを取得する。未生成の場合生成する。
'*
'******************************************************************************
Public Function GetRegExp() As Object
    If mRegExp Is Nothing Then
        Set mRegExp = CreateObject("VBScript.RegExp")
    End If
    Set GetRegExp = mRegExp
End Function

'******************************************************************************
'* [概  要] Shellオブジェクト取得処理。
'* [詳  細] Shellオブジェクトを取得する。未生成の場合生成する。
'*
'******************************************************************************
Public Function GetShell() As Object
    If mShell Is Nothing Then
        Set mShell = CreateObject("Shell.Application")
    End If
    Set GetShell = mShell
End Function

'******************************************************************************
'* [概  要] WScript.Networkオブジェクト取得処理。
'* [詳  細] WScript.Networkオブジェクトを取得する。未生成の場合生成する。
'*
'******************************************************************************
Public Function GetWshNetwork() As Object
    If mWshNetwork Is Nothing Then
        Set mWshNetwork = CreateObject("WScript.Network")
    End If
    Set GetWshNetwork = mWshNetwork
End Function

'******************************************************************************
'* [概  要] ScriptControlオブジェクト取得処理。
'* [詳  細] ScriptControlオブジェクトを取得する。未生成の場合生成する。
'*
'******************************************************************************
Public Function GetScriptControl() As Object
    If mSc Is Nothing Then
        Set mSc = CreateObject32bit("MSScriptControl.ScriptControl")
    End If
    Set GetScriptControl = mSc
End Function


'******************************************************************************
'* [概  要] CDO.Messageオブジェクト生成処理。
'* [詳  細] CDO.Messageオブジェクトを生成する。
'*
'******************************************************************************
Public Function CreateCDOMessage() As Object
    Set CreateCDOMessage = CreateObject("CDO.Message")
End Function

'******************************************************************************
'* [概  要] 設定情報オブジェクト取得処理。
'* [詳  細] 設定情報オブジェクトを取得する。未生成の場合生成する。
'*
'******************************************************************************
Public Function GetSettingInfo() As SettingInfo
    If mSettingInfo Is Nothing Then
        Set mSettingInfo = Core.Init(New SettingInfo, mSettingSheetName, mSettingSheetStartRow, mSettingSheetStartCol)
    End If
    Set GetSettingInfo = mSettingInfo
End Function

Public Function GetMasterValueByCode(masterName As String, code As String) As String
    If mSettingInfo Is Nothing Then
        Set mSettingInfo = GetSettingInfo()
    End If
    GetMasterValueByCode = mSettingInfo.GetMasterValueByCode(masterName, code)
End Function
Public Function GetMasterCodeByValue(masterName As String, val As String) As String
    If mSettingInfo Is Nothing Then
        Set mSettingInfo = GetSettingInfo()
    End If
    GetMasterCodeByValue = mSettingInfo.GetMasterCodeByValue(masterName, val)
End Function

'*******************************************************************************
'* [概  要] コンピュータ名設定処理
'* [詳  細] 実行端末のコンピュータ名を取得。
'*
'* @param String コンピュータ名
'*
'*******************************************************************************
Public Function GetComputerName() As String
    GetComputerName = Core.Wsh.ComputerName
End Function

'******************************************************************************
'* [概  要] 実行アプリケーション判定処理
'* [詳  細] 実行アプリケーションがExcelか判定する。
'*
'* @param Boolean 処理結果（True:Excel False：Excel以外）
'*
'******************************************************************************
Public Function CheckXlApplication() As Boolean
    CheckXlApplication = InStr(Application.Name, "Excel") > 0
End Function

'******************************************************************************
'* [概  要] Is32BitProcessorForApp
'* [詳  細] 使用するアプリケーションが32ビットかをチェックする。
'*
'* @return チェック結果（True: 32Bit、False: 64bit）
'*
'******************************************************************************
Public Function Is32BitProcessorForApp() As Boolean
    Dim proc As String: proc = Wsh.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
    If proc = "x86" Then
       Is32BitProcessorForApp = True
    Else
       Is32BitProcessorForApp = False
    End If
End Function

'******************************************************************************
'* [概  要] Is32BitProcessor
'* [詳  細] 使用する端末のプロセッサが32ビットかをチェックする。
'*
'* @return チェック結果（True: 32Bit、False: 64bit）
'*
'******************************************************************************
Public Function Is32BitProcessor() As Boolean
    Dim proc As String: proc = Wsh.ExpandEnvironmentStrings("%PROCESSOR_ARCHITEW6432%")
    If proc = "x86" Then
       Is32BitProcessor = True
    Else
       Is32BitProcessor = False
    End If
End Function

'******************************************************************************
'* [概  要] CreateObject32bit
'* [詳  細] 32ビット環境のObjectを生成する。
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
    
    Base.GetShell.Windows().Item(0).PutProperty strClassName, Nothing
    
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
    With IO.fso
        Dim strTempFile As String
        Do
            strTempFile = .BuildPath(.GetSpecialFolder(2), .GetTempName() & ".vbs")
        Loop While .FileExists(strTempFile)
        With .OpenTextFile(strTempFile, 2, True)
            .WriteLine strScriptCodes
            .Close
        End With
    End With
    
    ' 一時スクリプトファイル実行(32bit)
    With Core.Wsh.Environment("Process")
        .Item("SysWOW64") = IO.fso.BuildPath(.Item("SystemRoot"), "SysWOW64")
        .Item("WScriptName") = IO.fso.GetFileName("C:\WINDOWS\SysWOW64\cscript.exe")
        .Item("WScriptWOW64") = IO.fso.BuildPath(.Item("SysWOW64"), .Item("WScriptName"))
        .Item("Run") = .Item("WScriptWOW64") & " """ & strTempFile & """"
         Core.Wsh.Run .Item("Run"), True
    End With
    
    ' オブジェクト受け取り
    Do
        Set CreateObject32bit = Base.GetShell.Windows().Item(0).GetProperty(strClassName)
    Loop While CreateObject32bit Is Nothing
End Function

'******************************************************************************
'* [概  要] SetAppSettingsNormal
'* [詳  細] アプリケーションの設定を通常の設定にする。
'*
'******************************************************************************
Public Sub SetAppSettingsNormal()
    With Application
        .Cursor = xlDefault
        .DisplayAlerts = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub

'******************************************************************************
'* [概  要] JudgeCond
'* [詳  細] 2値を指定した比較演算子で比較・判定する。
'*
'* @param val1 比較する値1
'* @param val2 比較する値2
'* @param cond 比較演算子（"<", ">", "<=", ">=", "="）
'* @param flg 判定結果の反転用フラグ
'*            "<", ">"の場合に反転後の条件に"="を含まない。
'*            クイックソートの判定で使用。
'*
'* @return Boolean 判定結果
'*
'******************************************************************************
Public Function JudgeCond(val1, val2, cond As String, Optional flg As Boolean = True) As Boolean
    If flg Then
        Select Case cond
            Case ">"
                JudgeCond = val1 > val2
            Case "<"
                JudgeCond = val1 < val2
            Case ">="
                JudgeCond = val1 >= val2
            Case "<="
                JudgeCond = val1 <= val2
            Case "="
                JudgeCond = val1 = val2
            Case Else
                Err.Raise 9999, "JudgeCond", "Bad Condition."
        End Select
    Else
        Select Case cond
            Case ">"
                JudgeCond = val1 < val2
            Case "<"
                JudgeCond = val1 > val2
            Case ">="
                JudgeCond = val1 <= val2
            Case "<="
                JudgeCond = val1 >= val2
            Case "="
                JudgeCond = val1 = val2
            Case Else
                Err.Raise 9999, "JudgeCond", "Bad Condition."
        End Select
    End If
End Function

'******************************************************************************
'* [概  要] CreateUUID
'* [詳  細] UUID(GUID)を生成する。
'* [参  考] https://stackoverflow.com/a/46474125/918626
'*
'* @return String UUID
'*
'******************************************************************************
Public Function CreateUUID() As String
    Dim myUuid As String
    Randomize Timer() + Application.hWnd
    Do While Len(myUuid) < 32
        If Len(myUuid) = 16 Then
            myUuid = myUuid & Hex$(8 + CInt(Rnd * 3))
        End If
        myUuid = myUuid & Hex$(CInt(Rnd * 15))
    Loop
    CreateUUID = Mid(myUuid, 1, 8) & "-" & Mid(myUuid, 9, 4) & "-" & Mid(myUuid, 13, 4) & "-" & Mid(myUuid, 17, 4) & "-" & Mid(myUuid, 21, 12)
End Function

'******************************************************************************
'* [概  要] AppendEnvItem
'* [詳  細] 環境変数を追加する。
'*
'* @param itemName  項目名
'* @param itemValue 設定値
'* @param envType   環境変数の種類（デフォルトは"Process"）
'*                    "System"  : システム環境変数。全ユーザーに適用される。
'*                    "User"    : ユーザー環境変数。ユーザーに適用される。
'*                    "Volatile": 揮発性環境変数。ログオフ時に破棄される。
'*                    "Process" : プロセス環境変数。プロセス終了時に破棄。
'* @param appendHead 先頭に加えるかどうか。
'*
'******************************************************************************
Public Sub AppendEnvItem(itemName As String, itemValue, Optional envType As String = "Process", Optional appendHead As Boolean = True)
    With Core.Wsh
        Dim destEnvValue: destEnvValue = .Environment(envType).Item(itemName)
        Dim sep As String: sep = IIf(destEnvValue <> "", ";", "")
        
        If Not StringUtils.Contains(destEnvValue, itemValue) Then
            If appendHead Then
                .Environment(envType).Item(itemName) = itemValue & sep & destEnvValue
            Else
                .Environment(envType).Item(itemName) = destEnvValue & sep & itemValue
            End If
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] EditEnvItem
'* [詳  細] 環境変数を編集する。
'*
'* @param itemName  項目名
'* @param itemValue 設定値
'* @param envType   環境変数の種類（デフォルトは"Process"）
'*                    "System"  : システム環境変数。全ユーザーに適用される。
'*                    "User"    : ユーザー環境変数。ユーザーに適用される。
'*                    "Volatile": 揮発性環境変数。ログオフ時に破棄される。
'*                    "Process" : プロセス環境変数。プロセス終了時に破棄。
'*
'******************************************************************************
Public Sub EditEnvItem(itemName As String, itemValue, Optional envType As String = "Process")
    With Core.Wsh
        .Environment(envType).Item(itemName) = itemValue
    End With
End Sub
