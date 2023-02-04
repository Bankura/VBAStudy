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
Private mCallbackObjCol As Collection
Private mCallbackParamCol As Collection
Private mCallbackResultCol As Collection

'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* Logger プロパティ
'*-----------------------------------------------------------------------------
Public Property Get Logger(Optional logFName As String = "bankuravba.log") As Logger
    If mLogger Is Nothing Then
        Set mLogger = New Logger
        Call mLogger.Init(LogLevelEnum.lvTrace, True, IO.ExecPath, logFName)
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

'*-----------------------------------------------------------------------------
'* ActiveSheetEx プロパティ
'*-----------------------------------------------------------------------------
Public Property Get ActiveSheetEx() As WorkSheetEx
    Set ActiveSheetEx = Core.Init(New WorkSheetEx, Application.ActiveSheet.Name)
End Property

'******************************************************************************
'* メソッド定義
'******************************************************************************

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
'* [概  要] CreateWmiServices メソッド
'* [詳  細] GetObject メソッドを使用してSWbemServicesオブジェクトを生成する。
'*
'* [補  足] WMIの名前空間は、以下のように「コンピュータの管理」から確認可能。
'*          ・コントロールパネル→管理ツール→コンピュータの管理 を選択
'*            以下のコマンドでも起動可能
'*            %windir%\system32\compmgmt.msc /s
'*          ・サービスとアプリケーション→WMIコントロールを選択し、
'*            右クリック→プロパティを選択
'*          ・表示された「WMIコントロールのプロパティ」のセキュリティタブを選択
'*
'* [参  考] http://dodonpa.la.coocan.jp/windows_service_wmi_1.htm
'*
'* @param strComputer 省略可。コンピュータ名。
'* @param ns          省略可。名前空間。
'* @return SWbemServicesオブジェクト。
'*
'******************************************************************************
Function CreateSWbemServices(Optional strComputer As String = ".", Optional ns As String = "\root\cimv2", Optional userId As String, Optional passwd As String) As Object
    If userId = "" Then
        Set CreateSWbemServices = GetObject("winmgmts:\\" & strComputer & ns)
    Else
        Set CreateSWbemServices = Core.wmi.ConnectServer(strComputer, ns, userId, passwd)
    End If
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
    Dim strTempFile As String: strTempFile = FileUtils.GetTempFilePath(, ".vbs")
    Call FileUtils.WriteUTF8TextFile(strTempFile, strScriptCodes)
    
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
'* [概  要] Swap
'* [詳  細] 値を入れ替える。
'*
'* @param val1 値1
'* @param val2 値2
'******************************************************************************
Public Sub Swap(val1, val2)
    Dim tmp: tmp = val1
    val1 = val2
    val2 = tmp
End Sub

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

'******************************************************************************
'* [概  要] ForEach
'* [詳  細] 「For Each」で繰り返し可能なオブジェクトに対して、処理を適用する。
'*
'* @param obj  「For Each」で繰り返し可能なオブジェクト
'* @param proc 関数名、またはFuncオブジェクト、または
'*                   Exec（x As Object）メソッドを
'*                   持つオブジェクト。
'*
'******************************************************************************
Public Sub ForEach(ByVal obj As Object, ByVal proc As Variant)
    Dim o
    For Each o In obj
        If ValidateUtils.IsFunc(proc) Then
            Call proc.Apply(o)
        ElseIf ValidateUtils.IsString(proc) And proc <> "" Then
            Call Application.Run(proc, o)
        ElseIf IsObject(proc) Then
            Call proc.Exec(o)
        End If
    Next
End Sub

'******************************************************************************
'* [概  要] GetRandom
'* [詳  細] ランダム値を生成する。
'*
'* @param willRandomize Randomizeを実行するかどうか
'* @return Single ランダム値
'*
'******************************************************************************
Public Function GetRandom(Optional willRandomize As Boolean = False) As Single
    If willRandomize Then
        Randomize
    End If
    GetRandom = Rnd
End Function

'******************************************************************************
'* [概  要] GetRandomInt
'* [詳  細] ランダムな整数を生成する。
'*
'* @param minVal 最小値
'* @param maxVal 最大値
'* @param willRandomize Randomizeを実行するかどうか
'* @return Long ランダムな整数
'*
'******************************************************************************
Public Function GetRandomInt(Optional minVal As Long = 0, Optional maxVal As Long = MAX_LONG - 1, Optional willRandomize As Boolean = False) As Long
    Dim randomVal As Single: randomVal = GetRandom(willRandomize)
    If minVal > maxVal Then
        Swap minVal, maxVal
    End If
    GetRandomInt = Int((maxVal * randomVal) + (0 - (minVal * randomVal) + minVal) + randomVal)
End Function

'******************************************************************************
'* [概  要] GetRandomDate
'* [詳  細] ランダムな日付を生成する。
'*
'* @param minVal 最小値
'* @param maxVal 最大値
'* @param noTimeVal 時刻を含まないか（True：含まない）
'* @param willRandomize Randomizeを実行するかどうか
'* @return Date ランダムな日付（時分秒含む）
'*
'******************************************************************************
Public Function GetRandomDate(Optional minVal As Single = 0, Optional maxVal As Single = 2958465, Optional noTimeVal As Boolean = False, Optional willRandomize As Boolean = False) As Date
    Dim randomVal As Single: randomVal = GetRandom(willRandomize)
    If minVal = 0 Then
        minVal = Now
    End If
    If minVal > maxVal Then
        Swap minVal, maxVal
    End If
    If noTimeVal Then
        GetRandomDate = CDate(Int((maxVal * randomVal) + (0 - (minVal * randomVal) + minVal) + randomVal))
    Else
        GetRandomDate = CDate((maxVal * randomVal) + (0 - (minVal * randomVal) + minVal) + randomVal)
    End If
End Function


'******************************************************************************
'* [概  要] GetArrayDataRandom
'* [詳  細] 配列のデータをランダムに取得する。
'*
'* @param arr 対象の配列
'* @param willRandomize Randomizeを実行するかどうか
'* @return Variant ランダムな要素
'*
'******************************************************************************
Public Function GetArrayDataRandom(Arr As Variant, Optional willRandomize As Boolean = False) As Variant
    GetArrayDataRandom = Arr(Int((UBound(Arr) - LBound(Arr) + 1) * GetRandom(willRandomize) + LBound(Arr)))
End Function

'******************************************************************************
'* [概  要] GetRandomIntArray
'* [詳  細] ランダムな整数を格納した配列のデータを生成する。
'*
'* @param numOfItems 配列の個数
'* @param minVal 最小値
'* @param maxVal 最大値
'* @param noOverLap 重複値を許容しないか（True:重複なし）
'* @param willRandomize Randomizeを実行するかどうか
'* @return Variant 配列データ
'*
'******************************************************************************
Public Function GetRandomIntArray(numOfItems As Long, _
                                  Optional minVal As Long = 0, _
                                  Optional maxVal As Long = MAX_LONG - 1, _
                                  Optional noOverLap As Boolean = True, _
                                  Optional willRandomize As Boolean = False) As Variant
    If willRandomize Then
        Randomize
    End If
    
    If minVal > maxVal Then
        Swap minVal, maxVal
    End If
    
    If numOfItems > (maxVal - minVal + 1) Then
        noOverLap = False
    End If
    
    Dim Arr() As Long: ReDim Arr(0 To numOfItems - 1)
    Dim i As Long, j As Long, numVal As Long
    For i = 0 To numOfItems - 1
        If Not noOverLap Then
            Arr(i) = GetRandomInt(minVal, maxVal)
        Else
            Do
                Dim flg As Boolean: flg = True
                numVal = GetRandomInt(minVal, maxVal)
                
                If i = 0 Then
                    Arr(i) = numVal
                    Exit Do
                End If
                
                For j = 0 To i - 1
                    If Arr(j) = numVal Then
                       flg = False
                       Exit For
                    End If
                Next
                
                If flg Then
                    Arr(i) = numVal
                    Exit Do
                End If
            Loop
        End If
    Next
    GetRandomIntArray = Arr
End Function

'******************************************************************************
'* [概  要] GetRandomString
'* [詳  細] ランダムな文字列を生成する。
'*
'* @param textLength   文字列長
'* @param useableChars 使用する文字
'* @param willRandomize Randomizeを実行するかどうか
'* @return String ランダムな文字列
'*
'******************************************************************************
Public Function GetRandomString(ByVal textLength As Long, _
                                Optional useableChars As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", _
                                Optional willRandomize As Boolean = False) As String
    Dim tmpStr As StringEx: Set tmpStr = New StringEx
    Dim maxLen As Long: maxLen = StringUtils.CheckLength(useableChars)
    
    Dim v As Variant
    For Each v In GetRandomIntArray(textLength, 1, maxLen, , willRandomize)
        tmpStr.Append Mid$(useableChars, CLng(v), 1)
    Next
    GetRandomString = tmpStr.ToString
End Function

'******************************************************************************
'* [概  要] GetRandomHiragana
'* [詳  細] ランダムなひらがな文字列を生成する。
'*
'* @param textLength   文字列長
'* @param willRandomize Randomizeを実行するかどうか
'* @return String ランダムなひらがな文字列
'*
'******************************************************************************
Public Function GetRandomHiragana(ByVal textLength As Long, _
                                  Optional willRandomize As Boolean = False) As String
    GetRandomHiragana = GetRandomString(textLength, _
        "あいうえおかきくけこさしすせそたちつてとなにぬねの" & _
        "はひふへほまみむめもやゆよらりるれろわをん" & _
        "がぎぐげござじずぜぞだぢづでどばびぶべぼぱぴぷぺぽ", willRandomize)
End Function

'******************************************************************************
'* [概  要] GetRandomKatakana
'* [詳  細] ランダムなカタカナ文字列を生成する。
'*
'* @param textLength   文字列長
'* @param willRandomize Randomizeを実行するかどうか
'* @return String ランダムなカタカナ文字列
'*
'******************************************************************************
Public Function GetRandomKatakana(ByVal textLength As Long, _
                                  Optional willRandomize As Boolean = False) As String
    GetRandomKatakana = StrConv(GetRandomHiragana(textLength, willRandomize), vbKatakana)
End Function

'******************************************************************************
'* [概  要] GetRandomNumString
'* [詳  細] ランダムな数字文字列を生成する。
'*
'* @param textLength   文字列長
'* @param willRandomize Randomizeを実行するかどうか
'* @return String ランダムな数字文字列
'*
'******************************************************************************
Public Function GetRandomNumString(ByVal textLength As Long, _
                                  Optional willRandomize As Boolean = False) As String
    GetRandomNumString = GetRandomString(textLength, "0123456789", willRandomize)
End Function

'******************************************************************************
'* [概  要] GetRandomHalfAlphaNumeric
'* [詳  細] ランダムな半角英数文字列を生成する。
'*
'* @param textLength   文字列長
'* @param willRandomize Randomizeを実行するかどうか
'* @return String ランダムな半角英数文字列
'*
'******************************************************************************
Public Function GetRandomHalfAlphaNumeric(ByVal textLength As Long, _
                                  Optional willRandomize As Boolean = False) As String
    GetRandomHalfAlphaNumeric = GetRandomString(textLength, "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", willRandomize)
End Function

'******************************************************************************
'* [概  要] GetRandomHalfAlphaNumericSymbol
'* [詳  細] ランダムな半角英数文字列を生成する。
'*
'* @param textLength   文字列長
'* @param willRandomize Randomizeを実行するかどうか
'* @return String ランダムな半角英数文字列
'*
'******************************************************************************
Public Function GetRandomHalfAlphaNumericSymbol(ByVal textLength As Long, _
                                                Optional allowSpace As Boolean = True, _
                                                Optional willRandomize As Boolean = False) As String
    GetRandomHalfAlphaNumericSymbol = GetRandomString(textLength, IIf(allowSpace, " ", "") & "!""#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~", willRandomize)
End Function

'******************************************************************************
'* [概  要] ParallelExec
'* [詳  細] 並列処理を行う。
'* [注  意] 速くはない。
'* [参  考] https://www.excel-chunchun.com/entry/2019/03/27/005233
'*
'* @param fncName 実行する関数名
'*                呼び出す関数内で、終了時に以下の処理を行うこと。
'*                  Application.DisplayAlerts = False
'*                  ThisWorkbook.Close False
'*                処理を行わない場合、子プロセスの実行を検知できず、
'*                無限ループとなる。
'* @param procNum 並列実行するプロセス数
'* @param paParams 関数に渡すパラメータ
'*
'******************************************************************************
Public Sub ParallelExec(ByVal fncName As String, ByVal procNum As Integer, ParamArray paParams())
    Dim params As Variant: params = VariantUtils.EmptyArrayIfParamArrayMissing(paParams)
    Dim apps As Collection: Set apps = New Collection
    Dim app As Excel.Application
    Dim wb As Workbook
    Dim i As Long
    For i = 1 To procNum
        Set app = New Application
        apps.Add app
        
        ' 自WorkBookを別のインスタンスでOpen（読取り専用）
        Set wb = app.Workbooks.Open(ThisWorkbook.FullName, _
                                    UpdateLinks:=False, _
                                    ReadOnly:=True)
        ' 子プロセス実行
        app.Run "'" & wb.Name & "'!ParallelSubExec", i, fncName, params
        
        DoEvents
    Next
    Set app = Nothing: Set wb = Nothing
    
    ' 子プロセス終了待ち
    For i = 1 To apps.Count
        Do While apps(i).Workbooks.Count > 0
            Application.Wait [Now() + "00:00:00.2"]
            DoEvents
        Loop
    Next
    
    ' 子Excelのインスタンスの破棄
    On Error Resume Next
    For i = 1 To apps.Count
        apps(1).Quit
        apps.Remove 1
    Next
    On Error GoTo 0
End Sub

'******************************************************************************
'* [概  要] ParallelSubExec
'* [詳  細] ParallelExecのサブ処理。
'* [参  考] https://www.excel-chunchun.com/entry/2019/03/27/005233
'*
'* @param n       実行番号
'* @param fncName 実行する関数名
'* @param params  関数に渡すパラメータ
'*
'******************************************************************************
Private Sub ParallelSubExec(n As Long, fncName As String, params As Variant)
    Dim param
    Dim sb As StringEx: Set sb = New StringEx
    Call sb.Append("'").Append(fncName).Append(" """).Append(n - 1).Append("""")
    For Each param In params
        sb.Append ", """
        sb.Append CStr(param)
        sb.Append """"
    Next
    sb.Append "'"
    Application.OnTime [Now() + "00:00:00.2"], sb.ToString
End Sub

'******************************************************************************
'* [概  要] OnTimeForClass
'* [詳  細] クラスのメソッドに対してOntime処理を実行する。
'* [注  意] ・OnTime同様、別の処理が実行中の場合、別処理が終了するまで待機する。
'*            （並列で実行はされない）
'*          ・処理実行後、実行結果がmCallbackResultColに蓄積されるため、
'*            不要になった実行結果は、ClearResultOnTimeForClass を呼び出して、
'*            クリアすること。
'*
'* @param startSec    実行開始までの待機時間（秒）
'* @param callbackObj 実行するクラスのオブジェクト
'* @param fncName     実行するメソッド名
'* @param paParams    関数に渡すパラメータ
'* @return String     実行予約キー（実行結果を確認する際に使用）
'*
'******************************************************************************
Public Function OnTimeForClass(startSec As Long, callbackObj As Object, fncName As String, ParamArray paParams()) As String
    If mCallbackObjCol Is Nothing Then
        Set mCallbackObjCol = New Collection
    End If
    If mCallbackParamCol Is Nothing Then
        Set mCallbackParamCol = New Collection
    End If
    If mCallbackResultCol Is Nothing Then
        Set mCallbackResultCol = New Collection
    End If
    
    Dim keyStr As String: keyStr = Base.CreateUUID()
    Dim params As Variant: params = VariantUtils.EmptyArrayIfParamArrayMissing(paParams)
    mCallbackObjCol.Add callbackObj, keyStr
    mCallbackParamCol.Add params, keyStr
    
    Application.OnTime Now + TimeSerial(0, 0, startSec), "'OnTimeForClassSubExec """ & fncName & """, """ & keyStr & """'"
    OnTimeForClass = keyStr
End Function

'******************************************************************************
'* [概  要] OnTimeForClassSubExec
'* [詳  細] OnTimeForClassのサブ処理。
'*
'* @param fncName 実行するメソッド名
'* @param keyStr  実行予約キー（実行結果を確認する際に使用）
'*
'******************************************************************************
Public Sub OnTimeForClassSubExec(fncName As String, keyStr As String)
    Dim ret, p: p = mCallbackParamCol(keyStr)
    Dim callbackObj As Object: Set callbackObj = mCallbackObjCol(keyStr)

    Select Case ArrayUtils.GetLength(p)
        Case 1
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod, CVar(p(0))), keyStr)
        Case 2
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod, CVar(p(0)), CVar(p(1))), keyStr)
        Case 3
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod, CVar(p(0)), CVar(p(1)), CVar(p(2))), keyStr)
        Case 4
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod, CVar(p(0)), CVar(p(1)), CVar(p(2)), CVar(p(3))), keyStr)
        Case 5
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod, CVar(p(0)), CVar(p(1)), CVar(p(2)), CVar(p(3)), CVar(p(4))), keyStr)
        Case 6
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod, CVar(p(0)), CVar(p(1)), CVar(p(2)), CVar(p(3)), CVar(p(4)), CVar(p(5))), keyStr)
        Case 7
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod, CVar(p(0)), CVar(p(1)), CVar(p(2)), CVar(p(3)), CVar(p(4)), CVar(p(5)), CVar(p(6))), keyStr)
        Case 8
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod, CVar(p(0)), CVar(p(1)), CVar(p(2)), CVar(p(3)), CVar(p(4)), CVar(p(5)), CVar(p(6)), CVar(p(7))), keyStr)
        Case 9
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod, CVar(p(0)), CVar(p(1)), CVar(p(2)), CVar(p(3)), CVar(p(4)), CVar(p(5)), CVar(p(6)), CVar(p(7)), CVar(p(8))), keyStr)
        Case 10
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod, CVar(p(0)), CVar(p(1)), CVar(p(2)), CVar(p(3)), CVar(p(4)), CVar(p(5)), CVar(p(6)), CVar(p(7)), CVar(p(8)), CVar(p(9))), keyStr)
        Case Else
            Call mCallbackResultCol.Add(CallByName(callbackObj, fncName, vbMethod), keyStr)
    End Select
    mCallbackObjCol.Remove keyStr
    mCallbackParamCol.Remove keyStr
End Sub

'******************************************************************************
'* [概  要] GetResultOnTimeForClass
'* [詳  細] OnTimeForClassで実行した処理の実行結果を取得する。
'*
'* @param keyStr  実行予約キー（OnTimeForClassの戻り値）
'* @@return Variant 実行結果（戻り値がない処理の場合Empty）
'*
'******************************************************************************
Public Function GetResultOnTimeForClass(keyStr As String) As Variant
    GetResultOnTimeForClass = mCallbackResultCol(keyStr)
End Function

'******************************************************************************
'* [概  要] GetResultOnTimeForClass
'* [詳  細] OnTimeForClassで実行した処理の実行結果をクリアする。
'*
'******************************************************************************
Public Sub ClearResultOnTimeForClass()
    Set mCallbackResultCol = New Collection
End Sub
