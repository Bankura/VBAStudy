VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'高機能ストップウォッチクラス
'Byことりちゅん
'
'ver 1.0 : 2019/2/3 : 当初            https://www.excel-chunchun.com/entry/2019/02/03/233535
'ver 2.0 : 2019/2/5 : 大幅な仕様変更  https://www.excel-chunchun.com/entry/2019/02/05/014303
'ver 2.1 : 2019/2/7 : APIを変更

Private Const class_name = "clsTimer"

'※エラーコードは適当
Private Const NOSTART_ERROR_CODE = 8001
Private Const NORESET_ERROR_CODE = 8002
Private Const NOSTART_ERROR_MESSAGE = "タイマーが開始されていません。"
Private Const NORESET_ERROR_MESSAGE = "タイマーが初期化されていません。"

Private StartNum As Long
Private LapNum   As Long
Private PauseNum As Long  '停止時間の一時記憶変数
Private LapTimes As Collection
Private DefaultFormat As String

#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongPtr
    #Else
        Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    #End If
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

#If VBA7 Then
    Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If

'面倒なので32bitと64bitの型を揃える
Public Function GetTime() As Long

'#If Win64 Then
'    GetTime = CLng(GetTickCount64)
'#Else
'    GetTime = GetTickCount
'#End If
    GetTime = timeGetTime()

End Function

'初期化
Private Sub Class_Initialize()

    Call Reset
    
End Sub

'スタート：計測を開始／再開
Public Sub Start(Optional defFormat As String = "")
    If defFormat <> "" Then DefaultFormat = defFormat

    '停止していた時間分を開始時間に足すことで辻褄をあわせる
    StartNum = StartNum + (GetTime - PauseNum)
    LapNum = LapNum + (GetTime - PauseNum)
    PauseNum = 0
    
End Sub

'リセット：全ての時間を初期化
Public Sub Reset()

    StartNum = 0
    LapNum = 0
    PauseNum = 0
    Set LapTimes = New Collection
    
End Sub

'ポーズ：一時停止し停止時間を記憶
Public Sub Pause()
    If StartNum = 0 Then Err.Raise NOSTART_ERROR_CODE, class_name, NOSTART_ERROR_MESSAGE
    
    PauseNum = GetTime

End Sub

'スプリット：開始からの経過時間
Public Function SplitTime(Optional ByVal timeFormat As String = "") As Variant
    If StartNum = 0 Then Err.Raise NOSTART_ERROR_CODE, class_name, NOSTART_ERROR_MESSAGE
    Dim tm As Long
    
    tm = (GetTime - StartNum)

    If timeFormat = "" Then timeFormat = DefaultFormat
    If timeFormat = "" Then
        SplitTime = tm
    Else
        SplitTime = GetTimeFormat(tm, timeFormat)
    End If
    
End Function

'ラップ：直前の計測からの経過時間
Public Function LapTime(Optional ByVal timeFormat As String = "") As Variant
    If StartNum = 0 Then Err.Raise NOSTART_ERROR_CODE, class_name, NOSTART_ERROR_MESSAGE
    Dim tm As Long
    
    tm = (GetTime - LapNum)
    LapNum = GetTime

    If timeFormat = "" Then timeFormat = DefaultFormat
    If timeFormat = "" Then
        LapTime = tm
    Else
        LapTime = GetTimeFormat(tm, timeFormat)
    End If
    LapTimes.Add LapTime
    
End Function

'ラップタイムのコレクションを返す
Public Property Get Laps() As Variant
    If StartNum = 0 Then Err.Raise NOSTART_ERROR_CODE, class_name, NOSTART_ERROR_MESSAGE
    
    If LapTimes.Count = 0 Then
        Laps = Split(vbNullString)
    Else
        Dim arr() As Variant
        ReDim arr(1 To LapTimes.Count)
        Dim i As Long
        For i = 1 To LapTimes.Count
            arr(i) = LapTimes(i)
        Next
        Laps = arr
    End If
    
End Property

'ミリセコンドから任意の書式に変換
Private Function GetTimeFormat(tm As Long, timeFormat As String) As Variant

    Select Case True
        Case timeFormat = "s"
            GetTimeFormat = CDbl(tm) / 1000
            
        Case timeFormat = "ms"
            GetTimeFormat = tm
            
        Case timeFormat Like "*ms"
            GetTimeFormat = Right(String(10, " ") & Format(tm, timeFormat), 10)
        
        Case timeFormat Like "*s"
            GetTimeFormat = Right(String(10, " ") & Format(CDbl(tm) / 1000, timeFormat), 10)
        
        Case Else
            GetTimeFormat = Format(tm, timeFormat)
    End Select
            
End Function
