VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SoundUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] サウンド関連ユーティリティクラス
'* [詳  細] 音声・効果音・音楽関連のユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'*
'* [依　存]
'*
'* [参  考] https://excel.syogyoumujou.com/vba/mci.html
'*          http://home.att.ne.jp/zeta/gen/excel/c04p04.htm
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, Optional ByVal lpstrReturnString As String = vbNullString, Optional ByVal uReturnLength As Long = 0, Optional ByVal hwndback As LongPtr = 0) As Long
Private Declare PtrSafe Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare PtrSafe Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Private Declare PtrSafe Function PlaySoundAPI Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function PlaySoundOfByteData Lib "winmm.dll" Alias "PlaySoundA" (ByRef pszSound As Byte, ByVal hModule As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'******************************************************************************
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 定数定義
'******************************************************************************
Private Const MCI_SEND_CON_FILENAME As String = "MySoundConnect"
Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_MEMORY = &H4
Private Const SND_LOOP = &H8
Private Const SND_NOSTOP = &H10

'******************************************************************************
'* 変数定義
'******************************************************************************
Private mIsSoundOpened As Boolean

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    mIsSoundOpened = False
End Sub
Private Sub Class_Terminate()
    Call CloseSound
End Sub
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] PlaySoundJust
'* [詳  細] 選択した音声ファイルをすぐに再生する｡
'*
'* @param useByteData バイトデータに
'* @return String xxx
'*
'******************************************************************************
Public Sub PlaySoundJust(Optional ByVal useByteData As Boolean = False)
    '音声ファイルの選択
    Dim strFilePath As String
    strFilePath = Application.GetOpenFileName("音声ファイル (*.wav),*.wav")
    If strFilePath <> "False" Then
        If useByteData Then
            Dim buf() As Byte
            Dim fileNum As Long:  fileNum = FreeFile
            Open strFilePath For Binary As fileNum
            ReDim buf(LOF(fileNum))
            Get fileNum, , buf
            Close fileNum
    
            PlaySoundOfByteData buf(0), 0, SND_ASYNC + SND_MEMORY
            Exit Sub
        End If
        PlaySoundAPI strFilePath, 0, SND_ASYNC
    End If
End Sub

'******************************************************************************
'* [概  要] OpenSound
'* [詳  細] 音声ファイルをメモリに読み込み｡
'*
'******************************************************************************
Public Sub OpenSound()
    '音声ファイルの選択
    Dim strFilePath As String
    strFilePath = Application.GetOpenFileName("音声ファイル (*.wav;*.mid;*.mp3;*.wma),*.wav;*.mid;*.mp3;*.wma")
    If strFilePath <> "False" Then
        If 0 < InStrRev(strFilePath, ".") Then
            Call mciSendString("Open """ & strFilePath & """ alias " & MCI_SEND_CON_FILENAME)
            mIsSoundOpened = True
        End If
    End If
End Sub

'******************************************************************************
'* [概  要] CloseSound
'* [詳  細] 音声ファイルをメモリからクリア｡
'* [補  足] ブックを閉じる前に実行すること｡
'*
'******************************************************************************
Public Sub CloseSound()
    If mIsSoundOpened Then
        Call mciSendString("Close All")
        mIsSoundOpened = False
    End If
End Sub

'******************************************************************************
'* [概  要] PlaySound
'* [詳  細] 音声ファイルを再生｡
'* [補  足] 事前に音声ファイルを読み込まないと、再生できない｡
'*
'******************************************************************************
Public Sub PlaySound()
    Call mciSendString("Play " & MCI_SEND_CON_FILENAME & " from 0")
End Sub

'******************************************************************************
'* [概  要] PauseSound
'* [詳  細] 音声ファイルの再生を一時停止｡
'*
'******************************************************************************
Public Sub PauseSound()
    Call mciSendString("Pause " & MCI_SEND_CON_FILENAME)
End Sub

'******************************************************************************
'* [概  要] ResumeSound
'* [詳  細] 音声ファイルの再生一時停止解除｡
'*
'******************************************************************************
Public Sub ResumeSound()
    Call mciSendString("Resume " & MCI_SEND_CON_FILENAME)
End Sub

'******************************************************************************
'* [概  要] StopSound
'* [詳  細] 音声ファイルの再生を停止｡
'*
'******************************************************************************
Public Sub StopSound()
    Call mciSendString("Stop " & MCI_SEND_CON_FILENAME)
End Sub

'******************************************************************************
'* [概  要] BeepSound
'* [詳  細] 音声ファイルの再生を停止｡
'*
'* @param pitch 音の高さ
'* @param length 音の長さ（milliseconds）
'*
'******************************************************************************
Public Sub BeepSound(pitch As Long, length As Long)
    Call Beep(pitch, length)
End Sub

Public Sub ErrorMessageBeep()
    Call MessageBeep(&H10&)
End Sub
Public Sub InformationMessageBeep()
    Call MessageBeep(&H1&)
End Sub

