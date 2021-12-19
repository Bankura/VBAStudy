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
'* Copyright (c) 2019-2021 Bankura
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
Private Declare PtrSafe Function midiOutGetNumDevs Lib "winmm.dll" () As Long
Private Declare PtrSafe Function midiOutOpen Lib "winmm.dll" (lphMidiOut As LongPtr, ByVal uDeviceID As Long, ByVal dwCallback As LongPtr, ByVal dwInstance As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As LongPtr, ByVal dwMsg As Long) As Long
Private Declare PtrSafe Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As LongPtr) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'******************************************************************************
'* Enum定義
'******************************************************************************
Public Enum MidiNote0OctaveEnum
    MidiNoteC = 24   ' C        ド
    MidiNoteCS = 25  ' C# / Db  ド#
    MidiNoteD = 26   ' D        レ
    MidiNoteDS = 27  ' D# / Eb  レ#
    MidiNoteE = 28   ' E        ミ
    MidiNoteF = 29   ' F        ファ
    MidiNoteFS = 30  ' F# / Gb  ファ#
    MidiNoteG = 31   ' G        ソ
    MidiNoteGS = 32  ' G# / Ab  ソ#
    MidiNoteA = 33   ' A        ラ
    MidiNoteAS = 34  ' A# / Bb  ラ#
    MidiNoteB = 35   ' B        シ
End Enum

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
Private mSapiTokenCategory As Object
Private mIsMidiOpened As Boolean
Private mMidiHandle As LongPtr
Private mMidiStdTimbre  As Long  ' 1～128 音色
Private mMidiStdVolume  As Long  ' 1～127 ボリューム（ベロシティ）
Private mMidiStdChannel As Long  ' 0～15  チャネル
Private mMidiStdOctave  As Long  ' -2～8
Private mMidiStdTempo   As Long

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    mIsSoundOpened = False
    mIsMidiOpened = False
    mMidiHandle = 0
    mMidiStdTimbre = 1
    mMidiStdVolume = 127
    mMidiStdChannel = 0
    mMidiStdOctave = 3
    mMidiStdTempo = 120
    Set mSapiTokenCategory = CreateObject("SAPI.SpObjectTokenCategory")
    mSapiTokenCategory.SetID "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_OneCore\Voices", False
End Sub
Private Sub Class_Terminate()
    Call CloseSound
    Call CloseMidi
    Set mSapiTokenCategory = Nothing
End Sub

'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* MidiStdTimbre プロパティ
'*
'* Midiの標準の音色。1～127を指定。
'*-----------------------------------------------------------------------------
Public Property Get MidiStdTimbre() As Long
    MidiStdTimbre = mMidiStdTimbre
End Property
Public Property Let MidiStdTimbre(ByVal arg As Long)
    mMidiStdTimbre = arg
End Property

'*-----------------------------------------------------------------------------
'* MidiStdVolume プロパティ
'*
'* Midiの標準のボリューム（ベロシティ）。0～127を指定。
'*-----------------------------------------------------------------------------
Public Property Get MidiStdVolume() As Long
    MidiStdVolume = mMidiStdVolume
End Property
Public Property Let MidiStdVolume(ByVal arg As Long)
    mMidiStdVolume = arg
End Property

'*-----------------------------------------------------------------------------
'* MidiStdChannel プロパティ
'*
'* Midiの標準のチャンネル。0～15を指定。9の場合はドラム。
'*-----------------------------------------------------------------------------
Public Property Get MidiStdChannel() As Long
    MidiStdChannel = mMidiStdChannel
End Property
Public Property Let MidiStdChannel(ByVal arg As Long)
    mMidiStdChannel = arg
End Property

'*-----------------------------------------------------------------------------
'* MidiStdOctave プロパティ
'*
'* Midiの標準のオクターブ。-2～8を指定。
'*-----------------------------------------------------------------------------
Public Property Get MidiStdOctave() As Long
    MidiStdOctave = mMidiStdOctave
End Property
Public Property Let MidiStdOctave(ByVal arg As Long)
    If arg < -2 Or 8 < arg Then
        mMidiStdOctave = 3
        Exit Property
    End If
    mMidiStdOctave = arg
End Property

'*-----------------------------------------------------------------------------
'* MidiStdTempo プロパティ
'*
'* Midiの標準のテンポ。
'*-----------------------------------------------------------------------------
Public Property Get MidiStdTempo() As Long
    MidiStdTempo = mMidiStdTempo
End Property
Public Property Let MidiStdTempo(ByVal arg As Long)
    mMidiStdTempo = arg
End Property


'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] PlaySoundJust
'* [詳  細] 選択した音声ファイルをすぐに再生する｡
'*
'* @param useByteData バイトデータに変換するか
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
'* [詳  細] ビープ音を鳴らす｡
'*
'* @param pitch 音の高さ
'* @param length 音の長さ（milliseconds）
'*
'******************************************************************************
Public Sub BeepSound(pitch As Long, Length As Long)
    Call Beep(pitch, Length)
End Sub

'******************************************************************************
'* [概  要] ErrorMessageBeep
'* [詳  細] エラーメッセージ表示時の効果音を鳴らす｡
'*
'******************************************************************************
Public Sub ErrorMessageBeep()
    Call MessageBeep(&H10&)
End Sub

'******************************************************************************
'* [概  要] InformationMessageBeep
'* [詳  細] 情報メッセージ表示時の効果音を鳴らす｡
'*
'******************************************************************************
Public Sub InformationMessageBeep()
    Call MessageBeep(&H1&)
End Sub

'******************************************************************************
'* [概  要] Speak
'* [詳  細] テキストを音声読み上げする｡
'* [参  考] https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ee125647(v=vs.85)
'*
'* @param spText          読み上げるテキスト
'* @param engineNo        使用する音声合成エンジンの番号（デフォルト 0:日本語、1:英語(女性)）
'* @param engineNoOneCore 使用する音声合成エンジンの番号（OneCore）
'* @param spRate          音声の読み上げ速度（-10～10の範囲）
'* @param spVolume        音声の音量（0～100の範囲）
'* @param spAsync         非同期で読み上げを行うかどうか
'******************************************************************************
Public Sub Speak(ByVal spText As String, Optional engineNo As Integer = 0, Optional engineNoOneCore As Integer = 9999, _
                 Optional spRate As Integer = 0, Optional spVolume As Integer = 100, Optional spAsync As Boolean = False)
    With CreateObject("SAPI.SpVoice")
        If engineNoOneCore = 9999 Then
            Set .Voice = .GetVoices.Item(engineNo)
        Else
            Set .Voice = mSapiTokenCategory.EnumerateTokens.Item(engineNoOneCore)
        End If
        .Rate = spRate
        .Volume = spVolume
        If spAsync Then
            .Speak spText, 1
        Else
            .Speak spText
        End If
        
    End With
End Sub

'******************************************************************************
'* [概  要] SaveSpeakToWavFile
'* [詳  細] テキストを音声読み上げした内容をWavファイルに保存する｡
'*          （音声は出力されずファイル作成される）
'*
'* @param spText          読み上げるテキスト
'* @param saveFilePath    保存先Wavファイルパス
'* @param engineNo        使用する音声合成エンジンの番号（デフォルト 0:日本語、1:英語(女性)）
'* @param engineNoOneCore 使用する音声合成エンジンの番号（OneCore）
'* @param spRate          音声の読み上げ速度（-10～10の範囲）
'* @param spVolume        音声の音量（0～100の範囲）
'******************************************************************************
Public Sub SaveSpeakToWavFile(ByVal spText As String, ByVal saveFilePath As String, Optional engineNo As Integer = 0, Optional engineNoOneCore As Integer = 9999, _
                 Optional spRate As Integer = 0, Optional spVolume As Integer = 100)
    With CreateObject("SAPI.SpVoice")
        If engineNoOneCore = 9999 Then
            Set .Voice = .GetVoices.Item(engineNo)
        Else
            Set .Voice = mSapiTokenCategory.EnumerateTokens.Item(engineNoOneCore)
        End If
        .Rate = spRate
        .Volume = spVolume
        
        Dim fs: Set fs = CreateObject("SAPI.SpFileStream")
        fs.Open saveFilePath, 3
        Set .AudioOutputStream = fs
        .Speak spText
        Call fs.Close
    End With
End Sub

'******************************************************************************
'* [概  要] ShowAvailableSpeakEngines
'* [詳  細] 利用可能な音声合成エンジンを表示する｡
'*
'******************************************************************************
Public Sub ShowAvailableSpeakEngines()
    Dim i As Long, token
    Debug.Print "[Nomal Engine]"
    With CreateObject("SAPI.SpVoice")
        For i = 0 To .GetVoices.Count - 1
            Debug.Print i & ":" & .GetVoices.Item(i).GetDescription
        Next
    End With
    
    Debug.Print
    Debug.Print "[One Core Engine]"
    i = 0
    For Each token In mSapiTokenCategory.EnumerateTokens
        Debug.Print i & ":" & token.GetDescription
        i = i + 1
    Next
End Sub

'******************************************************************************
'* [概  要] GetMp3FileProperty
'* [詳  細] Mp3ファイルのプロパティを取得する｡
'*
'* @param mp3FilePath Mp3ファイルパス
'* @return DictionaryEx Mp3ファイルのプロパティ
'******************************************************************************
Public Function GetMp3FileProperty(ByVal mp3FilePath As String) As DictionaryEx
    Dim oFolder As Object: Set oFolder = SystemUtils.NameSpace(IO.GetFile(mp3FilePath).ParentFolder & "\")
    Dim dic As DictionaryEx: Set dic = New DictionaryEx
    dic.Add "File", oFolder.GetDetailsOf(oFolder.ParseName(IO.GetFile(mp3FilePath).Name), 0)       ' ファイル名
    dic.Add "Album", oFolder.GetDetailsOf(oFolder.ParseName(IO.GetFile(mp3FilePath).Name), 14)     ' アルバム
    dic.Add "Year", oFolder.GetDetailsOf(oFolder.ParseName(IO.GetFile(mp3FilePath).Name), 15)      ' 年
    dic.Add "Genre", oFolder.GetDetailsOf(oFolder.ParseName(IO.GetFile(mp3FilePath).Name), 16)     ' ジャンル
    dic.Add "Artist", oFolder.GetDetailsOf(oFolder.ParseName(IO.GetFile(mp3FilePath).Name), 20)    ' アーティスト名
    dic.Add "Title", oFolder.GetDetailsOf(oFolder.ParseName(IO.GetFile(mp3FilePath).Name), 21)     ' タイトル(曲名)
    dic.Add "TruckNo", oFolder.GetDetailsOf(oFolder.ParseName(IO.GetFile(mp3FilePath).Name), 26)   ' トラック番号
    dic.Add "Time", oFolder.GetDetailsOf(oFolder.ParseName(IO.GetFile(mp3FilePath).Name), 27)      ' 長さ(再生時間)
    dic.Add "Path", oFolder.GetDetailsOf(oFolder.ParseName(IO.GetFile(mp3FilePath).Name), 180)     ' パス
    dic.Add "Publish", oFolder.GetDetailsOf(oFolder.ParseName(IO.GetFile(mp3FilePath).Name), 195)  ' 発行元(発売レーベル)
    dic.Add "Composer", oFolder.GetDetailsOf(oFolder.ParseName(IO.GetFile(mp3FilePath).Name), 220) ' 作曲者
   Set GetMp3FileProperty = dic
End Function

'******************************************************************************
'* [概  要] CanMidiOut
'* [詳  細] MIDI音源があるかチェックする｡
'*          MIDI出力デバイス数を取得し、0より多ければMIDI出力可能とする。
'*
'* @return Boolean MIDI音源があるかどうか（True:有）
'******************************************************************************
Public Function CanMidiOut() As Boolean
    CanMidiOut = (midiOutGetNumDevs > 0)
End Function

'******************************************************************************
'* [概  要] OpenMidi
'* [詳  細] MIDIデバイスを開く｡
'*
'******************************************************************************
Public Sub OpenMidi()
    If CanMidiOut Then
        Call midiOutOpen(mMidiHandle, -1, 0, 0, 0)
        mIsMidiOpened = True
    Else
        Err.Raise 9999, "SoundUtils#OpenMidi", "MIDIデバイスが利用できません。"
    End If
End Sub

'******************************************************************************
'* [概  要] CloseMidi
'* [詳  細] MIDIデバイスを閉じる｡
'* [補  足] ブックを閉じる前に実行すること｡
'*
'******************************************************************************
Public Sub CloseMidi()
    If mIsMidiOpened Then
        Call midiOutClose(mMidiHandle)
        mIsMidiOpened = False
    End If
End Sub

'******************************************************************************
'* [概  要] GetMidiNote
'* [詳  細] MIDIノート番号を取得する｡
'* [参  考] https://computermusicresource.com/midikeys.html
'*
'* @param mdNote   MIDIノート番号（0オクターブ）
'* @param mdOctave オクターブ。-2～8の範囲で指定。指定しない場合、
'*                 または範囲外の場合、mMidiStdOctave の値が使用される。
'* @return Long MIDIノート番号
'******************************************************************************
Public Function GetMidiNote(ByVal mdNote As MidiNote0OctaveEnum, Optional ByVal mdOctave As Long = -9) As Long
    If mdOctave < -2 Or 8 < mdOctave Then
        mdOctave = mMidiStdOctave
    End If
    GetMidiNote = mdNote + (12 * mdOctave)
End Function

'******************************************************************************
'* [概  要] MidiNoteOn
'* [詳  細] MIDI音源で音を鳴らす｡
'* [補  足] 事前に OpenMidi を実行しておく必要がある｡
'* [参  考] https://codezine.jp/article/detail/172
'*          http://fft-0064.blogspot.com/2018/04/midi.html
'*
'* @param mdNote    MIDIノート番号（0オクターブ）
'* @param mdOctave  オクターブ。-2～8の範囲で指定。指定しない場合、
'*                  または範囲外の場合、mMidiStdOctave の値が使用される。
'* @param mdTimbre  音色番号
'* @param mdVolume  音量
'* @param mdChannel チャネル
'* @param mdTempo   テンポ
'******************************************************************************
Public Sub MidiNoteOn(ByVal mdNote As MidiNote0OctaveEnum, _
                      Optional ByVal mdOctave As Long = -9, _
                      Optional ByVal mdTimbre As Long = -1, _
                      Optional ByVal mdVolume As Long = -1, _
                      Optional ByVal mdChannel As Long = -1, _
                      Optional ByVal mdTempo As Long = -1)
    If Not mIsMidiOpened Then Exit Sub
    
    Dim noteNum As Long: noteNum = GetMidiNote(mdNote, mdOctave)
    If mdTimbre < 0 Or 127 < mdTimbre Then mdTimbre = mMidiStdTimbre
    If mdVolume < 0 Or 127 < mdVolume Then mdVolume = mMidiStdVolume
    If mdChannel < 0 Or 15 < mdChannel Then mdChannel = mMidiStdChannel
    If mdTempo < 0 Then mdTempo = mMidiStdTempo
                 
    ' 音色を指定 -- 音色番号（0-127） 音色変更:0xCn
    Call midiOutShortMsg(mMidiHandle, (mdTimbre) * (16 ^ 2) + (12 * 16) + mdChannel)
    
    ' 発音 -- 音量（0-127） ノート番号（0-127）  ノートオン（発音）:0x9n
    Call midiOutShortMsg(mMidiHandle, (mdVolume * (16 ^ 4)) + (noteNum * (16 ^ 2)) + ((9 * 16) + mdChannel))
    
    ' TODO: Sleep以外で音の長さを指定する方法
    Sleep 60000 / mdTempo
    
    ' 消音
    Call MidiNoteOff(mdNote, mdOctave, mdVolume, mdChannel)
    
    'Call Base.OnTimeForClass(60000 / mdTempo, SoundUtils, "MidiNoteOff", mdNote, mdOctave, mdVolume, mdChannel)
    
End Sub

'******************************************************************************
'* [概  要] MidiNoteOff
'* [詳  細] MIDI音源で音を止める｡
'* [補  足] 事前に OpenMidi を実行しておく必要がある｡
'*
'* @param mdNote    MIDIノート番号（0オクターブ）
'* @param mdOctave  オクターブ。-2～8の範囲で指定。指定しない場合、
'*                  または範囲外の場合、mMidiStdOctave の値が使用される。
'* @param mdVolume  音量
'* @param mdChannel チャネル
'******************************************************************************
Public Sub MidiNoteOff(ByVal mdNote As MidiNote0OctaveEnum, _
                       Optional ByVal mdOctave As Long = -9, _
                       Optional ByVal mdVolume As Long = -1, _
                       Optional ByVal mdChannel As Long = -1)
    If Not mIsMidiOpened Then Exit Sub

    Dim noteNum As Long: noteNum = GetMidiNote(mdNote, mdOctave)
    If mdVolume < 0 Or 127 < mdVolume Then mdVolume = mMidiStdVolume
    If mdChannel < 0 Or 15 < mdChannel Then mdChannel = mMidiStdChannel
                 
    ' 消音 -- 音量（0-127） ノート番号（0-127）  ノートオフ（消音）:0x8n
    Call midiOutShortMsg(mMidiHandle, (mdVolume * (16 ^ 4)) + (noteNum * (16 ^ 2)) + ((8 * 16) + mdChannel))
End Sub

