VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiWinmm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WindowAPI winmm クラス
'* [詳  細] WindowAPIのwinmm.dllの機能を提供する。
'*
'* [参  考] Win32API_PtrSafe.TXT
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Function mixerGetNumDevs Lib "winmm.dll" () As Long
Private Declare PtrSafe Function mixerGetDevCaps Lib "winmm.dll" Alias "mixerGetDevCapsA" (ByVal uMxId As LongPtr, pmxcaps As MIXERCAPS, ByVal cbmxcaps As Long) As Long
Private Declare PtrSafe Function mixerOpen Lib "winmm.dll" (phmx As LongPtr, ByVal uMxId As Long, ByVal dwCallback As LongPtr, ByVal dwInstance As LongPtr, ByVal fdwOpen As Long) As Long
Private Declare PtrSafe Function mixerClose Lib "winmm.dll" (ByVal hmx As LongPtr) As Long
Private Declare PtrSafe Function mixerMessage Lib "winmm.dll" (ByVal hmx As LongPtr, ByVal uMsg As Long, ByVal dwParam1 As LongPtr, ByVal dwParam2 As LongPtr) As Long
Private Declare PtrSafe Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hmxobj As LongPtr, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long

Private Declare PtrSafe Function mixerGetID Lib "winmm.dll" (ByVal hmxobj As LongPtr, pumxID As Long, ByVal fdwId As Long) As Long

Private Declare PtrSafe Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj As LongPtr, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long

Private Declare PtrSafe Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hmxobj As LongPtr, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long

Private Declare PtrSafe Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As LongPtr, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long

Private Declare PtrSafe Function joyGetPos Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFO) As Long
Private Declare PtrSafe Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFOEX) As Long

Private Declare PtrSafe Function midiStreamOpen Lib "winmm.dll" (phms As LongPtr, puDeviceID As Long, ByVal cMidi As Long, ByVal dwCallback As LongPtr, ByVal dwInstance As LongPtr, ByVal fdwOpen As Long) As Long
Private Declare PtrSafe Function midiStreamClose Lib "winmm.dll" (ByVal hms As LongPtr) As Long

Private Declare PtrSafe Function midiStreamProperty Lib "winmm.dll" (ByVal hms As LongPtr, lppropdata As Byte, ByVal dwProperty As Long) As Long
Private Declare PtrSafe Function midiStreamPosition Lib "winmm.dll" (ByVal hms As LongPtr, lpmmt As MMTIME, ByVal cbmmt As Long) As Long
Private Declare PtrSafe Function midiStreamOut Lib "winmm.dll" (ByVal hms As LongPtr, pmh As MIDIHDR, ByVal cbmh As Long) As Long
Private Declare PtrSafe Function midiStreamPause Lib "winmm.dll" (ByVal hms As LongPtr) As Long
Private Declare PtrSafe Function midiStreamRestart Lib "winmm.dll" (ByVal hms As LongPtr) As Long
Private Declare PtrSafe Function midiStreamStop Lib "winmm.dll" (ByVal hms As LongPtr) As Long
Private Declare PtrSafe Function midiConnect Lib "winmm.dll" (ByVal hmi As LongPtr, ByVal hmo As LongPtr, pReserved As Any) As Long
Private Declare PtrSafe Function midiDisconnect Lib "winmm.dll" (ByVal hmi As LongPtr, ByVal hmo As LongPtr, pReserved As Any) As Long

Private Declare PtrSafe Function CloseDriver Lib "winmm.dll" (ByVal hDriver As LongPtr, ByVal lParam1 As LongPtr, ByVal lParam2 As LongPtr) As LongPtr
Private Declare PtrSafe Function OpenDriver Lib "winmm.dll" (ByVal szDriverName As String, ByVal szSectionName As String, ByVal lParam2 As LongPtr) As LongPtr
Private Declare PtrSafe Function SendDriverMessage Lib "winmm.dll" (ByVal hDriver As LongPtr, ByVal message As Long, ByVal lParam1 As LongPtr, ByVal lParam2 As LongPtr) As LongPtr
Private Declare PtrSafe Function DrvGetModuleHandle Lib "winmm.dll" (ByVal hDriver As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDriverModuleHandle Lib "winmm.dll" (ByVal hDriver As LongPtr) As LongPtr
Private Declare PtrSafe Function DefDriverProc Lib "winmm.dll" (ByVal dwDriverIdentifier As LongPtr, ByVal hdrvr As LongPtr, ByVal uMsg As Long, ByVal lParam1 As LongPtr, ByVal lParam2 As LongPtr) As LongPtr

Private Declare PtrSafe Function mmsystemGetVersion Lib "winmm.dll" () As Long
Private Declare PtrSafe Sub OutputDebugStr Lib "winmm.dll" (ByVal lpszOutputString As String)

Private Declare PtrSafe Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As LongPtr, ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Private Declare PtrSafe Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveOutGetDevCapsA" (ByVal uDeviceID As LongPtr, lpCaps As WAVEOUTCAPS, ByVal uSize As Long) As Long

Private Declare PtrSafe Function waveOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As LongPtr, lpdwVolume As Long) As Long
Private Declare PtrSafe Function waveOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As LongPtr, ByVal dwVolume As Long) As Long

Private Declare PtrSafe Function waveOutGetErrorText Lib "winmm.dll" Alias "waveOutGetErrorTextA" (ByVal Err As Long, ByVal lpText As String, ByVal uSize As Long) As Long

Private Declare PtrSafe Function waveOutOpen Lib "winmm.dll" (lphWaveOut As LongPtr, ByVal uDeviceID As Long, lpFormat As WAVEFORMATEX, ByVal dwCallback As LongPtr, ByVal dwInstance As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function waveOutClose Lib "winmm.dll" (ByVal hWaveOut As LongPtr) As Long
Private Declare PtrSafe Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hWaveOut As LongPtr, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare PtrSafe Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hWaveOut As LongPtr, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare PtrSafe Function waveOutWrite Lib "winmm.dll" (ByVal hWaveOut As LongPtr, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare PtrSafe Function waveOutPause Lib "winmm.dll" (ByVal hWaveOut As LongPtr) As Long
Private Declare PtrSafe Function waveOutRestart Lib "winmm.dll" (ByVal hWaveOut As LongPtr) As Long
Private Declare PtrSafe Function waveOutReset Lib "winmm.dll" (ByVal hWaveOut As LongPtr) As Long
Private Declare PtrSafe Function waveOutBreakLoop Lib "winmm.dll" (ByVal hWaveOut As LongPtr) As Long
Private Declare PtrSafe Function waveOutGetPosition Lib "winmm.dll" (ByVal hWaveOut As LongPtr, lpInfo As MMTIME, ByVal uSize As Long) As Long
Private Declare PtrSafe Function waveOutGetPitch Lib "winmm.dll" (ByVal hWaveOut As LongPtr, lpdwPitch As Long) As Long
Private Declare PtrSafe Function waveOutSetPitch Lib "winmm.dll" (ByVal hWaveOut As LongPtr, ByVal dwPitch As Long) As Long
Private Declare PtrSafe Function waveOutGetPlaybackRate Lib "winmm.dll" (ByVal hWaveOut As LongPtr, lpdwRate As Long) As Long
Private Declare PtrSafe Function waveOutSetPlaybackRate Lib "winmm.dll" (ByVal hWaveOut As LongPtr, ByVal dwRate As Long) As Long
Private Declare PtrSafe Function waveOutGetID Lib "winmm.dll" (ByVal hWaveOut As LongPtr, lpuDeviceID As Long) As Long
Private Declare PtrSafe Function waveOutMessage Lib "winmm.dll" (ByVal hWaveOut As LongPtr, ByVal msg As Long, ByVal dw1 As LongPtr, ByVal dw2 As LongPtr) As Long
Private Declare PtrSafe Function waveInGetNumDevs Lib "winmm.dll" () As Long

Private Declare PtrSafe Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As LongPtr, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long

Private Declare PtrSafe Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal Err As Long, ByVal lpText As String, ByVal uSize As Long) As Long

Private Declare PtrSafe Function waveInOpen Lib "winmm.dll" (lphWaveIn As LongPtr, ByVal uDeviceID As Long, lpFormat As WAVEFORMATEX, ByVal dwCallback As LongPtr, ByVal dwInstance As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As LongPtr) As Long
Private Declare PtrSafe Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As LongPtr, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare PtrSafe Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As LongPtr, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare PtrSafe Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As LongPtr, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare PtrSafe Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As LongPtr) As Long
Private Declare PtrSafe Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As LongPtr) As Long
Private Declare PtrSafe Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As LongPtr) As Long
Private Declare PtrSafe Function waveInGetPosition Lib "winmm.dll" (ByVal hWaveIn As LongPtr, lpInfo As MMTIME, ByVal uSize As Long) As Long
Private Declare PtrSafe Function waveInGetID Lib "winmm.dll" (ByVal hWaveIn As LongPtr, lpuDeviceID As Long) As Long
Private Declare PtrSafe Function waveInMessage Lib "winmm.dll" (ByVal hWaveIn As LongPtr, ByVal msg As Long, ByVal dw1 As LongPtr, ByVal dw2 As LongPtr) As Long


Private Declare PtrSafe Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As LongPtr, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long

Private Declare PtrSafe Function midiOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As LongPtr, lpdwVolume As Long) As Long
Private Declare PtrSafe Function midiOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As LongPtr, ByVal dwVolume As Long) As Long

Private Declare PtrSafe Function midiOutGetErrorText Lib "winmm.dll" Alias "midiOutGetErrorTextA" (ByVal Err As Long, ByVal lpText As String, ByVal uSize As Long) As Long

Private Declare PtrSafe Function midiOutOpen Lib "winmm.dll" (lphMidiOut As LongPtr, ByVal uDeviceID As Long, ByVal dwCallback As LongPtr, ByVal dwInstance As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As LongPtr) As Long
Private Declare PtrSafe Function midiOutPrepareHeader Lib "winmm.dll" (ByVal hMidiOut As LongPtr, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Private Declare PtrSafe Function midiOutUnprepareHeader Lib "winmm.dll" (ByVal hMidiOut As LongPtr, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Private Declare PtrSafe Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As LongPtr, ByVal dwMsg As Long) As Long
Private Declare PtrSafe Function midiOutLongMsg Lib "winmm.dll" (ByVal hMidiOut As LongPtr, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Private Declare PtrSafe Function midiOutReset Lib "winmm.dll" (ByVal hMidiOut As LongPtr) As Long
Private Declare PtrSafe Function midiOutCachePatches Lib "winmm.dll" (ByVal hMidiOut As LongPtr, ByVal uBank As Long, lpPatchArray As Long, ByVal uFlags As Long) As Long
Private Declare PtrSafe Function midiOutCacheDrumPatches Lib "winmm.dll" (ByVal hMidiOut As LongPtr, ByVal uPatch As Long, lpKeyArray As Long, ByVal uFlags As Long) As Long
Private Declare PtrSafe Function midiOutGetID Lib "winmm.dll" (ByVal hMidiOut As LongPtr, lpuDeviceID As Long) As Long
Private Declare PtrSafe Function midiOutMessage Lib "winmm.dll" (ByVal hMidiOut As LongPtr, ByVal msg As Long, ByVal dw1 As LongPtr, ByVal dw2 As LongPtr) As Long
Private Declare PtrSafe Function midiInGetNumDevs Lib "winmm.dll" () As Long

Private Declare PtrSafe Function midiInGetDevCaps Lib "winmm.dll" Alias "midiInGetDevCapsA" (ByVal uDeviceID As LongPtr, lpCaps As MIDIINCAPS, ByVal uSize As Long) As Long

Private Declare PtrSafe Function midiInGetErrorText Lib "winmm.dll" Alias "midiInGetErrorTextA" (ByVal Err As Long, ByVal lpText As String, ByVal uSize As Long) As Long

Private Declare PtrSafe Function midiInOpen Lib "winmm.dll" (lphMidiIn As LongPtr, ByVal uDeviceID As Long, ByVal dwCallback As LongPtr, ByVal dwInstance As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function midiInClose Lib "winmm.dll" (ByVal hMidiIn As LongPtr) As Long
Private Declare PtrSafe Function midiInPrepareHeader Lib "winmm.dll" (ByVal hMidiIn As LongPtr, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
Private Declare PtrSafe Function midiInUnprepareHeader Lib "winmm.dll" (ByVal hMidiIn As LongPtr, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
Private Declare PtrSafe Function midiInAddBuffer Lib "winmm.dll" (ByVal hMidiIn As LongPtr, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
Private Declare PtrSafe Function midiInStart Lib "winmm.dll" (ByVal hMidiIn As LongPtr) As Long
Private Declare PtrSafe Function midiInStop Lib "winmm.dll" (ByVal hMidiIn As LongPtr) As Long
Private Declare PtrSafe Function midiInReset Lib "winmm.dll" (ByVal hMidiIn As LongPtr) As Long
Private Declare PtrSafe Function midiInGetID Lib "winmm.dll" (ByVal hMidiIn As LongPtr, lpuDeviceID As Long) As Long
Private Declare PtrSafe Function midiInMessage Lib "winmm.dll" (ByVal hMidiIn As LongPtr, ByVal msg As Long, ByVal dw1 As LongPtr, ByVal dw2 As LongPtr) As Long

Private Declare PtrSafe Function auxGetNumDevs Lib "winmm.dll" () As Long
Private Declare PtrSafe Function auxGetDevCaps Lib "winmm.dll" Alias "auxGetDevCapsA" (ByVal uDeviceID As LongPtr, lpCaps As AUXCAPS, ByVal uSize As Long) As Long

Private Declare PtrSafe Function auxSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare PtrSafe Function auxGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Private Declare PtrSafe Function auxOutMessage Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal msg As Long, ByVal dw1 As LongPtr, ByVal dw2 As LongPtr) As Long

Private Declare PtrSafe Function timeGetSystemTime Lib "winmm.dll" (lpTime As MMTIME, ByVal uSize As Long) As Long
Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
Private Declare PtrSafe Function timeSetEvent Lib "winmm.dll" (ByVal uDelay As Long, ByVal uResolution As Long, ByVal lpFunction As LongPtr, ByVal dwUser As LongPtr, ByVal uFlags As Long) As Long
Private Declare PtrSafe Function timeKillEvent Lib "winmm.dll" (ByVal uID As Long) As Long
Private Declare PtrSafe Function timeGetDevCaps Lib "winmm.dll" (lpTimeCaps As TIMECAPS, ByVal uSize As Long) As Long
Private Declare PtrSafe Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare PtrSafe Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

Private Declare PtrSafe Function joyGetDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" (ByVal id As LongPtr, lpCaps As JOYCAPS, ByVal uSize As Long) As Long

Private Declare PtrSafe Function joyGetNumDevs Lib "winmm.dll" Alias "joyGetNumDev" () As Long
Private Declare PtrSafe Function joyGetThreshold Lib "winmm.dll" (ByVal id As Long, lpuThreshold As Long) As Long
Private Declare PtrSafe Function joyReleaseCapture Lib "winmm.dll" (ByVal id As Long) As Long
Private Declare PtrSafe Function joySetCapture Lib "winmm.dll" (ByVal hWnd As LongPtr, ByVal uID As Long, ByVal uPeriod As Long, ByVal bChanged As Long) As Long
Private Declare PtrSafe Function joySetThreshold Lib "winmm.dll" (ByVal id As Long, ByVal uThreshold As Long) As Long

Private Declare PtrSafe Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long

Private Declare PtrSafe Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, lpmmioinfo As MMIOINFO, ByVal dwOpenFlags As Long) As LongPtr

Private Declare PtrSafe Function mmioRename Lib "winmm.dll" Alias "mmioRenameA" (ByVal szFileName As String, ByVal SzNewFileName As String, lpmmioinfo As MMIOINFO, ByVal dwRenameFlags As Long) As Long

Private Declare PtrSafe Function mmioClose Lib "winmm.dll" (ByVal hmmio As LongPtr, ByVal uFlags As Long) As Long
Private Declare PtrSafe Function mmioRead Lib "winmm.dll" (ByVal hmmio As LongPtr, ByVal pch As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function mmioWrite Lib "winmm.dll" (ByVal hmmio As LongPtr, ByVal pch As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function mmioSeek Lib "winmm.dll" (ByVal hmmio As LongPtr, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Private Declare PtrSafe Function mmioGetInfo Lib "winmm.dll" (ByVal hmmio As LongPtr, lpmmioinfo As MMIOINFO, ByVal uFlags As Long) As Long
Private Declare PtrSafe Function mmioSetInfo Lib "winmm.dll" (ByVal hmmio As LongPtr, lpmmioinfo As MMIOINFO, ByVal uFlags As Long) As Long
Private Declare PtrSafe Function mmioSetBuffer Lib "winmm.dll" (ByVal hmmio As LongPtr, ByVal pchBuffer As String, ByVal cchBuffer As Long, ByVal uFlags As Long) As Long
Private Declare PtrSafe Function mmioFlush Lib "winmm.dll" (ByVal hmmio As LongPtr, ByVal uFlags As Long) As Long
Private Declare PtrSafe Function mmioAdvance Lib "winmm.dll" (ByVal hmmio As LongPtr, lpmmioinfo As MMIOINFO, ByVal uFlags As Long) As Long
Private Declare PtrSafe Function mmioSendMessage Lib "winmm.dll" (ByVal hmmio As LongPtr, ByVal uMsg As Long, ByVal lParam1 As LongPtr, ByVal lParam2 As LongPtr) As LongPtr

Private Declare PtrSafe Function mmioDescend Lib "winmm.dll" (ByVal hmmio As LongPtr, lpck As MMCKINFO, lpckParent As MMCKINFO, ByVal uFlags As Long) As Long
Private Declare PtrSafe Function mmioAscend Lib "winmm.dll" (ByVal hmmio As LongPtr, lpck As MMCKINFO, ByVal uFlags As Long) As Long
Private Declare PtrSafe Function mmioCreateChunk Lib "winmm.dll" (ByVal hmmio As LongPtr, lpck As MMCKINFO, ByVal uFlags As Long) As Long

' MCI functions

Private Declare PtrSafe Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As LongPtr, ByVal dwParam2 As Any) As Long

Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As LongPtr) As Long

Private Declare PtrSafe Function mciGetCreatorTask Lib "winmm.dll" (ByVal wDeviceID As Long) As LongPtr

Private Declare PtrSafe Function mciGetDeviceID Lib "winmm.dll" Alias "mciGetDeviceIDA" (ByVal lpstrName As String) As Long

Private Declare PtrSafe Function mciGetDeviceIDFromElementID Lib "winmm.dll" Alias "mciGetDeviceIDFromElementIDA" (ByVal dwElementID As Long, ByVal lpstrType As String) As Long

Private Declare PtrSafe Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

Private Declare PtrSafe Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long


Private Declare PtrSafe Function mciGetYieldProc Lib "winmm" (ByVal mciId As Long, pdwYieldData As Long) As LongPtr
Private Declare PtrSafe Function mciSetYieldProc Lib "winmm" (ByVal mciId As Long, ByVal fpYieldProc As LongPtr, ByVal dwYieldData As Long) As Long
Private Declare PtrSafe Function midiOutGetNumDevs Lib "winmm" () As Long
Private Declare PtrSafe Function mmioInstallIOProcA Lib "winmm" (ByVal fccIOProc As Long, ByVal pIOProc As LongPtr, ByVal dwFlags As Long) As LongPtr





