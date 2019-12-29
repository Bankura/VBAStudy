Attribute VB_Name = "LogUtil"
Option Explicit

'==============================================================================
'���O�o�͗p���[�e�B���e�B�֐����W���[��
'==============================================================================


'******************************************************************************
'���O�o�͋@�\���ʒ萔
'******************************************************************************
Public Const CMN_LOG_LEVEL_ERROR = "ERROR"
Public Const CMN_LOG_LEVEL_INFO = "INFO"
Public Const CMN_LOG_LEVEL_DEBUG = "DEBUG"

'******************************************************************************
'���O�o�͋@�\���ʕϐ�
'******************************************************************************
' ���O�o�̓t���O
'  0 �c �o�͂��Ȃ�
'  1 �c �uDEBUG�v�ȏ�o��
'  2 �c �uINFO�v�ȏ���o��
'  3 �c �uERROR�v�̂ݏo��
Public cmn_outLogFlg As Integer

' �C�~�f�B�G�C�g�E�B���h�E�o�͉ۃt���O
' True �c �o�͂���
' False �c �o�͂��Ȃ�
Public cmn_iwFlg As Boolean

' ���O�o�͐�f�B���N�g��
Public cmn_strLogDirPath As String

' ���O�t�@�C����
Public cmn_strLogFileName As String

' �X�e�[�^�X�o�[�\���ۃt���O
' True �c �o�͂���
' False �c �o�͂��Ȃ�
Public cmn_sbFlg As Boolean

'******************************************************************************
' [�֐���] ���O�o�͏����ݒ�
' [���@��] ���O���o�͂��邽�߂̏����ݒ���s���֐��B
' [���@��] outLogFlg      ���O�o�̓t���O
'          iwFlg          �C�~�f�B�G�C�g�E�B���h�E�o��
'                         �ۃt���O
'          strLogDirPath  ���O�o�͐�f�B���N�g��
'          strLogFileName ���O�t�@�C����
'          lrotateFileSize ���O���[�e�[�g�t�@�C���T�C�Y
'******************************************************************************
Public Sub InitLogSetting(outLogFlg As Integer, iwFlg As Boolean, _
                          strLogDirPath As String, strLogFileName As String, _
                          Optional divPersonFlg As Boolean = False, _
                          Optional lrotateFileSize As Long = 10485760)
    cmn_outLogFlg = outLogFlg
    cmn_iwFlg = iwFlg
    cmn_strLogDirPath = AddPathSeparator(strLogDirPath)
    
    If divPersonFlg Then
        cmn_strLogFileName = Mid(strLogFileName, 1, InStrRev(strLogFileName, ".") - 1) & "_" & _
                             CStr(CreateObject("WScript.Network").UserName) & _
                             Mid(strLogFileName, InStrRev(strLogFileName, "."), _
                                 Len(strLogFileName) - InStrRev(strLogFileName, ".") + 1)
    Else
        cmn_strLogFileName = strLogFileName
    End If
    
    ' ���O�T�C�Y���w�肵���T�C�Y�𒴂���Ɛ؂�ւ����s���i�f�t�H���g10MB�j
    If Dir(cmn_strLogDirPath & cmn_strLogFileName) <> "" Then
        If FileLen(cmn_strLogDirPath & cmn_strLogFileName) > lrotateFileSize Then
            Name cmn_strLogDirPath & cmn_strLogFileName _
                As cmn_strLogDirPath & cmn_strLogFileName & "." & Format(Now, "yyyymmdd-hhmm")
        End If
    End If
End Sub

'******************************************************************************
' [�֐���] ���O�o�͊֐�
' [���@��] ���O���o�͂���֐��B
' [���@��] strLogText ���O�ɏo�͂�����e
'******************************************************************************
Public Sub OutLog(strLogText As String, strLogLevel As String)
    On Error GoTo ErrorHandler
    Dim lngLogLevel As Long
    Dim lngFileNum As Long
    Dim strLogFile As String
    Dim strOutMsg As String
    
    If cmn_outLogFlg = 0 And cmn_iwFlg = False Then
        Exit Sub
    End If

    Select Case strLogLevel
        Case CMN_LOG_LEVEL_ERROR
            lngLogLevel = 3
        Case CMN_LOG_LEVEL_INFO
            lngLogLevel = 2
        Case CMN_LOG_LEVEL_DEBUG
            lngLogLevel = 1
        Case Else
            lngLogLevel = 0
    End Select
    
    strOutMsg = GetNowWithMSec() & " [" & strLogLevel & "] " & strLogText
    If lngLogLevel >= cmn_outLogFlg Then
        strLogFile = cmn_strLogDirPath & cmn_strLogFileName
        lngFileNum = FreeFile()
        Open strLogFile For Append As #lngFileNum
        Print #lngFileNum, strOutMsg
        Close #lngFileNum
    End If
    If cmn_iwFlg Then
        Debug.Print strOutMsg
    End If
    Exit Sub
ErrorHandler:
    Debug.Print GetNowWithMSec() & " [FATAL] �yOutLog�z�G���[�������FNumber=" & Err.Number & " Description=" & Err.Description
    Debug.Print GetNowWithMSec() & " �yCAUTION�z���O�o�͕s�B�C�~�f�B�G�C�g�E�B���h�E�o�͂ɐ؂�ւ��܂��B"
    cmn_iwFlg = True
    cmn_outLogFlg = 9
    Debug.Print strOutMsg
End Sub

'******************************************************************************
' [�֐���] OutStatusBar �X�e�[�^�X�o�[�\���֐�
' [���@��] �X�e�[�^�X�o�[��\������֐��B
' [���@��] varText �X�e�[�^�X�o�[�ɕ\��������e
'******************************************************************************
Public Sub OutStatusBar(varText As Variant)
    If cmn_sbFlg Then
        Application.StatusBar = varText
    End If
End Sub


'******************************************************************************
' [�֐���] GetNowWithMSec
' [���@��] ���ݎ����̔N���������b�~���b���uYYYY/MM/DD HH:NN:SS.000�v�`����
'          ������Ƃ��Ď擾����B
' [���@��] �Ȃ�
' [�߂�l] String �uyyyy/mm/dd hh:nn:ss.000�v�`�����ݎ���������
'******************************************************************************
Private Function GetNowWithMSec() As String
    Dim dblTimer As Double
    Dim hour As Integer
    Dim minute As Integer
    Dim second As Integer
    Dim mSec As Double
    Dim rtn As String
    
    dblTimer = CDbl(Timer)
    hour = dblTimer \ 3600
    minute = (dblTimer Mod 3600) \ 60
    second = dblTimer Mod 60
    mSec = Fix((dblTimer - Fix(dblTimer)) * 1000)
    
    rtn = Format(Now, "yyyy/mm/dd") & " " & Format(hour, "00") & ":" & _
          Format(minute, "00") & ":" & Format(second, "00") & "." & Format(mSec, "000")
          
    GetNowWithMSec = rtn
End Function




