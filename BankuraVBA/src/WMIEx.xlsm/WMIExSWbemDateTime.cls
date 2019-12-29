VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WMIExSWbemDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] WbemScripting.SWbemDateTime�N���X�̃��b�v�E�g���N���X
'* [��  ��] WbemScripting.SWbemDateTime�N���X��Wrapper�Ƃ��Ă̋@�\��񋟂��鑼�A
'*          WbemScripting.SWbemDateTime�N���X���g�p�������[�e�B���e�B��񋟂���B
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [��  ��]
' WbemScripting.SWbemDateTime�N���X�ɂ��āB
' <https://msdn.microsoft.com/ja-jp/windows/aa393687(v=vs.80)>
' Date & Time
'------------------------------------------------------------------------------

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemDateTime
Private mOrigin As Object
    
'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************
'*-----------------------------------------------------------------------------
'* �I���W�i��Object �v���p�e�B
'*-----------------------------------------------------------------------------
Property Get Origin() As Object
    Set Origin = mOrigin
End Property
Property Set Origin(arg As Object)
    Set mOrigin = arg
End Property

'*-----------------------------------------------------------------------------
'* Day �v���p�e�B
'*
'* [��  ��]
'* �EThe Day component of the value (must be in the range 1-31, or 0-999999 for interval values).
'*-----------------------------------------------------------------------------
Property Get Day() As Long
    Day = mOrigin.Day
End Property
Property Let Day(arg As Long)
    mOrigin.Day = arg
End Property

'*-----------------------------------------------------------------------------
'* DaySpecified �v���p�e�B
'*
'* [��  ��]
'* �EWhether the Day component is specified.
'*-----------------------------------------------------------------------------
Property Get DaySpecified() As Boolean
    DaySpecified = mOrigin.DaySpecified
End Property
Property Let DaySpecified(arg As Boolean)
    mOrigin.DaySpecified = arg
End Property
    
'*-----------------------------------------------------------------------------
'* Hours �v���p�e�B
'*
'* [��  ��]
'* �EThe Hours component of the value (must be in the range 0-23).
'*-----------------------------------------------------------------------------
Property Get Hours() As Long
    Hours = mOrigin.Hours
End Property
Property Let Hours(arg As Long)
    mOrigin.Hours = arg
End Property

'*-----------------------------------------------------------------------------
'* HoursSpecified �v���p�e�B
'*
'* [��  ��]
'* �EWhether the Hours component is specified.
'*-----------------------------------------------------------------------------
Property Get HoursSpecified() As Boolean
    HoursSpecified = mOrigin.HoursSpecified
End Property
Property Let HoursSpecified(arg As Boolean)
    mOrigin.HoursSpecified = arg
End Property
    
'*-----------------------------------------------------------------------------
'* IsInterval �v���p�e�B
'*
'* [��  ��]
'* �EIndicates whether this value describes an absolute date and time or is an interval.
'*-----------------------------------------------------------------------------
Property Get IsInterval() As Boolean
    IsInterval = mOrigin.IsInterval
End Property
Property Let IsInterval(arg As Boolean)
    mOrigin.IsInterval = arg
End Property
    
'*-----------------------------------------------------------------------------
'* Microseconds �v���p�e�B
'*
'* [��  ��]
'* �EThe Microseconds component of the value (must be in the range 0-999999).
'*-----------------------------------------------------------------------------
Property Get Microseconds() As Long
    Microseconds = mOrigin.Microseconds
End Property
Property Let Microseconds(arg As Long)
    mOrigin.Microseconds = arg
End Property
    
'*-----------------------------------------------------------------------------
'* MicrosecondsSpecified �v���p�e�B
'*
'* [��  ��]
'* �EWhether the Microseconds component is specified.
'*-----------------------------------------------------------------------------
Property Get MicrosecondsSpecified() As Boolean
    MicrosecondsSpecified = mOrigin.MicrosecondsSpecified
End Property
Property Let MicrosecondsSpecified(arg As Boolean)
    mOrigin.MicrosecondsSpecified = arg
End Property
    
'*-----------------------------------------------------------------------------
'* Minutes �v���p�e�B
'*
'* [��  ��]
'* �EThe Minutes component of the value (must be in the range 0-59).
'*-----------------------------------------------------------------------------
Property Get Minutes() As Long
    Minutes = mOrigin.Minutes
End Property
Property Let Minutes(arg As Long)
    mOrigin.Minutes = arg
End Property
    
'*-----------------------------------------------------------------------------
'* MinutesSpecified �v���p�e�B
'*
'* [��  ��]
'* �EWhether the Minutes component is specified.
'*-----------------------------------------------------------------------------
Property Get MinutesSpecified() As Boolean
    MinutesSpecified = mOrigin.MinutesSpecified
End Property
Property Let MinutesSpecified(arg As Boolean)
    mOrigin.MinutesSpecified = arg
End Property
    
'*-----------------------------------------------------------------------------
'* Month �v���p�e�B
'*
'* [��  ��]
'* �EThe Month component of the value (must be in the range 1-12).
'*-----------------------------------------------------------------------------
Property Get Month() As Long
    Month = mOrigin.Month
End Property
Property Let Month(arg As Long)
    mOrigin.Month = arg
End Property
    
'*-----------------------------------------------------------------------------
'* MonthSpecified �v���p�e�B
'*
'* [��  ��]
'* �EWhether the Month component is specified.
'*-----------------------------------------------------------------------------
Property Get MonthSpecified() As Boolean
    MonthSpecified = mOrigin.MonthSpecified
End Property
Property Let MonthSpecified(arg As Boolean)
    mOrigin.MonthSpecified = arg
End Property

    
'*-----------------------------------------------------------------------------
'* Seconds �v���p�e�B
'*
'* [��  ��]
'* �EThe Seconds component of the value (must be in the range 0-59).
'*-----------------------------------------------------------------------------
Property Get Seconds() As Long
    Seconds = mOrigin.Seconds
End Property
Property Let Seconds(arg As Long)
    mOrigin.Seconds = arg
End Property
    
'*-----------------------------------------------------------------------------
'* SecondsSpecified �v���p�e�B
'*
'* [��  ��]
'* �EWhether the Seconds component is specified.
'*-----------------------------------------------------------------------------
Property Get SecondsSpecified() As Boolean
    SecondsSpecified = mOrigin.SecondsSpecified
End Property
Property Let SecondsSpecified(arg As Boolean)
    mOrigin.SecondsSpecified = arg
End Property

'*-----------------------------------------------------------------------------
'* UTC �v���p�e�B
'*
'* [��  ��]
'* �EThe UTC component of the value (must be in the range -720 to 720).
'*-----------------------------------------------------------------------------
Property Get UTC() As Long
    UTC = mOrigin.UTC
End Property
Property Let UTC(arg As Long)
    mOrigin.UTC = arg
End Property
    
'*-----------------------------------------------------------------------------
'* UTCSpecified �v���p�e�B
'*
'* [��  ��]
'* �EWhether the UTC component is specified.
'*-----------------------------------------------------------------------------
Property Get UTCSpecified() As Boolean
    UTCSpecified = mOrigin.UTCSpecified
End Property
Property Let UTCSpecified(arg As Boolean)
    mOrigin.UTCSpecified = arg
End Property

'*-----------------------------------------------------------------------------
'* Value �v���p�e�B�i���胁���o�[�j
'*
'* [��  ��]
'* �EThe DMTF datetime.
'*-----------------------------------------------------------------------------
Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    Value = mOrigin.Value
End Property
Property Let Value(arg As String)
    mOrigin.Value = arg
End Property
   
    
'*-----------------------------------------------------------------------------
'* Year �v���p�e�B
'*
'* [��  ��]
'* �EThe Year component of the value (must be in the range 0-9999).
'*-----------------------------------------------------------------------------
Property Get Year() As Long
    Year = mOrigin.Year
End Property
Property Let Year(arg As Long)
    mOrigin.Year = arg
End Property
    
'*-----------------------------------------------------------------------------
'* YearSpecified �v���p�e�B
'*
'* [��  ��]
'* �EWhether the Year component is specified.
'*-----------------------------------------------------------------------------
Property Get YearSpecified() As Boolean
    YearSpecified = mOrigin.YearSpecified
End Property
Property Let YearSpecified(arg As Boolean)
    mOrigin.YearSpecified = arg
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("WbemScripting.SWbemDateTime")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] GetFileTime ���\�b�h
'* [��  ��] GetFileTime �̃��b�p�[���\�b�h�B
'*          Retrieve value in FILETIME compatible string representation.
'*
'* @param bIsLocal xx
'* @return String
'*
'******************************************************************************
Function GetFileTime(Optional bIsLocal As Boolean = True) As String
    GetFileTime = mOrigin.GetFileTime(bIsLocal)
End Function

'******************************************************************************
'* [�T  �v] GetVarDate ���\�b�h
'* [��  ��] GetVarDate �̃��b�p�[���\�b�h�B
'*          Retrieve value in Variant compatible (VT_DATE) format.
'*
'* @param bIsLocal xx
'* @return Date
'*
'******************************************************************************
Function GetVarDate(Optional bIsLocal As Boolean = True) As Date
    GetVarDate = mOrigin.GetVarDate(bIsLocal)
End Function

'******************************************************************************
'* [�T  �v] SetFileTime ���\�b�h
'* [��  ��] SetFileTime �̃��b�p�[���\�b�h�B
'*          Set the value using FILETIME compatible string representation.
'*
'* @param strFileTime xx
'* @param bIsLocal xx
'*
'******************************************************************************
Sub SetFileTime(strFileTime As String, Optional bIsLocal As Boolean = True)
    Call mOrigin.SetFileTime(strFileTime, bIsLocal)
End Sub

'******************************************************************************
'* [�T  �v] SetFileTime ���\�b�h
'* [��  ��] SetFileTime �̃��b�p�[���\�b�h�B
'*          Set the value using Variant compatible (VT_DATE) format.
'*
'* @param dVarDate xx
'* @param bIsLocal xx
'*
'******************************************************************************
Sub SetVarDate(dVarDate As Date, Optional bIsLocal As Boolean = True)
    Call mOrigin.SetVarDate(dVarDate, bIsLocal)
End Sub

'*-----------------------------------------------------------------------------
'* �g�����\�b�h
'*-----------------------------------------------------------------------------

'******************************************************************************
'* [�T  �v] GetUTCDate ���\�b�h
'* [��  ��] ������ݒ肵�AUTC�����̒l���擾����B
'*
'* @param targetDate �ݒ肷�����
'* @return UTC����
'*
'******************************************************************************
Function GetUTCDate(targetDate As Date) As String
    Me.SetVarDate targetDate, True
    GetUTCDate = Me.Value
End Function

'******************************************************************************
'* [�T  �v] UTCtoJST ���\�b�h
'* [��  ��] UTC�����l��W�����Ԃɒ����ē��t�^�ɕϊ�����B
'* [���p��] <https://selifelog.com/blog-entry-312.html>
'*
'* @param strUTCDate UTC�����l������
'* @return �W������
'*
'******************************************************************************
Function UTCtoJST(strUTCDate As String) As Date
    Dim strYear As String, strMonth As String
    Dim strDay As String, strHour As String, strMinute As String
    Dim strSec As String, JSTDateTime As Date
    '�����Ƃ��Ď󂯎�������t����������(YYYYMMDDHHMMSS)��
    '�N�A���A���A���A���A�b�Ƀo�����Ċe�ϐ��ɑ��
    strYear = Left(strUTCDate, 4)
    strMonth = Mid(strUTCDate, 5, 2)
    strDay = Mid(strUTCDate, 7, 2)
    strHour = Mid(strUTCDate, 9, 2)
    strMinute = Mid(strUTCDate, 11, 2)
    strSec = Mid(strUTCDate, 13, 2)
    '�����������̂�YYYY/MM/DD HH:MM:SS�ɍ��ς��ē��t�^�ɕϊ�����
    JSTDateTime = CDate(strYear & "/" & strMonth & "/" & strDay _
                    & " " & strHour & ":" & strMinute & ":" & strSec)
    'UTC�����l��9���ԉ��Z���ē��{�̕W�����ɒ������l��Ԃ�
    UTCtoJST = JSTDateTime + TimeValue("9:00:00")
End Function