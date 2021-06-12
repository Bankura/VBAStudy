VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTimeUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] ���t�E���ԑ��샆�[�e�B���e�B�N���X
'* [��  ��] ���t�E���ԑ���Ɋւ��郆�[�e�B���e�B��񋟂���B
'*          Static�N���X�iAttribute VB_PredeclaredId = True�j�Ƃ���B
'*
'* [�Q  �l] https://www.take1mg.com/plus1/2018/04/excel-vba-datecalculatorclass/
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI��`
'******************************************************************************

'******************************************************************************
'* Enum��`
'******************************************************************************

'******************************************************************************
'* �萔��`
'******************************************************************************
Private Const HOLIDAY_URL As String = "https://holidays-jp.github.io/api/v1/date.csv"
Private Const DATE_FORMAT_YYYYMMDD As String = "yyyymmdd"
Private Const DATE_FORMAT_YYYYMMDD_SLASH As String = "yyyy/mm/dd"
Private Const DATE_FORMAT_YYYYMMDD_HYPHEN As String = "yyyy-mm-dd"
Private Const DATE_FORMAT_YYYYMMDD_DOT As String = "yyyy.mm.dd"
Private Const DATE_FORMAT_YYYYMMDDHHNNSS As String = "yyyymmddhhnnss"
Private Const DATE_FORMAT_YYYYMMDDHHNNSS_SLASH As String = "yyyy/mm/dd hh:nn:ss"
Private Const DATE_FORMAT_YYYYMMDDHHNNSS_HYPHEN As String = "yyyy-mm-dd hh:nn:ss"
Private Const DATE_FORMAT_YYYYMMDDHHNNSS_DOT As String = "yyyy.mm.dd hh:nn:ss"
Private Const DATE_FORMAT_WAREKI As String = "ggge�Nmm��dd��"
Private Const DATE_FORMAT_WAREKI_YOUBI As String = "ggge�Nmm��dd��(aaa)"

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
Private mHolidayList As Collection

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************

'*-----------------------------------------------------------------------------
'* Today �v���p�e�B
'*
'*-----------------------------------------------------------------------------
Public Property Get Today() As Date
    Today = Date
End Property

'*-----------------------------------------------------------------------------
'* Yesterday �v���p�e�B
'*
'*-----------------------------------------------------------------------------
Public Property Get Yesterday() As Date
    Yesterday = Date - 1
End Property

'*-----------------------------------------------------------------------------
'* Tomorrow �v���p�e�B
'*
'*-----------------------------------------------------------------------------
Public Property Get Tomorrow() As Date
    Tomorrow = Date + 1
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
End Sub

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] BeginOfMonth
'* [��  ��] �w�肵�����t�̑����錎�̊J�n�����擾����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param dt �Ώۓ��t
'* @return Date �����̓��t
'******************************************************************************
Public Function BeginOfMonth(ByVal dt As Date) As Date
    BeginOfMonth = Core.BeginOfMonth(dt)
End Function

'******************************************************************************
'* [�T  �v] EndOfMonth
'* [��  ��] �w�肵�����t�̑����錎�̍ŏI�����擾����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param dt �Ώۓ��t
'* @return Date �����̓��t
'******************************************************************************
Public Function EndOfMonth(ByVal dt As Date) As Date
    EndOfMonth = Core.EndOfMonth(dt)
End Function

'******************************************************************************
'* [�T  �v] BeginOfWeek
'* [��  ��] �w�肵�����t�̑�����T�̊J�n�����擾����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param dt �Ώۓ��t
'* @param fstDayOfWeek �T�̎n�܂�̗j��
'* @return Date �T���̓��t
'******************************************************************************
Public Function BeginOfWeek(ByVal dt As Date, Optional fstDayOfWeek As VbDayOfWeek = vbSunday) As Date
    BeginOfWeek = Core.BeginOfWeek(dt, fstDayOfWeek)
End Function

'******************************************************************************
'* [�T  �v] EndOfWeek
'* [��  ��] �w�肵�����t�̑�����T�̍ŏI�����擾����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param dt �Ώۓ��t
'* @param fstDayOfWeek �T�̎n�܂�̗j��
'* @return Date �T���̓��t
'******************************************************************************
Public Function EndOfWeek(ByVal dt As Date, Optional fstDayOfWeek As VbDayOfWeek = vbSunday) As Date
    EndOfWeek = Core.EndOfWeek(dt, fstDayOfWeek)
End Function

'******************************************************************************
'* [�T  �v] IsLeapYear
'* [��  ��] �[�N���ǂ����`�F�b�N����B
'*
'* @param yearValue �Ώ۔N
'* @return Boolean �`�F�b�N���ʁiTrue�F�[�N�j
'******************************************************************************
Public Function IsLeapYear(ByVal yearValue As Integer) As Boolean
    If ((yearValue Mod 4) = 0 And (yearValue Mod 100) <> 0 Or (yearValue Mod 400) = 0) Then
        IsLeapYear = True
        Exit Function
    End If
    IsLeapYear = False
End Function

'******************************************************************************
'* [�T  �v] ToString
'* [��  ��] Date�^�̓��t����t������ɕϊ�����B
'*
'* @param dt Date�^���t
'* @param dateStrFormat �t�H�[�}�b�g
'* @return String ���t������
'******************************************************************************
Public Function ToString(ByVal dt As Date, Optional ByVal dateStrFormat As String = DATE_FORMAT_YYYYMMDDHHNNSS_SLASH) As String
    ToString = Format(dt, dateStrFormat)
End Function

'******************************************************************************
'* [�T  �v] ToStringOfYYYYMMDD
'* [��  ��] Date�^�̓��t����t������iyyyymmdd�j�ɕϊ�����B
'*
'* @param dt Date�^���t
'* @return String ���t������
'******************************************************************************
Public Function ToStringOfYYYYMMDD(ByVal dt As Date) As String
    ToStringOfYYYYMMDD = Format(dt, DATE_FORMAT_YYYYMMDD)
End Function

'******************************************************************************
'* [�T  �v] ToStringOfYYYYMMDDSlash
'* [��  ��] Date�^�̓��t����t������iyyyy/mm/dd�j�ɕϊ�����B
'*
'* @param dt Date�^���t
'* @return String ���t������
'******************************************************************************
Public Function ToStringOfYYYYMMDDSlash(ByVal dt As Date) As String
    ToStringOfYYYYMMDDSlash = Format(dt, DATE_FORMAT_YYYYMMDD_SLASH)
End Function

'******************************************************************************
'* [�T  �v] ToStringOfYYYYMMDDHyphen
'* [��  ��] Date�^�̓��t����t������iyyyy-mm-dd�j�ɕϊ�����B
'*
'* @param dt Date�^���t
'* @return String ���t������
'******************************************************************************
Public Function ToStringOfYYYYMMDDHyphen(ByVal dt As Date) As String
    ToStringOfYYYYMMDDHyphen = Format(dt, DATE_FORMAT_YYYYMMDD_HYPHEN)
End Function

'******************************************************************************
'* [�T  �v] ToStringOfYYYYMMDDDot
'* [��  ��] Date�^�̓��t����t������iyyyy.mm.dd�j�ɕϊ�����B
'*
'* @param dt Date�^���t
'* @return String ���t������
'******************************************************************************
Public Function ToStringOfYYYYMMDDDot(ByVal dt As Date) As String
    ToStringOfYYYYMMDDDot = Format(dt, DATE_FORMAT_YYYYMMDD_DOT)
End Function

'******************************************************************************
'* [�T  �v] ToStringOfYYYYMMDDHHNNSS
'* [��  ��] Date�^�̓��t����t������iyyyymmddhhnnss�j�ɕϊ�����B
'*
'* @param dt Date�^���t
'* @return String ���t������
'******************************************************************************
Public Function ToStringOfYYYYMMDDHHNNSS(ByVal dt As Date) As String
    ToStringOfYYYYMMDDHHNNSS = Format(dt, DATE_FORMAT_YYYYMMDDHHNNSS)
End Function

'******************************************************************************
'* [�T  �v] ToStringOfYYYYMMDDHHNNSSSlash
'* [��  ��] Date�^�̓��t����t������iyyyy/mm/dd hh:nn:ss�j�ɕϊ�����B
'*
'* @param dt Date�^���t
'* @return String ���t������
'******************************************************************************
Public Function ToStringOfYYYYMMDDHHNNSSSlash(ByVal dt As Date) As String
    ToStringOfYYYYMMDDHHNNSSSlash = Format(dt, DATE_FORMAT_YYYYMMDDHHNNSS_SLASH)
End Function

'******************************************************************************
'* [�T  �v] ToStringOfYYYYMMDDHHNNSSHyphen
'* [��  ��] Date�^�̓��t����t������iyyyy-mm-dd hh:nn:ss�j�ɕϊ�����B
'*
'* @param dt Date�^���t
'* @return String ���t������
'******************************************************************************
Public Function ToStringOfYYYYMMDDHHNNSSHyphen(ByVal dt As Date) As String
    ToStringOfYYYYMMDDHHNNSSHyphen = Format(dt, DATE_FORMAT_YYYYMMDDHHNNSS_HYPHEN)
End Function

'******************************************************************************
'* [�T  �v] ToStringOfYYYYMMDDHHNNSSDot
'* [��  ��] Date�^�̓��t����t������iyyyy.mm.dd hh:nn:ss�j�ɕϊ�����B
'*
'* @param dt Date�^���t
'* @return String ���t������
'******************************************************************************
Public Function ToStringOfYYYYMMDDHHNNSSDot(ByVal dt As Date) As String
    ToStringOfYYYYMMDDHHNNSSDot = Format(dt, DATE_FORMAT_YYYYMMDDHHNNSS_DOT)
End Function

'******************************************************************************
'* [�T  �v] ToStringOfWareki
'* [��  ��] Date�^�̓��t����t������iggge�Nmm��dd���j�ɕϊ�����B
'*
'* @param dt Date�^���t
'* @return String ���t������
'******************************************************************************
Public Function ToStringOfWareki(ByVal dt As Date) As String
    ToStringOfWareki = Format(dt, DATE_FORMAT_WAREKI)
End Function

'******************************************************************************
'* [�T  �v] ToStringOfWarekiYoubi
'* [��  ��] Date�^�̓��t����t������iggge�Nmm��dd��(aaa)�j�ɕϊ�����B
'*
'* @param dt Date�^���t
'* @return String ���t������
'******************************************************************************
Public Function ToStringOfWarekiYoubi(ByVal dt As Date) As String
    ToStringOfWarekiYoubi = Format(dt, DATE_FORMAT_WAREKI_YOUBI)
End Function

'******************************************************************************
'* [�T  �v] ParseDate
'* [��  ��] ���t�������Date�^�̓��t�ɕϊ�����B
'*
'* @param dateString ���t������
'* @return Date Date�^���t
'******************************************************************************
Public Function ParseDate(ByVal dateString As String) As Date
    On Error GoTo ErrorHandler
    Dim tmpDateStr As String
    tmpDateStr = StringUtils.RemoveEach(dateString, Array(".", "/", "-", " "))
    If Len(tmpDateStr) = 8 Then
        ParseDate = CDate(Format(tmpDateStr, "####/##/##"))
    ElseIf Len(tmpDateStr) = 14 Then
        ParseDate = CDate(Format(tmpDateStr, "####/##/## ##:##:##"))
    Else
        ParseDate = CDate(tmpDateStr)
    End If
    Exit Function
ErrorHandler:
    ParseDate = CDate(dateString)
End Function

'******************************************************************************
'* [�T  �v] AddDays
'* [��  ��] ���t����w�肵�����������Z����B
'*
'* @param dt ���t
'* @param days ����
'* @return Date Date�^���t
'******************************************************************************
Public Function AddDays(ByVal dt As Date, ByVal days As Long) As Date
    AddDays = dt + days
End Function

'******************************************************************************
'* [�T  �v] SubtractDays
'* [��  ��] ���t����w�肵�����������Z����B
'*
'* @param dt ���t
'* @param days ����
'* @return Date Date�^���t
'******************************************************************************
Public Function SubtractDays(ByVal dt As Date, ByVal days As Long) As Date
    SubtractDays = dt - days
End Function

'******************************************************************************
'* [�T  �v] AddYears
'* [��  ��] ���t����w�肵���N�������Z����B
'*
'* @param dt ���t
'* @param years �N��
'* @return Date Date�^���t
'******************************************************************************
Public Function AddYears(ByVal dt As Date, ByVal years As Long) As Date
    AddYears = DateAdd("yyyy", years, dt)
End Function

'******************************************************************************
'* [�T  �v] SubtractYears
'* [��  ��] ���t����w�肵���N�������Z����B
'*
'* @param dt ���t
'* @param years �N��
'* @return Date Date�^���t
'******************************************************************************
Public Function SubtractYears(ByVal dt As Date, ByVal years As Long) As Date
    SubtractYears = AddYears(dt, -years)
End Function

'******************************************************************************
'* [�T  �v] AddMonths
'* [��  ��] ���t����w�肵�����������Z����B
'*
'* @param dt ���t
'* @param months ����
'* @return Date Date�^���t
'******************************************************************************
Public Function AddMonths(ByVal dt As Date, ByVal months As Long) As Date
    AddMonths = DateAdd("m", months, dt)
End Function

'******************************************************************************
'* [�T  �v] SubtractMonths
'* [��  ��] ���t����w�肵�����������Z����B
'*
'* @param dt ���t
'* @param months ����
'* @return Date Date�^���t
'******************************************************************************
Public Function SubtractMonths(ByVal dt As Date, ByVal months As Long) As Date
    SubtractMonths = AddMonths(dt, -months)
End Function

'******************************************************************************
'* [�T  �v] AddWeeks
'* [��  ��] ���t����w�肵���T�������Z����B
'*
'* @param dt ���t
'* @param weeks �T��
'* @return Date Date�^���t
'******************************************************************************
Public Function AddWeeks(ByVal dt As Date, ByVal weeks As Long) As Date
    AddWeeks = DateAdd("ww", weeks, dt)
End Function

'******************************************************************************
'* [�T  �v] SubtractWeeks
'* [��  ��] ���t����w�肵���T�������Z����B
'*
'* @param dt ���t
'* @param weeks �T��
'* @return Date Date�^���t
'******************************************************************************
Public Function SubtractWeeks(ByVal dt As Date, ByVal weeks As Long) As Date
    SubtractWeeks = AddWeeks(dt, -weeks)
End Function

'******************************************************************************
'* [�T  �v] AddHours
'* [��  ��] ���t����w�肵�����Ԑ������Z����B
'*
'* @param dt ���t
'* @param hours ���Ԑ�
'* @return Date Date�^���t
'******************************************************************************
Public Function AddHours(ByVal dt As Date, ByVal hours As Long) As Date
    AddHours = DateAdd("h", hours, dt)
End Function

'******************************************************************************
'* [�T  �v] SubtractHours
'* [��  ��] ���t����w�肵�����Ԑ������Z����B
'*
'* @param dt ���t
'* @param hours ���Ԑ�
'* @return Date Date�^���t
'******************************************************************************
Public Function SubtractHours(ByVal dt As Date, ByVal hours As Long) As Date
    SubtractHours = AddHours(dt, -hours)
End Function

'******************************************************************************
'* [�T  �v] AddMinutes
'* [��  ��] ���t����w�肵�����������Z����B
'*
'* @param dt ���t
'* @param minutes ����
'* @return Date Date�^���t
'******************************************************************************
Public Function AddMinutes(ByVal dt As Date, ByVal minutes As Long) As Date
    AddMinutes = DateAdd("n", minutes, dt)
End Function

'******************************************************************************
'* [�T  �v] SubtractMinutes
'* [��  ��] ���t����w�肵�����������Z����B
'*
'* @param dt ���t
'* @param minutes ����
'* @return Date Date�^���t
'******************************************************************************
Public Function SubtractMinutes(ByVal dt As Date, ByVal minutes As Long) As Date
    SubtractMinutes = AddMinutes(dt, -minutes)
End Function

'******************************************************************************
'* [�T  �v] AddSeconds
'* [��  ��] ���t����w�肵���b�������Z����B
'*
'* @param dt ���t
'* @param seconds �b��
'* @return Date Date�^���t
'******************************************************************************
Public Function AddSeconds(ByVal dt As Date, ByVal seconds As Long) As Date
    AddSeconds = DateAdd("s", seconds, dt)
End Function

'******************************************************************************
'* [�T  �v] SubtractSeconds
'* [��  ��] ���t����w�肵���b�������Z����B
'*
'* @param dt ���t
'* @param seconds �b��
'* @return Date Date�^���t
'******************************************************************************
Public Function SubtractSeconds(ByVal dt As Date, ByVal seconds As Long) As Date
    SubtractSeconds = AddSeconds(dt, -seconds)
End Function

'******************************************************************************
'* [�T  �v] DiffYears
'* [��  ��] 2�̓��t�̍����N�����擾����B
'*
'* @param d1 ���t
'* @param d2 ���t
'* @return Long �����N��
'******************************************************************************
Public Function DiffYears(ByVal d1 As Date, ByVal d2 As Date) As Long
    DiffYears = DateDiff("yyyy", d1, d2)
End Function

'******************************************************************************
'* [�T  �v] DiffMonths
'* [��  ��] 2�̓��t�̍����������擾����B
'*
'* @param d1 ���t
'* @param d2 ���t
'* @return Long ��������
'******************************************************************************
Public Function DiffMonths(ByVal d1 As Date, ByVal d2 As Date) As Long
    DiffMonths = DateDiff("m", d1, d2)
End Function

'******************************************************************************
'* [�T  �v] DiffWeeks
'* [��  ��] 2�̓��t�̍����T�����擾����B
'*
'* @param d1 ���t
'* @param d2 ���t
'* @return Long �����T��
'******************************************************************************
Public Function DiffWeeks(ByVal d1 As Date, ByVal d2 As Date) As Long
    DiffWeeks = DateDiff("ww", d1, d2)
End Function

'******************************************************************************
'* [�T  �v] DiffDays
'* [��  ��] 2�̓��t�̍����������擾����B
'*
'* @param d1 ���t
'* @param d2 ���t
'* @return Long ��������
'******************************************************************************
Public Function DiffDays(ByVal d1 As Date, ByVal d2 As Date) As Long
    DiffDays = DateDiff("d", d1, d2)
End Function

'******************************************************************************
'* [�T  �v] DiffHours
'* [��  ��] 2�̓��t�̍������Ԑ����擾����B
'*
'* @param d1 ���t
'* @param d2 ���t
'* @return Long �������Ԑ�
'******************************************************************************
Public Function DiffHours(ByVal d1 As Date, ByVal d2 As Date) As Long
    DiffHours = DateDiff("h", d1, d2)
End Function

'******************************************************************************
'* [�T  �v] DiffMinutes
'* [��  ��] 2�̓��t�̍����������擾����B
'*
'* @param d1 ���t
'* @param d2 ���t
'* @return Long ��������
'******************************************************************************
Public Function DiffMinutes(ByVal d1 As Date, ByVal d2 As Date) As Long
    DiffMinutes = DateDiff("n", d1, d2)
End Function

'******************************************************************************
'* [�T  �v] DiffSeconds
'* [��  ��] 2�̓��t�̍����b�����擾����B
'*
'* @param d1 ���t
'* @param d2 ���t
'* @return Long �����b��
'******************************************************************************
Public Function DiffSeconds(ByVal d1 As Date, ByVal d2 As Date) As Long
    DiffSeconds = DateDiff("s", d1, d2)
End Function

'******************************************************************************
'* [�T  �v] IsHoliday
'* [��  ��] �����Ŏw�肳�ꂽ���t���j���ł��邩�ǂ������肷��B
'*
'* @param dt ���t
'* @return Boolean ���茋�ʁiTrue�F�j���j
'******************************************************************************
Public Function IsHoliday(ByVal dt As Date) As Boolean
    Dim hol As Variant
    If mHolidayList Is Nothing Then
        Set mHolidayList = GetHolidayList()
    End If
    IsHoliday = CollectionUtils.HasKey(mHolidayList, ToStringOfYYYYMMDDHyphen(dt))
End Function

'******************************************************************************
'* [�T  �v] GetHolidayName
'* [��  ��] �����Ŏw�肳�ꂽ���t����j���̖��̂��擾����B
'*
'* @param dt ���t
'* @return String ���茋�ʁiTrue�F�j���j
'******************************************************************************
Public Function GetHolidayName(ByVal dt As Date) As String
    Dim hol As Variant
    If mHolidayList Is Nothing Then
        Set mHolidayList = GetHolidayList()
    End If

    If CollectionUtils.HasKey(mHolidayList, ToStringOfYYYYMMDDHyphen(dt)) Then
        GetHolidayName = mHolidayList(ToStringOfYYYYMMDDHyphen(dt))
        Exit Function
    End If
    GetHolidayName = ""
End Function

'******************************************************************************
'* [�T  �v] GetHolidayList
'* [��  ��] Web����擾�����j�����X�g��Collection�iKey=�j��(yyyy-mm-dd)�A
'*          Value=�j�����́j�ɐݒ肵�ĕԋp����B
'*
'* @return Collection �j�����X�g
'******************************************************************************
Public Function GetHolidayList() As Collection
    Dim lines: lines = Split(GetHttpResponseText(HOLIDAY_URL), vbLf)
    Dim hList As Collection
    Set hList = New Collection
    Dim i As Long
    For i = 0 To UBound(lines) - 1
        Dim items: items = Split(lines(i), ",")
        hList.Add Replace(items(1), """", ""), items(0)
    Next
    Set GetHolidayList = hList
End Function


'******************************************************************************
'* [�T  �v] GetHttpResponseText
'* [��  ��] �w�肵��URL��HTTP Get���N�G�X�g�𑗂�A���X�|���X�̃e�L�X�g��
'*          �擾����B
'*
'* @param url URL
'* @return String ���X�|���X�e�L�X�g
'******************************************************************************
Private Function GetHttpResponseText(ByVal Url As String) As String
    Dim xmlHttp As Object
    Dim res As String
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
    xmlHttp.Open "GET", Url, False
    xmlHttp.Send
    If xmlHttp.Status = 200 Then
        res = xmlHttp.responseText
    End If
    Set xmlHttp = Nothing
    GetHttpResponseText = res
End Function

'******************************************************************************
'* [�T  �v] GetNendo
'* [��  ��] �w�肵�����t����N�x���擾����B
'*
'* @param sourceDate ���t
'* @return Long �N�x
'******************************************************************************
Public Function GetNendo(ByVal sourceDate As Date) As Long
    If Month(sourceDate) <= 3 Then
        GetNendo = Year(sourceDate) - 1
        Exit Function
    End If
    GetNendo = Year(sourceDate)
End Function