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
'* [機能名] 日付・時間操作ユーティリティクラス
'* [詳  細] 日付・時間操作に関するユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'*
'* [参  考] https://www.take1mg.com/plus1/2018/04/excel-vba-datecalculatorclass/
'*          https://vbabeginner.net/measure-milliseconds-and-microseconds/
'*          https://www.excel-chunchun.com/entry/2019/02/05/014303
'*
'* @author Bankura
'* Copyright (c) 2020-2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
#If Win64 Then
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
#Else
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (frequency As Double) As Long
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (procTime As Double) As Long
#End If
Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long

'******************************************************************************
'* 構造体定義
'******************************************************************************
Private Type LARGE_INTEGER
    quadPart As LongLong
End Type

'******************************************************************************
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 定数定義
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
Private Const DATE_FORMAT_WAREKI As String = "ggge年mm月dd日"
Private Const DATE_FORMAT_WAREKI_YOUBI As String = "ggge年mm月dd日(aaa)"

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mHolidayList As Collection
Private mSWbemDateTime As Object
#If Win64 Then
    Dim mPerfFrequency As LARGE_INTEGER
#Else
    Dim mPerfFrequency As Double
#End If
' Timer用 Start ----------------------
Private mStartNum As Double
Private mLapNum   As Double
Private mPauseNum As Double
Private mPauseFlg As Boolean
Private mLapTimes As Collection
Private mDefaultFormat As String
' Timer用 End   ----------------------

'******************************************************************************
'* プロパティ定義
'******************************************************************************

'*-----------------------------------------------------------------------------
'* Today プロパティ
'*
'*-----------------------------------------------------------------------------
Public Property Get Today() As Date
    Today = Date
End Property

'*-----------------------------------------------------------------------------
'* Yesterday プロパティ
'*
'*-----------------------------------------------------------------------------
Public Property Get Yesterday() As Date
    Yesterday = Date - 1
End Property

'*-----------------------------------------------------------------------------
'* Tomorrow プロパティ
'*
'*-----------------------------------------------------------------------------
Public Property Get Tomorrow() As Date
    Tomorrow = Date + 1
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mSWbemDateTime = CreateObject("WbemScripting.SWbemDateTime")
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] BeginOfMonth
'* [詳  細] 指定した日付の属する月の開始日を取得する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param dt 対象日付
'* @return Date 月初の日付
'******************************************************************************
Public Function BeginOfMonth(ByVal dt As Date) As Date
    BeginOfMonth = Core.BeginOfMonth(dt)
End Function

'******************************************************************************
'* [概  要] EndOfMonth
'* [詳  細] 指定した日付の属する月の最終日を取得する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param dt 対象日付
'* @return Date 月末の日付
'******************************************************************************
Public Function EndOfMonth(ByVal dt As Date) As Date
    EndOfMonth = Core.EndOfMonth(dt)
End Function

'******************************************************************************
'* [概  要] BeginOfWeek
'* [詳  細] 指定した日付の属する週の開始日を取得する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param dt 対象日付
'* @param fstDayOfWeek 週の始まりの曜日
'* @return Date 週初の日付
'******************************************************************************
Public Function BeginOfWeek(ByVal dt As Date, Optional fstDayOfWeek As VbDayOfWeek = vbSunday) As Date
    BeginOfWeek = Core.BeginOfWeek(dt, fstDayOfWeek)
End Function

'******************************************************************************
'* [概  要] EndOfWeek
'* [詳  細] 指定した日付の属する週の最終日を取得する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param dt 対象日付
'* @param fstDayOfWeek 週の始まりの曜日
'* @return Date 週末の日付
'******************************************************************************
Public Function EndOfWeek(ByVal dt As Date, Optional fstDayOfWeek As VbDayOfWeek = vbSunday) As Date
    EndOfWeek = Core.EndOfWeek(dt, fstDayOfWeek)
End Function

'******************************************************************************
'* [概  要] IsLeapYear
'* [詳  細] 閏年かどうかチェックする。
'*
'* @param yearValue 対象年
'* @return Boolean チェック結果（True：閏年）
'******************************************************************************
Public Function IsLeapYear(ByVal yearValue As Integer) As Boolean
    If ((yearValue Mod 4) = 0 And (yearValue Mod 100) <> 0 Or (yearValue Mod 400) = 0) Then
        IsLeapYear = True
        Exit Function
    End If
    IsLeapYear = False
End Function

'******************************************************************************
'* [概  要] ToString
'* [詳  細] Date型の日付を日付文字列に変換する。
'*
'* @param dt Date型日付
'* @param dateStrFormat フォーマット
'* @return String 日付文字列
'******************************************************************************
Public Function ToString(ByVal dt As Date, Optional ByVal dateStrFormat As String = DATE_FORMAT_YYYYMMDDHHNNSS_SLASH) As String
    ToString = Format(dt, dateStrFormat)
End Function

'******************************************************************************
'* [概  要] ToStringOfYYYYMMDD
'* [詳  細] Date型の日付を日付文字列（yyyymmdd）に変換する。
'*
'* @param dt Date型日付
'* @return String 日付文字列
'******************************************************************************
Public Function ToStringOfYYYYMMDD(ByVal dt As Date) As String
    ToStringOfYYYYMMDD = Format(dt, DATE_FORMAT_YYYYMMDD)
End Function

'******************************************************************************
'* [概  要] ToStringOfYYYYMMDDSlash
'* [詳  細] Date型の日付を日付文字列（yyyy/mm/dd）に変換する。
'*
'* @param dt Date型日付
'* @return String 日付文字列
'******************************************************************************
Public Function ToStringOfYYYYMMDDSlash(ByVal dt As Date) As String
    ToStringOfYYYYMMDDSlash = Format(dt, DATE_FORMAT_YYYYMMDD_SLASH)
End Function

'******************************************************************************
'* [概  要] ToStringOfYYYYMMDDHyphen
'* [詳  細] Date型の日付を日付文字列（yyyy-mm-dd）に変換する。
'*
'* @param dt Date型日付
'* @return String 日付文字列
'******************************************************************************
Public Function ToStringOfYYYYMMDDHyphen(ByVal dt As Date) As String
    ToStringOfYYYYMMDDHyphen = Format(dt, DATE_FORMAT_YYYYMMDD_HYPHEN)
End Function

'******************************************************************************
'* [概  要] ToStringOfYYYYMMDDDot
'* [詳  細] Date型の日付を日付文字列（yyyy.mm.dd）に変換する。
'*
'* @param dt Date型日付
'* @return String 日付文字列
'******************************************************************************
Public Function ToStringOfYYYYMMDDDot(ByVal dt As Date) As String
    ToStringOfYYYYMMDDDot = Format(dt, DATE_FORMAT_YYYYMMDD_DOT)
End Function

'******************************************************************************
'* [概  要] ToStringOfYYYYMMDDHHNNSS
'* [詳  細] Date型の日付を日付文字列（yyyymmddhhnnss）に変換する。
'*
'* @param dt Date型日付
'* @return String 日付文字列
'******************************************************************************
Public Function ToStringOfYYYYMMDDHHNNSS(ByVal dt As Date) As String
    ToStringOfYYYYMMDDHHNNSS = Format(dt, DATE_FORMAT_YYYYMMDDHHNNSS)
End Function

'******************************************************************************
'* [概  要] ToStringOfYYYYMMDDHHNNSSSlash
'* [詳  細] Date型の日付を日付文字列（yyyy/mm/dd hh:nn:ss）に変換する。
'*
'* @param dt Date型日付
'* @return String 日付文字列
'******************************************************************************
Public Function ToStringOfYYYYMMDDHHNNSSSlash(ByVal dt As Date) As String
    ToStringOfYYYYMMDDHHNNSSSlash = Format(dt, DATE_FORMAT_YYYYMMDDHHNNSS_SLASH)
End Function

'******************************************************************************
'* [概  要] ToStringOfYYYYMMDDHHNNSSHyphen
'* [詳  細] Date型の日付を日付文字列（yyyy-mm-dd hh:nn:ss）に変換する。
'*
'* @param dt Date型日付
'* @return String 日付文字列
'******************************************************************************
Public Function ToStringOfYYYYMMDDHHNNSSHyphen(ByVal dt As Date) As String
    ToStringOfYYYYMMDDHHNNSSHyphen = Format(dt, DATE_FORMAT_YYYYMMDDHHNNSS_HYPHEN)
End Function

'******************************************************************************
'* [概  要] ToStringOfYYYYMMDDHHNNSSDot
'* [詳  細] Date型の日付を日付文字列（yyyy.mm.dd hh:nn:ss）に変換する。
'*
'* @param dt Date型日付
'* @return String 日付文字列
'******************************************************************************
Public Function ToStringOfYYYYMMDDHHNNSSDot(ByVal dt As Date) As String
    ToStringOfYYYYMMDDHHNNSSDot = Format(dt, DATE_FORMAT_YYYYMMDDHHNNSS_DOT)
End Function

'******************************************************************************
'* [概  要] ToStringOfWareki
'* [詳  細] Date型の日付を日付文字列（ggge年mm月dd日）に変換する。
'*
'* @param dt Date型日付
'* @return String 日付文字列
'******************************************************************************
Public Function ToStringOfWareki(ByVal dt As Date) As String
    ToStringOfWareki = Format(dt, DATE_FORMAT_WAREKI)
End Function

'******************************************************************************
'* [概  要] ToStringOfWarekiYoubi
'* [詳  細] Date型の日付を日付文字列（ggge年mm月dd日(aaa)）に変換する。
'*
'* @param dt Date型日付
'* @return String 日付文字列
'******************************************************************************
Public Function ToStringOfWarekiYoubi(ByVal dt As Date) As String
    ToStringOfWarekiYoubi = Format(dt, DATE_FORMAT_WAREKI_YOUBI)
End Function

'******************************************************************************
'* [概  要] ParseDate
'* [詳  細] 日付文字列をDate型の日付に変換する。
'*
'* @param dateString 日付文字列
'* @return Date Date型日付
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
'* [概  要] AddDays
'* [詳  細] 日付から指定した日数分加算する。
'*
'* @param dt 日付
'* @param days 日数
'* @return Date Date型日付
'******************************************************************************
Public Function AddDays(ByVal dt As Date, ByVal days As Long) As Date
    AddDays = dt + days
End Function

'******************************************************************************
'* [概  要] SubtractDays
'* [詳  細] 日付から指定した日数分減算する。
'*
'* @param dt 日付
'* @param days 日数
'* @return Date Date型日付
'******************************************************************************
Public Function SubtractDays(ByVal dt As Date, ByVal days As Long) As Date
    SubtractDays = dt - days
End Function

'******************************************************************************
'* [概  要] AddYears
'* [詳  細] 日付から指定した年数分加算する。
'*
'* @param dt 日付
'* @param years 年数
'* @return Date Date型日付
'******************************************************************************
Public Function AddYears(ByVal dt As Date, ByVal years As Long) As Date
    AddYears = DateAdd("yyyy", years, dt)
End Function

'******************************************************************************
'* [概  要] SubtractYears
'* [詳  細] 日付から指定した年数分減算する。
'*
'* @param dt 日付
'* @param years 年数
'* @return Date Date型日付
'******************************************************************************
Public Function SubtractYears(ByVal dt As Date, ByVal years As Long) As Date
    SubtractYears = AddYears(dt, -years)
End Function

'******************************************************************************
'* [概  要] AddMonths
'* [詳  細] 日付から指定した月数分加算する。
'*
'* @param dt 日付
'* @param months 月数
'* @return Date Date型日付
'******************************************************************************
Public Function AddMonths(ByVal dt As Date, ByVal months As Long) As Date
    AddMonths = DateAdd("m", months, dt)
End Function

'******************************************************************************
'* [概  要] SubtractMonths
'* [詳  細] 日付から指定した月数分減算する。
'*
'* @param dt 日付
'* @param months 月数
'* @return Date Date型日付
'******************************************************************************
Public Function SubtractMonths(ByVal dt As Date, ByVal months As Long) As Date
    SubtractMonths = AddMonths(dt, -months)
End Function

'******************************************************************************
'* [概  要] AddWeeks
'* [詳  細] 日付から指定した週数分加算する。
'*
'* @param dt 日付
'* @param weeks 週数
'* @return Date Date型日付
'******************************************************************************
Public Function AddWeeks(ByVal dt As Date, ByVal weeks As Long) As Date
    AddWeeks = DateAdd("ww", weeks, dt)
End Function

'******************************************************************************
'* [概  要] SubtractWeeks
'* [詳  細] 日付から指定した週数分減算する。
'*
'* @param dt 日付
'* @param weeks 週数
'* @return Date Date型日付
'******************************************************************************
Public Function SubtractWeeks(ByVal dt As Date, ByVal weeks As Long) As Date
    SubtractWeeks = AddWeeks(dt, -weeks)
End Function

'******************************************************************************
'* [概  要] AddHours
'* [詳  細] 日付から指定した時間数分加算する。
'*
'* @param dt 日付
'* @param hours 時間数
'* @return Date Date型日付
'******************************************************************************
Public Function AddHours(ByVal dt As Date, ByVal hours As Long) As Date
    AddHours = DateAdd("h", hours, dt)
End Function

'******************************************************************************
'* [概  要] SubtractHours
'* [詳  細] 日付から指定した時間数分減算する。
'*
'* @param dt 日付
'* @param hours 時間数
'* @return Date Date型日付
'******************************************************************************
Public Function SubtractHours(ByVal dt As Date, ByVal hours As Long) As Date
    SubtractHours = AddHours(dt, -hours)
End Function

'******************************************************************************
'* [概  要] AddMinutes
'* [詳  細] 日付から指定した分数分加算する。
'*
'* @param dt 日付
'* @param minutes 分数
'* @return Date Date型日付
'******************************************************************************
Public Function AddMinutes(ByVal dt As Date, ByVal minutes As Long) As Date
    AddMinutes = DateAdd("n", minutes, dt)
End Function

'******************************************************************************
'* [概  要] SubtractMinutes
'* [詳  細] 日付から指定した分数分減算する。
'*
'* @param dt 日付
'* @param minutes 分数
'* @return Date Date型日付
'******************************************************************************
Public Function SubtractMinutes(ByVal dt As Date, ByVal minutes As Long) As Date
    SubtractMinutes = AddMinutes(dt, -minutes)
End Function

'******************************************************************************
'* [概  要] AddSeconds
'* [詳  細] 日付から指定した秒数分加算する。
'*
'* @param dt 日付
'* @param seconds 秒数
'* @return Date Date型日付
'******************************************************************************
Public Function AddSeconds(ByVal dt As Date, ByVal seconds As Long) As Date
    AddSeconds = DateAdd("s", seconds, dt)
End Function

'******************************************************************************
'* [概  要] SubtractSeconds
'* [詳  細] 日付から指定した秒数分減算する。
'*
'* @param dt 日付
'* @param seconds 秒数
'* @return Date Date型日付
'******************************************************************************
Public Function SubtractSeconds(ByVal dt As Date, ByVal seconds As Long) As Date
    SubtractSeconds = AddSeconds(dt, -seconds)
End Function

'******************************************************************************
'* [概  要] DiffYears
'* [詳  細] 2つの日付の差分年数を取得する。
'*
'* @param d1 日付
'* @param d2 日付
'* @return Long 差分年数
'******************************************************************************
Public Function DiffYears(ByVal d1 As Date, ByVal d2 As Date) As Long
    DiffYears = DateDiff("yyyy", d1, d2)
End Function

'******************************************************************************
'* [概  要] DiffMonths
'* [詳  細] 2つの日付の差分月数を取得する。
'*
'* @param d1 日付
'* @param d2 日付
'* @return Long 差分月数
'******************************************************************************
Public Function DiffMonths(ByVal d1 As Date, ByVal d2 As Date) As Long
    DiffMonths = DateDiff("m", d1, d2)
End Function

'******************************************************************************
'* [概  要] DiffWeeks
'* [詳  細] 2つの日付の差分週数を取得する。
'*
'* @param d1 日付
'* @param d2 日付
'* @return Long 差分週数
'******************************************************************************
Public Function DiffWeeks(ByVal d1 As Date, ByVal d2 As Date) As Long
    DiffWeeks = DateDiff("ww", d1, d2)
End Function

'******************************************************************************
'* [概  要] DiffDays
'* [詳  細] 2つの日付の差分日数を取得する。
'*
'* @param d1 日付
'* @param d2 日付
'* @return Long 差分日数
'******************************************************************************
Public Function DiffDays(ByVal d1 As Date, ByVal d2 As Date) As Long
    DiffDays = DateDiff("d", d1, d2)
End Function

'******************************************************************************
'* [概  要] DiffHours
'* [詳  細] 2つの日付の差分時間数を取得する。
'*
'* @param d1 日付
'* @param d2 日付
'* @return Long 差分時間数
'******************************************************************************
Public Function DiffHours(ByVal d1 As Date, ByVal d2 As Date) As Long
    DiffHours = DateDiff("h", d1, d2)
End Function

'******************************************************************************
'* [概  要] DiffMinutes
'* [詳  細] 2つの日付の差分分数を取得する。
'*
'* @param d1 日付
'* @param d2 日付
'* @return Long 差分分数
'******************************************************************************
Public Function DiffMinutes(ByVal d1 As Date, ByVal d2 As Date) As Long
    DiffMinutes = DateDiff("n", d1, d2)
End Function

'******************************************************************************
'* [概  要] DiffSeconds
'* [詳  細] 2つの日付の差分秒数を取得する。
'*
'* @param d1 日付
'* @param d2 日付
'* @return Long 差分秒数
'******************************************************************************
Public Function DiffSeconds(ByVal d1 As Date, ByVal d2 As Date) As Long
    DiffSeconds = DateDiff("s", d1, d2)
End Function

'******************************************************************************
'* [概  要] IsHoliday
'* [詳  細] 引数で指定された日付が祝日であるかどうか判定する。
'*
'* @param dt 日付
'* @return Boolean 判定結果（True：祝日）
'******************************************************************************
Public Function IsHoliday(ByVal dt As Date) As Boolean
    Dim hol As Variant
    If mHolidayList Is Nothing Then
        Set mHolidayList = GetHolidayList()
    End If
    IsHoliday = CollectionUtils.HasKey(mHolidayList, ToStringOfYYYYMMDDHyphen(dt))
End Function

'******************************************************************************
'* [概  要] GetHolidayName
'* [詳  細] 引数で指定された日付から祝日の名称を取得する。
'*
'* @param dt 日付
'* @return String 判定結果（True：祝日）
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
'* [概  要] GetHolidayList
'* [詳  細] Webから取得した祝日リストをCollection（Key=祝日(yyyy-mm-dd)、
'*          Value=祝日名称）に設定して返却する。
'*
'* @return Collection 祝日リスト
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
'* [概  要] GetHttpResponseText
'* [詳  細] 指定したURLにHTTP Getリクエストを送り、レスポンスのテキストを
'*          取得する。
'*
'* @param url URL
'* @return String レスポンステキスト
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
'* [概  要] GetNendo
'* [詳  細] 指定した日付から年度を取得する。
'*
'* @param sourceDate 日付
'* @return Long 年度
'******************************************************************************
Public Function GetNendo(ByVal sourceDate As Date) As Long
    If Month(sourceDate) <= 3 Then
        GetNendo = Year(sourceDate) - 1
        Exit Function
    End If
    GetNendo = Year(sourceDate)
End Function

'******************************************************************************
'* [概  要] GetMicroSecond
'* [詳  細] システム起動後の経過時間をマイクロ秒の精度で算出する。
'* [参  考] https://vbabeginner.net/measure-milliseconds-and-microseconds/
'*          http://www.alato.ne.jp/kazu-/vb/tip27.htm
'*
'* @return Double システム起動後の経過時間（マイクロ秒）
'******************************************************************************
Public Function GetMicroSecond() As Double
#If Win64 Then
    Dim perfCount As LARGE_INTEGER
    If mPerfFrequency.quadPart = 0 Then
#Else
    Dim perfCount As Double
    If mPerfFrequency = 0 Then
#End If
        If QueryPerformanceFrequency(mPerfFrequency) = 0 Then
            Debug.Print "QueryPerformanceCounter Is Not Supported."
            Exit Function
        End If
    End If
    Call QueryPerformanceCounter(perfCount)

#If Win64 Then
    GetMicroSecond = perfCount.quadPart / mPerfFrequency.quadPart
#Else
    GetMicroSecond = perfCount / mPerfFrequency
#End If
End Function
'******************************************************************************
'* [概  要] GetMilliSecond
'* [詳  細] システム起動後の経過時間をミリ秒の精度で算出する。
'*
'* @return Double システム起動後の経過時間（ミリ秒）
'******************************************************************************
Public Function GetMilliSecond() As Double
    GetMilliSecond = GetMicroSecond * 1000
End Function

'******************************************************************************
'* [概  要] GetTimestampString
'* [詳  細] 現在時刻の年月日時分秒ミリ秒を「YYYY/MM/DD HH:NN:SS.000」形式の
'*          文字列で取得する。
'*
'* @return String 現在時刻文字列
'******************************************************************************
Public Function GetTimestampString() As String
    Dim dblTimer As Double
    Dim hour As Integer
    Dim minute As Integer
    Dim second As Integer
    Dim mSec As Double
    
    dblTimer = CDbl(Timer)
    hour = dblTimer \ 3600
    minute = (dblTimer Mod 3600) \ 60
    second = dblTimer Mod 60
    mSec = Fix((dblTimer - Fix(dblTimer)) * 1000)

    GetTimestampString = Format(Now, "yyyy/mm/dd") & " " & Format(hour, "00") & ":" & _
                         Format(minute, "00") & ":" & Format(second, "00") & "." & Format(mSec, "000")
End Function

'******************************************************************************
'* [概  要] Get10Kan12Shi
'* [詳  細] 十干十二支を取得する。
'*
'* @param yearValue 対象年
'* @return String 十干十二支
'******************************************************************************
Public Function Get10Kan12Shi(ByVal yearValue As Integer) As String
    Dim list10Kan: list10Kan = Array("甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸")
    Dim list12Shi: list12Shi = Array("子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥")
    
    Get10Kan12Shi = list10Kan((yearValue + 6) Mod 10) & list12Shi((yearValue + 8) Mod 12)
End Function

'******************************************************************************
'* [概  要] GetUTCDate メソッド
'* [詳  細] 日時を設定し、UTC日時の値を取得する。
'*
'* @param targetDate 設定する日時
'* @return UTC日時
'*
'******************************************************************************
Function GetUTCDate(targetDate As Date) As String
    mSWbemDateTime.SetVarDate targetDate, True
    GetUTCDate = mSWbemDateTime.Value
End Function

'******************************************************************************
'* [概  要] UTCtoJST メソッド
'* [詳  細] UTC日時値を標準時間に直して日付型に変換する。
'* [参  考] https://selifelog.com/blog-entry-312.html
'*
'* @param strUTCDate UTC日時値文字列(YYYYMMDDHHMMSS)
'* @return 標準時間
'*
'******************************************************************************
Function UTCtoJST(strUTCDate As String) As Date
    Dim strYear As String:   strYear = Left(strUTCDate, 4)
    Dim strMonth As String:  strMonth = Mid(strUTCDate, 5, 2)
    Dim strDay As String:    strDay = Mid(strUTCDate, 7, 2)
    Dim strHour As String:   strHour = Mid(strUTCDate, 9, 2)
    Dim strMinute As String: strMinute = Mid(strUTCDate, 11, 2)
    Dim strSec As String:    strSec = Mid(strUTCDate, 13, 2)
    
    ' YYYY/MM/DD HH:MM:SSに作り変えて日付型に変換
    Dim JSTDateTime As Date: JSTDateTime = CDate(strYear & "/" & strMonth & "/" & strDay & " " & strHour & ":" & strMinute & ":" & strSec)
    
    ' UTC日時値に9時間加算して日本の標準時に直した値を返却
    UTCtoJST = JSTDateTime + TimeValue("9:00:00")
End Function

'******************************************************************************
'* [概  要] StartTimer
'* [詳  細] タイマーの計測を開始／再開。
'* [参  考] https://www.excel-chunchun.com/entry/2019/02/05/014303
'*
'* @param defFormat デフォルトの時間フォーマット（s, ms, 0ms, 0s etc.）
'*
'******************************************************************************
Public Sub StartTimer(Optional defFormat As String = "")
    If defFormat <> "" Then mDefaultFormat = defFormat

    If mPauseFlg Then
        ' 再開時は、停止していた時間分を開始時間に足すことで辻褄をあわせる
        mStartNum = mStartNum + (GetMilliSecond() - mPauseNum)
        mLapNum = mLapNum + (GetMilliSecond() - mPauseNum)
        mPauseFlg = False
    Else
        mStartNum = GetMilliSecond()
        mLapNum = GetMilliSecond()
    End If
    mPauseNum = 0

    If mLapTimes Is Nothing Then
        Set mLapTimes = New Collection
    End If
End Sub

'******************************************************************************
'* [概  要] ResetTimer
'* [詳  細] タイマーの全ての時間を初期化。
'* [参  考] https://www.excel-chunchun.com/entry/2019/02/05/014303
'*
'******************************************************************************
Public Sub ResetTimer()
    mStartNum = 0
    mLapNum = 0
    mPauseNum = 0
    mPauseFlg = False
    Set mLapTimes = New Collection
End Sub

'******************************************************************************
'* [概  要] PauseTimer
'* [詳  細] タイマーを一時停止し停止時間を記憶。
'* [参  考] https://www.excel-chunchun.com/entry/2019/02/05/014303
'*
'******************************************************************************
Public Sub PauseTimer()
    If mStartNum = 0 Then Exit Sub
    mPauseNum = GetMilliSecond()
    mPauseFlg = True
End Sub

'******************************************************************************
'* [概  要] SplitTime
'* [詳  細] タイマー開始からの経過時間を返却。
'* [参  考] https://www.excel-chunchun.com/entry/2019/02/05/014303
'*
'* @param timeFormat 時間フォーマット（s, ms, #,0 ms, 0.00 s etc.）
'* @param willReset 経過時間を計算後、タイマーのリセットを実行するかどうか
'* @return 開始からの経過時間
'******************************************************************************
Public Function SplitTime(Optional ByVal timeFormat As String = "", _
                          Optional ByVal willReset As Boolean = False) As Variant
    If mStartNum = 0 Then Exit Function
    
    Dim tm As Double
    If mPauseFlg Then
        tm = (mPauseNum - mStartNum)
    Else
        tm = (GetMilliSecond() - mStartNum)
    End If

    If timeFormat = "" Then timeFormat = mDefaultFormat
    If timeFormat = "" Then
        SplitTime = tm
    Else
        SplitTime = GetTimeFormatForTimer(tm, timeFormat)
    End If
    If willReset Then
        Call ResetTimer
    End If
End Function

'******************************************************************************
'* [概  要] LapTime
'* [詳  細] タイマーの直前の計測からの経過時間を返却。
'* [参  考] https://www.excel-chunchun.com/entry/2019/02/05/014303
'*
'* @param timeFormat 時間フォーマット（s, ms, #,0 ms, 0.00 s etc.）
'* @return 直前の計測からの経過時間
'******************************************************************************
Public Function LapTime(Optional ByVal timeFormat As String = "") As Variant
    If mStartNum = 0 Or mPauseFlg Then Exit Function
    Dim tmpTm As Double: tmpTm = GetMilliSecond
    Dim tm As Double: tm = (tmpTm - mLapNum)
    mLapNum = tmpTm

    If timeFormat = "" Then timeFormat = mDefaultFormat
    If timeFormat = "" Then
        LapTime = tm
    Else
        LapTime = GetTimeFormatForTimer(tm, timeFormat)
    End If
    mLapTimes.Add LapTime
End Function

'******************************************************************************
'* [概  要] LapTimes
'* [詳  細] タイマーで記録したラップタイムのコレクションを返却。
'* [参  考] https://www.excel-chunchun.com/entry/2019/02/05/014303
'*
'* @return ラップタイムのコレクション（一次元配列）
'******************************************************************************
Public Function LapTimes() As Variant
    If mStartNum = 0 Or mLapTimes.Count = 0 Then
        LapTimes = Split(vbNullString)
    Else
        Dim Arr() As Variant
        ReDim Arr(1 To mLapTimes.Count)
        Dim i As Long
        For i = 1 To mLapTimes.Count
            Arr(i) = mLapTimes(i)
        Next
        LapTimes = Arr
    End If
End Function

'******************************************************************************
'* [概  要] GetTimeFormatForTimer
'* [詳  細] ミリセコンドから任意の書式に変換（タイマー処理用）。
'* [参  考] https://www.excel-chunchun.com/entry/2019/02/05/014303
'*
'* @param tm 時間（ms）
'* @param timeFormat 時間フォーマット（s, ms, #,0 ms, 0.00 s etc.）
'* @return 指定書式に変換した時間
'******************************************************************************
Private Function GetTimeFormatForTimer(tm As Double, timeFormat As String) As Variant
    Select Case True
        Case timeFormat = "s"
            GetTimeFormatForTimer = tm / 1000
            
        Case timeFormat = "ms"
            GetTimeFormatForTimer = CLng(tm)
            
        Case timeFormat Like "*ms"
            GetTimeFormatForTimer = Right(String(10, " ") & Format(CLng(tm), timeFormat), 10)
        
        Case timeFormat Like "*s"
            GetTimeFormatForTimer = Right(String(10, " ") & Format(tm / 1000, timeFormat), 10)
        
        Case Else
            GetTimeFormatForTimer = Format(CLng(tm), timeFormat)
    End Select
End Function

