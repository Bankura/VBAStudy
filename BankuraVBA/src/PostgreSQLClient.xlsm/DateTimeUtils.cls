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
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************

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
