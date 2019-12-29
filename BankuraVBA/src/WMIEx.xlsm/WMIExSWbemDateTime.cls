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
'* [機能名] WbemScripting.SWbemDateTimeクラスのラップ・拡張クラス
'* [詳  細] WbemScripting.SWbemDateTimeクラスのWrapperとしての機能を提供する他、
'*          WbemScripting.SWbemDateTimeクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WbemScripting.SWbemDateTimeクラスについて。
' <https://msdn.microsoft.com/ja-jp/windows/aa393687(v=vs.80)>
' Date & Time
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As WbemScripting.SWbemDateTime
Private mOrigin As Object
    
'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* オリジナルObject プロパティ
'*-----------------------------------------------------------------------------
Property Get Origin() As Object
    Set Origin = mOrigin
End Property
Property Set Origin(arg As Object)
    Set mOrigin = arg
End Property

'*-----------------------------------------------------------------------------
'* Day プロパティ
'*
'* [補  足]
'* ・The Day component of the value (must be in the range 1-31, or 0-999999 for interval values).
'*-----------------------------------------------------------------------------
Property Get Day() As Long
    Day = mOrigin.Day
End Property
Property Let Day(arg As Long)
    mOrigin.Day = arg
End Property

'*-----------------------------------------------------------------------------
'* DaySpecified プロパティ
'*
'* [補  足]
'* ・Whether the Day component is specified.
'*-----------------------------------------------------------------------------
Property Get DaySpecified() As Boolean
    DaySpecified = mOrigin.DaySpecified
End Property
Property Let DaySpecified(arg As Boolean)
    mOrigin.DaySpecified = arg
End Property
    
'*-----------------------------------------------------------------------------
'* Hours プロパティ
'*
'* [補  足]
'* ・The Hours component of the value (must be in the range 0-23).
'*-----------------------------------------------------------------------------
Property Get Hours() As Long
    Hours = mOrigin.Hours
End Property
Property Let Hours(arg As Long)
    mOrigin.Hours = arg
End Property

'*-----------------------------------------------------------------------------
'* HoursSpecified プロパティ
'*
'* [補  足]
'* ・Whether the Hours component is specified.
'*-----------------------------------------------------------------------------
Property Get HoursSpecified() As Boolean
    HoursSpecified = mOrigin.HoursSpecified
End Property
Property Let HoursSpecified(arg As Boolean)
    mOrigin.HoursSpecified = arg
End Property
    
'*-----------------------------------------------------------------------------
'* IsInterval プロパティ
'*
'* [補  足]
'* ・Indicates whether this value describes an absolute date and time or is an interval.
'*-----------------------------------------------------------------------------
Property Get IsInterval() As Boolean
    IsInterval = mOrigin.IsInterval
End Property
Property Let IsInterval(arg As Boolean)
    mOrigin.IsInterval = arg
End Property
    
'*-----------------------------------------------------------------------------
'* Microseconds プロパティ
'*
'* [補  足]
'* ・The Microseconds component of the value (must be in the range 0-999999).
'*-----------------------------------------------------------------------------
Property Get Microseconds() As Long
    Microseconds = mOrigin.Microseconds
End Property
Property Let Microseconds(arg As Long)
    mOrigin.Microseconds = arg
End Property
    
'*-----------------------------------------------------------------------------
'* MicrosecondsSpecified プロパティ
'*
'* [補  足]
'* ・Whether the Microseconds component is specified.
'*-----------------------------------------------------------------------------
Property Get MicrosecondsSpecified() As Boolean
    MicrosecondsSpecified = mOrigin.MicrosecondsSpecified
End Property
Property Let MicrosecondsSpecified(arg As Boolean)
    mOrigin.MicrosecondsSpecified = arg
End Property
    
'*-----------------------------------------------------------------------------
'* Minutes プロパティ
'*
'* [補  足]
'* ・The Minutes component of the value (must be in the range 0-59).
'*-----------------------------------------------------------------------------
Property Get Minutes() As Long
    Minutes = mOrigin.Minutes
End Property
Property Let Minutes(arg As Long)
    mOrigin.Minutes = arg
End Property
    
'*-----------------------------------------------------------------------------
'* MinutesSpecified プロパティ
'*
'* [補  足]
'* ・Whether the Minutes component is specified.
'*-----------------------------------------------------------------------------
Property Get MinutesSpecified() As Boolean
    MinutesSpecified = mOrigin.MinutesSpecified
End Property
Property Let MinutesSpecified(arg As Boolean)
    mOrigin.MinutesSpecified = arg
End Property
    
'*-----------------------------------------------------------------------------
'* Month プロパティ
'*
'* [補  足]
'* ・The Month component of the value (must be in the range 1-12).
'*-----------------------------------------------------------------------------
Property Get Month() As Long
    Month = mOrigin.Month
End Property
Property Let Month(arg As Long)
    mOrigin.Month = arg
End Property
    
'*-----------------------------------------------------------------------------
'* MonthSpecified プロパティ
'*
'* [補  足]
'* ・Whether the Month component is specified.
'*-----------------------------------------------------------------------------
Property Get MonthSpecified() As Boolean
    MonthSpecified = mOrigin.MonthSpecified
End Property
Property Let MonthSpecified(arg As Boolean)
    mOrigin.MonthSpecified = arg
End Property

    
'*-----------------------------------------------------------------------------
'* Seconds プロパティ
'*
'* [補  足]
'* ・The Seconds component of the value (must be in the range 0-59).
'*-----------------------------------------------------------------------------
Property Get Seconds() As Long
    Seconds = mOrigin.Seconds
End Property
Property Let Seconds(arg As Long)
    mOrigin.Seconds = arg
End Property
    
'*-----------------------------------------------------------------------------
'* SecondsSpecified プロパティ
'*
'* [補  足]
'* ・Whether the Seconds component is specified.
'*-----------------------------------------------------------------------------
Property Get SecondsSpecified() As Boolean
    SecondsSpecified = mOrigin.SecondsSpecified
End Property
Property Let SecondsSpecified(arg As Boolean)
    mOrigin.SecondsSpecified = arg
End Property

'*-----------------------------------------------------------------------------
'* UTC プロパティ
'*
'* [補  足]
'* ・The UTC component of the value (must be in the range -720 to 720).
'*-----------------------------------------------------------------------------
Property Get UTC() As Long
    UTC = mOrigin.UTC
End Property
Property Let UTC(arg As Long)
    mOrigin.UTC = arg
End Property
    
'*-----------------------------------------------------------------------------
'* UTCSpecified プロパティ
'*
'* [補  足]
'* ・Whether the UTC component is specified.
'*-----------------------------------------------------------------------------
Property Get UTCSpecified() As Boolean
    UTCSpecified = mOrigin.UTCSpecified
End Property
Property Let UTCSpecified(arg As Boolean)
    mOrigin.UTCSpecified = arg
End Property

'*-----------------------------------------------------------------------------
'* Value プロパティ（既定メンバー）
'*
'* [補  足]
'* ・The DMTF datetime.
'*-----------------------------------------------------------------------------
Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    Value = mOrigin.Value
End Property
Property Let Value(arg As String)
    mOrigin.Value = arg
End Property
   
    
'*-----------------------------------------------------------------------------
'* Year プロパティ
'*
'* [補  足]
'* ・The Year component of the value (must be in the range 0-9999).
'*-----------------------------------------------------------------------------
Property Get Year() As Long
    Year = mOrigin.Year
End Property
Property Let Year(arg As Long)
    mOrigin.Year = arg
End Property
    
'*-----------------------------------------------------------------------------
'* YearSpecified プロパティ
'*
'* [補  足]
'* ・Whether the Year component is specified.
'*-----------------------------------------------------------------------------
Property Get YearSpecified() As Boolean
    YearSpecified = mOrigin.YearSpecified
End Property
Property Let YearSpecified(arg As Boolean)
    mOrigin.YearSpecified = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("WbemScripting.SWbemDateTime")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] GetFileTime メソッド
'* [詳  細] GetFileTime のラッパーメソッド。
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
'* [概  要] GetVarDate メソッド
'* [詳  細] GetVarDate のラッパーメソッド。
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
'* [概  要] SetFileTime メソッド
'* [詳  細] SetFileTime のラッパーメソッド。
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
'* [概  要] SetFileTime メソッド
'* [詳  細] SetFileTime のラッパーメソッド。
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
'* 拡張メソッド
'*-----------------------------------------------------------------------------

'******************************************************************************
'* [概  要] GetUTCDate メソッド
'* [詳  細] 日時を設定し、UTC日時の値を取得する。
'*
'* @param targetDate 設定する日時
'* @return UTC日時
'*
'******************************************************************************
Function GetUTCDate(targetDate As Date) As String
    Me.SetVarDate targetDate, True
    GetUTCDate = Me.Value
End Function

'******************************************************************************
'* [概  要] UTCtoJST メソッド
'* [詳  細] UTC日時値を標準時間に直して日付型に変換する。
'* [引用元] <https://selifelog.com/blog-entry-312.html>
'*
'* @param strUTCDate UTC日時値文字列
'* @return 標準時間
'*
'******************************************************************************
Function UTCtoJST(strUTCDate As String) As Date
    Dim strYear As String, strMonth As String
    Dim strDay As String, strHour As String, strMinute As String
    Dim strSec As String, JSTDateTime As Date
    '引数として受け取った日付時刻文字列(YYYYMMDDHHMMSS)を
    '年、月、日、時、分、秒にバラして各変数に代入
    strYear = Left(strUTCDate, 4)
    strMonth = Mid(strUTCDate, 5, 2)
    strDay = Mid(strUTCDate, 7, 2)
    strHour = Mid(strUTCDate, 9, 2)
    strMinute = Mid(strUTCDate, 11, 2)
    strSec = Mid(strUTCDate, 13, 2)
    '分解したものをYYYY/MM/DD HH:MM:SSに作り変えて日付型に変換する
    JSTDateTime = CDate(strYear & "/" & strMonth & "/" & strDay _
                    & " " & strHour & ":" & strMinute & ":" & strSec)
    'UTC日時値に9時間加算して日本の標準時に直した値を返す
    UTCtoJST = JSTDateTime + TimeValue("9:00:00")
End Function
