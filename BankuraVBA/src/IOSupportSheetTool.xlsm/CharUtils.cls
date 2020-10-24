VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CharUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] 文字操作ユーティリティクラス
'* [詳  細] 文字（１文字）操作を行うユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'*
'* [依　存] <内部依存>
'*           MathUtils クラス
'*
'* [参  考] https://commons.apache.org/proper/commons-lang/apidocs/src-html/org/apache/commons/lang3/CharUtils.html
'*          http://hg.openjdk.java.net/jdk8u/jdk8u60/jdk/file/935758609767/src/share/classes/java/lang/Character.java
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
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
Private Const HEX_DIGITS As String = "0,1,2,3,4,5,6,7,8,9,a,b,c,d,e,f"

'******************************************************************************
'* 変数定義
'******************************************************************************
Private mHexsDigits As Variant

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    mHexsDigits = Split(HEX_DIGITS, ",")
End Sub
'******************************************************************************
'* メソッド定義
'******************************************************************************
'******************************************************************************
'* [概  要] 半角英字チェック処理。
'* [詳  細] 値が半角英字かどうか判定する｡
'*
'* @param val 検証値（１文字）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsHalfAlpha(ByVal val As String) As Boolean
    IsHalfAlpha = (val Like "[a-zA-Z]")
End Function

'******************************************************************************
'* [概  要] 半角英数字チェック処理。
'* [詳  細] 値が半角英数字かどうか判定する｡
'*
'* @param val 検証値（１文字）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsHalfAlphaNumeric(ByVal val As String) As Boolean
    IsHalfAlphaNumeric = (val Like "[0-9a-zA-Z]")
End Function

'******************************************************************************
'* [概  要] 半角数字チェック処理。
'* [詳  細] 値が半角数字かどうか判定する｡
'*
'* @param val 検証値（１文字）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsHalfNumeric(ByVal val As String) As Boolean
    IsHalfNumeric = (val Like "[0-9]")
End Function

'******************************************************************************
'* [概  要] 半角大文字チェック処理。
'* [詳  細] 値が大文字かどうか判定する｡
'*
'* @param val 検証値（１文字）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsHalfUpperCase(ByVal val As String) As Boolean
    IsHalfUpperCase = (val Like "[Ａ-Ｚ]")
End Function

'******************************************************************************
'* [概  要] 半角小文字チェック処理。
'* [詳  細] 値が大文字かどうか判定する｡
'*
'* @param val 検証値（１文字）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsHalfLowerCase(ByVal val As String) As Boolean
    IsHalfLowerCase = (val Like "[a-z]")
End Function

'******************************************************************************
'* [概  要] 大文字チェック処理。
'* [詳  細] 値が大文字かどうか判定する｡
'*
'* @param val 検証値（１文字）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsUpperCase(ByVal val As String) As Boolean
    IsUpperCase = (val Like "[A-ZＡ-Ｚ]")
End Function

'******************************************************************************
'* [概  要] 小文字チェック処理。
'* [詳  細] 値が大文字かどうか判定する｡
'*
'* @param val 検証値（１文字）
'* @return Boolean チェック結果（True：OK、False：NG）
'*
'******************************************************************************
Public Function IsLowerCase(ByVal val As String) As Boolean
    IsLowerCase = (val Like "[a-zａ-ｚ]")
End Function

'******************************************************************************
'* [概  要] ToIntValue
'* [詳  細] 数値に変換する。
'*
'* @param char 文字（１文字）
'* @return Long 数値
'*
'******************************************************************************
Public Function ToIntValue(ByVal char As String, Optional ByVal defaultValue As Long = 0) As Long
    If Not IsAsciiNumeric(char) Then
        ToIntValue = defaultValue
        Exit Function
    End If
    ToIntValue = AscW(char) - 48
End Function

'******************************************************************************
'* [概  要] UnicodeEscaped
'* [詳  細] Converts the string to the Unicode format '\u0020'。
'*
'* [用  例]
'*   CharUtils.UnicodeEscaped(" ") = "\u0020"
'*   CharUtils.UnicodeEscaped("A") = "\u0041"
'*
'* @param char 文字（１文字）
'* @return String Unicodeフォーマット
'*
'******************************************************************************
Public Function UnicodeEscaped(ByVal char As String) As String
    Dim charCode As Long: charCode = AscW(char)
    UnicodeEscaped = "\u" & mHexsDigits(MathUtils.RShift(charCode, 12) And 15) & _
                            mHexsDigits(MathUtils.RShift(charCode, 8) And 15) & _
                            mHexsDigits(MathUtils.RShift(charCode, 4) And 15) & _
                            mHexsDigits(charCode And 15)
End Function

'******************************************************************************
'* [概  要] IsAscii
'* [詳  細] Checks whether the character is ASCII 7 bit.
'*
'* [用  例]
'*   CharUtils.isAscii("a")  = True
'*   CharUtils.isAscii("A")  = True
'*   CharUtils.isAscii("3")  = True
'*   CharUtils.isAscii("-")  = True
'*   CharUtils.isAscii(vbLf) = True
'*   CharUtils.isAscii('&copy;') = False
'*
'* @param char 文字（１文字）
'* @return Boolean 判定結果（True:OK, False:NG）
'*
'******************************************************************************
Public Function IsAscii(ByVal char As String) As Boolean
    Dim charCode As Integer: charCode = AscW(char)
    IsAscii = (charCode < 128)
End Function

'******************************************************************************
'* [概  要] IsAsciiPrintable
'* [詳  細] Checks whether the character is ASCII 7 bit printable.
'*
'* [用  例]
'*   CharUtils.IsAsciiPrintable("a")  = True
'*   CharUtils.IsAsciiPrintable("A")  = True
'*   CharUtils.IsAsciiPrintable("3")  = True
'*   CharUtils.IsAsciiPrintable("-")  = True
'*   CharUtils.IsAsciiPrintable(vbLf) = False
'*   CharUtils.IsAsciiPrintable('&copy;') = False
'*
'* @param char 文字（１文字）
'* @return Boolean 判定結果（True:OK, False:NG）
'*
'******************************************************************************
Public Function IsAsciiPrintable(ByVal char As String) As Boolean
    Dim charCode As Integer: charCode = AscW(char)
    IsAsciiPrintable = (charCode >= 32 And charCode < 127)
End Function

'******************************************************************************
'* [概  要] IsAsciiControl
'* [詳  細] Checks whether the character is ASCII 7 bit control.
'*
'* [用  例]
'*   CharUtils.IsAsciiControl("a")  = False
'*   CharUtils.IsAsciiControl("A")  = False
'*   CharUtils.IsAsciiControl("3")  = False
'*   CharUtils.IsAsciiControl("-")  = False
'*   CharUtils.IsAsciiControl(vbLf) = True
'*   CharUtils.IsAsciiControl('&copy;') = False
'*
'* @param char 文字（１文字）
'* @return Boolean 判定結果（True:OK, False:NG）
'*
'******************************************************************************
Public Function IsAsciiControl(ByVal char As String) As Boolean
    Dim charCode As Integer: charCode = AscW(char)
    IsAsciiControl = (charCode < 32 Or charCode = 127)
End Function

'******************************************************************************
'* [概  要] IsAsciiAlpha
'* [詳  細] Checks whether the character is ASCII 7 bit alphabetic.
'*
'* [用  例]
'*   CharUtils.IsAsciiAlpha("a")  = True
'*   CharUtils.IsAsciiAlpha("A")  = True
'*   CharUtils.IsAsciiAlpha("3")  = False
'*   CharUtils.IsAsciiAlpha("-")  = False
'*   CharUtils.IsAsciiAlpha(vbLf) = False
'*   CharUtils.IsAsciiAlpha('&copy;') = False
'*
'* @param char 文字（１文字）
'* @return Boolean 判定結果（True:OK, False:NG）
'*
'******************************************************************************
Public Function IsAsciiAlpha(ByVal char As String) As Boolean
    Dim charCode As Integer: charCode = AscW(char)
    IsAsciiAlpha = (IsAsciiAlphaUpper(charCode) Or IsAsciiAlphaLower(charCode))
End Function

'******************************************************************************
'* [概  要] IsAsciiAlphaUpper
'* [詳  細] Checks whether the character is ASCII 7 bit alphabetic upper case.
'*
'* [用  例]
'*   CharUtils.IsAsciiAlphaUpper("a")  = False
'*   CharUtils.IsAsciiAlphaUpper("A")  = True
'*   CharUtils.IsAsciiAlphaUpper("3")  = False
'*   CharUtils.IsAsciiAlphaUpper("-")  = False
'*   CharUtils.IsAsciiAlphaUpper(vbLf) = False
'*   CharUtils.IsAsciiAlphaUpper('&copy;') = False
'*
'* @param char 文字（１文字）
'* @return Boolean 判定結果（True:OK, False:NG）
'*
'******************************************************************************
Public Function IsAsciiAlphaUpper(ByVal char As String) As Boolean
    Dim charCode As Integer: charCode = AscW(char)
    IsAsciiAlphaUpper = (charCode >= 65 And charCode <= 90)
End Function

'******************************************************************************
'* [概  要] IsAsciiAlphaLower
'* [詳  細] Checks whether the character is ASCII 7 bit alphabetic lower case.
'*
'* [用  例]
'*   CharUtils.IsAsciiAlphaLower("a")  = True
'*   CharUtils.IsAsciiAlphaLower("A")  = False
'*   CharUtils.IsAsciiAlphaLower("3")  = False
'*   CharUtils.IsAsciiAlphaLower("-")  = False
'*   CharUtils.IsAsciiAlphaLower(vbLf) = False
'*   CharUtils.IsAsciiAlphaLower('&copy;') = False
'*
'* @param char 文字（１文字）
'* @return Boolean 判定結果（True:OK, False:NG）
'*
'******************************************************************************
Public Function IsAsciiAlphaLower(ByVal char As String) As Boolean
    Dim charCode As Integer: charCode = AscW(char)
    IsAsciiAlphaLower = (charCode >= 97 And charCode <= 122)
End Function

'******************************************************************************
'* [概  要] IsAsciiNumeric
'* [詳  細] Checks whether the character is ASCII 7 bit alphabetic lower case.
'*
'* [用  例]
'*   CharUtils.IsAsciiNumeric("a")  = False
'*   CharUtils.IsAsciiNumeric("A")  = False
'*   CharUtils.IsAsciiNumeric("3")  = True
'*   CharUtils.IsAsciiNumeric("-")  = False
'*   CharUtils.IsAsciiNumeric(vbLf) = False
'*   CharUtils.IsAsciiNumeric('&copy;') = False
'*
'* @param char 文字（１文字）
'* @return Boolean 判定結果（True:OK, False:NG）
'*
'******************************************************************************
Public Function IsAsciiNumeric(ByVal char As String) As Boolean
    Dim charCode As Integer: charCode = AscW(char)
    IsAsciiNumeric = (charCode >= 48 And charCode <= 57)
End Function

'******************************************************************************
'* [概  要] IsAsciiAlphanumeric
'* [詳  細] Checks whether the character is ASCII 7 bit alphanumeric.
'*
'* [用  例]
'*   CharUtils.IsAsciiNumeric("a")  = True
'*   CharUtils.IsAsciiNumeric("A")  = True
'*   CharUtils.IsAsciiNumeric("3")  = True
'*   CharUtils.IsAsciiNumeric("-")  = False
'*   CharUtils.IsAsciiNumeric(vbLf) = False
'*   CharUtils.IsAsciiNumeric('&copy;') = False
'*
'* @param char 文字（１文字）
'* @return Boolean 判定結果（True:OK, False:NG）
'*
'******************************************************************************
Public Function IsAsciiAlphanumeric(ByVal char As String) As Boolean
    Dim charCode As Integer: charCode = AscW(char)
    IsAsciiAlphanumeric = (IsAsciiAlpha(charCode) Or IsAsciiNumeric(charCode))
End Function

'******************************************************************************
'* [概  要] IsWhitespace
'* [詳  細] 半角スペースかを判定。
'*          ※JavaのCharacter#isWhitespaceとは仕様が異なる。
'*
'* [用  例]
'*   CharUtils.IsWhitespace("")  = True
'*   CharUtils.IsWhitespace(" ")  = True
'*   CharUtils.IsAsciiNumeric("a")  = False
'*
'* @param char 文字（１文字）
'* @return Boolean 判定結果（True:OK, False:NG）
'*
'******************************************************************************
Public Function IsWhitespace(ByVal char As String) As Boolean
    IsWhitespace = (char = " ")
End Function

'******************************************************************************
'* [概  要] Compare
'* [詳  細] Compares two char values numerically.
'*
'* @param charX 文字（１文字）
'* @param charX 文字（１文字）
'* @return Long 判定結果（数値）
'*
'******************************************************************************
Public Function Compare(ByVal charX As String, ByVal charY As String) As Long
    Dim charCodeX As Integer: charCodeX = AscW(charX)
    Dim charCodeY As Integer: charCodeY = AscW(charY)
    Compare = (charCodeX - charCodeY)
End Function
