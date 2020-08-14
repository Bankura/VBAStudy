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
'* [�@�\��] �������샆�[�e�B���e�B�N���X
'* [��  ��] �����i�P�����j������s�����[�e�B���e�B��񋟂���B
'*          Static�N���X�iAttribute VB_PredeclaredId = True�j�Ƃ���B
'*
'* [�ˁ@��] <�����ˑ�>
'*           MathUtils �N���X
'*
'* [�Q  �l] https://commons.apache.org/proper/commons-lang/apidocs/src-html/org/apache/commons/lang3/CharUtils.html
'*          http://hg.openjdk.java.net/jdk8u/jdk8u60/jdk/file/935758609767/src/share/classes/java/lang/Character.java
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
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
Private Const HEX_DIGITS As String = "0,1,2,3,4,5,6,7,8,9,a,b,c,d,e,f"

'******************************************************************************
'* �ϐ���`
'******************************************************************************
Private mHexsDigits As Variant

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    mHexsDigits = Split(HEX_DIGITS, ",")
End Sub
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************
'******************************************************************************
'* [�T  �v] ���p�p���`�F�b�N�����B
'* [��  ��] �l�����p�p�����ǂ������肷��
'*
'* @param val ���ؒl�i�P�����j
'* @return Boolean �`�F�b�N���ʁiTrue�FOK�AFalse�FNG�j
'*
'******************************************************************************
Public Function IsHalfAlpha(ByVal val As String) As Boolean
    IsHalfAlpha = (val Like "[a-zA-Z]")
End Function

'******************************************************************************
'* [�T  �v] ���p�p�����`�F�b�N�����B
'* [��  ��] �l�����p�p�������ǂ������肷��
'*
'* @param val ���ؒl�i�P�����j
'* @return Boolean �`�F�b�N���ʁiTrue�FOK�AFalse�FNG�j
'*
'******************************************************************************
Public Function IsHalfAlphaNumeric(ByVal val As String) As Boolean
    IsHalfAlphaNumeric = (val Like "[0-9a-zA-Z]")
End Function

'******************************************************************************
'* [�T  �v] ���p�����`�F�b�N�����B
'* [��  ��] �l�����p�������ǂ������肷��
'*
'* @param val ���ؒl�i�P�����j
'* @return Boolean �`�F�b�N���ʁiTrue�FOK�AFalse�FNG�j
'*
'******************************************************************************
Public Function IsHalfNumeric(ByVal val As String) As Boolean
    IsHalfNumeric = (val Like "[0-9]")
End Function

'******************************************************************************
'* [�T  �v] ���p�啶���`�F�b�N�����B
'* [��  ��] �l���啶�����ǂ������肷��
'*
'* @param val ���ؒl�i�P�����j
'* @return Boolean �`�F�b�N���ʁiTrue�FOK�AFalse�FNG�j
'*
'******************************************************************************
Public Function IsHalfUpperCase(ByVal val As String) As Boolean
    IsHalfUpperCase = (val Like "[�`-�y]")
End Function

'******************************************************************************
'* [�T  �v] ���p�������`�F�b�N�����B
'* [��  ��] �l���啶�����ǂ������肷��
'*
'* @param val ���ؒl�i�P�����j
'* @return Boolean �`�F�b�N���ʁiTrue�FOK�AFalse�FNG�j
'*
'******************************************************************************
Public Function IsHalfLowerCase(ByVal val As String) As Boolean
    IsHalfLowerCase = (val Like "[a-z]")
End Function

'******************************************************************************
'* [�T  �v] �啶���`�F�b�N�����B
'* [��  ��] �l���啶�����ǂ������肷��
'*
'* @param val ���ؒl�i�P�����j
'* @return Boolean �`�F�b�N���ʁiTrue�FOK�AFalse�FNG�j
'*
'******************************************************************************
Public Function IsUpperCase(ByVal val As String) As Boolean
    IsUpperCase = (val Like "[A-Z�`-�y]")
End Function

'******************************************************************************
'* [�T  �v] �������`�F�b�N�����B
'* [��  ��] �l���啶�����ǂ������肷��
'*
'* @param val ���ؒl�i�P�����j
'* @return Boolean �`�F�b�N���ʁiTrue�FOK�AFalse�FNG�j
'*
'******************************************************************************
Public Function IsLowerCase(ByVal val As String) As Boolean
    IsLowerCase = (val Like "[a-z��-��]")
End Function

'******************************************************************************
'* [�T  �v] ToIntValue
'* [��  ��] ���l�ɕϊ�����B
'*
'* @param char �����i�P�����j
'* @return Long ���l
'*
'******************************************************************************
Public Function ToIntValue(ByVal char As String, Optional ByVal defaultValue As Long = 0) As Long
    If IsAsciiNumeric(char) Then
        ToIntValue = defaultValue
        Exit Function
    End If
    ToIntValue = AscW(char) - 48
End Function

'******************************************************************************
'* [�T  �v] UnicodeEscaped
'* [��  ��] Converts the string to the Unicode format '\u0020'�B
'*
'* [�p  ��]
'*   CharUtils.UnicodeEscaped(" ") = "\u0020"
'*   CharUtils.UnicodeEscaped("A") = "\u0041"
'*
'* @param char �����i�P�����j
'* @return String Unicode�t�H�[�}�b�g
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
'* [�T  �v] IsAscii
'* [��  ��] Checks whether the character is ASCII 7 bit.
'*
'* [�p  ��]
'*   CharUtils.isAscii("a")  = True
'*   CharUtils.isAscii("A")  = True
'*   CharUtils.isAscii("3")  = True
'*   CharUtils.isAscii("-")  = True
'*   CharUtils.isAscii(vbLf) = True
'*   CharUtils.isAscii('&copy;') = False
'*
'* @param char �����i�P�����j
'* @return Boolean ���茋�ʁiTrue:OK, False:NG�j
'*
'******************************************************************************
Public Function IsAscii(ByVal char As String) As Boolean
    Dim charCode As Integer: charCode = AscW(char)
    IsAscii = (charCode < 128)
End Function

'******************************************************************************
'* [�T  �v] IsAsciiPrintable
'* [��  ��] Checks whether the character is ASCII 7 bit printable.
'*
'* [�p  ��]
'*   CharUtils.IsAsciiPrintable("a")  = True
'*   CharUtils.IsAsciiPrintable("A")  = True
'*   CharUtils.IsAsciiPrintable("3")  = True
'*   CharUtils.IsAsciiPrintable("-")  = True
'*   CharUtils.IsAsciiPrintable(vbLf) = False
'*   CharUtils.IsAsciiPrintable('&copy;') = False
'*
'* @param char �����i�P�����j
'* @return Boolean ���茋�ʁiTrue:OK, False:NG�j
'*
'******************************************************************************
Public Function IsAsciiPrintable(ByVal char As String) As Boolean
    Dim charCode As Integer: charCode = AscW(char)
    IsAsciiPrintable = (charCode >= 32 And charCode < 127)
End Function

'******************************************************************************
'* [�T  �v] IsAsciiControl
'* [��  ��] Checks whether the character is ASCII 7 bit control.
'*
'* [�p  ��]
'*   CharUtils.IsAsciiControl("a")  = False
'*   CharUtils.IsAsciiControl("A")  = False
'*   CharUtils.IsAsciiControl("3")  = False
'*   CharUtils.IsAsciiControl("-")  = False
'*   CharUtils.IsAsciiControl(vbLf) = True
'*   CharUtils.IsAsciiControl('&copy;') = False
'*
'* @param char �����i�P�����j
'* @return Boolean ���茋�ʁiTrue:OK, False:NG�j
'*
'******************************************************************************
Public Function IsAsciiControl(ByVal char As String) As Boolean
    Dim charCode As Integer: charCode = AscW(char)
    IsAsciiControl = (charCode < 32 Or charCode = 127)
End Function

'******************************************************************************
'* [�T  �v] IsAsciiAlpha
'* [��  ��] Checks whether the character is ASCII 7 bit alphabetic.
'*
'* [�p  ��]
'*   CharUtils.IsAsciiAlpha("a")  = True
'*   CharUtils.IsAsciiAlpha("A")  = True
'*   CharUtils.IsAsciiAlpha("3")  = False
'*   CharUtils.IsAsciiAlpha("-")  = False
'*   CharUtils.IsAsciiAlpha(vbLf) = False
'*   CharUtils.IsAsciiAlpha('&copy;') = False
'*
'* @param char �����i�P�����j
'* @return Boolean ���茋�ʁiTrue:OK, False:NG�j
'*
'******************************************************************************
Public Function IsAsciiAlpha(ByVal char As String) As Boolean
    Dim charCode As Integer: charCode = AscW(char)
    IsAsciiAlpha = (IsAsciiAlphaUpper(charCode) Or IsAsciiAlphaLower(charCode))
End Function

'******************************************************************************
'* [�T  �v] IsAsciiAlphaUpper
'* [��  ��] Checks whether the character is ASCII 7 bit alphabetic upper case.
'*
'* [�p  ��]
'*   CharUtils.IsAsciiAlphaUpper("a")  = False
'*   CharUtils.IsAsciiAlphaUpper("A")  = True
'*   CharUtils.IsAsciiAlphaUpper("3")  = False
'*   CharUtils.IsAsciiAlphaUpper("-")  = False
'*   CharUtils.IsAsciiAlphaUpper(vbLf) = False
'*   CharUtils.IsAsciiAlphaUpper('&copy;') = False
'*
'* @param char �����i�P�����j
'* @return Boolean ���茋�ʁiTrue:OK, False:NG�j
'*
'******************************************************************************
Public Function IsAsciiAlphaUpper(ByVal char As String) As Boolean
    Dim charCode As Integer: charCode = AscW(char)
    IsAsciiAlphaUpper = (charCode >= 65 And charCode <= 90)
End Function

'******************************************************************************
'* [�T  �v] IsAsciiAlphaLower
'* [��  ��] Checks whether the character is ASCII 7 bit alphabetic lower case.
'*
'* [�p  ��]
'*   CharUtils.IsAsciiAlphaLower("a")  = True
'*   CharUtils.IsAsciiAlphaLower("A")  = False
'*   CharUtils.IsAsciiAlphaLower("3")  = False
'*   CharUtils.IsAsciiAlphaLower("-")  = False
'*   CharUtils.IsAsciiAlphaLower(vbLf) = False
'*   CharUtils.IsAsciiAlphaLower('&copy;') = False
'*
'* @param char �����i�P�����j
'* @return Boolean ���茋�ʁiTrue:OK, False:NG�j
'*
'******************************************************************************
Public Function IsAsciiAlphaLower(ByVal char As String) As Boolean
    Dim charCode As Integer: charCode = AscW(char)
    IsAsciiAlphaLower = (charCode >= 97 And charCode <= 122)
End Function

'******************************************************************************
'* [�T  �v] IsAsciiNumeric
'* [��  ��] Checks whether the character is ASCII 7 bit alphabetic lower case.
'*
'* [�p  ��]
'*   CharUtils.IsAsciiNumeric("a")  = False
'*   CharUtils.IsAsciiNumeric("A")  = False
'*   CharUtils.IsAsciiNumeric("3")  = True
'*   CharUtils.IsAsciiNumeric("-")  = False
'*   CharUtils.IsAsciiNumeric(vbLf) = False
'*   CharUtils.IsAsciiNumeric('&copy;') = False
'*
'* @param char �����i�P�����j
'* @return Boolean ���茋�ʁiTrue:OK, False:NG�j
'*
'******************************************************************************
Public Function IsAsciiNumeric(ByVal char As String) As Boolean
    Dim charCode As Integer: charCode = AscW(char)
    IsAsciiNumeric = (charCode >= 48 And charCode <= 57)
End Function

'******************************************************************************
'* [�T  �v] IsAsciiAlphanumeric
'* [��  ��] Checks whether the character is ASCII 7 bit alphanumeric.
'*
'* [�p  ��]
'*   CharUtils.IsAsciiNumeric("a")  = True
'*   CharUtils.IsAsciiNumeric("A")  = True
'*   CharUtils.IsAsciiNumeric("3")  = True
'*   CharUtils.IsAsciiNumeric("-")  = False
'*   CharUtils.IsAsciiNumeric(vbLf) = False
'*   CharUtils.IsAsciiNumeric('&copy;') = False
'*
'* @param char �����i�P�����j
'* @return Boolean ���茋�ʁiTrue:OK, False:NG�j
'*
'******************************************************************************
Public Function IsAsciiAlphanumeric(ByVal char As String) As Boolean
    Dim charCode As Integer: charCode = AscW(char)
    IsAsciiAlphanumeric = (IsAsciiAlpha(charCode) Or IsAsciiNumeric(charCode))
End Function

'******************************************************************************
'* [�T  �v] IsWhitespace
'* [��  ��] ���p�X�y�[�X���𔻒�B
'*          ��Java��Character#isWhitespace�Ƃ͎d�l���قȂ�B
'*
'* [�p  ��]
'*   CharUtils.IsWhitespace("")  = True
'*   CharUtils.IsWhitespace(" ")  = True
'*   CharUtils.IsAsciiNumeric("a")  = False
'*
'* @param char �����i�P�����j
'* @return Boolean ���茋�ʁiTrue:OK, False:NG�j
'*
'******************************************************************************
Public Function IsWhitespace(ByVal char As String) As Boolean
    IsWhitespace = (char = " ")
End Function

'******************************************************************************
'* [�T  �v] Compare
'* [��  ��] Compares two char values numerically.
'*
'* @param charX �����i�P�����j
'* @param charX �����i�P�����j
'* @return Long ���茋�ʁi���l�j
'*
'******************************************************************************
Public Function Compare(ByVal charX As String, ByVal charY As String) As Long
    Dim charCodeX As Integer: charCodeX = AscW(charX)
    Dim charCodeY As Integer: charCodeY = AscW(charY)
    Compare = (charCodeX - charCodeY)
End Function