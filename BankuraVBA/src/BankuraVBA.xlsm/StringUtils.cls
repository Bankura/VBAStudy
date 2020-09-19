VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] 文字列操作ユーティリティクラス
'* [詳  細] 文字列操作を行うユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'*
'* [依　存] <内部依存>
'*           ValidateUtilsクラス
'*           StringExクラス
'*           ArrayUtilsクラス
'*          <外部依存>
'*           Ariawase ライブラリ
'*
'* [参  考] https://docs.microsoft.com/ja-jp/dotnet/api/system.string?view=netframework-4.8#methods
'*          https://commons.apache.org/proper/commons-lang/apidocs/org/apache/commons/lang3/StringUtils.html
'*          https://github.com/RelaxTools/Hidennotare/wiki/StringUtils
'*          https://spring.pleiades.io/spring/docs/5.2.3.RELEASE/javadoc-api/org/springframework/util/StringUtils.html
'*          https://guava.dev/releases/19.0/api/docs/com/google/common/base/Strings.html
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cchMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cchMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As Long) As Long

'******************************************************************************
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 定数定義
'******************************************************************************
Private Const CP_UTF8 As Long = 65001
Private Const MB_PRECOMPOSED As Long = &H1
Private Const INDEX_NOT_FOUND As Long = -1
Private Const PAD_LIMIT As Long = 8192

'******************************************************************************
'* 変数定義
'******************************************************************************

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] FromUTF8
'* [詳  細] UTF8からString型に変換｡
'*
'* @param bytBuf() Byte型配列(UTF8)
'* @return String xxx
'*
'*
'* [Hidennotare] v2.5 StringUtils クラス
'*
'* Copyright (c) 2019 Yasuhiro Watanabe
'* https://github.com/RelaxTools/Hidennotare
'* author:relaxtools@opensquare.net
'*
'******************************************************************************
Public Function FromUTF8(ByRef bytBuf() As Byte) As String
    Dim strBuf As String
    Dim lngSize As Long
    Dim lngByte As Long

    If UBound(bytBuf) < 0 Then
        Exit Function
    End If
    
    lngByte = UBound(bytBuf) - LBound(bytBuf) + 1

    lngSize = MultiByteToWideChar(CP_UTF8, 0, VarPtr(bytBuf(0)), lngByte, 0, 0)
    
    strBuf = String$(lngSize, vbNullChar)
    
    lngSize = MultiByteToWideChar(CP_UTF8, 0, VarPtr(bytBuf(0)), lngByte, StrPtr(strBuf), lngSize)

    FromUTF8 = strBuf
End Function

'******************************************************************************
'* [概  要] ToUTF8
'* [詳  細] String型からByte型配列(UTF8)に変換｡
'*
'* @param strBuf 文字列
'* @return Byte() Byte型配列(UTF8)
'*
'*
'* [Hidennotare] v2.5 StringUtils クラス
'*
'* Copyright (c) 2019 Yasuhiro Watanabe
'* https://github.com/RelaxTools/Hidennotare
'* author:relaxtools@opensquare.net
'*
'******************************************************************************
Public Function ToUTF8(ByVal strBuf As String) As Byte()
    
    If Len(strBuf) = 0 Then
        ToUTF8 = ""
        Exit Function
    End If
    
    Dim lngSize As Long
    Dim bytBuf() As Byte
    
    lngSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strBuf), Len(strBuf), 0, 0, 0, 0)
    
    ReDim bytBuf(0 To lngSize - 1)
    
    WideCharToMultiByte CP_UTF8, 0, StrPtr(strBuf), Len(strBuf), VarPtr(bytBuf(0)), lngSize, 0, 0
    
    ToUTF8 = bytBuf
End Function

'******************************************************************************
'* [概  要] ToHiragana
'* [詳  細] カタカナからひらがなに変換｡
'*
'* @param strBuf カタカナ文字列
'* @return String ひらがな文字列
'*
'*
'* [Hidennotare] v2.5 StringUtils クラス
'*
'* Copyright (c) 2019 Yasuhiro Watanabe
'* https://github.com/RelaxTools/Hidennotare
'* author:relaxtools@opensquare.net
'*
'******************************************************************************
Public Function ToHiragana(ByVal strBuf As String, Optional ByVal flag As Boolean = False) As String

    Dim bytBuf() As Byte
    Dim retBuf() As Byte
    Dim lngBuf As Long
    Dim i As Long
    Dim lngLen As Long
    Dim lngConv As Long
    Dim lngOpt As Long
    
    lngLen = 0
    
    If Len(strBuf) = 0 Then
        ToHiragana = ""
        Exit Function
    End If
    
    bytBuf = strBuf
    retBuf = strBuf
    
    If flag Then
        lngOpt = &H30F6&
    Else
        lngOpt = &H30F4&
    End If
    
    For i = LBound(bytBuf) To UBound(bytBuf) Step 2
    
        lngBuf = MathUtils.LShift(bytBuf(i + 1), 8) + bytBuf(i)
    
        Select Case lngBuf
            'カタカナ
            Case &H30A1& To lngOpt, &H30FD&, &H30FE&
            
                lngConv = lngBuf - &H60&
                retBuf(i) = Me.LByte(lngConv)
                retBuf(i + 1) = Me.UByte(lngConv)
            
        End Select
    
    Next
    
    ToHiragana = retBuf()

End Function

'******************************************************************************
'* [概  要] ToKatakana
'* [詳  細] ひらがなからカタカナに変換｡
'*
'* @param strBuf ひらがな文字列
'* @return String カタカナ文字列
'*
'*
'* [Hidennotare] v2.5 StringUtils クラス
'*
'* Copyright (c) 2019 Yasuhiro Watanabe
'* https://github.com/RelaxTools/Hidennotare
'* author:relaxtools@opensquare.net
'*
'******************************************************************************
Public Function ToKatakana(ByVal strBuf As String, Optional ByVal flag As Boolean = False) As String

    Dim bytBuf() As Byte
    Dim retBuf() As Byte
    Dim lngBuf As Long
    Dim i As Long
    Dim lngLen As Long
    Dim lngConv As Long
    Dim lngOpt As Long
    
    lngLen = 0
    
    If Len(strBuf) = 0 Then
        ToKatakana = ""
        Exit Function
    End If
    
    bytBuf = strBuf
    retBuf = strBuf
    
    If flag Then
        lngOpt = &H3096&
    Else
        lngOpt = &H3094&
    End If
    
    For i = LBound(bytBuf) To UBound(bytBuf) Step 2
    
        lngBuf = MathUtils.LShift(bytBuf(i + 1), 8) + bytBuf(i)
    
        Select Case lngBuf
            'ひらがな
            Case &H3041& To lngOpt, &H309D&, &H309E&
            
                lngConv = lngBuf + &H60&
                retBuf(i) = Me.LByte(lngConv)
                retBuf(i + 1) = Me.UByte(lngConv)
            
        End Select
    
    Next
    
    ToKatakana = retBuf()

End Function

'******************************************************************************
'* [概  要] CheckLength
'* [詳  細] サロゲートペアを考慮した文字数を取得｡
'*
'* @param strBuf 文字列
'* @return Long 文字数
'*
'*
'* [Hidennotare] v2.5 StringUtils クラス LenEx
'*
'* Copyright (c) 2019 Yasuhiro Watanabe
'* https://github.com/RelaxTools/Hidennotare
'* author:relaxtools@opensquare.net
'*
'******************************************************************************
Public Function CheckLength(ByVal strBuf As String) As Long

    Dim bytBuf() As Byte
    Dim lngBuf As Long
    Dim i As Long
    Dim lngLen As Long

    lngLen = 0

    If Len(strBuf) = 0 Then
        CheckLength = 0
        Exit Function
    End If

    bytBuf = strBuf

    For i = LBound(bytBuf) To UBound(bytBuf) Step 2
        lngBuf = MathUtils.LShift(bytBuf(i + 1), 8) + bytBuf(i)

        Select Case lngBuf
            '上位サロゲート
            Case &HD800& To &HDBFF&
                lngLen = lngLen + 1
            '下位サロゲート
            Case &HDC00& To &HDFFF&
                'カウントしない
            Case Else
                lngLen = lngLen + 1
        End Select

    Next
    CheckLength = lngLen
End Function

'******************************************************************************
'* [概  要] TrimZ
'* [詳  細] 全角も含む左右のスペースをトリム｡
'*
'* @param strBuf 文字列
'* @return String トリム文字列
'*
'*
'* [Hidennotare] v2.5 StringUtils クラス TrimEx
'*
'* Copyright (c) 2019 Yasuhiro Watanabe
'* https://github.com/RelaxTools/Hidennotare
'* author:relaxtools@opensquare.net
'*
'******************************************************************************
Public Function TrimZ(ByVal strBuf As String) As String

    Dim lngLen As Long
    Dim lngStart As Long
    Dim lngEnd As Long

    lngLen = Len(strBuf)

    lngStart = 1
    Do Until lngStart > lngLen
        Select Case Mid$(strBuf, lngStart, 1)
            Case Is <= " "
            Case Is = "　"
            Case Else
                Exit Do
        End Select
        lngStart = lngStart + 1
    Loop

    lngEnd = lngLen
    Do Until lngEnd < 1
        Select Case Mid$(strBuf, lngEnd, 1)
            Case Is <= " "
            Case Is = "　"
            Case Else
                Exit Do
        End Select
        lngEnd = lngEnd - 1
    Loop

    If lngEnd > 0 Or lngStart <= lngLen Then
        TrimZ = Mid$(strBuf, lngStart, (lngEnd - lngStart) + 1)
    Else
        TrimZ = ""
    End If

End Function

'******************************************************************************
'* [概  要] StartsWith
'* [詳  細] 指定した文字列で始まるかチェック｡
'*
'* @param str 文字列
'* @param startStr 開始文字列
'* @return Boolean チェック結果（True:該当）
'*
'*
'* [Hidennotare] v2.5 StringUtils クラス
'*
'* Copyright (c) 2019 Yasuhiro Watanabe
'* https://github.com/RelaxTools/Hidennotare
'* author:relaxtools@opensquare.net
'*
'******************************************************************************
Public Function StartsWith(ByVal str As String, ByVal startStr As String) As Boolean
   StartsWith = (startStr = Left$(str, Len(startStr)))
End Function


'******************************************************************************
'* [概  要] EndsWith
'* [詳  細] 指定した文字列で終るかチェック｡
'*
'* @param str 文字列
'* @param endStr 終了文字列
'* @return Boolean チェック結果（True:該当）
'*
'*
'* [Hidennotare] v2.5 StringUtils クラス
'*
'* Copyright (c) 2019 Yasuhiro Watanabe
'* https://github.com/RelaxTools/Hidennotare
'* author:relaxtools@opensquare.net
'*
'******************************************************************************
Public Function EndsWith(ByVal str As String, ByVal endStr As String) As Boolean
    EndsWith = (endStr = Right$(str, Len(endStr)))
End Function

'******************************************************************************
'* [概  要] PlaceHolder
'* [詳  細] 文字列のプレースホルダー{x}に値を埋め込む｡
'*
'* @param str 文字列
'* @param p 埋込文字列の配列
'* @return String 埋込後文字列
'*
'******************************************************************************
Public Function PlaceHolder(ByVal str As String, ParamArray p() As Variant) As String
    Dim strMsg As String
    strMsg = StringUtils.Unescape(str)
    If UBound(p) >= 0 Then
        Dim tmp: tmp = p
        Dim vArr: vArr = ArrayUtils.ToVArrayFlattened(tmp)
        Dim i As Long
        For i = 0 To UBound(vArr)
            strMsg = VBA.Replace(strMsg, "{" & CStr(i) & "}", vArr(i))
        Next
    End If
    PlaceHolder = strMsg
End Function

'******************************************************************************
'* [概  要] Escape
'* [詳  細] 文字列のエスケープを行う｡
'*
'* @param str 文字列
'* @return String エスケープ後文字列
'*
'*
'* [Hidennotare] v2.5 StringUtils クラス
'*
'* Copyright (c) 2019 Yasuhiro Watanabe
'* https://github.com/RelaxTools/Hidennotare
'* author:relaxtools@opensquare.net
'*
'******************************************************************************
Public Function Escape(ByVal v As String) As String

    Dim strChr As String
    Dim sb As StringEx
    Dim i As Long

    Set sb = New StringEx

    For i = 1 To Len(v)

        strChr = Mid$(v, i, 1)

        Select Case AscW(strChr)
            Case &H0 To &H7, &HB, &HE To &H1F, &H7F
                sb.Append "\u" & Hex$(AscW(strChr))
            Case &H8
                sb.Append "\b"
            Case &H9
                sb.Append "\t"
            Case &HA
                sb.Append "\n"
            Case &HC
                sb.Append "\f"
            Case &HD
                sb.Append "\r"
            Case &H22
                sb.Append "\"""
            Case &H2F
                sb.Append "\/"
            Case &H5C
                sb.Append "\\"
            Case Else
                sb.Append strChr
        End Select

    Next

    Escape = sb.ToString
End Function

'******************************************************************************
'* [概  要] Unescape
'* [詳  細] 文字列のアンエスケープを行う｡
'*
'* @param v 文字列
'* @return String アンエスケープ後文字列
'*
'*
'* [Hidennotare] v2.5 StringUtils クラス
'*
'* Copyright (c) 2019 Yasuhiro Watanabe
'* https://github.com/RelaxTools/Hidennotare
'* author:relaxtools@opensquare.net
'*
'******************************************************************************
Public Function Unescape(ByVal v As String) As String

    Dim strChr As String
    Dim strBuf As String
    Dim sb As StringEx
    Dim i As Long

    Set sb = New StringEx
    i = 1
    Do Until i > Len(v)

        strChr = Mid$(v, i, 2)

        Select Case strChr
            Case "\u"

                If i + 6 - 1 > Len(v) Then
                    Err.Raise 512 + 1, TypeName(Me) & ".Unescape", "format error"
                End If

                sb.Append ChrW$(val("&H" & Mid$(v, i + 2, 4)))
                i = i + 6
            Case "\b"
                sb.Append Chr$(&H8)
                i = i + 2
            Case "\t"
                sb.Append vbTab
                i = i + 2
            Case "\n"
                sb.Append vbLf
                i = i + 2
            Case "\f"
                sb.Append Chr$(&HC)
                i = i + 2
            Case "\r"
                sb.Append vbCr
                i = i + 2
            Case "\"""
                sb.Append """"
                i = i + 2
            Case "\/"
                sb.Append "/"
                i = i + 2
            Case "\\"
                sb.Append "\"
                i = i + 2
            Case Else
                sb.Append Left$(strChr, 1)
                i = i + 1
        End Select

    Loop

    Unescape = sb.ToString
End Function


'******************************************************************************
'* [概  要] UByte
'* [詳  細] 上位8ビット文字列を返却｡
'*
'* @param val 文字列
'* @return String 上位8ビット文字列
'*
'*
'* [Hidennotare] v2.5 StringUtils クラス
'*
'* Copyright (c) 2019 Yasuhiro Watanabe
'* https://github.com/RelaxTools/Hidennotare
'* author:relaxtools@opensquare.net
'*
'******************************************************************************
Public Function UByte(ByVal val As Long) As String
    UByte = MathUtils.RShift((val And &HFF00&), 8)
End Function


'******************************************************************************
'* [概  要] LByte
'* [詳  細] 下位8ビット文字列を返却｡
'*
'* @param val 文字列
'* @return String 下位8ビット文字列
'*
'*
'* [Hidennotare] v2.5 StringUtils クラス
'*
'* Copyright (c) 2019 Yasuhiro Watanabe
'* https://github.com/RelaxTools/Hidennotare
'* author:relaxtools@opensquare.net
'*
'******************************************************************************
Public Function LByte(ByVal val As Long) As String
    LByte = val And &HFF&
End Function

'******************************************************************************
'* [概  要] AscLen
'* [詳  細] 文字列のバイト数を求める。
'*          漢字2バイト､半角1バイト｡
'*
'* @param var 文字列
'* @return Long バイト数
'*
'*
'* [Hidennotare] v2.5 StringUtils クラス
'*
'* Copyright (c) 2019 Yasuhiro Watanabe
'* https://github.com/RelaxTools/Hidennotare
'* author:relaxtools@opensquare.net
'*
'******************************************************************************
Public Function AscLen(ByVal var As Variant) As Long
    Dim ascVar As Variant

    ascVar = VBA.StrConv(var, vbFromUnicode)

    AscLen = LenB(ascVar)
End Function

'******************************************************************************
'* [概  要] AscLeft
'* [詳  細] 文字列の左端から指定した文字数分の文字列を返す。
'*          漢字2バイト､半角1バイト｡
'*
'* @param var 文字列
'* @param lngSize 桁数
'* @return String 結果文字列
'*
'*
'* [Hidennotare] v2.5 StringUtils クラス
'*
'* Copyright (c) 2019 Yasuhiro Watanabe
'* https://github.com/RelaxTools/Hidennotare
'* author:relaxtools@opensquare.net
'*
'******************************************************************************
Public Function AscLeft(ByVal var As Variant, ByVal lngSize As Long) As String
    Dim lngLen As Long
    Dim i As Long

    Dim strChr As String
    Dim strResult As String

    lngLen = Len(var)
    strResult = ""

    For i = 1 To lngLen

        strChr = Mid$(var, i, 1)
        If AscLen(strResult & strChr) > lngSize Then
            Exit For
        End If
        strResult = strResult & strChr

    Next
    AscLeft = strResult
End Function

'******************************************************************************
'* [概  要] StrConvU
'* [詳  細] 文字化け対応StrConv(vbUnicode, vbFromUnicodeはサポート外)。
'*
'* @param strSource 文字列
'* @param Conv VbStrConv
'* @return String 結果文字列
'*
'*
'* [Hidennotare] v2.5 StringUtils クラス
'*
'* Copyright (c) 2019 Yasuhiro Watanabe
'* https://github.com/RelaxTools/Hidennotare
'* author:relaxtools@opensquare.net
'*
'******************************************************************************
Public Function StrConvU(ByVal strSource As String, Conv As VbStrConv) As String

    Dim i As Long
    Dim strBuf As String
    Dim c As String
    Dim strRet As String
    Dim strBefore As String
    Dim strChr As String
    Dim strNext As String

    strRet = ""
    strBuf = ""
    strBefore = ""
    strNext = ""

    For i = 1 To Len(strSource)

        c = Mid$(strSource, i, 1)
        
        If i = Len(strSource) Then
            strNext = ""
        Else
            strNext = Mid$(strSource, i + 1, 1)
        End If

        Select Case c
            '全角の￥
            Case "￥"
                If (Conv And vbNarrow) > 0 Then
                    strChr = "\"
                    strRet = strRet & VBA.StrConv(strBuf, Conv) & strChr
                    strBuf = ""
                Else
                    strBuf = strBuf & c
                End If
           
            '半角の\
            Case "\"
                If (Conv And vbWide) > 0 Then
                    strChr = "￥"
                    strRet = strRet & VBA.StrConv(strBuf, Conv) & strChr
                    strBuf = ""
                Else
                    strBuf = strBuf & c
                End If
            '全角の濁点、半濁点
            Case "゜", "゛"
                If (Conv And vbNarrow) > 0 Then
                    If c = "゜" Then
                        strChr = "ﾟ"
                    Else
                        strChr = "ﾞ"
                    End If
                    strRet = strRet & VBA.StrConv(strBuf, Conv) & strChr
                    strBuf = ""
                Else
                    strBuf = strBuf & c
                End If
                
            '半角の半濁点
            Case "ﾟ"
                '１つ前の文字
                Select Case strBefore
                    Case "ﾊ" To "ﾎ"
                        strBuf = strBuf & c
                    Case Else
                        If (Conv And vbWide) > 0 Then
                             strChr = "゜"
                            strRet = strRet & VBA.StrConv(strBuf, Conv) & strChr
                            strBuf = ""
                        Else
                            strBuf = strBuf & c
                        End If
                End Select
                
            '半角の濁点
            Case "ﾞ"
                '１つ前の文字
                Select Case strBefore
                    Case "ｳ", "ｶ" To "ｺ", "ｻ" To "ｿ", "ﾀ" To "ﾄ", "ﾊ" To "ﾎ"
                        strBuf = strBuf & c
                    Case Else
                        If (Conv And vbWide) > 0 Then
                            strChr = "゛"
                            strRet = strRet & VBA.StrConv(strBuf, Conv) & strChr
                            strBuf = ""
                        Else
                            strBuf = strBuf & c
                        End If
                End Select
            'ヴ
            Case "ヴ"
                If (Conv And vbHiragana) > 0 Then
                    Dim b() As Byte
                    ReDim b(0 To 1)
                    b(0) = &H94
                    b(1) = &H30
                    strChr = b
                    strRet = strRet & VBA.StrConv(strBuf, Conv) & strChr
                    strBuf = ""
                Else
                    strBuf = strBuf & c
                End If
            'う゛
            Case "う"
                If strNext = "゛" And (Conv And vbKatakana) > 0 Then
                    strChr = "ヴ"
                    strRet = strRet & VBA.StrConv(strBuf, Conv) & strChr
                    strBuf = ""
                    i = i + 1
                Else
                    strBuf = strBuf & c
                End If

            'ヶヵ
            Case "ヶ", "ヵ"
                If (Conv And vbHiragana) > 0 Then
                    strRet = strRet & VBA.StrConv(strBuf, Conv) & c
                    strBuf = ""
                Else
                    strBuf = strBuf & c
                End If

            'その他
            Case Else
                '第二水準等StrConvで文字化けするものを退避
                If Asc(c) = 63 And c <> "?" Then
                    strRet = strRet & VBA.StrConv(strBuf, Conv) & c
                    strBuf = ""
                Else
                    'う”
                    If AscW(c) = &H3094 Then
                        If Conv = vbKatakana Then
                            strRet = strRet & VBA.StrConv(strBuf, Conv) & "ヴ"
                            strBuf = ""
                        Else
                            strRet = strRet & VBA.StrConv(strBuf, Conv) & c
                            strBuf = ""
                        End If
                    Else
                        strBuf = strBuf & c
                    End If
                End If
        End Select
        
        '１個前の文字
        strBefore = c

    Next

    If strBuf <> "" Then
        strRet = strRet & VBA.StrConv(strBuf, Conv)
    End If

    StrConvU = strRet

End Function


'******************************************************************************
'* [概  要] AppendIfMissing
'* [詳  細] 末尾に接尾辞が存在しない場合は、Stringの末尾に接尾辞を追加｡
'*
'* [用  例]
'*   StringUtils.AppendIfMissing("", "xyz") = "xyz"
'*   StringUtils.AppendIfMissing("abc", "xyz") = "abcxyz"
'*   StringUtils.AppendIfMissing("abcxyz", "xyz") = "abcxyz"
'*   StringUtils.AppendIfMissing("abcXYZ", "xyz") = "abcXYZxyz"
'*   StringUtils.AppendIfMissing("abc", "xyz", "") = "abc"
'*   StringUtils.AppendIfMissing("abc", "xyz", "mno") = "abcxyz"
'*   StringUtils.AppendIfMissing("abcxyz", "xyz", "mno") = "abcxyz"
'*   StringUtils.AppendIfMissing("abcmno", "xyz", "mno") = "abcmno"
'*   StringUtils.AppendIfMissing("abcXYZ", "xyz", "mno") = "abcXYZxyz"
'*   StringUtils.AppendIfMissing("abcMNO", "xyz", "mno") = "abcMNOxyz"
'*
'* @param str 文字列
'* @param suffix 文字列の最後に追加する接尾辞
'* @param suffixes 有効なターミネーターである追加のサフィックス接尾辞
'* @return String 接尾辞が追加された場合は新しい文字列、それ以外の場合は同じ文字列
'*
'******************************************************************************
Public Function AppendIfMissing(ByVal str As String, ByVal suffix As String, ParamArray suffixes()) As String
    If Me.IsEmpty(suffix) Or Me.EndsWith(str, suffix) Then
        AppendIfMissing = str
        Exit Function
    End If
    If Not IsMissing(suffixes) Then
        Dim tmp: tmp = suffixes ' ParamArrayを別関数に渡すための対応
        Dim v
        For Each v In ArrayUtils.ToVArrayFlattened(tmp)
            If Me.EndsWith(str, CStr(v)) Then
                AppendIfMissing = str
                Exit Function
            End If
        Next
    End If
    AppendIfMissing = str & suffix
End Function

'******************************************************************************
'* [概  要] AppendIfMissingIgnoreCase
'* [詳  細] 末尾に接尾辞が存在しない場合は、Stringの末尾に接尾辞を追加｡
'*          接尾辞の存在チェックは大文字小文字を考慮しない。
'*
'* [用  例]
'*   StringUtils.AppendIfMissing("", "xyz") = "xyz"
'*   StringUtils.AppendIfMissingIgnoreCase("", "xyz") = "xyz"
'*   StringUtils.AppendIfMissingIgnoreCase("abc", "xyz") = "abcxyz"
'*   StringUtils.AppendIfMissingIgnoreCase("abcxyz", "xyz") = "abcxyz"
'*   StringUtils.AppendIfMissingIgnoreCase("abcXYZ", "xyz") = "abcXYZ"
'*   StringUtils.AppendIfMissingIgnoreCase("abc", "xyz", "") = "abc"
'*   StringUtils.AppendIfMissingIgnoreCase("abc", "xyz", "mno") = "abcxyz"
'*   StringUtils.AppendIfMissingIgnoreCase("abcxyz", "xyz", "mno") = "abcxyz"
'*   StringUtils.AppendIfMissingIgnoreCase("abcmno", "xyz", "mno") = "abcmno"
'*   StringUtils.AppendIfMissingIgnoreCase("abcXYZ", "xyz", "mno") = "abcXYZ"
'*   StringUtils.AppendIfMissingIgnoreCase("abcMNO", "xyz", "mno") = "abcMNO"
'*
'* @param str 文字列
'* @param suffix 文字列の最後に追加する接尾辞
'* @param suffixes 有効なターミネーターである追加のサフィックス接尾辞
'* @return String 接尾辞が追加された場合は新しい文字列、それ以外の場合は同じ文字列
'*
'******************************************************************************
Public Function AppendIfMissingIgnoreCase(ByVal str As String, ByVal suffix As String, ParamArray suffixes()) As String
    If Me.IsEmpty(suffix) Or Me.EndsWithIgnoreCase(str, suffix) Then
        AppendIfMissingIgnoreCase = str
        Exit Function
    End If
    If Not IsMissing(suffixes) Then
        Dim tmp: tmp = suffixes ' ParamArrayを別関数に渡すための対応
        Dim v
        For Each v In ArrayUtils.ToVArrayFlattened(tmp)
            If Me.EndsWithIgnoreCase(str, CStr(v)) Then
                AppendIfMissingIgnoreCase = str
                Exit Function
            End If
        Next
    End If
    AppendIfMissingIgnoreCase = str & suffix
End Function

'******************************************************************************
'* [概  要] Capitalize
'* [詳  細] 指定された文字列の最初の文字を大文字に変換する｡
'*
'* [用  例]
'*   StringUtils.Capitalize("") = ""
'*   StringUtils.Capitalize("cat") = "Cat"
'*   StringUtils.Capitalize("cAt") = "CAt"
'*   StringUtils.Capitalize("'cat'") = "'cat'"
'*
'* @param str 文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function Capitalize(ByVal str As String) As String
    If Me.IsEmpty(str) Then
       Capitalize = str
       Exit Function
    End If
    Capitalize = StrConv(Left(str, 1), vbUpperCase) & Mid(str, 2)
End Function

'******************************************************************************
'* [概  要] Center
'* [詳  細] 文字列をパディングして中央寄せにする｡
'*          指定されたパディング文字があれば、文字列に埋め込む値として使用する。
'*          デフォルトのパディング文字は半角スペース。
'*          サイズが文字列の長さより小さい場合、元の文字列が返される。
'*          空文字列は空文字列を返す。負のサイズはゼロとして扱われる。
'*
'* [用  例]
'*   StringUtils.Center("", 4) = "    "
'*   StringUtils.Center("ab", -1) = "ab"
'*   StringUtils.Center("ab", 4) = " ab "
'*   StringUtils.Center("abcd", 2) = "abcd"
'*   StringUtils.Center("a", 4) = " a  "
'*   StringUtils.Center("a", 4, "y") = "yayy"
'*   StringUtils.Center("a", 4, "yz") = "yayz"
'*   StringUtils.Center("abc", 7, "") = "  abc  "
'*
'* @param str  文字列
'* @param size 新しい文字列のサイズ
'* @param padStr パディング文字
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function Center(ByVal str As String, ByVal size As Long, Optional padStr As String = " ") As String
    If Me.IsEmpty(str) Or size <= 0 Then
       Center = str
       Exit Function
    End If
    Dim strLen As Long: strLen = Me.CheckLength(str)
    Dim pads As Long: pads = size - strLen
    If pads <= 0 Then
       Center = str
       Exit Function
    End If
    str = Me.LeftPad(str, strLen + pads / 2, padStr)
    str = Me.RightPad(str, size, padStr)
    Center = str
End Function

'******************************************************************************
'* [概  要] Chomp
'* [詳  細] 文字列の末尾にある場合は改行を1つ削除する。
'*          それ以外の場合はそのままにする｡
'*          改行は「\n」、「\r」、または「\r\n」｡
'*
'* [用  例]
'*   StringUtils.Chomp("") = ""
'*   StringUtils.Chomp("abc " & vbCr) = "abc "
'*   StringUtils.Chomp("abc" & vbLf) = "abc"
'*   StringUtils.Chomp("abc" & vbCrLf) = "abc"
'*   StringUtils.Chomp("abc" & vbCrLf & vbCrLf) = "abc\r\n"
'*   StringUtils.Chomp("abc" & vbLf & vbCr) = "abc\n"
'*   StringUtils.Chomp("abc" & vbLf & vbCr & "abc") = "abc\n\rabc"
'*   StringUtils.Chomp(vbCr) = ""
'*   StringUtils.Chomp(vbLf) = ""
'*   StringUtils.Chomp(vbCrLf) = ""
'*
'* @param str 文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function Chomp(ByVal str As String) As String
    If Me.IsEmpty(str) Then
        Chomp = str
        Exit Function
    End If
    If Len(str) = 1 Then
        If str = vbCr Or str = vbLf Then
            Chomp = ""
            Exit Function
        End If
    End If
    Dim lastIdx As Long: lastIdx = Len(str) - 1
    Dim lastChar As String: lastChar = Me.CharAt(str, lastIdx)
    If lastChar = vbLf Then
        If Me.CharAt(str, lastIdx - 1) = vbCr Then
            lastIdx = lastIdx - 1
        End If
    ElseIf lastChar <> vbCr Then
        lastIdx = lastIdx + 1
    End If
    Chomp = Me.Substring(str, 0, lastIdx)
End Function

'******************************************************************************
'* [概  要] Chop
'* [詳  細] 文字列から最後の文字を削除｡
'*          文字列が\r\nで終わっている場合は、両方を削除する。
'*
'* [用  例]
'*   StringUtils.Chop("") = ""
'*   StringUtils.Chop("abc \r") = "abc "
'*   StringUtils.Chop("abc\n") = "abc"
'*   StringUtils.Chop("abc\r\n") = "abc"
'*   StringUtils.Chop("abc") = "ab"
'*   StringUtils.Chop("abc\nabc") = "abc\nab"
'*   StringUtils.Chop("a") = ""
'*   StringUtils.Chop("\r") = ""
'*   StringUtils.Chop("\n") = ""
'*   StringUtils.Chop("\r\n") = ""
'*
'* @param str 文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function Chop(ByVal str As String) As String
    If Me.IsEmpty(str) Then
        Chop = ""
        Exit Function
    End If
    
    Dim strLen As Long: strLen = Len(str)
    If strLen < 2 Then
        Chop = ""
        Exit Function
    End If
    
    Dim lastIdx As Long: lastIdx = strLen - 1
    Dim ret As String: ret = Me.Substring(str, 0, lastIdx)
    Dim lastChar As String: lastChar = Me.CharAt(str, lastIdx)
    
    If lastChar = vbLf And Me.CharAt(ret, lastIdx - 1) = vbCr Then
        Chop = Me.Substring(ret, 0, lastIdx - 1)
        Exit Function
    End If
    Chop = ret
End Function

'******************************************************************************
'* [概  要] Compare
'* [詳  細] 2つの文字列を比較する｡
'*          ※JavaのString#compareToのような辞書的な比較結果（数値）は返さない。
'*
'* [用  例]
'*   StringUtils.Compare("abc", "abc") = 0
'*   StringUtils.Compare("a", "b")     < 0
'*   StringUtils.Compare("b", "a")     > 0
'*   StringUtils.Compare("a", "B")     > 0
'*   StringUtils.Compare("ab", "abc")  < 0
'*
'* @param str1 比較する文字列1
'* @param str2 比較する文字列2
'* @return Long チェック結果
'*         int = 0     ：str1 = str2
'*         int < 0(-1) ：str1 < str2
'*         int > 0( 1) ：str1 > str2
'*
'******************************************************************************
Public Function Compare(ByVal str1 As String, ByVal str2 As String) As Long
    If str1 = str2 Then
        Compare = 0
        Exit Function
    ElseIf str1 < str2 Then
        Compare = -1
        Exit Function
    Else
        Compare = 1
    End If
End Function

'******************************************************************************
'* [概  要] CompareIgnoreCase
'* [詳  細] 2つの文字列を比較する｡
'*          大文字小文字を考慮しない。
'*          ※JavaのString#compareToのような辞書的な比較結果（数値）は返さない。
'*
'* [用  例]
'*   StringUtils.CompareIgnoreCase("abc", "abc") = 0
'*   StringUtils.CompareIgnoreCase("abc", "ABC") = 0
'*   StringUtils.CompareIgnoreCase("a", "b")     < 0
'*   StringUtils.CompareIgnoreCase("b", "a")     > 0
'*   StringUtils.CompareIgnoreCase("a", "B")     < 0
'*   StringUtils.CompareIgnoreCase("A", "b")     < 0
'*   StringUtils.CompareIgnoreCase("ab", "ABC")  < 0
'*
'* @param str1 比較する文字列1
'* @param str2 比較する文字列2
'* @return Long チェック結果
'*         int = 0     ：str1 = str2
'*         int < 0(-1) ：str1 < str2
'*         int > 0( 1) ：str1 > str2
'*
'******************************************************************************
Public Function CompareIgnoreCase(ByVal str1 As String, ByVal str2 As String) As Long
    Dim str1Ic As String: str1Ic = Me.LowerCase(str1)
    Dim str2Ic As String: str2Ic = Me.LowerCase(str2)
    CompareIgnoreCase = Me.Compare(str1Ic, str2Ic)
End Function

'******************************************************************************
'* [概  要] Contains
'* [詳  細] 指定した文字列が含まれているかどうかをチェック｡
'*          空文字の文字列はFalseを返す｡
'*          空文字の検索文字はFalseを返す｡
'*
'* [用  例]
'*   StringUtils.Contains("", *)      = false
'*   StringUtils.Contains("abc", 'a') = true
'*   StringUtils.Contains("abc", 'z') = false
'*
'* @param str 文字列
'* @param searchStr 検索文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function Contains(ByVal str As String, ByVal searchStr As String) As Boolean
    If Me.IsEmpty(str) Or Me.IsEmpty(searchStr) Then
        Contains = False
        Exit Function
    End If
    Contains = (InStr(str, searchStr) <> 0)
End Function

'******************************************************************************
'* [概  要] ContainsAll
'* [詳  細] 指定した文字列のすべてが含まれているかチェック｡
'*          空文字の文字列はFalseを返す｡
'*          空文字または長さがゼロの検索配列はFalseを返す｡
'*
'* @param str 文字列
'* @param searchStr 検索文字配列（複数指定）
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function ContainsAll(ByVal str As String, ParamArray searchStr()) As Boolean
    If Me.IsEmpty(str) Or VBA.IsMissing(searchStr) Then
        ContainsAll = False
        Exit Function
    End If
    Dim tmp: tmp = searchStr ' ParamArrayを別関数に渡すための対応
    Dim v
    For Each v In ArrayUtils.ToVArrayFlattened(tmp)
        If Not Me.Contains(str, CStr(v)) Then
            ContainsAll = False
            Exit Function
        End If
    Next
    ContainsAll = True
End Function

'******************************************************************************
'* [概  要] ContainsAny
'* [詳  細] 指定した文字列のどれかが含まれているかどうかをチェック｡
'*          空文字の文字列はFalseを返す｡
'*          空文字または長さがゼロの検索配列はFalseを返す｡
'*
'* [用  例]
'*   StringUtils.ContainsAny("", *)                    = false
'*   StringUtils.ContainsAny(*, [])                    = false
'*   StringUtils.containsAny("abcd", "ab", "") = true
'*   StringUtils.containsAny("abcd", "ab", "cd") = true
'*   StringUtils.containsAny("abc", "d", "abc")  = true
'*
'* @param str 文字列
'* @param searchStr 検索文字配列（複数指定）
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function ContainsAny(ByVal str As String, ParamArray searchStr()) As Boolean
    If Me.IsEmpty(str) Or VBA.IsMissing(searchStr) Then
        ContainsAny = False
        Exit Function
    End If
    Dim tmp: tmp = searchStr ' ParamArrayを別関数に渡すための対応
    Dim v
    For Each v In ArrayUtils.ToVArrayFlattened(tmp)
        If Me.Contains(str, CStr(v)) Then
            ContainsAny = True
            Exit Function
        End If
    Next
    ContainsAny = False
End Function

'******************************************************************************
'* [概  要] ContainsIgnoreCase
'* [詳  細] 指定した文字列が含まれているかどうかをチェック｡
'*          大文字小文字を考慮しない。
'*          空文字の文字列はFalseを返す｡
'*          空文字の検索文字はFalseを返す｡
'*
'* [用  例]
'*   StringUtils.ContainsIgnoreCase("", "") = True
'*   StringUtils.ContainsIgnoreCase("abc", "") = True
'*   StringUtils.ContainsIgnoreCase("abc", "a") = True
'*   StringUtils.ContainsIgnoreCase("abc", "z") = False
'*   StringUtils.ContainsIgnoreCase("abc", "A") = True
'*   StringUtils.ContainsIgnoreCase("abc", "Z") = False
'*
'* @param str 文字列
'* @param searchStr 検索文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function ContainsIgnoreCase(ByVal str As String, ByVal searchStr As String) As Boolean
    If Me.IsEmpty(str) Or Me.IsEmpty(searchStr) Then
        ContainsIgnoreCase = False
        Exit Function
    End If
    Dim strIc As String: strIc = Me.LowerCase(str)
    Dim searchStrIc As String: searchStrIc = Me.LowerCase(searchStr)
    ContainsIgnoreCase = Me.Contains(strIc, searchStrIc)
End Function

'******************************************************************************
'* [概  要] ContainsNone
'* [詳  細] 指定した文字列が含まれていないことをチェック｡
'*          空文字の文字列はTrueを返す｡
'*          空文字の検索文字はTrueを返す｡
'*
'* [用  例]
'*   StringUtils.ContainsNone("", *)         = true
'*   StringUtils.ContainsNone("ab", "")      = true
'*   StringUtils.ContainsNone("abab", "xyz") = true
'*   StringUtils.ContainsNone("ab1", "xyz")  = true
'*   StringUtils.ContainsNone("abz", "xyz")  = false
'*
'* @param str 文字列
'* @param searchStr 検索文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function ContainsNone(ByVal str As String, ByVal searchStr As String) As Boolean
    If Me.IsEmpty(str) Or Me.IsEmpty(searchStr) Then
        ContainsNone = False
        Exit Function
    End If
    ContainsNone = Not ContainsAny(str, Me.ToCharArray(searchStr))
End Function

'******************************************************************************
'* [概  要] ContainsOnly
'* [詳  細] 特定の文字のみが含まれているかどうかをチェック｡
'*          空の文字列（length（）= 0）は常にTrueを返す。
'*
'* [用  例]
'*   StringUtils.ContainsOnly("", *)         = True
'*   StringUtils.ContainsOnly("ab", "")      = False
'*   StringUtils.ContainsOnly("abab", "abc") = True
'*   StringUtils.ContainsOnly("ab1", "abc")  = False
'*   StringUtils.ContainsOnly("abz", "abc")  = False
'*
'* @param str 文字列
'* @param searchStr 検索文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function ContainsOnly(ByVal str As String, ByVal searchStr As String) As Boolean
    If Me.IsEmpty(str) Then
        ContainsOnly = True
        Exit Function
    End If
    If Me.IsEmpty(searchStr) Then
        ContainsOnly = False
        Exit Function
    End If
    ContainsOnly = (Me.IndexOfAnyBut(str, searchStr) = INDEX_NOT_FOUND)
End Function

'******************************************************************************
'* [概  要] ContainsWhitespace
'* [詳  細] 空白文字が含まれているかどうかをチェック｡
'*
'* @param str 文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function ContainsWhitespace(ByVal str As String) As Boolean
    If Me.IsEmpty(str) Then
        ContainsWhitespace = False
        Exit Function
    End If
    ContainsWhitespace = Me.Contains(str, " ")
End Function

'******************************************************************************
'* [概  要] CountMatches
'* [詳  細] 文字が指定された文字列に出現する回数をカウントする｡
'*
'* [用  例]
'*   StringUtils.CountMatches("", *)       = 0
'*   StringUtils.CountMatches("abba", 0)   = 0
'*   StringUtils.CountMatches("abba", "a") = 2
'*   StringUtils.CountMatches("abba", "b") = 2
'*   StringUtils.CountMatches("abba", "x") = 0
'*
'* @param str 文字列
'* @param searchStr 検索文字列
'* @return Long 出現回数
'*
'******************************************************************************
Public Function CountMatches(ByVal str As String, ByVal searchStr As String) As Long
    If Me.IsEmpty(str) Or Me.IsEmpty(searchStr) Then
        CountMatches = 0
        Exit Function
    End If
    Dim Count As Long: Count = 0
    Dim idx As Long: idx = 0
    
    idx = Me.IndexOf(str, searchStr, idx)
    Do While idx <> INDEX_NOT_FOUND
        Count = Count + 1
        idx = idx + Len(searchStr)
        idx = Me.IndexOf(str, searchStr, idx)
    Loop
    CountMatches = Count
End Function

'******************************************************************************
'* [概  要] DefaultIfBlank
'* [詳  細] 渡された文字列がスペースか空（""）の場合は、指定したデフォルト値を返す｡
'*
'* [用  例]
'*   StringUtils.DefaultIfBlank("", "NULL") = "NULL"
'*   StringUtils.DefaultIfBlank(" ", "NULL") = "NULL"
'*   StringUtils.DefaultIfBlank("bat", "NULL") = "bat"
'*
'* @param str 文字列
'* @param defaultStr デフォルト文字列
'* @return String チェック済文字列
'*
'******************************************************************************
Public Function DefaultIfBlank(ByVal str As String, ByVal defaultStr As String) As String
    DefaultIfBlank = IIf(Me.IsBlank(str), defaultStr, str)
End Function

'******************************************************************************
'* [概  要] DefaultIfEmpty
'* [詳  細] 渡された文字列が空（""）の場合は、指定したデフォルト値を返す｡
'*
'* [用  例]
'*   StringUtils.DefaultIfEmpty("", "NULL") = "NULL"
'*   StringUtils.DefaultIfEmpty(" ", "NULL") = " "
'*   StringUtils.DefaultIfEmpty("bat", "NULL") = "bat"
'*
'* @param str 文字列
'* @param defaultStr デフォルト文字列
'* @return String チェック済文字列
'*
'******************************************************************************
Public Function DefaultIfEmpty(ByVal str As String, ByVal defaultStr As String) As String
    DefaultIfEmpty = IIf(Me.IsEmpty(str), defaultStr, str)
End Function

'******************************************************************************
'* [概  要] DeleteWhitespace
'* [詳  細] 文字列からすべての半角スペースを削除｡
'*   StringUtils.DeleteWhitespace("") = ""
'*   StringUtils.DeleteWhitespace("abc") = "abc"
'*   StringUtils.DeleteWhitespace("   ab  c  ") = "abc"
'*
'* [用  例]
'*
'* @param str 文字列
'* @return String 削除後文字列
'*
'******************************************************************************
Public Function DeleteWhitespace(ByVal str As String) As String
    DeleteWhitespace = VBA.Replace(str, " ", "")
End Function

'******************************************************************************
'* [概  要] Difference
'* [詳  細] 2つの文字列を比較し、それらが異なる部分を返す。
'*          より正確には、最初の文字列とは異なる場所から、2番目の文字列の残りを返す。
'*          「abc」と「ab」の違いは空の文字列であり、「c」ではない。
'*
'* [用  例]
'*   StringUtils.Difference("i am a machine", "i am a robot") = "robot"
'*   StringUtils.Difference("", "") = ""
'*   StringUtils.Difference("", "abc") = "abc"
'*   StringUtils.Difference("abc", "") = ""
'*   StringUtils.Difference("abc", "abc") = ""
'*   StringUtils.Difference("abc", "ab") = ""
'*   StringUtils.Difference("ab", "abxyz") = "xyz"
'*   StringUtils.Difference("abcde", "abxyz") = "xyz"
'*   StringUtils.Difference("abcde", "xyz") = "xyz"
'*
'* @param str1 比較する文字列1
'* @param str2 比較する文字列2
'* @return String 異なる部分の文字列。等しい場合は空の文字列を返す
'*
'******************************************************************************
Public Function Difference(ByVal str1 As String, ByVal str2 As String) As String
    If Me.IsEmpty(str1) Then
        Difference = str2
        Exit Function
    End If
    If Me.IsEmpty(str2) Then
        Difference = ""
        Exit Function
    End If
    Dim at As Long: at = Me.IndexOfDifference(str1, str2)
    If at = INDEX_NOT_FOUND Then
        Difference = ""
        Exit Function
    End If
    Difference = Me.Substring(str2, at)
End Function

'******************************************************************************
'* [概  要] EndsWithAny
'* [詳  細] 指定した文字列のセットのいづれかで終るかチェック｡
'*
'* [用  例]
'*   StringUtils.EndsWithAny("abcxyz", Array("xyz", "abc")) = True
'*   StringUtils.EndsWithAny("abcXYZ", "def", "XYZ") = True
'*   StringUtils.EndsWithAny("abcXYZ", "def", "xyz") = False
'*
'* @param str 文字列
'* @param endStr 終了文字列（複数指定）
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function EndsWithAny(ByVal str As String, ParamArray endStr() As Variant) As Boolean

    If Me.IsEmpty(str) Or VBA.IsMissing(endStr) Then
        EndsWithAny = False
        Exit Function
    End If
    Dim tmp: tmp = endStr ' ParamArrayを別関数に渡すための対応
    Dim v
    For Each v In ArrayUtils.ToVArrayFlattened(tmp)
        If Me.EndsWith(str, CStr(v)) Then
            EndsWithAny = True
            Exit Function
        End If
    Next
    EndsWithAny = False
End Function

'******************************************************************************
'* [概  要] EndsWithIgnoreCase
'* [詳  細] 指定した文字列で終るかチェック｡
'*          大文字小文字を考慮しない。
'*
'* [用  例]
'*   StringUtils.EndsWithIgnoreCase("abcdef", "def") = True
'*   StringUtils.EndsWithIgnoreCase("ABCDEF", "def") = True
'*   StringUtils.EndsWithIgnoreCase("ABCDEF", "cde") = False
'*
'* @param str 文字列
'* @param endStr 終了文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function EndsWithIgnoreCase(ByVal str As String, ByVal endStr As String) As Boolean
    If Me.IsEmpty(str) Or Me.IsEmpty(endStr) Then
        EndsWithIgnoreCase = False
        Exit Function
    End If
    Dim strIc As String: strIc = Me.LowerCase(str)
    Dim endStrIc As String: endStrIc = Me.LowerCase(endStr)
    EndsWithIgnoreCase = Me.EndsWith(strIc, endStrIc)
End Function

'******************************************************************************
'* [概  要] Equals
'* [詳  細] 2つの文字列が等しいかチェック｡
'*
'* [用  例]
'*   StringUtils.Equals("abc", "abc") = True
'*   StringUtils.Equals("abc", "ABC") = False
'*
'* @param str1 文字列1
'* @param str2 文字列2
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function Equals(ByVal str1 As String, ByVal str2 As String) As Boolean
    Equals = (str1 = str2)
End Function

'******************************************************************************
'* [概  要] EqualsIgnoreCase
'* [詳  細] 2つの文字列が等しいかチェック｡
'*          大文字小文字を考慮しない。
'*
'* [用  例]
'*   StringUtils.EqualsIgnoreCase("abc", "abc") = True
'*   StringUtils.EqualsIgnoreCase("abc", "ABC") = True
'*
'* @param str1 文字列1
'* @param str2 文字列2
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function EqualsIgnoreCase(ByVal str1 As String, ByVal str2 As String) As Boolean
    Dim str1Ic As String: str1Ic = Me.LowerCase(str1)
    Dim str2Ic As String: str2Ic = Me.LowerCase(str2)
    EqualsIgnoreCase = Me.Equals(str1Ic, str2Ic)
End Function

'******************************************************************************
'* [概  要] EqualsAny
'* [詳  細] 文字列が指定した文字列のいづれかに等しいかチェック｡
'*
'* [用  例]
'*   StringUtils.EqualsAny("abc", "abc", "def") = True
'*   StringUtils.EqualsAny("abc", "ABC", "DEF") = False
'*
'* @param str 文字列
'* @param searchStrings 比較文字列（複数指定）
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function EqualsAny(ByVal str As String, ParamArray searchStrings() As Variant) As Boolean
    If VBA.IsMissing(searchStrings) Then
        EqualsAny = False
        Exit Function
    End If
    Dim tmp: tmp = searchStrings ' ParamArrayを別関数に渡すための対応
    Dim v
    For Each v In ArrayUtils.ToVArrayFlattened(tmp)
        If Me.Equals(str, CStr(v)) Then
            EqualsAny = True
            Exit Function
        End If
    Next
    EqualsAny = False
End Function

'******************************************************************************
'* [概  要] EqualsAnyIgnoreCase
'* [詳  細] 文字列が指定した文字列のいづれかに等しいかチェック｡
'*          大文字小文字を考慮しない。
'*
'* [用  例]
'*   StringUtils.EqualsAnyIgnoreCase("abc", "abc", "def") = True
'*   StringUtils.EqualsAnyIgnoreCase("abc", "ABC", "DEF") = True
'*
'* @param str1 文字列1
'* @param str2 文字列2
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function EqualsAnyIgnoreCase(ByVal str As String, ParamArray searchStrings() As Variant) As Boolean
    If VBA.IsMissing(searchStrings) Then
        EqualsAnyIgnoreCase = False
        Exit Function
    End If
    Dim tmp: tmp = searchStrings ' ParamArrayを別関数に渡すための対応
    Dim v
    For Each v In ArrayUtils.ToVArrayFlattened(tmp)
        If Me.EqualsIgnoreCase(str, CStr(v)) Then
            EqualsAnyIgnoreCase = True
            Exit Function
        End If
    Next
    EqualsAnyIgnoreCase = False
End Function

'******************************************************************************
'* [概  要] FirstNonBlank
'* [詳  細] 空（""）、スペースのみではない文字配列の最初の値を返す。
'*          すべての値が空白であるか、配列が空である場合、空が返される。
'*
'* [用  例]
'*   StringUtils.FirstNonBlank("", " ") = ""
'*   StringUtils.FirstNonBlank("abc") = "abc"
'*   StringUtils.FirstNonBlank("", "xyz") = "xyz"
'*   StringUtils.FirstNonBlank("", " ", "xyz") = "xyz"
'*   StringUtils.FirstNonBlank() = ""
'*
'* @param values 文字列（複数指定）
'* @return String 該当文字列
'*
'******************************************************************************
Public Function FirstNonBlank(ParamArray values() As Variant) As String
    If IsMissing(values) Then
        FirstNonBlank = ""
        Exit Function
    End If
    Dim tmp: tmp = values ' ParamArrayを別関数に渡すための対応
    Dim v
    For Each v In ArrayUtils.ToVArrayFlattened(tmp)
        If Me.IsNotBlank(CStr(v)) Then
            FirstNonBlank = CStr(v)
            Exit Function
        End If
    Next
    FirstNonBlank = ""
End Function


'******************************************************************************
'* [概  要] FirstNonEmpty
'* [詳  細] 空（""）ではない文字配列の最初の値を返す。
'*          すべての値が空である場合、空が返される。
'*
'* [用  例]
'*   StringUtils.FirstNonEmpty("", " ") = " "
'*   StringUtils.FirstNonEmpty("abc") = "abc"
'*   StringUtils.FirstNonEmpty("", "xyz") = "xyz"
'*   StringUtils.FirstNonEmpty("", " ", "xyz") = " "
'*   StringUtils.FirstNonEmpty() = ""
'*
'* @param values 文字列（複数指定）
'* @return String 該当文字列
'*
'******************************************************************************
Public Function FirstNonEmpty(ParamArray values() As Variant) As String
    If IsMissing(values) Then
        FirstNonEmpty = ""
        Exit Function
    End If
    Dim tmp: tmp = values ' ParamArrayを別関数に渡すための対応
    Dim v
    For Each v In ArrayUtils.ToVArrayFlattened(tmp)
        If Me.IsNotEmpty(CStr(v)) Then
            FirstNonEmpty = CStr(v)
            Exit Function
        End If
    Next
    FirstNonEmpty = ""
End Function


'******************************************************************************
'* [概  要] GetBytes
'* [詳  細] 文字列をバイト配列に変換｡
'*
'* @param str 文字列
'* @param encode エンコードのキャラセット
'* @return Byte() バイト配列
'*
'******************************************************************************
Public Function GetBytes(ByVal str As String, Optional ByVal encode As String) As Byte()
    If Me.IsEmpty(str) Then
       GetBytes = ArrayUtils.EmptyByteArray
       Exit Function
    End If
    If encode = "" Then
       GetBytes = str
       Exit Function
    End If
    
    With IO.CreateADODBStream(adTypeText, encode)
        .mode = adModeReadWrite
        .Open
        .WriteText str
        .Position = 0
        .Type = adTypeBinary
        GetBytes = .Read
        .Close
    End With
End Function


'******************************************************************************
'* [概  要] GetCommonPrefix
'* [詳  細] 配列内のすべての文字列を比較し、共通する接頭辞を返却する。
'*
'* [用  例]
'*   StringUtils.GetCommonPrefix(Array("abc")) = "abc"
'*   StringUtils.GetCommonPrefix(Array("", "")) = ""
'*   StringUtils.GetCommonPrefix(Array("", "abc")) = ""
'*   StringUtils.GetCommonPrefix(Array("abc", "")) = ""
'*   StringUtils.GetCommonPrefix(Array("abc", "abc")) = "abc"
'*   StringUtils.GetCommonPrefix(Array("abc", "a")) = "a"
'*   StringUtils.GetCommonPrefix(Array("ab", "abxyz")) = "ab"
'*   StringUtils.GetCommonPrefix(Array("abcde", "abxyz")) = "ab"
'*   StringUtils.GetCommonPrefix(Array("abcde", "xyz")) = ""
'*   StringUtils.GetCommonPrefix(Array("xyz", "abcde")) = ""
'*   StringUtils.GetCommonPrefix(Array("i am a machine", "i am a robot")) = "i am a "
'*
'* @param strs 文字配列
'* @return String 共通する接頭辞文字列
'*
'******************************************************************************
Public Function GetCommonPrefix(ParamArray strs() As Variant) As String
    If IsMissing(strs) Then
        GetCommonPrefix = ""
        Exit Function
    End If
    Dim tmp: tmp = strs ' ParamArrayを別関数に渡すための対応
    Dim smallestIndexOfDiff As Long: smallestIndexOfDiff = Me.IndexOfDifference(tmp)
    If smallestIndexOfDiff = INDEX_NOT_FOUND Then
        ' all strings were identical
        GetCommonPrefix = CStr(strs(0))
        Exit Function
    ElseIf smallestIndexOfDiff = 0 Then
        ' there were no common initial characters
        GetCommonPrefix = ""
        Exit Function
    End If
    ' we found a common initial character sequence
    GetCommonPrefix = Me.Substring(CStr(strs(0)), 0, smallestIndexOfDiff)
End Function


'******************************************************************************
'* [概  要] GetDigits
'* [詳  細] 文字列strに数字が含まれているかどうかをチェックし、
'*          含まれている場合はstr内のすべての数字を連結して文字列として返す｡
'*          数字が見つからない場合は、空の（""）文字列が返される。
'*
'* [用  例]
'*   StringUtils.GetDigits("") = ""
'*   StringUtils.GetDigits("abc") = ""
'*   StringUtils.GetDigits("1000$") = "1000"
'*   StringUtils.GetDigits("1123~45") = "112345"
'*   StringUtils.GetDigits("(541) 754-3010") = "5417543010"
'*
'* @param str 文字列
'* @return String 数字文字列
'*
'******************************************************************************
Public Function GetDigits(ByVal str As String) As String
    If Me.IsEmpty(str) Then
        GetDigits = False
        Exit Function
    End If
    Dim sz As Long: sz = Len(str)
    Dim strDigits As StringEx: strDigits = Core.Init(New StringEx, "", sz)
    Dim i As Long
    For i = 0 To sz - 1
        Dim tempChar As String: tempChar = Me.CharAt(str, i)
        If ValidateUtils.IsNumber(tempChar) Then
            Call strDigits.Append(tempChar)
        End If
    Next
    GetDigits = strDigits.ToString
End Function


'******************************************************************************
'* [概  要] GetIfBlank
'* [詳  細] 渡された文字列がスペースか空（""）の場合は、指定した関数の結果を返す｡
'*
'* [用  例]
'*   StringUtils.GetIfBlank("", () -> "NULL")     = "NULL"
'*   StringUtils.GetIfBlank(" ", () -> "NULL")    = "NULL"
'*   StringUtils.GetIfBlank("bat", () -> "NULL")  = "bat"
'*    ※() -> **** 部分は実際はFuncクラスを指定する。
'*
'* [依　存]
'*  Ariawase 0.9.0 Funcクラス
'*
'* @param str 文字列
'* @param defaultSupplier デフォルト文字列を返すサプライヤー（Funcオブジェクト）
'* @return String チェック済文字列
'*
'******************************************************************************
Public Function GetIfBlank(ByVal str As String, ByVal defaultSupplier As Func) As String
    If Me.IsBlank(str) Then
        GetIfBlank = defaultSupplier.Apply()
        Exit Function
    End If
    GetIfBlank = str
End Function


'******************************************************************************
'* [概  要] GetIfEmpty
'* [詳  細] 渡された文字列が空（""）の場合は、指定した関数の結果を返す｡
'*
'* [用  例]
'*   StringUtils.GetIfBlank("", () -> "NULL")     = "NULL"
'*   StringUtils.GetIfBlank(" ", () -> "NULL")    = " "
'*   StringUtils.GetIfBlank("bat", () -> "NULL")  = "bat"
'*    ※() -> **** 部分は実際はFuncクラスを指定する。
'*
'* [依　存]
'*  Ariawase 0.9.0 Funcクラス
'*
'* @param str 文字列
'* @param defaultSupplier デフォルト文字列を返すサプライヤー（Funcオブジェクト）
'* @return String チェック済文字列
'*
'******************************************************************************
Public Function GetIfEmpty(ByVal str As String, ByVal defaultSupplier As Func) As String
    If Me.IsEmpty(str) Then
        GetIfEmpty = defaultSupplier.Apply()
        Exit Function
    End If
End Function


'******************************************************************************
'* [概  要] IndexOf
'* [詳  細] 文字列内の最初のインデックス（文字位置）を検索する｡
'*
'* [用  例]
'*   StringUtils.IndexOf("", *)            = -1 (except when * = "")
'*   StringUtils.IndexOf("aabaabaa", "a") = 0
'*   StringUtils.IndexOf("aabaabaa", "b") = 2
'*   StringUtils.IndexOf("aabaabaa", "ab") = 1
'*   StringUtils.IndexOf("aabaabaa", "") = 0
'*   StringUtils.IndexOf("aabaabaa", "a", 0) = 0
'*   StringUtils.IndexOf("aabaabaa", "b", 0) = 2
'*   StringUtils.IndexOf("aabaabaa", "ab", 0) = 1
'*   StringUtils.IndexOf("aabaabaa", "b", 3) = 5
'*   StringUtils.IndexOf("aabaabaa", "b", 9) = -1
'*   StringUtils.IndexOf("aabaabaa", "b", -1) = 2
'*   StringUtils.IndexOf("aabaabaa", "", 2) = 2
'*   StringUtils.IndexOf("abc", "", 9) = 3
'*
'* @param str 文字列
'* @param searchStr 検索文字列
'* @param startPos 開始位置。負の値はゼロとして扱われる。
'* @return Long 文字列の最初のインデックス。一致がない場合、空文字列入力の場合は-1
'*
'******************************************************************************
Public Function IndexOf(ByVal str As String, ByVal searchStr As String, Optional ByVal startPos As Long = 0) As Long
    If Me.IsEmpty(str) Then
        If Not Me.IsEmpty(searchStr) Then
            IndexOf = INDEX_NOT_FOUND
        Else
            IndexOf = 0
        End If
        Exit Function
    End If
    If startPos < 0 Then
        startPos = 0
    ElseIf startPos >= Len(str) Then
        IndexOf = Len(str) - 1
        Exit Function
    End If
    IndexOf = InStr(startPos + 1, str, searchStr) - 1
End Function


'******************************************************************************
'* [概  要] IndexOfAny
'* [詳  細] 指定された文字セット内の任意の文字の最初のインデックス（文字位置）を検索する｡
'*
'* [用  例]
'*   StringUtils.indexOfAny(*, [])                        = -1
'*   StringUtils.IndexOfAny("zzabyycdxx", ["ab", "cd"]) = 2
'*   StringUtils.IndexOfAny("zzabyycdxx", ["cd", "ab"]) = 2
'*   StringUtils.IndexOfAny("zzabyycdxx", ["mn", "op"]) = -1
'*   StringUtils.IndexOfAny("zzabyycdxx", ["zab", "aby"]) = 1
'*   StringUtils.IndexOfAny("zzabyycdxx", [""]) = 0
'*   StringUtils.IndexOfAny("", [""]) = 0
'*   StringUtils.IndexOfAny("", ["a"]) = -1
'*
'* @param str 文字列
'* @param searchStrs 検索文字列（複数指定）
'* @return Long 文字列の最初のインデックス。一致がない場合、空配列入力の場合は-1
'*
'******************************************************************************
Public Function IndexOfAny(ByVal str As String, ParamArray searchStrs() As Variant) As Long
    If IsMissing(searchStrs) Then
        IndexOfAny = INDEX_NOT_FOUND
         Exit Function
    End If
    Dim ret As Long: ret = MAX_LONG
    Dim i As Long: i = 0

    Dim tmp: tmp = searchStrs ' ParamArrayを別関数に渡すための対応
    Dim v
    For Each v In ArrayUtils.ToVArrayFlattened(tmp)
        i = Me.IndexOf(str, CStr(v))
        If i <> INDEX_NOT_FOUND Then
            If i < ret Then
                ret = i
            End If
        End If
    Next
    IndexOfAny = IIf(ret = MAX_LONG, INDEX_NOT_FOUND, ret)
End Function


'******************************************************************************
'* [概  要] IndexOfAnyBut
'* [詳  細] 指定された文字セットにない任意の文字の最初のインデックス（文字位置）
'*          を検索する｡
'*
'* [用  例]
'*   StringUtils.indexOfAnyBut("", *)                                = -1
'*   StringUtils.indexOfAnyBut(*, [])                                = -1
'*   StringUtils.indexOfAnyBut("zzabyycdxx", new char[] {'z', 'a'} ) = 3
'*   StringUtils.indexOfAnyBut("aba", new char[] {'z'} )             = 0
'*   StringUtils.indexOfAnyBut("aba", new char[] {'a', 'b'} )        = -1
'*
'* @param str 文字列
'* @param searchStrs 検索文字（１文字。複数指定）
'* @return Long 文字列の最初のインデックス。一致がない場合、空文字列入力の場合は-1
'*
'******************************************************************************
Public Function IndexOfAnyBut(ByVal str As String, ParamArray searchStrs() As Variant) As Long
    If Me.IsEmpty(str) Or IsMissing(searchStrs) Then
        IndexOfAnyBut = INDEX_NOT_FOUND
        Exit Function
    End If
    Dim strLen As Long: strLen = Len(str)
    Dim tmp: tmp = searchStrs ' ParamArrayを別関数に渡すための対応
    Dim v, i As Long: i = 0
    For Each v In Me.ToCharArray(str)
        If Me.EqualsAny(CStr(v), tmp) Then
            Exit For
        Else
            i = i + 1
        End If
    Next
    If i = Len(str) Then
        IndexOfAnyBut = INDEX_NOT_FOUND
        Exit Function
    End If
    IndexOfAnyBut = i
End Function


'******************************************************************************
'* [概  要] IndexOfDifference
'* [詳  細] 2つの文字列を比較し、文字が異なり始めるインデックスを返す｡
'*
'* [用  例]
'*   StringUtils.IndexOfDifference("", "") = -1
'*   StringUtils.IndexOfDifference("", "abc") = 0
'*   StringUtils.IndexOfDifference("abc", "") = 0
'*   StringUtils.IndexOfDifference("abc", "abc") = -1
'*   StringUtils.IndexOfDifference("ab", "abxyz") = 2
'*   StringUtils.IndexOfDifference("abcde", "abxyz") = 2
'*   StringUtils.IndexOfDifference("abcde", "xyz") = 0
'*
'* @param str 文字列
'* @param startStr 開始文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IndexOfDifference(ParamArray strs() As Variant) As Long
    If IsMissing(strs) Then
        IndexOfDifference = INDEX_NOT_FOUND
        Exit Function
    End If
    Dim var: var = strs ' ParamArrayを別関数に渡すための対応
    If ArrayUtils.GetLength(var) <= 1 Then
        IndexOfDifference = INDEX_NOT_FOUND
        Exit Function
    ElseIf ArrayUtils.GetLength(var) = 2 Then
        Dim str1 As String: str1 = strs(LBound(strs))
        Dim str2 As String: str2 = strs(UBound(strs))
        If Me.Equals(str1, str2) Then
            IndexOfDifference = INDEX_NOT_FOUND
            Exit Function
        End If
        Dim i As Long: i = 0
        Do While i < Len(str1) And i < Len(str2)
            If Me.CharAt(str1) <> Me.CharAt(str2) Then
                Exit Do
            End If
            i = i + 1
        Loop
        If i < Len(str1) Or i < Len(str2) Then
            IndexOfDifference = i
            Exit Function
        End If
        IndexOfDifference = INDEX_NOT_FOUND
        Exit Function
    End If

    Dim anyStringEmpty  As Boolean: anyStringEmpty = False
    Dim allStringsEmpty As Boolean: allStringsEmpty = True
    Dim ArrayLen        As Long:    ArrayLen = ArrayUtils.GetLength(var)
    Dim shortestStrLen  As Long:    shortestStrLen = MAX_LONG
    Dim longestStrLen   As Long:    longestStrLen = 0

    ' find the min and max string lengths; this avoids checking to make
    ' sure we are not exceeding the length of the string each time through
    ' the bottom loop.
    Dim s
    For Each s In strs

       If Me.IsEmpty(CStr(s)) Then
           anyStringEmpty = True
           shortestStrLen = 0
       Else
           allStringsEmpty = False
           shortestStrLen = Core.Min(Len(s), shortestStrLen)
           longestStrLen = Core.Max(Len(s), longestStrLen)
       End If
    Next

    ' handle lists containing all nulls or all empty strings
    If allStringsEmpty Or longestStrLen = 0 And Not anyStringEmpty Then
        IndexOfDifference = INDEX_NOT_FOUND
        Exit Function
    End If

    ' handle lists containing some nulls or some empty strings
    If shortestStrLen = 0 Then
        IndexOfDifference = INDEX_NOT_FOUND
        Exit Function
    End If

    ' find the position with the first difference across all strings
    Dim firstDiff As Long: firstDiff = -1
    Dim stringPos As Long
    For stringPos = 0 To shortestStrLen - 1
        Dim comparisonChar As String: comparisonChar = Me.CharAt(CStr(strs(0)), stringPos)
        Dim arrayPos As Long: arrayPos = 1
        For arrayPos = 1 To ArrayLen - 1
            If Me.CharAt(CStr(strs(arrayPos)), stringPos) <> comparisonChar Then
                firstDiff = stringPos
                Exit For
            End If
        Next
        If firstDiff <> -1 Then
            Exit For
        End If
    Next

    If firstDiff = -1 And shortestStrLen <> longestStrLen Then
        ' we compared all of the characters up to the length of the
        ' shortest string and didn't find a match, but the string lengths
        ' vary, so return the length of the shortest string.
        IndexOfDifference = shortestStrLen
        Exit Function
    End If
    IndexOfDifference = firstDiff
End Function


'******************************************************************************
'* [概  要] IndexOfIgnoreCase
'* [詳  細] 文字列内の最初のインデックス（文字位置）を検索する｡
'*          大文字小文字を考慮しない。
'*
'* [用  例]
'*   StringUtils.IndexOfIgnoreCase("", "") = 0
'*   StringUtils.IndexOfIgnoreCase("Aabaabaa", "a") = 0
'*   StringUtils.IndexOfIgnoreCase("aaBaabaa", "b") = 2
'*   StringUtils.IndexOfIgnoreCase("aABaabaa", "ab") = 1
'*
'* @param str 文字列
'* @param searchStr 検索文字列
'* @param startPos 開始位置。負の値はゼロとして扱われる。
'* @return Long 文字列の最初のインデックス。一致がない場合、空文字列入力の場合は-1
'*
'******************************************************************************
Public Function IndexOfIgnoreCase(ByVal str As String, ByVal searchStr As String, Optional ByVal startPos As Long = 0) As Long
    Dim strIc As String: strIc = Me.LowerCase(str)
    Dim searchStrIc As String: searchStrIc = Me.LowerCase(searchStr)
    IndexOfIgnoreCase = Me.IndexOf(strIc, searchStrIc, startPos)
End Function

'******************************************************************************
'* [概  要] IsAllBlank
'* [詳  細] すべての文字列が空（""）、またはスペースのみかどうかをチェック｡
'*
'* [用  例]
'*   StringUtils.IsAllBlank(" ", " ") = True
'*   StringUtils.IsAllBlank("", "") = True
'*   StringUtils.IsAllBlank("", "bar") = False
'*   StringUtils.IsAllBlank("bob", "") = False
'*   StringUtils.IsAllBlank("  bob  ", "") = False
'*   StringUtils.IsAllBlank(" ", "bar") = False
'*   StringUtils.IsAllBlank("foo", "bar") = False
'*
'* @param strs() 文字列（複数指定）
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsAllBlank(ParamArray strs() As Variant) As Boolean
    If VBA.IsMissing(strs) Then
        IsAllBlank = True
        Exit Function
    End If
    Dim tmp: tmp = strs ' ParamArrayを別関数に渡すための対応
    Dim v
    For Each v In ArrayUtils.ToVArrayFlattened(tmp)
        If Me.IsNotBlank(CStr(v)) Then
            IsAllBlank = False
            Exit Function
        End If
    Next
    IsAllBlank = True
End Function

'******************************************************************************
'* [概  要] IsAllEmpty
'* [詳  細] すべての文字列が空（""のみかどうかをチェック｡
'*
'* [用  例]
'*   StringUtils.IsAllBlank(" ", " ") = False
'*   StringUtils.IsAllBlank("", "") = True
'*   StringUtils.IsAllBlank("", "bar") = False
'*   StringUtils.IsAllBlank("bob", "") = False
'*   StringUtils.IsAllBlank("  bob  ", "") = False
'*   StringUtils.IsAllBlank(" ", "bar") = False
'*   StringUtils.IsAllBlank("foo", "bar") = False
'*
'* @param strs() 文字列（複数指定）
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsAllEmpty(ParamArray strs() As Variant) As Boolean
    If VBA.IsMissing(strs) Then
        IsAllEmpty = True
        Exit Function
    End If
    Dim tmp: tmp = strs ' ParamArrayを別関数に渡すための対応
    Dim v
    For Each v In ArrayUtils.ToVArrayFlattened(tmp)
        If Me.IsNotEmpty(CStr(v)) Then
            IsAllEmpty = False
            Exit Function
        End If
    Next
    IsAllEmpty = True
End Function

'******************************************************************************
'* [概  要] IsLowerCase
'* [詳  細] 小文字のみかどうかをチェック｡
'*
'* [用  例]
'*   StringUtils.IsAllLowerCase("") = False
'*   StringUtils.IsAllLowerCase("  ") = False
'*   StringUtils.IsAllLowerCase("abc") = True
'*   StringUtils.IsAllLowerCase("abC") = False
'*   StringUtils.IsAllLowerCase("ab c") = False
'*   StringUtils.IsAllLowerCase("ab1c") = False
'*   StringUtils.IsAllLowerCase("ab/c") = False
'*
'* @param str 文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsLowerCase(ByVal str As String) As Boolean
    If Me.IsEmpty(str) Then
        IsLowerCase = False
        Exit Function
    End If
    Dim v
    For Each v In Me.ToCharArray(str)
        If Not CharUtils.IsAsciiAlphaLower(CStr(v)) Then
            IsLowerCase = False
            Exit Function
        End If
    Next
    IsLowerCase = True
End Function

'******************************************************************************
'* [概  要] IsUpperCase
'* [詳  細] 大文字のみかどうかをチェック｡
'*
'* [用  例]
'*   StringUtils.IsAllUpperCase("") = False
'*   StringUtils.IsAllUpperCase("  ") = False
'*   StringUtils.IsAllUpperCase("ABC") = True
'*   StringUtils.IsAllUpperCase("aBC") = False
'*   StringUtils.IsAllUpperCase("A C") = False
'*   StringUtils.IsAllUpperCase("A1C") = False
'*   StringUtils.IsAllUpperCase("A/C") = False
'*
'* @param str 文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsUpperCase(ByVal str As String) As Boolean
    If Me.IsEmpty(str) Then
        IsUpperCase = False
        Exit Function
    End If
    Dim v
    For Each v In Me.ToCharArray(str)
        If Not CharUtils.IsAsciiAlphaUpper(CStr(v)) Then
            IsUpperCase = False
            Exit Function
        End If
    Next
    IsUpperCase = True
End Function

'******************************************************************************
'* [概  要] IsAlpha
'* [詳  細] 文字列が英字かどうかをチェック｡
'*
'* [用  例]
'*   StringUtils.IsAlpha("") = False
'*   StringUtils.IsAlpha("  ") = False
'*   StringUtils.IsAlpha("abc") = True
'*   StringUtils.IsAlpha("ab2c") = False
'*   StringUtils.IsAlpha("ab-c") = False
'*
'* @param str 文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsAlpha(ByVal str As String) As Boolean
    If Me.IsEmpty(str) Then
        IsAlpha = False
        Exit Function
    End If
    Dim v
    For Each v In Me.ToCharArray(str)
        If Not CharUtils.IsAsciiAlpha(CStr(v)) Then
            IsAlpha = False
            Exit Function
        End If
    Next
    IsAlpha = True
End Function


'******************************************************************************
'* [概  要] IsAlphanumeric
'* [詳  細] 文字列が英数字かどうかをチェック｡
'*
'* [用  例]
'*   StringUtils.IsAlphanumeric("") = False
'*   StringUtils.IsAlphanumeric("  ") = False
'*   StringUtils.IsAlphanumeric("abc") = True
'*   StringUtils.IsAlphanumeric("ab c") = False
'*   StringUtils.IsAlphanumeric("ab2c") = True
'*   StringUtils.IsAlphanumeric("ab-c") = False
'*
'* @param str 文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsAlphanumeric(ByVal str As String) As Boolean
    If Me.IsEmpty(str) Then
        IsAlphanumeric = False
        Exit Function
    End If
    Dim v
    For Each v In Me.ToCharArray(str)
        If Not CharUtils.IsAsciiAlphanumeric(CStr(v)) Then
            IsAlphanumeric = False
            Exit Function
        End If
    Next
    IsAlphanumeric = True
End Function


'******************************************************************************
'* [概  要] IsAlphanumericSpace
'* [詳  細] 文字列が英数字とスペースかどうかをチェック｡
'*
'* [用  例]
'*   StringUtils.IsAlphanumericSpace("") = True
'*   StringUtils.IsAlphanumericSpace("  ") = True
'*   StringUtils.IsAlphanumericSpace("abc") = True
'*   StringUtils.IsAlphanumericSpace("ab c") = True
'*   StringUtils.IsAlphanumericSpace("ab2c") = True
'*   StringUtils.IsAlphanumericSpace("ab-c") = False
'*
'* @param str 文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsAlphanumericSpace(ByVal str As String) As Boolean
    If Me.IsEmpty(str) Then
        IsAlphanumericSpace = False
        Exit Function
    End If
    Dim v
    For Each v In Me.ToCharArray(str)
        If Not CharUtils.IsAsciiAlphanumeric(CStr(v)) And CStr(v) <> " " Then
            IsAlphanumericSpace = False
            Exit Function
        End If
    Next
    IsAlphanumericSpace = True
End Function

'******************************************************************************
'* [概  要] IsAlphaSpace
'* [詳  細] 文字列が英字とスペースかどうかをチェック｡
'*
'* [用  例]
'*   StringUtils.IsAlphaSpace("") = True
'*   StringUtils.IsAlphaSpace("  ") = True
'*   StringUtils.IsAlphaSpace("abc") = True
'*   StringUtils.IsAlphaSpace("ab c") = True
'*   StringUtils.IsAlphaSpace("ab2c") = False
'*   StringUtils.IsAlphaSpace("ab-c") = False
'*
'* @param str 文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsAlphaSpace(ByVal str As String) As Boolean
    If Me.IsEmpty(str) Then
        IsAlphaSpace = False
        Exit Function
    End If
    Dim v
    For Each v In Me.ToCharArray(str)
        If Not CharUtils.IsAsciiAlpha(CStr(v)) And CStr(v) <> " " Then
            IsAlphaSpace = False
            Exit Function
        End If
    Next
    IsAlphaSpace = True
End Function

'******************************************************************************
'* [概  要] IsAnyBlank
'* [詳  細] 文字列のいずれかが空（""）か空白のみかをチェック｡
'*
'* [用  例]
'*   StringUtils.IsAnyBlank("", "bar") = True
'*   StringUtils.IsAnyBlank("bob", "") = True
'*   StringUtils.IsAnyBlank(" ", "bar") = True
'*   StringUtils.IsAnyBlank("foo", "bar") = False
'*
'* @param str 文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsAnyBlank(ParamArray strs() As Variant) As Boolean
    If IsMissing(strs) Then
        IsAnyBlank = False
        Exit Function
    End If
    
    Dim tmp: tmp = strs ' ParamArrayを別関数に渡すための対応
    Dim v
    For Each v In ArrayUtils.ToVArrayFlattened(tmp)
        If Me.IsBlank(CStr(v)) Then
            IsAnyBlank = True
            Exit Function
        End If
    Next
    
    IsAnyBlank = False
End Function

'******************************************************************************
'* [概  要] IsAnyEmpty
'* [詳  細] 文字列のいずれかが空（""）のみかをチェック｡
'*
'* [用  例]
'*   StringUtils.IsAnyEmpty("", "bar") = True
'*   StringUtils.IsAnyEmpty("bob", "") = True
'*   StringUtils.IsAnyEmpty(" ", "bar") = False
'*   StringUtils.IsAnyEmpty("foo", "bar") = False
'*
'* @param str 文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsAnyEmpty(ParamArray strs() As Variant) As Boolean
    If IsMissing(strs) Then
        IsAnyEmpty = False
        Exit Function
    End If
    
    Dim tmp: tmp = strs ' ParamArrayを別関数に渡すための対応
    Dim v
    For Each v In ArrayUtils.ToVArrayFlattened(tmp)
        If Me.IsEmpty(CStr(v)) Then
            IsAnyEmpty = True
            Exit Function
        End If
    Next
    
    IsAnyEmpty = False
End Function

'******************************************************************************
'* [概  要] IsAsciiPrintable
'* [詳  細] 文字列にASCII印刷可能文字のみが含まれているかどうかをチェック｡
'*
'* [用  例]
'*   StringUtils.IsAsciiPrintable("") = True
'*   StringUtils.IsAsciiPrintable(" ") = True
'*   StringUtils.IsAsciiPrintable("Ceki") = True
'*   StringUtils.IsAsciiPrintable("ab2c") = True
'*   StringUtils.IsAsciiPrintable("!ab-c~") = True
'*   StringUtils.IsAsciiPrintable(" ") = True
'*   StringUtils.IsAsciiPrintable("!") = True
'*   StringUtils.IsAsciiPrintable("~") = True
'*   StringUtils.IsAsciiPrintable("") = False
'*
'* @param str 文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsAsciiPrintable(ByVal str As String) As Boolean
    If Me.IsEmpty(str) Then
        IsAsciiPrintable = False
        Exit Function
    End If
    Dim v
    For Each v In Me.ToCharArray(str)
        If Not CharUtils.IsAsciiPrintable(CStr(v)) Then
            IsAsciiPrintable = False
            Exit Function
        End If
    Next
    IsAsciiPrintable = True
End Function

'******************************************************************************
'* [概  要] IsBlank
'* [詳  細] 文字列が空（""）、スペースのみかどうかをチェック｡
'*
'* [用  例]
'*   StringUtils.IsBlank("") = True
'*   StringUtils.IsBlank(" ") = True
'*   StringUtils.IsBlank("bob") = False
'*   StringUtils.IsBlank("  bob  ") = False
'*
'* @param str 文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsBlank(ByVal str As String) As Boolean
    If Me.IsEmpty(str) Then
        IsBlank = True
        Exit Function
    End If
    Dim v
    For Each v In Me.ToCharArray(str)
        If Not CharUtils.IsWhitespace(CStr(v)) Then
            IsBlank = False
            Exit Function
        End If
    Next
    IsBlank = True
End Function

'******************************************************************************
'* [概  要] IsEmpty
'* [詳  細] 文字列が空（""）かどうかをチェック｡
'*
'* [用  例]
'*   StringUtils.IsEmpty("") = True
'*   StringUtils.IsEmpty(" ") = False
'*   StringUtils.IsEmpty("bob") = False
'*   StringUtils.IsEmpty("  bob  ") = False
'*
'* @param str 文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsEmpty(ByVal str As String) As Boolean
    IsEmpty = (str = "")
End Function

'******************************************************************************
'* [概  要] IsMixedCase
'* [詳  細] 文字列が大文字と小文字の両方が混在しているかどうかをチェック｡
'*
'* [用  例]
'*   StringUtils.IsMixedCase("") = False
'*   StringUtils.IsMixedCase("ABC") = False
'*   StringUtils.IsMixedCase("abc") = False
'*   StringUtils.IsMixedCase("aBc") = True
'*   StringUtils.IsMixedCase("A c") = True
'*   StringUtils.IsMixedCase("A1c") = True
'*   StringUtils.IsMixedCase("a/C") = True
'*   StringUtils.IsMixedCase("aC\t") = True
'*
'* @param str 文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsMixedCase(ByVal str As String) As Boolean
    If Me.IsEmpty(str) Or Len(str) = 1 Then
        IsMixedCase = False
        Exit Function
    End If
    Dim containsUppercase As Boolean: containsUppercase = False
    Dim containsLowercase As Boolean: containsLowercase = False
    Dim v
    For Each v In Me.ToCharArray(str)
        If containsUppercase And containsLowercase Then
            IsMixedCase = True
            Exit Function
        ElseIf CharUtils.IsAsciiAlphaUpper(CStr(v)) Then
            containsUppercase = True
        ElseIf CharUtils.IsAsciiAlphaLower(CStr(v)) Then
            containsLowercase = True
        End If
    Next
    IsMixedCase = containsUppercase And containsLowercase
End Function

'******************************************************************************
'* [概  要] IsNoneBlank
'* [詳  細] 文字列のセットが空（""）、スペースのみでないかどうかをチェック｡
'*
'* [用  例]
'*   StringUtils.IsNoneBlank("", "bar") = False
'*   StringUtils.IsNoneBlank("bob", "") = False
'*   StringUtils.IsNoneBlank("  bob  ", Null) = False
'*   StringUtils.IsNoneBlank(" ", "bar") = False
'*   StringUtils.IsNoneBlank("foo", "bar") = True
'*
'* @param strs() 文字列（複数指定）
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsNoneBlank(ParamArray strs() As Variant) As Boolean
    Dim tmp: tmp = strs
    IsNoneBlank = Not Me.IsAnyBlank(tmp)
End Function

'******************************************************************************
'* [概  要] IsNoneEmpty
'* [詳  細] 文字列のセットが空（""）のみでないかどうかをチェック｡
'*
'* [用  例]
'*   StringUtils.IsNoneEmpty("", "bar") = False
'*   StringUtils.IsNoneEmpty("bob", "") = False
'*   StringUtils.IsNoneEmpty("  bob  ", Null) = False
'*   StringUtils.IsNoneEmpty(" ", "bar") = True
'*   StringUtils.IsNoneEmpty("foo", "bar") = True
'*
'* @param strs() 文字列（複数指定）
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsNoneEmpty(ParamArray strs() As Variant) As Boolean
    Dim tmp: tmp = strs
    IsNoneEmpty = Not Me.IsAnyEmpty(tmp)
End Function

'******************************************************************************
'* [概  要] IsNotBlank
'* [詳  細] 文字列が空（""）、スペースのみでないかどうかをチェック｡
'*
'* [用  例]
'*   StringUtils.IsNotBlank("") = False
'*   StringUtils.IsNotBlank(" ") = False
'*   StringUtils.IsNotBlank("bob") = True
'*   StringUtils.IsNotBlank("  bob  ") = True
'*
'* @param str 文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsNotBlank(ByVal str As String) As Boolean
    IsNotBlank = Not Me.IsBlank(str)
End Function

'******************************************************************************
'* [概  要] IsNotEmpty
'* [詳  細] 文字列が空（""）のみでないかどうかをチェック｡
'*
'* [用  例]
'*   StringUtils.IsNotEmpty("") = False
'*   StringUtils.IsNotEmpty(" ") = True
'*   StringUtils.IsNotEmpty("bob") = True
'*   StringUtils.IsNotEmpty("  bob  ") = True
'*
'* @param str 文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsNotEmpty(ByVal str As String) As Boolean
    IsNotEmpty = Not Me.IsEmpty(str)
End Function

'******************************************************************************
'* [概  要] IsNumeric
'* [詳  細] 文字列が数字かどうかをチェック｡
'*          小数は、Falseを返す。
'*
'* [用  例]
'*   StringUtils.IsNumeric("") = False
'*   StringUtils.IsNumeric("  ") = False
'*   StringUtils.IsNumeric("123") = True
'*   StringUtils.IsNumeric("12 3") = False
'*   StringUtils.IsNumeric("ab2c") = False
'*   StringUtils.IsNumeric("12-3") = False
'*   StringUtils.IsNumeric("12.3") = False
'*   StringUtils.IsNumeric("-123") = False
'*   StringUtils.IsNumeric("+123") = False
'*
'* @param str 文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsNumeric(ByVal str As String) As Boolean
    If Me.IsEmpty(str) Then
        IsNumeric = False
        Exit Function
    End If
    Dim v
    For Each v In Me.ToCharArray(str)
        If Not CharUtils.IsAsciiNumeric(CStr(v)) Then
            IsNumeric = False
            Exit Function
        End If
    Next
    IsNumeric = True
End Function

'******************************************************************************
'* [概  要] IsNumericSpace
'* [詳  細] 文字列が数字またはスペースかどうかをチェック｡
'*
'* [用  例]
'*   StringUtils.IsNumericSpace("") = True
'*   StringUtils.IsNumericSpace("  ") = True
'*   StringUtils.IsNumericSpace("123") = True
'*   StringUtils.IsNumericSpace("12 3") = True
'*   StringUtils.IsNumericSpace("ab2c") = False
'*   StringUtils.IsNumericSpace("12-3") = False
'*   StringUtils.IsNumericSpace("12.3") = False
'*
'* @param str 文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsNumericSpace(ByVal str As String) As Boolean
    If Me.IsEmpty(str) Then
        IsNumericSpace = False
        Exit Function
    End If
    Dim v
    For Each v In Me.ToCharArray(str)
        If Not CharUtils.IsAsciiNumeric(CStr(v)) And CStr(v) <> " " Then
            IsNumericSpace = False
            Exit Function
        End If
    Next
    IsNumericSpace = True
End Function

'******************************************************************************
'* [概  要] LastIndexOf
'* [詳  細] 文字列内の最後のインデックスを検索｡
'*
'* [用  例]
'*   StringUtils.LastIndexOf("", "") = 0
'*   StringUtils.LastIndexOf("aabaabaa", "a") = 7
'*   StringUtils.LastIndexOf("aabaabaa", "b") = 5
'*   StringUtils.LastIndexOf("aabaabaa", "ab") = 4
'*   StringUtils.LastIndexOf("aabaabaa", "") = 8
'*   StringUtils.LastIndexOf("aabaabaa", "a", 8) = 7
'*   StringUtils.LastIndexOf("aabaabaa", "b", 8) = 5
'*   StringUtils.LastIndexOf("aabaabaa", "ab", 8) = 4
'*   StringUtils.LastIndexOf("aabaabaa", "b", 9) = 5
'*   StringUtils.LastIndexOf("aabaabaa", "b", -1) = -1
'*   StringUtils.LastIndexOf("aabaabaa", "a", 0) = 0
'*   StringUtils.LastIndexOf("aabaabaa", "b", 0) = -1
'*   StringUtils.LastIndexOf("aabaabaa", "b", 1) = -1
'*   StringUtils.LastIndexOf("aabaabaa", "b", 2) = 2
'*   StringUtils.LastIndexOf("aabaabaa", "ba", 2) = 2
'*
'* @param str 文字列
'* @param searchStr 検索文字列
'* @param startPos 開始位置。負の値はゼロとして扱われる。
'* @return Long 検索文字列の最後のインデックス。一致しない場合、-1
'*
'******************************************************************************
Public Function LastIndexOf(ByVal str As String, ByVal searchStr As String, Optional ByVal startPos) As Long
    If Me.IsEmpty(str) Then
        If Not Me.IsEmpty(searchStr) Then
            LastIndexOf = INDEX_NOT_FOUND
        Else
            LastIndexOf = 0
        End If
        Exit Function
    End If
    If IsMissing(startPos) Or startPos >= Len(str) Then
        LastIndexOf = InStrRev(str, searchStr) - 1
        Exit Function
    ElseIf startPos < 0 Then
        startPos = 0

    End If
    LastIndexOf = InStrRev(str, searchStr, startPos + 1) - 1
End Function

'******************************************************************************
'* [概  要] LastIndexOfIgnoreCase
'* [詳  細] 文字列内の最後のインデックスを検索｡
'*          接尾辞の存在チェックは大文字小文字を考慮しない。
'*
'* [用  例]
'*   StringUtils.LastIndexOfIgnoreCase("aabaabaa", "A") = 7
'*   StringUtils.LastIndexOfIgnoreCase("aabaabaa", "B") = 5
'*   StringUtils.LastIndexOfIgnoreCase("aabaabaa", "AB") = 4
'*   StringUtils.LastIndexOfIgnoreCase("aabaabaa", "A", 8) = 7
'*   StringUtils.LastIndexOfIgnoreCase("aabaabaa", "B", 8) = 5
'*   StringUtils.LastIndexOfIgnoreCase("aabaabaa", "AB", 8) = 4
'*   StringUtils.LastIndexOfIgnoreCase("aabaabaa", "B", 9) = 5
'*   StringUtils.LastIndexOfIgnoreCase("aabaabaa", "B", -1) = -1
'*   StringUtils.LastIndexOfIgnoreCase("aabaabaa", "A", 0) = 0
'*   StringUtils.LastIndexOfIgnoreCase("aabaabaa", "B", 0) = -1
'*
'* @param str 文字列
'* @param searchStr 検索文字列
'* @param startPos 開始位置。負の値はゼロとして扱われる。
'* @return Long 検索文字列の最後のインデックス。一致しない場合、-1
'*
'******************************************************************************
Public Function LastIndexOfIgnoreCase(ByVal str As String, ByVal searchStr As String, Optional ByVal startPos) As Long
    Dim strIc As String: strIc = Me.LowerCase(str)
    Dim searchStrIc As String: searchStrIc = Me.LowerCase(searchStr)
    If IsMissing(startPos) Then
        LastIndexOfIgnoreCase = Me.LastIndexOf(strIc, searchStrIc)
        Exit Function
    End If
    LastIndexOfIgnoreCase = Me.LastIndexOf(strIc, searchStrIc, startPos)
End Function


'******************************************************************************
'* [概  要] LeftPad
'* [詳  細] 文字列を指定したパディング文字列で左詰めする｡
'*          デフォルトのパディング文字列は半角スペース。
'*
'* [用  例]
'*   StringUtils.LeftPad("", 3) = "   "
'*   StringUtils.LeftPad("bat", 3) = "bat"
'*   StringUtils.LeftPad("bat", 5) = "  bat"
'*   StringUtils.LeftPad("bat", 1) = "bat"
'*   StringUtils.LeftPad("bat", -1) = "bat"
'*   StringUtils.LeftPad("", 3, "z") = "zzz"
'*   StringUtils.LeftPad("bat", 3, "z") = "bat"
'*   StringUtils.LeftPad("bat", 5, "z") = "zzbat"
'*   StringUtils.LeftPad("bat", 1, "z") = "bat"
'*   StringUtils.LeftPad("bat", -1, "z") = "bat"
'*   StringUtils.LeftPad("bat", 3, "yz") = "bat"
'*   StringUtils.LeftPad("bat", 5, "yz") = "yzbat"
'*   StringUtils.LeftPad("bat", 8, "yz") = "yzyzybat"
'*   StringUtils.LeftPad("bat", 1, "yz") = "bat"
'*   StringUtils.LeftPad("bat", -1, "yz") = "bat"
'*
'* @param str 文字列
'* @param size パディングするサイズ
'* @param padStr パディング文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function LeftPad(ByVal str As String, ByVal size As Long, Optional ByVal padStr As String = " ") As String
    Dim padLen As Long: padLen = Len(padStr)
    Dim strLen As Long: strLen = Len(str)
    Dim pads As Long: pads = size - strLen
    If pads <= 0 Then
        LeftPad = str '  returns original String when possible
        Exit Function
    End If
    If padLen = 1 And pads <= PAD_LIMIT Then
        LeftPad = Me.Repeat(padStr, pads) & str
        Exit Function
    End If
    If pads = padLen Then
        LeftPad = padStr & str
        Exit Function
    ElseIf pads < padLen Then
        LeftPad = Me.Substring(str, 0, pads) & str
        Exit Function
    End If
    Dim padding: ReDim padding(0 To pads - 1)
    Dim padChars: padChars = Me.ToCharArray(padStr)
    Dim i As Long
    For i = 0 To pads - 1
        padding(i) = padChars(i Mod padLen)
    Next
    LeftPad = VBA.Join(padding) & str
End Function


'******************************************************************************
'* [概  要] LowerCase
'* [詳  細] 文字列を小文字に変換する｡
'*
'* [用  例]
'*   StringUtils.LowerCase("") = ""
'*   StringUtils.LowerCase("aBc") = "abc"
'*
'* @param str 文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function LowerCase(ByVal str As String) As String
    LowerCase = StrConv(str, vbLowerCase)
End Function


'******************************************************************************
'* [概  要] Narrow
'* [詳  細] 全角文字を半角文字に変換する｡
'*
'* [用  例]
'*   StringUtils.Proper("Ｈａｐｐｙ　Ｂｉｒｔｈｄａｙ！") = "happy birthday!"
'*
'* @param str 文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function Narrow(ByVal str As String) As String
    Narrow = StrConv(str, vbNarrow)
End Function

'******************************************************************************
'* [概  要] PrependIfMissing
'* [詳  細] 文字列がどの接頭辞（プレフィックス）でも始まっていない場合は、
'*          文字列の先頭に接頭辞を付加する｡
'*
'* [用  例]
'*   StringUtils.PrependIfMissing("", "xyz") = "xyz"
'*   StringUtils.PrependIfMissing("abc", "xyz") = "xyzabc"
'*   StringUtils.PrependIfMissing("xyzabc", "xyz") = "xyzabc"
'*   StringUtils.PrependIfMissing("XYZabc", "xyz") = "xyzXYZabc"
'*   StringUtils.PrependIfMissing("abc", "xyz", "") = "abc"
'*   StringUtils.PrependIfMissing("abc", "xyz", "mno") = "xyzabc"
'*   StringUtils.PrependIfMissing("xyzabc", "xyz", "mno") = "xyzabc"
'*   StringUtils.PrependIfMissing("mnoabc", "xyz", "mno") = "mnoabc"
'*   StringUtils.PrependIfMissing("XYZabc", "xyz", "mno") = "xyzXYZabc"
'*   StringUtils.PrependIfMissing("MNOabc", "xyz", "mno") = "xyzMNOabc"
'*
'* @param str 文字列
'* @param prefix  文字列の先頭に追加する接頭辞
'* @param prefixes() 追加のプレフィックス
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function PrependIfMissing(ByVal str As String, ByVal prefix As String, ParamArray prefixes() As Variant) As String
    If Me.IsEmpty(prefix) Or Me.StartsWith(str, prefix) Then
        PrependIfMissing = str
        Exit Function
    End If
    If Not IsMissing(prefixes) Then
        Dim tmp: tmp = prefixes ' ParamArrayを別関数に渡すための対応
        Dim v
        For Each v In ArrayUtils.ToVArrayFlattened(tmp)
            If Me.StartsWith(str, CStr(v)) Then
                PrependIfMissing = str
                Exit Function
            End If
        Next
    End If
    PrependIfMissing = prefix & str
End Function


'******************************************************************************
'* [概  要] PrependIfMissingIgnoreCase
'* [詳  細] 文字列がどの接頭辞（プレフィックス）でも始まっていない場合は、
'*          文字列の先頭に接頭辞を付加する｡
'*          大文字小文字を考慮しない。
'*
'* [用  例]
'*   StringUtils.PrependIfMissingIgnoreCase("", "xyz") = "xyz"
'*   StringUtils.PrependIfMissingIgnoreCase("abc", "xyz") = "xyzabc"
'*   StringUtils.PrependIfMissingIgnoreCase("xyzabc", "xyz") = "xyzabc"
'*   StringUtils.PrependIfMissingIgnoreCase("XYZabc", "xyz") = "XYZabc"
'*   StringUtils.PrependIfMissingIgnoreCase("abc", "xyz", "") = "abc"
'*   StringUtils.PrependIfMissingIgnoreCase("abc", "xyz", "mno") = "xyzabc"
'*   StringUtils.PrependIfMissingIgnoreCase("xyzabc", "xyz", "mno") = "xyzabc"
'*   StringUtils.PrependIfMissingIgnoreCase("mnoabc", "xyz", "mno") = "mnoabc"
'*   StringUtils.PrependIfMissingIgnoreCase("XYZabc", "xyz", "mno") = "XYZabc"
'*   StringUtils.PrependIfMissingIgnoreCase("MNOabc", "xyz", "mno") = "MNOabc"
'*
'* @param str 文字列
'* @param prefix  文字列の先頭に追加する接頭辞
'* @param prefixes() 追加のプレフィックス
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function PrependIfMissingIgnoreCase(ByVal str As String, ByVal prefix As String, ParamArray prefixes() As Variant) As String
    If Me.IsEmpty(prefix) Or Me.StartsWithIgnoreCase(str, prefix) Then
        PrependIfMissingIgnoreCase = str
        Exit Function
    End If
    If Not IsMissing(prefixes) Then
        Dim tmp: tmp = prefixes ' ParamArrayを別関数に渡すための対応
        Dim v
        For Each v In ArrayUtils.ToVArrayFlattened(tmp)
            If Me.StartsWithIgnoreCase(str, CStr(v)) Then
                PrependIfMissingIgnoreCase = str
                Exit Function
            End If
        Next
    End If
    PrependIfMissingIgnoreCase = prefix & str
End Function


'******************************************************************************
'* [概  要] Proper
'* [詳  細] 先頭文字のみを大文字にする｡
'*
'* [用  例]
'*   StringUtils.Proper("happy birthday!") = "Happy Birthday!"
'*
'* @param str 文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function Proper(ByVal str As String) As String
    Proper = StrConv(str, vbProperCase)
End Function

'******************************************************************************
'* [概  要] Remove
'* [詳  細] ソース文字列内から指定した文字に一致する全ての文字を削除｡
'*
'* [用  例]
'*   StringUtils.remove("", *)         = ""
'*   StringUtils.remove(*, "")          = *
'*   StringUtils.Remove("queued", "ue") = "qd"
'*   StringUtils.Remove("queued", "zz") = "queued"
'*
'* @param str 文字列
'* @param removeStr 削除する文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function Remove(ByVal str As String, ByVal removeStr As String) As String
    If Me.IsEmpty(str) Or Me.IsEmpty(removeStr) Then
        Remove = str
        Exit Function
    End If
    Remove = VBA.Replace(str, removeStr, "")
End Function

'******************************************************************************
'* [概  要] RemoveEach
'* [詳  細] 検索文字列リストで指定した文字列に合致するすべての文字列を
'*          削除する｡
'*
'* @param text 文字列
'* @param searchList 検索文字列リスト（Variant配列）
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function RemoveEach(ByVal text As String, ByVal searchList As Variant) As String
    If Me.IsEmpty(text) Or Not VBA.IsArray(searchList) Then
        RemoveEach = text
        Exit Function
    End If
    Dim i As Long, j As Long
    For i = LBound(searchList) To UBound(searchList)
        text = Remove(text, searchList(i))
    Next
    RemoveEach = text
End Function

'******************************************************************************
'* [概  要] RemoveEnd
'* [詳  細] ソース文字列の末尾にある場合にのみ部分文字列を削除し、
'*          それ以外の場合はソース文字列を返す。
'*
'* [用  例]
'*   StringUtils.removeEnd("", *)        = ""
'*   StringUtils.RemoveEnd("www.domain.com", ".com.") = "www.domain.com"
'*   StringUtils.RemoveEnd("www.domain.com", ".com") = "www.domain"
'*   StringUtils.RemoveEnd("www.domain.com", "domain") = "www.domain.com"
'*   StringUtils.RemoveEnd("abc", "") = "abc"
'*
'* @param str 文字列
'* @param removeStr 削除する文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function RemoveEnd(ByVal str As String, ByVal removeStr As String) As String
    If Me.IsEmpty(str) Or Me.IsEmpty(removeStr) Then
        RemoveEnd = str
        Exit Function
    End If
    If Me.EndsWith(str, removeStr) Then
        RemoveEnd = Me.Substring(str, 0, Len(str) - Len(removeStr))
        Exit Function
    End If
    RemoveEnd = str
End Function


'******************************************************************************
'* [概  要] RemoveEndIgnoreCase
'* [詳  細] ソース文字列の末尾にある場合にのみ部分文字列を削除し、
'*          それ以外の場合はソース文字列を返す。
'*          大文字小文字を考慮しない。
'*
'* [用  例]
'*   StringUtils.removeEndIgnoreCase("", *)        = ""
'*   StringUtils.RemoveEndIgnoreCase("www.domain.com", ".com.") = "www.domain.com"
'*   StringUtils.RemoveEndIgnoreCase("www.domain.com", ".com") = "www.domain"
'*   StringUtils.RemoveEndIgnoreCase("www.domain.com", "domain") = "www.domain.com"
'*   StringUtils.RemoveEndIgnoreCase("abc", "") = "abc"
'*   StringUtils.RemoveEndIgnoreCase("www.domain.com", ".COM") = "www.domain"
'*   StringUtils.RemoveEndIgnoreCase("www.domain.COM", ".com") = "www.domain"
'*
'* @param str 文字列
'* @param removeStr 削除する文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function RemoveEndIgnoreCase(ByVal str As String, ByVal removeStr As String) As String
    If Me.IsEmpty(str) Or Me.IsEmpty(removeStr) Then
        RemoveEndIgnoreCase = str
        Exit Function
    End If
    If Me.EndsWithIgnoreCase(str, removeStr) Then
        RemoveEndIgnoreCase = Me.Substring(str, 0, Len(str) - Len(removeStr))
        Exit Function
    End If
    RemoveEndIgnoreCase = str

End Function

'******************************************************************************
'* [概  要] RemoveIgnoreCase
'* [詳  細] ソース文字列内から指定した文字に一致する全ての文字を削除｡
'*          大文字小文字を考慮しない。
'*
'* [用  例]
'*   StringUtils.removeIgnoreCase("", *)          = ""
'*   StringUtils.removeIgnoreCase(*, "")          = *
'*   StringUtils.RemoveIgnoreCase("queued", "ue") = "qd"
'*   StringUtils.RemoveIgnoreCase("queued", "zz") = "queued"
'*   StringUtils.RemoveIgnoreCase("quEUed", "UE") = "qd"
'*   StringUtils.RemoveIgnoreCase("queued", "zZ") = "queued"
'*
'* @param str 文字列
'* @param removeStr 削除する文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function RemoveIgnoreCase(ByVal str As String, ByVal removeStr As String) As String
    If Me.IsEmpty(str) Or Me.IsEmpty(removeStr) Then
        RemoveIgnoreCase = str
        Exit Function
    End If
    RemoveIgnoreCase = Me.ReplaceIgnoreCase(str, removeStr, "")
End Function

'******************************************************************************
'* [概  要] RemoveStart
'* [詳  細] ソース文字列の先頭にある場合にのみ部分文字列を削除し、
'*          それ以外の場合はソース文字列を返す｡
'*
'* [用  例]
'*   StringUtils.RemoveStart("", *)        = ""
'*   StringUtils.RemoveStart("www.domain.com", "www.") = "domain.com"
'*   StringUtils.RemoveStart("domain.com", "www.") = "domain.com"
'*   StringUtils.RemoveStart("www.domain.com", "domain") = "www.domain.com"
'*   StringUtils.RemoveStart("abc", "") = "abc"
'*
'* @param str 文字列
'* @param removeStr 削除する文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function RemoveStart(ByVal str As String, ByVal removeStr As String) As String
    If Me.IsEmpty(str) Or Me.IsEmpty(removeStr) Then
        RemoveStart = str
        Exit Function
    End If
    If Me.StartsWith(str, removeStr) Then
        RemoveStart = Me.Substring(str, Len(removeStr))
        Exit Function
    End If
    RemoveStart = str
End Function

'******************************************************************************
'* [概  要] RemoveStartIgnoreCase
'* [詳  細] ソース文字列の先頭にある場合にのみ部分文字列を削除し、
'*          それ以外の場合はソース文字列を返す｡
'*          大文字小文字を考慮しない。
'*
'* [用  例]
'*   StringUtils.RemoveStartIgnoreCase("", *)        = ""
'*   StringUtils.RemoveStartIgnoreCase("www.domain.com", "www.") = "domain.com"
'*   StringUtils.RemoveStartIgnoreCase("www.domain.com", "WWW.") = "domain.com"
'*   StringUtils.RemoveStartIgnoreCase("domain.com", "www.") = "domain.com"
'*   StringUtils.RemoveStartIgnoreCase("www.domain.com", "domain") = "www.domain.com"
'*   StringUtils.RemoveStartIgnoreCase("abc", "") = "abc"
'*
'* @param str 文字列
'* @param removeStr 削除する文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function RemoveStartIgnoreCase(ByVal str As String, ByVal removeStr As String) As String
    If Me.IsEmpty(str) Or Me.IsEmpty(removeStr) Then
        RemoveStartIgnoreCase = str
        Exit Function
    End If
    If Me.StartsWithIgnoreCase(str, removeStr) Then
        RemoveStartIgnoreCase = Me.Substring(str, Len(removeStr))
        Exit Function
    End If
    RemoveStartIgnoreCase = str

End Function


'******************************************************************************
'* [概  要] Repeat
'* [詳  細] 文字列を繰返し回数分繰返し、新しい文字列を形成する。
'*          挿入文字列を指定している場合、繰返し毎に毎回挿入される。
'*
'* [用  例]
'*   StringUtils.Repeat("", 0) = ""
'*   StringUtils.Repeat("", 2) = ""
'*   StringUtils.Repeat("a", 3) = "aaa"
'*   StringUtils.Repeat("ab", 2) = "abab"
'*   StringUtils.Repeat("a", -2) = ""
'*   StringUtils.Repeat("", 2, "") = ""
'*   StringUtils.Repeat("", 3, "x") = "xxx"
'*   StringUtils.Repeat("?", 3, ", ") = "?, ?, ?"
'*
'* @param str 繰返し文字列
'* @param repeatCount 繰返し回数
'* @param separator 挿入文字列
'* @return String 元の文字列が繰り返された新しい文字列
'*
'******************************************************************************
Public Function Repeat(ByVal str As String, ByVal repeatCount As Long, Optional ByVal separator As String) As String
    If repeatCount <= 0 Then
        Repeat = ""
        Exit Function
    End If
    If separator <> "" Then
        Repeat = Repeat(str & separator, repeatCount)
        Repeat = Me.RemoveEnd(Repeat, separator)
        Exit Function
    End If
    
    Dim inputLength: inputLength = Len(str)
    If repeatCount = 0 Or inputLength = 0 Then
        Repeat = ""
        Exit Function
    End If
    If inputLength = 1 And repeatCount <= PAD_LIMIT Then
        Repeat = String(repeatCount, str)
        Exit Function
    End If
    Dim outputLength: outputLength = inputLength * repeatCount
    Select Case inputLength
        Case 1
            Repeat = String(repeatCount, str)
        Case 2
            Dim ch0: ch0 = Me.CharAt(str, 0)
            Dim ch1: ch0 = Me.CharAt(str, 1)
            Dim output2: ReDim output2(outputLength - 1)
            Dim i As Long
            For i = repeatCount * 2 - 2 To 0 Step -2
                output2(i) = ch0
                output2(i + 1) = ch1
            Next
            Repeat = VBA.Join(output2)
        Case Else
            Dim buf As StringEx: Set buf = Core.Init(New StringEx, outputLength)
            Dim j As Long
            For j = 0 To repeatCount - 1
                Call buf.Append(str)
            Next
            Repeat = buf.ToString
    
    End Select
End Function

'******************************************************************************
'* [概  要] ReplaceEach
'* [詳  細] 検索文字列リストで指定した文字列に合致するすべての文字列を
'*          置換文字列リストの文字列で置き換える｡
'*          配列の長さが同じでない場合は置換しない。
'*            ※JavaApacheCommonsLangのStringUtilsとは異なり単純に繰返し
'*              置換を行う。
'*
'* [用  例]
'*    StringUtils.ReplaceEach("aba", Array("a"), Array(""))  = "b"
'*    StringUtils.ReplaceEach("aba", Array(""), Array("a"))  = "aba"
'*    StringUtils.ReplaceEach("abcde", Array("ab", "d"), Array("w", "t"))  = "wcte"
'*
'* @param text 文字列
'* @param searchList 検索文字列リスト（Variant配列）
'* @param replacementList 置換文字列リスト（Variant配列）
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function ReplaceEach(ByVal text As String, ByVal searchList As Variant, ByVal replacementList As Variant) As String
    If Me.IsEmpty(text) Or Not VBA.IsArray(searchList) Or Not VBA.IsArray(replacementList) Then
        ReplaceEach = text
        Exit Function
    End If
    If (UBound(searchList) <> UBound(replacementList)) Or (LBound(searchList) <> LBound(replacementList)) Then
        ReplaceEach = text
        Exit Function
    End If
    Dim i As Long, j As Long
    For i = LBound(searchList) To UBound(searchList)
        text = VBA.Replace(text, searchList(i), replacementList(i))
    Next
    ReplaceEach = text
End Function

'******************************************************************************
'* [概  要] ReplaceIgnoreCase
'* [詳  細] 指定した検索文字列に合致するすべての文字列を
'*          置換文字列で置き換える｡
'*          大文字小文字を考慮しない。
'*
'* [用  例]
'*   StringUtils.replaceIgnoreCase("", *, *)          = ""
'*   StringUtils.replaceIgnoreCase("any", "", *)      = "any"
'*   StringUtils.ReplaceIgnoreCase("abA", "A", "") = "b"
'*   StringUtils.ReplaceIgnoreCase("aba", "A", "z") = "zbz"
'*
'* @param str 文字列
'* @param searchString 検索文字列
'* @param replacement 置換文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function ReplaceIgnoreCase(ByVal str As String, ByVal searchString As String, ByVal replacement As String) As String
    If Me.IsEmpty(str) Or Me.IsEmpty(searchString) Then
        ReplaceIgnoreCase = str
        Exit Function
    End If
    
    Dim strIc As String: strIc = Me.LowerCase(str)
    Dim searchStringIc As String: searchStringIc = Me.LowerCase(searchString)
    
    Dim sb As StringEx: Set sb = New StringEx
    Dim pos As Long: pos = 1
    Dim nextpos As Long: nextpos = InStr(strIc, searchStringIc)
    Dim Count As Long: Count = 0
    
    Do Until nextpos = 0
        Call sb.Append(Mid(str, pos, nextpos - pos)) _
               .Append(replacement)
        
        pos = nextpos + Len(searchStringIc)
        nextpos = InStr(pos, strIc, searchStringIc)
        Count = Count + 1
    Loop
    If Count = 0 Then
        ReplaceIgnoreCase = str
        Exit Function
    Else
        If pos <= Len(strIc) Then
            Call sb.Append(Mid(str, pos, Len(strIc) - pos + 1))
        End If
    End If
    ReplaceIgnoreCase = sb.ToString
End Function

'******************************************************************************
'* [概  要] Reverse
'* [詳  細] 文字列を反転させる｡
'*
'* [用  例]
'*   StringUtils.Reverse("") = ""
'*   StringUtils.Reverse("bat") = "tab"
'*
'* @param str 文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function Reverse(ByVal str As String) As String
    Reverse = VBA.StrReverse(str)
End Function


'******************************************************************************
'* [概  要] RightPad
'* [詳  細] 文字列を指定したパディング文字列で右詰めする｡
'*          デフォルトのパディング文字列は半角スペース。
'*
'* [用  例]
'*   StringUtils.RightPad("", 3) = "   "
'*   StringUtils.RightPad("bat", 3) = "bat"
'*   StringUtils.RightPad("bat", 5) = "bat  "
'*   StringUtils.RightPad("bat", 1) = "bat"
'*   StringUtils.RightPad("bat", -1) = "bat"
'*   StringUtils.RightPad("", 3, "z") = "zzz"
'*   StringUtils.RightPad("bat", 3, "yz") = "bat"
'*   StringUtils.RightPad("bat", 5, "yz") = "batyz"
'*   StringUtils.RightPad("bat", 8, "yz") = "batyzyzy"
'*   StringUtils.RightPad("bat", 1, "yz") = "bat"
'*   StringUtils.RightPad("bat", -1, "yz") = "bat"
'*
'* @param str 文字列
'* @param size パディングするサイズ
'* @param padStr パディング文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function RightPad(ByVal str As String, ByVal size As Long, Optional ByVal padStr As String = " ") As String
    Dim padLen As Long: padLen = Len(padStr)
    Dim strLen As Long: strLen = Len(str)
    Dim pads As Long: pads = size - strLen
    If pads <= 0 Then
        RightPad = str '  returns original String when possible
        Exit Function
    End If
    If padLen = 1 And pads <= PAD_LIMIT Then
        RightPad = str & Me.Repeat(padStr, pads)
        Exit Function
    End If
    If pads = padLen Then
        RightPad = str & padStr
        Exit Function
    ElseIf pads < padLen Then
        RightPad = str & Me.Substring(str, 0, pads)
        Exit Function
    End If
    Dim padding: ReDim padding(0 To pads - 1)
    Dim padChars: padChars = Me.ToCharArray(padStr)
    Dim i As Long
    For i = 0 To pads - 1
        padding(i) = padChars(i Mod padLen)
    Next
    RightPad = str & VBA.Join(padding)
End Function


'******************************************************************************
'* [概  要] Rotate
'* [詳  細] 文字列を循環シフトさせる｡
'*          シフト > 0の場合、右循環シフト（例：ABCDEF => FABCDE）
'*          シフト < 0の場合、左循環シフト（例：ABCDEF => BCDEFA）
'*
'* [用  例]
'*   StringUtils.rotate("", *)          = ""
'*   StringUtils.Rotate("abcdefg", 0) = "abcdefg"
'*   StringUtils.Rotate("abcdefg", 2) = "fgabcde"
'*   StringUtils.Rotate("abcdefg", -2) = "cdefgab"
'*   StringUtils.Rotate("abcdefg", 7) = "abcdefg"
'*   StringUtils.Rotate("abcdefg", -7) = "abcdefg"
'*   StringUtils.Rotate("abcdefg", 9) = "fgabcde"
'*   StringUtils.Rotate("abcdefg", -9) = "cdefgab"
'*
'* @param str 文字列
'* @param shiftCount シフトする回数（正：右シフト、負：左シフト）
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function Rotate(ByVal str As String, ByVal shiftCount As Long) As String
    Dim strLen As Long: strLen = Len(str)
    If (shiftCount = 0) Or (strLen = 0) Or (shiftCount Mod strLen = 0) Then
        Rotate = str
        Exit Function
    End If
    Dim builder As StringEx: builder = Core.Init(New StringEx, "", strLen)
    Dim offset As Long: offset = -(shiftCount Mod strLen)
    Call builder.Append(Me.Substring(str, offset)) _
                .Append(Me.Substring(str, 0, offset))
    Rotate = builder.ToString
End Function

'******************************************************************************
'* [概  要] StartsWithAny
'* [詳  細] 指定した文字列のいずれかで始まるかチェック｡
'*
'* [用  例]
'*   StringUtils.StartsWithAny("abcxyz", Null, "xyz", "ABCX") = False
'*   StringUtils.StartsWithAny("ABCXYZ", Null, "xyz", "abc") = False
'*
'* @param str 文字列
'* @param searchStrings 開始文字列（複数指定）
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function StartsWithAny(ByVal str As String, ParamArray searchStrings() As Variant) As Boolean
    If Me.IsEmpty(str) Or VBA.IsMissing(searchStrings) Then
        StartsWithAny = False
        Exit Function
    End If
    Dim tmp: tmp = searchStrings ' ParamArrayを別関数に渡すための対応
    Dim v
    For Each v In ArrayUtils.ToVArrayFlattened(tmp)
        If Me.StartsWith(str, CStr(v)) Then
            StartsWithAny = True
            Exit Function
        End If
    Next
    StartsWithAny = False
End Function

'******************************************************************************
'* [概  要] StartsWithIgnoreCase
'* [詳  細] 指定した文字列で始まるかチェック｡
'*          接尾辞の存在チェックは大文字小文字を考慮しない。
'*
'* [用  例]
'*   StringUtils.StartsWithIgnoreCase("abcdef", "abc") = True
'*   StringUtils.StartsWithIgnoreCase("ABCDEF", "abc") = True
'*
'* @param str 文字列
'* @param prefix 開始文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function StartsWithIgnoreCase(ByVal str As String, ByVal prefix As String) As Boolean
    If Me.IsEmpty(str) Or Me.IsEmpty(prefix) Then
        StartsWithIgnoreCase = False
        Exit Function
    End If
    Dim strIc As String: strIc = Me.LowerCase(str)
    Dim prefixIc As String: prefixIc = Me.LowerCase(prefix)
    StartsWithIgnoreCase = Me.StartsWith(strIc, prefixIc)
End Function

'******************************************************************************
'* [概  要] Substring
'* [詳  細] 指定された文字列から部分文字列を取得する｡
'*          負の開始位置を使用して、文字列の末尾から開始/終了位置を指定可能。
'*          返される部分文字列は、開始位置の文字で始まり、終了位置の前で終わる。
'*          位置はゼロベースでカウントされる。つまり、文字列の先頭から開始する
'*          には、start = 0を使用する。
'*          開始する位置の指定が終了する位置と逆転していた場合、 ""が返される。
'*
'* [用  例]
'*   StringUtils.substring("", *)     = ""
'*   StringUtils.Substring("abc", 0) = "abc"
'*   StringUtils.Substring("abc", 2) = "c"
'*   StringUtils.Substring("abc", 4) = ""
'*   StringUtils.Substring("abc", -2) = "bc"
'*   StringUtils.Substring("abc", -4) = "abc"
'*   StringUtils.substring("", * ,  *)    = "";
'*   StringUtils.Substring("abc", 0, 2) = "ab"
'*   StringUtils.Substring("abc", 2, 0) = ""
'*   StringUtils.Substring("abc", 2, 4) = "c"
'*   StringUtils.Substring("abc", 4, 6) = ""
'*   StringUtils.Substring("abc", 2, 2) = ""
'*   StringUtils.Substring("abc", -2, -1) = "b"
'*   StringUtils.Substring("abc", -4, 2) = "ab"
'*
'* @param str 文字列
'* @param startPos 開始する位置
'* @param endPos 終了する位置
'* @return String 開始位置から終了位置までの部分文字列
'*
'******************************************************************************
Public Function Substring(ByVal str As String, ByVal startPos As Long, Optional ByVal endPos) As String
    If Me.IsEmpty(str) Then
        Substring = ""
        Exit Function
    End If
    If IsMissing(endPos) Then
        endPos = Len(str) - 1
    End If
    ' handle negatives
    If endPos < 0 Then
        endPos = Len(str) + endPos ' remember end is negative
    End If
    If startPos < 0 Then
        startPos = Len(str) + startPos ' remember start is negative
    End If
    If endPos > Len(str) Then
        endPos = Len(str)
    End If
    ' if start is greater than end, return ""
    If startPos > endPos Then
        Substring = ""
        Exit Function
    End If
    If startPos < 0 Then
        startPos = 0
    End If
    If endPos < 0 Then
        endPos = 0
    End If
    Substring = Mid(str, startPos + 1, CLng(endPos) - startPos + 1)
End Function


'******************************************************************************
'* [概  要] SubstringAfter
'* [詳  細] セパレータが最初に出現した後の部分文字列を取得する。
'*
'* [用  例]
'*   StringUtils.substringAfter("", *)        = ""
'*   StringUtils.SubstringAfter("abc", "a") = "bc"
'*   StringUtils.SubstringAfter("abcba", "b") = "cba"
'*   StringUtils.SubstringAfter("abc", "c") = ""
'*   StringUtils.SubstringAfter("abc", "d") = ""
'*   StringUtils.SubstringAfter("abc", "") = "abc"
'*
'* @param str 文字列
'* @param separator 検索する文字列
'* @return String セパレータが最初に出現した後の部分文字列
'*
'******************************************************************************
Public Function SubstringAfter(ByVal str As String, ByVal separator As String) As String
    If Me.IsEmpty(str) Then
        SubstringAfter = str
        Exit Function
    End If
    If Me.IsEmpty(separator) Then
        SubstringAfter = ""
        Exit Function
    End If
    Dim pos As Long: pos = Me.IndexOf(str, separator)
    If pos = INDEX_NOT_FOUND Then
        SubstringAfter = ""
        Exit Function
    End If
    SubstringAfter = Me.Substring(str, pos + Len(separator))
End Function


'******************************************************************************
'* [概  要] SubstringAfterLast
'* [詳  細] セパレータが最後に出現した後の部分文字列を取得する｡
'*
'* [用  例]
'*   StringUtils.substringAfterLast("", *)        = ""
'*   StringUtils.substringAfterLast(*, "")        = ""
'*   StringUtils.SubstringAfterLast("abc", "a") = "bc"
'*   StringUtils.SubstringAfterLast("abcba", "b") = "a"
'*   StringUtils.SubstringAfterLast("abc", "c") = ""
'*   StringUtils.SubstringAfterLast("a", "a") = ""
'*   StringUtils.SubstringAfterLast("a", "z") = ""
'*
'* @param str 文字列
'* @param separator 検索する文字列
'* @return String セパレータが最後に出現した後の部分文字列
'*
'******************************************************************************
Public Function SubstringAfterLast(ByVal str As String, ByVal separator As String) As String
    If Me.IsEmpty(str) Then
        SubstringAfterLast = str
        Exit Function
    End If
    If Me.IsEmpty(separator) Then
        SubstringAfterLast = ""
        Exit Function
    End If
    Dim pos As Long: pos = Me.LastIndexOf(str, separator)
    If pos = INDEX_NOT_FOUND Or pos = Len(str) - Len(separator) Then
        SubstringAfterLast = ""
        Exit Function
    End If
    SubstringAfterLast = Me.Substring(str, pos + Len(separator))
End Function


'******************************************************************************
'* [概  要] SubstringBefore
'* [詳  細]セパレータが最初に出現する前の部分文字列を取得する｡
'*
'* [用  例]
'*   StringUtils.substringBefore("", *)        = ""
'*   StringUtils.SubstringBefore("abc", "a") = ""
'*   StringUtils.SubstringBefore("abcba", "b") = "a"
'*   StringUtils.SubstringBefore("abc", "c") = "ab"
'*   StringUtils.SubstringBefore("abc", "d") = "abc"
'*   StringUtils.SubstringBefore("abc", "") = ""
'*   StringUtils.SubstringBefore("abc", Null) = "abc"
'*
'* @param str 文字列
'* @param separator 検索する文字列
'* @return String セパレータが最初に出現する前の部分文字列
'*
'******************************************************************************
Public Function SubstringBefore(ByVal str As String, ByVal separator As String) As String
    If Me.IsEmpty(str) Then
        SubstringBefore = str
        Exit Function
    End If
    If Me.IsEmpty(separator) Then
        SubstringBefore = ""
        Exit Function
    End If
    Dim pos As Long: pos = Me.LastIndexOf(str, separator)
    If pos = INDEX_NOT_FOUND Then
        SubstringBefore = ""
        Exit Function
    End If
    SubstringBefore = Me.Substring(str, 0, pos)
End Function

'******************************************************************************
'* [概  要] SubstringBeforeLast
'* [詳  細] セパレータが最後に出現する前の部分文字列を取得する｡
'*
'* [用  例]
'*   StringUtils.substringBeforeLast("", *)        = ""
'*   StringUtils.SubstringBeforeLast("abcba", "b") = "abc"
'*   StringUtils.SubstringBeforeLast("abc", "c") = "ab"
'*   StringUtils.SubstringBeforeLast("a", "a") = ""
'*   StringUtils.SubstringBeforeLast("a", "z") = "a"
'*   StringUtils.SubstringBeforeLast("a", Null) = "a"
'*   StringUtils.SubstringBeforeLast("a", "") = "a"
'*
'* @param str 文字列
'* @param separator 検索する文字列
'* @return String セパレータが最後に出現する前の部分文字列
'*
'******************************************************************************
Public Function SubstringBeforeLast(ByVal str As String, ByVal separator As String) As String
    If Me.IsEmpty(str) Then
        SubstringBeforeLast = str
        Exit Function
    End If
    If Me.IsEmpty(separator) Then
        SubstringBeforeLast = ""
        Exit Function
    End If
    Dim pos As Long: pos = Me.LastIndexOf(str, separator)
    If pos = INDEX_NOT_FOUND Then
        SubstringBeforeLast = ""
        Exit Function
    End If
    SubstringBeforeLast = Me.Substring(str, 0, pos)
End Function

'******************************************************************************
'* [概  要] SubstringBetween
'* [詳  細] ２つの文字列の間にネストされている文字列を取得する｡
'*
'* [用  例]
'*   StringUtils.SubstringBetween("", "") = ""
'*   StringUtils.SubstringBetween("", "tag") = ""
'*   StringUtils.SubstringBetween("tagabctag", "") = ""
'*   StringUtils.SubstringBetween("tagabctag", "tag") = "abc"
'*   StringUtils.substringBetween("wx[b]yz", "[", "]") = "b"
'*   StringUtils.substringBetween("", *, *)          = ""
'*   StringUtils.substringBetween(*, "", *)          = ""
'*   StringUtils.SubstringBetween("tagabctag", "tag", "") = "abc"
'*   StringUtils.substringBetween("", "", "")          = ""
'*   StringUtils.substringBetween("", "", "]")         = ""
'*   StringUtils.substringBetween("", "[", "]")        = ""
'*   StringUtils.substringBetween("yabcz", "", "")     = ""
'*   StringUtils.substringBetween("yabcz", "y", "z")   = "abc"
'*   StringUtils.substringBetween("yabczyabcz", "y", "z")   = "abc"
'*
'* @param str 文字列
'* @param openStr 部分文字列の開始を識別する文字列
'* @param closeStr 部分文字列の終わりを識別する文字列
'* @return String 部分文字列
'*
'******************************************************************************
Public Function SubstringBetween(ByVal str As String, ByVal tagOrOpenStr As String, Optional ByVal closeStr As String = "") As String
    If IsMissing(closeStr) Then
        closeStr = tagOrOpenStr
    End If
    If Me.IsEmpty(str) Or Me.IsEmpty(tagOrOpenStr) Or Me.IsEmpty(closeStr) Then
        SubstringBetween = ""
        Exit Function
    End If

    Dim startPos As Long: startPos = Me.IndexOf(str, tagOrOpenStr)
    If startPos <> INDEX_NOT_FOUND Then
        Dim endPos As Long: endPos = Me.IndexOf(str, closeStr, startPos + Len(tagOrOpenStr))
        If endPos <> INDEX_NOT_FOUND Then
            SubstringBetween = Me.Substring(str, startPos + Len(tagOrOpenStr), endPos)
            Exit Function
        End If
    End If
    SubstringBetween = ""
End Function


'******************************************************************************
'* [概  要] ToEncodedString
'* [詳  細] 指定された文字エンコーディングを使用して、バイト配列を文字列に
'*          変換する｡
'*
'* @param bytes() 読込むバイト配列
'* @param charset 使用するエンコーディング
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function ToEncodedString(Bytes() As Byte, ByVal charSet As String) As String
    With IO.CreateADODBStream(adTypeBinary, charSet)
        .mode = adModeReadWrite
        .Open
        .Write Bytes
        .Position = 0
        .Type = adTypeText
        ToEncodedString = .ReadText
        .Close
    End With
End Function

'******************************************************************************
'* [概  要] Truncate
'* [詳  細] 文字列を切り捨てる｡
'*
'* [用  例]
'*   StringUtils.Truncate("", 4) = ""
'*   StringUtils.Truncate("abcdefg", 4) = "abcd"
'*   StringUtils.Truncate("abcdefg", 6) = "abcdef"
'*   StringUtils.Truncate("abcdefg", 7) = "abcdefg"
'*   StringUtils.Truncate("abcdefg", 8) = "abcdefg"
'*
'* @param str 文字列
'* @param maxWidth 結果文字列の最大長
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function Truncate(ByVal str As String, ByVal maxWidth As Long) As String
    If maxWidth < 0 Then
        Call Err.Raise(9999, "StringUtils#Truncate", "maxWith cannot be negative")
    End If
    If Me.IsEmpty(str) Then
        Truncate = str
        Exit Function
    End If

    Truncate = Left(str, maxWidth)
End Function

'******************************************************************************
'* [概  要] Uncapitalize
'* [詳  細] 文字列の最初の文字を小文字に変更する｡
'*
'* [用  例]
'*   StringUtils.Uncapitalize("") = ""
'*   StringUtils.Uncapitalize("cat") = "cat"
'*   StringUtils.Uncapitalize("Cat") = "cat"
'*   StringUtils.Uncapitalize("CAT") = "cAT"
'*
'* @param str 文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function Uncapitalize(ByVal str As String) As String
    If Me.IsEmpty(str) Then
        Uncapitalize = str
        Exit Function
    End If
    Uncapitalize = StrConv(Left(str, 1), vbLowerCase) & Mid(str, 2)
End Function

'******************************************************************************
'* [概  要] Unwrap
'* [詳  細] 指定された文字列をアンラップする｡
'*
'* [用  例]
'*   StringUtils.Unwrap("a", "a") = "a"
'*   StringUtils.Unwrap("aa", "a") = ""
'*   StringUtils.Unwrap("'abc'", "'") = "abc"
'*   StringUtils.Unwrap("""abc""", """") = "abc"
'*   StringUtils.Unwrap("AABabcBAA", "AA") = "BabcB"
'*   StringUtils.Unwrap("A", "#") = "A"
'*   StringUtils.Unwrap("#A", "#") = "#A"
'*   StringUtils.Unwrap("A#", "#") = "A#"
'*
'* @param str 文字列
'* @param wrapToken アンラップ文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function Unwrap(ByVal str As String, ByVal wrapToken As String) As String
    If Me.IsEmpty(str) Or Me.IsEmpty(wrapToken) Or Len(str) = 1 Then
        Unwrap = str
        Exit Function
    End If

    If Me.StartsWith(str, wrapToken) And Me.EndsWith(str, wrapToken) Then
        Dim startIndex As Integer: startIndex = Me.IndexOf(str, wrapToken)
        Dim endIndex As Integer:   endIndex = Me.LastIndexOf(str, wrapToken)
        Dim wrapLength As Integer: wrapLength = Len(wrapToken)
        
        If startIndex <> -1 And endIndex <> -1 Then
            Unwrap = Me.Substring(startIndex + wrapLength, endIndex)
            Exit Function
        End If
    End If
    Unwrap = str
End Function

'******************************************************************************
'* [概  要] UpperCase
'* [詳  細] 文字列を大文字に変換する｡
'*
'* [用  例]
'*   StringUtils.UpperCase("") = ""
'*   StringUtils.UpperCase("aBc") = "ABC"
'*
'* @param str 文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function UpperCase(ByVal str As String) As String
    UpperCase = StrConv(str, vbUpperCase)
End Function

'******************************************************************************
'* [概  要] Wide
'* [詳  細] 半角文字を全角文字に変換する｡
'*
'* [用  例]
'*   StringUtils.Proper("happy birthday!") = "Ｈａｐｐｙ　Ｂｉｒｔｈｄａｙ！"
'*
'* @param str 文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function Wide(ByVal str As String) As String
    Wide = StrConv(str, vbWide)
End Function

'******************************************************************************
'* [概  要] Wrap
'* [詳  細] 指定された文字列をラップする｡
'*
'* [用  例]
'*   StringUtils.Wrap("ab", "x") = "xabx"
'*   StringUtils.Wrap("ab", """") = """ab"""
'*   StringUtils.Wrap("""ab""", """") = """""ab"""""
'*   StringUtils.Wrap("ab", "'") = "'ab'"
'*   StringUtils.Wrap("'abcd'", "'") = "''abcd''"
'*   StringUtils.Wrap("""abcd""", "'") = "'""abcd""'"
'*   StringUtils.Wrap("'abcd'", """") = """'abcd'"""
'*
'* @param str 文字列
'* @param wrapWith ラップ文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function Wrap(ByVal str As String, ByVal wrapWith As String) As String
    If Me.IsEmpty(str) Or Me.IsEmpty(wrapWith) Then
        Wrap = str
        Exit Function
    End If
    Wrap = wrapWith & str & wrapWith
End Function

'******************************************************************************
'* [概  要] WrapIfMissing
'* [詳  細] 文字列が指定された文字列の先頭または末尾にない場合、
'*          文字列をラップする｡
'*
'* [用  例]
'*   StringUtils.wrapIfMissing("", *)           = ""
'*   StringUtils.WrapIfMissing("ab", Null) = "ab"
'*   StringUtils.WrapIfMissing("ab", "x") = "xabx"
'*   StringUtils.WrapIfMissing("ab", """") = """ab"""
'*   StringUtils.WrapIfMissing("""ab""", """") = """ab"""
'*   StringUtils.WrapIfMissing("ab", "'") = "'ab'"
'*   StringUtils.WrapIfMissing("'abcd'", "'") = "'abcd'"
'*   StringUtils.WrapIfMissing("""abcd""", "'") = "'""abcd""'"
'*   StringUtils.WrapIfMissing("'abcd'", """") = """'abcd'"""
'*   StringUtils.WrapIfMissing("/", "/") = "/"
'*   StringUtils.WrapIfMissing("a/b/c", "/") = "/a/b/c/"
'*   StringUtils.WrapIfMissing("/a/b/c", "/") = "/a/b/c/"
'*   StringUtils.WrapIfMissing("a/b/c/", "/") = "/a/b/c/"
'*
'* @param str 文字列
'* @param wrapWith ラップ文字列
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function WrapIfMissing(ByVal str As String, ByVal wrapWith As String) As String
    If Me.IsEmpty(str) Or Me.IsEmpty(wrapWith) Then
        WrapIfMissing = str
        Exit Function
    End If

    Dim wrapStart As Boolean: wrapStart = Not Me.StartsWith(str, wrapWith)
    Dim wrapEnd As Boolean:   wrapEnd = Not Me.EndsWith(str, wrapWith)
    If Not wrapStart And Not wrapEnd Then
        WrapIfMissing = str
        Exit Function
    End If
    
    Dim builder As StringEx
    builder = Core.Init(New StringEx, "", Len(str) + Len(wrapWith) + Len(wrapWith))
    If wrapStart Then
        builder.Append wrapWith
    End If
    builder.Append str
    If wrapEnd Then
        builder.Append wrapWith
    End If
    WrapIfMissing = builder.ToString
End Function


'******************************************************************************
'* [概  要] ToCharArray
'* [詳  細] 文字列を１文字ずつの配列に変換する｡
'*
'* [用  例]
'*   StringUtils.ToCharArray("abc") = ["a", "b", "c"]
'*
'* [依  存]
'*   Ariawase.ArrayEx
'*
'* @param str 文字列
'* @return Variant 変換後文字配列
'*
'******************************************************************************
Public Function ToCharArray(ByVal str As String) As Variant
    Dim arr As ArrayEx: Set arr = New ArrayEx
    If Me.IsEmpty(str) Then
        ToCharArray = arr.ToArray
        Exit Function
    End If
    Dim i As Long
    For i = 1 To Len(str)
        arr.AddVal Mid(str, i, 1)
    Next
    ToCharArray = arr.ToArray
End Function

'******************************************************************************
'* [概  要] ToCharArrayReverse
'* [詳  細] 文字列を並び順を逆にして１文字ずつの配列に変換する｡
'*
'* [用  例]
'*   StringUtils.ToCharArrayReverse("abc") = ["c", "b", "a"]
'*
'* [依  存]
'*   Ariawase.ArrayEx
'*
'* @param str 文字列
'* @return Variant 変換後文字配列
'*
'******************************************************************************
Public Function ToCharArrayReverse(ByVal str As String) As Variant
    Dim arr As ArrayEx: Set arr = New ArrayEx
    If Me.IsEmpty(str) Then
        ToCharArrayReverse = arr.ToArray
        Exit Function
    End If
    Dim i As Long
    For i = Len(str) To 1 Step -1
        arr.AddVal Mid(str, i, 1)
    Next
    ToCharArrayReverse = arr.ToArray
End Function

'******************************************************************************
'* [概  要] CharAt
'* [詳  細] 文字列から指定したインデックス位置の文字を取得する｡
'*
'* [用  例]
'*   StringUtils.CharAt("abc", 0) = "a"
'*
'* @param str 文字列
'* @param index インデックス
'* @return String 文字
'*
'******************************************************************************
Public Function CharAt(ByVal str As String, Optional ByVal Index As Long = 0) As String
    CharAt = Mid(str, Index + 1, 1)
End Function

'*******************************************************************************
'* [概  要] V2Str
'* [詳  細] Variant型データを文字列に変換する。
'*
'* @param v Variant型データ
'* @return String 変換後文字列
'*
'*******************************************************************************
Public Function V2Str(ByVal v As Variant) As String
    If VBA.IsEmpty(v) Or VBA.IsNull(v) Then
        V2Str = ""
    Else
        V2Str = CStr(v)
    End If
End Function

'*******************************************************************************
'* [概  要] 左ゼロ埋め処理
'* [詳  細] 指定した桁数になる範囲で左ゼロ埋め編集を行う。
'*
'* @param str 編集対象文字列
'* @param size 桁数
'* @return String 左ゼロ埋めした値
'*
'*******************************************************************************
Public Function LeftPad0(ByVal str As String, ByVal size As Long) As String
    LeftPad0 = Me.LeftPad(str, size, "0")
End Function

'*******************************************************************************
'* [概  要] 右ゼロ埋め処理
'* [詳  細] 指定した桁数になる範囲で右ゼロ埋め編集を行う。
'*
'* @param str 編集対象文字列
'* @param size 桁数
'* @return String 右ゼロ埋めした値
'*
'*******************************************************************************
Public Function RightPad0(ByVal str As String, ByVal size As Long) As String
    RightPad0 = Me.RightPad(str, size, "0")
End Function

'*******************************************************************************
'* [概  要] シングルクォーテーションエスケープ処理
'* [詳  細] SQLに埋め込む文字列にシングルクォーテーションがある場合にエスケープ
'*          を行うための処理。
'*
'* @param val エスケープ対象文字列
'* @return String エスケープ後文字列
'*
'*******************************************************************************
Public Function EscapeSqlQuot(ByVal val As String) As String
    '「'」が2つ連続している場合はエスケープ済みと見做して対応
    EscapeSqlQuot = VBA.Replace(VBA.Replace(val, "''", "'"), "'", "''")
End Function


'*******************************************************************************
'* [概  要] SplitByLength
'* [詳  細] 指定した文字数で文字列を分割する。
'*
'* @param str Variant型データ
'* @param cutLength 文字数
'* @return Variant Variant文字配列
'*
'*******************************************************************************
Public Function SplitByLength(ByVal str As String, ByVal cutLength) As Variant
    If cutLength <= 0 Then
        SplitByLength = ArrayUtils.EmptyVariantArray
        Exit Function
    End If
    If Me.IsEmpty(str) Or (Len(str) Mod cutLength) <> 0 Then
        SplitByLength = ArrayUtils.EmptyVariantArray
        Exit Function
    End If
    Dim i As Long
    Dim arr As ArrayEx: Set arr = New ArrayEx
    For i = 1 To Len(str) / cutLength
        Call arr.AddVal(Mid(str, 1 + ((i - 1) * cutLength), cutLength))
    Next
    SplitByLength = arr.ToArray
End Function

'*******************************************************************************
'* [概  要] EditTextByFunc
'* [詳  細] FuncオブジェクトまたはFunction名を指定して、テキストを編集する。
'*
'* @param txt Variant型データ
'* @param fun FuncオブジェクトまたはFunction名
'* @return String 結果文字列
'*
'*******************************************************************************
Public Function EditTextByFunc(ByVal txt As String, ByRef fun As Variant) As String
    If ValidateUtils.IsString(fun) Then
        If fun <> "" Then
            EditTextByFunc = Application.Run(fun, txt)
            Exit Function
        End If
    ElseIf ValidateUtils.IsFunc(fun) Then
        EditTextByFunc = fun.Apply(txt)
        Exit Function
    End If
    EditTextByFunc = txt
End Function

'******************************************************************************
'* [概  要] CSV行分割処理。
'* [詳  細] CSV文字列をカンマごとに分割してVariant配列に格納し、返却する。
'*          項目両端の「"」を削除する。
'*
'* @param csvLine CSV行
'* @param quotExists ダブルクォート存在有無（True：あり）。省略可能。
'* @param delimiter デリミタ。省略可能。
'* @return 分割後Variant配列
'******************************************************************************
Public Function SplitCsvRow(csvLine As String, Optional ByVal quotExists As Boolean = False, Optional delimiter As String = ",") As Variant
    If quotExists Then
        Dim lQuoteCount As Long, lItemCount As Long, i As Long, tmpTxt As String, v()
        For i = 1 To Len(csvLine)
            Select Case Mid(csvLine, i, 1)
                Case delimiter
                    If lQuoteCount Mod 2 = 0 Then
                        lItemCount = lItemCount + 1
                    End If
                Case """"
                    lQuoteCount = lQuoteCount + 1
            End Select
        Next
        ReDim v(lItemCount) As Variant
        lQuoteCount = 0: lItemCount = 0
        For i = 1 To Len(csvLine)
            Select Case Mid(csvLine, i, 1)
                Case delimiter
                    If lQuoteCount Mod 2 = 0 Then
                        v(lItemCount) = EditIncludeQuotItem(tmpTxt)
                        tmpTxt = ""
                        lItemCount = lItemCount + 1
                    Else
                        tmpTxt = tmpTxt & Mid(csvLine, i, 1)
                    End If
                Case """"
                    lQuoteCount = lQuoteCount + 1
                    tmpTxt = tmpTxt & Mid(csvLine, i, 1)
                Case Else
                    tmpTxt = tmpTxt & Mid(csvLine, i, 1)
            End Select
        Next
        If tmpTxt <> "" Then
            v(lItemCount) = EditIncludeQuotItem(tmpTxt)
        End If
        SplitCsvRow = v
    Else
        SplitCsvRow = Split(csvLine, delimiter)
    End If
End Function

'******************************************************************************
'* [概  要] ダブルクォート入り項目編集処理。
'* [詳  細] エスケープされているダブルクォート「""」を「"」に変換し、
'*          項目両端の「"」を削除する。
'*
'* @param text ダブルクォート入り項目
'* @return 編集後項目
'******************************************************************************
Private Function EditIncludeQuotItem(ByVal text As String) As String
    text = Replace(text, """""", """")
    If Left(text, 1) = """" And Right(text, 1) = """" Then
        If Len(text) <= 2 Then
            EditIncludeQuotItem = ""
        Else
            EditIncludeQuotItem = Mid(text, 2, Len(text) - 2)
            
        End If
    End If
End Function

'******************************************************************************
'* [概  要] RemoveLineSeparator
'* [詳  細] 改行コード除去処理。
'*
'* @param val 文字列
'* @return 編集後文字列
'******************************************************************************
Public Function RemoveLineSeparator(ByVal val As String) As String
    RemoveLineSeparator = RemoveEach(val, Array(vbLf, vbCr))
End Function

'******************************************************************************
'* [概  要] EscapeDoubleQuote
'* [詳  細] ダブルクォーテーションエスケープ処理。
'*
'* @param val 文字列
'* @return 編集後文字列
'******************************************************************************
Public Function EscapeDoubleQuote(ByVal val As String) As String
    EscapeDoubleQuote = VBA.Replace(val, """", """""")
End Function
