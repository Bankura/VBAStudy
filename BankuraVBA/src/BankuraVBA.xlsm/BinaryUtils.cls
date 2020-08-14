VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BinaryUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] バイナリ操作ユーティリティクラス
'* [詳  細] バイナリ操作に関するユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'* [参  考]
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

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mWinApi As WinAPI

'******************************************************************************
'* プロパティ定義
'******************************************************************************

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mWinApi = Base.GetWinAPI()
End Sub
Private Sub Class_Terminate()
    Set mWinApi = Nothing
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] CopyMemory2ByteArray
'* [詳  細] 指定したメモリアドレスから指定バイトサイズ分のバイト情報を取得
'*          し、バイト配列として返却する。
'*
'* @param addr アドレス
'*             （VarPtr, StrPtr, ObjPtrの値をベースに指定する想定）
'* @param size 取得するバイトサイズ
'* @return Byte() メモリから取得したバイト情報
'******************************************************************************
Public Function CopyMemory2ByteArray(ByVal addr As LongPtr, ByVal size As Long) As Byte()
    If addr <= 0 Or size <= 0 Then
        CopyMemory2ByteArray = ArrayUtils.EmptyByteArray
        Exit Function
    End If
  
    ReDim bData(1 To size) As Byte
    Call mWinApi.CopyMemory(VarPtr(bData(1)), addr, size)
    CopyMemory2ByteArray = bData
End Function

'******************************************************************************
'* [概  要] ByteArray2HexStrings
'* [詳  細] バイト配列を16進文字列に変換する。
'*
'* @param bData() バイト配列
'* @param delimiter 区切り文字（デフォルト：半角スペース）
'* @return String 16進文字列
'******************************************************************************
Public Function ByteArray2HexStrings(bData() As Byte, Optional ByVal delimiter As String = " ") As String
    Dim size As Long
    size = ArrayUtils.GetLength(bData)
    If size <= 0 Then
        ByteArray2HexStrings = ""
        Exit Function
    End If
    
    Dim i As Long, hVal As String
    Dim pStr As StringEx: Set pStr = New StringEx
    
    For i = LBound(bData) To UBound(bData)
        hVal = Byte2HexString(bData(i))
        If i > LBound(bData) Then
            Call pStr.Append(delimiter)
        End If
        Call pStr.Append(hVal)
    Next i
    ByteArray2HexStrings = pStr.ToString
End Function

'******************************************************************************
'* [概  要] HexStrings2ByteArray
'* [詳  細] 16進文字列をバイト配列を変換する。
'*
'* @param hstr 16進文字列
'* @param delimiter 区切り文字（デフォルト：半角スペース）
'* @return String 16進文字列
'******************************************************************************
Public Function HexStrings2ByteArray(ByVal hstr As String, Optional ByVal delimiter As String = " ") As Byte()
    If hstr = "" Then
        HexStrings2ByteArray = ArrayUtils.EmptyByteArray
        Exit Function
    End If
    
    Dim varr
    If delimiter = "" Then
        If Len(hstr) Mod 2 <> 0 Then
            HexStrings2ByteArray = ArrayUtils.EmptyByteArray
            Exit Function
        End If
        varr = StringUtils.SplitByLength(hstr, 2)
    Else
        varr = Split(hstr, delimiter)
    End If
    
    Dim bData() As Byte, i As Long
    ReDim bData(LBound(varr) To UBound(varr)) As Byte
    
    For i = LBound(varr) To UBound(varr)
        bData(i) = HexString2Byte(varr(i))
    Next i
    HexStrings2ByteArray = bData
End Function

'******************************************************************************
'* [概  要] ByteArray2BinaryStrings
'* [詳  細] バイト配列をバイナリ文字列に変換する。
'*
'* @param bData() バイト配列
'* @return String バイナリ文字列
'******************************************************************************
Public Function ByteArray2BinaryStrings(bData() As Byte) As String
    Dim size As Long
    size = ArrayUtils.GetLength(bData)
    If size <= 0 Then
        ByteArray2BinaryStrings = ""
        Exit Function
    End If
    
    Dim i As Long, bVal As String, hVal As String
    For i = LBound(bData) To UBound(bData)
        hVal = Byte2HexString(bData(i))
        bVal = Hex2Bin(hVal) & bVal
    Next i
    ByteArray2BinaryStrings = bVal
End Function

'******************************************************************************
'* [概  要] ByteArray2LongPtr
'* [詳  細] バイト配列を整数値に変換する。
'*
'* @param bData() バイト配列
'* @return LongPtr 整数値
'******************************************************************************
Public Function ByteArray2LongPtr(bData() As Byte) As LongPtr
    Dim size As Long
    size = ArrayUtils.GetLength(bData)
    If size <= 0 Or 8 < size Then
        ByteArray2LongPtr = 0
        Exit Function
    End If
    
    Dim i As Long, dVal As LongPtr, hVal As String
    For i = LBound(bData) To UBound(bData)
        hVal = Byte2HexString(bData(i))
        dVal = dVal + Hex2Dec(hVal, i, LBound(bData))
    Next i
    
    ByteArray2LongPtr = dVal
End Function

'******************************************************************************
'* [概  要] Hex2dec
'* [詳  細] 1バイト分の16進文字列をバイト整数値に変換する。
'*          複数バイト分の16進文字列の位置を指定することで重みづけが可能。
'*
'* @param hVal 1バイト分の16進文字列
'* @param pos 位置
'* @param startPos 開始位置
'* @return LongPtr バイト整数値
'******************************************************************************
Public Function Hex2Dec(ByVal hVal As String, Optional ByVal pos As LongPtr = 0, Optional ByVal startPos As LongPtr = 0) As LongPtr
    Hex2Dec = HexString2LongPtr(hVal) * (256 ^ (pos - startPos))
End Function

'******************************************************************************
'* [概  要] Hex2bin
'* [詳  細] 1バイト分の16進文字列をバイナリ文字列に変換する。
'*
'* @param hVal 1バイト分の16進文字列
'* @return String バイナリ文字列
'******************************************************************************
Public Function Hex2Bin(ByVal hVal As String) As String
    Dim dVal As LongPtr
    dVal = HexString2LongPtr(hVal)
    Dim power2 As LongPtr: power2 = 1
    Do Until dVal < power2
        If (dVal And power2) <> 0 Then
            Hex2Bin = "1" & Hex2Bin
        Else
            Hex2Bin = "0" & Hex2Bin
        End If
        power2 = power2 * 2
    Loop
    Hex2Bin = StringUtils.LeftPad0(Hex2Bin, 8)
End Function

'******************************************************************************
'* [概  要] Bin2Hex
'* [詳  細] バイナリ文字列を16進文字列に変換する。
'*
'* @param bVal バイナリ文字列
'* @return LongPtr 16進文字列
'******************************************************************************
Public Function Bin2Hex(ByVal bVal As String) As String
    If Not IsBinaryString(bVal) Then
        Bin2Hex = ""
        Exit Function
    End If
    
    Dim str As StringEx: Set str = New StringEx
    Dim v, tmp As Long: tmp = 0
    For Each v In StringUtils.SplitByLength(bVal, 4)
        Dim i As Long
        tmp = 0
        For i = Len(v) To 1 Step -1
            If Mid(v, Len(v) - i + 1, 1) = "1" Then
               tmp = tmp + (2 ^ (i - 1))
            End If
        Next
        str.Append Hex$(tmp)
    Next

    Bin2Hex = str.ToString
End Function

'******************************************************************************
'* [概  要] Byte2HexString
'* [詳  細] 1バイトを16進文字列に変換する。
'*
'* @param b 1バイト
'* @param prependSuffix 先頭に"&H"を付与するか（デフォルト：False）
'* @return String 16進文字列
'******************************************************************************
Public Function Byte2HexString(ByVal b As Byte, Optional ByVal prependSuffix As Boolean = False) As String
    Byte2HexString = StringUtils.LeftPad0(Hex$(b), 2)
    If prependSuffix Then Byte2HexString = "&H" & Byte2HexString
End Function

'******************************************************************************
'* [概  要] HexString2LongPtr
'* [詳  細] 1バイト分の16進文字列をバイト整数値に変換する。
'*
'* @param hVal 1バイト分の16進文字列
'* @return LongPtr バイト整数値
'******************************************************************************
Public Function HexString2LongPtr(ByVal hstr As String) As LongPtr
    HexString2LongPtr = CLngPtr(StringUtils.PrependIfMissingIgnoreCase(hstr, "&H"))
End Function

'******************************************************************************
'* [概  要] HexString2Byte
'* [詳  細] 1バイト分の16進文字列をバイトに変換する。
'*
'* @param hVal 1バイト分の16進文字列
'* @return Byte バイト
'******************************************************************************
Public Function HexString2Byte(ByVal hstr As String) As Byte
    HexString2Byte = CByte(StringUtils.PrependIfMissingIgnoreCase(hstr, "&H"))
End Function

'******************************************************************************
'* [概  要] IsHexString
'* [詳  細] 16進文字列かどうかチェック。
'*
'* @param val チェック対象文字列
'* @return Boolean チェック結果（True：OK、False：NG）
'******************************************************************************
Public Function IsHexString(ByVal val As String) As Boolean
    Dim i As Long

    For i = 1 To Len(val)
        If Not Mid(val, i, 1) Like "[0-9a-fA-F]" Then
            IsHexString = False
            Exit Function
        End If
    Next
    IsHexString = True
End Function

'******************************************************************************
'* [概  要] IsBinaryString
'* [詳  細] 2進文字列かどうかチェック。
'*
'* @param val チェック対象文字列
'* @return Boolean チェック結果（True：OK、False：NG）
'******************************************************************************
Public Function IsBinaryString(ByVal val As String) As Boolean
    If Len(val) Mod 4 <> 0 Then
        IsBinaryString = False
        Exit Function
    End If
    Dim i As Long

    For i = 1 To Len(val)
        If Not Mid(val, i, 1) Like "[0-1]" Then
            IsBinaryString = False
            Exit Function
        End If
    Next
    IsBinaryString = True
End Function

'******************************************************************************
'* [概  要] 文字列→バイトデータ変換処理。
'* [詳  細] 文字列からバイトデータへ変換を行う。
'*
'* @param text 対象となる文字列
'* @param cSet キャラセット（文字コード指定）。省略時は「UTF-8」。
'*               ※使用可能なCharset例：
'*                 utf-8、shift_jis、euc-jp、ISO-2022-JP、unicode（既定）
'*                 「HKEY_CLASSES_ROOT\MIME\Database\Charset」参照。
'* @return 変換後バイトデータ
'******************************************************************************
Public Function String2Binary(ByVal text As String, Optional ByVal cSet As String = "UTF-8") As Byte()
    String2Binary = StringUtils.GetBytes(text, cSet)
End Function

'******************************************************************************
'* [概  要] バイトデータ→文字列変換処理。
'* [詳  細] バイトデータから文字列へ変換を行う。
'*
'* @param text 対象となるバイトデータ
'* @param cSet キャラセット（文字コード指定）。省略時は「UTF-8」。
'*               ※使用可能なCharset例：
'*                 utf-8、shift_jis、euc-jp、ISO-2022-JP、unicode（既定）
'*                 「HKEY_CLASSES_ROOT\MIME\Database\Charset」参照。
'* @return 変換後文字列
'******************************************************************************
Public Function Binary2String(bin() As Byte, Optional ByVal cSet As String = "UTF-8") As String
    Binary2String = StringUtils.ToEncodedString(bin, cSet)
End Function
