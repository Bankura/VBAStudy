VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] デバッグユーティリティクラス
'* [詳  細] デバッグに関するユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'*
'* [参  考] https://excel-ubara.com/excelvba4/EXCEL_VBA_422.html
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

'******************************************************************************
'* 内部変数定義
'******************************************************************************

'******************************************************************************
'* プロパティ定義
'******************************************************************************


'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
End Sub
'******************************************************************************
'* メソッド定義
'******************************************************************************

Public Sub Show(var)
    On Error GoTo ErrorHandler
    Debug.Print VariantUtils.ToString(var)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number, Err.Source, Err.Description
End Sub


'******************************************************************************
'* [概  要] Dump
'* [詳  細] 指定した変数の内容をダンプ文字列に変換する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param x 対象の変数
'* @return String ダンプ文字列
'*
'******************************************************************************
Public Function Dump(ByVal x As Variant) As String
    Dump = Core.Dump(x)
End Function

'******************************************************************************
'* [概  要] PrintDump
'* [詳  細] 指定した変数の内容をダンプ文字列に変換し、イミディエイトウィンドウ
'*          に出力する。
'*
'* @param x 対象の変数
'*
'******************************************************************************
Public Sub PrintDump(ByVal x As Variant)
    Debug.Print Dump(x)
End Sub

'******************************************************************************
'* [概  要] Variant配列デバッグ出力処理
'* [詳  細] Variant配列の内容をイミディエイトウィンドウに出力する。
'*          2次元配列まで対応。
'*
'* @param vArr Variant配列
'******************************************************************************
Public Sub PrintArray(vArr)
    Dim arr As Variant
    If ValidateUtils.IsArrayEx(vArr) Then
        arr = vArr.ToArray
    ElseIf ValidateUtils.IsArray2DEx(vArr) Then
        arr = vArr.To2DArray
    ElseIf ValidateUtils.IsArray(vArr) Then
        arr = vArr
    Else
        Exit Sub
    End If
    If ArrayUtils.IsEmptyArray(arr) Then Exit Sub
    
    Dim i As Long, j As Long, tmp As StringEx, v
    If ArrayUtils.Is2DArray(arr) Then
    
        Dim collen() As Long
        ReDim collen(LBound(arr, 2) To UBound(arr, 2)) As Long
        Dim colNumAttr() As Boolean
        ReDim colNumAttr(LBound(arr, 2) To UBound(arr, 2)) As Boolean
        
        For j = LBound(arr, 2) To UBound(arr, 2)
            collen(j) = ArrayUtils.Array2DMaxTxtLength(arr, j, True)
            v = arr(LBound(arr, 1), j)
            colNumAttr(j) = ValidateUtils.IsNumericType(v)
        Next

        For i = LBound(arr, 1) To UBound(arr, 1)
            Set tmp = New StringEx
            Call tmp.Append("| ")
            For j = LBound(arr, 2) To UBound(arr, 2)
                If colNumAttr(j) Then
                    Call tmp.Append(StringUtils.LeftPad(arr(i, j), collen(j), , True))
                Else
                    Call tmp.Append(StringUtils.RightPad(arr(i, j), collen(j), , True))
                End If
                Call tmp.Append(" | ")
            Next
            Debug.Print tmp.ToString
        Next
    Else
        For i = LBound(arr, 1) To UBound(arr, 1)
            Debug.Print arr(i)
        Next
    End If
End Sub

Public Sub PrintEnumerator(obj)
    On Error Resume Next
    Dim v
    For Each v In obj
        Debug.Print v.Name
        Debug.Print v.Value
    Next
End Sub

'******************************************************************************
'* [概  要] Print2DArrayRow
'* [詳  細] Variant2次元配列の指定した行（1次元インデックス）の内容を
'*          イミディエイトウィンドウに出力する。
'*
'* @param vArr Variant配列
'* @param rowNum 行（1次元インデックス）
'******************************************************************************
Public Sub Print2DArrayRow(vArr, ByVal rowNum As Long)
    If Not ArrayUtils.Is2DArray(vArr) Then Exit Sub
    If Not ArrayUtils.IsCorrectRange(vArr, rowNum) Then Exit Sub
    Dim j As Long, tmp As String
    For j = LBound(vArr, 2) To UBound(vArr, 2)
        tmp = tmp & vArr(rowNum, j) & " | "
    Next
    Debug.Print tmp
    tmp = ""

End Sub

'******************************************************************************
'* [概  要] RecordFormatデバッグ出力処理
'* [詳  細] RecordFormatの内容をイミディエイトウィンドウに出力する。
'*
'* @param rf RecordFormat
'******************************************************************************
Public Sub PrintRecordFormat(rf As RecordFormat)
    Dim record As Collection, itm As Item, tmp As String
    For Each record In rf.RecordSet
        For Each itm In record
            tmp = tmp & itm.Value & " | "
        Next
        Debug.Print tmp
        tmp = ""
    Next
End Sub

'******************************************************************************
'* [概  要] DumpMemoryFromVariant
'* [詳  細] 変数のアドレス(VarPtr)とメモリのダンプをイミディエイトウィンドウ
'*          に出力する。
'*
'* @param var メモリのダンプを取得する変数
'* @return LongPtr メモリから取得したバイト整数値
'******************************************************************************
Public Function DumpMemoryFromVariant(ByRef var) As LongPtr
    Dim addr As LongPtr, size As Long, opt As String
    Dim vType As Long
    vType = CLng(BinaryUtils.ByteArray2LongPtr(BinaryUtils.CopyMemory2ByteArray(VarPtr(var), 2)))
    Select Case vType
        Case VT_I2, VT_BYREF + VT_I2     ' Integer
            size = 2
            opt = "dec"
        Case VT_I4, VT_DATE, VT_BYREF + VT_I4, VT_BYREF + VT_DATE   ' Long, Date
            size = 4
            opt = "dec"
        Case VT_R4, VT_BYREF + VT_R4     ' Single
            size = 4
            opt = "bin"
        Case VT_R8, VT_BYREF + VT_R8     ' Double
            size = 8
            opt = "bin"
        Case VT_BOOL, VT_UI1, VT_BYREF + VT_BOOL, VT_BYREF + VT_UI1  ' Boolean, Byte
            size = 1
            opt = "bin"
        Case VT_BSTR, VT_DISPATCH, VT_BYREF + VT_BSTR, VT_BYREF + VT_DISPATCH ' String, Object
            size = 8
            opt = "dec"
        Case Else ' LongLong, Currency
            size = 8
            opt = "bin"
    End Select
    
    If (vType And VT_BYREF) = VT_BYREF Then
        addr = BinaryUtils.ByteArray2LongPtr(BinaryUtils.CopyMemory2ByteArray(VarPtr(var) + LONGPTR_SIZE, LONGPTR_SIZE))
    Else
        addr = VarPtr(var) + LONGPTR_SIZE
    End If
    
    DumpMemoryFromVariant = DumpMemory(addr, size, opt)
End Function

'******************************************************************************
'* [概  要] DumpMemoryFromString
'* [詳  細] 変数のアドレス(StrPtr)とメモリのダンプをイミディエイトウィンドウ
'*          に出力する。
'*
'* @param var メモリのダンプを取得する変数
'* @return LongPtr メモリから取得したバイト整数値
'******************************************************************************
Public Function DumpMemoryFromString(ByRef var) As LongPtr
    DumpMemoryFromString = DumpMemory(StrPtr(var), LenB(var), "str")
End Function

'******************************************************************************
'* [概  要] DumpMemoryFromObject
'* [詳  細] 変数のアドレス(ObjPtr)とメモリのダンプをイミディエイトウィンドウ
'*          に出力する。
'*
'* @param var メモリのダンプを取得するObject変数
'* @return LongPtr メモリから取得したバイト整数値
'******************************************************************************
Public Function DumpMemoryFromObject(ByRef obj As Object) As LongPtr
    DumpMemoryFromObject = DumpMemory(ObjPtr(obj), LONGPTR_SIZE, "dec")
End Function

'******************************************************************************
'* [概  要] DumpMemory
'* [詳  細] 変数のアドレス(StrPtr)とメモリのダンプをイミディエイトウィンドウ
'*          に出力する。
'*
'* @param addr メモリダンプを取得するアドレス
'*             （VarPtr, StrPtr, ObjPtrの値をベースに指定する想定）
'* @param size メモリのダンプを取得するバイトサイズ
'* @param opt メモリダンプの内容に加えて表示する値の変換形式
'*              dec     ： バイト整数値
'*              str     ： 文字列
'*              bin     ： バイナリ（2進）
'*              指定なし： 表示しない
'* @return LongPtr メモリから取得したバイト整数値
'******************************************************************************
Public Function DumpMemory(ByVal addr As LongPtr, _
                    ByVal size As Long, _
                    Optional ByVal opt As String) As LongPtr

    If addr <= 0 Or size <= 0 Then
        Debug.Print addr & " :"
        Exit Function
    End If

    ReDim bData(1 To size) As Byte
    bData = BinaryUtils.CopyMemory2ByteArray(addr, size)
  
    Select Case LCase(opt)
        Case "str"
            Call DumpByteArrayAsString(bData, addr)
        Case "dec"
            Call DumpByteArrayAsDecimal(bData, addr)
        Case "bin"
            Call DumpByteArrayAsBinary(bData, addr)
        Case Else
            Call DumpByteArray(bData, addr)
    End Select

    DumpMemory = BinaryUtils.ByteArray2LongPtr(bData)
End Function

'******************************************************************************
'* [概  要] DumpByteArray
'* [詳  細] バイト情報を16進ダンプ文字列に変換してイミディエイトウィンドウに
'*          に出力する。
'*
'* @param bData() バイト情報（配列）
'* @param addr アドレス（追加表示用）。バイト情報の取得元情報として指定。
'******************************************************************************
Public Sub DumpByteArray(bData() As Byte, Optional ByVal addr)
    Dim pStr As String
    If ValidateUtils.IsNotMissing(addr) Then
        pStr = addr & " : "
    End If
    pStr = pStr & BinaryUtils.ByteArray2HexStrings(bData)
    Debug.Print pStr
End Sub

'******************************************************************************
'* [概  要] DumpByteArrayAsString
'* [詳  細] バイト情報を16進ダンプ文字列に変換してイミディエイトウィンドウに
'*          に出力する。
'*          メモリを文字列変換した内容を追加して出力する。
'*
'* @param bData() バイト情報（配列）
'* @param addr アドレス（追加表示用）。バイト情報の取得元情報として指定。
'******************************************************************************
Public Sub DumpByteArrayAsString(bData() As Byte, Optional ByVal addr)
    Dim pStr As String
    If ValidateUtils.IsNotMissing(addr) Then
        pStr = addr & " : "
    End If
    pStr = pStr & BinaryUtils.ByteArray2HexStrings(bData) & " ⇒ " & CStr(bData)
    Debug.Print pStr
End Sub

'******************************************************************************
'* [概  要] DumpByteArrayAsDecimal
'* [詳  細] バイト情報を16進ダンプ文字列に変換してイミディエイトウィンドウに
'*          に出力する。
'*          メモリを整数値に変換した内容を追加して出力する。
'*
'* @param bData() バイト情報（配列）
'* @param addr アドレス（追加表示用）。バイト情報の取得元情報として指定。
'******************************************************************************
Public Sub DumpByteArrayAsDecimal(bData() As Byte, Optional ByVal addr)
    Dim pStr As String
    If ValidateUtils.IsNotMissing(addr) Then
        pStr = addr & " : "
    End If
    pStr = pStr & BinaryUtils.ByteArray2HexStrings(bData) & " ⇒ " & BinaryUtils.ByteArray2LongPtr(bData)
    Debug.Print pStr
End Sub

'******************************************************************************
'* [概  要] DumpByteArrayAsBinary
'* [詳  細] バイト情報を16進ダンプ文字列に変換してイミディエイトウィンドウに
'*          に出力する。
'*          メモリをバイナリ変換した内容を追加して出力する。
'*
'* @param bData() バイト情報（配列）
'* @param addr アドレス（追加表示用）。バイト情報の取得元情報として指定。
'******************************************************************************
Public Sub DumpByteArrayAsBinary(bData() As Byte, Optional ByVal addr)
    Dim pStr As String
    If ValidateUtils.IsNotMissing(addr) Then
        pStr = addr & " : "
    End If
    pStr = pStr & BinaryUtils.ByteArray2HexStrings(bData) & " ⇒ " & BinaryUtils.ByteArray2BinaryStrings(bData)
    Debug.Print pStr
End Sub

'******************************************************************************
'* [概  要] MesureExcuteTime
'* [詳  細] 関数を実行し、実行した時間をイミディエイトウィンドウに出力する。
'*
'* @param fnc       実行する関数（bas内のFunctionであること）
'* @param arg1-arg8 関数に渡す引数
'* @return          実行結果
'******************************************************************************
Public Function MesureExcuteTime(fnc As Func, _
    Optional ByRef arg1 As Variant, _
    Optional ByRef arg2 As Variant, _
    Optional ByRef arg3 As Variant, _
    Optional ByRef arg4 As Variant, _
    Optional ByRef arg5 As Variant, _
    Optional ByRef arg6 As Variant, _
    Optional ByRef arg7 As Variant, _
    Optional ByRef arg8 As Variant _
    ) As Variant
    
    Dim StartTime, endTime
    Dim startMs As Double, endMs As Double, excuteMs As Double
    StartTime = DateTimeUtils.GetTimestampString
    startMs = DateTimeUtils.GetMilliSecond
    
    Dim ret
    If IsMissing(arg1) Then ret = fnc.Apply Else _
    If IsMissing(arg2) Then ret = fnc.Apply(arg1) Else _
    If IsMissing(arg3) Then ret = fnc.Apply(arg1, arg2) Else _
    If IsMissing(arg4) Then ret = fnc.Apply(arg1, arg2, arg3) Else _
    If IsMissing(arg5) Then ret = fnc.Apply(arg1, arg2, arg3, arg4) Else _
    If IsMissing(arg6) Then ret = fnc.Apply(arg1, arg2, arg3, arg4, arg5) Else _
    If IsMissing(arg7) Then ret = fnc.Apply(arg1, arg2, arg3, arg4, arg5, arg6) Else _
    If IsMissing(arg8) Then ret = fnc.Apply(arg1, arg2, arg3, arg4, arg5, arg6, arg7) Else _
    ret = fnc.Apply(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8)
    
    endMs = DateTimeUtils.GetMilliSecond
    endTime = DateTimeUtils.GetTimestampString
    excuteMs = endMs - startMs
    
    Debug.Print "処理開始： " & StartTime
    Debug.Print "処理終了： " & endTime
    Debug.Print "実行時間： " & excuteMs & " ms"
    
    MesureExcuteTime = ret
End Function

'******************************************************************************
'* [概  要] WashImmediateWindow
'* [詳  細] イミディエイトウィンドウに1000行分の改行を出力し、現在の表示を
'*          内容を流す。
'*
'******************************************************************************
Public Sub WashImmediateWindow()
    Debug.Print String(1000, vbNewLine)
End Sub

'******************************************************************************
'* [概  要] ShowFormatCondition
'* [詳  細] 設定されている条件付き書式の情報を表示する。
'*
'* @param sheetName 対象のシート名（省略すると選択中のシートで処理）
'* @param wb ワークブックオブジェクト
'******************************************************************************
Public Sub ShowFormatCondition(Optional sheetName As String, Optional ByRef wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim ws As Worksheet
    If sheetName = "" Then
        Set ws = wb.ActiveSheet
    Else
        Set ws = wb.Worksheets(sheetName)
    End If
    
    Dim fc As FormatCondition, cnt As Long
    For Each fc In ws.Cells.FormatConditions
        cnt = cnt + 1
        Debug.Print "No." & cnt & _
                    " タイプ:[" & GetFormatConditionsTypeName(fc.Type) & "]" & _
                    " 条件:[" & fc.Formula1 & "]" & _
                    " 範囲:[" & fc.AppliesTo.Address & "]" & _
                    " 背景色:[" & GetRGBColorName(fc.Interior.Color) & "]" & _
                    " フォントスタイル:[" & fc.Font.FontStyle & "]" & _
                    " フォント色:[" & GetColorIndexName(fc.Font.ColorIndex) & "]" & _
                    " 優先度:[" & fc.Priority & "]"
    Next
End Sub

'******************************************************************************
'* [概  要] GetFormatConditionsTypeName
'* [詳  細] 条件付き書式のタイプの名称を取得する。
'*
'* @param typeNum 条件付き書式のタイプ
'* @return        条件付き書式のタイプ名称
'******************************************************************************
Public Function GetFormatConditionsTypeName(typeNum As XlFormatConditionType) As String
    Select Case typeNum
        Case XlFormatConditionType.xlAboveAverageCondition
            GetFormatConditionsTypeName = "平均以上の条件"
        Case XlFormatConditionType.xlBlanksCondition
            GetFormatConditionsTypeName = "空白の条件"
        Case XlFormatConditionType.xlCellValue
            GetFormatConditionsTypeName = "セルの値"
        Case XlFormatConditionType.xlColorScale
            GetFormatConditionsTypeName = "カラースケール"
        Case XlFormatConditionType.xlDatabar
            GetFormatConditionsTypeName = "データバー"
        Case XlFormatConditionType.xlErrorsCondition
            GetFormatConditionsTypeName = "エラー条件"
        Case XlFormatConditionType.xlExpression
            GetFormatConditionsTypeName = "演算"
        Case XlFormatConditionType.xlIconSets
            GetFormatConditionsTypeName = "アイコンセット"
        Case XlFormatConditionType.xlNoBlanksCondition
            GetFormatConditionsTypeName = "空白の条件なし"
        Case XlFormatConditionType.xlNoErrorsCondition
            GetFormatConditionsTypeName = "エラー条件なし"
        Case XlFormatConditionType.xlTextString
            GetFormatConditionsTypeName = "テキスト文字列"
        Case XlFormatConditionType.xlTimePeriod
            GetFormatConditionsTypeName = "期間"
        Case XlFormatConditionType.xlTop10
            GetFormatConditionsTypeName = "上位10の値"
        Case XlFormatConditionType.xlUniqueValues
            GetFormatConditionsTypeName = "一意の値"
    End Select
End Function

'******************************************************************************
'* [概  要] GetRGBColorName
'* [詳  細] Color値に対応する名称を取得する。
'*
'* @param lColorValue Color値
'* @return            Color名称
'******************************************************************************
Public Function GetRGBColorName(lColorValue) As String
    If IsNull(lColorValue) Then
        Exit Function
    End If
    
    Select Case lColorValue
        Case XlRgbColor.rgbBlack
            GetRGBColorName = "黒"
        Case XlRgbColor.rgbNavy
            GetRGBColorName = "ネイビー"
        Case XlRgbColor.rgbNavyBlue
            GetRGBColorName = "ネイビーブルー"
        Case XlRgbColor.rgbDarkBlue
            GetRGBColorName = "濃い青"
        Case XlRgbColor.rgbMediumBlue
            GetRGBColorName = "淡い青"
        Case XlRgbColor.rgbBlue
            GetRGBColorName = "青"
        Case XlRgbColor.rgbDarkGreen
            GetRGBColorName = "濃い緑"
        Case XlRgbColor.rgbGreen
            GetRGBColorName = "緑"
        Case XlRgbColor.rgbTeal
            GetRGBColorName = "青緑"
        Case XlRgbColor.rgbDarkCyan
            GetRGBColorName = "濃いシアン"
        Case XlRgbColor.rgbLightCyan
            GetRGBColorName = "明るい水色"
        Case XlRgbColor.rgbDeepSkyBlue
            GetRGBColorName = "深いスカイブルー"
        Case XlRgbColor.rgbDarkTurquoise
            GetRGBColorName = "濃いターコイズ"
        Case XlRgbColor.rgbMediumSpringGreen
            GetRGBColorName = "淡いスプリンググリーン"
        Case XlRgbColor.rgbLime
            GetRGBColorName = "黄緑"
        Case XlRgbColor.rgbSpringGreen
            GetRGBColorName = "スプリンググリーン"
        Case XlRgbColor.rgbAqua
            GetRGBColorName = "水色"
        Case XlRgbColor.rgbMidnightBlue
            GetRGBColorName = "ミッドナイトブルー"
        Case XlRgbColor.rgbDodgerBlue
            GetRGBColorName = "ドジャーブルー"
        Case XlRgbColor.rgbLightSeaGreen
            GetRGBColorName = "薄いシーグリーン"
        Case XlRgbColor.rgbForestGreen
            GetRGBColorName = "フォレストグリーン"
        Case XlRgbColor.rgbSeaGreen
            GetRGBColorName = "シーグリーン"
        Case XlRgbColor.rgbDarkSlateGray
            GetRGBColorName = "濃いスレートグレー"
        Case XlRgbColor.rgbLimeGreen
            GetRGBColorName = "ライムグリーン"
        Case XlRgbColor.rgbMediumSeaGreen
            GetRGBColorName = "淡いシーグリーン"
        Case XlRgbColor.rgbTurquoise
            GetRGBColorName = "ターコイズ"
        Case XlRgbColor.rgbRoyalBlue
            GetRGBColorName = "ロイヤルブルー"
        Case XlRgbColor.rgbSteelBlue
            GetRGBColorName = "スチールブルー"
        Case XlRgbColor.rgbDarkSlateBlue
            GetRGBColorName = "濃いスレートブルー"
        Case XlRgbColor.rgbMediumTurquoise
            GetRGBColorName = "淡いターコイズ"
        Case XlRgbColor.rgbIndigo
            GetRGBColorName = "インディゴ"
        Case XlRgbColor.rgbDarkOliveGreen
            GetRGBColorName = "濃いオリーブグリーン"
        Case XlRgbColor.rgbCadetBlue
            GetRGBColorName = "カデットブルー"
        Case XlRgbColor.rgbCornflowerBlue
            GetRGBColorName = "コーンフラワーブルー"
        Case XlRgbColor.rgbMediumAquamarine
            GetRGBColorName = "淡いアクアマリン"
        Case XlRgbColor.rgbDimGray
            GetRGBColorName = "ディムグレー"
        Case XlRgbColor.rgbSlateBlue
            GetRGBColorName = "スレートブルー"
        Case XlRgbColor.rgbOliveDrab
            GetRGBColorName = "オリーブドラブ"
        Case XlRgbColor.rgbSlateGray
            GetRGBColorName = "スレートグレー"
        Case XlRgbColor.rgbLightSlateGray
            GetRGBColorName = "薄いスレートグレー"
        Case XlRgbColor.rgbMediumSlateBlue
            GetRGBColorName = "淡いスレートブルー"
        Case XlRgbColor.rgbLawnGreen
            GetRGBColorName = "若草色"
        Case XlRgbColor.rgbChartreuse
            GetRGBColorName = "シャルトルーズ"
        Case XlRgbColor.rgbAquamarine
            GetRGBColorName = "アクアマリン"
        Case XlRgbColor.rgbMaroon
            GetRGBColorName = "栗色"
        Case XlRgbColor.rgbPurple
            GetRGBColorName = "紫"
        Case XlRgbColor.rgbOlive
            GetRGBColorName = "オリーブ"
        Case XlRgbColor.rgbGray
            GetRGBColorName = "灰色"
        Case XlRgbColor.rgbSkyBlue
            GetRGBColorName = "スカイブルー"
        Case XlRgbColor.rgbLightSkyBlue
            GetRGBColorName = "薄いスカイブルー"
        Case XlRgbColor.rgbBlueViolet
            GetRGBColorName = "青紫"
        Case XlRgbColor.rgbDarkRed
            GetRGBColorName = "濃い赤"
        Case XlRgbColor.rgbDarkMagenta
            GetRGBColorName = "濃いマゼンタ"
        Case XlRgbColor.rgbDarkSeaGreen
            GetRGBColorName = "濃いシーグリーン"
        Case XlRgbColor.rgbLightGreen
            GetRGBColorName = "明るい緑"
        Case XlRgbColor.rgbMediumPurple
            GetRGBColorName = "淡い紫"
        Case XlRgbColor.rgbDarkViolet
            GetRGBColorName = "濃い紫"
        Case XlRgbColor.rgbPaleGreen
            GetRGBColorName = "ペールグリーン"
        Case XlRgbColor.rgbDarkOrchid
            GetRGBColorName = "濃いオーキッド"
        Case XlRgbColor.rgbYellowGreen
            GetRGBColorName = "イエローグリーン"
        Case XlRgbColor.rgbSienna
            GetRGBColorName = "シェンナ"
        Case XlRgbColor.rgbBrown
            GetRGBColorName = "茶"
        Case XlRgbColor.rgbDarkGray
            GetRGBColorName = "濃い灰色"
        Case XlRgbColor.rgbLightBlue
            GetRGBColorName = "明るい青"
        Case XlRgbColor.rgbGreenYellow
            GetRGBColorName = "グリーンイエロー"
        Case XlRgbColor.rgbPaleTurquoise
            GetRGBColorName = "ペールターコイズ"
        Case XlRgbColor.rgbLightSteelBlue
            GetRGBColorName = "薄いスチールブルー"
        Case XlRgbColor.rgbPowderBlue
            GetRGBColorName = "パウダーブルー"
        Case XlRgbColor.rgbFireBrick
            GetRGBColorName = "れんが色"
        Case XlRgbColor.rgbDarkGoldenrod
            GetRGBColorName = "濃いゴールデンロッド"
        Case XlRgbColor.rgbMediumOrchid
            GetRGBColorName = "淡いオーキッド"
        Case XlRgbColor.rgbRosyBrown
            GetRGBColorName = "ローズブラウン"
        Case XlRgbColor.rgbDarkKhaki
            GetRGBColorName = "濃いカーキ"
        Case XlRgbColor.rgbSilver
            GetRGBColorName = "銀色"
        Case XlRgbColor.rgbMediumVioletRed
            GetRGBColorName = "淡いバイオレットレッド"
        Case XlRgbColor.rgbIndianRed
            GetRGBColorName = "インディアンレッド"
        Case XlRgbColor.rgbPeru
            GetRGBColorName = "ペルー"
        Case XlRgbColor.rgbTan
            GetRGBColorName = "タン"
        Case XlRgbColor.rgbLightGray
            GetRGBColorName = "薄い灰色"
        Case XlRgbColor.rgbThistle
            GetRGBColorName = "あざみ色"
        Case XlRgbColor.rgbOrchid
            GetRGBColorName = "オーキッド"
        Case XlRgbColor.rgbGoldenrod
            GetRGBColorName = "ゴールデンロッド"
        Case XlRgbColor.rgbPaleVioletRed
            GetRGBColorName = "ペールバイオレットレッド"
        Case XlRgbColor.rgbCrimson
            GetRGBColorName = "深紅"
        Case XlRgbColor.rgbGainsboro
            GetRGBColorName = "ゲーンズボロ"
        Case XlRgbColor.rgbPlum
            GetRGBColorName = "プラム"
        Case XlRgbColor.rgbBurlyWood
            GetRGBColorName = "バーリーウッド"
        Case XlRgbColor.rgbLavender
            GetRGBColorName = "ラベンダー"
        Case XlRgbColor.rgbDarkSalmon
            GetRGBColorName = "濃いサーモンピンク"
        Case XlRgbColor.rgbViolet
            GetRGBColorName = "紫色"
        Case XlRgbColor.rgbPaleGoldenrod
            GetRGBColorName = "ペールゴールデンロッド"
        Case XlRgbColor.rgbLightCoral
            GetRGBColorName = "薄いさんご"
        Case XlRgbColor.rgbKhaki
            GetRGBColorName = "カーキ"
        Case XlRgbColor.rgbAliceBlue
            GetRGBColorName = "アリスブルー"
        Case XlRgbColor.rgbHoneydew
            GetRGBColorName = "ハニーデュー"
        Case XlRgbColor.rgbAzure
            GetRGBColorName = "空色"
        Case XlRgbColor.rgbSandyBrown
            GetRGBColorName = "サンディブラウン"
        Case XlRgbColor.rgbWheat
            GetRGBColorName = "小麦"
        Case XlRgbColor.rgbBeige
            GetRGBColorName = "ベージュ"
        Case XlRgbColor.rgbWhiteSmoke
            GetRGBColorName = "ホワイトスモーク"
        Case XlRgbColor.rgbMintCream
            GetRGBColorName = "ミントクリーム"
        Case XlRgbColor.rgbGhostWhite
            GetRGBColorName = "ゴーストホワイト"
        Case XlRgbColor.rgbSalmon
            GetRGBColorName = "サーモンピンク"
        Case XlRgbColor.rgbAntiqueWhite
            GetRGBColorName = "アンティークホワイト"
        Case XlRgbColor.rgbLinen
            GetRGBColorName = "リネン"
        Case XlRgbColor.rgbLightGoldenrodYellow
            GetRGBColorName = "薄いゴールデンロッドイエロー"
        Case XlRgbColor.rgbOldLace
            GetRGBColorName = "オールドレース"
        Case XlRgbColor.rgbRed
            GetRGBColorName = "赤"
        Case XlRgbColor.rgbFuchsia
            GetRGBColorName = "明るい紫"
        Case XlRgbColor.rgbDeepPink
            GetRGBColorName = "深いピンク"
        Case XlRgbColor.rgbOrangeRed
            GetRGBColorName = "オレンジレッド"
        Case XlRgbColor.rgbTomato
            GetRGBColorName = "トマト"
        Case XlRgbColor.rgbHotPink
            GetRGBColorName = "ホットピンク"
        Case XlRgbColor.rgbCoral
            GetRGBColorName = "さんご"
        Case XlRgbColor.rgbDarkOrange
            GetRGBColorName = "濃いオレンジ"
        Case XlRgbColor.rgbLightSalmon
            GetRGBColorName = "薄いサーモンピンク"
        Case XlRgbColor.rgbOrange
            GetRGBColorName = "オレンジ"
        Case XlRgbColor.rgbLightPink
            GetRGBColorName = "薄いピンク"
        Case XlRgbColor.rgbPink
            GetRGBColorName = "ピンク"
        Case XlRgbColor.rgbGold
            GetRGBColorName = "ゴールド"
        Case XlRgbColor.rgbPeachPuff
            GetRGBColorName = "ピーチパフ"
        Case XlRgbColor.rgbNavajoWhite
            GetRGBColorName = "ナバホホワイト"
        Case XlRgbColor.rgbMoccasin
            GetRGBColorName = "モカシン"
        Case XlRgbColor.rgbBisque
            GetRGBColorName = "ビスク"
        Case XlRgbColor.rgbMistyRose
            GetRGBColorName = "ミスティローズ"
        Case XlRgbColor.rgbBlanchedAlmond
            GetRGBColorName = "ブランシュアーモンド"
        Case XlRgbColor.rgbPapayaWhip
            GetRGBColorName = "パパイヤホイップ"
        Case XlRgbColor.rgbLavenderBlush
            GetRGBColorName = "ラベンダーブラッシュ"
        Case XlRgbColor.rgbSeashell
            GetRGBColorName = "シーシェル"
        Case XlRgbColor.rgbCornsilk
            GetRGBColorName = "コーンシルク"
        Case XlRgbColor.rgbLemonChiffon
            GetRGBColorName = "レモンシフォン"
        Case XlRgbColor.rgbFloralWhite
            GetRGBColorName = "フローラルホワイト"
        Case XlRgbColor.rgbSnow
            GetRGBColorName = "スノー"
        Case XlRgbColor.rgbYellow
            GetRGBColorName = "黄"
        Case XlRgbColor.rgbLightYellow
            GetRGBColorName = "明るい黄"
        Case XlRgbColor.rgbIvory
            GetRGBColorName = "アイボリー"
        Case XlRgbColor.rgbWhite
            GetRGBColorName = "白"
        Case Else
            GetRGBColorName = "？"
    End Select
End Function

'******************************************************************************
'* [概  要] GetColorIndexName
'* [詳  細] ColorIndex値に対応する名称を取得する。
'*
'* @param lColorIndexValue ColorIndex値
'* @return                 ColorIndex名称
'******************************************************************************
Public Function GetColorIndexName(lColorIndexValue) As String
    If IsNull(lColorIndexValue) Then
        Exit Function
    End If
    
    Select Case lColorIndexValue
        Case 1
            GetColorIndexName = "黒"
        Case 2
            GetColorIndexName = "白"
        Case 3
            GetColorIndexName = "赤"
        Case 4
            GetColorIndexName = "黄緑"
        Case 5
            GetColorIndexName = "青"
        Case 6
            GetColorIndexName = "黄"
        Case 7
            GetColorIndexName = "明るい紫"
        Case 8
            GetColorIndexName = "水色"
        Case 9
            GetColorIndexName = "栗色"
        Case 10
            GetColorIndexName = "緑"
        Case 11
            GetColorIndexName = "ネイビー"
        Case 12
            GetColorIndexName = "オリーブ"
        Case 13
            GetColorIndexName = "紫"
        Case 14
            GetColorIndexName = "青緑"
        Case 15
            GetColorIndexName = "銀色"
        Case 16
            GetColorIndexName = "灰色"
        Case Else
            GetColorIndexName = "？"
    End Select
End Function

'******************************************************************************
'* [概  要] GetRGBRedValue
'* [詳  細] Color値に対応するRGB値のRedの値を取得する。
'*
'* @param lColorValue Color値
'* @return            RGB(Red)値
'******************************************************************************
Public Function GetRGBRedValue(lColorValue As Long) As Long
    GetRGBRedValue = lColorValue Mod 256
End Function

'******************************************************************************
'* [概  要] GetRGBGreenValue
'* [詳  細] Color値に対応するRGB値のGreenの値を取得する。
'*
'* @param lColorValue Color値
'* @return            RGB(Green)値
'******************************************************************************
Public Function GetRGBGreenValue(lColorValue As Long) As Long
    GetRGBGreenValue = Int(lColorValue / 256) Mod 256
End Function

'******************************************************************************
'* [概  要] GetRGBBlueValue
'* [詳  細] Color値に対応するRGB値のBlueの値を取得する。
'*
'* @param lColorValue Color値
'* @return            RGB(Blue)値
'******************************************************************************
Public Function GetRGBBlueValue(lColorValue As Long) As Long
    GetRGBBlueValue = Int(lColorValue / 256 / 256)
End Function

'******************************************************************************
'* [概  要] GetRGBHexValue
'* [詳  細] Color値に対応するHex値を取得する。
'*
'* @param lColorValue Color値
'* @return            Hex値
'******************************************************************************
Public Function GetRGBHexValue(lColorValue As Long) As String
    GetRGBHexValue = "#" & Hex(GetRGBRedValue(lColorValue)) & Hex(GetRGBGreenValue(lColorValue)) & Hex(GetRGBBlueValue(lColorValue))
End Function

'******************************************************************************
'* [概  要] GetRGBText
'* [詳  細] Color値に対応するRGB値を取得する。
'*          形式： RGB(xxx, xxx, xxx)
'*
'* @param lColorValue Color値
'* @return            RGB値（形式： RGB(xxx, xxx, xxx)）
'******************************************************************************
Public Function GetRGBText(lColorValue As Long) As String
    GetRGBText = "RGB(" & GetRGBRedValue(lColorValue) & ", " & GetRGBGreenValue(lColorValue) & ", " & GetRGBBlueValue(lColorValue) & ")"
End Function
