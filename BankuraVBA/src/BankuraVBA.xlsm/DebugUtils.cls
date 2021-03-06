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
    Debug.Print var
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number, Err.source, Err.Description
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
Public Sub PrintVariantArray(vArr)
    Dim arr As Variant
    If ValidateUtils.IsArrayEx(vArr) Then
        arr = vArr.ToArray
    ElseIf ValidateUtils.IsArray2DEx(vArr) Then
        arr = vArr.ToArray
    ElseIf ValidateUtils.IsArray(vArr) Then
        arr = vArr
    Else
        Exit Sub
    End If
    If ArrayUtils.IsEmptyArray(arr) Then Exit Sub
    
    Dim i As Long, j As Long, tmp As String
    For i = LBound(arr, 1) To UBound(arr, 1)
        If ArrayUtils.Is2DArray(arr) Then
            For j = LBound(arr, 2) To UBound(arr, 2)
                tmp = tmp & arr(i, j) & " | "
            Next
        Else
            tmp = arr(i)
        End If
        Debug.Print tmp
        tmp = ""
    Next
End Sub

Public Sub PrintEnumerator(obj)
    On Error Resume Next
    Dim v
    For Each v In obj
        Debug.Print v.name
        Debug.Print v.value
    Next
End Sub


'******************************************************************************
'* [概  要] Print2DArrayRow
'* [詳  細] Variant2次元配列の指定した行（1次元インデックス）の内容を
'*          イミディエイトウィンドウに出力する。
'*
'* @param vArr Variant配列
'* @param row 行（1次元インデックス）
'******************************************************************************
Public Sub Print2DArrayRow(vArr, ByVal row As Long)
    Dim j As Long, tmp As String
    For j = LBound(vArr, 2) To UBound(vArr, 2)
        tmp = tmp & vArr(row, j) & " | "
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
            tmp = tmp & itm.value & " | "
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
    Dim addr As LongPtr, Size As Long, opt As String
    Dim vType As Long
    vType = CLng(BinaryUtils.ByteArray2LongPtr(BinaryUtils.CopyMemory2ByteArray(VarPtr(var), 2)))
    Select Case vType
        Case VT_I2, VT_BYREF + VT_I2     ' Integer
            Size = 2
            opt = "dec"
        Case VT_I4, VT_DATE, VT_BYREF + VT_I4, VT_BYREF + VT_DATE   ' Long, Date
            Size = 4
            opt = "dec"
        Case VT_R4, VT_BYREF + VT_R4     ' Single
            Size = 4
            opt = "bin"
        Case VT_R8, VT_BYREF + VT_R8     ' Double
            Size = 8
            opt = "bin"
        Case VT_BOOL, VT_UI1, VT_BYREF + VT_BOOL, VT_BYREF + VT_UI1  ' Boolean, Byte
            Size = 1
            opt = "bin"
        Case VT_BSTR, VT_DISPATCH, VT_BYREF + VT_BSTR, VT_BYREF + VT_DISPATCH ' String, Object
            Size = 8
            opt = "dec"
        Case Else ' LongLong, Currency
            Size = 8
            opt = "bin"
    End Select
    
    If (vType And VT_BYREF) = VT_BYREF Then
        addr = BinaryUtils.ByteArray2LongPtr(BinaryUtils.CopyMemory2ByteArray(VarPtr(var) + LONGPTR_SIZE, LONGPTR_SIZE))
    Else
        addr = VarPtr(var) + LONGPTR_SIZE
    End If
    
    DumpMemoryFromVariant = DumpMemory(addr, Size, opt)
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
                    ByVal Size As Long, _
                    Optional ByVal opt As String) As LongPtr

    If addr <= 0 Or Size <= 0 Then
        Debug.Print addr & " :"
        Exit Function
    End If

    ReDim bData(1 To Size) As Byte
    bData = BinaryUtils.CopyMemory2ByteArray(addr, Size)
  
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



