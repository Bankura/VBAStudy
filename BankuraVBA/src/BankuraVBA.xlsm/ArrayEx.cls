VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''+----                                                                   --+
'''|                             Ariawase 0.9.0                              |
'''|                Ariawase is free library for VBA cowboys.                |
'''|          The Project Page: https://github.com/vbaidiot/Ariawase         |
'''+--                                                                   ----+
Option Explicit

Private xItems As Variant
Private xLength As Long
Private xIndex As Long

' ADD Bankura
Private mBaseStartIndex As Long

Private Sub Class_Initialize()
    xIndex = -1
    xLength = -1 + 32
    ReDim xItems(xLength - 1)
End Sub

Private Sub Extend()
    If xIndex < xLength Then GoTo Escape
    xLength = xLength + 1 'possible overflow (Err.Raise 6)
    xLength = -1 + xLength + xLength
    ReDim Preserve xItems(xLength - 1)
    
Escape:
End Sub

Public Function AddVal(ByVal val As Variant)
    xIndex = xIndex + 1
    Extend
    Let xItems(xIndex) = val
End Function

Public Function AddObj(ByVal obj As Variant)
    xIndex = xIndex + 1
    Extend
    Set xItems(xIndex) = obj
End Function

'MOD Bankura
Public Function ToArray() As Variant
    Dim arr As Variant: arr = xItems
    If xIndex > -1 Then
        ReDim Preserve arr(xIndex)
    Else
        arr = Array()
    End If
    If mBaseStartIndex <= 0 Then
        ToArray = arr
    Else
        ToArray = ArrayUtils.ArrayExToArray(Me, mBaseStartIndex)
    End If
End Function


'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* 拡張メソッド
'*
'* @author Bankura
'* Copyright (c) 2020-2022 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'*-----------------------------------------------------------------------------
'* Item プロパティ
'*-----------------------------------------------------------------------------
Public Property Get Item(ByVal idx As Long) As Variant
Attribute Item.VB_UserMemId = 0
    If IsInvalidIndex(idx) Then
        Item = Empty
        Exit Property
    End If

    If IsObject(xItems(idx)) Then
        Set Item = xItems(idx)
    Else
        Item = xItems(idx)
    End If
End Property
Public Property Let Item(ByVal idx As Long, ByVal arg As Variant)
    If IsInvalidIndex(idx) Then
        Exit Property
    End If
    
    If IsObject(arg) Then
        Set xItems(idx) = arg
    Else
        xItems(idx) = arg
    End If
End Property

'*-----------------------------------------------------------------------------
'* BaseStartIndex プロパティ
'* 1次元配列変換時の添え字の開始インデックス
'*-----------------------------------------------------------------------------
Public Property Get BaseStartIndex() As Long
    BaseStartIndex = mBaseStartIndex
End Property
Public Property Let BaseStartIndex(ByVal arg As Long)
    If arg < 0 Then
        mBaseStartIndex = 0
    Else
        mBaseStartIndex = arg
    End If
End Property

Private Function IsInvalidIndex(ByVal idx As Long) As Boolean
    IsInvalidIndex = (idx < 0 Or xIndex < idx)
End Function


'******************************************************************************
'* [概  要] Init メソッド
'* [詳  細] 初期化用メソッド｡
'*
'* @param val 初期配列
'* @return ArrayExオブジェクト
'*
'******************************************************************************
Public Function Init(Optional ByVal val) As ArrayEx
    If VBA.IsMissing(val) Then
        Call Clear
        Set Init = Me
        Exit Function
    End If
    If IsArray(val) Then
        xIndex = UBound(val)
        xLength = xIndex + 32
        xItems = val
        ReDim Preserve xItems(xLength - 1)
    Else
        Call Clear
        Call Add(val)
    End If
    Set Init = Me
End Function

'******************************************************************************
'* [概  要] Clear
'* [詳  細] 初期化する｡
'*
'* @return ArrayExオブジェクト
'*
'******************************************************************************
Public Function Clear() As ArrayEx
    xIndex = -1
    xLength = -1 + 32
    ReDim xItems(xLength - 1)
    Set Clear = Me
End Function

'******************************************************************************
'* [概  要] Length
'* [詳  細] 要素数を返却する｡
'*
'* @return Long 要素数
'*
'******************************************************************************
Property Get Length() As Long
    Length = xIndex + 1
End Property

'******************************************************************************
'* [概  要] LastIndex
'* [詳  細] 最終インデックスを返却する｡
'*
'* @return Long 最終インデックス
'*
'******************************************************************************
Property Get LastIndex() As Long
    LastIndex = xIndex
End Property

'******************************************************************************
'* [概  要] Add
'* [詳  細] 項目を追加する｡
'*
'* @param var 追加項目
'* @return ArrayEx
'*
'******************************************************************************
Public Function Add(ByVal var As Variant) As ArrayEx
    If IsObject(var) Then
        Call AddObj(var)
    Else
        Call AddVal(var)
    End If
    Set Add = Me
End Function

'******************************************************************************
'* [概  要] AddAll
'* [詳  細] 項目を複数追加する｡
'*
'* @param vars 追加項目（複数）
'* @return ArrayEx
'*
'******************************************************************************
Public Function AddAll(ParamArray vars() As Variant) As ArrayEx
    Dim tmp: tmp = vars
    Dim v
    For Each v In ArrayUtils.ToVArrayFlattened(tmp)
        Add v
    Next
    Set AddAll = Me
End Function

'******************************************************************************
'* [概  要] Expand
'* [詳  細] 配列領域を拡張する｡
'*
'* @param xAddCount 追加項目数
'* @return ArrayEx
'*
'******************************************************************************
Public Function Expand(ByVal xAddCount As Long) As ArrayEx
    xIndex = xIndex + xAddCount
    Call Extend
    Set Expand = Me
End Function

'******************************************************************************
'* [概  要] IsEmptyArray
'* [詳  細] 配列が未初期化・空配列か判定する｡
'*
'* @return Boolean 判定結果(True:未初期化・空)
'*
'******************************************************************************
Public Function IsEmptyArray() As Boolean
    IsEmptyArray = (xIndex = -1)
End Function

'******************************************************************************
'* [概  要] Equals
'* [詳  細] 一次元配列を比較し、イコールであればTrueを返却する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param arr2 As Variant(Of Array(Of T))
'* @param swAllowNull As Boolean
'* @return As Variant(Of Boolean Or Null Or Empty)
'******************************************************************************
Public Function Equals(ByVal arr2 As Variant, Optional ByVal swAllowNull As Boolean = False) As Variant
    Equals = Core.ArrEquals(Me.ToArray, arr2, swAllowNull)
End Function

'******************************************************************************
'* [概  要] Compare
'* [詳  細] 一次元配列（x,y）を比較し、x=yであれば0、x<yであれば-1、x>y
'*          であれば1を返却する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param arr2 As Variant(Of Array(Of T))
'* @return As Variant(Of Integer Or Null)
'******************************************************************************
Public Function Compare(ByVal arr2 As Variant, Optional ByVal swAllowNull As Boolean = False) As Variant
    Compare = Core.ArrCompare(Me.ToArray, arr2, swAllowNull)
End Function

'******************************************************************************
'* [概  要] IndexOf
'* [詳  細] 一次元配列から指定した値と一致する要素のインデックスを返却する。
'*          値の捜査開始位置と要素数を任意指定可能。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param val As Variant(Of T)
'* @param ixStart As Variant(Of Long)
'* @param cnt As Variant(Of Long)
'* @return As Long
'******************************************************************************
Public Function IndexOf(ByVal val As Variant, Optional ByVal ixStart As Variant, Optional ByVal cnt As Variant) As Long
    IndexOf = Core.ArrIndexOf(Me.ToArray, val, ixStart, cnt)
End Function

'******************************************************************************
'* [概  要] Reverse
'* [詳  細] 一次元配列の要素の並び順を反転させる。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param arr As Variant(Of Array(Of T))
'* @return ArrayEx
'******************************************************************************
Public Function reverse() As ArrayEx
    Dim arr: arr = Me.ToArray
    Call Core.ArrRev(arr)
    Set reverse = Init(arr)
End Function

'******************************************************************************
'* [概  要] Sort
'* [詳  細] 一次元配列のソートを行う。デフォルトは昇順。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param orderAsc As Boolean
'* @return ArrayEx
'******************************************************************************
Public Function Sort(Optional ByVal orderAsc As Boolean = True) As ArrayEx
    Dim arr: arr = Me.ToArray
    Call Core.ArrSort(arr, orderAsc)
    Set Sort = Init(arr)
End Function

'******************************************************************************
'* [概  要] Uniq
'* [詳  細] 一次元配列の重複要素を排除する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @return ArrayEx
'******************************************************************************
Public Function Uniq() As ArrayEx
    Set Uniq = Core.Init(New ArrayEx, Core.ArrUniq(Me.ToArray))
End Function

'******************************************************************************
'* [概  要] Concat
'* [詳  細] 一次元配列と合体（マージ）して返却する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param arr1 As Variant(Of Array(Of T))
'* @param arr2 As Variant(Of Array(Of T))
'* @return ArrayEx
'******************************************************************************
Public Function Concat(ByRef arr2 As Variant) As ArrayEx
    If ValidateUtils.IsArrayEx(arr2) Then
        Set Concat = Core.Init(New ArrayEx, Core.ArrConcat(Me.ToArray, arr2.ToArray))
    ElseIf IsArray(arr2) Then
        Set Concat = Core.Init(New ArrayEx, Core.ArrConcat(Me.ToArray, arr2))
    Else
        Set Concat = Core.Init(New ArrayEx, Core.ArrConcat(Me.ToArray, Array(arr2)))
    End If
End Function

'******************************************************************************
'* [概  要] Slice
'* [詳  細] 一次元配列をスライス（一部を配列として切り出）して返却する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param ixStart As Variant(Of Long)
'* @param ixEnd As Variant(Of Long)
'* @return ArrayEx
'******************************************************************************
Public Function Slice(Optional ByVal ixStart As Variant, Optional ByVal ixEnd As Variant) As ArrayEx
    Set Slice = Core.Init(New ArrayEx, Core.ArrSlice(Me.ToArray, ixStart, ixEnd))
End Function

'******************************************************************************
'* [概  要] FromJagArray
'* [詳  細] ジャグ配列のフラット化。ジャグ配列を一次元配列(ArrayEx)に変換して
'*          返却する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param jagArray As Variant(Of Array(Of Array(Of T)))
'* @return ArrayEx(Of Array(Of T))
'******************************************************************************
Public Function FromJagArray(ByVal jagArr As Variant) As ArrayEx
    Dim arr: arr = Core.ArrFlatten(jagArr)
    Set FromJagArray = Init(arr)
End Function

'******************************************************************************
'* [概  要] ToCollection
'* [詳  細] 一次元配列をCollectionオブジェクトに変換する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @return As Collection(Of T)
'******************************************************************************
Public Function ToCollection() As Collection
    Set ToCollection = Core.ArrToClct(Me.ToArray)
End Function

'******************************************************************************
'* [概  要] FromCollection
'* [詳  細] Collectionオブジェクトを一次元配列(ArrayEx)に変換する。
'*          AriawaseのCoreモジュールラッパーメソッド。
'*
'* @param clct As Collection(Of T)
'* @return ArrayEx(Of Array(Of T))
'******************************************************************************
Public Function FromCollection(ByVal clct As Collection) As ArrayEx
    Dim arr: arr = Core.ClctToArr(clct)
    Set FromCollection = Init(arr)
End Function

'******************************************************************************
'* [概  要] FromEnumerator
'* [詳  細] 繰り返し処理（For Each）可能なオブジェクトを一次元配列(ArrayEx)に
'*          変換する。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param enumr As Enumerator(Of T)
'* @return ArrayEx(Of Array(Of T))
'******************************************************************************
Public Function FromEnumerator(ByVal enumr As Object) As ArrayEx
    Dim arr: arr = ext.EnumeratorToArr(enumr)
    Set FromEnumerator = Init(arr)
End Function

'******************************************************************************
'* [概  要] Range
'* [詳  細] 指定した範囲の数値の配列を作成する。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fromVal As Variant(Of T)
'* @param toVal As Variant(Of T)
'* @param stepVal As Variant(Of T)
'* @return ArrayEx(Of Array(Of T))
'******************************************************************************
Public Function Range(ByVal fromVal As Variant, ByVal toVal As Variant, Optional ByVal stepVal As Variant = 1) As ArrayEx
    Dim arr: arr = ext.ArrRange(fromVal, toVal, stepVal)
    Set Range = Init(arr)
End Function

'******************************************************************************
'* [概  要] Map
'* [詳  細] 配列の全要素に対して指定したファンクションの処理を適用する。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func(Of T, U)
'* @return ArrayEx(Of Array(Of U))
'******************************************************************************
Public Function Map(ByVal fun As Func) As ArrayEx
    Dim arr: arr = ext.ArrMap(fun, Me.ToArray)
    Set Map = Init(arr)
End Function

'******************************************************************************
'* [概  要] Zip
'* [詳  細] ２つの配列に対して指定したファンクションの処理を適用し
'*          １つの配列を返す。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func(Of T, U, R)
'* @param arr2 As Variant(Of Array(Of U))
'* @return ArrayEx(Of Array(Of R))
'******************************************************************************
Public Function Zip(ByVal fun As Func, ByVal arr2 As Variant) As ArrayEx
    Dim arr: arr = ext.ArrZip(fun, Me.ToArray, arr2)
    Set Zip = Init(arr)
End Function

'******************************************************************************
'* [概  要] Filter
'* [詳  細] 配列に対して指定したフィルタ処理用のファンクションを適用し、
'*          条件に一致する要素のみの配列を返す。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func(Of T, Boolean)
'* @return ArrayEx(Of Array(Of T))
'******************************************************************************
Public Function Filter(ByVal fun As Func) As ArrayEx
    Dim arr: arr = ext.ArrFilter(fun, Me.ToArray)
    Set Filter = Init(arr)
End Function

'******************************************************************************
'* [概  要] GroupBy
'* [詳  細] 配列に対して指定したグルーピング処理用のファンクションを適用し、
'*          結果を返す。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func(Of T, K)
'* @return ArrayEx(Of Array(Of Tuple`2(Of K, T)))
'******************************************************************************
Public Function GroupBy(ByVal fun As Func) As ArrayEx
    Dim arr: arr = ext.ArrGroupBy(fun, Me.ToArray)
    Set GroupBy = Init(arr)
End Function

'******************************************************************************
'* [概  要] Fold
'* [詳  細] 畳み込み関数。配列内の要素に対して指定した演算用ファンクション
'*          を適用し、その結果と次の要素にまたファンクションを適用する
'*          といった処理を繰り返し最終的な結果を返す。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func(Of U, T, U)
'* @param seedv As Variant(Of U)
'* @return Variant(Of U)
'******************************************************************************
Public Function Fold(ByVal fun As Func, Optional ByVal seedv As Variant) As Variant
    Fold = ext.ArrFold(fun, Me.ToArray, seedv)
End Function

'******************************************************************************
'* [概  要] Scan
'* [詳  細] 配列の各要素に対して繰り返し指定した演算用ファンクションを適用し、
'*          累積した計算結果のリストを返す。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func(Of U, T, U)
'* @param seedv As Variant(Of U)
'* @return ArrayEx(Of Array(Of U))
'******************************************************************************
Public Function Scan(ByVal fun As Func, Optional ByVal seedv As Variant) As ArrayEx
    Dim arr: arr = ext.ArrScan(fun, Me.ToArray, seedv)
    Set Scan = Init(arr)
End Function

'******************************************************************************
'* [概  要] Unfold
'* [詳  細] Foldの逆。初期値から指定した演算用ファンクションを適用しリストを
'*          生成する。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func
'* @param seedv As Variant(Of T)
'* @return ArrayEx(Of Array(Of U))
'******************************************************************************
Public Function Unfold(ByVal fun As Func, ByVal seedv As Variant) As ArrayEx
    Dim arr: arr = ext.ArrUnfold(fun, seedv)
    Set Unfold = Init(arr)
End Function

'******************************************************************************
'* [概  要] CopyTo2DArrayRow
'* [詳  細] 2次元配列の指定行に対して、1次元配列の値を設定する。
'*
'* @param v2DArr 2次元配列
'* @param rowNum 指定行
'* @return Variant 変更後2次元配列
'******************************************************************************
Public Function CopyTo2DArrayRow(ByRef v2DArr As Variant, ByVal rowNum As Long) As Variant
    CopyTo2DArrayRow = ArrayUtils.Copy1DArrayTo2DArrayRow(v2DArr, rowNum, Me.ToArray)
End Function

'******************************************************************************
'* [概  要] Contains
'* [詳  細] 指定した項目が含まれるかチェックする。
'*
'* @param itm 項目
'* @return Boolean チェック結果（True：含む）
'******************************************************************************
Public Function Contains(ByVal itm As Variant) As Boolean
    Contains = ArrayUtils.Contains(Me.ToArray, itm)
End Function

'******************************************************************************
'* [概  要] ContainsLike
'* [詳  細] 指定した検索文字列（正規表現）が含まれるかチェックする。
'*
'* @param searchStr 検索文字列（正規表現）
'* @param igCase 大文字小文字の区別
'* @return Boolean チェック結果（True：含む）
'******************************************************************************
Public Function ContainsLike(ByVal searchStr As String, Optional ByVal igCase As Boolean = False) As Boolean
    ContainsLike = ArrayUtils.ContainsLike(Me.ToArray, searchStr, igCase)
End Function

'******************************************************************************
'* [概  要] CountOf
'* [詳  細] 配列に指定した項目が含まれる数をカウントする。
'*
'* @param itm 項目
'* @return Long カウント結果
'******************************************************************************
Public Function CountOf(ByVal itm As Variant) As Long
    CountOf = ArrayUtils.CountOf(Me.ToArray, itm)
End Function

'******************************************************************************
'* [概  要] Clone
'* [詳  細] ArrayExオブジェクトを複製する。
'*
'* @return ArrayEx 新しいArrayEx
'******************************************************************************
Public Function Clone() As ArrayEx
    Dim ax As New ArrayEx
    Set Clone = ax.AddAll(Me.ToArray)
End Function

'******************************************************************************
'* [概  要] Insert
'* [詳  細] 項目を指定位置に追加する。
'*
'* @param itm 項目
'* @return ArrayEx
'******************************************************************************
Public Function Insert(ByVal idx As Long, ByVal itm As Variant) As ArrayEx
    Set Insert = Init(ArrayUtils.Insert(Me.ToArray, idx, itm))
End Function

'******************************************************************************
'* [概  要] InsertAll
'* [詳  細] 項目（複数）を指定位置に追加する。
'*
'* @param itms 項目
'* @return ArrayEx
'******************************************************************************
Public Function InsertAll(ByVal idx As Long, ParamArray itms() As Variant) As ArrayEx
    Dim tmp: tmp = itms
    Set InsertAll = Init(ArrayUtils.InsertAll(Me.ToArray, idx, tmp))
End Function

'******************************************************************************
'* [概  要] LastIndexOf
'* [詳  細] 指定した値と一致する要素のインデックスを返却する。
'*          最後尾から検索して最初に一致したものを返却。
'*          値の捜査開始位置と要素数を任意指定可能。
'*
'* @param val As Variant(Of T)
'* @param ixStart As Variant(Of Long)
'* @param cnt As Variant(Of Long)
'* @return As Long
'******************************************************************************
Public Function LastIndexOf(ByVal val As Variant, _
    Optional ByVal ixStart As Variant, Optional ByVal cnt As Variant _
    ) As Long
    LastIndexOf = ArrayUtils.LastIndexOf(Me.ToArray, val, ixStart, cnt)
End Function

'******************************************************************************
'* [概  要] IndexListOf
'* [詳  細] 1次元配列から指定した値と一致する要素のインデックスリストを返却する。
'*
'* @param val 値
'* @param reverseOrder リストの並びを逆順にするか
'* @return Variant インデックスリスト（1次元配列）
'******************************************************************************
Public Function IndexListOf(ByVal val As Variant, Optional ByVal reverseOrder As Boolean = False) As Variant
    IndexListOf = ArrayUtils.IndexListOf(Me.ToArray, val, reverseOrder)
End Function

'******************************************************************************
'* [概  要] Remove
'* [詳  細] 1次元配列の最初に見つかった特定の要素を削除する。
'*
'* @param itm 項目
'* @return ArrayEx 削除後1次元配列
'******************************************************************************
Public Function Remove(ByVal itm As Variant) As ArrayEx
    Set Remove = Init(ArrayUtils.Remove(Me.ToArray, itm))
End Function

'******************************************************************************
'* [概  要] RemoveAt
'* [詳  細] 指定したインデックスにある要素を削除する。
'*
'* @param idx インデックス
'* @return ArrayEx 削除後1次元配列
'******************************************************************************
Public Function RemoveAt(ByVal idx As Long) As ArrayEx
    Set RemoveAt = Init(ArrayUtils.RemoveAt(Me.ToArray, idx))
End Function

'******************************************************************************
'* [概  要] RemoveFirst
'* [詳  細] 先頭の要素を削除する。
'*
'* @return ArrayEx 削除後1次元配列
'******************************************************************************
Public Function RemoveFirst() As ArrayEx
    Set RemoveFirst = Init(ArrayUtils.RemoveAt(Me.ToArray, 0))
End Function

'******************************************************************************
'* [概  要] RemoveLast
'* [詳  細] 最後の要素を削除する。
'*
'* @return ArrayEx 削除後1次元配列
'******************************************************************************
Public Function RemoveLast() As ArrayEx
    Set RemoveLast = Init(ArrayUtils.RemoveAt(Me.ToArray, xIndex))
End Function

'******************************************************************************
'* [概  要] RemoveRange
'* [詳  細] 1次元配列の要素の範囲を削除する。
'*
'* @param idx1 開始インデックス
'* @param idx2 終了インデックス
'* @return ArrayEx 削除後1次元配列
'******************************************************************************
Public Function RemoveRange(ByVal idx1 As Long, ByVal idx2 As Long) As ArrayEx
    Set RemoveRange = Init(ArrayUtils.RemoveRange(Me.ToArray, idx1, idx2))
End Function

'******************************************************************************
'* [概  要] Repeat
'* [詳  細] 指定した要素を指定回数分追加した1次元配列を作成する。
'*
'* @param itm 項目
'* @param count 項目
'* @return Variant 1次元配列
'******************************************************************************
Public Function Repeat(ByVal itm As Variant, ByVal Count As Long) As ArrayEx
    Set Repeat = Init(ArrayUtils.Repeat(itm, Count))
End Function

'******************************************************************************
'* [概  要] GetFirst
'* [詳  細] 最初の要素を取得する。
'*
'* @return Variant 最初の要素
'******************************************************************************
Public Function GetFirst() As Variant
    If IsEmptyArray Then Exit Function
    If IsObject(xItems(0)) Then
        Set GetFirst = xItems(0)
    Else
        GetFirst = xItems(0)
    End If
End Function

'******************************************************************************
'* [概  要] GetLast
'* [詳  細] 最後の要素を取得する。
'*
'* @return Variant 最後の要素
'******************************************************************************
Public Function GetLast(ByRef vArr As Variant) As Variant
    If IsEmptyArray Then Exit Function
    If IsObject(xItems(xIndex)) Then
        Set GetLast = xItems(xIndex)
    Else
        GetLast = xItems(xIndex)
    End If
End Function

'******************************************************************************
'* [概  要] Join
'* [詳  細] 配列を連結して文字列にする。
'*
'* @param delimiter デリミタ
'* @param wrapChar 囲み文字
'* @param rmLineSeparator 改行を除去するか
'* @return String 連結文字列
'******************************************************************************
Public Function Join(Optional ByVal Delimiter As String, Optional ByVal wrapChar As String, Optional ByVal rmLineSeparator As Boolean = True) As String
    Join = ArrayUtils.Join(Me.ToArray, Delimiter, wrapChar, rmLineSeparator)
End Function

'******************************************************************************
'* [概  要] Search
'* [詳  細] 配列（文字列）から指定した文字列と部分一致する項目を
'*          取得し、1次元配列で返却する。
'*
'* @param searchVal 検索文字列（配列で複数指定可）
'* @return Variant 検索結果（1次元配列）
'******************************************************************************
Public Function Search(ByVal searchVal As Variant) As ArrayEx
    Set Search = Core.Init(New ArrayEx, ArrayUtils.Search(Me.ToArray, searchVal))
End Function

'******************************************************************************
'* [概  要] RegexSearch
'* [詳  細] 配列（文字列）から指定した文字列と正規表現が一致する項目を
'*          取得し、1次元配列で返却する。
'*
'* @param searchStr 検索文字列（正規表現）
'* @param igCase 大文字小文字の区別
'* @return Variant 検索結果（1次元配列）
'******************************************************************************
Public Function RegexSearch(ByVal searchStr As String, Optional ByVal igCase As Boolean = False) As ArrayEx
    Set RegexSearch = Core.Init(New ArrayEx, ArrayUtils.RegexSearch(Me.ToArray, searchStr, igCase))
End Function

'******************************************************************************
'* [概  要] ToString
'* [詳  細] JSON文字列に変換する。
'*
'* @return String JSON文字列
'******************************************************************************
Public Function ToString() As String
    ToString = ArrayUtils.ToString(Me.ToArray)
End Function

'******************************************************************************
'* [概  要] To2DArray
'* [詳  細] 2次元配列に変換する。
'*
'* @param transpose Trueにすると1次元配列を列方向に設定する
'* @return Variant 2次元配列
'******************************************************************************
Public Function To2DArray(Optional ByVal Transpose As Boolean = False) As Variant
    To2DArray = ArrayUtils.ArrayTo2DArray(Me, Transpose)
End Function

'*******************************************************************************
'* [概  要] MaxTxtLength
'* [詳  細] 配列の中の項目で最大の文字数を返す。
'*
'* @return Long 最大文字数
'*******************************************************************************
Public Function MaxTxtLength(Optional useZenkaku2Width As Boolean = False) As Long
    MaxTxtLength = ArrayUtils.ArrayMaxTxtLength(Me.ToArray, useZenkaku2Width)
End Function

'******************************************************************************
'* [概  要] RemoveEmpty
'* [詳  細] 配列の空データを除去する。
'*
'* @return ArrayEx 変更後配列
'******************************************************************************
Public Function RemoveEmpty() As ArrayEx
    Set RemoveEmpty = Core.Init(New ArrayEx, ArrayUtils.RemoveEmpty(Me.ToArray))
End Function

'******************************************************************************
'* [概  要] NaturalSort
'* [詳  細] エクスプローラのファイル順のように自然順ソートで並び替える。
'*
'******************************************************************************
Public Sub NaturalSort()
    Dim arr: arr = Me.ToArray
    Call ArrayUtils.NaturalSort(arr)
    Call Init(arr)
End Sub

