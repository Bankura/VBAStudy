VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Array2DEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] 2次元配列拡張クラス
'* [詳  細] 2次元配列操作を行うクラス。
'*
'* [依　存] <内部依存>
'*           Base, ValidateUtils, ArrayUtils
'*          <外部依存>
'*           Ariawase Core, Ext, ArrayEx
'*
'* @author Bankura
'* Copyright (c) 2020-2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private m2DItems As ArrayEx

Private mRowLength As Long
Private mColLength As Long
Private mRowMaxIndex As Long
Private mColMaxIndex As Long
Private mDefaultInitValue As Variant

'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* RowLength プロパティ
'*-----------------------------------------------------------------------------
Public Property Get RowLength() As Long
    RowLength = mRowLength
End Property
'*-----------------------------------------------------------------------------
'* ColLength プロパティ
'*-----------------------------------------------------------------------------
Public Property Get ColLength() As Long
    ColLength = mColLength
End Property
'*-----------------------------------------------------------------------------
'* RowMaxIndex プロパティ
'*-----------------------------------------------------------------------------
Public Property Get RowMaxIndex() As Long
    RowMaxIndex = mRowMaxIndex
End Property
'*-----------------------------------------------------------------------------
'* ColMaxIndex プロパティ
'*-----------------------------------------------------------------------------
Public Property Get ColMaxIndex() As Long
    ColMaxIndex = mColMaxIndex
End Property

'*-----------------------------------------------------------------------------
'* Item プロパティ
'*-----------------------------------------------------------------------------
Public Property Get Item(ByVal x As Long, ByVal y As Long) As Variant
Attribute Item.VB_UserMemId = 0
    If IsInvalidIndex(x, y) Then
        Item = mDefaultInitValue
        Exit Property
    End If
    If IsObject(m2DItems(x)(y)) Then
        Set Item = m2DItems(x)(y)
    Else
        Item = m2DItems(x)(y)
    End If
End Property
Public Property Let Item(ByVal x As Long, ByVal y As Long, ByVal arg As Variant)
    If IsInvalidIndex(x, y) Then
        Exit Property
    End If
    
    If IsObject(arg) Then
        Set m2DItems(x)(y) = arg
    Else
        m2DItems(x)(y) = arg
    End If
End Property

'*-----------------------------------------------------------------------------
'* XRow プロパティ
'*-----------------------------------------------------------------------------
Public Property Get XRow(ByVal x As Long) As ArrayEx
    If IsInvalidRowIndex(x) Then
        Exit Property
    End If
    Set XRow = Core.Init(New ArrayEx, m2DItems(x).ToArray)
End Property
Public Property Let XRow(ByVal x As Long, ByVal arg As ArrayEx)
    If IsInvalidRowIndex(x) Then
        Exit Property
    End If
    If mColLength < arg.Length Then
        Call ExpandColumns(arg.LastIndex - mColMaxIndex)
    End If
    Set m2DItems(x) = arg
End Property

'*-----------------------------------------------------------------------------
'* YColumn プロパティ
'*-----------------------------------------------------------------------------
Public Property Get YColumn(ByVal y As Long) As ArrayEx
    If IsInvalidColIndex(y) Then
        Set YColumn = Empty
        Exit Property
    End If
    Set YColumn = m2DItems(y)
    Dim x As Long, col As ArrayEx: Set col = New ArrayEx
    For x = 0 To mRowMaxIndex
        col.Add m2DItems(x)(y)
    Next
    Set YColumn = col
End Property
Public Property Let YColumn(ByVal y As Long, ByVal arg As ArrayEx)
    If IsInvalidColIndex(y) Then
        Exit Property
    End If
    If mRowLength < arg.Length Then
        Call ExpandRows(arg.LastIndex - mRowMaxIndex)
    End If
    Dim x As Long
    For x = 0 To mRowMaxIndex
        If IsObject(arg) Then
            Set m2DItems(x)(y) = arg(x)
        Else
            m2DItems(x)(y) = arg(x)
        End If
    Next
End Property

'*-----------------------------------------------------------------------------
'* DefaultInitValue プロパティ
'*-----------------------------------------------------------------------------
Public Property Get DefaultInitValue() As Variant
    DefaultInitValue = mDefaultInitValue
End Property
Public Property Let DefaultInitValue(ByVal arg)
    If IsObject(arg) Then
        Set mDefaultInitValue = arg
    Else
        mDefaultInitValue = arg
    End If
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Call Clear
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Init メソッド
'* [詳  細] 初期化用メソッド｡
'*
'* @param val 初期配列
'* @return ArrayExオブジェクト
'*
'******************************************************************************
Public Function Init(Optional ByVal val, Optional ByVal transpose As Boolean = False) As Array2DEx
    If VBA.IsMissing(val) Then
        Call Clear
        Set Init = Me
        Exit Function
    End If
    If ArrayUtils.Is2DArray(val) Then
        mRowMaxIndex = UBound(val, 1) - LBound(val, 1)
        mColMaxIndex = UBound(val, 2) - LBound(val, 2)
        mRowLength = UBound(val, 1) - LBound(val, 1) + 1
        mColLength = UBound(val, 2) - LBound(val, 2) + 1
        Set m2DItems = New ArrayEx
        Dim i As Long, j As Long
        For i = LBound(val, 1) To UBound(val, 1)
            Dim row As ArrayEx
            Set row = New ArrayEx
            For j = LBound(val, 2) To UBound(val, 2)
                row.Add val(i, j)
            Next
            m2DItems.Add row
        Next
    ElseIf ValidateUtils.IsArray2DEx(val) Then
        Set m2DItems = val.ToArrayExOfArrayEx
        mRowMaxIndex = val.RowMaxIndex
        mColMaxIndex = val.ColMaxIndex
        mRowLength = val.RowLength
        mColLength = val.ColLength
    ElseIf ValidateUtils.IsArrayExOfArrayEx(val) Then
        If val.IsEmptyArray Then
            Call Clear
        Else
            Set m2DItems = val
            mRowMaxIndex = val.LastIndex
            mRowLength = val.Length
            mColMaxIndex = val(0).LastIndex
            mColLength = val(0).Length
        End If
    ElseIf ArrayUtils.IsJagArray(val) Then
        Call Clear
        Dim v
        For Each v In val
            Me.AddRow v
        Next
    ElseIf ArrayUtils.Is1DArray(val) Or ValidateUtils.IsArrayEx(val) Then
        Call Clear
        If transpose Then
            Call Me.AddColumn(val)
        Else
            Call Me.AddRow(val)
        End If
    Else
        Call Clear
        Call Me.AddRow(ArrayUtils.CreateArrayEx().Add(val))
    End If
    Set Init = Me
End Function

'******************************************************************************
'* [概  要] Clear
'* [詳  細] 初期化する｡
'*
'* @return Array2DExオブジェクト
'*
'******************************************************************************
Public Function Clear() As Array2DEx
    mRowLength = 0
    mColLength = 0
    mRowMaxIndex = -1
    mColMaxIndex = -1
    mDefaultInitValue = Empty
    Set m2DItems = New ArrayEx
    Set Clear = Me
End Function

Private Function IsInvalidIndex(ByVal x As Long, ByVal y As Long) As Boolean
    IsInvalidIndex = IsInvalidRowIndex(x) Or IsInvalidColIndex(y)
End Function
Private Function IsInvalidRowIndex(ByVal idx As Long) As Boolean
    IsInvalidRowIndex = (idx < 0 Or mRowMaxIndex < idx)
End Function
Private Function IsInvalidColIndex(ByVal idx As Long) As Boolean
    IsInvalidColIndex = (idx < 0 Or mColMaxIndex < idx)
End Function
Private Function IsInvalidRowIndexes(ParamArray idxs() As Variant) As Boolean
    If IsMissing(idxs) Then
        IsInvalidRowIndexes = True
        Exit Function
    End If
    Dim v
    For Each v In idxs
        If IsInvalidRowIndex(CLng(v)) Then
            IsInvalidRowIndexes = True
            Exit Function
        End If
    Next
    IsInvalidRowIndexes = False
End Function
Private Function IsInvalidColIndexes(ParamArray idxs() As Variant) As Boolean
    If IsMissing(idxs) Then
        IsInvalidColIndexes = True
        Exit Function
    End If
    Dim v
    For Each v In idxs
        If IsInvalidColIndex(CLng(v)) Then
            IsInvalidColIndexes = True
            Exit Function
        End If
    Next
    IsInvalidColIndexes = False
End Function

'******************************************************************************
'* [概  要] To2DArray
'* [詳  細] 2次元配列に変換する｡
'*
'* @return Variant 2次元配列
'*
'******************************************************************************
Public Function To2DArray() As Variant
    If mRowMaxIndex > -1 Then
        To2DArray = ArrayUtils.Array2DExTo2DArray(Me)
    Else
        To2DArray = ArrayUtils.EmptyVariant2DArray
    End If
End Function

'******************************************************************************
'* [概  要] ToArrayExOfArrayEx
'* [詳  細] ArrayExを要素とするArrayExの形式でデータを返却する｡
'*
'* @return ArrayEx(Of ArrayEx)
'*
'******************************************************************************
Public Function ToArrayExOfArrayEx() As ArrayEx
    Set ToArrayExOfArrayEx = m2DItems
End Function

'******************************************************************************
'* [概  要] ToArrayOfArrayEx
'* [詳  細] ArrayExを要素とするVariant配列の形式でデータを返却する｡
'*
'* @return Array(Of ArrayEx)
'*
'******************************************************************************
Public Function ToArrayOfArrayEx() As Variant
    ToArrayOfArrayEx = m2DItems.ToArray
End Function

'******************************************************************************
'* [概  要] ToJagArray
'* [詳  細] JagArrayの形式でデータを返却する｡
'*
'* @return Array(Of Array)
'*
'******************************************************************************
Public Function ToJagArray() As Variant
    Dim arrex As ArrayEx: Set arrex = New ArrayEx
    Dim v, var
    var = m2DItems.ToArray
    For Each v In var
        arrex.Add v.ToArray
    Next
    ToJagArray = arrex.ToArray
End Function

'******************************************************************************
'* [概  要] ExpandRows
'* [詳  細] 行を追加する｡
'*          追加した行の値はDefaultInitValueの値を設定する。
'*
'* @param xAddRowCount 追加行数
'*
'******************************************************************************
Public Sub ExpandRows(ByVal xAddRowCount As Long)
    Dim i As Long, j As Long
    For i = 1 To xAddRowCount
        Dim row As ArrayEx: Set row = New ArrayEx
        For j = 0 To mColMaxIndex
            row.Add mDefaultInitValue
        Next
        m2DItems.Add row
    Next
    mRowLength = mRowLength + xAddRowCount
    mRowMaxIndex = mRowMaxIndex + xAddRowCount
End Sub

'******************************************************************************
'* [概  要] ExpandColumns
'* [詳  細] 列を追加する｡
'*          追加した列の値はDefaultInitValueの値を設定する。
'*
'* @param yAddColCount 追加列数
'*
'******************************************************************************
Public Sub ExpandColumns(ByVal yAddColCount As Long)
    Dim i As Long, j As Long
    For i = 0 To mRowMaxIndex
        Dim row As ArrayEx: Set row = m2DItems(i)
        For j = 1 To yAddColCount
            row.Add mDefaultInitValue
        Next
    Next
    mColLength = mColLength + yAddColCount
    mColMaxIndex = mColMaxIndex + yAddColCount
End Sub

'******************************************************************************
'* [概  要] Expand
'* [詳  細] 行列を追加する｡
'*          追加した行/列の値はDefaultInitValueの値を設定する。
'*
'* @param xAddRowCount 追加行数
'* @param yAddColCount 追加列数
'*
'******************************************************************************
Public Sub Expand(ByVal xAddRowCount As Long, ByVal yAddColCount As Long)
    Call ExpandRows(xAddRowCount)
    Call ExpandColumns(yAddColCount)
End Sub

'******************************************************************************
'* [概  要] AddRow
'* [詳  細] 行を追加する｡
'*
'* @param var 追加行（1次元Variant配列 or ArrayEx）
'* @return Array2DEx
'*
'******************************************************************************
Public Function AddRow(ByVal var As Variant) As Array2DEx
    Dim colLen As Long
    Dim arrex As ArrayEx
    If ArrayUtils.Is1DArray(var) Then
        colLen = ArrayUtils.GetLength(var)
        Set arrex = Core.Init(New ArrayEx, var)
    ElseIf ValidateUtils.IsArrayEx(var) Then
        colLen = var.Length
        Set arrex = var
    Else
        Err.raise 9999, "Array2DEx#AddRow", "Invalid argument."
    End If

    If mColLength < colLen Then
        Dim addCount As Long: addCount = colLen - mColLength
        Call ExpandColumns(addCount)
    End If
    
    Dim y As Long
    For y = arrex.LastIndex + 1 To mColMaxIndex
        arrex.Add mDefaultInitValue
    Next
    m2DItems.Add arrex
    mRowLength = mRowLength + 1
    mRowMaxIndex = mRowMaxIndex + 1
    Set AddRow = Me
End Function

'******************************************************************************
'* [概  要] AddRows
'* [詳  細] 行を複数追加する｡
'*
'* @param var 追加行（複数:2次元Variant配列 or Array2DEx）
'* @return Array2DEx
'*
'******************************************************************************
Public Function AddRows(ByVal var As Variant, ParamArray vars() As Variant) As Array2DEx
    Dim colLen As Long
    Dim arr2dex As Array2DEx
    Set arr2dex = Core.Init(New Array2DEx, var)
    colLen = arr2dex.ColLength

    If mColLength < colLen Then
        Dim addCount As Long: addCount = colLen - mColLength
        Call ExpandColumns(addCount)
    End If
    
    Dim x As Long
    For x = 0 To arr2dex.RowMaxIndex
        Dim y As Long, rowArrex As ArrayEx
        Set rowArrex = arr2dex.XRow(x)
        For y = arr2dex.ColMaxIndex + 1 To mColMaxIndex
            rowArrex.Add mDefaultInitValue
        Next
        m2DItems.Add rowArrex
    Next
    mRowLength = mRowLength + arr2dex.RowLength
    mRowMaxIndex = mRowMaxIndex + arr2dex.RowLength
    
    If Not VBA.IsMissing(vars) Then
        Dim v
        For Each v In vars
            Call AddRows(v)
        Next
    End If
    
    Set AddRows = Me
End Function

'******************************************************************************
'* [概  要] AddColumn
'* [詳  細] 列を追加する｡
'*
'* @param var 追加列（1次元Variant配列 or ArrayEx）
'* @return Array2DEx
'*
'******************************************************************************
Public Function AddColumn(ByVal var As Variant) As Array2DEx
    Dim rowLen As Long
    Dim arrex As ArrayEx
    If ArrayUtils.Is1DArray(var) Then
        rowLen = ArrayUtils.GetLength(var)
        Set arrex = Core.Init(New ArrayEx, var)
    ElseIf ValidateUtils.IsArrayEx(var) Then
        rowLen = var.Length
        Set arrex = var
    Else
        Err.raise 9999, "Array2DEx#AddColumn", "Invalid argument."
    End If

    If mRowLength < rowLen Then
        Dim addCount As Long: addCount = rowLen - mRowLength
        Call ExpandRows(addCount)
    End If
    
    Dim x As Long
    For x = 0 To mRowMaxIndex
        m2DItems(x).Add arrex(x)
    Next
    mColLength = mColLength + 1
    mColMaxIndex = mColMaxIndex + 1
    Set AddColumn = Me
End Function

'******************************************************************************
'* [概  要] AddColumns
'* [詳  細] 列を複数追加する｡
'*
'* @param var 追加列（複数:2次元Variant配列 or Array2DEx）
'* @return Array2DEx
'*
'******************************************************************************
Public Function AddColumns(ByVal var As Variant, ParamArray vars() As Variant) As Array2DEx
    Dim rowLen As Long
    Dim arr2dex As Array2DEx
    Set arr2dex = Core.Init(New Array2DEx, var, True)
    rowLen = arr2dex.RowLength
    
    If mRowLength < rowLen Then
        Dim addCount As Long: addCount = rowLen - mRowLength
        Call ExpandRows(addCount)
    End If
    
    Dim x As Long, y As Long
    For x = 0 To arr2dex.RowMaxIndex
        For y = 0 To arr2dex.ColMaxIndex
            m2DItems(x).Add arr2dex(x, y)
        Next
    Next
    For x = arr2dex.RowMaxIndex + 1 To mRowMaxIndex
        For y = 0 To arr2dex.ColMaxIndex
            m2DItems(x).Add mDefaultInitValue
        Next
    Next
    mColLength = mColLength + arr2dex.ColLength
    mColMaxIndex = mColMaxIndex + arr2dex.ColLength
    
    If Not VBA.IsMissing(vars) Then
        Dim v
        For Each v In vars
            Call AddColumns(v)
        Next
    End If
    
    Set AddColumns = Me
End Function

'******************************************************************************
'* [概  要] IsEmptyArray
'* [詳  細] 配列が未初期化・空配列か判定する｡
'*
'* @return Boolean 判定結果(True:未初期化・空)
'*
'******************************************************************************
Public Function IsEmptyArray() As Boolean
    IsEmptyArray = (mRowMaxIndex = -1 And mColMaxIndex = -1)
End Function

'******************************************************************************
'* [概  要] Equals
'* [詳  細] 2次元配列を比較し、イコールであればTrueを返却する。
'*
'* @param arr2 2次元Variant配列
'* @return Boolean 比較結果
'******************************************************************************
Public Function Equals(ByVal arr2 As Variant) As Boolean
    Dim rowLen As Long, colLen As Long
    
    If ArrayUtils.Is2DArray(arr2) Then
        rowLen = ArrayUtils.GetLength(arr2)
        colLen = ArrayUtils.GetLength(arr2, 2)
    ElseIf ValidateUtils.IsArray2DEx(arr2) Then
        rowLen = arr2.RowLength
        colLen = arr2.ColLength
    Else
        Err.raise 9999, "Array2DEx#Equals", "Invalid argument."
    End If
    
    If mRowLength <> rowLen Or mColLength <> colLen Then
        Equals = False
        Exit Function
    End If
    
    Dim arr2dex As Array2DEx
    If ArrayUtils.Is2DArray(arr2) Then
        Set arr2dex = Core.Init(New Array2DEx, arr2)
    ElseIf ValidateUtils.IsArray2DEx(arr2) Then
        Set arr2dex = arr2
    End If
    Dim x As Long, y As Long
    For x = 0 To mRowMaxIndex
        For y = 0 To mColMaxIndex
            Dim ret As Boolean
            ret = Core.Eq(m2DItems(x)(y), arr2dex(x, y))
            If ret <> True Then
                Equals = False
                Exit Function
            End If
        Next
    Next
    Equals = True
End Function

'******************************************************************************
'* [概  要] RowIndexOf
'* [詳  細] 2次元配列の指定した行について、指定した値と一致する要素の
'*          列インデックスを返却する。
'*          値の捜査開始位置と要素数を任意指定可能。
'*
'* @param rowidx 行インデックス
'* @param val 値
'* @param ixStart 捜査開始位置
'* @param cnt 要素数
'* @return Long 列インデックス（一致なしの場合は-1）
'******************************************************************************
Public Function RowIndexOf(ByVal rowIdx As Variant, ByVal val As Variant, Optional ByVal ixStart As Long = 0, Optional ByVal cnt As Variant) As Long
    If IsInvalidRowIndex(rowIdx) Then
        Err.raise 9999, "Array2DEx#RowIndexOf", "Invalid argument."
    End If
    RowIndexOf = Core.ArrIndexOf(m2DItems(rowIdx).ToArray, val, ixStart, cnt)
    If RowIndexOf < ixStart Then RowIndexOf = -1
End Function

'******************************************************************************
'* [概  要] ColIndexOf
'* [詳  細] 2次元配列の指定した列について、指定した値と一致する要素の
'*          行インデックスを返却する。
'*          値の捜査開始位置と要素数を任意指定可能。
'*
'* @param colidx 列インデックス
'* @param val 値
'* @param ixStart 捜査開始位置
'* @param cnt 要素数
'* @return Long 行インデックス（一致なしの場合は-1）
'******************************************************************************
Public Function ColIndexOf(ByVal colidx As Variant, ByVal val As Variant, Optional ByVal ixStart As Long = 0, Optional ByVal cnt As Variant) As Long
    If IsInvalidColIndex(colidx) Then
        Err.raise 9999, "Array2DEx#ColIndexOf", "Invalid argument."
    End If
    ColIndexOf = Core.ArrIndexOf(YColumn(colidx).ToArray, val, ixStart, cnt)
    If ColIndexOf < ixStart Then ColIndexOf = -1
End Function

'******************************************************************************
'* [概  要] IndexOf
'* [詳  細] 2次元配列について、指定した値と一致する要素の
'*          行列インデックスを返却する。
'*
'* @param val 値
'* @param ixStart 捜査開始行インデックス
'* @return Array2DIndex 行列インデックス（一致なしの場合は行列ともに-1）
'******************************************************************************
Public Function IndexOf(ByVal val As Variant, Optional ByVal ixStart As Long) As Array2DIndex
    If IsInvalidRowIndex(ixStart) Then
        Err.raise 9999, "Array2DEx#IndexOf", "Invalid argument."
    End If

    Dim x As Long, y As Long, arr2dIdx As Array2DIndex
    For x = ixStart To mRowMaxIndex
        y = Core.ArrIndexOf(m2DItems(x).ToArray, val)
        If y <> -1 Then
            arr2dIdx.x = x
            arr2dIdx.y = y
            IndexOf = arr2dIdx
            Exit Function
        End If
    Next
    arr2dIdx.x = -1
    arr2dIdx.y = -1
    IndexOf = arr2dIdx
End Function

'******************************************************************************
'* [概  要] EntireRowIndexOf
'* [詳  細] 2次元配列について、指定した行データと一致する行の
'*          行インデックスを返却する。
'*          値の捜査開始位置と要素数を任意指定可能。
'*
'* @param rowData 行（1次元配列 Or ArrayEx）
'* @param ixStart 捜査開始行インデックス
'* @return Long 行インデックス（一致なしの場合は-1）
'******************************************************************************
Public Function EntireRowIndexOf(ByVal rowData As Variant, Optional ByVal ixStart As Long = 0) As Long
    If IsInvalidRowIndex(ixStart) Then
        Err.raise 9999, "Array2DEx#RowsIndexOf", "Invalid argument."
    End If
    
    Dim arr As Variant
    If ArrayUtils.Is1DArray(rowData) Then
        arr = rowData
    ElseIf ValidateUtils.IsArrayEx(rowData) Then
        arr = rowData.ToArray
    Else
        Err.raise 9999, "Array2DEx#RowsIndexOf", "Invalid argument."
    End If
    
    Dim x As Long
    For x = ixStart To mRowMaxIndex
        If m2DItems(x).Equals(arr) Then
            EntireRowIndexOf = x
            Exit Function
        End If
    Next
    EntireRowIndexOf = -1
End Function

'******************************************************************************
'* [概  要] ReverseRow
'* [詳  細] 2次元配列の行の並び順を反転させる。
'*
'******************************************************************************
Public Sub ReverseRow()
    If mRowMaxIndex = -1 Then Exit Sub
    Dim ixL As Long: ixL = 0
    Dim ixU As Long: ixU = mRowMaxIndex

    Dim tmprow As ArrayEx
    While ixL < ixU
        Set tmprow = m2DItems(ixL): m2DItems(ixL) = m2DItems(ixU): m2DItems(ixU) = tmprow
        ixL = ixL + 1: ixU = ixU - 1
    Wend
End Sub

'******************************************************************************
'* [概  要] ReverseColumns
'* [詳  細] 2次元配列の列の並び順を反転させる。
'*
'******************************************************************************
Public Sub ReverseColumns()
    If mColMaxIndex = -1 Then Exit Sub
    Dim x As Long: x = 0
    Dim tmprow As ArrayEx
    For x = 0 To mRowMaxIndex
        m2DItems(x).Reverse
    Next
End Sub

'******************************************************************************
'* [概  要] Reverse
'* [詳  細] 2次元配列の行列の並び順を反転させる。
'*
'******************************************************************************
Public Sub Reverse()
    Call ReverseRow
    Call ReverseColumns
End Sub

'******************************************************************************
'* [概  要] Sort
'* [詳  細] 指定した列を基準に2次元配列のソートを行う。デフォルトは昇順。
'*
'* @param colIdx 列インデックス
'* @param orderAsc 昇順ソートかどうか
'******************************************************************************
Public Sub Sort(ByVal colidx As Long, Optional ByVal orderAsc As Boolean = True)
    If IsInvalidColIndex(colidx) Then
        Err.raise 9999, "Array2DEx#Sort", "Invalid argument."
    End If
    Dim arr: arr = YColumn(colidx).ToArray
    Dim idxArr: idxArr = ArrSort(arr, orderAsc)
    
    Dim arrex As ArrayEx: Set arrex = New ArrayEx
    Call arrex.Expand(mRowLength)
    Dim x As Long
    For x = 0 To mRowMaxIndex
        arrex(x) = m2DItems(idxArr(x))
    Next
    Set m2DItems = arrex
End Sub

''' @param arr As Variant(Of Array(Of T))
''' @param orderAsc As Boolean
Private Function ArrSort(ByRef arr As Variant, ByVal orderAsc As Boolean) As Variant

    If Not IsArray(arr) Then Err.raise 13
    If ArrLen(arr) <= 1 Then GoTo Escape

    Dim ix0 As Long: ix0 = LBound(arr)
    
    Dim idxArr: idxArr = ArrayUtils.Range(ix0, UBound(arr))
    If IsObject(arr(ix0)) Then
        ObjArrMSort arr, ix0, idxArr, orderAsc
    Else
        ValArrMSort arr, ix0, idxArr, orderAsc
    End If
    ArrSort = idxArr
Escape:
End Function
Private Sub ObjArrMSort(arr As Variant, lb As Long, idxArr As Variant, orderAsc As Boolean)
    Dim aLen As Long: aLen = ArrLen(arr)
    If aLen <= 1 Then GoTo Escape
    
    '' optimization
    If aLen <= 8 Then
        ObjArrISort arr, lb, idxArr, orderAsc
        GoTo Escape
    End If
    
    Dim i As Long
    Dim l1 As Long: l1 = Fix(aLen / 2)
    Dim l2 As Long: l2 = aLen - l1
    
    Dim ub1 As Long:   ub1 = lb + l1 - 1
    Dim a1 As Variant: ReDim a1(lb To ub1)
    For i = lb To ub1: Set a1(i) = arr(i): Next
    Dim idx1 As Variant: ReDim idx1(lb To ub1)
    For i = lb To ub1: idx1(i) = idxArr(i): Next
    ObjArrMSort a1, lb, idx1, orderAsc
    
    Dim ub2 As Long:   ub2 = lb + l2 - 1
    Dim a2 As Variant: ReDim a2(lb To ub2)
    For i = lb To ub2: Set a2(i) = arr(l1 + i): Next
    Dim idx2 As Variant: ReDim idx2(lb To ub2)
    For i = lb To ub2: idx2(i) = idxArr(l1 + i): Next
    ObjArrMSort a2, lb, idx2, orderAsc
    
    Dim i1 As Long: i1 = lb
    Dim i2 As Long: i2 = lb
    While i1 <= ub1 Or i2 <= ub2
        If ArrMergeSw(a1, i1, ub1, a2, i2, ub2, orderAsc) Then
            Set arr(i1 + i2 - lb) = a1(i1)
            idxArr(i1 + i2 - lb) = idx1(IncrPst(i1))
        Else
            Set arr(i1 + i2 - lb) = a2(i2)
            idxArr(i1 + i2 - lb) = idx2(IncrPst(i2))
        End If
    Wend
    
Escape:
End Sub
Private Sub ValArrMSort(arr As Variant, lb As Long, idxArr As Variant, orderAsc As Boolean)
    Dim aLen As Long: aLen = ArrLen(arr)
    If aLen <= 1 Then GoTo Escape
    
    '' optimization
    If aLen <= 8 Then
        ValArrISort arr, lb, idxArr, orderAsc
        GoTo Escape
    End If
    
    Dim i As Long
    Dim l1 As Long: l1 = Fix(aLen / 2)
    Dim l2 As Long: l2 = aLen - l1
    
    Dim ub1 As Long:   ub1 = lb + l1 - 1
    Dim a1 As Variant: ReDim a1(lb To ub1)
    For i = lb To ub1: Let a1(i) = arr(i): Next
    Dim idx1 As Variant: ReDim idx1(lb To ub1)
    For i = lb To ub1: idx1(i) = idxArr(i): Next
    ValArrMSort a1, lb, idxArr, orderAsc
    
    Dim ub2 As Long:   ub2 = lb + l2 - 1
    Dim a2 As Variant: ReDim a2(lb To ub2)
    For i = lb To ub2: Let a2(i) = arr(l1 + i): Next
    Dim idx2 As Variant: ReDim idx2(lb To ub2)
    For i = lb To ub2: idx2(i) = idxArr(l1 + i): Next
    ValArrMSort a2, lb, idxArr, orderAsc
    
    Dim i1 As Long: i1 = lb
    Dim i2 As Long: i2 = lb
    While i1 <= ub1 Or i2 <= ub2
        If ArrMergeSw(a1, i1, ub1, a2, i2, ub2, orderAsc) Then
            Let arr(i1 + i2 - lb) = a1(i1)
            idxArr(i1 + i2 - lb) = idx1(IncrPst(i1))
        Else
            Let arr(i1 + i2 - lb) = a2(i2)
            idxArr(i1 + i2 - lb) = idx2(IncrPst(i2))
        End If
    Wend
    
Escape:
End Sub
Private Sub ObjArrISort(arr As Variant, lb As Long, idxArr As Variant, orderAsc As Boolean)
    Dim i As Long, j As Long, x As Variant, idx As Variant
    For i = lb + 1 To UBound(arr)
        j = i
        Do While j > lb
            If Core.Compare(arr(j - 1), arr(j), True) * (Abs(orderAsc) * 2 - 1) <= 0 Then Exit Do
            Set x = arr(j): Set arr(j) = arr(j - 1): Set arr(j - 1) = x
            idx = idxArr(j): idxArr(j) = idxArr(j - 1): idxArr(j - 1) = idx
            j = j - 1
        Loop
    Next
End Sub
Private Sub ValArrISort(arr As Variant, lb As Long, idxArr As Variant, orderAsc As Boolean)
    Dim i As Long, j As Long, x As Variant, idx As Variant
    For i = lb + 1 To UBound(arr)
        j = i
        Do While j > lb
            If Core.Compare(arr(j - 1), arr(j), True) * (Abs(orderAsc) * 2 - 1) <= 0 Then Exit Do
            Let x = arr(j): Let arr(j) = arr(j - 1): Let arr(j - 1) = x
            idx = idxArr(j): idxArr(j) = idxArr(j - 1): idxArr(j - 1) = idx
            j = j - 1
        Loop
    Next
End Sub
Private Function ArrMergeSw( _
    arr1 As Variant, i1 As Long, ub1 As Long, _
    arr2 As Variant, i2 As Long, ub2 As Long, _
    orderAsc As Boolean _
    ) As Boolean
    
    If i1 > ub1 Then ArrMergeSw = False Else _
    If i2 > ub2 Then ArrMergeSw = True Else _
    ArrMergeSw = Core.Compare(arr1(i1), arr2(i2), True) * (Abs(orderAsc) * 2 - 1) < 1
End Function

'******************************************************************************
'* [概  要] Uniq
'* [詳  細] 2次元配列の重複行を排除する。
'*
'* @return Array2DEx
'******************************************************************************
Public Function Uniq() As Array2DEx
    Dim arr2dex As Array2DEx: Set arr2dex = New Array2DEx
    Dim x As Long
    For x = 0 To mRowMaxIndex
        If x = mRowMaxIndex Then
            arr2dex.AddRow m2DItems(x).ToArray
        ElseIf EntireRowIndexOf(m2DItems(x), x + 1) = -1 Then
            arr2dex.AddRow m2DItems(x).ToArray
        End If
    Next
    Set Uniq = arr2dex
End Function

'******************************************************************************
'* [概  要] Concat
'* [詳  細] 2次元配列と合体（マージ）して返却する。
'*
'* @param arr2 As Variant(Of Array(Of T))
'* @return As Variant(Of Array(Of T))
'******************************************************************************
Public Function Concat(ByVal arr2 As Variant) As Array2DEx
    Dim arr2dex As Array2DEx: Set arr2dex = New Array2DEx
    arr2dex.AddRows Me.Clone, arr2
    Set Concat = arr2dex
End Function

'******************************************************************************
'* [概  要] RowSlice
'* [詳  細] 2次元配列の行をスライス（一部を配列として切り出す）して返却する。
'*
'* @param idxStart As Variant(Of Long)
'* @param idxEnd As Variant(Of Long)
'* @return Array2DEx
'******************************************************************************
Public Function RowSlice(Optional ByVal idxStart As Variant, Optional ByVal idxEnd As Variant) As Array2DEx
    If VBA.IsMissing(idxStart) Then
        idxStart = 0
    End If
    If VBA.IsMissing(idxEnd) Then
        idxEnd = mRowMaxIndex
    End If
    If IsInvalidRowIndex(idxStart) Or IsInvalidRowIndex(idxEnd) Then
        Err.raise 9999, "Array2DEx#RowSlice", "Invalid argument."
    End If
    
    Dim arr2dex As Array2DEx: Set arr2dex = New Array2DEx
    Dim x As Long, s As Long: s = 1
    If idxStart > idxEnd Then s = -1
    For x = idxStart To idxEnd Step s
        arr2dex.AddRow m2DItems(x).ToArray
    Next
    Set RowSlice = arr2dex
End Function

'******************************************************************************
'* [概  要] ColSlice
'* [詳  細] 2次元配列の列をスライス（一部を配列として切り出す）して返却する。
'*
'* @param idxStart As Variant(Of Long)
'* @param idxEnd As Variant(Of Long)
'* @return Array2DEx
'******************************************************************************
Public Function ColSlice(Optional ByVal idxStart As Variant, Optional ByVal idxEnd As Variant) As Array2DEx
    If VBA.IsMissing(idxStart) Then
        idxStart = 0
    End If
    If VBA.IsMissing(idxEnd) Then
        idxEnd = mColMaxIndex
    End If
    If IsInvalidColIndex(idxStart) Or IsInvalidColIndex(idxEnd) Then
        Err.raise 9999, "Array2DEx#ColSlice", "Invalid argument."
    End If
    
    Dim arr2dex As Array2DEx: Set arr2dex = New Array2DEx
    Dim x As Long
    For x = 0 To mRowMaxIndex
        Dim row As ArrayEx: Set row = m2DItems(x)
        arr2dex.AddRow row.Slice(idxStart, idxEnd)
    Next
    Set ColSlice = arr2dex
End Function

'******************************************************************************
'* [概  要] Range
'* [詳  細] 指定した範囲の数値の2次元配列を作成する。
'*
'* @param fromVal
'* @param xIdx
'* @param yIdx
'* @param stepVal
'* @return Array2DEx
'******************************************************************************
Public Function Range(ByVal fromVal As Variant, ByVal xIdx As Variant, ByVal yIdx As Variant, Optional ByVal stepVal As Variant = 1) As Array2DEx
    Dim arr2dex As Array2DEx: Set arr2dex = New Array2DEx
    Dim cnt As Long: cnt = fromVal
    Dim x As Long, y As Long
    For x = 0 To xIdx
        Dim row As ArrayEx: Set row = New ArrayEx
        For y = 0 To yIdx
            row.Add cnt
            cnt = cnt + stepVal
        Next
        arr2dex.AddRow row
    Next
    Set Range = Core.Init(New Array2DEx, arr2dex)
End Function

'******************************************************************************
'* [概  要] Map
'* [詳  細] 配列の全要素に対して指定したファンクションの処理を適用する。
'*
'* @param fun Func(Of T, U)オブジェクト
'* @return Array2DEx
'******************************************************************************
Public Function Map(ByVal fun As Func) As Array2DEx
    If IsEmptyArray Then
        Set Map = Me
        Exit Function
    End If
    Dim arr2d: arr2d = Me.To2DArray()

    Dim lb1 As Long: lb1 = LBound(arr2d)
    Dim ub1 As Long: ub1 = UBound(arr2d)
    Dim lb2 As Long: lb2 = LBound(arr2d, 2)
    Dim ub2 As Long: ub2 = UBound(arr2d, 2)
    
    Dim ret
    ReDim ret(lb1 To ub1, lb2 To ub2)
    
    Dim x As Long, y As Long
    For x = lb1 To ub1
        For y = lb2 To ub2
            fun.FastApply ret(x, y), arr2d(x, y)
        Next
    Next
    Set Map = Core.Init(New Array2DEx, ret)
End Function

'******************************************************************************
'* [概  要] Zip
'* [詳  細] ２つの2次元配列に対して指定したファンクションの処理を適用し
'*          １つの配列を返す。
'*
'* @param fun As Func(Of T, U, R)オブジェクト
'* @param arr2 As Variant(Of Array(Of U))
'* @return Array2DEx
'******************************************************************************
Public Function Zip(ByVal fun As Func, ByVal arr2 As Variant) As Array2DEx
    If IsEmptyArray Then
        Set Zip = Me
        Exit Function
    End If
    If Not ArrayUtils.Is2DArray(arr2) Then Err.raise 9999, "Array2DEx#Zip", "Invalid argument."
    
    Dim arr1: arr1 = Me.To2DArray()

    Dim lb1 As Long: lb1 = LBound(arr1)
    Dim ub1 As Long: ub1 = UBound(arr1)
    Dim lb2 As Long: lb2 = LBound(arr1, 2)
    Dim ub2 As Long: ub2 = UBound(arr1, 2)
    Dim len1 As Long: len1 = ub1 - lb1
    Dim len2 As Long: len2 = ub2 - lb2
    
    Dim a2lb1 As Long: a2lb1 = LBound(arr2)
    Dim a2ub1 As Long: a2ub1 = UBound(arr2)
    Dim a2lb2 As Long: a2lb2 = LBound(arr2, 2)
    Dim a2ub2 As Long: a2ub2 = UBound(arr2, 2)
    If (ub1 - lb1) <> (a2ub1 - a2lb1) Or _
       (ub2 - lb2) <> (a2ub2 - a2lb2) Then
        Err.raise 9999, "Array2DEx#Zip", "Invalid argument. Defferent count Of elements."
    End If
    
    Dim ret
    ReDim ret(lb1 To ub1, lb2 To ub2)
    
    Dim x As Long, y As Long
    For x = 0 To len1
        For y = 0 To len2
            fun.FastApply ret(x, y), arr1(lb1 + x, lb2 + y), arr2(a2lb1 + x, a2lb2 + y)
        Next
    Next
    Set Zip = Core.Init(New Array2DEx, ret)
End Function

'******************************************************************************
'* [概  要] RowFilter
'* [詳  細] 2次元配列に対して指定したフィルタ処理用のファンクションを適用し、
'*          条件に一致する行のみの2次元配列を返す。
'*
'* @param fun As Func(Of T, Boolean)
'* @return Array2DEx
'******************************************************************************
Public Function RowFilter(ByVal fun As Func) As Array2DEx
    If IsEmptyArray Then
        Set RowFilter = Me
        Exit Function
    End If

    Dim lb As Long: lb = 0
    Dim ub As Long: ub = mRowMaxIndex
    Dim ret As Array2DEx: Set ret = New Array2DEx
    
    Dim flg As Boolean
    Dim x As Long
    For x = lb To ub
        fun.FastApply flg, XRow(x).ToArray
        If flg Then ret.AddRow XRow(x)
    Next
    
    If ret.IsEmptyArray Then
        Set RowFilter = New Array2DEx
    Else
        Set RowFilter = Core.Init(New Array2DEx, ret)
    End If
End Function

'******************************************************************************
'* [概  要] ColFilter
'* [詳  細] 2次元配列に対して指定したフィルタ処理用のファンクションを適用し、
'*          条件に一致する列のみの2次元配列を返す。
'*
'* @param fun As Func(Of T, Boolean)
'* @return Array2DEx
'******************************************************************************
Public Function ColFilter(ByVal fun As Func) As Array2DEx
    If IsEmptyArray Then
        Set ColFilter = Me
        Exit Function
    End If

    Dim lb As Long: lb = 0
    Dim ub As Long: ub = mColMaxIndex
    Dim ret As Array2DEx: Set ret = New Array2DEx
    
    Dim flg As Boolean
    Dim x As Long, y As Long
    For y = lb To ub
        fun.FastApply flg, YColumn(y).ToArray
        If flg Then ret.AddColumn YColumn(y)
    Next
    
    If ret.IsEmptyArray Then
        Set ColFilter = New Array2DEx
    Else
        Set ColFilter = Core.Init(New Array2DEx, ret)
    End If
End Function

'******************************************************************************
'* [概  要] RowFold
'* [詳  細] 畳み込み関数。2次元配列内の各行の要素に対して指定した演算用ファン
'*          クションを適用し、その結果と次の要素にまたファンクションを適用する
'*          といった処理を繰り返し最終的な結果を返す。
'*
'* @param fun Func(Of U, T, U)
'* @param seedv Variant(Of U)
'* @return ArrayEx
'******************************************************************************
Public Function RowFold(ByVal fun As Func, Optional ByVal seedv As Variant) As ArrayEx
    If IsEmptyArray Then
        Set RowFold = New ArrayEx
        Exit Function
    End If

    Dim arrex As ArrayEx: Set arrex = New ArrayEx
    Dim x As Long
    For x = 0 To mRowMaxIndex
        arrex.Add XRow(x).Fold(fun, seedv)
    Next
    
    Set RowFold = arrex
End Function

'******************************************************************************
'* [概  要] ColFold
'* [詳  細] 畳み込み関数。2次元配列内の各列の要素に対して指定した演算用ファン
'*          クションを適用し、その結果と次の要素にまたファンクションを適用する
'*          といった処理を繰り返し最終的な結果を返す。
'*
'* @param fun Func(Of U, T, U)
'* @param seedv Variant(Of U)
'* @return ArrayEx
'******************************************************************************
Public Function ColFold(ByVal fun As Func, Optional ByVal seedv As Variant) As ArrayEx
    If IsEmptyArray Then
        Set ColFold = New ArrayEx
        Exit Function
    End If

    Dim arrex As ArrayEx: Set arrex = New ArrayEx
    Dim y As Long
    For y = 0 To mColMaxIndex
        arrex.Add YColumn(y).Fold(fun, seedv)
    Next
    
    Set ColFold = arrex
End Function

'******************************************************************************
'* [概  要] RowScan
'* [詳  細] 2次元配列内の各行の各要素に対して繰り返し指定した演算用ファンク
'*          ションを適用し、累積した計算結果のリストを返す。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func(Of U, T, U)
'* @param seedv As Variant(Of U)
'* @return Array2DEx(Of Array(Of U))
'******************************************************************************
Public Function RowScan(ByVal fun As Func, Optional ByVal seedv As Variant) As Array2DEx
    If IsEmptyArray Then
        Set RowScan = New Array2DEx
        Exit Function
    End If

    Dim arr2dex As Array2DEx: Set arr2dex = New Array2DEx
    Dim x As Long
    For x = 0 To mRowMaxIndex
        arr2dex.AddRow XRow(x).Scan(fun, seedv)
    Next
    
    Set RowScan = arr2dex
End Function

'******************************************************************************
'* [概  要] ColScan
'* [詳  細] 2次元配列内の各列の各要素に対して繰り返し指定した演算用ファンク
'*          ションを適用し、累積した計算結果のリストを返す。
'*          AriawaseのExtモジュールラッパーメソッド。
'*
'* @param fun As Func(Of U, T, U)
'* @param seedv As Variant(Of U)
'* @return Array2DEx(Of Array(Of U))
'******************************************************************************
Public Function ColScan(ByVal fun As Func, Optional ByVal seedv As Variant) As Array2DEx
    If IsEmptyArray Then
        Set ColScan = New Array2DEx
        Exit Function
    End If

    Dim arr2dex As Array2DEx: Set arr2dex = New Array2DEx
    Dim y As Long
    For y = 0 To mColMaxIndex
        arr2dex.AddColumn YColumn(y).Scan(fun, seedv)
    Next
    
    Set ColScan = arr2dex
End Function

'******************************************************************************
'* [概  要] RowContains
'* [詳  細] 指定した行について、指定した項目が含まれるかチェックする。
'*
'* @param x 行インデックス
'* @param itm 項目
'* @return Boolean チェック結果（True：含む）
'******************************************************************************
Public Function RowContains(ByVal x As Long, ByVal itm As Variant) As Boolean
    If IsInvalidRowIndex(x) Then
        RowContains = False
        Exit Function
    End If
    RowContains = ArrayUtils.Contains(XRow(x).ToArray, itm)
End Function

'******************************************************************************
'* [概  要] ColContains
'* [詳  細] 指定した列について、指定した項目が含まれるかチェックする。
'*
'* @param y 列インデックス
'* @param itm 項目
'* @return Boolean チェック結果（True：含む）
'******************************************************************************
Public Function ColContains(ByVal y As Long, ByVal itm As Variant) As Boolean
    If IsInvalidColIndex(y) Then
        ColContains = False
        Exit Function
    End If
    ColContains = ArrayUtils.Contains(YColumn(y).ToArray, itm)
End Function

'******************************************************************************
'* [概  要] Contains
'* [詳  細] 指定した項目が含まれるかチェックする。
'*
'* @param itm 項目
'* @return Boolean チェック結果（True：含む）
'******************************************************************************
Public Function Contains(ByVal itm As Variant) As Boolean
    If IsEmptyArray Then
        Contains = False
        Exit Function
    End If
    Dim x As Long
    For x = 0 To mRowMaxIndex
        If RowContains(x, itm) Then
            Contains = True
            Exit Function
        End If
    Next
    Contains = False
End Function

'******************************************************************************
'* [概  要] Clone
'* [詳  細] Array2DExオブジェクトを複製する。
'*
'* @return Array2DEx 新しいArray2DEx
'******************************************************************************
Public Function Clone() As Array2DEx
    Set Clone = Core.Init(New Array2DEx, Me.To2DArray)
End Function

'******************************************************************************
'* [概  要] RowInsert
'* [詳  細] 行を指定位置に追加する。
'*
'* @param idx 行インデックス
'* @param row 行データ（1次元配列 or ArrayEx）
'* @return Array2DEx
'******************************************************************************
Public Function RowInsert(ByVal idx As Long, ByVal row As Variant) As Array2DEx
    If idx < 0 Or idx > mRowMaxIndex + 1 Then
        Err.raise 9999, "Array2DEx#RowInsert", "Invalid argument."
    End If

    Dim colLen As Long
    Dim arrex As ArrayEx
    If ArrayUtils.Is1DArray(row) Then
        colLen = ArrayUtils.GetLength(row)
        Set arrex = Core.Init(New ArrayEx, row)
    ElseIf ValidateUtils.IsArrayEx(row) Then
        colLen = row.Length
        Set arrex = row
    Else
        Err.raise 9999, "Array2DEx#RowInsert", "Invalid argument."
    End If

    If mColLength < colLen Then
        Dim addCount As Long: addCount = colLen - mColLength
        Call ExpandColumns(addCount)
    End If

    m2DItems.Insert idx, arrex
    mRowMaxIndex = mRowMaxIndex + 1
    mRowLength = mRowLength + 1
    Set RowInsert = Me
End Function

'******************************************************************************
'* [概  要] ColInsert
'* [詳  細] 列を指定位置に追加する。
'*
'* @param idx 列インデックス
'* @param col 列データ（1次元配列 or ArrayEx）
'* @return Array2DEx
'******************************************************************************
Public Function ColInsert(ByVal idx As Long, ByVal col As Variant) As Array2DEx
    If idx < 0 Or idx > mColMaxIndex + 1 Then
        Err.raise 9999, "Array2DEx#ColInsert", "Invalid argument."
    End If
    
    Dim rowLen As Long
    Dim arrex As ArrayEx
    If ArrayUtils.Is1DArray(col) Then
        rowLen = ArrayUtils.GetLength(col)
        Set arrex = Core.Init(New ArrayEx, col)
    ElseIf ValidateUtils.IsArrayEx(col) Then
        rowLen = col.Length
        Set arrex = col
    Else
        Err.raise 9999, "Array2DEx#ColInsert", "Invalid argument."
    End If

    If mRowLength < rowLen Then
        Dim addCount As Long: addCount = rowLen - mRowLength
        Call ExpandRows(addCount)
    End If
    
    Dim x As Long
    For x = 0 To mRowMaxIndex
        m2DItems(x).Insert idx, arrex(x)
    Next

    mColMaxIndex = mColMaxIndex + 1
    mColLength = mColLength + 1
    Set ColInsert = Me
End Function

'******************************************************************************
'* [概  要] RowsInsert
'* [詳  細] 行（複数）を指定位置に追加する｡
'*
'* @param idx 行インデックス
'* @param rows 追加行（複数:2次元Variant配列 or Array2DEx）
'* @return Array2DEx
'******************************************************************************
Public Function RowsInsert(ByVal idx As Long, ByVal Rows As Variant, ParamArray rows2() As Variant) As Array2DEx
    If idx < 0 Or idx > mRowMaxIndex + 1 Then
        Err.raise 9999, "Array2DEx#RowsInsert", "Invalid argument."
    End If
    
    Dim colLen As Long
    Dim arr2dex As Array2DEx
    If Not VBA.IsMissing(rows2) Then
        Dim tmp: tmp = rows2
        Set arr2dex = New Array2DEx
        arr2dex.AddRows Rows, tmp
    Else
        Set arr2dex = Core.Init(New Array2DEx, Rows)
    End If
    colLen = arr2dex.ColLength

    If mColLength < colLen Then
        Dim addCount As Long: addCount = colLen - mColLength
        Call ExpandColumns(addCount)
    End If
    
    Dim v
    For Each v In arr2dex.ToArrayOfArrayEx
        m2DItems.Insert idx, v
        idx = idx + 1
    Next
    mRowLength = mRowLength + arr2dex.RowLength
    mRowMaxIndex = mRowMaxIndex + arr2dex.RowLength

    Set RowsInsert = Me
End Function

'******************************************************************************
'* [概  要] ColsInsert
'* [詳  細] 列（複数）を指定位置に追加する｡
'*
'* @param idx 列インデックス
'* @param cols 追加列（複数:2次元Variant配列 or Array2DEx）
'* @return Array2DEx
'******************************************************************************
Public Function ColsInsert(ByVal idx As Long, ByVal cols As Variant, ParamArray cols2() As Variant) As Array2DEx
    If idx < 0 Or idx > mColMaxIndex + 1 Then
        Err.raise 9999, "Array2DEx#ColsInsert", "Invalid argument."
    End If
    
    Dim rowLen As Long
    Dim arr2dex As Array2DEx
    If Not VBA.IsMissing(cols2) Then
        Dim tmp: tmp = cols2
        Set arr2dex = New Array2DEx
        arr2dex.AddColumns cols, tmp
    Else
        Set arr2dex = Core.Init(New Array2DEx, cols)
    End If
    rowLen = arr2dex.RowLength

    If mRowLength < rowLen Then
        Dim addCount As Long: addCount = rowLen - mRowLength
        Call ExpandRows(addCount)
    End If
    
    Dim x As Long
    For x = 0 To mRowMaxIndex
        m2DItems(x).InsertAll idx, arr2dex.XRow(x).ToArray
    Next
    mColLength = mColLength + arr2dex.ColLength
    mColMaxIndex = mColMaxIndex + arr2dex.ColLength

    Set ColsInsert = Me
End Function

'******************************************************************************
'* [概  要] RowLastIndexOf
'* [詳  細] 2次元配列の指定した行について、指定した値と一致する要素の
'*          列インデックスを返却する。
'*          最後尾から検索して最初に一致したものを返却。
'*          値の捜査開始位置と要素数を任意指定可能。
'*
'* @param rowidx 行インデックス
'* @param val 値
'* @param ixStart 捜査開始位置
'* @param cnt 要素数
'* @return Long 列インデックス（一致なしの場合は-1）
'******************************************************************************
Public Function RowLastIndexOf(ByVal rowIdx As Variant, ByVal val As Variant, Optional ByVal ixStart As Variant, Optional ByVal cnt As Variant) As Long
    If IsInvalidRowIndex(rowIdx) Then
        Err.raise 9999, "Array2DEx#RowLastIndexOf", "Invalid argument."
    End If
    RowLastIndexOf = ArrayUtils.LastIndexOf(m2DItems(rowIdx).ToArray, val, ixStart, cnt)
End Function

'******************************************************************************
'* [概  要] ColLastIndexOf
'* [詳  細] 2次元配列の指定した列について、指定した値と一致する要素の
'*          行インデックスを返却する。
'*          最後尾から検索して最初に一致したものを返却。
'*          値の捜査開始位置と要素数を任意指定可能。
'*
'* @param colidx 列インデックス
'* @param val 値
'* @param ixStart 捜査開始位置
'* @param cnt 要素数
'* @return Long 行インデックス（一致なしの場合は-1）
'******************************************************************************
Public Function ColLastIndexOf(ByVal colidx As Variant, ByVal val As Variant, Optional ByVal ixStart As Variant, Optional ByVal cnt As Variant) As Long
    If IsInvalidColIndex(colidx) Then
        Err.raise 9999, "Array2DEx#ColLastIndexOf", "Invalid argument."
    End If
    ColLastIndexOf = ArrayUtils.LastIndexOf(YColumn(colidx).ToArray, val, ixStart, cnt)
End Function

'******************************************************************************
'* [概  要] LastIndexOf
'* [詳  細] 2次元配列について、指定した値と一致する要素の
'*          行列インデックスを返却する。
'*          最後尾から検索して最初に一致したものを返却。
'*
'* @param val 値
'* @param ixStart 捜査開始行インデックス
'* @return Array2DIndex 行列インデックス（一致なしの場合は行列ともに-1）
'******************************************************************************
Public Function LastIndexOf(ByVal val As Variant, Optional ByVal ixStart) As Array2DIndex
    If VBA.IsMissing(ixStart) Then
        ixStart = mRowMaxIndex
    End If
    If IsInvalidRowIndex(ixStart) Then
        Err.raise 9999, "Array2DEx#LastIndexOf", "Invalid argument."
    End If

    Dim x As Long, y As Long, arr2dIdx As Array2DIndex
    For x = ixStart To 0 Step -1
        y = ArrayUtils.LastIndexOf(m2DItems(x).ToArray, val)
        If y <> -1 Then
            arr2dIdx.x = x
            arr2dIdx.y = y
            LastIndexOf = arr2dIdx
            Exit Function
        End If
    Next
    arr2dIdx.x = -1
    arr2dIdx.y = -1
    LastIndexOf = arr2dIdx
End Function

'******************************************************************************
'* [概  要] RowIndexListOf
'* [詳  細] 2次元配列の指定した行について、指定した値と一致する要素の
'*          列インデックスリストを返却する。
'*
'* @param rowidx 行インデックス
'* @param val 値
'* @param reverseOrder リストの並びを逆順にするか
'* @return Variant インデックスリスト（1次元配列）
'******************************************************************************
Public Function RowIndexListOf(ByVal rowIdx As Variant, ByVal val As Variant, Optional ByVal reverseOrder As Boolean = False) As Variant
    If IsInvalidRowIndex(rowIdx) Then
        Err.raise 9999, "Array2DEx#RowIndexListOf", "Invalid argument."
    End If
    RowIndexListOf = m2DItems(rowIdx).IndexListOf(val, reverseOrder)
End Function

'******************************************************************************
'* [概  要] ColIndexListOf
'* [詳  細] 2次元配列の指定した列について、指定した値と一致する要素の
'*          行インデックスリストを返却する。
'*
'* @param colidx 列インデックス
'* @param val 値
'* @param reverseOrder リストの並びを逆順にするか
'* @return Variant インデックスリスト（1次元配列）
'******************************************************************************
Public Function ColIndexListOf(ByVal colidx As Variant, ByVal val As Variant, Optional ByVal reverseOrder As Boolean = False) As Variant
    If IsInvalidColIndex(colidx) Then
        Err.raise 9999, "Array2DEx#ColIndexListOf", "Invalid argument."
    End If
    ColIndexListOf = YColumn(colidx).IndexListOf(val, reverseOrder)
End Function

'******************************************************************************
'* [概  要] RowRemoveAt
'* [詳  細] 指定したインデックスにある行を削除する。
'*
'* @param idx 行インデックス
'* @return Array2DEx 削除後2次元配列
'******************************************************************************
Public Function RowRemoveAt(ByVal idx As Long) As Array2DEx
    If IsInvalidRowIndex(idx) Then
        Err.raise 9999, "Array2DEx#RowRemoveAt", "Invalid argument."
    End If
    m2DItems.RemoveAt idx
    mRowLength = mRowLength - 1
    mRowMaxIndex = mRowMaxIndex - 1
    Set RowRemoveAt = Me
End Function

'******************************************************************************
'* [概  要] ColRemoveAt
'* [詳  細] 指定したインデックスにある列を削除する。
'*
'* @param idx 列インデックス
'* @return Array2DEx 削除後2次元配列
'******************************************************************************
Public Function ColRemoveAt(ByVal idx As Long) As Array2DEx
    If IsInvalidColIndex(idx) Then
        Err.raise 9999, "Array2DEx#ColRemoveAt", "Invalid argument."
    End If
    Dim x As Long
    For x = 0 To mRowMaxIndex
        m2DItems(x).RemoveAt idx
    Next
    mColLength = mColLength - 1
    mColMaxIndex = mColMaxIndex - 1
    Set ColRemoveAt = Me
End Function

'******************************************************************************
'* [概  要] RowRemove
'* [詳  細] 指定した列インデックスの項目値が、指定した値に一致する行を削除する。
'*
'* @param colidx 列インデックス
'* @param keyvalue キー値
'* @return Array2DEx 削除後2次元配列
'******************************************************************************
Public Function RowRemove(ByVal colidx As Long, ByVal keyvalue) As Array2DEx
    If IsInvalidColIndex(colidx) Then
        Err.raise 9999, "Array2DEx#RowRemove", "Invalid argument."
    End If
    Dim v
    For Each v In Me.ColIndexListOf(colidx, keyvalue, True)
        Me.RowRemoveAt v
    Next
    Set RowRemove = Me
End Function

'******************************************************************************
'* [概  要] ColRemove
'* [詳  細] 指定した行インデックスの項目値が、指定した値に一致する列を削除する。
'*
'* @param rowidx 列インデックス
'* @param keyvalue キー値
'* @return Array2DEx 削除後2次元配列
'******************************************************************************
Public Function ColRemove(ByVal rowIdx As Long, ByVal keyvalue) As Array2DEx
    If IsInvalidRowIndex(rowIdx) Then
        Err.raise 9999, "Array2DEx#ColRemove", "Invalid argument."
    End If
    Dim v
    For Each v In Me.RowIndexListOf(rowIdx, keyvalue, True)
        Me.ColRemoveAt v
    Next
    Set ColRemove = Me
End Function

'******************************************************************************
'* [概  要] RowRemoveRange
'* [詳  細] 2次元配列の行の範囲を削除する。
'*
'* @param idx1 開始行インデックス
'* @param idx2 終了行インデックス
'* @return Array2DEx 削除後2次元配列
'******************************************************************************
Public Function RowRemoveRange(ByVal idx1 As Long, ByVal idx2 As Long) As Array2DEx
    If IsInvalidRowIndexes(idx1, idx2) Then
        Err.raise 9999, "Array2DEx#RowRemoveRange", "Invalid argument."
    End If
    If idx1 > idx2 Then
        Err.raise 9999, "Array2DEx#RowRemoveRange", "Invalid argument."
    End If
    Dim v
    For Each v In ArrayUtils.Range(idx2, idx1, -1)
        Me.RowRemoveAt v
    Next
    Set RowRemoveRange = Me
End Function

'******************************************************************************
'* [概  要] ColRemoveRange
'* [詳  細] 2次元配列の列の範囲を削除する。
'*
'* @param idx1 開始列インデックス
'* @param idx2 終了列インデックス
'* @return Array2DEx 削除後2次元配列
'******************************************************************************
Public Function ColRemoveRange(ByVal idx1 As Long, ByVal idx2 As Long) As Array2DEx
    If IsInvalidColIndexes(idx1, idx2) Then
        Err.raise 9999, "Array2DEx#ColRemoveRange", "Invalid argument."
    End If
    If idx1 > idx2 Then
        Err.raise 9999, "Array2DEx#ColRemoveRange", "Invalid argument."
    End If
    Dim v
    For Each v In ArrayUtils.Range(idx2, idx1, -1)
        Me.ColRemoveAt v
    Next
    Set ColRemoveRange = Me
End Function

'******************************************************************************
'* [概  要] Join
'* [詳  細] 配列を連結して文字列にする。
'*
'* @param itemDelimiter 項目デリミタ
'* @param rowDelimiter 行デリミタ
'* @param wrapChar 囲み文字
'* @param rmLineSeparator 改行を除去するか
'* @return String 連結文字列
'******************************************************************************
Public Function Join(Optional ByVal itemDelimiter As String, Optional ByVal rowDelimiter As String = vbCrLf, Optional ByVal wrapChar As String, Optional ByVal rmLineSeparator As Boolean = True) As String
    Dim x As Long, strex As StringEx: Set strex = New StringEx
    For x = 0 To mRowMaxIndex
        strex.AppendLine m2DItems(x).Join(itemDelimiter, wrapChar, rmLineSeparator), rowDelimiter
    Next
    Join = strex.ToString
End Function

'******************************************************************************
'* [概  要] RowSearch
'* [詳  細] 配列（文字列）から指定した文字列と部分一致する項目を
'*          指定した行から取得し、1次元配列で返却する。
'*
'* @param rowidx 行インデックス
'* @param searchVal 検索文字列（配列で複数指定可）
'* @return ArrayEx 検索結果
'******************************************************************************
Public Function RowSearch(ByVal rowIdx As Variant, ByVal searchVal As Variant) As ArrayEx
    If IsInvalidRowIndex(rowIdx) Then
        Err.raise 9999, "Array2DEx#RowSearch", "Invalid argument."
    End If
    Set RowSearch = m2DItems(rowIdx).Search(searchVal)
End Function

'******************************************************************************
'* [概  要] ColSearch
'* [詳  細] 配列（文字列）から指定した文字列と部分一致する項目を
'*          指定した列から取得し、1次元配列で返却する。
'*
'* @param colidx 列インデックス
'* @param searchVal 検索文字列（配列で複数指定可）
'* @return ArrayEx 検索結果
'******************************************************************************
Public Function ColSearch(ByVal colidx As Variant, ByVal searchVal As Variant) As ArrayEx
    If IsInvalidColIndex(colidx) Then
        Err.raise 9999, "Array2DEx#ColSearch", "Invalid argument."
    End If
    Set ColSearch = YColumn(colidx).Search(searchVal)
End Function

'******************************************************************************
'* [概  要] Search
'* [詳  細] 配列（文字列）から指定した文字列と部分一致する項目を
'*          取得し、1次元配列で返却する。
'*
'* @param searchVal 検索文字列（配列で複数指定可）
'* @return ArrayEx 検索結果
'******************************************************************************
Public Function Search(ByVal searchVal As Variant) As ArrayEx
    Dim x As Long, arrex As ArrayEx: Set arrex = New ArrayEx
    For x = 0 To mRowMaxIndex
        Dim retarrex As ArrayEx
        Set retarrex = Me.RowSearch(x, searchVal)
        arrex.AddAll retarrex.ToArray
    Next
    Set Search = arrex
End Function

'******************************************************************************
'* [概  要] RowRegexSearch
'* [詳  細] 配列（文字列）から指定した文字列と正規表現が一致する項目を
'*          指定した行から取得し、1次元配列で返却する。
'*
'* @param rowidx 行インデックス
'* @param searchStr 検索文字列（正規表現）
'* @param igCase 大文字小文字の区別
'* @return ArrayEx 検索結果
'******************************************************************************
Public Function RowRegexSearch(ByVal rowIdx As Variant, ByVal searchStr As String, Optional ByVal igCase As Boolean = False) As ArrayEx
    If IsInvalidRowIndex(rowIdx) Then
        Err.raise 9999, "Array2DEx#RowRegexSearch", "Invalid argument."
    End If
    Set RowRegexSearch = m2DItems(rowIdx).RegexSearch(searchStr, igCase)
End Function

'******************************************************************************
'* [概  要] ColRegexSearch
'* [詳  細] 配列（文字列）から指定した文字列と正規表現が一致する項目を
'*          指定した列から取得し、1次元配列で返却する。
'*
'* @param colidx 列インデックス
'* @param searchStr 検索文字列（正規表現）
'* @param igCase 大文字小文字の区別
'* @return ArrayEx 検索結果
'******************************************************************************
Public Function ColRegexSearch(ByVal colidx As Variant, ByVal searchStr As String, Optional ByVal igCase As Boolean = False) As ArrayEx
    If IsInvalidColIndex(colidx) Then
        Err.raise 9999, "Array2DEx#ColRegexSearch", "Invalid argument."
    End If
    Set ColRegexSearch = YColumn(colidx).RegexSearch(searchStr, igCase)
End Function

'******************************************************************************
'* [概  要] RegexSearch
'* [詳  細] 配列（文字列）から指定した文字列と正規表現が一致する項目を
'*          取得し、1次元配列で返却する。
'*
'* @param searchStr 検索文字列（正規表現）
'* @param igCase 大文字小文字の区別
'* @return ArrayEx 検索結果
'******************************************************************************
Public Function RegexSearch(ByVal searchStr As String, Optional ByVal igCase As Boolean = False) As ArrayEx
    Dim x As Long, arrex As ArrayEx: Set arrex = New ArrayEx
    For x = 0 To mRowMaxIndex
        Dim retarrex As ArrayEx
        Set retarrex = Me.RowRegexSearch(x, searchStr, igCase)
        arrex.AddAll retarrex.ToArray
    Next
    Set RegexSearch = arrex
End Function

'******************************************************************************
'* [概  要] ToString
'* [詳  細] JSON文字列に変換する。
'*
'* @return String JSON文字列
'******************************************************************************
Public Function ToString() As String
    ToString = ArrayUtils.ToString(Me.ToJagArray)
End Function


'******************************************************************************
'* [概  要] FromCsvString
'* [詳  細] CSVファイルを行単位で読込み、Array2DExに設定して返却する。
'*
'* @param filePath CSVファイルパス
'* @param headerExists ヘッダ行の有無
'* @param quotExists ""で括っているか
'* @param separator 改行コード
'* @return Array2DEx
'******************************************************************************
Public Function FromCsvString(ByVal csvData As String, _
                              Optional ByVal HeaderExists As Boolean = True, _
                              Optional ByVal QuotExists As Boolean = False, _
                              Optional ByVal separator As LineSeparatorsEnum = adCRLF) As Array2DEx
    Set FromCsvString = Init(ArrayUtils.CsvStringToArray(csvData, HeaderExists, QuotExists, separator))
End Function
