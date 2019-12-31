VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] レコードフォーマットクラス
'* [詳  細] 項目のフォーマット定義情報をレコード単位で束ねたクラス。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mCol As Collection
Private mRecordSetCol As Collection
Private mColumnCount As Long
Private mDataRowCount As Long

Private mErrMessage As String
Private mErrRowNo As Long
Private mErrColNo As Long

'******************************************************************************
'* プロパティ定義
'******************************************************************************

'*-----------------------------------------------------------------------------
'* Item プロパティ
'*-----------------------------------------------------------------------------
Property Get Item(index) As Item
    Item = mCol.Item(index)
End Property
Property Let Item(index, arg As Item)
    mCol.Item(index) = arg
End Property

'*-----------------------------------------------------------------------------
'* RecordSet プロパティ（読み取り専用）
'*-----------------------------------------------------------------------------
Property Get RecordSet() As Collection
    Set RecordSet = mRecordSetCol
End Property

'*-----------------------------------------------------------------------------
'* RecordDefinition プロパティ（読み取り専用）
'*-----------------------------------------------------------------------------
Property Get RecordDefinition() As Collection
    Set RecordDefinition = mCol
End Property

'*-----------------------------------------------------------------------------
'* ColumnCount プロパティ
'*-----------------------------------------------------------------------------
Property Get ColumnCount() As Long
    ColumnCount = mColumnCount
End Property
Property Let ColumnCount(arg As Long)
    mColumnCount = arg
End Property

'*-----------------------------------------------------------------------------
'* DataRowCount プロパティ
'*-----------------------------------------------------------------------------
Property Get DataRowCount() As Long
    DataRowCount = mDataRowCount
End Property
Property Let DataRowCount(arg As Long)
    mDataRowCount = arg
End Property

'*-----------------------------------------------------------------------------
'* ErrMessage プロパティ
'*-----------------------------------------------------------------------------
Property Get ErrMessage() As String
    ErrMessage = mErrMessage
End Property
Property Let ErrMessage(arg As String)
    mErrMessage = arg
End Property

'*-----------------------------------------------------------------------------
'* ErrRowNo プロパティ
'*-----------------------------------------------------------------------------
Property Get ErrRowNo() As Long
    ErrRowNo = mErrRowNo
End Property
Property Let ErrRowNo(arg As Long)
    mErrRowNo = arg
End Property

'*-----------------------------------------------------------------------------
'* mErrColNo プロパティ
'*-----------------------------------------------------------------------------
Property Get ErrColNo() As Long
    ErrColNo = mErrColNo
End Property
Property Let ErrColNo(arg As Long)
    mErrColNo = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mCol = New Collection
    Set mRecordSetCol = New Collection
End Sub
Private Sub Class_Terminate()
    Set mCol = Nothing
    Set mRecordSetCol = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] 項目表情報取得・設定処理。
'* [詳  細] worksheetの項目表から情報を取得し、コレクションに項目を一括設定します｡
'*
'* @param dataSheet 項目表ワークシート
'* @param lStartRow 項目表データ開始行番号
'* @param lStartCol 項目表データ開始列番号
'* @param itemCount 項目列数
'*
'******************************************************************************
Public Function GetItemDataFromSheet(dataSheet As Worksheet, lStartRow As Long, lStartCol As Long, Optional colCount As Long) As Boolean
    mDataRowCount = 0
    Dim vArr: vArr = GetVariantDataFromSheet(dataSheet, lStartRow, lStartCol, colCount)
    If IsEmpty(vArr) Then
        GetItemDataFromSheet = False
        Exit Function
    End If
    
    Call SetItems(vArr)
    
    GetItemDataFromSheet = True
End Function


'******************************************************************************
'* [概  要] 項目一括設定処理。
'* [詳  細] 項目一括設定します｡
'*
'* @param vArr Variant型2次元配列（項目表データ）
'*
'******************************************************************************
Private Sub SetItems(vArr)
    Dim i As Long

    For i = LBound(vArr, 1) To UBound(vArr, 1)
        Call AddItem(vArr, i)
    Next
    mColumnCount = UBound(vArr, 1) - LBound(vArr, 1) + 1
End Sub

'******************************************************************************
'* [概  要] 項目設定処理。
'* [詳  細] Itemに値を設定しコレクションに追加します｡
'*
'* @param vArr Variant型2次元配列（項目表データ）
'* @param rownum 配列行（1次元添え字）
'*
'******************************************************************************
Private Sub AddItem(vArr, rownum As Long)
    Dim itm As Item
    Set itm = New Item
    itm.Name = vArr(rownum, 1)
    If vArr(rownum, 2) = "○" Then
        itm.Required = True
    End If
    Select Case vArr(rownum, 3)
        Case "半角"
            itm.Attr = AttributeEnum.attrHalf
        Case "半角英数"
            itm.Attr = AttributeEnum.attrHalfAlphaNumeric
        Case "半角英数記号"
            itm.Attr = AttributeEnum.attrHalfAlphaNumericSymbol
        Case "数値"
            itm.Attr = AttributeEnum.attrNumeric
        Case "日付"
            itm.Attr = AttributeEnum.attrDate
        Case Else
            itm.Attr = AttributeEnum.attrString
    End Select
    Select Case vArr(rownum, 4)
        Case "固定"
            itm.KindOfDigits = KindOfDigitsEnum.digitFixed
        Case "以内"
            itm.KindOfDigits = KindOfDigitsEnum.digitWithin
        Case "範囲"
            itm.KindOfDigits = KindOfDigitsEnum.digitRange
        Case Else
            itm.KindOfDigits = KindOfDigitsEnum.digitNone
    End Select
    If vArr(rownum, 5) <> "" And VBA.IsNumeric(vArr(rownum, 5)) Then
        itm.MinNumOfDigits = CLng(vArr(rownum, 5))
    End If
    If vArr(rownum, 6) <> "" And VBA.IsNumeric(vArr(rownum, 6)) Then
        itm.MaxNumOfDigits = CLng(vArr(rownum, 6))
    End If
    itm.Pattern = vArr(rownum, 7)
    If UBound(vArr, 2) = 13 Then
        If vArr(rownum, 8) <> "" And VBA.IsNumeric(vArr(rownum, 8)) Then
            itm.InputCsvColNo = vArr(rownum, 8)
        End If
        itm.InitValueKind = vArr(rownum, 9)
        itm.InitValue = vArr(rownum, 10)
        If vArr(rownum, 11) = "○" Then
            itm.OutputTarget = True
        End If
        itm.OutputEditKind = vArr(rownum, 12)
        itm.OutputEditValue = vArr(rownum, 13)
    End If
    mCol.Add itm
End Sub

'******************************************************************************
'* [概  要] 検証処理。
'* [詳  細] レコードデータを検証します｡
'*
'* @param vArr Variant型2次元配列（レコードデータ）
'* @return 検証結果
'*
'******************************************************************************
Public Function Validate(vArr)
    If mColumnCount <> (UBound(vArr, 2) - LBound(vArr, 2) + 1) Then
        Validate = False
        Exit Function
    End If
    Dim statusBarTxt As String: statusBarTxt = "CSVデータ検証中："
    Application.StatusBar = statusBarTxt
    
    Dim i As Long, j As Long, itm As Item, index As Long
    index = 1
    For i = LBound(vArr, 1) To UBound(vArr, 1)

        If i Mod 5000 = 0 Then
            statusBarTxt = statusBarTxt & "■"
            Application.StatusBar = statusBarTxt
            Call CheckEvents
        End If
            
        For j = LBound(vArr, 2) To UBound(vArr, 2)
            Set itm = mCol(index)
            itm.Value = vArr(i, j)
            If Not itm.Validate Then
                Validate = False
                Call SetErrorMessage(itm, i, j)
                Exit Function
            End If
            index = index + 1
        Next
        index = 1
    Next
    Validate = True
End Function

'******************************************************************************
'* [概  要] メッセージ設定処理。
'* [詳  細] エラー発生項目からエラーメッセージを取得して設定します｡
'*
'* @param itm エラー発生項目
'* @param rownum エラー発生データ行番号
'* @param colnum エラー発生データ列番号
'*
'******************************************************************************
Private Sub SetErrorMessage(ByVal itm As Item, ByVal rownum As Long, ByVal colnum As Long)
    mErrMessage = "[行" + CStr(rownum) + " 列" + CStr(colnum) + "]  " + itm.Name + vbNewLine + vbNewLine + _
          itm.GetErrorMessage
    mErrRowNo = rownum
    mErrColNo = colnum
End Sub

'******************************************************************************
'* [概  要] レコード情報取得・設定処理。
'* [詳  細] worksheetのデータ表から情報を取得し、コレクションに項目を
'*          一括設定します｡
'*
'* @param dataSheet data表ワークシート
'* @param lStartRow data表データ開始行番号
'* @param lStartCol data表データ開始列番号
'* @param itemCount 項目列数
'*
'******************************************************************************
Public Function GetRecordDataFromSheet(dataSheet As Worksheet, lStartRow As Long, lStartCol As Long, Optional colCount As Long, Optional ignoreColnum As Long) As Boolean
    mDataRowCount = 0
    Dim rng As Range: Set rng = GetActualUsedRangeFromSheet(dataSheet, lStartRow, lStartCol, colCount, ignoreColnum)
    If rng Is Nothing Then
        mErrMessage = "データが入力されていません。"
        mErrRowNo = 1
        mErrColNo = 1
        GetRecordDataFromSheet = False
        Exit Function
    End If
    Dim vArr: vArr = rng
    mDataRowCount = UBound(vArr, 1) - LBound(vArr, 1) + 1
    
    GetRecordDataFromSheet = SetDataItems(vArr)
End Function


'******************************************************************************
'* [概  要] レコード情報一括設定処理。
'* [詳  細] レコード情報を一括設定します｡
'*
'* @param vArr Variant型2次元配列（レコード情報データ）
'*
'******************************************************************************
Public Function SetDataItems(vArr) As Boolean
    Dim i As Long, j As Long, itm As Item, index As Long
    Dim recordCol As Collection
    For i = LBound(vArr, 1) To UBound(vArr, 1)
        Set recordCol = New Collection

        index = 1
        For j = LBound(vArr, 2) To UBound(vArr, 2)
            Set itm = CopyItem(mCol(index))
            itm.ForOutput = True
            itm.Value = vArr(i, j)
            
            If Not itm.Validate Then
                SetDataItems = False
                Call SetErrorMessage(itm, i, j)
                Exit Function
            End If
            recordCol.Add itm
            index = index + 1
        Next
        mRecordSetCol.Add recordCol
    Next
    SetDataItems = True
End Function

'******************************************************************************
'* [概  要] Itemコピー処理。
'* [詳  細] Itemオブジェクトのコピーを返却します｡
'*
'* @param itm Itemオブジェクト
'* @return コピーしたItemオブジェクト
'*
'******************************************************************************
Public Function CopyItem(itm As Item) As Item
    Dim newItem As Item: Set newItem = New Item
    newItem.Name = itm.Name
    newItem.Value = itm.Value
    newItem.Required = itm.Required
    newItem.Attr = itm.Attr
    newItem.KindOfDigits = itm.KindOfDigits
    newItem.MinNumOfDigits = itm.MinNumOfDigits
    newItem.MaxNumOfDigits = itm.MaxNumOfDigits
    newItem.Pattern = itm.Pattern
    newItem.InitValueKind = itm.InitValueKind
    newItem.InitValue = itm.InitValue
    newItem.InputCsvColNo = itm.InputCsvColNo
    newItem.OutputTarget = itm.OutputTarget
    newItem.OutputEditKind = itm.OutputEditKind
    newItem.OutputEditValue = itm.OutputEditValue
    Set CopyItem = newItem
End Function

'******************************************************************************
'* [概  要] フォーム用Variant配列取得処理。
'* [詳  細] 読込レコードデータをフォーム定義情報に基づき、フォーム用データ
'*          に変換し返却します｡
'*
'* @param vArr Variant型2次元配列（読込レコードデータ）
'* @return Variant型2次元配列（フォーム用データ）
'*
'******************************************************************************
Public Function GetFormVariantData(vArr) As Variant
    Dim vFormArr()
    ReDim vFormArr(0 To UBound(vArr, 1) - LBound(vArr, 1), 0 To mColumnCount - 1)
    
    Dim sInfo As SettingInfo: Set sInfo = GetSettingInfo()
    Dim statusBarTxt As String: statusBarTxt = "シートデータ変換処理中："
    Application.StatusBar = statusBarTxt

    Dim i As Long, j As Long, itm As Item, index As Long, val As String
    For i = LBound(vFormArr, 1) To UBound(vFormArr, 1)

        If i Mod 5000 = 0 Then
            statusBarTxt = statusBarTxt & "■"
            Application.StatusBar = statusBarTxt
            Call CheckEvents
        End If
        
        index = 1
        For j = LBound(vFormArr, 2) To UBound(vFormArr, 2)
            Set itm = mCol(index)
            
            '値を取得
            If itm.InputCsvColNo = 0 Then
                val = ""
            Else
                val = vArr(i, itm.InputCsvColNo - 1)
            End If
            
            If itm.InitValueKind <> "" Then
                If itm.InitValueKind = "デフォルト" Then
                    val = itm.InitValue
                ElseIf itm.InitValueKind = "マスタ変換（Code→Value）" Then
                    val = sInfo.GetMasterValueByCode(itm.InitValue, val)
                ElseIf itm.InitValueKind = "マスタ変換（Value→Code）" Then
                    val = sInfo.GetMasterCodeByValue(itm.InitValue, val)
                End If
            End If
            vFormArr(i, j) = val
            index = index + 1
        Next
    Next
    GetFormVariantData = vFormArr
End Function
