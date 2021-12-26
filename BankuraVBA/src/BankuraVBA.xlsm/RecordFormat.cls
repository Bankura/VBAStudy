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
Implements IRowEditor

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] レコードフォーマットクラス
'* [詳  細] 項目のフォーマット定義情報をレコード単位で束ねたクラス。
'*
'* @author Bankura
'* Copyright (c) 2019-2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* Enum定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* 処理の種類。
'*
'*-----------------------------------------------------------------------------
Private Enum ProcessTypeEnum
    ptValidate
    ptConvert
End Enum

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

Private mRowCount As Long
Private mProgressReporter As IProgressReporter
Private mReportUnit As Long

Private mProcessType As ProcessTypeEnum
Private mForOutput As Boolean

'******************************************************************************
'* プロパティ定義
'******************************************************************************

'*-----------------------------------------------------------------------------
'* Item プロパティ
'*-----------------------------------------------------------------------------
Property Get Item(Index) As Item
    Item = mCol.Item(Index)
End Property
Property Let Item(Index, arg As Item)
    mCol.Item(Index) = arg
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
'* ItemNames プロパティ（読み取り専用）
'*-----------------------------------------------------------------------------
Property Get ItemNames() As Collection
    Dim itm As Item
    Dim itmNamesCol As Collection: Set itmNamesCol = New Collection
    For Each itm In mCol
        itmNamesCol.Add itm.Name
    Next
    Set ItemNames = itmNamesCol
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

'*-----------------------------------------------------------------------------
'* ProgressReporter プロパティ
'*-----------------------------------------------------------------------------
Property Get ProgressReporter() As IProgressReporter
    Set ProgressReporter = mProgressReporter
End Property
Property Set ProgressReporter(arg As IProgressReporter)
    Set mProgressReporter = arg
End Property

'*-----------------------------------------------------------------------------
'* ReportUnit  プロパティ
'*-----------------------------------------------------------------------------
Property Get ReportUnit() As Long
    ReportUnit = mReportUnit
End Property
Property Let ReportUnit(arg As Long)
    mReportUnit = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mCol = New Collection
    Set mRecordSetCol = New Collection
    mRowCount = 0
    mColumnCount = 0
    mReportUnit = 500
    mDataRowCount = 0
End Sub
Private Sub Class_Terminate()
    Set mCol = Nothing
    Set mRecordSetCol = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Init メソッド
'* [詳  細] 初期化用メソッド｡
'*
'* @param itmdefs 定義済みItemコレクション
'* @return RecordFormatオブジェクト
'*
'******************************************************************************
Public Function Init(ByVal itmdefs As Collection) As RecordFormat
    Set mCol = itmdefs
    mColumnCount = itmdefs.Count
    Set Init = Me
End Function

'******************************************************************************
'* [概  要] 検証処理。
'* [詳  細] レコードデータを検証します｡
'*
'* @param vArr Variant型2次元配列（レコードデータ）
'* @param forOutput Output時
'* @return 検証結果
'*
'******************************************************************************
Public Function Validate(vArr, Optional forOutput As Boolean = False) As Boolean
    mProcessType = ptValidate
    mForOutput = forOutput
    
    Dim arr2dex As Array2DEx
    If ValidateUtils.IsArray2DEx(vArr) Then
        Set arr2dex = vArr
    ElseIf ArrayUtils.Is2DArray(vArr) Then
        Set arr2dex = Core.Init(New Array2DEx, vArr)
    Else
        Validate = False
        Exit Function
    End If
    
    If mColumnCount <> arr2dex.ColLength Then
        Validate = False
        Exit Function
    End If
    
    ' RowEnumeratorの設定
    Dim rEnumerator As RowEnumerator: Set rEnumerator = New RowEnumerator
    Set rEnumerator.RowEditor = Me
    rEnumerator.SourceData = arr2dex
    rEnumerator.IsActionable = False
    rEnumerator.MaxReportCount = arr2dex.RowLength
    rEnumerator.startRow = 0
    rEnumerator.EndRow = arr2dex.RowMaxIndex
    mRowCount = arr2dex.RowLength
    
    ' ProgressControllerの設定
    Dim pController As ProgressController: Set pController = New ProgressController
    If mProgressReporter Is Nothing Then
        pController.IsReportable = False
    Else
        Set pController.ProgressReporter = mProgressReporter
        pController.IsReportable = True
        rEnumerator.ReportUnit = mReportUnit
    End If

    ' 処理開始
    On Error GoTo ErrorHandler
    pController.Start rEnumerator

    Validate = True
    Exit Function
    
ErrorHandler:
    If Err.Number = 9998 Then
        Validate = False
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
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
Private Sub SetErrorMessage(ByVal itm As Item, ByVal rowNum As Long, ByVal colNum As Long)
    mErrMessage = "[行" + CStr(rowNum) + " 列" + CStr(colNum) + "]  " + itm.Name _
                  + vbNewLine + "[" + itm.Value + "]" + vbNewLine + vbNewLine + _
                    itm.GetErrorMessage
    mErrRowNo = rowNum
    mErrColNo = colNum
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
    Dim rng As Range: Set rng = XlWorkSheetUtils.GetActualUsedRangeFromSheet(dataSheet, lStartRow, lStartCol, colCount, ignoreColnum)
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
    Dim i As Long, j As Long, itm As Item, Index As Long
    Dim recordCol As Collection
    For i = LBound(vArr, 1) To UBound(vArr, 1)
        Set recordCol = New Collection

        Index = 1
        For j = LBound(vArr, 2) To UBound(vArr, 2)
            Set itm = mCol(Index).Copy()
            itm.forOutput = True
            itm.Value = vArr(i, j)
            
            If Not itm.Validate Then
                SetDataItems = False
                Call SetErrorMessage(itm, i, j)
                Exit Function
            End If
            recordCol.Add itm
            Index = Index + 1
        Next
        mRecordSetCol.Add recordCol
    Next
    SetDataItems = True
End Function

'******************************************************************************
'* [概  要] 項目定義による入力データ変換処理。
'* [詳  細] 読込レコードデータを項目定義情報に基づき、変換｡
'*
'* @param vArr Variant型2次元配列（読込レコードデータ）
'* @return Array2DEx フォーム用データ
'*
'******************************************************************************
Public Function Convert(vArr, Optional forOutput As Boolean = False) As Array2DEx
    mProcessType = ptConvert
    mForOutput = forOutput
    
    Dim arr2dex As Array2DEx, destArr2dex As Array2DEx
    Set destArr2dex = New Array2DEx
    If ValidateUtils.IsArray2DEx(vArr) Then
        Set arr2dex = vArr
    ElseIf ArrayUtils.Is2DArray(vArr) Then
        Set arr2dex = Core.Init(New Array2DEx, vArr)
    Else
        Set Convert = destArr2dex
        Exit Function
    End If
    
    ' RowEnumeratorの設定
    Dim rEnumerator As RowEnumerator: Set rEnumerator = New RowEnumerator
    Set rEnumerator.RowEditor = Me
    rEnumerator.SourceData = arr2dex
    rEnumerator.DestTarget = destArr2dex
    rEnumerator.IsActionable = False
    rEnumerator.MaxReportCount = arr2dex.RowLength
    rEnumerator.startRow = 0
    rEnumerator.EndRow = arr2dex.RowMaxIndex
    mRowCount = arr2dex.RowLength
    
    ' ProgressControllerの設定
    Dim pController As ProgressController: Set pController = New ProgressController
    If mProgressReporter Is Nothing Then
        pController.IsReportable = False
    Else
        Set pController.ProgressReporter = mProgressReporter
        pController.IsReportable = True
        rEnumerator.ReportUnit = mReportUnit
    End If

    ' 処理開始
    pController.Start rEnumerator

    Set Convert = destArr2dex
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
Public Function ConvertInToVArray(vArr) As Variant
    ConvertInToVArray = Convert(vArr).To2DArray
End Function

'******************************************************************************
'* [概  要] StartProcess
'* [詳  細] 開始時の処理を行う。
'*
'******************************************************************************
Private Sub StartProcess(enumerator As Object)
    ' Nothing
End Sub

'******************************************************************************
'* [概  要] Edit
'* [詳  細] 行編集処理を行う。
'*
'******************************************************************************
Private Sub Edit(enumerator As Object)
    Dim rEnumerator As RowEnumerator
    Set rEnumerator = enumerator
    
    Select Case mProcessType
        Case ProcessTypeEnum.ptValidate
            Call Edit_Validate(rEnumerator)
        Case ProcessTypeEnum.ptConvert
            Call Edit_Convert(rEnumerator)
    End Select
End Sub

Private Sub Edit_Validate(enumerator As RowEnumerator)
    Dim arr2dex As Array2DEx
    Set arr2dex = enumerator.SourceData
    Dim colIdx As Long, itm As Item, ind As Long
    ind = 1
    For colIdx = 0 To arr2dex.ColMaxIndex
        Set itm = mCol(ind)
        itm.forOutput = mForOutput
        itm.Value = arr2dex(enumerator.CurrentRow, colIdx)
        If Not itm.Validate Then
            Call SetErrorMessage(itm, enumerator.CurrentRow + 1, colIdx + 1)
            Err.Raise 9998, "RecordFormat#Edit_Validate", "不正データが見つかりました。"
        End If
        ind = ind + 1
    Next
End Sub

Private Sub Edit_Convert(enumerator As RowEnumerator)
    Dim srcArr2dex As Array2DEx, destArr2dex As Array2DEx, record As ArrayEx
    Set srcArr2dex = enumerator.SourceData
    Set destArr2dex = enumerator.DestTarget
    Set record = New ArrayEx

    Dim colIdx As Long, itm As Item, ind As Long, val As String
    ind = 1
    For colIdx = 0 To mColumnCount - 1
        Set itm = mCol(ind)
        itm.forOutput = mForOutput
        If mForOutput Then
            If itm.OutputTarget Then
                itm.Value = srcArr2dex(enumerator.CurrentRow, colIdx)
                val = itm.EditedValue
                record.Add val
            End If
        Else
            ' 値を取得
            If itm.InputColNo = 0 Then
                val = ""
            Else
                itm.Value = srcArr2dex(enumerator.CurrentRow, itm.InputColNo - 1)
                val = itm.EditedValue
            End If
            record.Add val
        End If
        ind = ind + 1
    Next
    destArr2dex.AddRow record
End Sub


'******************************************************************************
'* [概  要] EndProcess
'* [詳  細] 終了時の処理を行う。
'*
'******************************************************************************
Private Sub EndProcess(enumerator As Object)
    ' Nothing
End Sub

'******************************************************************************
'* [概  要] CancelProcess
'* [詳  細] キャンセル時の処理を行う。
'*
'******************************************************************************
Private Sub CancelProcess(enumerator As Object)
    ' Nothing
End Sub

'******************************************************************************
'* インタフェース実装
'******************************************************************************
'*-----------------------------------------------------------------------------
'* IRowEditor#StartProcess
'*
'*-----------------------------------------------------------------------------
Private Sub IRowEditor_StartProcess(enumerator As Object)
    StartProcess enumerator
End Sub
'*-----------------------------------------------------------------------------
'* IRowEditor#Edit
'*
'*-----------------------------------------------------------------------------
Private Sub IRowEditor_Edit(enumerator As Object)
    Edit enumerator
End Sub
'*-----------------------------------------------------------------------------
'* IRowEditor#EndProcess
'*
'*-----------------------------------------------------------------------------
Private Sub IRowEditor_EndProcess(enumerator As Object)
    EndProcess enumerator
End Sub
'*-----------------------------------------------------------------------------
'* IRowEditor#CancelProcess
'*
'*-----------------------------------------------------------------------------
Public Sub IRowEditor_CancelProcess(enumerator As Object)
    CancelProcess enumerator
End Sub

