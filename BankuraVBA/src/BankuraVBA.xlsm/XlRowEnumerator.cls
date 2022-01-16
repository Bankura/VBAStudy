VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XlRowEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IEnumerator

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Excel行繰り返し処理クラス
'* [詳  細] Excel行毎に繰り返し処理を行う。
'*
'* [参  考] https://qiita.com/kazu3516/items/1223ce0838dd4a8d2629
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mSheet As Worksheet
Private mSourceData As Variant
Private mDestTarget As Variant
Private mCurrentRow As Long
Private mEventObject As EnumeratorEventObject
Private mStartRow As Long
Private mEndRow As Long
Private mStartCol As Long
Private mEndCol As Long
Private mReportUnit As Long
Private mCurrentReportCount As Long
Private mMaxReportCount As Long
Private mIsActionable As Long
Private mRowProcessFunction As Func
Private mStartProcessFunction As Func
Private mEndProcessFunction As Func
Private mCancelProcessFunction As Func
Private mRowEditor As IRowEditor
Private mIsCancel As Boolean
Private mIsError As Boolean

'******************************************************************************
'* イベント定義
'******************************************************************************
Public Event StartAction(sender As XlRowEnumerator)
Public Event RowProcessAction(sender As XlRowEnumerator)
Public Event EndAction(sender As XlRowEnumerator)
Public Event CancelAction(sender As RowEnumerator)

'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* Target プロパティ
'*
'* 対象シート
'*-----------------------------------------------------------------------------
Public Property Set Target(arg As Worksheet)
    Set mSheet = arg
End Property

'*-----------------------------------------------------------------------------
'* SourceData プロパティ
'*
'* 元となるデータ
'*-----------------------------------------------------------------------------
Public Property Get SourceData() As Variant
    If IsObject(mSourceData) Then
        Set SourceData = mSourceData
    Else
        SourceData = mSourceData
    End If
End Property
Public Property Let SourceData(arg As Variant)
    If IsObject(arg) Then
        Set mSourceData = arg
    Else
        mSourceData = arg
    End If
End Property

'*-----------------------------------------------------------------------------
'* DestTarget プロパティ
'*
'* 出力先
'*-----------------------------------------------------------------------------
Public Property Get DestTarget() As Variant
    If IsObject(mSourceData) Then
        Set DestTarget = mDestTarget
    Else
        DestTarget = mDestTarget
    End If
End Property
Public Property Let DestTarget(arg As Variant)
    If IsObject(arg) Then
        Set mDestTarget = arg
    Else
        mDestTarget = arg
    End If
End Property

'*-----------------------------------------------------------------------------
'* CurrentRow プロパティ
'*
'* 現在行番号
'*-----------------------------------------------------------------------------
Public Property Get CurrentRow() As Long
    CurrentRow = mCurrentRow
End Property

'*-----------------------------------------------------------------------------
'* StartRow プロパティ
'*
'* 開始行番号
'*-----------------------------------------------------------------------------
Public Property Get startRow() As Long
    startRow = mStartRow
End Property
Public Property Let startRow(arg As Long)
    mStartRow = arg
End Property

'*-----------------------------------------------------------------------------
'* EndRow プロパティ
'*
'* 終了行番号
'*-----------------------------------------------------------------------------
Public Property Get endRow() As Long
    endRow = mEndRow
End Property
Public Property Let endRow(arg As Long)
    mEndRow = arg
End Property

'*-----------------------------------------------------------------------------
'* StartCol プロパティ
'*
'* 開始列番号
'*-----------------------------------------------------------------------------
Public Property Get startCol() As Long
    startCol = mStartCol
End Property
Public Property Let startCol(arg As Long)
    mStartCol = arg
End Property

'*-----------------------------------------------------------------------------
'* EndCol プロパティ
'*
'* 終了列番号
'*-----------------------------------------------------------------------------
Public Property Get endCol() As Long
    endCol = mEndCol
End Property
Public Property Let endCol(arg As Long)
    mEndCol = arg
End Property

'*-----------------------------------------------------------------------------
'* ReportUnit プロパティ
'*
'* 進捗単位。進捗のリポートを行う（イベントを発生させる）単位の処理数。
'*-----------------------------------------------------------------------------
Public Property Get ReportUnit() As Long
    ReportUnit = mReportUnit
End Property
Public Property Let ReportUnit(arg As Long)
    mReportUnit = arg
End Property

'*-----------------------------------------------------------------------------
'* CurrentReportCount プロパティ
'*
'* 現在の進捗処理数
'*-----------------------------------------------------------------------------
Public Property Get CurrentReportCount() As Long
    CurrentReportCount = mCurrentReportCount
End Property
Public Property Let CurrentReportCount(arg As Long)
    mCurrentReportCount = arg
End Property

'*-----------------------------------------------------------------------------
'* MaxReportCount プロパティ
'*
'* 現在の進捗処理数
'*-----------------------------------------------------------------------------
Public Property Get MaxReportCount() As Long
    MaxReportCount = mMaxReportCount
End Property
Public Property Let MaxReportCount(arg As Long)
    mMaxReportCount = arg
End Property

'*-----------------------------------------------------------------------------
'* IsActionable プロパティ
'*
'* 行毎のActionイベントを発生させるかどうか。
'*-----------------------------------------------------------------------------
Public Property Get IsActionable() As Boolean
    IsActionable = mIsActionable
End Property
Public Property Let IsActionable(arg As Boolean)
    mIsActionable = arg
End Property

'*-----------------------------------------------------------------------------
'* RowProcessFunction プロパティ
'*
'* 行毎の処理を行うFunction。設定すると行毎に呼び出しが行われる。
'*-----------------------------------------------------------------------------
Public Property Get RowProcessFunction() As Func
    RowProcessFunction = mRowProcessFunction
End Property
Public Property Set RowProcessFunction(arg As Func)
    Set mRowProcessFunction = arg
End Property

'*-----------------------------------------------------------------------------
'* StartProcessFunction プロパティ
'*
'* 開始時の処理を行うFunction。設定すると開始時に呼び出しが行われる。
'*-----------------------------------------------------------------------------
Public Property Get StartProcessFunction() As Func
    StartProcessFunction = mStartProcessFunction
End Property
Public Property Set StartProcessFunction(arg As Func)
    Set mStartProcessFunction = arg
End Property

'*-----------------------------------------------------------------------------
'* EndProcessFunction プロパティ
'*
'* 終了時の処理を行うFunction。設定すると終了時に呼び出しが行われる。
'*-----------------------------------------------------------------------------
Public Property Get EndProcessFunction() As Func
    EndProcessFunction = mEndProcessFunction
End Property
Public Property Set EndProcessFunction(arg As Func)
    Set mEndProcessFunction = arg
End Property

'*-----------------------------------------------------------------------------
'* RowEditor プロパティ
'*
'* 行毎の処理を行うオブジェクト。設定すると行毎に呼び出しが行われる。
'* 開始・終了時の処理でも対応するメソッドの呼び出しが行われる。
'*-----------------------------------------------------------------------------
Public Property Get RowEditor() As IRowEditor
    RowEditor = mRowEditor
End Property
Public Property Set RowEditor(arg As IRowEditor)
    Set mRowEditor = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mEventObject = New EnumeratorEventObject
    mIsActionable = True
    mReportUnit = 1
    mCurrentReportCount = 0
    mMaxReportCount = 0
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************
'******************************************************************************
'* [概  要] Init メソッド
'* [詳  細] 初期化用メソッド｡
'*
'* @param wb ワークブック
'* @param sheetName シート名
'* @param startRowNum 開始行番号
'* @param startColNum 開始列番号
'* @return WorkSheetExオブジェクト
'*
'******************************************************************************
Public Function Init(wb As Workbook, ByVal sheetName As String, _
                     ByVal startRowNum As Long, ByVal startColNum As Long, _
                     Optional ByVal endRowNum As Long, Optional ByVal endColNum As Long) As XlRowEnumerator
    Set mSheet = wb.Worksheets(sheetName)
    mStartRow = startRowNum
    mStartCol = startColNum
    
    If endRowNum = 0 Then
        mEndRow = XlWorkSheetUtils.GetFinalKeyRow(mSheet, startColNum)
    Else
        mEndRow = endRowNum
    End If
    If endColNum = 0 Then
        mEndCol = XlWorkSheetUtils.GetFinalKeyCol(mSheet, startRowNum)
    Else
        mEndCol = endColNum
    End If
    Set Init = Me
End Function

'******************************************************************************
'* [概  要] SetStartPosition
'* [詳  細] 開始位置を設定する。
'*
'******************************************************************************
Sub SetStartPosition(ByVal startRowNum As Long, ByVal startColNum As Long)
    mStartRow = startRowNum
    mStartCol = startColNum
End Sub

'******************************************************************************
'* [概  要] Start
'* [詳  細] 行毎に繰り返し処理を行う。
'*
'******************************************************************************
Public Sub Start()
    ' 開始時の処理
    Call StartProcess
        
    ' 繰り返し処理
    Dim i As Long
    For i = mStartRow To mEndRow
        mCurrentReportCount = mCurrentReportCount + 1
        mCurrentRow = i
        
        ' 行処理
        Call RowProcess
        
        ' 進捗レポート
        Call ReportProgress(i)
        
        If mIsCancel Then
            Call CancelProcess
            Exit Sub
        ElseIf mIsError Then
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    Next
    
    ' 終了時の処理
    Call EndProcess
End Sub

'******************************************************************************
'* [概  要] StartProcess
'* [詳  細] 開始時の初期処理を行う（設定・読込処理等）。
'*
'******************************************************************************
Private Sub StartProcess()
    mSourceData = XlWorkSheetUtils.GetVariantDataFromSheet(mSheet, mStartRow, mStartCol)
    mDestTarget = ArrayUtils.Create2DArrayFromPreExisting(mSourceData)
    mMaxReportCount = ArrayUtils.GetLength(mSourceData)

    If IsActionable Then
        RaiseEvent StartAction(Me)
    End If
    If Not mStartProcessFunction Is Nothing Then
        Call mStartProcessFunction.Apply(Me)
    End If
    If Not mRowEditor Is Nothing Then
        Call mRowEditor.StartProcess(Me)
    End If
End Sub

'******************************************************************************
'* [概  要] RowProcess
'* [詳  細] 行に対する処理を行う（編集処理等）。
'*
'******************************************************************************
Private Sub RowProcess()
    If IsActionable Then
        RaiseEvent RowProcessAction(Me)
    End If
    If Not mRowProcessFunction Is Nothing Then
        Call mRowProcessFunction.Apply(Me)
    End If
    If Not mRowEditor Is Nothing Then
        Call mRowEditor.Edit(Me)
    End If
End Sub

'******************************************************************************
'* [概  要] ReportProgress
'* [詳  細] 進捗のリポート処理（イベント通知）を行う。
'*
'******************************************************************************
Private Sub ReportProgress(i As Long)
    If i Mod mReportUnit = 0 Then
        mEventObject.Raise Me
        DoEvents
    End If
End Sub

'******************************************************************************
'* [概  要] EndProcess
'* [詳  細] 終了時の処理を行う（出力処理等）。
'*
'******************************************************************************
Private Sub EndProcess()
    If IsActionable Then
        RaiseEvent EndAction(Me)
    End If
    If Not mEndProcessFunction Is Nothing Then
        Call mEndProcessFunction.Apply(Me)
    End If
    If Not mRowEditor Is Nothing Then
        Call mRowEditor.EndProcess(Me)
    End If
End Sub

'******************************************************************************
'* [概  要] CancelProcess
'* [詳  細] キャンセル時の処理を行う（出力処理等）。
'*
'******************************************************************************
Private Sub CancelProcess()
    If IsActionable Then
        RaiseEvent CancelAction(Me)
    End If
    If Not mCancelProcessFunction Is Nothing Then
        Call mCancelProcessFunction.Apply(Me)
    End If
    If Not mRowEditor Is Nothing Then
        Call mRowEditor.CancelProcess(Me)
    End If
End Sub

'******************************************************************************
'* インタフェース実装
'******************************************************************************
'*-----------------------------------------------------------------------------
'* IEnumerator#EventObject
'*
'*-----------------------------------------------------------------------------
Private Property Get IEnumerator_EventObject() As EnumeratorEventObject
    Set IEnumerator_EventObject = mEventObject
End Property

'*-----------------------------------------------------------------------------
'* IEnumerator#CurrentReportCount
'*
'*-----------------------------------------------------------------------------
Private Property Get IEnumerator_CurrentReportCount() As Long
    IEnumerator_CurrentReportCount = mCurrentReportCount
End Property
Public Property Let IEnumerator_CurrentReportCount(arg As Long)
    mCurrentReportCount = arg
End Property

'*-----------------------------------------------------------------------------
'* IEnumerator#MaxReportCount
'*
'*-----------------------------------------------------------------------------
Private Property Get IEnumerator_MaxReportCount() As Long
    IEnumerator_MaxReportCount = mMaxReportCount
End Property
Public Property Let IEnumerator_MaxReportCount(arg As Long)
    mMaxReportCount = arg
End Property

'*-----------------------------------------------------------------------------
'* IEnumerator#IsCancel
'*
'*-----------------------------------------------------------------------------
Private Property Get IEnumerator_IsCancel() As Boolean
    IEnumerator_IsCancel = mIsCancel
End Property
Private Property Let IEnumerator_IsCancel(arg As Boolean)
    mIsCancel = arg
End Property

'*-----------------------------------------------------------------------------
'* IEnumerator#IsError
'*
'*-----------------------------------------------------------------------------
Private Property Get IEnumerator_IsError() As Boolean
    IEnumerator_IsError = mIsError
End Property
Private Property Let IEnumerator_IsError(arg As Boolean)
    mIsError = arg
End Property

'*-----------------------------------------------------------------------------
'* IEnumerator#Start
'*
'*-----------------------------------------------------------------------------
Private Sub IEnumerator_Start()
    Me.Start
End Sub
