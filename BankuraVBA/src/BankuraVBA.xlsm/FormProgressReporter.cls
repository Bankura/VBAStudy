VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormProgressReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IProgressReporter

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] 進捗バーフォーム進捗表示クラス
'* [詳  細] 進捗バーフォームに進捗状況を表示する。
'*
'* [参  考]
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mProgressBarForm As ProgressBarForm
Private mBaseMessage As String
Private mProgressString As String
Private mCountVisible As Boolean

'******************************************************************************
'* プロパティ定義
'******************************************************************************

'*-----------------------------------------------------------------------------
'* BaseMessage プロパティ
'*
'* 進捗表示メッセージの固定部分（「CSV読込処理中」等）
'*-----------------------------------------------------------------------------
Public Property Let BaseMessage(arg As String)
    mBaseMessage = arg
End Property

'*-----------------------------------------------------------------------------
'* CountVisible プロパティ
'*
'* 件数表示をするか
'*-----------------------------------------------------------------------------
Public Property Let CountVisible(arg As Boolean)
    mCountVisible = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mProgressBarForm = New ProgressBarForm
    mBaseMessage = "処理中"
    mCountVisible = True
End Sub
Private Sub Class_Terminate()
    Set mProgressBarForm = Nothing
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************
'******************************************************************************
'* [概  要] Start
'* [詳  細] 開始時の表示処理を行う。
'*
'* @param enumerator IEnumeratorオブジェクト
'******************************************************************************
Public Sub Start(enumerator As IEnumerator)
    With mProgressBarForm
        .ProcessTitle = mBaseMessage
        With .ProgressBar
            .Min = 0
            .Max = enumerator.MaxReportCount
            .Value = 0
        End With
        .ProgressText.Caption = GetProgressMessage(enumerator)
        .Show vbModeless
    End With
End Sub

'******************************************************************************
'* [概  要] Finish
'* [詳  細] 終了時の表示処理を行う。
'*
'* @param enumerator IEnumeratorオブジェクト
'******************************************************************************
Public Sub Finish(enumerator As IEnumerator)
    mProgressBarForm.Hide
End Sub

'******************************************************************************
'* [概  要] Proceed
'* [詳  細] 進捗が進んだ際の表示処理を行う。
'*
'* @param enumerator IEnumeratorオブジェクト
'******************************************************************************
Public Sub Proceed(enumerator As IEnumerator)
    With mProgressBarForm
         .ProgressBar.Value = enumerator.CurrentReportCount
         .ProgressText.Caption = GetProgressMessage(enumerator)
         .Repaint
    End With
End Sub

'******************************************************************************
'* [概  要] GetProgressMessage
'* [詳  細] 進捗メッセージを取得する。
'*
'* @param enumerator IEnumeratorオブジェクト
'* @return String 進捗メッセージ
'******************************************************************************
Private Function GetProgressMessage(enumerator As IEnumerator) As String
    Dim msg As New StringEx
    If mCountVisible Then
        msg.Append enumerator.CurrentReportCount
        msg.Append " / "
        msg.Append enumerator.MaxReportCount
        msg.Append " 件"
    End If
    GetProgressMessage = msg.ToString
End Function

'******************************************************************************
'* インタフェース実装
'******************************************************************************
'*-----------------------------------------------------------------------------
'* IProgressReporter#Start
'*
'*-----------------------------------------------------------------------------
Private Sub IProgressReporter_Start(enumerator As IEnumerator)
    Me.Start enumerator
End Sub
'*-----------------------------------------------------------------------------
'* IProgressReporter#Finish
'*
'*-----------------------------------------------------------------------------
Private Sub IProgressReporter_Finish(enumerator As IEnumerator)
    Me.Finish enumerator
End Sub
'*-----------------------------------------------------------------------------
'* IProgressReporter#Proceed
'*
'*-----------------------------------------------------------------------------
Private Sub IProgressReporter_Proceed(enumerator As IEnumerator)
    Me.Proceed enumerator
End Sub

