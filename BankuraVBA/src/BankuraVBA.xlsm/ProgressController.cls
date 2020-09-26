VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] 繰り返し処理進捗管理クラス
'* [詳  細] IEnumeratorインタフェースを実装した繰返しデータ処理を行うクラス
'*          について、処理をスタートさせ、進捗イベントを受け取る。
'*          進捗状況を、プレゼンテーション用のIProgressReporterインタフェース
'*          を実装したクラスで表示させる。
'*
'* [参  考] https://qiita.com/kazu3516/items/1223ce0838dd4a8d2629
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private WithEvents eventObject As EnumeratorEventObject
Attribute eventObject.VB_VarHelpID = -1
Private mProgressReporter As IProgressReporter
Private mIsReportable As Long

'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* ProgressReporter プロパティ
'*
'* 進捗レポートオブジェクト
'*-----------------------------------------------------------------------------
Public Property Get ProgressReporter() As IProgressReporter
    Set ProgressReporter = mProgressReporter
End Property
Public Property Set ProgressReporter(arg As IProgressReporter)
    Set mProgressReporter = arg
End Property

'*-----------------------------------------------------------------------------
'* IsReportable プロパティ
'*
'* 進捗のリポートを行うかどうか。
'*-----------------------------------------------------------------------------
Public Property Get IsReportable() As Boolean
    IsReportable = mIsReportable
End Property
Public Property Let IsReportable(arg As Boolean)
    mIsReportable = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    mIsReportable = True
End Sub
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Start
'* [詳  細] 繰り返し処理を開始する。
'*
'* @param enumerator 繰り返し処理実装オブジェクト
'*
'******************************************************************************
Public Sub Start(enumerator As IEnumerator)
    Set eventObject = enumerator.eventObject
    
    ' 進捗表示開始
    If Not mProgressReporter Is Nothing And mIsReportable Then
        mProgressReporter.Start enumerator
    End If

    ' 列挙開始
    enumerator.Start
    
    ' 進捗表示完了
    If Not mProgressReporter Is Nothing And mIsReportable Then
        mProgressReporter.Finish enumerator
    End If
End Sub

'******************************************************************************
'* [概  要] eventObject_OnAction
'* [詳  細] 進捗レポートごとに呼び出されるイベントハンドラ。
'*
'* @param sender 繰り返し処理実装オブジェクト
'*
'******************************************************************************
Private Sub eventObject_OnAction(sender As IEnumerator)
    If Not mProgressReporter Is Nothing And mIsReportable Then
        mProgressReporter.Proceed sender
    End If
    DoEvents
End Sub
