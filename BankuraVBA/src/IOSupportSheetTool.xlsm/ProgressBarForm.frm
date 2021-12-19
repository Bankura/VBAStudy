VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBarForm 
   Caption         =   "しばらくお待ちください…"
   ClientHeight    =   1456
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   4298
   OleObjectBlob   =   "ProgressBarForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ProgressBarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] 進捗バーフォーム
'* [詳  細] 進捗状況を表示する進捗バーフォーム。
'*          FormProgressReporterクラスから使用する想定。
'* [参  考]
'*          https://excel-ubara.com/excelvba3/EXCELFORM026.html
'*
'* @author Bankura
'* Copyright (c) 2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Public IsCancel As Boolean
Private mProgressBarLabel As MSForms.Label
Private mMaxValue As Long
Private mBarColor As Long
Private mCurValue As Double  ' プログレスバー現在値
Private mInteractive As Long
Private mSelfDoEvents As Boolean

'******************************************************************************
'* プロパティ定義
'******************************************************************************

'*-----------------------------------------------------------------------------
'* MaxValue プロパティ
'*
'* 最大値プロパティ
'*-----------------------------------------------------------------------------
Public Property Let MaxValue(arg As Long)
    If arg = 0 Then Exit Property
    mMaxValue = arg
End Property
Public Property Get MaxValue() As Long
    MaxValue = mMaxValue
End Property

'*-----------------------------------------------------------------------------
'* BarColor プロパティ
'*
'* プログレスバーの色指定
'*-----------------------------------------------------------------------------
Public Property Let BarColor(arg As Long)
    mBarColor = arg
    mProgressBarLabel.BackColor = mBarColor
End Property

'*-----------------------------------------------------------------------------
'* Interactive プロパティ
'*
'* 割込み拒否指定（False: 拒否）
'*-----------------------------------------------------------------------------
Public Property Let Interactive(arg As Boolean)
    mInteractive = arg
End Property

'*-----------------------------------------------------------------------------
'* SelfDoEvents プロパティ
'*
'* Form自身の処理でDoEventsを呼び出すか
'*-----------------------------------------------------------------------------
Public Property Let SelfDoEvents(arg As Boolean)
    mSelfDoEvents = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub UserForm_Initialize()
    mBarColor = rgb(0, 0, 128)
    mInteractive = True
    IsCancel = False
    mSelfDoEvents = True
    
    mMaxValue = 100
    mCurValue = 0

    ' ラベルコントロール追加
    Set mProgressBarLabel = Me.ProgressBar.Controls.Add("Forms.Label.1", "lblProgress")
    mProgressBarLabel.Width = 0
    mProgressBarLabel.Height = Me.ProgressBar.Height
    mProgressBarLabel.BackColor = mBarColor

    ' プログレスバーの背景をへこませる
    Me.ProgressBar.SpecialEffect = fmSpecialEffectSunken
    
    Me.Caption = "しばらくお待ちください…"
    Me.ProgressText.Caption = ""
    
End Sub
Private Sub UserForm_Terminate()
    If mInteractive = False Then
        Application.Interactive = True
        Application.EnableCancelKey = xlInterrupt
    End If
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************
'******************************************************************************
'* [概  要] ShowModeless
'* [詳  細] フォームをモードレスで表示する。
'*
'* @param formCaptionTxt フォームのタイトルテキスト
'******************************************************************************
Public Sub ShowModeless(Optional ByVal formCaptionTxt As String)
    ' 割込み拒否の設定
    If mInteractive = False Then
        Me.Enabled = False
        Application.Interactive = False
        Application.EnableCancelKey = xlDisabled
    End If
  
    ' フォームをモードレスで表示
    If formCaptionTxt <> "" Then
        Me.Caption = formCaptionTxt
    End If
    Me.Show vbModeless
End Sub

'******************************************************************************
'* [概  要] SetProgressValue
'* [詳  細] プログレスバーの進捗を指定した値で更新する。
'*
'* @param aValue 進捗値（指定値）
'* @param progressTxt 進捗表示のテキスト
'* @param formCaptionTxt フォームのタイトルテキスト
'******************************************************************************
Public Sub SetProgressValue(ByVal aValue As Double, Optional ByVal progressTxt As String, Optional ByVal formCaptionTxt As String)
    mCurValue = aValue
  
    ' 最大値を超えないように調整
    If mCurValue > mMaxValue Then
        mCurValue = mMaxValue
    End If
  
    ' プログレスバーの描画
    mProgressBarLabel.BackColor = mBarColor
    mProgressBarLabel.Width = Me.ProgressBar.Width * (mCurValue / mMaxValue)
    If formCaptionTxt <> "" Then
        Me.Caption = formCaptionTxt
    End If
    If progressTxt <> "" Then
        Me.ProgressText.Caption = progressTxt
    End If
    
    ' 再描画
    Me.Repaint
    If mSelfDoEvents Then
        UXUtils.CheckEvents
    End If
End Sub

'******************************************************************************
'* [概  要] AddProgressValue
'* [詳  細] プログレスバーの進捗を指定した値で加算し、更新する。
'*
'* @param aValue 進捗値（加算値）
'* @param progressTxt 進捗表示のテキスト
'* @param formCaptionTxt フォームのタイトルテキスト
'******************************************************************************
Public Sub AddProgressValue(ByVal aValue As Double, Optional ByVal progressTxt As String, Optional ByVal formCaptionTxt As String)
    mCurValue = mCurValue + aValue
    Call SetProgressValue(mCurValue, formCaptionTxt)
End Sub

'******************************************************************************
'* [概  要] Unload
'* [詳  細] 自身のフォームをUnloadする。
'*
'******************************************************************************
Public Sub Unload()
    VBA.Unload Me
End Sub

'******************************************************************************
'* イベント処理
'******************************************************************************

'******************************************************************************
'* [概  要] UserForm：QueryClose イベント処理
'* [詳  細] フォームが閉じる前に発生するQueryClose イベントの処理。
'*          「×」ボタン等のユーザ操作によりフォームが閉じられる際に、
'*          割込み拒否指定プロパティが「True」（割り込みを許容）である場合は、
'*          処理の中断を確認するダイアログを表示する。
'*          ダイアログにて「はい」選択時は、IsCancelプロパティを「True」（中断）
'*          に設定し、フォームを閉じる。
'*          割込み拒否指定プロパティが「False」（割り込みを拒否）である場合、
'*          ダイアログにて「いいえ」選択時は、フォームを閉じる処理をキャンセル
'*          する。
'* [参  考] https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/queryclose-event
'*
'* @param Cancel イベントをキャンセルするか。この 引数を 0 以外の値に設定すると
'*               読込済のすべてのユーザーフォームで QueryClose イベントが停止さ
'*               れ、UserForm とアプリケーションを閉じる処理がキャンセルされる。
'* @param CloseMode QueryClose イベントの原因を示す値または 定数
'******************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        If mInteractive Then
            If MsgBox("処理を中断しますか?", vbYesNo, "中断確認") = vbYes Then
                IsCancel = True
                Exit Sub
            End If
        End If
        Cancel = True
    End If
End Sub
