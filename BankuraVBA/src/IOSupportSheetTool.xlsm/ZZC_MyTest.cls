VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZZC_MyTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private WithEvents rowEnumerator1 As RowEnumerator
Private WithEvents rowEnumerator1 As XlRowEnumerator
Attribute rowEnumerator1.VB_VarHelpID = -1
Public Sub Main()
    ' RowEnumeratorの設定
    'Set rowEnumerator1 = New RowEnumerator
    Set rowEnumerator1 = Init(New XlRowEnumerator, ThisWorkbook, "Sheet1", 4, 2)
'    rowEnumerator1.StartRow = 1
'    rowEnumerator1.EndRow = 10000
'    rowEnumerator1.ReportUnit = 500
'    rowEnumerator1.IsActionable = True
'    rowEnumerator1.MaxReportCount = 10000
'    rowEnumerator1.SourceData = ArrayUtils.ArrayRange(1, 10000)
    
    Dim pRowEditor As RowDebugPrinter: Set pRowEditor = New RowDebugPrinter
    Set rowEnumerator1.RowEditor = pRowEditor
    
    ' ProgressControllerの設定
    Dim pController As ProgressController: Set pController = New ProgressController
    Dim pReporter As SBProgressReporter: Set pReporter = New SBProgressReporter
    Set pController.ProgressReporter = pReporter
    pController.IsReportable = True

    ' 処理開始
    pController.Start rowEnumerator1
End Sub

'Private Sub rowEnumerator1_StartAction(sender As RowEnumerator)
Private Sub rowEnumerator1_StartAction(sender As XlRowEnumerator)
    Debug.Print "★StartAction"
End Sub

'Private Sub rowEnumerator1_RowProcessAction(sender As RowEnumerator)
Private Sub rowEnumerator1_RowProcessAction(sender As XlRowEnumerator)
    Debug.Print "★RowProcessAction"
    DebugUtils.Print2DArrayRow sender.SourceData, sender.CurrentReportCount
End Sub

'Private Sub rowEnumerator1_EndAction(sender As RowEnumerator)
Private Sub rowEnumerator1_EndAction(sender As XlRowEnumerator)
    Debug.Print "★EndAction"
End Sub


