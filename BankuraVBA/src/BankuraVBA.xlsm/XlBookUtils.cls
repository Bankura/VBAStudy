VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XlBookUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Excel Bookユーティリティクラス
'* [詳  細] Excel Bookに関するユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'* [参  考]
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
'******************************************************************************
'* [概  要] ChangeExcelFiles2007Format メソッド
'* [詳  細] 指定フォルダから、2003形式Excelファイルを開き、Office2007形式
'*          フォーマット（xlsxまたはxlsm）で保存する。
'*
'* @param srcPath Excelファイルが格納されているフォルダ
'* @param destPath 保存先フォルダ
'* @param format Excelフォーマット（xlOpenXMLWorkbook）
'******************************************************************************
Public Sub ChangeExcelFiles2007Format(srcPath As String, destPath As String)
    Dim strXlsFileName As String
    strXlsFileName = Dir(FileUtils.AddPathSeparator(srcPath) & "*.xls*")

    Do Until strXlsFileName = ""
        If LCase(Right(strXlsFileName, 3)) = "xls" Then
            Workbooks.Open fileName:=FileUtils.AddPathSeparator(srcPath) & strXlsFileName
            Application.DisplayAlerts = False
            If ActiveWorkbook.HasVBProject = True Then
                ActiveWorkbook.SaveAs fileName:=FileUtils.AddPathSeparator(destPath) & strXlsFileName & "m", FileFormat:=xlOpenXMLWorkbookMacroEnabled
            Else
                ActiveWorkbook.SaveAs fileName:=FileUtils.AddPathSeparator(destPath) & strXlsFileName & "x", FileFormat:=xlOpenXMLWorkbook
            End If
            ActiveWorkbook.Close
            Application.DisplayAlerts = True
        End If
        strXlsFileName = Dir()
    Loop
End Sub

'******************************************************************************
'* [概  要] ChangeExcelFilesFormatXlsx メソッド
'* [詳  細] 指定フォルダから、Excelファイルを開き、xlsx形式で
'*          保存する。
'*
'* @param srcPath Excelファイルが格納されているフォルダ
'* @param destPath 保存先フォルダ
'******************************************************************************
Public Sub ChangeExcelFilesFormatXlsx(srcPath As String, destPath As String)
    Dim strXlsFileName As String
    strXlsFileName = Dir(FileUtils.AddPathSeparator(srcPath) & "*.xls*")

    Do Until strXlsFileName = ""
        Call ChangeExcelFileFormatXlsx(strXlsFileName, srcPath, destPath)
        strXlsFileName = Dir()
    Loop
End Sub

'******************************************************************************
'* [概  要] ChangeExcelFileFormatXlsx メソッド
'* [詳  細] Excelファイルを開き、xlsx形式で保存する。
'*
'* @param xlsFileName Excelファイル名
'* @param srcPath Excelファイルが格納されているフォルダ
'* @param destPath 保存先フォルダ
'******************************************************************************
Public Sub ChangeExcelFileFormatXlsx(xlsFileName As String, srcPath As String, destPath As String)
    Workbooks.Open fileName:=FileUtils.AddPathSeparator(srcPath) & xlsFileName

    Dim displayAlertsState As Boolean: displayAlertsState = Application.DisplayAlerts
    If Application.DisplayAlerts Then Application.DisplayAlerts = False

    ActiveWorkbook.SaveAs fileName:=FileUtils.AddPathSeparator(destPath) & Split(xlsFileName, ".")(0) & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = displayAlertsState
End Sub

'******************************************************************************
'* [概  要] ExtractKeyData
'* [詳  細] 指定したシートのデータ(1列)を抽出し、別シートA列に貼付ける。
'*          シート名とデータが記載されている列の先頭セルの行列番号を指定
'*          することにより対象の先頭セルから最後のセルまでのデータを抽出
'*          する。空行と重複するデータは除外する。
'*
'* @param inputSheetName 読込元シート名
'* @param startRow 読込列の先頭セル行番号
'* @param startCol 読込列の先頭セル列番号
'* @param outputSheetName 出力先シート名
'* @return Boolean 処理結果（True:正常 False：異常）
'*
'******************************************************************************
Public Function ExtractKeyData(inputSheetName As String, StartRow As Integer, StartCol As Integer, _
                           outputSheetName As String) As Boolean
           
    Dim myDic As Object, myKey As Variant
    Dim c As Variant, varData As Variant
    Set myDic = CreateObject("Scripting.Dictionary")
    
    ' 入力情報を取得
    With Worksheets(inputSheetName)
        varData = .Range(.Cells(StartRow, StartCol), .Cells(Rows.count, StartCol).End(xlUp)).value
    End With

    '空データ、重複データを除外
    If IsArray(varData) Then
        For Each c In varData
            If Not c = Empty Then
                If Not myDic.Exists(c) Then
                    myDic.Add c, Null
                End If
            End If
        Next
    Else
        myDic.Add varData, Null
    End If
    myKey = myDic.keys
    
    With Worksheets(outputSheetName)
        '読込先のシートをクリア
        .Range("A:A").ClearContents

        '読込先シートにデータを出力
        .Range("A1").Resize(myDic.count) = Application.WorksheetFunction.Transpose(myKey)
        
        'ソート処理
        .Range("A1").Resize(myDic.count).Sort Key1:=.Range("A1")
    End With
    Set myDic = Nothing
    ExtractKeyData = True
    Exit Function
End Function

'******************************************************************************
'* [概  要] ExtractKeyData2
'* [詳  細] 指定したシートのデータ(1列)を抽出し、別シートA列に貼付ける。
'*          シート名とデータが記載されている列の先頭セルの行列番号を指定
'*          することにより対象の先頭セルから最後のセルまでのデータを抽出
'*          する。空行と重複するデータは除外する。
'*
'* @param inputSheet 読込元シート
'* @param startRow 読込列の先頭セル行番号
'* @param startCol 読込列の先頭セル列番号
'* @return Variant データ配列
'*
'******************************************************************************
Public Function ExtractKeyData2(inputSheet As Workbook, StartRow As Integer, StartCol As Integer) As Variant
    On Error GoTo ErrorHandler
           
    Dim myDic As Object, myKey As Variant
    Dim c As Variant, varData As Variant
    Set myDic = CreateObject("Scripting.Dictionary")
    
    ' 入力情報を取得
    With inputSheet
        varData = .Range(.Cells(StartRow, StartCol), .Cells(Rows.count, StartCol).End(xlUp)).value
    End With

    '空データ、重複データを除外
    If IsArray(varData) Then
        For Each c In varData
            If Not c = Empty Then
                If Not myDic.Exists(c) Then
                    myDic.Add c, Null
                End If
            End If
        Next
    Else
        myDic.Add varData, Null
    End If
    myKey = myDic.keys
    ExtractKeyData2 = Application.WorksheetFunction.Transpose(myKey)
    Set myDic = Nothing

    Exit Function
ErrorHandler:
    ExtractKeyData2 = False
End Function

'******************************************************************************
'* [概  要] IsBookOpened
'* [詳  細] オープン済みブックの存在チェック。
'*
'* @param bookName ブック
'* @return Boolean 判定結果（True:存在 False：存在しない）
'*
'******************************************************************************
Public Function IsBookOpened(bookName As String) As Boolean
    Dim wb As Workbook

    ' Excel起動チェック
    For Each wb In Application.Workbooks
        If wb.Name = Dir(bookName) Then
            IsBookOpened = False
            Exit Function
        End If
    Next
    IsBookOpened = True
End Function
