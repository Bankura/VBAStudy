VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XlWorkSheetUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Excel WorkSheet Cell/Rangeユーティリティクラス
'* [詳  細] Excel WorkSheet Cell/Rangeに関するユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'* [参  考]
'*
'* @author Bankura
'* Copyright (c) 2019-2021 Bankura
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
'* [概  要] シート貼り付け処理。
'* [詳  細] Variant配列データをシートに出力する。
'*
'* @param dataSheet ワークシート
'* @param vArray Variant配列データ
'* @param lStartRow データ開始行番号
'* @param lStartCol データ開始列番号
'*
'******************************************************************************
Public Sub InjectVariantArrayToCells(ByVal dataSheet As Worksheet, vArray, lStartRow As Long, lStartCol As Long)
    If ArrayUtils.IsEmptyArray(vArray) Then Exit Sub
    dataSheet.Cells(lStartRow, lStartCol).Resize(UBound(vArray, 1) + 1, UBound(vArray, 2) + 1).Value = vArray
End Sub

'******************************************************************************
'* [概  要] 項番設定処理。
'* [詳  細] 項番に連番を出力する。
'*
'* @param dataSheet ワークシート
'* @param lStartRow データ開始行番号
'* @param lStartCol データ開始列番号
'* @param rowNum 番号数
'*
'******************************************************************************
Public Sub InjectNumbersToIndexCells(ByVal dataSheet As Worksheet, lStartRow As Long, lStartCol As Long, rowNum As Long)
    If rowNum < 1 Then
        Exit Sub
    End If
    With dataSheet
        .Cells(lStartRow, lStartCol) = 1
        If rowNum > 1 Then
            .Cells(lStartRow, lStartCol).AutoFill _
              Destination:=.Range(.Cells(lStartRow, lStartCol), .Cells(lStartRow + rowNum - 1, lStartCol)), Type:=xlLinearTrend
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] シート保護解除処理。
'* [詳  細] シートの保護を解除する。
'*
'* @param sheetName シート名
'* @param sheetPassword 保護パスワード
'* @param wb ワークブックオブジェクト
'*
'******************************************************************************
Public Sub UnprotectSheet(sheetName As String, Optional sheetPassword As String, Optional wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    'シート保護解除
    If sheetPassword = "" Then
        wb.Sheets(sheetName).Unprotect
    Else
        wb.Sheets(sheetName).Unprotect Password:=sheetPassword
    End If
End Sub

'******************************************************************************
'* [概  要] シート保護処理。
'* [詳  細] シートの保護をする。
'*
'* @param sheetName シート名
'* @param sheetPassword 保護パスワード
'* @param wb ワークブックオブジェクト
'*
'******************************************************************************
Public Sub ProtectSheet(sheetName As String, Optional sheetPassword As String, Optional wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    With wb.Sheets(sheetName)
        .EnableOutlining = True  'アウトライン有効
        .EnableAutoFilter = True 'オートフィルタ有効
        
        'シート保護
        If sheetPassword = "" Then
            .Protect Contents:=True, UserInterfaceOnly:=True
        Else
            .Protect Contents:=True, UserInterfaceOnly:=True, Password:=sheetPassword
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] GotoSheet
'* [詳  細] ブックの指定したシート・アドレスへ移動する。
'*
'* @param sheetName 移動先シート名
'* @param strAddr 移動先セルのアドレス
'* @param wb ワークブックオブジェクト
'*
'******************************************************************************
Public Sub GotoSheet(sheetName As String, Optional strAddr As String = "A1", Optional wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    wb.Activate
    wb.Worksheets(sheetName).Select
    wb.Worksheets(sheetName).Range(strAddr).Activate
End Sub

'******************************************************************************
'* [概  要] InitializeWorkSheets
'* [詳  細] ワークシート初期化
'*
'* @param varMakeSheets 作成するシート名の配列
'* @param varFormSheets 作成するシートのフォーマットとなるシート名の配列
'* @param strAfShName     作成位置の基準となるシートの名称
'*                        （このシートの右側に作成する）
'* @return Boolean 処理結果（True:正常 False：異常）
'******************************************************************************
Public Function InitializeWorkSheets(varMakeSheets As Variant, _
                                     varFormSheets As Variant, _
                                     strAfShName As String) As Boolean
    On Error GoTo ErrorHandler
    If IsArray(varMakeSheets) <> IsArray(varFormSheets) Then GoTo ErrorHandler

    ' 配列ではない場合も正常に処理
    If Not IsArray(varMakeSheets) And Not IsArray(varFormSheets) Then
        If DeleteWorkSheet(varMakeSheets) = False Then GoTo ErrorHandler
        ThisWorkbook.Worksheets(varFormSheets).Copy After:=Worksheets(strAfShName)
        ActiveSheet.Name = varMakeSheets
        InitializeWorkSheets = True
        Exit Function
    End If
    
    If UBound(varMakeSheets) <> UBound(varFormSheets) Then GoTo ErrorHandler

    ' 作成するワークシートと同名のシートがあれば削除
    If DeleteWorkSheets(varMakeSheets) = False Then GoTo ErrorHandler
    
    ' ワークシートを作成
    Dim i As Long
    For i = LBound(varMakeSheets) To UBound(varMakeSheets)
        ThisWorkbook.Worksheets(varFormSheets(i)).Copy After:=Worksheets(strAfShName)
        ActiveSheet.Name = varMakeSheets(i)
    Next i

    InitializeWorkSheets = True
    Exit Function
    
ErrorHandler:
    InitializeWorkSheets = False
End Function

'******************************************************************************
'* [概  要] DeleteWorkSheets
'* [詳  細] ワークシート削除
'*
'* @param delSheetNames 削除するシート名の配列
'* @param wb ブックオブジェクト
'* @return Boolean 処理結果（True:正常 False：異常）
'*
'******************************************************************************
Public Function DeleteWorkSheets(delSheetNames As Variant, Optional ByRef wb As Workbook) As Boolean
    On Error GoTo ErrorHandler
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    If Not IsArray(delSheetNames) Then
        DeleteWorkSheets = Me.DeleteWorkSheet(delSheetNames, wb)
        Exit Function
    End If

    Dim ws As Worksheet
    Dim delSheetName As Variant
    For Each delSheetName In delSheetNames
        Call Me.DeleteWorkSheet(delSheetName, wb)
    Next

    DeleteWorkSheets = True
    Exit Function

ErrorHandler:
    DeleteWorkSheets = False
End Function

'******************************************************************************
'* [概  要] DeleteWorkSheet
'* [詳  細] ワークシート削除
'*
'* @param sheetName 削除するシート名
'* @param wb ブックオブジェクト
'* @return Boolean 処理結果（True:正常 False：異常）
'*
'******************************************************************************
Public Function DeleteWorkSheet(ByVal sheetName As String, Optional ByRef wb As Workbook) As Boolean
    On Error GoTo ErrorHandler
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    If Me.SheetExists(sheetName, wb) Then
        wb.Worksheets(sheetName).Delete
    End If
    DeleteWorkSheet = True
    Exit Function
ErrorHandler:
    DeleteWorkSheet = False
End Function

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
Public Function ExtractKeyData(inputSheetName As String, startRow As Integer, startCol As Integer, _
                           outputSheetName As String) As Boolean
           
    Dim myDic As Object, myKey As Variant
    Dim C As Variant, varData As Variant
    Set myDic = CreateObject("Scripting.Dictionary")
    
    ' 入力情報を取得
    With Worksheets(inputSheetName)
        varData = .Range(.Cells(startRow, startCol), .Cells(Rows.count, startCol).End(xlUp)).Value
    End With

    '空データ、重複データを除外
    If IsArray(varData) Then
        For Each C In varData
            If Not C = Empty Then
                If Not myDic.Exists(C) Then
                    myDic.Add C, Null
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
        .Range("A1").Resize(myDic.count).Sort key1:=.Range("A1")
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
Public Function ExtractKeyData2(inputSheet As Worksheet, startRow As Integer, startCol As Integer) As Variant
    On Error GoTo ErrorHandler
           
    Dim myDic As Object, myKey As Variant
    Dim C As Variant, varData As Variant
    Set myDic = CreateObject("Scripting.Dictionary")
    
    ' 入力情報を取得
    With inputSheet
        varData = .Range(.Cells(startRow, startCol), .Cells(Rows.count, startCol).End(xlUp)).Value
    End With

    '空データ、重複データを除外
    If IsArray(varData) Then
        For Each C In varData
            If Not C = Empty Then
                If Not myDic.Exists(C) Then
                    myDic.Add C, Null
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
'* [概  要] シート存在チェック処理
'* [詳  細] シートが存在するか判定する。
'*
'* @param sheetName シート名
'* @param wb ブックオブジェクト
'* @return Boolean 処理結果（True:存在する False：存在しない）
'*
'******************************************************************************
Public Function SheetExists(ByVal sheetName As String, Optional ByRef wb As Workbook) As Boolean
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    'シートの検索
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next
    
    SheetExists = False
End Function

'******************************************************************************
'* [概  要] SheetExistsNoOpen
'* [詳  細] シートが存在するかをブックを開かずに判定する
'*
'* @param bookFolderPath ブックが配置されているフォルダパス
'* @param bookName ブック名
'* @param sheetName シート名
'* @return Boolean 処理結果（True:存在する False：存在しない）
'*
'******************************************************************************
Public Function SheetExistsNoOpen(bookFolderPath As String, _
                                  bookName As String, _
                                  sheetName As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = ExecuteExcel4Macro("'" & bookFolderPath & "[" & bookName & "]" & sheetName & "'!R1C1")
    If CStr(v) = "エラー 2023" Then
        SheetExistsNoOpen = False
    Else
        SheetExistsNoOpen = True
    End If
End Function


'******************************************************************************
'* [概  要] 最終行取得処理。
'* [詳  細] WorksheetのUsedRangeを下から走査し、最終行番号を取得する｡
'*
'* @param dataSheet ワークシート
'* @param ignoreColnum 走査対象外の列番号
'* @return 最終行番号
'*
'******************************************************************************
Public Function GetFinalRow(ByVal dataSheet As Worksheet, Optional ignoreColnum As Long) As Long
    Dim ret As Long
    Dim i As Long, cnta As Long
    With dataSheet.UsedRange
        For i = .Rows.count To 1 Step -1
            cnta = WorksheetFunction.CountA(.Rows(i))
            If cnta > 0 Then
                If cnta <> 1 Then
                    ret = i
                    Exit For
                Else
                    If ignoreColnum > 0 Then
                        If .Cells(i, ignoreColnum) = "" Then
                            ret = i
                            Exit For
                        End If
                    Else
                        ret = i
                        Exit For
                    End If
                End If
            End If
        Next
        If ret > 0 Then
            ret = ret + .row - 1
        End If
    End With
    GetFinalRow = ret
End Function

'******************************************************************************
'* [概  要] 最終列取得処理。
'* [詳  細] WorksheetのUsedRangeを右から走査し、最終列番号を取得する｡
'*
'* @param dataSheet ワークシート
'* @return 最終列番号
'*
'******************************************************************************
Public Function GetFinalCol(ByVal dataSheet As Worksheet) As Long
    Dim ret As Long
    Dim i As Long
    With dataSheet.UsedRange
        For i = .Columns.count To 1 Step -1
            If WorksheetFunction.CountA(.Columns(i)) > 0 Then
                ret = i
                Exit For
            End If
        Next
        If ret > 0 Then
            ret = ret + .Column - 1
        End If
    End With
    GetFinalCol = ret
End Function

'******************************************************************************
'* [概  要] 指定列最終入力行取得処理。
'* [詳  細] 指定列の最終入力行番号を取得する｡
'*
'* @param dataSheet ワークシート
'* @param startRow 対象列の番号
'* @return 最終行番号
'*
'******************************************************************************
Public Function GetFinalKeyRow(ByVal dataSheet As Worksheet, startCol As Long) As Long
    GetFinalKeyRow = dataSheet.Cells(Rows.count, startCol).End(xlUp).row
End Function

'******************************************************************************
'* [概  要] 指定行最終入力列取得処理。
'* [詳  細] 指定行の最終入力列番号を取得する｡
'*
'* @param dataSheet ワークシート
'* @param startRow 対象行の番号
'* @return 最終列番号
'*
'******************************************************************************
Public Function GetFinalKeyCol(ByVal dataSheet As Worksheet, startRow As Long) As Long
    GetFinalKeyCol = dataSheet.Cells(startRow, Columns.count).End(xlToLeft).Column
End Function

'******************************************************************************
'* [概  要] 表情報取得処理。
'* [詳  細] worksheetの表から情報を取得し、Variant配列を返す｡
'*
'* @param dataSheet ワークシート
'* @param lStartRow データ開始行番号
'* @param lStartCol データ開始列番号
'* @param itemCount 項目列数
'*
'******************************************************************************
Public Function GetVariantDataFromSheet(dataSheet As Worksheet, lStartRow As Long, lStartCol As Long, Optional colCount As Long)
    Dim lMaxRow As Long: lMaxRow = GetFinalKeyRow(dataSheet, lStartCol)
    Dim lMaxCol As Long
    If colCount = 0 Then
        lMaxCol = GetFinalKeyCol(dataSheet, lStartRow)
    Else
        lMaxCol = lStartCol + colCount - 1
    End If
    
    'レコードが存在しない場合
    If lMaxRow < lStartRow Or lMaxCol < lStartCol Then
        GetVariantDataFromSheet = Empty
        Exit Function
    End If
    
    Dim vArr: vArr = dataSheet.Range(dataSheet.Cells(lStartRow, lStartCol), dataSheet.Cells(lMaxRow, lMaxCol))
    
    GetVariantDataFromSheet = vArr
End Function

'******************************************************************************
'* [概  要] 使用セル範囲クリア処理。
'* [詳  細] worksheetのデータ表の使用セル範囲をクリアする｡
'*
'* @param dataSheet data表ワークシート
'* @param lStartRow data表データ開始行番号
'* @param lStartCol data表データ開始列番号
'* @param itemCount 項目列数
'* @param ignoreColnum 走査対象外の列番号
'*
'******************************************************************************
Public Sub ClearActualUsedRangeFromSheet(dataSheet As Worksheet, _
                                         lStartRow As Long, _
                                         lStartCol As Long, _
                                         Optional colCount As Long, _
                                         Optional ignoreColnum As Long)
    Dim rng As Range
    Set rng = GetActualUsedRangeFromSheet(dataSheet, lStartRow, lStartCol, colCount, ignoreColnum)
    If rng Is Nothing Then
        Exit Sub
    End If
    rng.ClearContents
End Sub

'******************************************************************************
'* [概  要] 未使用範囲行削除処理。
'* [詳  細] worksheetのデータ表の未使用範囲行を削除する（UsedRangeを縮小）｡
'*
'* @param dataSheet data表ワークシート
'* @param lStartRow data表データ開始行番号
'*
'******************************************************************************
Public Sub DeleteNoUsedRange(ByVal dataSheet As Worksheet, lStartRow As Long)
    Dim delStartRow As Long
    Dim delEndRow As Long
    
    Dim rng As Range
    Set rng = GetActualUsedRangeFromSheet(dataSheet, lStartRow, 1)
    If rng Is Nothing Then
        delStartRow = lStartRow
    Else
        delStartRow = rng.Item(rng.count).row + 1
    End If
    delEndRow = dataSheet.UsedRange.Item(dataSheet.UsedRange.count).row
    
    If delStartRow > delEndRow Then
        Exit Sub
    End If
    With dataSheet
        .Range(.Rows(delStartRow), .Rows(delEndRow)).Delete
    End With
End Sub

'******************************************************************************
'* [概  要] 使用セル範囲取得処理。
'* [詳  細] worksheetのデータ表の使用セル範囲を取得する｡
'*
'* @param dataSheet data表ワークシート
'* @param lStartRow data表データ開始行番号
'* @param lStartCol data表データ開始列番号
'* @param itemCount 項目列数
'* @param ignoreColnum 走査対象外の列番号
'* @return 使用セル範囲
'*
'******************************************************************************
Public Function GetActualUsedRangeFromSheet(dataSheet As Worksheet, lStartRow As Long, lStartCol As Long, Optional colCount As Long, Optional ignoreColnum As Long) As Range
    Dim lMaxRow As Long: lMaxRow = GetFinalRow(dataSheet, ignoreColnum)
    Dim lMaxCol As Long
    If colCount = 0 Then
        lMaxCol = GetFinalCol(dataSheet)
    Else
        lMaxCol = lStartCol + colCount - 1
    End If

    'レコードが存在しない場合
    If lMaxRow < lStartRow Or lMaxCol < lStartCol Then
        Set GetActualUsedRangeFromSheet = Nothing
        Exit Function
    End If
    
    Set GetActualUsedRangeFromSheet = dataSheet.Range(dataSheet.Cells(lStartRow, lStartCol), dataSheet.Cells(lMaxRow, lMaxCol))
End Function

'******************************************************************************
'* [概  要] SetHyperLink
'* [詳  細] ハイパーリンクを追加する（フォント・フォントサイズの変更をしない）
'*
'* @param sh シートオブジェクト
'* @param row 行番号
'* @param col 列番号
'* @param strLinkAddress リンクアドレス
'*
'******************************************************************************
Public Sub SetHyperLink(ByVal sh As Worksheet, row As Long, col As Long, strLinkAddress As String)
    Dim strFontName As String
    Dim iFontSize As Integer
    With sh
        strFontName = .Range(.Cells(row, col), .Cells(row, col)).Font.Name
        iFontSize = sh.Range(.Cells(row, col), .Cells(row, col)).Font.size

        .Hyperlinks.Add Anchor:=.Range(.Cells(row, col), .Cells(row, col)), _
                          Address:=strLinkAddress
                        
        .Range(.Cells(row, col), .Cells(row, col)).Font.Name = strFontName
        .Range(.Cells(row, col), .Cells(row, col)).Font.size = iFontSize
    End With
End Sub

'******************************************************************************
'* [概  要] シートコピー処理
'* [詳  細] フォーマットとなるシートをコピーし、対象のWorksheetオブジェクトを
'*          返却する。
'*          countAppend が True の場合、シート名+ nの名称（nは重複しない番号）
'*          を付与する。
'*
'* @param formatSheetName フォーマットとなるコピー元シート
'* @param destSheetName 作成するシート名
'* @param wb 対象のワークブック
'* @param countAppend 番号を付与するか
'* @return 作成したシートオブジェクト
'******************************************************************************
Public Function CopySheet(ByVal formatSheetName As String, ByVal destSheetName As String, _
                   Optional ByRef wb As Workbook, Optional ByVal countAppend As Boolean = False) As Worksheet
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim cnt As Long: cnt = 0
    If countAppend Then
        cnt = 1
        Do While SheetExists(destSheetName & cnt, wb)
            cnt = cnt + 1
        Loop
    End If
    
    Dim tmpVisible As Excel.XlSheetVisibility
    tmpVisible = wb.Worksheets(formatSheetName).Visible
    wb.Worksheets(formatSheetName).Visible = xlSheetVisible
    
    wb.Worksheets(formatSheetName).Copy After:=wb.Worksheets(formatSheetName)
    wb.ActiveSheet.Name = destSheetName & IIf(cnt > 0, CStr(cnt), "")
    Set CopySheet = wb.ActiveSheet
    
    wb.Worksheets(formatSheetName).Visible = tmpVisible
End Function

'******************************************************************************
'* [概  要] 行情報取得処理。
'* [詳  細] worksheetの表から1行分の情報を取得し、Variant配列を返却する｡
'*
'* @param dataSheet ワークシート
'* @param lStartRow データ開始行番号
'* @param lStartCol データ開始列番号
'* @param colCount 項目列数
'* @return Variant 1行分のデータ（1次元配列）
'*
'******************************************************************************
Public Function GetRowData(dataSheet As Worksheet, lStartRow As Long, Optional lStartCol As Long, Optional colCount As Long) As Variant
    Dim lMaxCol As Long
    If colCount = 0 Then
        lMaxCol = GetFinalKeyCol(dataSheet, lStartRow)
    Else
        lMaxCol = lStartCol + colCount - 1
    End If
    If lStartCol = 0 Then
        lStartCol = dataSheet.Cells(lStartRow, 1).End(xlToRight).Column
    End If
    
    'レコードが存在しない場合
    If lMaxCol < lStartCol Then
        GetRowData = Empty
        Exit Function
    End If
    
    Dim vArr: vArr = dataSheet.Range(dataSheet.Cells(lStartRow, lStartCol), dataSheet.Cells(lStartRow, lMaxCol))
    
    GetRowData = ArrayUtils.Copy2DArrayRow(vArr, LBound(vArr))
End Function

'******************************************************************************
'* [概  要] SetCenterAcrossSelection
'* [詳  細] 選択中の範囲に「選択範囲内で中央揃え」を設定する。
'*
'* @param targetRng 対象の範囲（省略すると選択中の範囲で処理）
'* @param resetCas  設定を解除するか（True:解除）
'*
'******************************************************************************
Public Sub SetCenterAcrossSelection(Optional ByRef targetRng As Range, Optional ByVal resetCas As Boolean = False)
    If targetRng Is Nothing Then
        If TypeName(Application.Selection) = "Range" Then
            Set targetRng = Application.Selection
        Else
            Exit Sub
        End If
    End If
    
    With targetRng
        If resetCas Then
            .HorizontalAlignment = xlGeneral
        Else
            .HorizontalAlignment = xlCenterAcrossSelection
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] SetValidateList
'* [詳  細] 選択中の範囲に入力規則のリストを設定する。
'*
'* @param vlist 入力規則のリスト（1次元配列）
'* @param targetRng 対象の範囲（省略すると選択中の範囲で処理）
'* @param resetVl  設定を解除するか（True:解除）
'*
'******************************************************************************
Public Sub SetValidateList(Optional vlist, Optional ByRef targetRng As Range, Optional ByVal resetVl As Boolean = False)
    If targetRng Is Nothing Then
        If TypeName(Application.Selection) = "Range" Then
            Set targetRng = Application.Selection
        Else
            Exit Sub
        End If
    End If

    With targetRng.Validation
        If resetVl Then
            .Delete
        Else
            If IsMissing(vlist) Then Exit Sub
            If ArrayUtils.IsEmptyArray(vlist) Then Exit Sub
            If Not ArrayUtils.Is1DArray(vlist) Then Exit Sub

            .Delete
            .Add Type:=xlValidateList, _
                 Formula1:=Join(vlist, ",")
        End If
    End With
   
End Sub

'******************************************************************************
'* [概  要] ChangeDisplayGridlines
'* [詳  細] 罫線の表示・非表示を切り替える。
'*
'* @param sheetName 対象のシート名（省略すると選択中のシートで処理）
'* @param wb ワークブックオブジェクト
'******************************************************************************
Public Sub ChangeDisplayGridlines(Optional ByVal sheetName As String, Optional ByRef wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim wd As Object
    If sheetName = "" Then
        Set wd = wb.Windows(1)
    Else
        Set wd = wb.Windows(1).SheetViews.Item(sheetName)
    End If
    
    With wd
        If .DisplayGridlines Then
            '罫線を非表示
            .DisplayGridlines = False
        Else
            '罫線を表示
            .DisplayGridlines = True
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] ChangeDisplayHeadings
'* [詳  細] 行列番号の表示・非表示を切り替える。
'*
'* @param sheetName 対象のシート名（省略すると選択中のシートで処理）
'* @param wb ワークブックオブジェクト
'******************************************************************************
Public Sub ChangeDisplayHeadings(Optional ByVal sheetName As String, Optional ByRef wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim wd As Object
    If sheetName = "" Then
        Set wd = wb.Windows(1)
    Else
        Set wd = wb.Windows(1).SheetViews.Item(sheetName)
    End If
    
    With wd
        If .DisplayHeadings Then
            '行列番号を非表示
            .DisplayHeadings = False
        Else
            '行列番号を表示
            .DisplayHeadings = True
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] ChangeDisplayFormulas
'* [詳  細] 数式の表示・非表示を切り替える。
'*
'* @param sheetName 対象のシート名（省略すると選択中のシートで処理）
'* @param wb ワークブックオブジェクト
'******************************************************************************
Public Sub ChangeDisplayFormulas(Optional ByVal sheetName As String, Optional ByRef wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim wd As Object
    If sheetName = "" Then
        Set wd = wb.Windows(1)
    Else
        Set wd = wb.Windows(1).SheetViews.Item(sheetName)
    End If
    
    With wd
        If .DisplayFormulas Then
            .DisplayFormulas = False
        Else
            .DisplayFormulas = True
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] ChangeDisplayOutline
'* [詳  細] アウトライン記号の表示・非表示を切り替える。
'*
'* @param sheetName 対象のシート名（省略すると選択中のシートで処理）
'* @param wb ワークブックオブジェクト
'******************************************************************************
Public Sub ChangeDisplayOutline(Optional ByVal sheetName As String, Optional ByRef wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim wd As Object
    If sheetName = "" Then
        Set wd = wb.Windows(1)
    Else
        Set wd = wb.Windows(1).SheetViews.Item(sheetName)
    End If
    
    With wd
        If .DisplayOutline Then
            .DisplayOutline = False
        Else
            .DisplayOutline = True
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] ChangeDisplayZeros
'* [詳  細] ゼロ値の表示・非表示を切り替える。
'*
'* @param sheetName 対象のシート名（省略すると選択中のシートで処理）
'* @param wb ワークブックオブジェクト
'******************************************************************************
Public Sub ChangeDisplayZeros(Optional ByVal sheetName As String, Optional ByRef wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim wd As Object
    If sheetName = "" Then
        Set wd = wb.Windows(1)
    Else
        Set wd = wb.Windows(1).SheetViews.Item(sheetName)
    End If
    
    With wd
        If .DisplayZeros Then
            .DisplayZeros = False
        Else
            .DisplayZeros = True
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] ChangeZoom
'* [詳  細] 表示倍率を変更する。
'*
'* @param zoomVal   表示倍率
'* @param sheetName 対象のシート名（省略すると選択中のシートで処理）
'* @param wb ワークブックオブジェクト
'******************************************************************************
Public Sub ChangeZoom(ByVal zoomVal As Long, Optional ByVal sheetName As String, Optional ByRef wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim wd As Object: Set wd = wb.Windows(1)
    Dim ss As Object: Set ss = wd.SelectedSheets
    
    If sheetName <> "" Then
        If Not SheetExists(sheetName, wb) Then Exit Sub
        wb.Sheets(sheetName).Select
    End If
    
    wd.Zoom = zoomVal
    ss.Select
End Sub

'******************************************************************************
'* [概  要] MergeRangeSilently
'* [詳  細] セルを結合する。結合時の警告は無視する。
'*
'* @param rngAddress 結合するセル範囲のアドレス
'* @param sheetName 対象のシート名（省略すると選択中のシートで処理）
'* @param wb ワークブックオブジェクト
'******************************************************************************
Public Sub MergeRangeSilently(ByVal rngAddress As String, Optional ByVal sheetName As String, Optional ByRef wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim ws As Worksheet
    If sheetName <> "" Then
        If Not SheetExists(sheetName, wb) Then Exit Sub
        Set ws = wb.Sheets(sheetName)
    Else
        Set ws = wb.ActiveSheet
    End If
    
    Dim dAlerts As Boolean: dAlerts = wb.Application.DisplayAlerts
    wb.Application.DisplayAlerts = False
    ws.Range(rngAddress).Merge
    wb.Application.DisplayAlerts = dAlerts
End Sub

'******************************************************************************
'* [概  要] SaveFileOLEObject
'* [詳  細] ファイルを指定したシートにOLEObjectとして保存する。
'*
'* @param filePath  対象となるファイルのフルパス
'* @param rowIdx    挿入先セル行インデックス
'* @param colIdx    挿入先セル列インデックス
'* @param sheetName 挿入先シート名
'* @param wb        挿入先ワークブックオブジェクト
'******************************************************************************
Public Function SaveFileOLEObject(ByVal filePath As String, _
                             ByVal rowIdx As Long, _
                             ByVal colIdx As Long, _
                             Optional ByVal sheetName As String, _
                             Optional ByRef wb As Workbook) As OLEObject
    
    If Not IO.FileExists(filePath) Then
        Err.Raise 9999, "XlWorkSheetUtils#SaveFileOLEObject", "File not found. [" & filePath & "]"
    End If
    
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim ws As Worksheet
    If sheetName <> "" Then
        If Not SheetExists(sheetName, wb) Then Exit Function
        Set ws = wb.Sheets(sheetName)
    Else
        Set ws = wb.ActiveSheet
    End If
    
    ws.Activate
    ws.Cells(rowIdx, colIdx).Select
    Set SaveFileOLEObject = ws.OLEObjects.Add(fileName:=filePath, DisplayAsIcon:=True, IconFileName:=IO.GetFileName(filePath), IconLabel:=IO.GetFileName(filePath))
    SaveFileOLEObject.Name = IO.GetFileName(filePath)
End Function

'******************************************************************************
'* [概  要] LoadFileOLEObject
'* [詳  細] ブック内のOLEObjectをファイルに保存する。
'*
'* @param folderPath 保存先フォルダのフルパス
'* @param oleName    OLEObject名
'* @param sheetName 対象のシート名（省略すると選択中のシートで処理）
'* @param wb         ワークブックオブジェクト
'* @param oleObj     OLEObject（指定するとこのOLEObjectを保存）
'******************************************************************************
Public Sub LoadFileOLEObject(ByVal folderPath As String, _
                             ByVal oleName As String, _
                             Optional ByVal sheetName As String, _
                             Optional ByRef wb As Workbook, _
                             Optional ByVal oleObj As OLEObject)
        
    If Not oleObj Is Nothing Then
            oleObj.Select
            oleObj.Copy
            Call SystemUtils.PasteInFolder(folderPath)
            Exit Sub
    End If
    
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim ws As Worksheet
    If sheetName <> "" Then
        If Not SheetExists(sheetName, wb) Then Exit Sub
        Set ws = wb.Sheets(sheetName)
    Else
        Set ws = wb.ActiveSheet
    End If

    
    For Each oleObj In ws.OLEObjects
        If oleObj.Name = oleName Then
            oleObj.Select
            oleObj.Copy
            Call SystemUtils.PasteInFolder(folderPath)
            Exit Sub
        End If
    Next
End Sub

'******************************************************************************
'* [概  要] InsertPicture
'* [詳  細] シートに画像を挿入する。
'*          挿入位置（上端/左端）、幅、高さは数値以外に、Rangeオブジェクトや
'*          A1形式セルアドレス文字列を指定可能（指定したセルの位置、幅、高さ
'*          で設定を行う）。
'*
'* @param imgFilePath      画像ファイルのフルパス
'* @param topPos           挿入位置（上端）
'* @param leftPos          挿入位置（左端）
'* @param widthVal         画像幅
'* @param heightVal        画像高さ
'* @param pLockAspectRatio 縦横比を固定にするか
'* @param pPlacement       セルにあわせて移動やサイズ変更するか
'* @param sheetName        対象のシート名（省略すると選択中のシートで処理）
'* @param wb               ワークブックオブジェクト
'* @return Shape 挿入した画像オブジェクト
'******************************************************************************
Public Function InsertPicture(ByVal imgFilePath As String, _
                              ByRef topPos As Variant, _
                              ByRef leftPos As Variant, _
                              ByRef widthVal As Variant, _
                              ByRef heightVal As Variant, _
                              Optional ByVal pLockAspectRatio As MsoTriState = msoTrue, _
                              Optional ByVal pPlacement As XlPlacement = xlMoveAndSize, _
                              Optional ByVal sheetName As String, _
                              Optional ByRef wb As Workbook) As Shape
    If Not IO.FileExists(imgFilePath) Then
        Err.Raise 9999, "XlWorkSheetUtils#InsertPicture", "File not found. [" & imgFilePath & "]"
    End If
    
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim ws As Worksheet
    If sheetName <> "" Then
        If Not SheetExists(sheetName, wb) Then Exit Function
        Set ws = wb.Sheets(sheetName)
    Else
        Set ws = wb.ActiveSheet
    End If
    
    Dim sp As Shape
    Set sp = ws.Shapes.AddPicture(imgFilePath, msoFalse, msoTrue, 0, 0, 0, 0)
    With sp
        .ScaleHeight 1, msoTrue
        .ScaleWidth 1, msoTrue
        
        .LockAspectRatio = pLockAspectRatio
        .Placement = pPlacement
        
        If TypeName(topPos) = "Range" Then
            .Top = topPos.Top
        ElseIf ValidateUtils.ValidateCellAddress(topPos) Then
            .Top = Range(topPos).Top
        Else
            .Top = CSng(topPos)
        End If
        If TypeName(leftPos) = "Range" Then
            .Left = leftPos.Left
        ElseIf ValidateUtils.ValidateCellAddress(leftPos) Then
            .Left = Range(leftPos).Left
        Else
            .Left = CSng(leftPos)
        End If
        If TypeName(widthVal) = "Range" Then
            .Width = widthVal.Width
        ElseIf ValidateUtils.ValidateCellAddress(widthVal) Then
            .Width = Range(widthVal).Width
        Else
            .Width = CSng(widthVal)
        End If
        If TypeName(heightVal) = "Range" Then
            .Height = heightVal.Height
        ElseIf ValidateUtils.ValidateCellAddress(heightVal) Then
            .Height = Range(heightVal).Height
        Else
            .Height = CSng(heightVal)
        End If
    End With
    Set InsertPicture = sp
End Function

'******************************************************************************
'* [概  要] RemoveDuplicates
'* [詳  細] 対象の範囲で指定した列の重複をチェックし、重複行を削除する。
'*
'* @param targetRng 対象の範囲（Nothingを指定すると選択中の範囲で処理）
'* @param colnums() 重複をチェックする列（Excel列番号ではなく、範囲内の番号）
'*
'******************************************************************************
Public Sub RemoveDuplicates(targetRng As Range, ParamArray colnums() As Variant)
    If targetRng Is Nothing Then
        If TypeName(Application.Selection) = "Range" Then
            Set targetRng = Application.Selection
        Else
            Exit Sub
        End If
    End If
    Dim tmp: tmp = colnums
    Dim Arr: Arr = ArrayUtils.ToVArrayFlattened(tmp)
    targetRng.RemoveDuplicates Columns:=CVar(Arr), Header:=xlNo
End Sub

'******************************************************************************
'* [概  要] GetActiveCell
'* [詳  細] 対象のシート中でアクティブ（選択中）なセルを返却する。
'*
'* @param sheetName 対象のシート名（省略すると選択中のシートで処理）
'* @param wb ワークブックオブジェクト
'******************************************************************************
Public Function GetActiveCell(Optional ByVal sheetName As String, Optional ByRef wb As Workbook) As Range
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If

    Dim ws As Worksheet
    If sheetName <> "" Then
        If Not SheetExists(sheetName, wb) Then Exit Function
        Set ws = wb.Sheets(sheetName)
    Else
        Set ws = wb.ActiveSheet
    End If
    
    Dim orgActiveSheet As Worksheet: Set orgActiveSheet = wb.ActiveSheet
    
    With wb.Application
        Dim tmpScrUpd As Boolean: tmpScrUpd = .ScreenUpdating
        .ScreenUpdating = False
        ws.Select
        Set GetActiveCell = .ActiveCell
        orgActiveSheet.Select
        .ScreenUpdating = tmpScrUpd
    End With
End Function

'******************************************************************************
'* [概  要] GetRangeSelection
'* [詳  細] 対象のシート中で選択中のセル範囲を返却する。
'*
'* @param sheetName 対象のシート名（省略すると選択中のシートで処理）
'* @param wb ワークブックオブジェクト
'******************************************************************************
Public Function GetRangeSelection(Optional ByVal sheetName As String, Optional ByRef wb As Workbook) As Range
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If

    Dim ws As Worksheet
    If sheetName <> "" Then
        If Not SheetExists(sheetName, wb) Then Exit Function
        Set ws = wb.Sheets(sheetName)
    Else
        Set ws = wb.ActiveSheet
    End If

    Dim orgActiveSheet As Worksheet: Set orgActiveSheet = wb.ActiveSheet
    
    With wb.Application
        Dim tmpScrUpd As Boolean: tmpScrUpd = .ScreenUpdating
        .ScreenUpdating = False
        ws.Select
        Set GetRangeSelection = wb.Windows(1).RangeSelection
        orgActiveSheet.Select
        .ScreenUpdating = tmpScrUpd
    End With
End Function

'******************************************************************************
'* [概  要] Column1ToA
'* [詳  細] 列番号の整数を文字列に変換する（1→A）。
'* [参  考] https://www.excel-chunchun.com/entry/2019/02/08/020935#%E3%81%A7VBA%E3%81%A7%E6%9C%80%E9%80%9F%E3%81%AE%E9%96%A2%E6%95%B0%E3%81%AF
'*
'* @param colNo 列番号
'* @return String 列文字列
'******************************************************************************
Function Column1ToA(colNo As Long) As String
    If colNo <= 0 Then
        Column1ToA = ""
    ElseIf colNo <= 26 Then
        Column1ToA = Chr(65 + ((colNo - 1) Mod 26))
    ElseIf colNo <= 26 * 27 Then
        Column1ToA = Chr(65 + ((colNo - 27) \ 26) Mod 26) & Chr(65 + ((colNo - 1) Mod 26))
    ElseIf colNo <= 26 * 27 * 27 Then
        Column1ToA = Chr(65 + ((colNo - 27 * 27) / 26 \ 26)) & Chr(65 + ((colNo - 27) \ 26) Mod 26) & Chr(65 + ((colNo - 1) Mod 26))
    Else
        Column1ToA = ""
    End If
End Function

'******************************************************************************
'* [概  要] ColumnATo1
'* [詳  細] 列番号の文字列を整数に変換する（A→1）。
'* [参  考] https://www.excel-chunchun.com/entry/2019/02/08/020935#%E3%81%A7VBA%E3%81%A7%E6%9C%80%E9%80%9F%E3%81%AE%E9%96%A2%E6%95%B0%E3%81%AF
'*
'* @param colLabel 列文字列
'* @return Long 列番号
'******************************************************************************
Function ColumnATo1(ByVal colLabel As String) As Long
    ColumnATo1 = 0
    
    If Len(colLabel) = 1 Then
        ColumnATo1 = Asc(colLabel) Mod 32
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To Len(colLabel)
        ColumnATo1 = ColumnATo1 * 26 + (Asc(Mid(colLabel, i, 1)) Mod 32)
    Next
End Function

'******************************************************************************
'* [概  要] DeleteRowBlank
'* [詳  細] 対象の列範囲でセルが空白の行を削除する。
'*
'* @param colNo      列番号
'* @param startRowNo 開始行番号
'* @param endRowNo   終了行番号
'* @param sheetName  対象のシート名（省略すると選択中のシートで処理）
'* @param wb         ワークブックオブジェクト
'*
'******************************************************************************
Public Sub DeleteRowBlank(colNo As Long, startRowNo As Long, endRowNo As Long, Optional ByVal sheetName As String, Optional ByRef wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim ws As Worksheet
    If sheetName <> "" Then
        If Not SheetExists(sheetName, wb) Then Exit Sub
        Set ws = wb.Sheets(sheetName)
    Else
        Set ws = wb.ActiveSheet
    End If
    
    With ws.Range(ws.Cells(startRowNo, colNo), ws.Cells(endRowNo, colNo)).SpecialCells(xlCellTypeBlanks).EntireRow
        .Hidden = True
        .Delete
    End With
End Sub

'******************************************************************************
'* [概  要] CellToPng
'* [詳  細] セル範囲をPNG画像に変換して保存する。
'* [参  考] https://vbabeginner.net/save-the-cell-as-a-png-image/
'*
'* @param saveDirPath  作成するPNGファイルを保存するフォルダ
'* @param rngAddress   PNG画像に変換するセルのアドレス。単一セルまたは結合セルを指定する。
'* @param baseFileName 作成するPNGファイルのファイル名の先頭文字列。省略可。
'* @param sheetName    対象のシート名（省略すると選択中のシートで処理）
'* @param wb           ワークブックオブジェクト
'*
'******************************************************************************
Public Sub CellToPng(ByVal saveDirPath As String, ByVal rngAddress As String, Optional baseFileName As String = "PNG_", Optional ByVal sheetName As String, Optional ByRef wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim ws As Worksheet
    If sheetName <> "" Then
        If Not SheetExists(sheetName, wb) Then Exit Sub
        Set ws = wb.Sheets(sheetName)
    Else
        Set ws = wb.ActiveSheet
    End If
    Dim orgActiveSheet As Worksheet: Set orgActiveSheet = wb.ActiveSheet
    
    Dim rng As Range: Set rng = ws.Range(rngAddress)
    
    With wb.Application
        Dim tmpScrUpd As Boolean: tmpScrUpd = .ScreenUpdating
        .ScreenUpdating = False

        ' セルアドレスを取得
        Dim sAddress As String
        If (rng.count = 1) Then
            If (rng.MergeArea.count > 1) Then
                sAddress = rng.MergeArea.Address(False, False)
                If (Left(sAddress, Len(rng.Address(False, False))) <> rng.Address(False, False)) Then
                    Exit Sub
                End If
            Else
                sAddress = rng.Address(False, False)
            End If
        Else
            sAddress = rng.Address(False, False)
        End If

        ' 対象セル範囲を画像としてコピー
        ws.Range(sAddress).CopyPicture Appearance:=xlScreen, Format:=xlPicture
        
        ' 図として貼り付け（オートシェイプ化）
        ws.PasteSpecial Format:="図 (拡張メタファイル)", Link:=False, DisplayAsIcon:=False
        
        ' 図を移動
        Dim oShape: Set oShape = Selection
        oShape.Top = rng.Top
        oShape.Left = rng.Left
        
        ' HTMLファイル名を設定
        Dim sAddressStr As String: sAddressStr = Replace(sAddress, ":", "")
        saveDirPath = FileUtils.AddPathSeparator(saveDirPath)
        Dim sHtmPath As String: sHtmPath = saveDirPath & baseFileName & sAddressStr & ".htm"
        
        ' HTML形式で保存（ファイル名にセルアドレスを付与する）
        With wb.PublishObjects.Add(xlSourceRange, sHtmPath, rng.Parent.Name, rng.Address, xlHtmlStatic, sAddressStr)
            .Publish True
            .AutoRepublish = False
        End With
        
        ' 作成したオートシェイプを削除
        Selection.Delete
        
        ' 単一セル指定の場合（セル範囲の場合はPNGファイルが複数出力され、image_xxx.pngのどれが正しいのか不明のため選別しない）
        Dim sPngFileName As String
        If (rng.count = 1) Then
            sPngFileName = sAddressStr & "_image001.png"
            
            ' 既に作成済みの場合は再作成するため削除
            If (IO.FileExists(saveDirPath & sPngFileName) = True) Then
                Call IO.DeleteFile(saveDirPath & sPngFileName)
            End If
            
            ' PNGファイルを移動（PNG_A1.filesフォルダから上階層に移動）
            Call IO.MoveFile(saveDirPath & baseFileName & sAddressStr & ".files\" & sPngFileName, saveDirPath & sPngFileName)
            
            ' PNG_A1.filesフォルダを削除
            Kill saveDirPath & baseFileName & sAddressStr & ".files\*"
            RmDir saveDirPath & baseFileName & sAddressStr & ".files"
        End If
        
        ' .htmファイルを削除
        Kill sHtmPath
        
        orgActiveSheet.Select
        .ScreenUpdating = tmpScrUpd
    End With
End Sub
