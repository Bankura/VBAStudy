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
'* [概  要] シート貼り付け処理。
'* [詳  細] Variant配列データをシートに出力する。
'*
'* @param dataSheet ワークシート
'* @param vArray Variant配列データ
'* @param lStartRow データ開始行番号
'* @param lStartCol データ開始列番号
'*
'******************************************************************************
Public Sub InjectVariantArrayToCells(ByVal dataSheet As Worksheet, ByVal vArray, lStartRow As Long, lStartCol As Long)
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
Private Sub InjectNumbersToIndexCells(ByVal dataSheet As Worksheet, lStartRow As Long, lStartCol As Long, rownum As Long)
    If rownum < 1 Then
        Exit Sub
    End If
    With dataSheet
        .Cells(lStartRow, lStartCol) = 1
        If rownum > 1 Then
            .Cells(lStartRow, lStartCol).AutoFill _
              Destination:=.Range(.Cells(lStartRow, lStartCol), .Cells(lStartRow + rownum - 1, lStartCol)), Type:=xlLinearTrend
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
        For i = .Rows.Count To 1 Step -1
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
        For i = .Columns.Count To 1 Step -1
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
Public Function GetFinalKeyRow(ByVal dataSheet As Worksheet, StartCol As Long) As Long
    GetFinalKeyRow = dataSheet.Cells(Rows.Count, StartCol).End(xlUp).row
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
Public Function GetFinalKeyCol(ByVal dataSheet As Worksheet, StartRow As Long) As Long
    GetFinalKeyCol = dataSheet.Cells(StartRow, Columns.Count).End(xlToLeft).Column
End Function

'******************************************************************************
'* [概  要] 表情報取得処理。
'* [詳  細] worksheetの表から情報を取得し、Variant配列を返却します｡
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
'* [詳  細] worksheetのデータ表の使用セル範囲をクリアします｡
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
'* [詳  細] worksheetのデータ表の未使用範囲行を削除します（UsedRangeを縮小）｡
'*
'* @param dataSheet data表ワークシート
'* @param lStartRow data表データ開始行番号
'*
'******************************************************************************
Public Sub DeleteNoUsedRange(dataSheet As Worksheet, lStartRow As Long)
    Dim delStartRow As Long
    Dim delEndRow As Long
    
    Dim rng As Range
    Set rng = GetActualUsedRangeFromSheet(dataSheet, lStartRow, 1)
    If rng Is Nothing Then
        delStartRow = lStartRow
    Else
        delStartRow = rng.Item(rng.Count).row + 1
    End If
    delEndRow = dataSheet.UsedRange.Item(dataSheet.UsedRange.Count).row
    
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
Public Sub SetHyperLink(sh As Worksheet, row As Long, col As Long, strLinkAddress As String)
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
'* [詳  細] フォーマットとなるシートをコピーし、シート名+ nの名称
'*          （nは重複しない番号）をつけて、対象のWorksheetオブジェクトを返却。
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

Public Sub ClearBookAuthor(Optional ByRef wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    wb.BuiltinDocumentProperties("Author").Value = " "
    wb.BuiltinDocumentProperties("Last Author").Value = " "
End Sub

Public Sub RemoveBookInformationAndSave(Optional ByRef wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    With wb
        .RemoveDocumentInformation xlRDIAll
        .RemoveDocumentInformation xlRDIComments
        .RemoveDocumentInformation xlRDIDocumentProperties
        .RemoveDocumentInformation xlRDIRemovePersonalInformation
        .RemoveDocumentInformation xlRDIPrinterPath
        .RemovePersonalInformation = False
        Dim dp As DocumentProperty
        Dim tmpUserName As String
        tmpUserName = Application.UserName
        Application.UserName = "unknown"
         For Each dp In .BuiltinDocumentProperties
            If dp.Type = msoPropertyTypeString Then
                dp.Value = ""
            End If
        Next
        .Save
        Application.UserName = tmpUserName
    End With
End Sub

