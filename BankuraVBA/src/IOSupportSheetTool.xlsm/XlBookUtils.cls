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
Private Const XML_DEFAULT_NAMESPACE As String = "http://mynamespace/CustomProperties" '名前空間

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
'* [概  要] ApplyProc2Books メソッド
'* [詳  細] 指定フォルダから、Excelファイルを開き、指定した処理を行う。
'*
'* @param srcPath Excelファイルが格納されているフォルダ
'* @param proc2Apply 関数名、またはFuncオブジェクト、または
'*                   Exec（x As Workbook）メソッドを
'*                   持つオブジェクト。
'* @param recursive 子フォルダ内も含め再帰的に処理するか（省略可）
'*
'******************************************************************************
Public Sub ApplyProc2Books(ByVal srcPath As String, ByVal proc2Apply, Optional ByVal recursive As Boolean = True)
    Dim displayAlertsState As Boolean: displayAlertsState = Application.DisplayAlerts
    On Error GoTo ErrorHandler

    If Application.DisplayAlerts Then Application.DisplayAlerts = False
    
    Dim v
    For Each v In FileUtils.SearchFile(srcPath, "", "*.xls*", , recursive)
        Dim bookObj As Workbook
        Set bookObj = Workbooks.Open(fileName:=CStr(v))
        
        If ValidateUtils.IsFunc(proc2Apply) Then
            Call proc2Apply.Apply(bookObj)
        ElseIf ValidateUtils.IsString(proc2Apply) And proc2Apply <> "" Then
            Call Application.Run(proc2Apply, bookObj)
        ElseIf IsObject(proc2Apply) Then
            Call proc2Apply.Exec(bookObj)
        End If
        
        bookObj.Close
    Next
    Application.DisplayAlerts = displayAlertsState
    Exit Sub

ErrorHandler:
    Application.DisplayAlerts = displayAlertsState
    Err.Raise 9999, "XlBookUtils#ApplyProc2Books", Err.Number & ":" & Err.Discription
End Sub

'******************************************************************************
'* [概  要] SaveNewBook メソッド
'* [詳  細] 新規Excelブックを作成し、保存する。
'*          処理を指定すると、新規Excelブックに処理を適用後保存する。
'*
'* @param saveBookPath Excelブック保存先パス
'* @param bookTemplate XlWBATemplateで指定するブックテンプレート
'* @param proc2Apply 関数名、またはFuncオブジェクト、または
'*                   Exec（x As Workbook）メソッドを
'*                   持つオブジェクト。
'* @param willClose 保存後クローズするか
'* @return Workbook 保存済み新規Excelブック
'******************************************************************************
Public Function SaveNewBook(ByVal saveBookPath As String, _
                            Optional ByVal bookTemplate As XlWBATemplate = xlWBATWorksheet, _
                            Optional ByVal proc2Apply, _
                            Optional ByVal willClose As Boolean = False, _
                            Optional FileFormat, _
                            Optional Password, _
                            Optional WriteResPassword, _
                            Optional ReadOnlyRecommended) As Workbook
    Dim displayAlertsState As Boolean: displayAlertsState = Application.DisplayAlerts
    On Error GoTo ErrorHandler

    If Application.DisplayAlerts Then Application.DisplayAlerts = False
    
    Dim bookObj As Workbook
    Set bookObj = Workbooks.Add(bookTemplate)
  
    If ValidateUtils.IsFunc(proc2Apply) Then
        Call proc2Apply.Apply(bookObj)
    ElseIf ValidateUtils.IsString(proc2Apply) And proc2Apply <> "" Then
        Call Application.Run(proc2Apply, bookObj)
    ElseIf IsObject(proc2Apply) Then
        Call proc2Apply.Exec(bookObj)
    End If
    
    bookObj.SaveAs saveBookPath, FileFormat, Password, WriteResPassword, ReadOnlyRecommended
    If willClose Then
        bookObj.Close
    End If
    
    Application.DisplayAlerts = displayAlertsState
    Set SaveNewBook = bookObj
    Exit Function

ErrorHandler:
    Application.DisplayAlerts = displayAlertsState
    Err.Raise 9999, "XlBookUtils#SaveNewBook", Err.Number & ":" & Err.Discription
End Function

'******************************************************************************
'* [概  要] ChangeExcelFiles2007Format メソッド
'* [詳  細] 指定フォルダから、2003形式Excelファイルを開き、Office2007形式
'*          フォーマット（xlsxまたはxlsm）で保存する。
'*
'* @param srcPath Excelファイルが格納されているフォルダ
'* @param destPath 保存先フォルダ
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

'******************************************************************************
'* [概  要] DeleteNames
'* [詳  細] 名前定義を全削除する。
'*
'* @param wb 対象のワークブック
'* @param refErrorOnly 参照エラーのみを削除する
'* @param externalRefOnly 外部参照のみを削除する
'******************************************************************************
Public Sub DeleteNames(Optional ByRef wb As Workbook, _
                       Optional ByVal refErrorOnly As Boolean = False, _
                       Optional ByVal externalRefOnly As Boolean = False)

    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim n As Name
    For Each n In wb.Names
        n.Delete
        If refErrorOnly Or externalRefOnly Then
            If refErrorOnly And InStr(n.RefersTo, "#REF") > 0 Then
                n.Delete
            End If
            If externalRefOnly And InStr(n.RefersTo, ".xl") > 0 Then
                n.Delete
            End If
        Else
            n.Delete
        End If
    Next
End Sub

'******************************************************************************
'* [概  要] DeleteDataValidations
'* [詳  細] 入力規則を全削除する。
'*
'* @param wb 対象のワークブック
'* @param refErrorOnly 参照エラーのみを削除する
'* @param externalRefOnly 外部参照のみを削除する
'******************************************************************************
Public Sub DeleteDataValidations(Optional ByRef wb As Workbook, _
                       Optional ByVal refErrorOnly As Boolean = False, _
                       Optional ByVal externalRefOnly As Boolean = False)

    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim sh As Worksheet

    For Each sh In wb.Worksheets
        On Error Resume Next
        Dim rng As Range, rngChild As Range
        Set rng = sh.Cells.SpecialCells(xlCellTypeAllValidation)
        On Error GoTo 0

        If refErrorOnly Or externalRefOnly Then
            For Each rngChild In rng
                If rngChild.Validation.Type = xlValidateList Then
                    If refErrorOnly And InStr(rngChild.Validation.Formula1, "#REF") > 0 Then
                        rngChild.Validation.Delete
                    End If
                    If externalRefOnly And InStr(rngChild.Validation.Formula1, ".xl") > 0 Then
                        rngChild.Validation.Delete
                    End If
                End If
            Next
        Else
            sh.Cells.Validation.Delete
        End If
    Next
End Sub

'******************************************************************************
'* [概  要] DeleteFormatConditions
'* [詳  細] 条件付き書式を全削除する。
'*
'* @param wb 対象のワークブック
'* @param refErrorOnly 参照エラーのみを削除する
'* @param externalRefOnly 外部参照のみを削除する
'******************************************************************************
Public Sub DeleteFormatConditions(Optional ByRef wb As Workbook, _
                       Optional ByVal refErrorOnly As Boolean = False, _
                       Optional ByVal externalRefOnly As Boolean = False)

    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim sh As Worksheet

    For Each sh In wb.Worksheets
        On Error Resume Next
        Dim rng As Range, rngChild As Range
        Set rng = sh.Cells.SpecialCells(xlCellTypeAllFormatConditions)
        On Error GoTo 0

        If refErrorOnly Or externalRefOnly Then
            Dim fcds As FormatConditions
            Dim fcd As FormatCondition
            
            Set fcds = rng.FormatConditions
            
            Dim i As Long
            For i = fcds.Count To 1 Step -1
                If TypeName(fcds(i)) = "FormatCondition" Then
                    Set fcd = fcds(i)
                    If refErrorOnly And InStr(fcd.Formula1, "#REF") > 0 Then
                        fcd.Delete
                    End If
                    If externalRefOnly And InStr(fcd.Formula1, ".xl") > 0 Then
                        fcd.Delete
                    End If
                End If
            Next
        Else
            sh.Cells.FormatConditions.Delete
        End If
    Next
End Sub

'******************************************************************************
'* [概  要] DeleteShapeOnActions
'* [詳  細] Shapeオブジェクトのマクロを全削除する。
'*
'* @param wb 対象のワークブック
'* @param externalRefOnly 外部参照のみを削除する
'******************************************************************************
Public Sub DeleteShapeOnActions(Optional ByRef wb As Workbook, _
                       Optional ByVal externalRefOnly As Boolean = False)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim sh As Worksheet
    For Each sh In wb.Worksheets
        Dim sp As Shape
        If externalRefOnly Then
            For Each sp In sh.Shapes
                If InStr(sp.OnAction, ".xl") > 0 Then
                    sp.OnAction = ""
                End If
            Next
        Else
            For Each sp In sh.Shapes
                sp.OnAction = ""
            Next
        End If
    Next
End Sub

'******************************************************************************
'* [概  要] DeleteStyles
'* [詳  細] スタイルを全削除する。
'*
'* @param wb 対象のワークブック
'******************************************************************************
Public Sub DeleteStyles(Optional ByRef wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    Dim i As Long
    For i = wb.Styles.Count To 1 Step -1
        wb.Styles(i).Delete
    Next
End Sub

'******************************************************************************
'* [概  要] ChangeDisplayWorkbookTabs
'* [詳  細] シート見出しの表示・非表示を切り替える。
'*
'* @param wb ワークブックオブジェクト
'******************************************************************************
Public Sub ChangeDisplayWorkbookTabs(Optional ByRef wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim wd As Object
    Set wd = wb.Windows(1)
    
    With wd
        If .DisplayWorkbookTabs Then
            'シート見出しを非表示
            .DisplayWorkbookTabs = False
        Else
            'シート見出しを表示
            .DisplayWorkbookTabs = True
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] ChangeDisplayHorizontalScrollBar
'* [詳  細] 横スクロールバーの表示・非表示を切り替える。
'*
'* @param wb ワークブックオブジェクト
'******************************************************************************
Public Sub ChangeDisplayHorizontalScrollBar(Optional ByRef wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim wd As Object
    Set wd = wb.Windows(1)
    
    With wd
        If .DisplayHorizontalScrollBar Then
            .DisplayHorizontalScrollBar = False
        Else
            .DisplayHorizontalScrollBar = True
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] ChangeDisplayVerticalScrollBar
'* [詳  細] 縦スクロールバーの表示・非表示を切り替える。
'*
'* @param wb ワークブックオブジェクト
'******************************************************************************
Public Sub ChangeDisplayVerticalScrollBar(Optional ByRef wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim wd As Object
    Set wd = wb.Windows(1)
    
    With wd
        If .DisplayVerticalScrollBar Then
            .DisplayVerticalScrollBar = False
        Else
            .DisplayVerticalScrollBar = True
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] ChangeDisplayScrollBar
'* [詳  細] スクロールバーの表示・非表示を切り替える。
'*
'* @param wb ワークブックオブジェクト
'******************************************************************************
Public Sub ChangeDisplayScrollBar(Optional ByRef wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim wd As Object
    Set wd = wb.Windows(1)
    
    With wd
        If .DisplayHorizontalScrollBar Then
            'スクロールバーを非表示
            .DisplayHorizontalScrollBar = False
            .DisplayVerticalScrollBar = False
        Else
            'スクロールバーを表示
            .DisplayHorizontalScrollBar = True
            .DisplayVerticalScrollBar = True
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] ChangeReferenceStyle
'* [詳  細] Excel参照形式を切り替える。
'*
'* @param wb ワークブックオブジェクト
'******************************************************************************
Public Sub ChangeReferenceStyle(Optional ByRef wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
        
    With wb.Application
        If .ReferenceStyle = xlA1 Then
            .ReferenceStyle = xlR1C1
        Else
            .ReferenceStyle = xlA1
        End If
    End With
End Sub

'******************************************************************************
'* [概  要] ClearBookAuthor
'* [詳  細] Excelブックの作成者（Author）をクリアする。
'*          保存時に再設定されるため、本メソッドは有用ではないと思われる。
'*
'* @param wb ワークブックオブジェクト
'******************************************************************************
Public Sub ClearBookAuthor(Optional ByRef wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    wb.BuiltinDocumentProperties("Author").Value = " "
    wb.BuiltinDocumentProperties("Last Author").Value = " "
End Sub

'******************************************************************************
'* [概  要] RemoveBookInformationAndSave
'* [詳  細] Excelブックのドキュメント情報をクリアして保存する。
'*          作成者は未設定にできないため「unknown」を設定する。
'*
'* @param wb ワークブックオブジェクト
'******************************************************************************
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

'******************************************************************************
'* [概  要] SaveCustomXMLProperty
'* [詳  細] ExcelブックのカスタムXMLにデータを保存する。
'* [参  考] https://www.ka-net.org/blog/?p=11930
'*
'* @param propName  プロパティ名
'* @param propValue プロパティ値
'* @param wb        ワークブックオブジェクト
'* @param ns        名前空間
'******************************************************************************
Public Sub SaveCustomXMLProperty(ByVal propName As String, _
                                 ByVal propValue As String, _
                                 Optional ByRef wb As Workbook, _
                                 Optional ByVal ns As String = XML_DEFAULT_NAMESPACE)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If

    Dim parts As Office.CustomXMLParts
    Set parts = wb.CustomXMLParts.SelectByNamespace(ns)
    
    Dim part As Office.CustomXMLPart
    If parts.Count < 1 Then
        Set part = InitCustomXMLParts(wb, ns)
    Else
        Set part = parts.Item(1)
    End If

    Dim root As Office.CustomXMLNode
    Set root = part.DocumentElement
    
    Dim Target As Office.CustomXMLNode
    Set Target = root.SelectSingleNode("//CustomProperty[@id='" & propName & "']")
    
    If Target Is Nothing Then
        Dim d As Object: Set d = CreateObject("MSXML2.DOMDocument.6.0")
        Dim child As Object: Set child = d.createElement("CustomProperty")
        
        ' id属性
        Dim attrId As Object: Set attrId = d.createAttribute("id")
        attrId.NodeValue = propName
        child.Attributes.setNamedItem attrId

        ' value属性
        Dim attrVal As Object: Set attrVal = d.createAttribute("value")
        attrVal.NodeValue = propValue
        child.Attributes.setNamedItem attrVal
        
        d.appendChild child
        root.AppendChildSubtree d.XML
    Else
        Dim targetAttr As Office.CustomXMLNode
        For Each targetAttr In Target.Attributes
            If targetAttr.BaseName = "value" Then
                targetAttr.NodeValue = propValue
                Exit For
            End If
        Next
    End If
End Sub

'******************************************************************************
'* [概  要] InitCustomXMLParts
'* [詳  細] カスタムXMLを初期化する。
'* [参  考] https://www.ka-net.org/blog/?p=11930
'*
'* @param propName プロパティ名
'* @param wb       ワークブックオブジェクト
'* @param ns       名前空間
'* @return String  プロパティ値
'******************************************************************************
Private Function InitCustomXMLParts(Optional ByRef wb As Workbook, _
                                    Optional ByVal ns As String = XML_DEFAULT_NAMESPACE) As Office.CustomXMLPart
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim d As Object: Set d = CreateObject("MSXML2.DOMDocument.6.0")
    Dim root As Object: Set root = d.createElement("CustomProperties")
    root.setAttribute "xmlns", ns
    d.appendChild root
  
    Set InitCustomXMLParts = wb.CustomXMLParts.Add(d.XML)
End Function

'******************************************************************************
'* [概  要] LoadCustomXMLProperty
'* [詳  細] ExcelブックのカスタムXMLからプロパティを読込む。
'* [参  考] https://www.ka-net.org/blog/?p=11930
'*
'* @param propName プロパティ名
'* @param wb       ワークブックオブジェクト
'* @param ns       名前空間
'* @return String  プロパティ値
'******************************************************************************
Public Function LoadCustomXMLProperty(ByVal propName As String, _
                                      Optional ByRef wb As Workbook, _
                                      Optional ByVal ns As String = XML_DEFAULT_NAMESPACE) As String
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If

    Dim parts As Office.CustomXMLParts
    Set parts = wb.CustomXMLParts.SelectByNamespace(ns)
    
    If parts.Count > 0 Then
        Dim root As Office.CustomXMLNode
        Set root = parts.Item(1).DocumentElement
        
        Dim Target As Office.CustomXMLNode
        Set Target = root.SelectSingleNode("//CustomProperty[@id='" & propName & "']")
        
        If Not Target Is Nothing Then
            Dim targetAttr As Office.CustomXMLNode
            For Each targetAttr In Target.Attributes
                If targetAttr.BaseName = "value" Then
                    LoadCustomXMLProperty = targetAttr.NodeValue
                End If
            Next
        End If
    End If
End Function

'******************************************************************************
'* [概  要] DeleteCustomXMLProperty
'* [詳  細] ExcelブックのカスタムXMLからプロパティを削除する。
'* [参  考] https://www.ka-net.org/blog/?p=11930
'*
'* @param propName プロパティ名
'* @param wb       ワークブックオブジェクト
'* @param ns       名前空間
'* @return String  プロパティ値
'******************************************************************************
Public Sub DeleteCustomXMLProperty(ByVal propName As String, _
                                   Optional ByRef wb As Workbook, _
                                   Optional ByVal ns As String = XML_DEFAULT_NAMESPACE)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If

    Dim parts As Office.CustomXMLParts
    Set parts = wb.CustomXMLParts.SelectByNamespace(ns)
    
    If parts.Count > 0 Then
        Dim root As Office.CustomXMLNode
        Set root = parts.Item(1).DocumentElement
        
        Dim Target As Office.CustomXMLNode
        Set Target = root.SelectSingleNode("//CustomProperty[@id='" & propName & "']")
        
        If Not Target Is Nothing Then
            Target.Delete
        End If
    End If
End Sub

'******************************************************************************
'* [概  要] DeleteCustomXMLPart
'* [詳  細] カスタムXMLを削除する。
'* [参  考] https://www.ka-net.org/blog/?p=11930
'*
'* @param wb       ワークブックオブジェクト
'* @param ns       名前空間
'* @return String  プロパティ値
'******************************************************************************
Public Sub DeleteCustomXMLPart(Optional ByRef wb As Workbook, _
                               Optional ByVal ns As String = XML_DEFAULT_NAMESPACE)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If

    Dim parts As Office.CustomXMLParts
    Set parts = wb.CustomXMLParts.SelectByNamespace(ns)
    If parts.Count > 0 Then parts.Item(1).Delete
End Sub

'******************************************************************************
'* [概  要] SaveFileCustomXML
'* [詳  細] ファイルをBase64変換し、ブック内のカスタムXMLに保存する。
'* [参  考] https://www.ka-net.org/blog/?p=11930
'*
'* @param filePath 対象となるファイルのフルパス
'* @param propName プロパティ名
'* @param wb       ワークブックオブジェクト
'* @param ns       名前空間
'******************************************************************************
Public Sub SaveFileCustomXML(ByVal FilePath As String, _
                             Optional ByVal propName As String, _
                             Optional ByRef wb As Workbook, _
                             Optional ByVal ns As String = XML_DEFAULT_NAMESPACE)
    If propName = "" Then
        propName = IO.GetFileName(FilePath)
    End If
    Call SaveCustomXMLProperty(propName, FileUtils.ReadFileBase64(FilePath), wb, ns)
End Sub

'******************************************************************************
'* [概  要] LoadFileCustomXML
'* [詳  細] ブック内のカスタムXMLに保存しているBase64データをバイナリ変換し
'*          ファイルに保存する。
'* [参  考] https://www.ka-net.org/blog/?p=11930
'*
'* @param filePath   保存先ファイルのフルパス
'* @param propName   プロパティ名
'* @param wb         ワークブックオブジェクト
'* @param ns         名前空間
'* @param willAppend 追記有無（省略可。規定はFalse:追記なし）
'******************************************************************************
Public Sub LoadFileCustomXML(ByVal FilePath As String, _
                             Optional ByVal propName As String, _
                             Optional ByRef wb As Workbook, _
                             Optional ByVal ns As String = XML_DEFAULT_NAMESPACE, _
                             Optional willAppend As Boolean = False)
    If propName = "" Then
        propName = IO.GetFileName(FilePath)
    End If
    Dim base64Text As String: base64Text = LoadCustomXMLProperty(propName, wb, ns)
    
    Call FileUtils.WriteFileBase64(FilePath, base64Text, willAppend)
End Sub
