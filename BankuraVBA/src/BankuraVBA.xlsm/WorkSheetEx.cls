VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkSheetEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Worksheet拡張クラス
'* [詳  細] WorksheetクラスのWrapperとしての機能を提供する他、
'*          Excel.Worksheetクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2020-2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

#Const XLVERSION = 2013

'******************************************************************************
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mOrigin As Worksheet
Private mCriterias() As Collection

'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* オリジナルObject プロパティ
'*-----------------------------------------------------------------------------
Property Get Origin() As Worksheet
Attribute Origin.VB_UserMemId = 0
    Set Origin = mOrigin
End Property
Property Set Origin(arg As Worksheet)
    Set mOrigin = arg
End Property

'*-----------------------------------------------------------------------------
'* Application プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Microsoft Excel アプリケーションを表すApplicationオブジェクトを返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.application
'*-----------------------------------------------------------------------------
Property Get Application() As Application
    Set Application = mOrigin.Application
End Property

'*-----------------------------------------------------------------------------
'* AutoFilter プロパティ（読み取り専用）
'*
'* [補  足]
'* ・フィルターがオンの場合は、AutoFilter オブジェクトを返す。
'* ・フィルターがオフの設定の場合は、Nothing を返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.autofilter
'*-----------------------------------------------------------------------------
Property Get AutoFilter() As AutoFilter
    Set AutoFilter = mOrigin.AutoFilter
End Property

'*-----------------------------------------------------------------------------
'* AutoFilterMode プロパティ
'*
'* [補  足]
'* ・Trueの場合、オートフィルターのドロップダウン矢印がシートに表示される。
'* ・このプロパティをFalseに設定すると矢印は削除されるが、 Trueに設定することはできない。
'* ・矢印を表示する（Trueにする）には、AutoFilter のメソッドを使用して、リスト
'*   にフィルターを適用する。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.autofiltermode
'*-----------------------------------------------------------------------------
Property Get AutoFilterMode() As Boolean
    AutoFilterMode = mOrigin.AutoFilterMode
End Property
Property Let AutoFilterMode(arg As Boolean)
    mOrigin.AutoFilterMode = arg
End Property

'*-----------------------------------------------------------------------------
'* Cells プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ワークシートのすべてのセル (現在使用されているセルだけでなく) を表す Range
'*   オブジェクトを返す。
'* ・Cellsキーワードの直後に行と列のインデックスを指定してRange.Item プロパティ
'*   を呼び出す用途でよく使用される。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.cells
'*-----------------------------------------------------------------------------
Property Get Cells() As Range
    Set Cells = mOrigin.Cells
End Property

'*-----------------------------------------------------------------------------
'* CircularReference プロパティ（読み取り専用）
'*
'* [補  足]
'* ・シートの最初の循環参照が含まれる範囲を表す Range オブジェクトを返す。
'* ・ワークシートに循環参照がない場合は、 Nothingを返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.circularreference
'*-----------------------------------------------------------------------------
Property Get CircularReference() As Range
    Set CircularReference = mOrigin.CircularReference
End Property

'*-----------------------------------------------------------------------------
'* CodeName プロパティ（読み取り専用）
'*
'* [補  足]
'* ・オブジェクトのコード名を返す（Sheet1の場合"Sheet1"）。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.codename
'*-----------------------------------------------------------------------------
Property Get CodeName() As String
    CodeName = mOrigin.CodeName
End Property

'*-----------------------------------------------------------------------------
'* Columns プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ワークシートのすべての列を表す Range オブジェクトを返す。
'* ・Columnsキーワードの直後に列のインデックスを指定してRange.Item プロパティ
'*   を呼び出す用途でよく使用される。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.columns
'*-----------------------------------------------------------------------------
Property Get Columns() As Range
    Set Columns = mOrigin.Columns
End Property

'*-----------------------------------------------------------------------------
'* Comments プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ワークシートのすべてのコメントを表す Comments コレクションを返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.comments
'*-----------------------------------------------------------------------------
Property Get Comments() As Comments
    Set Comments = mOrigin.Comments
End Property

#If XLVERSION >= 2016 Then
'*-----------------------------------------------------------------------------
'* CommentsThreaded プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ワークシートの最上位のコメント (返信なし) を表す CommentsThreaded コレクションを返す。
'* ・レガシおよびモダンコメントが含まれている。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.commentsthreaded
'*-----------------------------------------------------------------------------
Property Get CommentsThreaded() As CommentsThreaded
    Set CommentsThreaded = mOrigin.CommentsThreaded
End Property
#End If

'*-----------------------------------------------------------------------------
'* ConsolidationFunction プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Excelの一般機能である[データ]→[統合]で使用されている関数コードを返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.consolidationfunction
'*-----------------------------------------------------------------------------
Property Get ConsolidationFunction() As XlConsolidationFunction
    ConsolidationFunction = mOrigin.ConsolidationFunction
End Property

'*-----------------------------------------------------------------------------
'* ConsolidationOptions プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Excelの一般機能である[データ]→[統合]で使用されているオプションの3つの要素
'*   の配列を返す。
'* ・要素の値はBoolean型でオプションにチェックが入っている場合はTrueとなる。
'*   ConsolidationOptions(1) 先頭行にラベルを使用
'*   ConsolidationOptions(2) 左の列のラベルを使用
'*   ConsolidationOptions(3) ソースデータへのリンクを作成
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.consolidationoptions
'*-----------------------------------------------------------------------------
Property Get ConsolidationOptions() As Variant
    ConsolidationOptions = mOrigin.ConsolidationOptions
End Property

'*-----------------------------------------------------------------------------
'* ConsolidationSources プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Excelの一般機能である[データ]→[統合]で使用されているソース範囲の名前を
'*   示す文字列型 (string) の値の配列を返す。
'* ・シートに統合がない場合はEmptyを返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.consolidationsources
'*-----------------------------------------------------------------------------
Property Get ConsolidationSources() As Variant
    ConsolidationSources = mOrigin.ConsolidationSources
End Property

'*-----------------------------------------------------------------------------
'* CustomProperties プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ワークシートに関連付けられている識別子情報を表す CustomProperties オブジェクトを返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.customproperties
'*-----------------------------------------------------------------------------
Property Get CustomProperties() As CustomProperties
    Set CustomProperties = mOrigin.CustomProperties
End Property

'*-----------------------------------------------------------------------------
'* Creator プロパティ（読み取り専用）
'*
'* [補  足]
'* ・オブジェクトを作成したアプリケーションを示す、32 ビットの整数値を返す。
'* ・Microsoft Excel で作成された場合、文字列 XCEL を返す(16 進数の 5843454C と同値)。
'* ・アプリケーションに 4 文字の作成元のコードが付加される Macintosh 版
'*   Microsoft Excel で使用するように設計されたもの。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.creator
'*-----------------------------------------------------------------------------
Property Get Creator() As XlCreator
    Creator = mOrigin.Creator
End Property

'*-----------------------------------------------------------------------------
'* DisplayPageBreaks プロパティ
'*
'* [補  足]
'* ・ワークシートの改ページ (自動と手動) の表示有無。Trueは表示有り。
'* ・プリンターがセットアップされていない場合、このプロパティを設定できない。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.displaypagebreaks
'*-----------------------------------------------------------------------------
Property Get DisplayPageBreaks() As Boolean
    DisplayPageBreaks = mOrigin.DisplayPageBreaks
End Property
Property Let DisplayPageBreaks(arg As Boolean)
    mOrigin.DisplayPageBreaks = arg
End Property

'*-----------------------------------------------------------------------------
'* DisplayRightToLeft プロパティ
'*
'* [補  足]
'* ・Trueの場合、ワークシートが左から右にではなく、右から左に表示される。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.displayrighttoleft
'*-----------------------------------------------------------------------------
Property Get DisplayRightToLeft() As Boolean
    DisplayRightToLeft = mOrigin.DisplayRightToLeft
End Property
Property Let DisplayRightToLeft(arg As Boolean)
    mOrigin.DisplayRightToLeft = arg
End Property

'*-----------------------------------------------------------------------------
'* EnableAutoFilter プロパティ
'*
'* [補  足]
'* ・保護されたワークシートのオートフィルター矢印を使用可能にするか指定する。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.enableautofilter
'*-----------------------------------------------------------------------------
Property Get EnableAutoFilter() As Boolean
    EnableAutoFilter = mOrigin.EnableAutoFilter
End Property
Property Let EnableAutoFilter(arg As Boolean)
    mOrigin.EnableAutoFilter = arg
End Property

'*-----------------------------------------------------------------------------
'* EnableCalculation プロパティ
'*
'* [補  足]
'* ・ワークシートを自動的に再計算するか指定する。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.enablecalculation
'*-----------------------------------------------------------------------------
Property Get EnableCalculation() As Boolean
    EnableCalculation = mOrigin.EnableCalculation
End Property
Property Let EnableCalculation(arg As Boolean)
    mOrigin.EnableCalculation = arg
End Property

'*-----------------------------------------------------------------------------
'* EnableFormatConditionsCalculation プロパティ
'*
'* [補  足]
'* ・条件付き書式の評価は自動的に行うか指定する。
'* ・このフラグの目的は、評価の前に VBA プログラマがルールを完全に構成できるようにすること。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.enableformatconditionscalculation
'*-----------------------------------------------------------------------------
Property Get EnableFormatConditionsCalculation() As Boolean
    EnableFormatConditionsCalculation = mOrigin.EnableFormatConditionsCalculation
End Property
Property Let EnableFormatConditionsCalculation(arg As Boolean)
    mOrigin.EnableFormatConditionsCalculation = arg
End Property

'*-----------------------------------------------------------------------------
'* EnableOutlining プロパティ
'*
'* [補  足]
'* ・保護されたワークシートのアウトライン記号を使用可能にするか指定する。
'* ・userInterfaceOnlyが有効になっている場合に有効。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.enableoutlining
'*-----------------------------------------------------------------------------
Property Get EnableOutlining() As Boolean
    EnableOutlining = mOrigin.EnableOutlining
End Property
Property Let EnableOutlining(arg As Boolean)
    mOrigin.EnableOutlining = arg
End Property

'*-----------------------------------------------------------------------------
'* EnablePivotTable プロパティ
'*
'* [補  足]
'* ・保護されたワークシートのピボットテーブルのコントロールを使用可能にするか指定する。
'* ・userInterfaceOnlyが有効になっている場合に有効。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.enablepivottable
'*-----------------------------------------------------------------------------
Property Get EnablePivotTable() As Boolean
    EnablePivotTable = mOrigin.EnablePivotTable
End Property
Property Let EnablePivotTable(arg As Boolean)
    mOrigin.EnablePivotTable = arg
End Property

'*-----------------------------------------------------------------------------
'* EnableSelection プロパティ
'*
'* [補  足]
'* ・シートで選択できるものを設定・取得する。
'* ・ワークシートを保護しているときにだけ、このプロパティを使用可能。
'* ・xlNoSelection シートのすべてのセルの選択を禁止
'*   xlUnlockedCells Locked プロパティが False のセルに限り選択を許可
'*   xlNoRestrictions どのセルの選択も許可
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.enableselection
'*-----------------------------------------------------------------------------
Property Get EnableSelection() As XlEnableSelection
    EnableSelection = mOrigin.EnableSelection
End Property
Property Let EnableSelection(arg As XlEnableSelection)
    mOrigin.EnableSelection = arg
End Property

'*-----------------------------------------------------------------------------
'* FilterMode プロパティ（読み取り専用）
'*
'* [補  足]
'* ・オートフィルタが設定されていて、なおかつ何らかの条件で絞り込まれている場合、
'*   True を返す。絞り込まれていないときはFalseを返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.filtermode
'* ・http://officetanaka.net/excel/vba/tips/tips129.htm
'*-----------------------------------------------------------------------------
Property Get FilterMode() As Boolean
    FilterMode = mOrigin.FilterMode
End Property

'*-----------------------------------------------------------------------------
'* HPageBreaks プロパティ（読み取り専用）
'*
'* [補  足]
'* ・シートの水平方向の改ページを表す HPageBreaks コレクションを取得。
'* ・1 つのシートに使用できる水平な改ページの数は 1026 。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.hpagebreaks
'*-----------------------------------------------------------------------------
Property Get HPageBreaks() As HPageBreaks
    Set HPageBreaks = mOrigin.HPageBreaks
End Property

'*-----------------------------------------------------------------------------
'* Hyperlinks プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ワークシートのハイパーリンクを表す hyperlinks コレクションを取得。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.hyperlinks
'*-----------------------------------------------------------------------------
Property Get Hyperlinks() As Hyperlinks
    Set Hyperlinks = mOrigin.Hyperlinks
End Property

'*-----------------------------------------------------------------------------
'* Index プロパティ（読み取り専用）
'*
'* [補  足]
'* ・類似するオブジェクトのコレクション内でのオブジェクトのインデックス番号を
'*   表す長整数型 (Long) の値を返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.index
'*-----------------------------------------------------------------------------
Property Get Index() As Long
    Index = mOrigin.Index
End Property

'*-----------------------------------------------------------------------------
'* ListObjects プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ワークシートの ListObject オブジェクトのコレクションを返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.listobjects
'*-----------------------------------------------------------------------------
Property Get ListObjects() As ListObjects
    Set ListObjects = mOrigin.ListObjects
End Property

'*-----------------------------------------------------------------------------
'* MailEnvelope プロパティ（読み取り専用）
'*
'* [補  足]
'* ・文書の電子メールヘッダーを表す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.mailenvelope
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/office.msoenvelope
'*-----------------------------------------------------------------------------
Property Get MailEnvelope() As MsoEnvelope
    Set MailEnvelope = mOrigin.MailEnvelope
End Property

'*-----------------------------------------------------------------------------
'* Name プロパティ
'*
'* [補  足]
'* ・ワークシートの名前。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.name
'*-----------------------------------------------------------------------------
Property Get Name() As String
    Name = mOrigin.Name
End Property
Property Let Name(arg As String)
    mOrigin.Name = arg
End Property

#If XLVERSION >= 2016 Then
'*-----------------------------------------------------------------------------
'* NamedSheetViews プロパティ（読み取り専用）
'*
'* [補  足]
'* ・他のユーザーと共同作業中に他のユーザーによって中断されることなく、カスタマイズ
'*   されたビューを作成できるシートビューを表す NamedSheetViewCollectionコレクションを返す。。
'* ・https://support.microsoft.com/ja-jp/office/excel-%E3%81%AE%E3%82%B7%E3%83%BC%E3%83%88%E3%83%93%E3%83%A5%E3%83%BC-0eea3dc5-d7d1-44c5-a953-25ebfbd6c1a6#ID0EBAEAAA=%E3%83%87%E3%82%B9%E3%82%AF%E3%83%88%E3%83%83%E3%83%97
'*-----------------------------------------------------------------------------
Property Get NamedSheetViews() As NamedSheetViewCollection
    Set NamedSheetViews = mOrigin.NamedSheetViews
End Property
#End If
    
'*-----------------------------------------------------------------------------
'* Names プロパティ（読み取り専用）
'*
'* [補  足]
'* ・すべてのワークシート固有の名前 ("ワークシート名!" で定義された名前) を表す Names コレクションを返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.names
'*-----------------------------------------------------------------------------
Property Get Names() As Names
    Set Names = mOrigin.Names
End Property

'*-----------------------------------------------------------------------------
'* Next_ プロパティ（読み取り専用）
'*
'* [補  足]
'* ・次のシートを表す Worksheet オブジェクトを取得。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.next
'*-----------------------------------------------------------------------------
Property Get Next_() As Object
    Set Next_ = mOrigin.Next
End Property

'*-----------------------------------------------------------------------------
'* Outline プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ワークシートのアウトラインを表す Outline オブジェクトを取得。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.outline
'*-----------------------------------------------------------------------------
Property Get Outline() As Outline
    Set Outline = mOrigin.Outline
End Property

'*-----------------------------------------------------------------------------
'* PageSetup プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ワークシートのすべてのページ設定を含む PageSetup オブジェクトを取得。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.pagesetup
'*-----------------------------------------------------------------------------
Property Get PageSetup() As PageSetup
    Set PageSetup = mOrigin.PageSetup
End Property

'*-----------------------------------------------------------------------------
'* Parent プロパティ（読み取り専用）
'*
'* [補  足]
'* ・指定されたオブジェクトの親オブジェクトを返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.parent
'*-----------------------------------------------------------------------------
Property Get Parent() As Object
    Set Parent = mOrigin.Parent
End Property

'*-----------------------------------------------------------------------------
'* Previous プロパティ（読み取り専用）
'*
'* [補  足]
'* ・前のシートを表すWorksheetオブジェクトを返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.previous
'*-----------------------------------------------------------------------------
Property Get Previous() As Object
    Set Previous = mOrigin.Previous
End Property

'*-----------------------------------------------------------------------------
'* PrintedCommentPages プロパティ（読み取り専用）
'*
'* [補  足]
'* ・現在のワークシートに対して印刷されるコメント ページの数を返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.printedcommentpages
'*-----------------------------------------------------------------------------
Property Get PrintedCommentPages() As Long
    PrintedCommentPages = mOrigin.PrintedCommentPages
End Property

'*-----------------------------------------------------------------------------
'* ProtectContents プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Trueの場合は、シートの内容が保護されている。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.protectcontents
'*-----------------------------------------------------------------------------
Property Get ProtectContents() As Boolean
    ProtectContents = mOrigin.ProtectContents
End Property

'*-----------------------------------------------------------------------------
'* ProtectDrawingObjects プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Trueの場合は図形は保護されている。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.protectdrawingobjects
'*-----------------------------------------------------------------------------
Property Get ProtectDrawingObjects() As Boolean
    ProtectDrawingObjects = mOrigin.ProtectDrawingObjects
End Property

'*-----------------------------------------------------------------------------
'* Protection プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ワークシートの保護オプションを表す Protection オブジェクトを返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.protection
'*-----------------------------------------------------------------------------
Property Get Protection() As Protection
    Set Protection = mOrigin.Protection
End Property

'*-----------------------------------------------------------------------------
'* ProtectionMode プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Trueの場合は、ユーザーインターフェイスのみの保護が有効。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.protectionmode
'*-----------------------------------------------------------------------------
Property Get ProtectionMode() As Boolean
    ProtectionMode = mOrigin.ProtectionMode
End Property

'*-----------------------------------------------------------------------------
'* ProtectScenarios プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ワークシートのシナリオが保護されている場合は True を指定します。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.protectscenarios
'*-----------------------------------------------------------------------------
Property Get ProtectScenarios() As Boolean
    ProtectScenarios = mOrigin.ProtectScenarios
End Property

'*-----------------------------------------------------------------------------
'* QueryTables プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ワークシートのすべてのクエリテーブルを表す QueryTables コレクションを取得。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.querytables
'*-----------------------------------------------------------------------------
Property Get QueryTables() As QueryTables
    Set QueryTables = mOrigin.QueryTables
End Property

'*-----------------------------------------------------------------------------
'* Range プロパティ（読み取り専用）
'*
'* [補  足]
'* ・セルまたはセル範囲を表す Range オブジェクトを返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.range
'*-----------------------------------------------------------------------------
Property Get Range(Cell1, Optional Cell2) As Range
    Set Range = mOrigin.Range(Cell1, Cell2)
End Property

'*-----------------------------------------------------------------------------
'* Rows プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ワークシートのすべての行を表す Range オブジェクトを返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.rows
'*-----------------------------------------------------------------------------
Property Get Rows() As Range
    Set Rows = mOrigin.Rows
End Property

'*-----------------------------------------------------------------------------
'* ScrollArea プロパティ
'*
'* [補  足]
'* ・A1 スタイルの範囲参照として、スクロールを許可する範囲。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.scrollarea
'*-----------------------------------------------------------------------------
Property Get ScrollArea() As String
    ScrollArea = mOrigin.ScrollArea
End Property
Property Let ScrollArea(arg As String)
    mOrigin.ScrollArea = arg
End Property

'*-----------------------------------------------------------------------------
'* Shapes プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ワークシート上のすべての図形を表す Shapes コレクションを返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.shapes
'*-----------------------------------------------------------------------------
Property Get Shapes() As Shapes
    Set Shapes = mOrigin.Shapes
End Property

'*-----------------------------------------------------------------------------
'* Sort プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Sort オブジェクトを取得。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.sort
'*-----------------------------------------------------------------------------
Property Get Sort() As Sort
    Set Sort = mOrigin.Sort
End Property

'*-----------------------------------------------------------------------------
'* StandardHeight プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ワークシートのすべての行の標準 (既定) の高さをポイント単位で返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.standardheight
'*-----------------------------------------------------------------------------
Property Get StandardHeight() As Double
    StandardHeight = mOrigin.StandardHeight
End Property

'*-----------------------------------------------------------------------------
'* StandardWidth プロパティ
'*
'* [補  足]
'* ・ワークシートのすべての列の標準 (既定) 幅。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.standardwidth
'*-----------------------------------------------------------------------------
Property Get StandardWidth() As Double
    StandardWidth = mOrigin.StandardWidth
End Property
Property Let StandardWidth(arg As Double)
    mOrigin.StandardWidth = arg
End Property

'*-----------------------------------------------------------------------------
'* Tab_ プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ワークシートの Tab オブジェクトを返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.tab
'*-----------------------------------------------------------------------------
Property Get Tab_() As Excel.Tab
    Tab_ = mOrigin.Tab
End Property

'*-----------------------------------------------------------------------------
'* TransitionExpEval プロパティ
'*
'* [補  足]
'* ・ワークシートに Lotus 1-2-3 式の評価ルールを使用する場合は True 。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.transitionexpeval
'*-----------------------------------------------------------------------------
Property Get TransitionExpEval() As Boolean
    TransitionExpEval = mOrigin.TransitionExpEval
End Property
Property Let TransitionExpEval(arg As Boolean)
    mOrigin.TransitionExpEval = arg
End Property

'*-----------------------------------------------------------------------------
'* TransitionFormEntry プロパティ
'*
'* [補  足]
'* ・Trueの場合、ワークシートの Lotus 1-2-3 数式入力規則を使用。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.transitionformentry
'*-----------------------------------------------------------------------------
Property Get TransitionFormEntry() As Boolean
    TransitionFormEntry = mOrigin.TransitionFormEntry
End Property
Property Let TransitionFormEntry(arg As Boolean)
    mOrigin.TransitionFormEntry = arg
End Property

'*-----------------------------------------------------------------------------
'* Type_ プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ワークシートの種類を表す Xlsheettype の値を返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.type
'*-----------------------------------------------------------------------------
Property Get Type_() As XlSheetType
    Type_ = mOrigin.Type
End Property

'*-----------------------------------------------------------------------------
'* UsedRange プロパティ（読み取り専用）
'*
'* [補  足]
'* ・指定したワークシートの使用されている範囲を表す Range オブジェクトを返す。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.usedrange
'*-----------------------------------------------------------------------------
Property Get UsedRange() As Range
    Set UsedRange = mOrigin.UsedRange
End Property

'*-----------------------------------------------------------------------------
'* Visible プロパティ
'*
'* [補  足]
'* ・オブジェクトを表示するかどうかを決定する Xlsheetvisibility の値。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.visible
'*-----------------------------------------------------------------------------
Property Get Visible() As XlSheetVisibility
    Visible = mOrigin.Visible
End Property
Property Let Visible(arg As XlSheetVisibility)
    mOrigin.Visible = arg
End Property

'*-----------------------------------------------------------------------------
'* VPageBreaks プロパティ（読み取り専用）
'*
'* [補  足]
'* ・シートの垂直方向の改ページを表す VPageBreaks コレクションを取得。
'* ・https://docs.microsoft.com/ja-jp/office/vba/api/excel.worksheet.vpagebreaks
'*-----------------------------------------------------------------------------
Property Get VPageBreaks() As VPageBreaks
    Set VPageBreaks = mOrigin.VPageBreaks
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************
'******************************************************************************
'* [概  要] Init メソッド
'* [詳  細] 初期化用メソッド｡
'*
'* @param sheetName シート名
'* @param wb ワークブック
'* @return WorkSheetExオブジェクト
'*
'******************************************************************************
Public Function Init(ByVal sheetName As String, Optional ByVal wb As Workbook) As WorkSheetEx
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    Set mOrigin = wb.Worksheets(sheetName)
    Set Init = Me
End Function

'******************************************************************************
'* [概  要] Activate メソッド
'* [詳  細] 現在のシートを作業中のシートにします。
'*
'******************************************************************************
Sub Activate()
    Call mOrigin.Activate
End Sub

'******************************************************************************
'* [概  要] Calculate メソッド
'* [詳  細] 計算を実行します。計算の対象となるのは、ブック内の特定のワークシートです。
'*
'******************************************************************************
Sub Calculate()
    Call mOrigin.Calculate
End Sub

'******************************************************************************
'* [概  要] ChartObjects メソッド
'* [詳  細] 単一の埋め込みグラフ ( ChartObject オブジェクト)、またはシート上のすべての埋め込みグラフ ( ChartObjects オブジェクト) のコレクションを表すオブジェクトを返します。
'*
'* @param Index グラフの名前またはインデックス番号を指定します。 この引数には、複数のグラフを指定する配列を指定できます。。省略可能。
'* @return Object ChartObject / ChartObjects オブジェクト
'*
'******************************************************************************
Function ChartObjects(Optional Index) As Object
    Set ChartObjects = mOrigin.ChartObjects(Index)
End Function

'******************************************************************************
'* [概  要] CheckSpelling メソッド
'* [詳  細] オブジェクトのスペルをチェックする [スペルチェック] ダイアログ
'*           ボックスが表示される。
'*
'* @param CustomDictionary メイン辞書に単語が見つからない場合に調べるユーザー辞書のファイル名を示す文字列。省略可能。
'* @param IgnoreUppercase すべて大文字の単語を無視するには、 True を指定します。省略可能。
'* @param AlwaysSuggest Trueを設定すると、スペルミスが検出されたときに修正候補の一覧を表示することができます。省略可能。
'* @param SpellLang 辞書に使用した言語を指定します。省略可能。
'*
'******************************************************************************
Sub CheckSpelling(Optional CustomDictionary, Optional IgnoreUppercase, Optional AlwaysSuggest, Optional SpellLang)
    Call mOrigin.CheckSpelling(CustomDictionary, IgnoreUppercase, AlwaysSuggest, SpellLang)
End Sub

'******************************************************************************
'* [概  要] CircleInvalid メソッド
'* [詳  細] ワークシートの [入力規則] で設定された制限外の値が入力されているセルに丸印を付けます。
'*
'******************************************************************************
Sub CircleInvalid()
    Call mOrigin.CircleInvalid
End Sub

'******************************************************************************
'* [概  要] ClearArrows メソッド
'* [詳  細] ワークシートからトレース矢印を削除します。 なお、トレース矢印とは使用したワークシート分析機能によって追加される矢印です。
'*
'******************************************************************************
Sub ClearArrows()
    Call mOrigin.ClearArrows
End Sub

'******************************************************************************
'* [概  要] ClearCircles メソッド
'* [詳  細] ワークシートの [入力規則] で設定された制限外の値が入力されているセルに付けられた丸印を消去します。
'*
'******************************************************************************
Sub ClearCircles()
    Call mOrigin.ClearCircles
End Sub

'******************************************************************************
'* [概  要] Copy メソッド
'* [詳  細] シートを現在のブックか新しいブック内の他の場所にコピーします。
'*
'* @param Before コピーしたシートを特定のシートの直前の位置に挿入するときに、
'*              そのシートを指定｡省略可能｡
'*              引数 Before を指定すると、引数 After を指定することはできません。
'* @param After コピーしたシートを特定のシートの直後の位置に挿入するときに、
'*              そのシートを指定｡省略可能｡
'*              引数 After を指定すると、引数 Before を指定することはできません。
'*
'******************************************************************************
Sub Copy(Optional Before, Optional After)
    Call mOrigin.Copy(Before, After)
End Sub

'******************************************************************************
'* [概  要] Delete メソッド
'* [詳  細] シートを削除します。このメソッドは、削除を確認するダイアログボックスを表示します。ダイアログボックスを表示せずにワークシートを削除するには、 DisplayAlerts プロパティをFalseに設定します。
'*
'******************************************************************************
Sub Delete()
    Call mOrigin.Delete
End Sub

'******************************************************************************
'* [概  要] Evaluate メソッド
'* [詳  細] 使用する名前を、オブジェクトまたは値に変換します。Evaluate メソッドの代わりに角かっこを使っても、同じ結果が得られます ([A1:C5] など)。
'*
'* @param Name Excel の名前付け規則を使用して、オブジェクトの名前を指定します。
'* @return Variant オブジェクトまたは値
'*
'******************************************************************************
Function Evaluate(Name)
    Evaluate = mOrigin.Evaluate(Name)
End Function

#If XLVERSION >= 2016 Then
'******************************************************************************
'* [概  要] ExportAsFixedFormat メソッド
'* [詳  細] 指定された形式のファイルにエクスポートします。
'*
'* @param Type_ エクスポート後のファイル形式を指定します。
'* @param fileName 保存されるファイルの名前。省略可能。
'* @param Quality ファイルの品質を指定。省略可能。
'* @param IncludeDocProperties ドキュメントプロパティを含める場合は True。省略可能。
'* @param IgnorePrintAreas 発行時に設定を無視する場合はTrue 。省略可能。
'* @param From 開始するページのページ番号を指定。省略可能。
'* @param To_ 終了するページの番号を指定。省略可能。
'* @param OpenAfterPublish 発行後にビューアーにファイルを表示する場合は True。省略可能。
'* @param FixedFormatExtClassPtr FixedFormatExt クラスへのポインター。省略可能。
'* @param WorkIdentity xxxxx。省略可能。
'*
'******************************************************************************
Sub ExportAsFixedFormat(Type_ As XlFixedFormatType, Optional FileName, Optional Quality, Optional IncludeDocProperties, Optional IgnorePrintAreas, Optional From, Optional To_, Optional OpenAfterPublish, Optional FixedFormatExtClassPtr, Optional WorkIdentity)
    Call mOrigin.ExportAsFixedFormat(Type_, FileName, Quality, IncludeDocProperties, IgnorePrintAreas, From, To_, OpenAfterPublish, FixedFormatExtClassPtr, WorkIdentity)
End Sub
#End If

'******************************************************************************
'* [概  要] Move メソッド
'* [詳  細] シートをブック内の他の場所に移動します。
'*          引数 Before と引数 After を共に省略した場合は、新規ブックが自動的に作成され、シートはそのブック内に移動します。
'*
'* @param Before 移動したシートを特定のシートの直前の位置に挿入するときに、そのシートを指定します。省略可能。
'* @param After  移動したシートを特定のシートの直後の位置に挿入するときに、そのシートを指定します。省略可能。
'*
'******************************************************************************
Sub Move(Optional Before, Optional After)
    Call mOrigin.Move(Before, After)
End Sub
'******************************************************************************
'* [概  要] OLEObjects メソッド
'* [詳  細] 1つの OLE オブジェクト ( OLEObject)、または、グラフまたはシート上のすべての ole オブジェクト ( OLEObjects コレクション) のコレクションを表すオブジェクトを返します。
'*
'* @param Index OLE オブジェクトの名前またはインデックス番号を指定。省略可能。
'* @return Object OLEObject / OLEObjects オブジェクト
'*
'******************************************************************************
Function OLEObjects(Optional Index) As Object
    Set OLEObjects = mOrigin.OLEObjects(Index)
End Function
    
'******************************************************************************
'* [概  要] Paste メソッド
'* [詳  細] クリップボードの内容をシートに貼り付けます。
'*
'* @param Destination クリップボードのコンテンツを貼り付ける場所を指定する Range オブジェクト。省略可能。
'* @param Link 貼り付けたデータと元のデータの間にリンクを設定するには、 True を指定。省略可能。
'*
'******************************************************************************
Sub Paste(Optional Destination, Optional Link)
    Call mOrigin.Paste(Destination, Link)
End Sub

'******************************************************************************
'* [概  要] PasteSpecial メソッド
'* [詳  細] 指定された形式で、クリップボードの内容をシートに貼り付けます。 他のアプリケーションからデータを貼り付けるときや、あるいは特別な形式でデータを貼り付ける場合に使います。
'*
'* @param Format クリップボードのデータの形式を文字列で指定。省略可能。
'* @param Link 元のデータと貼り付けたデータの間にリンクを設定するには、 True を指定。省略可能。既定値は False。
'* @param DisplayAsIcon True の場合は、貼り付けられたデータがアイコンとして表示。省略可能。既定値は False。
'* @param IconFileName DisplayAsIcon がTrueの場合に使用するアイコンを含むファイルの名前を指定。省略可能。
'* @param IconIndex アイコン ファイルのどのアイコンを使うかを示すインデックス番号を指定。省略可能。
'* @param IconLabel アイコンのラベルを文字列で指定。省略可能。
'* @param NoHTMLFormatting True を設定すると、HTML から書式設定、ハイパーリンク、およびイメージを削除。省略可能。既定値は False。
'*
'******************************************************************************
Sub PasteSpecial(Optional Format, Optional Link, Optional DisplayAsIcon, Optional IconFileName, Optional iconIndex, Optional IconLabel, Optional NoHTMLFormatting)
    Call mOrigin.PasteSpecial(Format, Link, DisplayAsIcon, IconFileName, iconIndex, IconLabel, NoHTMLFormatting)
End Sub

'******************************************************************************
'* [概  要] PivotTables メソッド
'* [詳  細] ワークシートの1つのピボットテーブルレポート ( PivotTable オブジェクト)、またはすべてのピボットテーブルレポート (PivotTables オブジェクト) のコレクションを表すオブジェクトを返します。
'*
'* @param Index ピボットテーブル レポートの名前またはインデックス番号を指定。省略可能。
'* @return Object PivotTable / PivotTables オブジェクト
'*
'******************************************************************************
Function PivotTables(Optional Index) As Object
    Set PivotTables = mOrigin.PivotTables(Index)
    mOrigin.PivotTableWizard
End Function

'******************************************************************************
'* [概  要] PivotTableWizard メソッド
'* [詳  細] ピボットテーブル レポートを新規作成します。 このメソッドは、ピボットテーブルウィザードを表示しません。 このメソッドは、OLE DB データソースでは使用できません。 Add メソッドを使用してピボットテーブル キャッシュを作成した後、そのキャッシュを基にピボットテーブル レポートを作成します。
'*
'* @param SourceType レポートデータのソースを表す Xlpivotwhere Ourcetype 値。省略可能。
'* @param SourceData レポートの作成に使うデータを指定 。省略可能。
'* @param TableDestination レポートの配置場所を表す Range オブジェクトを指定 。省略可能。
'* @param TableName 新しいレポートの名前 。省略可能。
'* @param RowGrand レポートに行の総計を表示するには、 True を指定 。省略可能。
'* @param ColumnGrand レポートに列の総計を表示するには、 True を指定 。省略可能。
'* @param SaveData レポートと共にデータを保存するには、 True を指定 。省略可能。
'* @param HasAutoFormat Trueの場合、レポートが更新されたとき、またはフィールドが移動されたときに、自動的に書式を設定 。省略可能。
'* @param AutoPage SourceType_がxlconsolidationの場合にのみ有効です。 Trueを設定すると、統合するページフィールドが作成されます。 。省略可能。
'* @param Reserved 使用されません 。省略可能。
'* @param BackgroundQuery True の場合、Excel はレポートのクエリをバック グラウンド (非同期) で実行します。 既定値は False です。 。省略可能。
'* @param OptimizeCache True の場合、ピボットテーブルのキャッシュの構成時に最適化を行います。 既定値は False です。 。省略可能。
'* @param PageFieldOrder ピボットテーブルレポートのレイアウトにページフィールドを追加する順序を指定します。 。省略可能。
'* @param PageFieldWrapCount ピボットテーブルの各列または各行のページ フィールドの数を指定します。 既定値は 0 (ゼロ) です。 。省略可能。
'* @param ReadData True の場合は、外部データベースのすべてのレコードを含むピボットテーブルキャッシュを作成します。 。省略可能。
'* @param Connection ExcelがODBCデータソースに接続できるようにするODBC設定を含む文字列 。省略可能。
'* @return PivotTable PivotTableオブジェクト
'*
'******************************************************************************
Function PivotTableWizard(Optional SourceType, Optional SourceData, Optional TableDestination, Optional TableName, Optional RowGrand, Optional ColumnGrand, Optional SaveData, Optional HasAutoFormat, Optional AutoPage, Optional Reserved, Optional BackgroundQuery, Optional OptimizeCache, Optional PageFieldOrder, Optional PageFieldWrapCount, Optional ReadData, Optional Connection) As PivotTable
    Set PivotTableWizard = mOrigin.PivotTableWizard(SourceType, SourceData, TableDestination, TableName, RowGrand, ColumnGrand, SaveData, HasAutoFormat, AutoPage, Reserved, BackgroundQuery, OptimizeCache, PageFieldOrder, PageFieldWrapCount, ReadData, Connection)
End Function

'******************************************************************************
'* [概  要] PrintOut メソッド
'* [詳  細] シートを印刷します。
'*
'* @param From 印刷を開始するページの番号を指定。省略可能。
'* @param To_ 印刷を終了するページの番号を指定。省略可能。
'* @param Copies 印刷部数を指定。省略可能。
'* @param Preview True の場合、印刷をする前に印刷プレビューを実行。省略可能。
'* @param ActivePrinter アクティブなプリンターの名前を指定。省略可能。
'* @param PrintToFile True の場合、ファイルへ出力。省略可能。
'* @param Collate True の場合、部単位で印刷。省略可能。
'* @param PrToFileName PrintToFileがTrueに設定されている場合、印刷先のファイル名を指定。省略可能。
'* @param IgnorePrintAreas True の場合、印刷範囲を無視してシート全体を印刷。省略可能。
'*
'******************************************************************************
Sub PrintOut(Optional From, Optional To_, Optional Copies, Optional Preview, Optional ActivePrinter, Optional PrintToFile, Optional Collate, Optional PrToFileName, Optional IgnorePrintAreas)
    Call mOrigin.PrintOut(From, To_, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName, IgnorePrintAreas)
End Sub

'******************************************************************************
'* [概  要] PrintPreview メソッド
'* [詳  細] 印刷プレビュー (印刷時のイメージ) を表示します。
'*
'* @param EnableChanges ユーザーが印刷プレビューで余白などのページ設定オプションを変更できるかどうかを指定。省略可能。
'*
'******************************************************************************
Sub PrintPreview(Optional EnableChanges)
    Call mOrigin.PrintPreview(EnableChanges)
End Sub

'******************************************************************************
'* [概  要] Protect メソッド
'* [詳  細] ワークシートを保護し、変更できないようにします。
'*
'* @param Password ワークシートまたはブックのパスワード文字列を指定。省略可能。
'* @param DrawingObjects 描画オブジェクトを保護するには、 True を指定します。 既定値は True です。省略可能。
'* @param Contents Trueの内容を保護します。 グラフの場合は、グラフ全体が保護されます。 ワークシートの場合は、ロックされたセルを保護します。 既定値は True です。省略可能。
'* @param Scenarios シナリオを保護するには、 True を指定します。 この引数はワークシートの場合のみ有効です。 既定値は True です。省略可能。
'* @param UserInterfaceOnly True を指定すると、画面上からの変更は保護されますが、マクロからの変更は保護されません。 この引数を省略すると、マクロからも、画面上からも変更することができなくなります。省略可能。
'* @param AllowFormattingCells True を指定すると、ユーザーは保護されたワークシートのセルを書式化することができます。 既定値は False です。省略可能。
'* @param AllowFormattingColumns True を指定すると、ユーザーは保護されたワークシートの列を書式化することができます。 既定値は False です。省略可能。
'* @param AllowFormattingRows Trueを指定すると、ユーザーは保護されたワークシートの任意の行を書式設定できます。 既定値は False です。省略可能。
'* @param AllowInsertingColumns True を指定すると、ユーザーは保護されたワークシートに列を挿入することができます。 既定値は False です。省略可能。
'* @param AllowInsertingRows True を指定すると、ユーザーは保護されたワークシートに行を挿入することができます。 既定値は False です。省略可能。
'* @param AllowInsertingHyperlinks Trueは、保護されたワークシートにハイパーリンクを挿入することをユーザーに許可します。 既定値は False です。省略可能。
'* @param AllowDeletingColumns True を指定すると、ユーザーは保護されたワークシートの列を削除することができ、削除される列のセルはすべてロック解除されます。 既定値は False です。省略可能。
'* @param AllowDeletingRows True を指定すると、ユーザーは保護されたワークシートの行を削除することができ、削除される行のセルはすべてロック解除されます。 既定値は False です。省略可能。
'* @param AllowSorting True を指定すると、ユーザーは保護されたワークシートで並べ替えを行うことができます。 並べ替え範囲内のセルは、ロックと保護が解除されている必要があります。 既定値は False です。省略可能。
'* @param AllowFiltering True を指定すると、ユーザーは保護されたワークシートにフィルターを設定することができます。 ユーザーは、フィルター条件を変更できますが、オートフィルターの有効と無効を切り替えることはできません。 ユーザーは、既存のオート フィルターにフィルターを設定できます。 既定値は False です。省略可能。
'* @param AllowUsingPivotTables Trueは、保護されたワークシートでピボットテーブルレポートを使用することをユーザーに許可します。 既定値は False です。省略可能。
'*
'******************************************************************************
Sub Protect(Optional Password, Optional DrawingObjects, Optional Contents, Optional Scenarios, Optional UserInterfaceOnly, Optional AllowFormattingCells, Optional AllowFormattingColumns, Optional AllowFormattingRows, Optional AllowInsertingColumns, Optional AllowInsertingRows, Optional AllowInsertingHyperlinks, Optional AllowDeletingColumns, Optional AllowDeletingRows, Optional AllowSorting, Optional AllowFiltering, Optional AllowUsingPivotTables)
    Call mOrigin.Protect(Password, DrawingObjects, Contents, Scenarios, UserInterfaceOnly, AllowFormattingCells, AllowFormattingColumns, AllowFormattingRows, AllowInsertingColumns, AllowInsertingRows, AllowInsertingHyperlinks, AllowDeletingColumns, AllowDeletingRows, AllowSorting, AllowFiltering, AllowUsingPivotTables)
End Sub

'******************************************************************************
'* [概  要] ResetAllPageBreaks メソッド
'* [詳  細] 指定したワークシートのすべての改ページをリセットします。
'*
'******************************************************************************
Sub ResetAllPageBreaks()
    Call mOrigin.ResetAllPageBreaks
End Sub

'******************************************************************************
'* [概  要] SaveAs メソッド
'* [詳  細] グラフまたはワークシートへの変更を別のファイルに保存します。
'*
'* @param fileName 保存するファイルの名前。省略可能。
'* @param FileFormat ファイルを保存するときに使用するファイル形式を指定。省略可能。
'* @param Password ファイルを保護するためのパスワードを表す 15 文字以内の文字列を指定。省略可能。
'* @param WriteResPassword ファイルの書き込みパスワードを表す文字列を指定。省略可能。
'* @param ReadOnlyRecommended ファイルを開くとき、ファイルを読み取り専用として開くように勧めるメッセージを表示するには、 True を指定。省略可能。
'* @param CreateBackup バックアップ ファイルを作成するには、 True を指定。省略可能。
'* @param AddToMru 最近使用したファイルの一覧にブックを追加するには、 True を指定します。 既定値は False です。省略可能。
'* @param TextCodepage 英語の Excel では使用できません。省略可能。
'* @param TextVisualLayout 英語の Excel では使用できません。省略可能。
'* @param Local_ Trueは、Excel の言語 (コントロールパネルの設定を含む) に対してファイルを保存します。省略可能。
'*
'******************************************************************************
Sub SaveAs(FileName As String, Optional FileFormat, Optional Password, Optional WriteResPassword, Optional ReadOnlyRecommended, Optional CreateBackup, Optional AddToMru, Optional TextCodepage, Optional TextVisualLayout, Optional Local_)
    Call mOrigin.SaveAs(FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AddToMru, TextCodepage, TextVisualLayout, Local_)
End Sub

'******************************************************************************
'* [概  要] Scenarios メソッド
'* [詳  細] ワークシートの単一のシナリオ (Scenario オブジェクト)、またはシナリオのコレクション ( Scenarios オブジェクト) を表すオブジェクトを返します。
'*
'* @param Index シナリオの名前またはインデックス番号を指定。省略可能。
'* @return Object Scenario / Scenarios オブジェクト
'*
'******************************************************************************
Function Scenarios(Optional Index) As Object
    Set Scenarios = mOrigin.Scenarios(Index)
End Function

'******************************************************************************
'* [概  要] Select_ メソッド
'* [詳  細] オブジェクトを選択します。単一のシートをアクティブシートにするには、 Activate メソッドを使用します。
'*
'* @param rep 複数シートでのみ使用。省略可能。
'*
'******************************************************************************
Sub Select_(Optional rep)
    Call mOrigin.Select(rep)
End Sub

'******************************************************************************
'* [概  要] SetBackgroundPicture メソッド
'* [詳  細] ワークシートの背景のグラフィックスを設定します。
'*
'* @param fileName 表示するグラフィック ファイルを指定
'*
'******************************************************************************
Sub SetBackgroundPicture(FileName As String)
    Call mOrigin.SetBackgroundPicture(FileName)
End Sub

'******************************************************************************
'* [概  要] ShowAllData メソッド
'* [詳  細] 現在フィルターがかけられているリストに対して、すべての行を表示します。 オートフィルターを使っているときにこのメソッドを使うと、下向き矢印が青から黒に変わります。
'*
'******************************************************************************
Sub ShowAllData()
    Call mOrigin.ShowAllData
End Sub

'******************************************************************************
'* [概  要] ShowDataForm メソッド
'* [詳  細] 対象ワークシートに関連付けられたデータ フォームを表示します。
'*
'******************************************************************************
Sub ShowDataForm()
    Call mOrigin.ShowDataForm
End Sub

'******************************************************************************
'* [概  要] Unprotect メソッド
'* [詳  細] シートまたはブックの保護を解除します。 保護されていないシートやブックに対しては、このメソッドは無効です。
'*
'* @param Password パスワード。省略可能。
'*
'******************************************************************************
Sub Unprotect(Optional Password)
    Call mOrigin.Unprotect(Password)
End Sub

'******************************************************************************
'* [概  要] XmlDataQuery メソッド
'* [詳  細] 特定の XPath にマップされたセルを表す Range オブジェクトを返します。 指定した XPath がワークシートにマッピングされていない場合、またはマッピングされた範囲が空の場合は、Nothing を返します。
'*
'* @param XPath 対象となる XPath を指定。
'* @param SelectionNamespaces XPath パラメーターで参照される名前空間を含む、スペースで区切られた文字列。省略可能。
'* @param Map 特定のマップ内の XPath を照会する場合は、 XmlMap を指定。省略可能。
'* @return Range 特定のXPathにマップされたセルを表すRangeオブジェクト
'*
'******************************************************************************
Function XmlDataQuery(XPath As String, Optional SelectionNamespaces, Optional Map) As Range
    Set XmlDataQuery = mOrigin.XmlDataQuery(XPath, SelectionNamespaces, Map)
End Function

'******************************************************************************
'* [概  要] XmlMapQuery メソッド
'* [詳  細] 特定の XPath にマップされたセルを表す Range オブジェクトを返します。 指定した XPath がワークシートにマッピングされていない場合、またはマッピングされた範囲が空の場合は、Nothing を返します。
'*          XmlDataQuery メソッドとは異なり、ヘッダー行を含む XML リストの列全体を返します。
'*
'* @param XPath 対象となる XPath を指定。
'* @param SelectionNamespaces XPath パラメーターで参照される名前空間を含む、スペースで区切られた文字列。省略可能。
'* @param Map 特定のマップ内の XPath を照会する場合は、 XmlMap を指定。省略可能。
'* @return Range 特定のXPathにマップされたセルを表すRangeオブジェクト
'*
'******************************************************************************
Function XmlMapQuery(XPath As String, Optional SelectionNamespaces, Optional Map) As Range
    Set XmlMapQuery = mOrigin.XmlMapQuery(XPath, SelectionNamespaces, Map)
End Function

'*-----------------------------------------------------------------------------
'* 拡張メソッド/プロパティ
'*-----------------------------------------------------------------------------
'*-----------------------------------------------------------------------------
'* Range2 プロパティ（読み取り専用）
'*
'* [補  足]
'* ・セルまたはセル範囲を表す Range オブジェクトを返す。
'* ・Range(Cells(x1, y1), Cells(x2, y2))を簡易指定できるようにしたもの
'*-----------------------------------------------------------------------------
Public Property Get Range2(startRow As Long, startCol As Long, endRow As Long, endCol As Long) As Range
    Set Range2 = mOrigin.Range(mOrigin.Cells(startRow, startCol), mOrigin.Cells(endRow, endCol))
End Property

'*-----------------------------------------------------------------------------
'* SpecialCells プロパティ（読み取り専用）
'*
'* [補  足]
'* ・特定の種類の Range オブジェクトを返す。
'* ・指定した種類のセルがない場合は、 Nothing を返す。
'* ・Range(Cells(x1, y1), Cells(x2, y2)).SpecialCells を簡易指定できるようにしたもの
'*-----------------------------------------------------------------------------
Public Property Get SpecialCells(startRow As Long, startCol As Long, endRow As Long, endCol As Long, typeOfCell As XlCellType, Optional typeOfval As XlSpecialCellsValue = 0) As Range
    Dim rng As Range
    On Error Resume Next
    If typeOfval = 0 Then
        Set rng = Range2(startRow, startCol, endRow, endCol).SpecialCells(typeOfCell)
    Else
        Set rng = Range2(startRow, startCol, endRow, endCol).SpecialCells(typeOfCell, typeOfval)
    End If
    On Error GoTo 0
    Err.Clear
    Set SpecialCells = rng
End Property

'*-----------------------------------------------------------------------------
'* CurrentRegion プロパティ（読み取り専用）
'*
'* [補  足]
'* ・アクティブセル領域の Range オブジェクトを返す。
'* ・Cells(x, y).CurrentRegion を簡易指定できるようにしたもの
'*-----------------------------------------------------------------------------
Public Property Get CurrentRegion(rowNo As Long, colNo As Long) As Range
    Set CurrentRegion = mOrigin.Cells(rowNo, colNo).CurrentRegion
End Property

'*-----------------------------------------------------------------------------
'* EntireRow プロパティ（読み取り専用）
'*
'* [補  足]
'* ・行全体を表す Range オブジェクトを返す。
'* ・Cells(x, 1).EntireRow を簡易指定できるようにしたもの
'*-----------------------------------------------------------------------------
Public Property Get EntireRow(rowNo As Long, Optional endRowNo As Long) As Range
    If endRowNo = 0 Then
        Set EntireRow = mOrigin.Cells(rowNo, 1).EntireRow
    Else
        Set EntireRow = Range2(rowNo, 1, endRowNo, 1).EntireRow
    End If
End Property

'*-----------------------------------------------------------------------------
'* EntireColumn プロパティ（読み取り専用）
'*
'* [補  足]
'* ・列全体を表す Range オブジェクトを返す。
'* ・Cells(1, y).EntireColumn を簡易指定できるようにしたもの
'*-----------------------------------------------------------------------------
Public Property Get EntireColumn(colNo As Long, Optional endColNo As Long) As Range
    If endColNo = 0 Then
        Set EntireColumn = mOrigin.Cells(1, colNo).EntireColumn
    Else
        Set EntireColumn = Range2(1, colNo, 1, endColNo).EntireColumn
    End If
End Property

'*-----------------------------------------------------------------------------
'* RangeSelection プロパティ（読み取り専用）
'*
'* [補  足]
'* ・選択中のセル範囲を返却する。
'*-----------------------------------------------------------------------------
Public Property Get RangeSelection() As Range
    Set RangeSelection = XlWorkSheetUtils.GetRangeSelection(mOrigin.Name, mOrigin.Parent)
End Property

'*-----------------------------------------------------------------------------
'* VisibleRange プロパティ（読み取り専用）
'*
'* [補  足]
'* ・表示中のセル範囲を返却する。
'*-----------------------------------------------------------------------------
Public Property Get VisibleRange() As Range
    Set VisibleRange = XlWorkSheetUtils.GetVisibleRange(mOrigin.Name, mOrigin.Parent)
End Property

'******************************************************************************
'* [概  要] ImportArray シート貼り付け処理。
'* [詳  細] Variant配列データをシートに出力する。
'*
'* @param vArray Variant配列データ
'* @param lStartRow データ開始行番号
'* @param lStartCol データ開始列番号
'* @param transpose Trueにすると1次元配列を列方向に設定する
'* @param willTxtFormat 書式を文字列にして出力するか
'*
'******************************************************************************
Public Sub ImportArray(vArray, lStartRow As Long, lStartCol As Long, Optional Transpose As Boolean = True, Optional willTxtFormat As Boolean = False)
    Call StoreFilterCriterias
    Call FilterOff

    If ValidateUtils.IsArray2DEx(vArray) Then
        Call XlWorkSheetUtils.InjectVariantArrayToCells(mOrigin, vArray.To2DArray, lStartRow, lStartCol, willTxtFormat)
    ElseIf ArrayUtils.Is2DArray(vArray) Then
        Call XlWorkSheetUtils.InjectVariantArrayToCells(mOrigin, vArray, lStartRow, lStartCol, willTxtFormat)
    ElseIf ValidateUtils.IsArrayEx(vArray) Then
        Call XlWorkSheetUtils.InjectVariantArrayToCells(mOrigin, vArray.To2DArray(Transpose), lStartRow, lStartCol, willTxtFormat)
    ElseIf ArrayUtils.Is1DArray(vArray) Then
        Call XlWorkSheetUtils.InjectVariantArrayToCells(mOrigin, ArrayUtils.ArrayTo2DArray(vArray, Transpose), lStartRow, lStartCol, willTxtFormat)
    ElseIf ValidateUtils.IsPrimitive(vArray) Then
        If willTxtFormat Then mOrigin.Cells(lStartRow, lStartCol).NumberFormat = "@"
        mOrigin.Cells(lStartRow, lStartCol).Value = vArray
    End If
    
    Call ReConstructFilter
End Sub

'******************************************************************************
'* [概  要] NumbersToIndexCells 項番設定処理。
'* [詳  細] 項番に連番を出力する。
'*
'* @param lStartRow データ開始行番号
'* @param lStartCol データ開始列番号
'* @param rowNum 番号数
'*
'******************************************************************************
Public Sub NumbersToIndexCells(lStartRow As Long, lStartCol As Long, rowNum As Long)
    Call XlWorkSheetUtils.InjectNumbersToIndexCells(mOrigin, lStartRow, lStartCol, rowNum)
End Sub

'******************************************************************************
'* [概  要] GetFinalRow 最終行取得処理。
'* [詳  細] WorksheetのUsedRangeを下から走査し、最終行番号を取得する｡
'*
'* @param ignoreColnum 走査対象外の列番号
'* @return 最終行番号
'*
'******************************************************************************
Public Function GetFinalRow(Optional ignoreColnum As Long) As Long
    GetFinalRow = XlWorkSheetUtils.GetFinalRow(mOrigin, ignoreColnum)
End Function

'******************************************************************************
'* [概  要] GetFinalCol 最終列取得処理。
'* [詳  細] WorksheetのUsedRangeを右から走査し、最終列番号を取得する｡
'*
'* @return 最終列番号
'*
'******************************************************************************
Public Function GetFinalCol() As Long
    GetFinalCol = XlWorkSheetUtils.GetFinalCol(mOrigin)
End Function

'******************************************************************************
'* [概  要] GetFinalKeyRow 指定列最終入力行取得処理。
'* [詳  細] 指定列の最終入力行番号を取得する｡
'*
'* @param dataSheet ワークシート
'* @param startRow 対象列の番号
'* @return 最終行番号
'*
'******************************************************************************
Public Function GetFinalKeyRow(startCol As Long) As Long
    GetFinalKeyRow = XlWorkSheetUtils.GetFinalKeyRow(mOrigin, startCol)
End Function

'******************************************************************************
'* [概  要] GetFinalKeyCol 指定行最終入力列取得処理。
'* [詳  細] 指定行の最終入力列番号を取得する｡
'*
'* @param dataSheet ワークシート
'* @param startRow 対象行の番号
'* @return 最終列番号
'*
'******************************************************************************
Public Function GetFinalKeyCol(startRow As Long) As Long
    GetFinalKeyCol = XlWorkSheetUtils.GetFinalKeyCol(mOrigin, startRow)
End Function

'******************************************************************************
'* [概  要] ExportArray 表情報取得処理。
'* [詳  細] worksheetの表から情報を取得し、Variant配列を返却｡
'*
'* @param lStartRow データ開始行番号
'* @param lStartCol データ開始列番号
'* @param itemCount 項目列数
'* @return Variant2次元配列
'*
'******************************************************************************
Public Function ExportArray(lStartRow As Long, lStartCol As Long, Optional colCount As Long, Optional startZeroIndex As Boolean = False)
    ExportArray = XlWorkSheetUtils.GetVariantDataFromSheet(mOrigin, lStartRow, lStartCol, colCount)
    If startZeroIndex Then
        ExportArray = ArrayUtils.Copy2DArray(ExportArray)
    End If
End Function

'******************************************************************************
'* [概  要] ExportArray2DEx 表情報取得処理。
'* [詳  細] worksheetの表から情報を取得し、Array2DExを返却｡
'*
'* @param lStartRow データ開始行番号
'* @param lStartCol データ開始列番号
'* @param itemCount 項目列数
'* @return Array2DEx
'*
'******************************************************************************
Public Function ExportArray2DEx(lStartRow As Long, lStartCol As Long, Optional colCount As Long) As Array2DEx
    Set ExportArray2DEx = Core.Init(New Array2DEx, XlWorkSheetUtils.GetVariantDataFromSheet(mOrigin, lStartRow, lStartCol, colCount))
End Function

'******************************************************************************
'* [概  要] GetRowData 表情報取得処理。
'* [詳  細] worksheetの表から1行分の情報を取得し、Variant配列を返却｡
'*
'* @param lStartRow データ開始行番号
'* @param lStartCol データ開始列番号
'* @param itemCount 項目列数
'* @return Variant 1行分のデータ
'*
'******************************************************************************
Public Function GetRowData(lStartRow As Long, Optional lStartCol As Long, Optional colCount As Long) As Variant
    GetRowData = XlWorkSheetUtils.GetRowData(mOrigin, lStartRow, lStartCol, colCount)
End Function

'******************************************************************************
'* [概  要] GetRowToArrayEx 表情報取得処理。
'* [詳  細] worksheetの表から1行分の情報を取得し、Variant配列を返却｡
'*
'* @param lStartRow データ開始行番号
'* @param lStartCol データ開始列番号
'* @param itemCount 項目列数
'* @return ArrayEx 1行分のデータ
'*
'******************************************************************************
Public Function GetRowToArrayEx(lStartRow As Long, Optional lStartCol As Long, Optional colCount As Long) As ArrayEx
    Set GetRowToArrayEx = Core.Init(New ArrayEx, XlWorkSheetUtils.GetRowData(mOrigin, lStartRow, lStartCol, colCount))
End Function

'******************************************************************************
'* [概  要] ClearActualUsedRange 使用セル範囲クリア処理。
'* [詳  細] worksheetのデータ表の使用セル範囲をクリア｡
'*
'* @param lStartRow data表データ開始行番号
'* @param lStartCol data表データ開始列番号
'* @param itemCount 項目列数
'* @param ignoreColnum 走査対象外の列番号
'*
'******************************************************************************
Public Sub ClearActualUsedRange(lStartRow As Long, _
                                lStartCol As Long, _
                                Optional colCount As Long, _
                                Optional ignoreColnum As Long)
    Call XlWorkSheetUtils.ClearActualUsedRangeFromSheet(mOrigin, lStartRow, lStartCol, colCount, ignoreColnum)
End Sub

'******************************************************************************
'* [概  要] DeleteNoUsedRange 未使用範囲行削除処理。
'* [詳  細] worksheetのデータ表の未使用範囲行を削除（UsedRangeを縮小）｡
'*
'* @param lStartRow data表データ開始行番号
'*
'******************************************************************************
Public Sub DeleteNoUsedRange(lStartRow As Long)
    Call XlWorkSheetUtils.DeleteNoUsedRange(mOrigin, lStartRow)
End Sub

'******************************************************************************
'* [概  要] GetActualUsedRange 使用セル範囲取得処理。
'* [詳  細] worksheetのデータ表の使用セル範囲を取得する｡
'*
'* @param lStartRow data表データ開始行番号
'* @param lStartCol data表データ開始列番号
'* @param itemCount 項目列数
'* @param ignoreColnum 走査対象外の列番号
'* @return 使用セル範囲
'*
'******************************************************************************
Public Function GetActualUsedRange(lStartRow As Long, lStartCol As Long, Optional colCount As Long, Optional ignoreColnum As Long) As Range
    Set GetActualUsedRange = XlWorkSheetUtils.GetActualUsedRangeFromSheet(mOrigin, lStartRow, lStartCol, colCount, ignoreColnum)
End Function

Public Function GetActualUsedRangeToArray2DEx(lStartRow As Long, lStartCol As Long, Optional colCount As Long, Optional ignoreColnum As Long) As Array2DEx
    Dim rng As Range, vArr As Variant
    Set rng = XlWorkSheetUtils.GetActualUsedRangeFromSheet(mOrigin, lStartRow, lStartCol, colCount, ignoreColnum)
    If rng Is Nothing Then
        Set GetActualUsedRangeToArray2DEx = New Array2DEx
    Else
        vArr = rng
        Set GetActualUsedRangeToArray2DEx = Core.Init(New Array2DEx, vArr)
    End If
End Function

'******************************************************************************
'* [概  要] SetHyperLink
'* [詳  細] ハイパーリンクを追加する（フォント・フォントサイズの変更をしない）
'*
'* @param rowNo 行番号
'* @param colNo 列番号
'* @param strLinkAddress リンクアドレス
'*
'******************************************************************************
Public Sub SetHyperLink(rowNo As Long, colNo As Long, strLinkAddress As String)
    Call XlWorkSheetUtils.SetHyperLink(mOrigin, rowNo, colNo, strLinkAddress)
End Sub

'******************************************************************************
'* [概  要] CopySheet シートコピー処理
'* [詳  細] シートをコピーし、シート名+ nの名称
'*          （nは重複しない番号）をつけて、対象のWorksheetオブジェクトを返却。
'*
'* @param destSheetName 作成するシート名
'* @param wb 対象のワークブック
'* @param countAppend 番号を付与するか
'* @return 作成したシートオブジェクト
'******************************************************************************
Public Function CopySheet(ByVal destSheetName As String, _
                   Optional ByRef wb As Workbook, Optional ByVal countAppend As Boolean = False) As Worksheet
    Set CopySheet = XlWorkSheetUtils.CopySheet(mOrigin.Name, destSheetName, wb, countAppend)
End Function

'******************************************************************************
'* [概  要] StoreFilterCriterias
'* [詳  細] オートフィルタの条件を記憶する。
'* [参  考] http://yamav102.cocolog-nifty.com/blog/2018/02/post-82ac.html
'*
'******************************************************************************
Public Sub StoreFilterCriterias()
    If mOrigin.AutoFilter Is Nothing Then Exit Sub

    Dim fltr As AutoFilter: Set fltr = mOrigin.AutoFilter
    Dim clms As Filters: Set clms = fltr.Filters
    
    If fltr.FilterMode Then
        ReDim mCriterias(1 To clms.Count)
        Dim i As Long
        For i = 1 To clms.Count
            Set mCriterias(i) = New Collection

            If clms.Item(i).On Then
                mCriterias(i).Add clms.Item(i).Operator, "operator"

                If clms.Item(i).Count > 0 Then
                    mCriterias(i).Add clms.Item(i).Criteria1, "Criteria1"
                End If
                If (7 <> clms.Item(i).Operator) And (clms.Item(i).Count > 1) Then
                    mCriterias(i).Add clms.Item(i).Criteria2, "Criteria2"
                End If
            End If
        Next
    End If
End Sub

'******************************************************************************
'* [概  要] SetAutoFilter
'* [詳  細] オートフィルタを設定する。
'*
'* @param startRowNo 開始行番号
'* @param startColNo 開始列番号
'* @param endRowNo   終了行番号
'* @param endColNo   終了列番号
'*
'******************************************************************************
Public Function SetAutoFilter(ByVal startRowNo As Long, ByVal startColNo As Long, ByVal endRowNo As Long, endColNo As Long) As WorkSheetEx
    If mOrigin.AutoFilterMode = True Then
        Set SetAutoFilter = Me
        Exit Function
    End If

    Dim rng As Range
    Set rng = Range2(startRowNo, startColNo, startRowNo, endColNo)
    Call rng.AutoFilter
    Set SetAutoFilter = Me
End Function

'******************************************************************************
'* [概  要] RemoveAutoFilter
'* [詳  細] オートフィルタを解除する。
'*
'******************************************************************************
Public Sub RemoveAutoFilter()
    If mOrigin.AutoFilterMode = False Then
        Exit Sub
    End If

    mOrigin.AutoFilter.Range.AutoFilter
End Sub

'******************************************************************************
'* [概  要] FilterOn
'* [詳  細] オートフィルタのフィルタ条件を指定する。
'*
'* @param choiceTable テーブルオブジェクトの行数を取得するか
'* @param choiceTable テーブルオブジェクト
'******************************************************************************
Public Sub FilterOn(Optional Field, Optional Criteria1, Optional Operator As XlAutoFilterOperator = xlAnd, Optional Criteria2, Optional VisibleDropDown)
    If mOrigin.AutoFilter Is Nothing Then Exit Sub
    If mOrigin.AutoFilterMode = False Then
        Exit Sub
    End If
    
    Call mOrigin.AutoFilter.Range.AutoFilter(Field, Criteria1, Operator, Criteria2, VisibleDropDown)
End Sub

'******************************************************************************
'* [概  要] FilterOff
'* [詳  細] オートフィルタのフィルタ条件を解除して全行表示する。
'* [参  考] http://yamav102.cocolog-nifty.com/blog/2018/02/post-82ac.html
'*
'******************************************************************************
Public Sub FilterOff()
    If mOrigin.AutoFilter Is Nothing Then Exit Sub
    
    If mOrigin.FilterMode Then
        mOrigin.ShowAllData
    End If
End Sub
'******************************************************************************
'* [概  要] CountOfFilter
'* [詳  細] オートフィルタで絞り込んだ行数を取得する。
'* [参  考] https://www.excel-chunchun.com/entry/20200613-excel-vba-autofilter-rows-count
'*
'* @param choiceTable テーブルオブジェクトの行数を取得するか
'* @param choiceTable テーブルオブジェクト
'* @return オートフィルタで絞り込んだ行数
'******************************************************************************
Public Function CountOfFilter(Optional choiceTable As Boolean = False, Optional tblNo As Long = 1) As Long
    If choiceTable Then
        If mOrigin.ListObjects.Count < tblNo Then Exit Function
        CountOfFilter = mOrigin.ListObjects(tblNo).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    Else
        If mOrigin.AutoFilter Is Nothing Then Exit Function
        
        ' ActiveCellがテーブル（ListObject）の中にある時、AutoFilterはListObject.AutoFilterを返すため
        ' 非テーブルの表からの結果が得られないことに注意
        CountOfFilter = mOrigin.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
    End If
End Function

'******************************************************************************
'* [概  要] ReConstructFilter
'* [詳  細] フィルタ条件を復元する。
'* [参  考] http://yamav102.cocolog-nifty.com/blog/2018/02/post-82ac.html
'*
'******************************************************************************
Public Sub ReConstructFilter()
    If mOrigin.AutoFilter Is Nothing Then Exit Sub
    If ArrayUtils.IsEmptyArray(mCriterias) Then Exit Sub
    
    Dim rng As Range: Set rng = mOrigin.AutoFilter.Range
    Dim i As Long
    For i = 1 To mOrigin.AutoFilter.Filters.Count
        With mCriterias(i)
            If CollectionUtils.HasKey(mCriterias(i), "Criteria2") Then
                rng.AutoFilter Field:=i, _
                               Criteria1:=.Item("Criteria1"), _
                               Criteria2:=.Item("Criteria2"), _
                               Operator:=.Item("operator")

            ElseIf CollectionUtils.HasKey(mCriterias(i), "Criteria1") Then
                rng.AutoFilter Field:=i, _
                               Criteria1:=.Item("Criteria1"), _
                               Operator:=.Item("operator")
            End If
        End With
    Next
End Sub

'******************************************************************************
'* [概  要] SetCenterAcrossSelection
'* [詳  細] 指定範囲に「選択範囲内で中央揃え」を設定する。
'*
'* @param startRowNo 開始行番号
'* @param startColNo 開始列番号
'* @param endRowNo   終了行番号
'* @param endColNo   終了列番号
'*
'******************************************************************************
Public Sub SetCenterAcrossSelection(startRowNo As Long, startColNo As Long, endRowNo As Long, endColNo As Long)
    Dim targetRng As Range
    Set targetRng = Range2(startRowNo, startColNo, endRowNo, endColNo)
    Call XlWorkSheetUtils.SetCenterAcrossSelection(targetRng)
End Sub

'******************************************************************************
'* [概  要] SetValidateList
'* [詳  細] 指定範囲に入力規則のリストを設定する。
'*
'* @param startRowNo 開始行番号
'* @param startColNo 開始列番号
'* @param endRowNo   終了行番号
'* @param endColNo   終了列番号
'* @param vlist 入力規則のリスト（1次元配列）
'* @param resetVl  設定を解除するか（True:解除）
'*
'******************************************************************************
Public Sub SetValidateList(startRowNo As Long, startColNo As Long, endRowNo As Long, endColNo As Long, Optional vlist, Optional ByVal resetVl As Boolean = False)
    Dim targetRng As Range
    Set targetRng = Range2(startRowNo, startColNo, endRowNo, endColNo)
    Call XlWorkSheetUtils.SetValidateList(vlist, targetRng, resetVl)
End Sub

'******************************************************************************
'* [概  要] ChangeDisplayGridlines
'* [詳  細] 罫線の表示・非表示を切り替える。
'*
'******************************************************************************
Public Sub ChangeDisplayGridlines()
    Call XlWorkSheetUtils.ChangeDisplayGridlines(mOrigin.Name, mOrigin.Parent)
End Sub

'******************************************************************************
'* [概  要] ChangeDisplayHeadings
'* [詳  細] 行列番号の表示・非表示を切り替える。
'*
'******************************************************************************
Public Sub ChangeDisplayHeadings()
    Call XlWorkSheetUtils.ChangeDisplayHeadings(mOrigin.Name, mOrigin.Parent)
End Sub

'******************************************************************************
'* [概  要] ChangeDisplayFormulas
'* [詳  細] 数式の表示・非表示を切り替える。
'*
'******************************************************************************
Public Sub ChangeDisplayFormulas()
    Call XlWorkSheetUtils.ChangeDisplayFormulas(mOrigin.Name, mOrigin.Parent)
End Sub

'******************************************************************************
'* [概  要] ChangeDisplayOutline
'* [詳  細] アウトライン記号の表示・非表示を切り替える。
'*
'******************************************************************************
Public Sub ChangeDisplayOutline()
    Call XlWorkSheetUtils.ChangeDisplayOutline(mOrigin.Name, mOrigin.Parent)
End Sub

'******************************************************************************
'* [概  要] ChangeDisplayZeros
'* [詳  細] ゼロ値の表示・非表示を切り替える。
'*
'******************************************************************************
Public Sub ChangeDisplayZeros()
    Call XlWorkSheetUtils.ChangeDisplayZeros(mOrigin.Name, mOrigin.Parent)
End Sub

'******************************************************************************
'* [概  要] ChangeZoom
'* [詳  細] 表示倍率を変更する。
'*
'* @param zoomVal 表示倍率
'******************************************************************************
Public Sub ChangeZoom(ByVal zoomVal As Long)
    Call XlWorkSheetUtils.ChangeZoom(zoomVal, mOrigin.Name, mOrigin.Parent)
End Sub

'******************************************************************************
'* [概  要] MergeRangeSilently
'* [詳  細] セルを結合する。結合時の警告は無視する。
'*
'* @param rngAddress 結合するセル範囲のアドレス
'******************************************************************************
Public Sub MergeRangeSilently(ByVal rngAddress As String)
    Call XlWorkSheetUtils.MergeRangeSilently(rngAddress, mOrigin.Name, mOrigin.Parent)
End Sub

'******************************************************************************
'* [概  要] SaveFileOLEObject
'* [詳  細] ファイルを指定したセル位置にOLEObjectとして保存する。
'*
'* @param filePath  対象となるファイルのフルパス
'* @param rowIdx    挿入先セル行インデックス
'* @param colIdx    挿入先セル列インデックス
'******************************************************************************
Public Function SaveFileOLEObject(ByVal filePath As String, _
                             ByVal rowIdx As Long, _
                             ByVal colIdx As Long) As OLEObject
    Set SaveFileOLEObject = XlWorkSheetUtils.SaveFileOLEObject(filePath, rowIdx, colIdx, mOrigin.Name, mOrigin.Parent)
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
    Call XlWorkSheetUtils.LoadFileOLEObject(folderPath, oleName, mOrigin.Name, mOrigin.Parent, oleObj)
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
'* @return Shape 挿入した画像オブジェクト
'******************************************************************************
Public Function InsertPicture(ByVal imgFilePath As String, _
                              ByRef topPos As Variant, _
                              ByRef leftPos As Variant, _
                              ByRef widthVal As Variant, _
                              ByRef heightVal As Variant, _
                              Optional ByVal pLockAspectRatio As MsoTriState = msoTrue, _
                              Optional ByVal pPlacement As XlPlacement = xlMoveAndSize) As Shape
    Set InsertPicture = XlWorkSheetUtils.InsertPicture(imgFilePath, topPos, leftPos, widthVal, heightVal, pLockAspectRatio, pPlacement, mOrigin.Name, mOrigin.Parent)
End Function

'******************************************************************************
'* [概  要] RemoveDuplicates
'* [詳  細] 対象の範囲で指定した列の重複をチェックし、重複行を削除する。
'*
'* @param startRowNo 開始行番号
'* @param startColNo 開始列番号
'* @param endRowNo   終了行番号
'* @param endColNo   終了列番号
'* @param colnums() 重複をチェックする列（Excel列番号ではなく、範囲内の番号）
'*
'******************************************************************************
Public Sub RemoveDuplicates(startRowNo As Long, startColNo As Long, endRowNo As Long, endColNo As Long, ParamArray colnums() As Variant)
    Dim rng As Range
    Set rng = Range2(startRowNo, startColNo, endRowNo, endColNo)
    Dim tmp: tmp = colnums
    Dim arr: arr = ArrayUtils.ToVArrayFlattened(tmp)
    
    Call XlWorkSheetUtils.RemoveDuplicates(rng, arr)
End Sub

'******************************************************************************
'* [概  要] Scroll
'* [詳  細] 指定したセルアドレスが画面左上隅に表示されるようにスクロールする。
'*
'* @param rowNo  スクロール先セル位置行
'* @param colNo  スクロール先セル位置列
'******************************************************************************
Public Sub Scroll(ByVal rowNo As Long, ByVal colNo As Long)
    Call XlWorkSheetUtils.Scroll(rowNo, colNo, mOrigin.Name, mOrigin.Parent)
End Sub

'******************************************************************************
'* [概  要] DataRangeAdress
'* [詳  細] 指定した開始行列番号を起点に、表データの終了行列番号を検索し
'*          表データのデータ範囲を示すRangeアドレスに変換する。
'*
'* @param startRowNo 開始行番号
'* @param startColNo 開始列番号
'* @return String Rangeアドレス
'******************************************************************************
Public Function DataRangeAdress(ByVal startRowNo As Long, ByVal startColNo As Long) As String
    DataRangeAdress = XlWorkSheetUtils.DataRangeAdress(startRowNo, startColNo, mOrigin.Name, mOrigin.Parent)
End Function

'******************************************************************************
'* [概  要] DeleteRowBlank
'* [詳  細] 対象の列範囲でセルが空白の行を削除する。
'*
'* @param colNo      列番号
'* @param startRowNo 開始行番号
'* @param endRowNo   終了行番号
'*
'******************************************************************************
Public Sub DeleteRowBlank(colNo As Long, startRowNo As Long, endRowNo As Long)
    Call XlWorkSheetUtils.DeleteRowBlank(colNo, startRowNo, endRowNo, mOrigin.Name, mOrigin.Parent)
End Sub

'******************************************************************************
'* [概  要] GetVisibleRangeData
'* [詳  細] Range範囲の表示領域のデータを2次元配列で取得する。
'*
'* @param rngAddress  対象とするアドレス。
'* @return Variant 対象範囲の表示領域データ（2次元配列）
'*
'******************************************************************************
Public Function GetVisibleRangeData(ByVal rngAddress As String) As Variant
    GetVisibleRangeData = XlWorkSheetUtils.GetVisibleRangeData(rngAddress, mOrigin.Name, mOrigin.Parent)
End Function

'******************************************************************************
'* [概  要] GetVisibleRangeToArray2DEx
'* [詳  細] Range範囲の表示領域のデータをArray2DExで取得する。
'*
'* @param rngAddress  対象とするアドレス。
'* @return Array2DEx 対象範囲の表示領域データ
'*
'******************************************************************************
Public Function GetVisibleRangeToArray2DEx(ByVal rngAddress As String) As Array2DEx
    Set GetVisibleRangeToArray2DEx = XlWorkSheetUtils.GetVisibleRangeToArray2DEx(rngAddress, mOrigin.Name, mOrigin.Parent)
End Function

'******************************************************************************
'* [概  要] InjectDataToVisibleRange
'* [詳  細] 配列データを、指定したRange範囲の表示領域に出力する。
'*
'* @param vArray     Variant配列データ
'* @param startRowNo データ開始行番号
'* @param startColNo データ開始列番号
'* @param rowCount   行項目数
'* @param colCount   列項目数
'*
'******************************************************************************
Public Sub InjectDataToVisibleRange(vArray, startRowNo As Long, startColNo As Long, Optional rowCount As Long, Optional colCount As Long, Optional ByVal sheetName As String, Optional ByRef wb As Workbook)
    Call XlWorkSheetUtils.InjectDataToVisibleRange(vArray, startRowNo, startColNo, rowCount, colCount, mOrigin.Name, mOrigin.Parent)
End Sub

'******************************************************************************
'* [概  要] HasHiddenCells
'* [詳  細] 非表示のセル（行・列）が存在するかチェックする。
'*
'* @param sheetName  対象のシート名（省略すると選択中のシートで処理）
'* @param wb         ワークブックオブジェクト
'* @return Boolean 存在する場合、True
'******************************************************************************
Public Function HasHiddenCells() As Boolean
    HasHiddenCells = XlWorkSheetUtils.HasHiddenCells(mOrigin.Name, mOrigin.Parent)
End Function

'******************************************************************************
'* [概  要] DeleteRow
'* [詳  細] 指定した行番号の行を削除する。
'*
'* @param rowNo    削除する行の行番号
'* @param endRowNo 終了行番号（範囲指定時）
'******************************************************************************
Public Sub DeleteRow(rowNo As Long, Optional endRowNo As Long)
    Call EntireRow(rowNo, endRowNo).Delete
End Sub

'******************************************************************************
'* [概  要] DeleteColumn
'* [詳  細] 指定した列番号の列を削除する。
'*
'* @param colNo    削除する列の列番号
'* @param endColNo 終了列番号（範囲指定時）
'******************************************************************************
Public Sub DeleteColumn(colNo As Long, Optional endColNo As Long)
    Call EntireColumn(colNo, endColNo).Delete
End Sub

'******************************************************************************
'* [概  要] InsertRow
'* [詳  細] 指定した行番号の行を基準に空行を追加する。
'*
'* @param rowNo    追加する行位置の行番号
'* @param cntOfAdd 追加行数
'* @param cpOrigin 挿入行の形式コピー元の方向（上／下）を指定
'******************************************************************************
Public Sub InsertRow(rowNo As Long, Optional cntOfAdd As Long = 1, Optional cpOrigin As XlInsertFormatOrigin = xlFormatFromLeftOrAbove)
    Call EntireRow(rowNo, rowNo + cntOfAdd - 1).Insert(, cpOrigin)
End Sub

'******************************************************************************
'* [概  要] InsertColumn
'* [詳  細] 指定した列番号の列を基準に空列を追加する。
'*
'* @param colNo    追加する列位置の列番号
'* @param cntOfAdd 追加列数
'* @param cpOrigin 挿入列の形式コピー元の方向（左／右）を指定
'******************************************************************************
Public Sub InsertColumn(colNo As Long, Optional cntOfAdd As Long = 1, Optional cpOrigin As XlInsertFormatOrigin = xlFormatFromLeftOrAbove)
    Call EntireColumn(colNo, colNo + cntOfAdd - 1).Insert(, cpOrigin)
End Sub

'******************************************************************************
'* [概  要] IsEmptyRange
'* [詳  細] Range範囲のデータが空か判定する。
'*
'* @param rngAddress  対象とするアドレス
'* @return Boolean 判定結果（"":True）
'*
'******************************************************************************
Public Function IsEmptyRange(ByVal rngAdress As String) As Boolean
    Dim rng As Range
    For Each rng In Me.Range(rngAdress)
        If rng.Value <> "" Then
            IsEmptyRange = False
            Exit Function
        End If
    Next
    IsEmptyRange = True
End Function

'******************************************************************************
'* [概  要] IsNumericRange
'* [詳  細] Range範囲のデータが数値か判定する。
'*
'* @param rngAddress  対象とするアドレス
'* @return Boolean 判定結果（数値:True）
'*
'******************************************************************************
Public Function IsNumericRange(ByVal rngAdress As String) As Boolean
    Dim rng As Range
    For Each rng In Me.Range(rngAdress)
        If rng.Value = "" Or Not IsNumeric(rng.Value) Then
            IsNumericRange = False
            Exit Function
        End If
    Next
    IsNumericRange = True
End Function

'******************************************************************************
'* [概  要] IsStringRange
'* [詳  細] Range範囲のデータが数値か判定する。
'*
'* @param rngAddress  対象とするアドレス
'* @return Boolean 判定結果（数値:True）
'*
'******************************************************************************
Public Function IsStringRange(ByVal rngAdress As String) As Boolean
    Dim rng As Range
    For Each rng In Me.Range(rngAdress)
        If TypeName(rng.Value) <> "String" Then
            IsStringRange = False
            Exit Function
        End If
    Next
    IsStringRange = True
End Function

'******************************************************************************
'* [概  要] IsDateRange
'* [詳  細] Range範囲のデータが日付か判定する。
'*
'* @param rngAddress  対象とするアドレス
'* @return Boolean 判定結果（日付:True）
'*
'******************************************************************************
Public Function IsDateRange(ByVal rngAdress As String) As Boolean
    Dim rng As Range
    For Each rng In Me.Range(rngAdress)
        If Not IsDate(rng.Value) Then
            IsDateRange = False
            Exit Function
        End If
    Next
    IsDateRange = True
End Function

'******************************************************************************
'* [概  要] IsErrorRange
'* [詳  細] Range範囲のデータがエラーか判定する。
'*
'* @param rngAddress  対象とするアドレス
'* @return Boolean 判定結果（エラー:True）
'*
'******************************************************************************
Public Function IsErrorRange(ByVal rngAdress As String) As Boolean
    Dim rng As Range
    For Each rng In Me.Range(rngAdress)
        If rng.HasFormula Then
            If Not IsError(rng.Value) Then
                IsErrorRange = False
                Exit Function
            End If
        Else
            IsErrorRange = False
            Exit Function
        End If
    Next
    IsErrorRange = True
End Function

'******************************************************************************
'* [概  要] ArrangeTableFormat
'* [詳  細] 指定したアドレスのテーブル（セル範囲）を整形する。
'*
'* @param rngAddress  対象とするアドレス
'* @param useHeader   ヘッダ行があるか
'* @param numFormat   数値のフォーマット形式
'* @param dateFormat  日付のフォーマット形式
'******************************************************************************
Public Sub ArrangeTableFormat(ByVal rngAddress As String, _
                              Optional ByVal useHeader As Boolean = False, _
                              Optional ByVal numFormat As String = "#,###", _
                              Optional ByVal dateFormat As String = "yyyy/mm/dd")
    Call XlWorkSheetUtils.ArrangeTableFormat(rngAddress, mOrigin.Name, useHeader, numFormat, dateFormat, mOrigin.Parent)
End Sub

