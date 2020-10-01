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
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mOrigin As Worksheet

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
Public Function Init(ByVal sheetName As String, Optional ByVal wb As Workbook) As WorksheetEx
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    Set mOrigin = wb.Worksheets(sheetName)
    Set Init = Me
End Function

''******************************************************************************
''* [概  要] Activate メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''******************************************************************************
Sub Activate()
    Call mOrigin.Activate
End Sub

''******************************************************************************
''* [概  要] Calculate メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''******************************************************************************
Sub Calculate()
    Call mOrigin.Calculate
End Sub

''******************************************************************************
''* [概  要] ChartObjects メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''* @param Index xxxx。省略可能。
''* @return Object
''*
''******************************************************************************
Function ChartObjects(Optional Index) As Object
    Set ChartObjects = mOrigin.ChartObjects(Index)
End Function

''******************************************************************************
''* [概  要] CheckSpelling メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''* @param CustomDictionary xxxxx。省略可能。
''* @param IgnoreUppercase xxxxx。省略可能。
''* @param AlwaysSuggest xxxxx。省略可能。
''* @param SpellLang xxxxx。省略可能。
''*
''******************************************************************************
Sub CheckSpelling(Optional CustomDictionary, Optional IgnoreUppercase, Optional AlwaysSuggest, Optional SpellLang)
    Call mOrigin.CheckSpelling(CustomDictionary, IgnoreUppercase, AlwaysSuggest, SpellLang)
End Sub

''******************************************************************************
''* [概  要] CircleInvalid メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''******************************************************************************
Sub CircleInvalid()
    Call mOrigin.CircleInvalid
End Sub

''******************************************************************************
''* [概  要] ClearArrows メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''******************************************************************************
Sub ClearArrows()
    Call mOrigin.ClearArrows
End Sub

''******************************************************************************
''* [概  要] ClearCircles メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''******************************************************************************
Sub ClearCircles()
    Call mOrigin.ClearCircles
End Sub

''******************************************************************************
''* [概  要] Copy メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''* @param Before xxxxx。省略可能。
''* @param After xxxxx。省略可能。
''*
''******************************************************************************
Sub Copy(Optional Before, Optional After)
    Call mOrigin.Copy(Before, After)
End Sub

''******************************************************************************
''* [概  要] Delete メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''******************************************************************************
Sub Delete()
    Call mOrigin.Delete
End Sub

''******************************************************************************
''* [概  要] Evaluate メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''* @param Name xxxx
''* @return Variant xxx
''*
''******************************************************************************
Function Evaluate(Name)
    Evaluate = mOrigin.Evaluate(Name)
End Function

''******************************************************************************
''* [概  要] ExportAsFixedFormat メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''* @param Type_ xxxx
''* @param fileName xxxxx。省略可能。
''* @param Quality xxxxx。省略可能。
''* @param IncludeDocProperties xxxxx。省略可能。
''* @param IgnorePrintAreas xxxxx。省略可能。
''* @param From xxxxx。省略可能。
''* @param To_ xxxxx。省略可能。
''* @param OpenAfterPublish xxxxx。省略可能。
''* @param FixedFormatExtClassPtr xxxxx。省略可能。
''* @param WorkIdentity xxxxx。省略可能。
''*
''******************************************************************************
Sub ExportAsFixedFormat(Type_ As XlFixedFormatType, Optional fileName, Optional Quality, Optional IncludeDocProperties, Optional IgnorePrintAreas, Optional From, Optional To_, Optional OpenAfterPublish, Optional FixedFormatExtClassPtr, Optional WorkIdentity)
    Call mOrigin.ExportAsFixedFormat(Type_, fileName, Quality, IncludeDocProperties, IgnorePrintAreas, From, To_, OpenAfterPublish, FixedFormatExtClassPtr, WorkIdentity)
End Sub

''******************************************************************************
''* [概  要] Move メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''* @param Before xxxx。省略可能。
''* @param After  xxxx。省略可能。
''*
''******************************************************************************
Sub Move(Optional Before, Optional After)
    Call mOrigin.Move(Before, After)
End Sub
''******************************************************************************
''* [概  要] OLEObjects メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''* @param Index xxxx。省略可能。
''* @return Object xxxx
''*
''******************************************************************************
Function OLEObjects(Optional Index) As Object
    Set OLEObjects = mOrigin.OLEObjects(Index)
End Function
    
''******************************************************************************
''* [概  要] Paste メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''* @param Destination xxxx。省略可能。
''* @param Link xxxx。省略可能。
''*
''******************************************************************************
Sub Paste(Optional Destination, Optional Link)
    Call mOrigin.Paste(Destination, Link)
End Sub

''******************************************************************************
''* [概  要] PasteSpecial メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''* @param Format xxxx。省略可能。
''* @param Link xxxx。省略可能。
''* @param DisplayAsIcon xxxx。省略可能。
''* @param IconFileName xxxx。省略可能。
''* @param IconIndex xxxx。省略可能。
''* @param IconLabel xxxx。省略可能。
''* @param NoHTMLFormatting xxxx。省略可能。
''*
''******************************************************************************
Sub PasteSpecial(Optional Format, Optional Link, Optional DisplayAsIcon, Optional IconFileName, Optional IconIndex, Optional IconLabel, Optional NoHTMLFormatting)
    Call mOrigin.PasteSpecial(Format, Link, DisplayAsIcon, IconFileName, IconIndex, IconLabel, NoHTMLFormatting)
End Sub

''******************************************************************************
''* [概  要] PivotTables メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''* @param Index xxxx。省略可能。
''* @return Object xxxx
''*
''******************************************************************************
Function PivotTables(Optional Index) As Object
    Set PivotTables = mOrigin.PivotTables(Index)
End Function

''******************************************************************************
''* [概  要] PrintOut メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''* @param From xxxx。省略可能。
''* @param To_ xxxx。省略可能。
''* @param Copies xxxx。省略可能。
''* @param Preview xxxx。省略可能。
''* @param ActivePrinter xxxx。省略可能。
''* @param PrintToFile xxxx。省略可能。
''* @param Collate xxxx。省略可能。
''* @param PrToFileName xxxx。省略可能。
''* @param IgnorePrintAreas xxxx。省略可能。
''*
''******************************************************************************
Sub PrintOut(Optional From, Optional To_, Optional Copies, Optional Preview, Optional ActivePrinter, Optional PrintToFile, Optional Collate, Optional PrToFileName, Optional IgnorePrintAreas)
    Call mOrigin.PrintOut(From, To_, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName, IgnorePrintAreas)
End Sub

''******************************************************************************
''* [概  要] PrintPreview メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''* @param EnableChanges xxxx。省略可能。
''*
''******************************************************************************
Sub PrintPreview(Optional EnableChanges)
    Call mOrigin.PrintPreview(EnableChanges)
End Sub

''******************************************************************************
''* [概  要] Protect メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''* @param Password xxxx。省略可能。
''* @param DrawingObjects xxxx。省略可能。
''* @param Contents xxxx。省略可能。
''* @param Scenarios xxxx。省略可能。
''* @param UserInterfaceOnly xxxx。省略可能。
''* @param AllowFormattingCells xxxx。省略可能。
''* @param AllowFormattingColumns xxxx。省略可能。
''* @param AllowFormattingRows xxxx。省略可能。
''* @param AllowInsertingColumns xxxx。省略可能。
''* @param AllowInsertingRows xxxx。省略可能。
''* @param AllowInsertingHyperlinks xxxx。省略可能。
''* @param AllowDeletingColumns xxxx。省略可能。
''* @param AllowDeletingRows xxxx。省略可能。
''* @param AllowSorting xxxx。省略可能。
''* @param AllowFiltering xxxx。省略可能。
''* @param AllowUsingPivotTables xxxx。省略可能。
''*
''******************************************************************************
Sub Protect(Optional Password, Optional DrawingObjects, Optional Contents, Optional Scenarios, Optional UserInterfaceOnly, Optional AllowFormattingCells, Optional AllowFormattingColumns, Optional AllowFormattingRows, Optional AllowInsertingColumns, Optional AllowInsertingRows, Optional AllowInsertingHyperlinks, Optional AllowDeletingColumns, Optional AllowDeletingRows, Optional AllowSorting, Optional AllowFiltering, Optional AllowUsingPivotTables)
    Call mOrigin.Protect(Password, DrawingObjects, Contents, Scenarios, UserInterfaceOnly, AllowFormattingCells, AllowFormattingColumns, AllowFormattingRows, AllowInsertingColumns, AllowInsertingRows, AllowInsertingHyperlinks, AllowDeletingColumns, AllowDeletingRows, AllowSorting, AllowFiltering, AllowUsingPivotTables)
End Sub

''******************************************************************************
''* [概  要] ResetAllPageBreaks メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''******************************************************************************
Sub ResetAllPageBreaks()
    Call mOrigin.ResetAllPageBreaks
End Sub

''******************************************************************************
''* [概  要] SaveAs メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''* @param fileName xxxx。省略可能。
''* @param FileFormat xxxx。省略可能。
''* @param Password xxxx。省略可能。
''* @param WriteResPassword xxxx。省略可能。
''* @param ReadOnlyRecommended xxxx。省略可能。
''* @param CreateBackup xxxx。省略可能。
''* @param AddToMru xxxx。省略可能。
''* @param TextCodepage xxxx。省略可能。
''* @param TextVisualLayout xxxx。省略可能。
''* @param Local_ xxxx。省略可能。
''*
''******************************************************************************
Sub SaveAs(fileName As String, Optional FileFormat, Optional Password, Optional WriteResPassword, Optional ReadOnlyRecommended, Optional CreateBackup, Optional AddToMru, Optional TextCodepage, Optional TextVisualLayout, Optional Local_)
    Call mOrigin.SaveAs(fileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AddToMru, TextCodepage, TextVisualLayout, Local_)
End Sub

''******************************************************************************
''* [概  要] Scenarios メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''* @param Index xxxx。省略可能。
''* @return Object xxxx
''*
''******************************************************************************
Function Scenarios(Optional Index) As Object
    Set Scenarios = mOrigin.Scenarios(Index)
End Function

''******************************************************************************
''* [概  要] Select_ メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''* @param rep xxxx。省略可能。
''*
''******************************************************************************
Sub Select_(Optional rep)
    Call mOrigin.Select(rep)
End Sub

''******************************************************************************
''* [概  要] SetBackgroundPicture メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''* @param fileName xxxxx
''*
''******************************************************************************
Sub SetBackgroundPicture(fileName As String)
    Call mOrigin.SetBackgroundPicture(fileName)
End Sub

''******************************************************************************
''* [概  要] ShowAllData メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''******************************************************************************
Sub ShowAllData()
    Call mOrigin.ShowAllData
End Sub

''******************************************************************************
''* [概  要] ShowDataForm メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''******************************************************************************
Sub ShowDataForm()
    Call mOrigin.ShowDataForm
End Sub

''******************************************************************************
''* [概  要] Unprotect メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''* @param Password xxxx。省略可能。
''*
''******************************************************************************
Sub Unprotect(Optional Password)
    Call mOrigin.Unprotect(Password)
End Sub

''******************************************************************************
''* [概  要] XmlDataQuery メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''* @param XPath xxxx
''* @param SelectionNamespaces ファイルパス。省略可能。
''* @param Map ファイルパス。省略可能。
''* @return Range xxxx
''*
''******************************************************************************
Function XmlDataQuery(XPath As String, Optional SelectionNamespaces, Optional Map) As Range
    Set XmlDataQuery = mOrigin.XmlDataQuery(XPath, SelectionNamespaces, Map)
End Function

''******************************************************************************
''* [概  要] XmlMapQuery メソッド
''* [詳  細] ファイルが存在するかどうか判定します｡
''*
''* @param XPath xxxx
''* @param SelectionNamespaces ファイルパス。省略可能。
''* @param Map ファイルパス。省略可能。
''* @return Range xxxx
''*
''******************************************************************************
Function XmlMapQuery(XPath As String, Optional SelectionNamespaces, Optional Map) As Range
    Set XmlMapQuery = mOrigin.XmlMapQuery(XPath, SelectionNamespaces, Map)
End Function



