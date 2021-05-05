VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GoogleSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Selenium使用Google検索クラス
'* [詳  細] Google検索を行うクラス。
'*
'* @author Bankura
'* Copyright (c) 2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'******************************************************************************
'* Enum定義
'******************************************************************************
Public Enum GSearchTypeEnum
    GSEARCH_NORMAL
    GSEARCH_IMAGE
    GSEARCH_VIDEO
    GSEARCH_NEWS
    GSEARCH_SHOPPING
    GSEARCH_BOOK
    GSEARCH_PATENT
End Enum

'******************************************************************************
'* 定数定義
'******************************************************************************
Private Const URL_GOOGLE_SEARCH As String = "https://www.google.com/search"

'******************************************************************************
'* 変数定義
'******************************************************************************
Private mDriver As SeleniumExDriver
Private mGoogleSearchCountOfPage As Long
Private mGoogleSearchType As GSearchTypeEnum
Private mUseFilter As Boolean
Private mMaxSearchCount As Long
Private mSearchWord As String

'******************************************************************************
'* プロパティ定義
'******************************************************************************

'*-----------------------------------------------------------------------------
'* WebDriver プロパティ
'*
'* [説  明]
'* ・SeleniumExDriver オブジェクト。
'*-----------------------------------------------------------------------------
Public Property Get WebDriver() As SeleniumExDriver
    Set WebDriver = mDriver
End Property

'*-----------------------------------------------------------------------------
'* GoogleSearchCountOfPage プロパティ
'*
'* [説  明]
'* ・Googleの1ページ当たりの検索結果表示件数。
'*-----------------------------------------------------------------------------
Public Property Get GoogleSearchCountOfPage() As Long
    GoogleSearchCountOfPage = mGoogleSearchCountOfPage
End Property
Public Property Let GoogleSearchCountOfPage(arg As Long)
    mGoogleSearchCountOfPage = arg
End Property

'*-----------------------------------------------------------------------------
'* GoogleSearchType プロパティ
'*
'* [説  明]
'* ・Google検索の種類。
'*-----------------------------------------------------------------------------
Public Property Get GoogleSearchType() As GSearchTypeEnum
    GoogleSearchType = mGoogleSearchType
End Property
Public Property Let GoogleSearchType(arg As GSearchTypeEnum)
    mGoogleSearchType = arg
End Property

'*-----------------------------------------------------------------------------
'* UseFilter プロパティ
'*
'* [説  明]
'* ・検索時に類似検索結果を除くフィルタを使用するかを指定。
'*-----------------------------------------------------------------------------
Public Property Get UseFilter() As Boolean
    UseFilter = mUseFilter
End Property
Public Property Let UseFilter(arg As Boolean)
    mUseFilter = arg
End Property

'*-----------------------------------------------------------------------------
'* MaxSearchCount プロパティ
'*
'* [説  明]
'* ・検索結果から情報を取得する際の最大件数。
'*-----------------------------------------------------------------------------
Public Property Get MaxSearchCount() As Long
    MaxSearchCount = mMaxSearchCount
End Property
Public Property Let MaxSearchCount(arg As Long)
    mMaxSearchCount = arg
End Property

'*-----------------------------------------------------------------------------
'* SearchWord プロパティ
'*
'* [説  明]
'* ・検索するワード。
'*-----------------------------------------------------------------------------
Public Property Get SearchWord() As String
    SearchWord = mSearchWord
End Property
Public Property Let SearchWord(arg As String)
    mSearchWord = arg
End Property


'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mDriver = New SeleniumExDriver
    mDriver.SetArgumentMaximized
    mDriver.SetTimeoutsImplicitWait 10000
    mGoogleSearchCountOfPage = 10
    mGoogleSearchType = GSEARCH_NORMAL
End Sub
Private Sub Class_Terminate()
    mDriver.CloseWindow
    mDriver.Quit
    Set mDriver = Nothing
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] GetGSearchTypeString メソッド
'* [詳  細] 検索タイプに対応する検索タイプ文字列を取得する。
'*
'* @param gsType  検索タイプ
'* @return String 検索タイプ文字列
'******************************************************************************
Private Function GetGSearchTypeString(gsType As GSearchTypeEnum) As String
    Select Case gsType
        Case GSearchTypeEnum.GSEARCH_NORMAL   ' 通常検索
            GetGSearchTypeString = ""
        Case GSearchTypeEnum.GSEARCH_IMAGE    ' 画像検索
            GetGSearchTypeString = "isch"
        Case GSearchTypeEnum.GSEARCH_VIDEO    ' 動画検索
            GetGSearchTypeString = "vid"
        Case GSearchTypeEnum.GSEARCH_NEWS     ' ニュース検索
            GetGSearchTypeString = "nws"
        Case GSearchTypeEnum.GSEARCH_SHOPPING ' ショッピング
            GetGSearchTypeString = "shop"
        Case GSearchTypeEnum.GSEARCH_BOOK     ' 書籍検索
            GetGSearchTypeString = "bks"
        Case GSearchTypeEnum.GSEARCH_PATENT   ' 特許検索
            GetGSearchTypeString = "pts"
    End Select
End Function

'******************************************************************************
'* [概  要] Go メソッド
'* [詳  細] Webブラウザ（Chrome）を起動する。
'*
'******************************************************************************
Public Sub Go()
    mDriver.StartChrome
End Sub

'******************************************************************************
'* [概  要] SearchOnly メソッド
'* [詳  細] 検索のみ行う。
'*
'* @param pageNumber 取得ページ番号
'******************************************************************************
Public Sub SearchOnly(Optional pageNumber As Long = 0)
    mDriver.GetPage GetSearchUrl(pageNumber)
End Sub

'******************************************************************************
'* [概  要] SearchAndShow メソッド
'* [詳  細] Google検索を行い、検索結果のリンクをイミディエイトウィンドウに
'*          表示する。
'*
'******************************************************************************
Public Sub SearchAndShow()
    Dim navPageNum As Long: navPageNum = 0
    Dim pageCnt As Long:    pageCnt = 0

    Do
        ' Webページ表示
        Call mDriver.GetPage(GetSearchUrl(navPageNum))
        
        ' 検索結果リンク要素取得
        Dim elems As Selenium.WebElements
        Set elems = GetSearchResultElements()

        ' 検索結果なしの場合は終了
        If elems.count = 0 Then
            Exit Do
        End If
        
        ' 検索結果表示
        Call ShowPageLinkInfo(elems)
        
        ' 画像検索・ショッピングの場合はページネーションなしのため終了
        If mGoogleSearchType = GSEARCH_IMAGE Or mGoogleSearchType = GSEARCH_SHOPPING Then
            Exit Do
        End If

        ' 検索ページ範囲インクリメント
        pageCnt = pageCnt + 1
        navPageNum = mGoogleSearchCountOfPage * pageCnt
        
        ' 指定した最大件数に達した場合は検索終了
        If mMaxSearchCount > 0 And mMaxSearchCount <= navPageNum Then
            Exit Do
        End If
    Loop
End Sub

'******************************************************************************
'* [概  要] GetSearchUrl メソッド
'* [詳  細] 検索タイプに対応する検索URLを取得する。
'*
'* @param  startPageNumber 取得ページ番号
'* @return String          検索URL
'******************************************************************************
Private Function GetSearchUrl(startPageNumber As Long) As String
    Dim tbm As String: tbm = GetGSearchTypeString(mGoogleSearchType)
    Select Case mGoogleSearchType
        Case GSearchTypeEnum.GSEARCH_IMAGE, GSearchTypeEnum.GSEARCH_SHOPPING
            GetSearchUrl = URL_GOOGLE_SEARCH & "?q=" & EncodeUrl(mSearchWord) & "&tbm=" & tbm
        Case Else
            GetSearchUrl = URL_GOOGLE_SEARCH & "?q=" & EncodeUrl(mSearchWord) & "&start=" & startPageNumber & IIf(tbm = "", "", "&tbm=" & tbm) & IIf(mUseFilter, "", "&filter=0")
    End Select
End Function

'******************************************************************************
'* [概  要] GetSearchResultElements メソッド
'* [詳  細] 検索タイプに対応する検索結果ページ情報を取得する。
'*
'* @return Selenium.WebElements 検索結果ページ情報
'******************************************************************************
Private Function GetSearchResultElements() As Selenium.WebElements
    Select Case mGoogleSearchType
        Case GSearchTypeEnum.GSEARCH_IMAGE
            Set GetSearchResultElements = mDriver.FindElementById("islrg").FindElementsByXPath("//div[@class='isv-r PNCib MSM1fd BUooTd']")
        Case GSearchTypeEnum.GSEARCH_NEWS
            Set GetSearchResultElements = mDriver.FindElementById("search").FindElementsByXPath("//div[@class='dbsr']")
        Case GSearchTypeEnum.GSEARCH_BOOK
            Set GetSearchResultElements = mDriver.FindElementById("search").FindElementsByXPath("//div[@class='bHexk Tz5Hvf']")
        Case Else
            Set GetSearchResultElements = mDriver.FindElementsByXPath("//div[@class='yuRUbf']/a")
    End Select
End Function

'******************************************************************************
'* [概  要] ShowPageLinkInfo メソッド
'* [詳  細] 検索結果のページ情報をイミディエイトウィンドウに表示する。
'*
'* @param Selenium.WebElements 検索結果ページ情報
'******************************************************************************
Private Sub ShowPageLinkInfo(elems As Selenium.WebElements)
    Dim elem As WebElement, elem2 As WebElement, elems2 As WebElements
    
    Select Case mGoogleSearchType
        Case GSearchTypeEnum.GSEARCH_BOOK
            For Each elem In elems
                Set elem2 = elem.FindElementByTag("a")
                Set elems2 = elem.FindElementByClass("N96wpd").FindElementsByTag("span")
                Debug.Print elem2.FindElementByTag("h3").text, elems2.Item(1).text, elems2.Item(2).text, elem2.Attribute("href")
            Next
        Case GSearchTypeEnum.GSEARCH_NEWS
            For Each elem In elems
                Set elem2 = elem.FindElementByTag("a")
                Debug.Print elem.FindElementByClass("JheGif").text, elem2.Attribute("href")
            Next
        Case GSearchTypeEnum.GSEARCH_IMAGE
            For Each elem In elems
                Set elems2 = elem.FindElementsByTag("a")
                Debug.Print elems2.Item(1).FindElementByTag("img").Attribute("alt"), elems2.Item(2).Attribute("href")
            Next
        Case Else
            For Each elem In elems
                Debug.Print elem.FindElementByTag("h3").text, elem.Attribute("href")
            Next
    End Select
End Sub


'******************************************************************************
'* [概  要] URLエンコード処理。
'* [詳  細] URLエンコードを行う。
'*
'* @param target 対象となる文字列
'* @return エンコード後文字列
'******************************************************************************
Private Function EncodeUrl(target As String, Optional cSet As String = "UTF-8") As String
    Dim buf() As Byte, s As String, i As Long
    
    buf = GetBytes(target, cSet)
    
    For i = 0 To UBound(buf)
        Dim flg As Boolean: flg = False
        Select Case buf(i)
            Case 45, 46, 95, 126 '-._~
                flg = True
            Case 48 To 57  '0-9
                flg = True
            Case 65 To 90  'A-Z
                flg = True
            Case 97 To 122 'a-z
                flg = True
        End Select
        If flg Then
            s = s & Chr(buf(i))
        Else
            s = s & "%" & Hex(buf(i))
        End If
    Next
    EncodeUrl = s
End Function

'******************************************************************************
'* [概  要] GetBytes
'* [詳  細] 文字列をバイト配列に変換｡
'*
'* @param str 文字列
'* @param encode エンコードのキャラセット
'* @return Byte() バイト配列
'******************************************************************************
Private Function GetBytes(ByVal str As String, Optional ByVal encode As String) As Byte()
    If str = "" Then
       GetBytes = VBA.vbNullString
       Exit Function
    End If
    If encode = "" Then
       GetBytes = str
       Exit Function
    End If
    
    With CreateObject("ADODB.Stream")
        .Charset = encode
        .LineSeparator = -1 'adCRLF
        .Type = 2 ' adTypeText
        .Mode = 3 ' adModeReadWrite
        .Open
        .WriteText str
        .Position = 0
        .Type = 1 ' adTypeBinary
        GetBytes = .Read
        .Close
    End With
End Function
