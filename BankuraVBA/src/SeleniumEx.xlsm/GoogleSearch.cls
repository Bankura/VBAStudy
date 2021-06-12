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
'* [�@�\��] Selenium�g�pGoogle�����N���X
'* [��  ��] Google�������s���N���X�B
'*
'* @author Bankura
'* Copyright (c) 2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* WindowsAPI��`
'******************************************************************************
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'******************************************************************************
'* Enum��`
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
'* �萔��`
'******************************************************************************
Private Const URL_GOOGLE_SEARCH As String = "https://www.google.com/search"

'******************************************************************************
'* �ϐ���`
'******************************************************************************
Private mDriver As SeleniumExDriver
Private mGoogleSearchCountOfPage As Long
Private mGoogleSearchType As GSearchTypeEnum
Private mUseFilter As Boolean
Private mMaxSearchCount As Long
Private mSearchWord As String

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************

'*-----------------------------------------------------------------------------
'* WebDriver �v���p�e�B
'*
'* [��  ��]
'* �ESeleniumExDriver �I�u�W�F�N�g�B
'*-----------------------------------------------------------------------------
Public Property Get WebDriver() As SeleniumExDriver
    Set WebDriver = mDriver
End Property

'*-----------------------------------------------------------------------------
'* GoogleSearchCountOfPage �v���p�e�B
'*
'* [��  ��]
'* �EGoogle��1�y�[�W������̌������ʕ\�������B
'*-----------------------------------------------------------------------------
Public Property Get GoogleSearchCountOfPage() As Long
    GoogleSearchCountOfPage = mGoogleSearchCountOfPage
End Property
Public Property Let GoogleSearchCountOfPage(arg As Long)
    mGoogleSearchCountOfPage = arg
End Property

'*-----------------------------------------------------------------------------
'* GoogleSearchType �v���p�e�B
'*
'* [��  ��]
'* �EGoogle�����̎�ށB
'*-----------------------------------------------------------------------------
Public Property Get GoogleSearchType() As GSearchTypeEnum
    GoogleSearchType = mGoogleSearchType
End Property
Public Property Let GoogleSearchType(arg As GSearchTypeEnum)
    mGoogleSearchType = arg
End Property

'*-----------------------------------------------------------------------------
'* UseFilter �v���p�e�B
'*
'* [��  ��]
'* �E�������ɗގ��������ʂ������t�B���^���g�p���邩���w��B
'*-----------------------------------------------------------------------------
Public Property Get UseFilter() As Boolean
    UseFilter = mUseFilter
End Property
Public Property Let UseFilter(arg As Boolean)
    mUseFilter = arg
End Property

'*-----------------------------------------------------------------------------
'* MaxSearchCount �v���p�e�B
'*
'* [��  ��]
'* �E�������ʂ�������擾����ۂ̍ő匏���B
'*-----------------------------------------------------------------------------
Public Property Get MaxSearchCount() As Long
    MaxSearchCount = mMaxSearchCount
End Property
Public Property Let MaxSearchCount(arg As Long)
    mMaxSearchCount = arg
End Property

'*-----------------------------------------------------------------------------
'* SearchWord �v���p�e�B
'*
'* [��  ��]
'* �E�������郏�[�h�B
'*-----------------------------------------------------------------------------
Public Property Get SearchWord() As String
    SearchWord = mSearchWord
End Property
Public Property Let SearchWord(arg As String)
    mSearchWord = arg
End Property


'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
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
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] GetGSearchTypeString ���\�b�h
'* [��  ��] �����^�C�v�ɑΉ����錟���^�C�v��������擾����B
'*
'* @param gsType  �����^�C�v
'* @return String �����^�C�v������
'******************************************************************************
Private Function GetGSearchTypeString(gsType As GSearchTypeEnum) As String
    Select Case gsType
        Case GSearchTypeEnum.GSEARCH_NORMAL   ' �ʏ팟��
            GetGSearchTypeString = ""
        Case GSearchTypeEnum.GSEARCH_IMAGE    ' �摜����
            GetGSearchTypeString = "isch"
        Case GSearchTypeEnum.GSEARCH_VIDEO    ' ���挟��
            GetGSearchTypeString = "vid"
        Case GSearchTypeEnum.GSEARCH_NEWS     ' �j���[�X����
            GetGSearchTypeString = "nws"
        Case GSearchTypeEnum.GSEARCH_SHOPPING ' �V���b�s���O
            GetGSearchTypeString = "shop"
        Case GSearchTypeEnum.GSEARCH_BOOK     ' ���Ќ���
            GetGSearchTypeString = "bks"
        Case GSearchTypeEnum.GSEARCH_PATENT   ' ��������
            GetGSearchTypeString = "pts"
    End Select
End Function

'******************************************************************************
'* [�T  �v] Go ���\�b�h
'* [��  ��] Web�u���E�U�iChrome�j���N������B
'*
'******************************************************************************
Public Sub Go()
    mDriver.StartChrome
End Sub

'******************************************************************************
'* [�T  �v] SearchOnly ���\�b�h
'* [��  ��] �����̂ݍs���B
'*
'* @param pageNumber �擾�y�[�W�ԍ�
'******************************************************************************
Public Sub SearchOnly(Optional pageNumber As Long = 0)
    mDriver.GetPage GetSearchUrl(pageNumber)
End Sub

'******************************************************************************
'* [�T  �v] SearchAndShow ���\�b�h
'* [��  ��] Google�������s���A�������ʂ̃����N���C�~�f�B�G�C�g�E�B���h�E��
'*          �\������B
'*
'******************************************************************************
Public Sub SearchAndShow()
    Dim navPageNum As Long: navPageNum = 0
    Dim pageCnt As Long:    pageCnt = 0

    Do
        ' Web�y�[�W�\��
        Call mDriver.GetPage(GetSearchUrl(navPageNum))
        
        ' �������ʃ����N�v�f�擾
        Dim elems As Selenium.WebElements
        Set elems = GetSearchResultElements()

        ' �������ʂȂ��̏ꍇ�͏I��
        If elems.count = 0 Then
            Exit Do
        End If
        
        ' �������ʕ\��
        Call ShowPageLinkInfo(elems)
        
        ' �摜�����E�V���b�s���O�̏ꍇ�̓y�[�W�l�[�V�����Ȃ��̂��ߏI��
        If mGoogleSearchType = GSEARCH_IMAGE Or mGoogleSearchType = GSEARCH_SHOPPING Then
            Exit Do
        End If

        ' �����y�[�W�͈̓C���N�������g
        pageCnt = pageCnt + 1
        navPageNum = mGoogleSearchCountOfPage * pageCnt
        
        ' �w�肵���ő匏���ɒB�����ꍇ�͌����I��
        If mMaxSearchCount > 0 And mMaxSearchCount <= navPageNum Then
            Exit Do
        End If
    Loop
End Sub

'******************************************************************************
'* [�T  �v] GetSearchUrl ���\�b�h
'* [��  ��] �����^�C�v�ɑΉ����錟��URL���擾����B
'*
'* @param  startPageNumber �擾�y�[�W�ԍ�
'* @return String          ����URL
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
'* [�T  �v] GetSearchResultElements ���\�b�h
'* [��  ��] �����^�C�v�ɑΉ����錟�����ʃy�[�W�����擾����B
'*
'* @return Selenium.WebElements �������ʃy�[�W���
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
'* [�T  �v] ShowPageLinkInfo ���\�b�h
'* [��  ��] �������ʂ̃y�[�W�����C�~�f�B�G�C�g�E�B���h�E�ɕ\������B
'*
'* @param Selenium.WebElements �������ʃy�[�W���
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
'* [�T  �v] URL�G���R�[�h�����B
'* [��  ��] URL�G���R�[�h���s���B
'*
'* @param target �ΏۂƂȂ镶����
'* @return �G���R�[�h�㕶����
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
'* [�T  �v] GetBytes
'* [��  ��] ��������o�C�g�z��ɕϊ��
'*
'* @param str ������
'* @param encode �G���R�[�h�̃L�����Z�b�g
'* @return Byte() �o�C�g�z��
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