VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeleniumExDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Selenium使用ブラウザアクセスクラス
'* [詳  細] SeleniumBasicを使用したブラウザ制御クラス。
'*          Selenium.WebDriverクラスをラップしたメソッドのほか、
'*          各種ユーティリティ機能を提供する。
'*
'* [参  考] https://excel-ubara.com/excelvba4/EXCEL_VBA_401.html
'*          https://www.tagindex.com/html5/elements/
'*          https://html.web-tora.com/basic-global-attribute.html
'*
'* [依  存] <SeleniumBasic>
'*           https://florentbr.github.io/SeleniumBasic/
'*           https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0
'*
'*          <Chromeドライバ>
'*           https://chromedriver.chromium.org/downloads
'*          <Edgeドライバ>
'*           https://developer.microsoft.com/ja-jp/microsoft-edge/tools/webdriver/
'*          <Firefoxドライバ>
'*           https://github.com/mozilla/geckodriver/releases
'*          <InternetExplorerドライバ>
'*           https://www.selenium.dev/downloads/
'*
'* [備  考]
'*         ・「参照設定」で「Selenium Type Library」を設定する。
'*         ・SeleniumBasicを動かすには、「.NET Framework 3.5」を有効化する必要
'*           がある。
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
Public Enum HKeyEnum
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
End Enum

Public Enum SearchElemTypeEnum
    SEARCH_CLASS
    SEARCH_CSS
    SEARCH_ID
    SEARCH_LINKTEXT
    SEARCH_NAME
    SEARCH_PARTIALLINKTEXT
    SEARCH_TAG
    SEARCH_XPATH
    SEARCH_TEXT
    SEARCH_ATTRIBUTE
End Enum

'******************************************************************************
'* 構造体定義
'******************************************************************************

'******************************************************************************
'* 定数定義
'******************************************************************************
Private Const URL_CHROME_DRIVER_LATEST_RELEASE_VERSION As String = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_"
Private Const URL_CHROME_DRIVER_ZIP As String = "https://chromedriver.storage.googleapis.com/{driver_version}/chromedriver_win32.zip"
Private Const JS_GET_TIMINGS As String = "var t=performance.timing; " & _
                                         "return [ " & _
                                                   "t.responseEnd - t.navigationStart, " & _
                                                   "t.loadEventEnd - t.responseEnd " & _
                                                "];  "
Private Const JS_LIST_ATTRIBUTES As String = "var d={}, a=this.attributes;  " & _
                                             "for(var i=0; i<a.length; i++) " & _
                                             "  d[a[i].name]=a[i].value;    " & _
                                             "return d;"
                                             
'******************************************************************************
'* 変数定義
'******************************************************************************
Private mApplication As Selenium.Application
Private mDriver As Selenium.WebDriver
Private mKeys As Selenium.Keys
Private mFso As Object

'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* Application プロパティ
'*
'* [説  明]
'* ・Selenium.Application オブジェクトを取得する。
'*-----------------------------------------------------------------------------
Public Property Get Application() As Selenium.Application
    Set Application = mApplication
End Property

'*-----------------------------------------------------------------------------
'* WebDriver プロパティ
'*
'* [説  明]
'* ・Selenium.WebDriver オブジェクトを取得する。
'*   Selenium.Application のメンバー
'*-----------------------------------------------------------------------------
Public Property Get WebDriver() As Selenium.WebDriver
    Set WebDriver = mDriver
End Property

'*-----------------------------------------------------------------------------
'* ChromeDriver プロパティ
'*
'* [説  明]
'* ・Selenium.WebDriver オブジェクトをブラウザにchromeを指定して取得する。
'*   Selenium.Application のメンバー
'*-----------------------------------------------------------------------------
Public Property Get ChromeDriver() As Selenium.WebDriver
    Set mDriver = mApplication.ChromeDriver
    Set ChromeDriver = mDriver
End Property

'*-----------------------------------------------------------------------------
'* IEDriver プロパティ
'*
'* [説  明]
'* ・Selenium.WebDriver オブジェクトをブラウザにIEを指定して取得する。
'*   Selenium.Application のメンバー
'*-----------------------------------------------------------------------------
Public Property Get IEDriver() As Selenium.WebDriver
    Set mDriver = mApplication.IEDriver
    Set IEDriver = mDriver
End Property

'*-----------------------------------------------------------------------------
'* OperaDriver プロパティ
'*
'* [説  明]
'* ・Selenium.WebDriver オブジェクトをブラウザOperaを指定して取得する。
'*   Selenium.Application のメンバー
'*-----------------------------------------------------------------------------
Public Property Get OperaDriver() As Selenium.WebDriver
    Set mDriver = mApplication.OperaDriver
    Set OperaDriver = mDriver
End Property

'*-----------------------------------------------------------------------------
'* PhantomJSDriver プロパティ
'*
'* [説  明]
'* ・Selenium.WebDriver オブジェクトをブラウザPhantomJSを指定して取得する。
'*   Selenium.Application のメンバー
'*-----------------------------------------------------------------------------
Public Property Get PhantomJSDriver() As Selenium.WebDriver
    Set mDriver = mApplication.PhantomJSDriver
    Set PhantomJSDriver = mDriver
End Property

'*-----------------------------------------------------------------------------
'* Keys プロパティ
'*
'* [説  明]
'* ・Selenium.Keys オブジェクトを取得する。
'*   Selenium.Application のメンバー
'*-----------------------------------------------------------------------------
Public Property Get Keys() As Selenium.Keys
    Set Keys = mKeys
End Property

'*-----------------------------------------------------------------------------
'* Assert プロパティ
'*
'* [説  明]
'* ・Selenium.Keys オブジェクトを取得する。
'*   Selenium.Application のメンバー
'*-----------------------------------------------------------------------------
Public Property Get Assert() As Selenium.Assert
    Set Assert = mApplication.Assert
End Property

'*-----------------------------------------------------------------------------
'* By プロパティ
'*
'* [説  明]
'* ・Selenium.Keys オブジェクトを取得する。
'*   Selenium.Application のメンバー
'*-----------------------------------------------------------------------------
Public Property Get By() As Selenium.By
    Set By = mApplication.By
End Property

'*-----------------------------------------------------------------------------
'* Utils プロパティ
'*
'* [説  明]
'* ・Selenium.Keys オブジェクトを取得する。
'*   Selenium.Application のメンバー
'*-----------------------------------------------------------------------------
Public Property Get Utils() As Selenium.Utils
    Set Utils = mApplication.Utils
End Property

'*-----------------------------------------------------------------------------
'* Verify プロパティ
'*
'* [説  明]
'* ・Selenium.Keys オブジェクトを取得する。
'*   Selenium.Application のメンバー
'*-----------------------------------------------------------------------------
Public Property Get Verify() As Selenium.Verify
    Set Verify = mApplication.Verify
End Property

'*-----------------------------------------------------------------------------
'* Waiter プロパティ
'*
'* [説  明]
'* ・Selenium.Keys オブジェクトを取得する。
'*   Selenium.Application のメンバー
'*-----------------------------------------------------------------------------
Public Property Get Waiter() As Selenium.Waiter
    Set Waiter = mApplication.Waiter
End Property

'*-----------------------------------------------------------------------------
'* Actions プロパティ
'*
'* [説  明]
'* ・Selenium.Actions オブジェクトを取得する。
'*   Selenium.WebDriver のメンバー
'*-----------------------------------------------------------------------------
Public Property Get Actions() As Selenium.Actions
    Set Assert = mDriver.Assert
End Property

'*-----------------------------------------------------------------------------
'* Actions プロパティ
'*
'* [説  明]
'* ・Selenium.Actions オブジェクトを取得する。
'*   Selenium.WebDriver のメンバー
'*-----------------------------------------------------------------------------
Public Property Get baseUrl() As String
    baseUrl = mDriver.baseUrl
End Property

'*-----------------------------------------------------------------------------
'* Keyboard プロパティ
'*
'* [説  明]
'* ・Selenium.Keyboard オブジェクトを取得する。
'*   Selenium.WebDriver のメンバー
'*-----------------------------------------------------------------------------
Public Property Get Keyboard() As Selenium.Keyboard
    Set Keyboard = mDriver.Keyboard
End Property

'*-----------------------------------------------------------------------------
'* Manage プロパティ
'*
'* [説  明]
'* ・Selenium.Manage オブジェクトを取得する。
'*   Selenium.WebDriver のメンバー
'*-----------------------------------------------------------------------------
Public Property Get Manage() As Selenium.Manage
    Set Manage = mDriver.Manage
End Property

'*-----------------------------------------------------------------------------
'* Mouse プロパティ
'*
'* [説  明]
'* ・Selenium.Mouse オブジェクトを取得する。
'*   Selenium.WebDriver のメンバー
'*-----------------------------------------------------------------------------
Public Property Get Mouse() As Selenium.Mouse
    Set Mouse = mDriver.Mouse
End Property

'*-----------------------------------------------------------------------------
'* Proxy プロパティ
'*
'* [説  明]
'* ・Selenium.Proxy オブジェクトを取得する。
'*   Selenium.WebDriver のメンバー
'*-----------------------------------------------------------------------------
Public Property Get Proxy() As Selenium.Proxy
    Set Proxy = mDriver.Proxy
End Property

'*-----------------------------------------------------------------------------
'* Timeouts プロパティ
'*
'* [説  明]
'* ・Selenium.Timeouts オブジェクトを取得する。
'*   Selenium.WebDriver のメンバー
'*-----------------------------------------------------------------------------
Public Property Get Timeouts() As Selenium.Timeouts
    Set Timeouts = mDriver.Timeouts
End Property

'*-----------------------------------------------------------------------------
'* Title プロパティ
'*
'* [説  明]
'* ・Selenium.Title オブジェクトを取得する。
'*   Selenium.WebDriver のメンバー
'*-----------------------------------------------------------------------------
Public Property Get Title() As String
    Title = mDriver.Title
End Property

'*-----------------------------------------------------------------------------
'* TouchActions プロパティ
'*
'* [説  明]
'* ・Selenium.TouchActions オブジェクトを取得する。
'*   Selenium.WebDriver のメンバー
'*-----------------------------------------------------------------------------
Public Property Get TouchActions() As Selenium.TouchActions
    Set TouchActions = mDriver.TouchActions
End Property

'*-----------------------------------------------------------------------------
'* TouchScreen プロパティ
'*
'* [説  明]
'* ・Selenium.TouchScreen オブジェクトを取得する。
'*   Selenium.WebDriver のメンバー
'*-----------------------------------------------------------------------------
Public Property Get TouchScreen() As Selenium.TouchScreen
    Set TouchScreen = mDriver.TouchScreen
End Property

'*-----------------------------------------------------------------------------
'* url プロパティ
'*
'* [説  明]
'* ・Selenium.url オブジェクトを取得する。
'*   Selenium.WebDriver のメンバー
'*-----------------------------------------------------------------------------
Public Property Get url() As String
    url = mDriver.url
End Property

'*-----------------------------------------------------------------------------
'* Window プロパティ
'*
'* [説  明]
'* ・Selenium.Window オブジェクトを取得する。
'*   Selenium.WebDriver のメンバー
'*-----------------------------------------------------------------------------
Public Property Get Window() As Selenium.Window
    Set Window = mDriver.Window
End Property

'*-----------------------------------------------------------------------------
'* Windows プロパティ
'*
'* [説  明]
'* ・Selenium.Windows オブジェクトのSelenium.Listコレクションを取得する。
'*   Selenium.WebDriver のメンバー
'*-----------------------------------------------------------------------------
Public Property Get Windows() As Selenium.List
    Set Windows = mDriver.Windows
End Property

'*-----------------------------------------------------------------------------
'* Capabilities プロパティ
'*
'* [説  明]
'* ・Capabilitiesを Selenium.Dictionary オブジェクトとして取得する。
'*   Selenium.Manage のメンバー
'*-----------------------------------------------------------------------------
Public Property Get Capabilities() As Selenium.Dictionary
    Set Capabilities = mDriver.Manage.Capabilities
End Property

'*-----------------------------------------------------------------------------
'* Cookies プロパティ
'*
'* [説  明]
'* ・Selenium.Cookies オブジェクトを取得する。
'*   Selenium.Manage のメンバー
'*-----------------------------------------------------------------------------
Public Property Get Cookies() As Selenium.Cookies
    Set Cookies = mDriver.Manage.Cookies
End Property

'*-----------------------------------------------------------------------------
'* StorageLocal プロパティ
'*
'* [説  明]
'* ・StorageLocalを Selenium.Storage オブジェクトとして取得する。
'*   Selenium.Manage のメンバー
'*-----------------------------------------------------------------------------
Public Property Get StorageLocal() As Selenium.Storage
    Set StorageLocal = mDriver.Manage.StorageLocal
End Property

'*-----------------------------------------------------------------------------
'* StorageSession プロパティ
'*
'* [説  明]
'* ・StorageSessionを Selenium.Storage オブジェクトとして取得する。
'*   Selenium.Manage のメンバー
'*-----------------------------------------------------------------------------
Public Property Get StorageSession() As Selenium.Storage
    Set StorageSession = mDriver.Manage.StorageSession
End Property


'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mApplication = New Selenium.Application
    Set mDriver = mApplication.WebDriver
    Set mKeys = mApplication.Keys
    Set mFso = CreateObject("Scripting.FileSystemObject")
    
End Sub
Private Sub Class_Terminate()
    Set mApplication = Nothing
    Set mDriver = Nothing
    Set mKeys = Nothing
    Set mFso = Nothing
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Dictionary メソッド
'* [詳  細] Selenium.Dictionary オブジェクトを取得し返却する。
'*          Selenium.Application のメンバー
'*
'* @return Selenium.Dictionary オブジェクト
'******************************************************************************
Public Function Dictionary() As Selenium.Dictionary
    Set Dictionary = mApplication.Dictionary
End Function

'******************************************************************************
'* [概  要] List メソッド
'* [詳  細] Selenium.List オブジェクトを取得し返却する。
'*          Selenium.Application のメンバー
'*
'* @return Selenium.List オブジェクト
'******************************************************************************
Public Function List() As Selenium.List
    Set List = mApplication.List
End Function

'******************************************************************************
'* [概  要] PdfFile メソッド
'* [詳  細] Selenium.PdfFile オブジェクトを取得し返却する。
'*          Selenium.Application のメンバー
'*
'* @return Selenium.PdfFile オブジェクト
'******************************************************************************
Public Function PdfFile() As Selenium.PdfFile
    Set PdfFile = mApplication.PdfFile
End Function

'******************************************************************************
'* [概  要] Point メソッド
'* [詳  細] Selenium.Point オブジェクトを取得し返却する。
'*          Selenium.Application のメンバー
'*
'* @param x 横位置
'* @param y 縦位置
'* @return Selenium.Point オブジェクト
'******************************************************************************
Public Function Point(x As Long, y As Long) As Selenium.Point
    Set Point = mApplication.Point(x, y)
End Function

'******************************************************************************
'* [概  要] Size メソッド
'* [詳  細] Selenium.Size オブジェクトを取得し返却する。
'*          Selenium.Application のメンバー
'*
'* @param Width  幅
'* @param Height 高さ
'* @return Selenium.Size オブジェクト
'******************************************************************************
Public Function Size(Width As Long, Height As Long) As Selenium.Size
    Set Size = mApplication.Size(Width, Height)
End Function

'******************************************************************************
'* [概  要] Table メソッド
'* [詳  細] Selenium.Table オブジェクトを取得し返却する。
'*          Selenium.Application のメンバー
'*
'* @return Selenium.Table オブジェクト
'******************************************************************************
Public Function Table() As Selenium.Table
    Set Table = mApplication.Table
End Function


'******************************************************************************
'* [概  要] ActiveElement メソッド
'* [詳  細] Gets the currently active element on the document
'*          Selenium.WebDriver のメンバー
'*
'* @return Selenium.WebElement オブジェクト
'******************************************************************************
Public Function ActiveElement() As Selenium.WebElement
    Set ActiveElement = mDriver.ActiveElement
End Function

'******************************************************************************
'* [概  要] AddArgument メソッド
'* [詳  細] Selenium.Table オブジェクトを取得し返却する。
'*          Selenium.WebDriver のメンバー
'*
'* @param argument   xxx
'******************************************************************************
Public Sub AddArgument(argument As String)
    Call mDriver.AddArgument(argument)
End Sub

'******************************************************************************
'* [概  要] AddExtension メソッド
'* [詳  細] Selenium.Table オブジェクトを取得し返却する。
'*          Selenium.WebDriver のメンバー
'*
'* @param extensionPath   xxx
'******************************************************************************
Public Sub AddExtension(extensionPath As String)
    Call mDriver.AddExtension(extensionPath)
End Sub

'******************************************************************************
'* [概  要] CacheStatus メソッド
'* [詳  細] Get the status of the html5 application cache.
'*          Selenium.WebDriver のメンバー
'*
'* @return Selenium.CacheState オブジェクト
'******************************************************************************
Public Function CacheStatus() As Selenium.CacheState
    CacheStatus = mDriver.CacheStatus
End Function

'******************************************************************************
'* [概  要] CloseWindow メソッド（Close メソッド）
'* [詳  細] Closes the current window。
'*          Selenium.WebDriver のメンバー
'*
'******************************************************************************
Public Sub CloseWindow()
    Call mDriver.Close
End Sub
    
'******************************************************************************
'* [概  要] ExecuteAsyncScript メソッド
'* [詳  細] Execute an asynchronous piece of JavaScript in the current frame or window. Returns the first argument of the callback function.
'*          Selenium.WebDriver のメンバー
'*
'* @param script   xxx
'* @param arguments   xxx
'* @param timeout   xxx
'* @return Variant
'******************************************************************************
Public Function ExecuteAsyncScript(script As String, Optional arguments, Optional timeout As Long = -1)
    ExecuteAsyncScript = mDriver.ExecuteAsyncScript(script, arguments, timeout)
End Function
    
'******************************************************************************
'* [概  要] ExecuteScript メソッド
'* [詳  細] Execute a piece of JavaScript in the current frame or window. Returns the value specified by the return statement.
'*          Selenium.WebDriver のメンバー
'*
'* @param script   xxx
'* @param arguments   xxx
'* @return Variant
'******************************************************************************
Public Function ExecuteScript(script As String, Optional arguments)
    ExecuteScript = mDriver.ExecuteScript(script, arguments)
End Function
    
'******************************************************************************
'* [概  要] FindElement メソッド
'* [詳  細] Find the first WebElement using the given method.
'*          Selenium.WebDriver のメンバー
'*
'* @param By   xxx
'* @param timeout   xxx
'* @param raise   xxx
'* @return Selenium.WebElement オブジェクト
'******************************************************************************
Public Function FindElement(By As By, Optional timeout As Long = -1, Optional raise As Boolean = True) As Selenium.WebElement
    Set FindElement = mDriver.FindElement(By, timeout, raise)
End Function
    
'******************************************************************************
'* [概  要] FindElementBy メソッド
'* [詳  細] Find the first WebElement for a given strategy and value.
'*          Selenium.WebDriver のメンバー
'*
'* @param Strategy   xxx
'* @param value   xxx
'* @param timeout   xxx
'* @param raise   xxx
'* @return Selenium.WebElement オブジェクト
'******************************************************************************
Public Function FindElementBy(Strategy As Integer, value As String, Optional timeout As Long = -1, Optional raise As Boolean = True) As Selenium.WebElement
    Set FindElementBy = mDriver.FindElementBy(Strategy, value, timeout, raise)
End Function
    
'******************************************************************************
'* [概  要] FindElementByClass メソッド
'* [詳  細] Finds the first element matching the specified CSS class.
'*          Selenium.WebDriver のメンバー
'*
'* @param classname   xxx
'* @param timeout   xxx
'* @param raise   xxx
'* @return Selenium.WebElement オブジェクト
'******************************************************************************
Public Function FindElementByClass(classname As String, Optional timeout As Long = -1, Optional raise As Boolean = True) As Selenium.WebElement
    Set FindElementByClass = mDriver.FindElementByClass(classname, timeout, raise)
End Function
    
'******************************************************************************
'* [概  要] FindElementByCss メソッド
'* [詳  細] Finds the first element matching the specified CSS selector.
'*          Selenium.WebDriver のメンバー
'*
'* @param cssselector   xxx
'* @param timeout   xxx
'* @param raise   xxx
'* @return Selenium.WebElement オブジェクト
'******************************************************************************
Public Function FindElementByCss(cssselector As String, Optional timeout As Long = -1, Optional raise As Boolean = True) As Selenium.WebElement
    Set FindElementByCss = mDriver.FindElementByCss(cssselector, timeout, raise)
End Function

'******************************************************************************
'* [概  要] FindElementById メソッド
'* [詳  細] Finds the first element matching the specified id.
'*          Selenium.WebDriver のメンバー
'*
'* @param id   xxx
'* @param timeout   xxx
'* @param raise   xxx
'* @return Selenium.WebElement オブジェクト
'******************************************************************************
Public Function FindElementById(id As String, Optional timeout As Long = -1, Optional raise As Boolean = True) As Selenium.WebElement
    Set FindElementById = mDriver.FindElementById(id, timeout, raise)
End Function

'******************************************************************************
'* [概  要] FindElementByLinkText メソッド
'* [詳  細] Finds the first element matching the specified link text.
'*          Selenium.WebDriver のメンバー
'*
'* @param linktext   xxx
'* @param timeout   xxx
'* @param raise   xxx
'* @return Selenium.WebElement オブジェクト
'******************************************************************************
Public Function FindElementByLinkText(linktext As String, Optional timeout As Long = -1, Optional raise As Boolean = True) As Selenium.WebElement
    Set FindElementByLinkText = mDriver.FindElementByLinkText(linktext, timeout, raise)
End Function

'******************************************************************************
'* [概  要] FindElementByName メソッド
'* [詳  細] Finds the first element matching the specified name.
'*          Selenium.WebDriver のメンバー
'*
'* @param name   xxx
'* @param timeout   xxx
'* @param raise   xxx
'* @return Selenium.WebElement オブジェクト
'******************************************************************************
Public Function FindElementByName(name As String, Optional timeout As Long = -1, Optional raise As Boolean = True) As Selenium.WebElement
    Set FindElementByName = mDriver.FindElementByName(name, timeout, raise)
End Function

'******************************************************************************
'* [概  要] FindElementByPartialLinkText メソッド
'* [詳  細] Finds the first of elements that match the part of the link text supplied
'*          Selenium.WebDriver のメンバー
'*
'* @param partiallinktext   xxx
'* @param timeout   xxx
'* @param raise   xxx
'* @return Selenium.WebElement オブジェクト
'******************************************************************************
Public Function FindElementByPartialLinkText(partiallinktext As String, Optional timeout As Long = -1, Optional raise As Boolean = True) As Selenium.WebElement
    Set FindElementByPartialLinkText = mDriver.FindElementByPartialLinkText(partiallinktext, timeout, raise)
End Function

'******************************************************************************
'* [概  要] FindElementByTag メソッド
'* [詳  細] Finds the first element matching the specified tag name.
'*          Selenium.WebDriver のメンバー
'*
'* @param tagname   xxx
'* @param timeout   xxx
'* @param raise   xxx
'* @return Selenium.WebElement オブジェクト
'******************************************************************************
Public Function FindElementByTag(tagName As String, Optional timeout As Long = -1, Optional raise As Boolean = True) As Selenium.WebElement
    Set FindElementByTag = mDriver.FindElementByTag(tagName, timeout, raise)
End Function

'******************************************************************************
'* [概  要] FindElementByXPath メソッド
'* [詳  細] Finds the first element matching the specified XPath query.
'*          Selenium.WebDriver のメンバー
'*
'* @param xpath   xxx
'* @param timeout   xxx
'* @param raise   xxx
'* @return Selenium.WebElement オブジェクト
'******************************************************************************
Public Function FindElementByXPath(xpath As String, Optional timeout As Long = -1, Optional raise As Boolean = True) As Selenium.WebElement
    Set FindElementByXPath = mDriver.FindElementByXPath(xpath, timeout, raise)
End Function

'******************************************************************************
'* [概  要] FindElements メソッド
'* [詳  細] Find all elements within the current context using the given mechanism.
'*          Selenium.WebDriver のメンバー
'*
'* @param By   xxx
'* @param minimum   xxx
'* @param timeout   xxx
'* @return Selenium.WebElements オブジェクト
'******************************************************************************
Public Function FindElements(By As By, Optional minimum As Long, Optional timeout As Long) As Selenium.WebElements
    Set FindElements = mDriver.FindElements(By, minimum, timeout)
End Function

'******************************************************************************
'* [概  要] FindElementsBy メソッド
'* [詳  細] Find all elements for a given strategy and value.
'*          Selenium.WebDriver のメンバー
'*
'* @param Strategy   xxx
'* @param value   xxx
'* @param minimum   xxx
'* @param timeout   xxx
'* @return Selenium.WebElements オブジェクト
'******************************************************************************
Public Function FindElementsBy(Strategy As Integer, value As String, Optional minimum As Long, Optional timeout As Long) As Selenium.WebElements
    Set FindElementsBy = mDriver.FindElementsBy(Strategy, value, minimum, timeout)
End Function

'******************************************************************************
'* [概  要] FindElementsByClass メソッド
'* [詳  細] Finds elements matching the specified CSS class.
'*          Selenium.WebDriver のメンバー
'*
'* @param classname   xxx
'* @param minimum   xxx
'* @param timeout   xxx
'* @return Selenium.WebElements オブジェクト
'******************************************************************************
Public Function FindElementsByClass(classname As String, Optional minimum As Long, Optional timeout As Long) As Selenium.WebElements
    Set FindElementsByClass = mDriver.FindElementsByClass(classname, minimum, timeout)
End Function

'******************************************************************************
'* [概  要] FindElementsByCss メソッド
'* [詳  細] Finds elements matching the specified CSS selector.
'*          Selenium.WebDriver のメンバー
'*
'* @param cssselector   xxx
'* @param minimum   xxx
'* @param timeout   xxx
'* @return Selenium.WebElements オブジェクト
'******************************************************************************
Public Function FindElementsByCss(cssselector As String, Optional minimum As Long, Optional timeout As Long) As Selenium.WebElements
    Set FindElementsByCss = mDriver.FindElementsByCss(cssselector, minimum, timeout)
End Function

'******************************************************************************
'* [概  要] FindElementsById メソッド
'* [詳  細] Finds elements matching the specified id.
'*          Selenium.WebDriver のメンバー
'*
'* @param id   xxx
'* @param minimum   xxx
'* @param timeout   xxx
'* @return Selenium.WebElements オブジェクト
'******************************************************************************
Public Function FindElementsById(id As String, Optional minimum As Long, Optional timeout As Long) As Selenium.WebElements
    Set FindElementsById = mDriver.FindElementsById(id, minimum, timeout)
End Function

'******************************************************************************
'* [概  要] FindElementsByLinkText メソッド
'* [詳  細] Finds elements matching the specified link text.
'*          Selenium.WebDriver のメンバー
'*
'* @param linktext   xxx
'* @param minimum   xxx
'* @param timeout   xxx
'* @return Selenium.WebElements オブジェクト
'******************************************************************************
Public Function FindElementsByLinkText(linktext As String, Optional minimum As Long, Optional timeout As Long) As Selenium.WebElements
    Set FindElementsByLinkText = mDriver.FindElementsByLinkText(linktext, minimum, timeout)
End Function

'******************************************************************************
'* [概  要] FindElementsByName メソッド
'* [詳  細] Finds elements matching the specified name.
'*          Selenium.WebDriver のメンバー
'*
'* @param name   xxx
'* @param minimum   xxx
'* @param timeout   xxx
'* @return Selenium.WebElements オブジェクト
'******************************************************************************
Public Function FindElementsByName(name As String, Optional minimum As Long, Optional timeout As Long) As Selenium.WebElements
    Set FindElementsByName = mDriver.FindElementsByName(name, minimum, timeout)
End Function

'******************************************************************************
'* [概  要] FindElementsByPartialLinkText メソッド
'* [詳  細] Finds the first of elements that match the part of the link text supplied
'*          Selenium.WebDriver のメンバー
'*
'* @param partiallinktext   xxx
'* @param minimum   xxx
'* @param timeout   xxx
'* @return Selenium.WebElements オブジェクト
'******************************************************************************
Public Function FindElementsByPartialLinkText(partiallinktext As String, Optional minimum As Long, Optional timeout As Long) As Selenium.WebElements
    Set FindElementsByPartialLinkText = mDriver.FindElementsByPartialLinkText(partiallinktext, minimum, timeout)
End Function

'******************************************************************************
'* [概  要] FindElementsByTag メソッド
'* [詳  細] Finds elements matching the specified tag name.
'*          Selenium.WebDriver のメンバー
'*
'* @param tagname   xxx
'* @param minimum   xxx
'* @param timeout   xxx
'* @return Selenium.WebElements オブジェクト
'******************************************************************************
Public Function FindElementsByTag(tagName As String, Optional minimum As Long, Optional timeout As Long) As Selenium.WebElements
    Set FindElementsByTag = mDriver.FindElementsByTag(tagName, minimum, timeout)
End Function

'******************************************************************************
'* [概  要] FindElementsByXPath メソッド
'* [詳  細] Finds elements matching the specified XPath query.
'*          Selenium.WebDriver のメンバー
'*
'* @param xpath   xxx
'* @param minimum   xxx
'* @param timeout   xxx
'* @return Selenium.WebElements オブジェクト
'******************************************************************************
Public Function FindElementsByXPath(xpath As String, Optional minimum As Long, Optional timeout As Long) As Selenium.WebElements
    Set FindElementsByXPath = mDriver.FindElementsByXPath(xpath, minimum, timeout)
End Function

'******************************************************************************
'* [概  要] GetPage メソッド（Get メソッド）
'* [詳  細] Loads a new web page with a relative or absolute URL
'*          Selenium.WebDriver のメンバー
'*
'* @param url   xxx
'* @param timeout   xxx
'* @param raise   xxx
'* @return Selenium.Table オブジェクト
'******************************************************************************
Public Function GetPage(url As String, Optional timeout As Long = -1, Optional raise As Boolean = True) As Boolean
    GetPage = mDriver.Get(url, timeout, raise)
End Function

'******************************************************************************
'* [概  要] GetClipBoard メソッド
'* [詳  細] Get text from the Clipboard
'*          Selenium.WebDriver のメンバー
'*
'* @return String
'******************************************************************************
Public Function GetClipBoard() As String
    GetClipBoard = mDriver.GetClipBoard
End Function

'******************************************************************************
'* [概  要] GoBack メソッド
'* [詳  細] Goes one step backward in the browser history.
'*          Selenium.WebDriver のメンバー
'*
'******************************************************************************
Public Sub GoBack()
    Call mDriver.GoBack
End Sub

'******************************************************************************
'* [概  要] GoForward メソッド
'* [詳  細] Goes one step forward in the browser history.
'*          Selenium.WebDriver のメンバー
'*
'******************************************************************************
Public Sub GoForward()
    Call mDriver.GoForward
End Sub
 
'******************************************************************************
'* [概  要] IsElementPresent メソッド
'* [詳  細] Indicates whether a WebElement is present using the given method.
'*          Selenium.WebDriver のメンバー
'*
'* @param By   xxx
'* @param timeout   xxx
'* @return Boolean
'******************************************************************************
Public Function IsElementPresent(By As By, Optional timeout As Long) As Boolean
    IsElementPresent = mDriver.IsElementPresent(By, timeout)
End Function

'******************************************************************************
'* [概  要] PageSource メソッド
'* [詳  細] Returns the page source
'*          Selenium.WebDriver のメンバー
'*
'* @return String
'******************************************************************************
Public Function PageSource() As String
    PageSource = mDriver.PageSource
End Function

'******************************************************************************
'* [概  要] PageSourceMatch メソッド
'* [詳  細] Returns the first occurence matching the regular expression.
'*          Selenium.WebDriver のメンバー
'*
'* @param pattern   xxx
'* @param group   xxx
'* @return String
'******************************************************************************
Public Function PageSourceMatch(pattern As String, Optional group As Integer) As String
    PageSourceMatch = mDriver.PageSourceMatch(pattern, group)
End Function

'******************************************************************************
'* [概  要] PageSourceMatches メソッド
'* [詳  細] Returns all the occurences matching the regular expression.
'*          Selenium.WebDriver のメンバー
'*
'* @param pattern   xxx
'* @param group   xxx
'* @return Selenium.List オブジェクト
'******************************************************************************
Public Function PageSourceMatches(pattern As String, Optional group As Integer) As Selenium.List
    Set PageSourceMatches = mDriver.PageSourceMatches(pattern, group)
End Function

'******************************************************************************
'* [概  要] Quit メソッド
'* [詳  細] Quits this driver, closing every associated window. Same as stop.
'*          Selenium.WebDriver のメンバー
'*
'******************************************************************************
Public Sub Quit()
    Call mDriver.Quit
End Sub

'******************************************************************************
'* [概  要] Refresh メソッド
'* [詳  細] Refreshes the current page.
'*          Selenium.WebDriver のメンバー
'*
'******************************************************************************
Public Sub Refresh()
    Call mDriver.Refresh
End Sub

'******************************************************************************
'* [概  要] Send メソッド
'* [詳  細] Sends a customized command.
'*          Selenium.WebDriver のメンバー
'*
'* @param method   xxx
'* @param relativeUri   xxx
'* @param param1   xxx
'* @param value1   xxx
'* @param param2   xxx
'* @param value2   xxx
'* @param param3   xxx
'* @param value3   xxx
'* @param param4   xxx
'* @param value4   xxx
'* @return Variant
'******************************************************************************
Public Function Send(method As String, relativeUri As String, Optional param1 As String, Optional value1, Optional param2 As String, Optional value2, Optional param3 As String, Optional value3, Optional param4 As String, Optional value4)
    Set Send = mDriver.Send(method, relativeUri, param1, value1, param2, value2, param3, value3, param4, value4)
End Function

'******************************************************************************
'* [概  要] SendKeys メソッド
'* [詳  細] Sends a sequence of keystrokes to the browser.
'*          Selenium.WebDriver のメンバー
'*
'* @param keysOrModifier   xxx
'* @param Keys   xxx
'******************************************************************************
Public Sub SendKeys(keysOrModifier As String, Optional Keys As String)
    Call mDriver.SendKeys(keysOrModifier, Keys)
End Sub

'******************************************************************************
'* [概  要] SetBinary メソッド
'* [詳  細] Set the path to the browser executable to use
'*          Selenium.WebDriver のメンバー
'*
'* @param path   xxx
'******************************************************************************
Public Sub SetBinary(path As String)
    Call mDriver.SetBinary(path)
End Sub

'******************************************************************************
'* [概  要] SetCapability メソッド
'* [詳  細] Selenium.Table オブジェクトを取得し返却する。
'*          Selenium.WebDriver のメンバー
'*
'* @param key   xxx
'* @param value   xxx
'******************************************************************************
Public Sub SetCapability(key As String, value)
    Call mDriver.SetCapability(key, value)
End Sub

'******************************************************************************
'* [概  要] SetClipBoard メソッド
'* [詳  細] Set text in the Clipboard。
'*          Selenium.WebDriver のメンバー
'*
'* @param text   xxx
'******************************************************************************
Public Sub SetClipBoard(text As String)
    Call mDriver.SetClipBoard(text)
End Sub

'******************************************************************************
'* [概  要] SetPreference メソッド
'* [詳  細] Selenium.Table オブジェクトを取得し返却する。
'*          Selenium.WebDriver のメンバー
'*
'* @param key   xxx
'* @param value   xxx
'******************************************************************************
Public Sub SetPreference(key As String, value)
    Call mDriver.SetPreference(key, value)
End Sub

'******************************************************************************
'* [概  要] SetProfile メソッド
'* [詳  細] Selenium.Table オブジェクトを取得し返却する。
'*          Selenium.WebDriver のメンバー
'*
'* @param nameOrDirectory   xxx
'* @param persistant   xxx
'******************************************************************************
Public Sub SetProfile(nameOrDirectory As String, Optional persistant As Boolean = False)
    Call mDriver.SetProfile(nameOrDirectory, persistant)
End Sub

'******************************************************************************
'* [概  要] Start メソッド
'* [詳  細] Starts a new Selenium session
'*          Selenium.WebDriver のメンバー
'*
'* @param browser   xxx
'* @param baseUrl   xxx
'******************************************************************************
Public Sub Start(Optional browser As String, Optional baseUrl As String)
    Call mDriver.Start(browser, baseUrl)
End Sub

'******************************************************************************
'* [概  要] StartRemotely メソッド
'* [詳  細] Starts remotely a new Selenium session
'*          Selenium.WebDriver のメンバー
'*
'* @param executorUri   xxx
'* @param browser   xxx
'* @param version   xxx
'* @param platform   xxx
'******************************************************************************
Public Sub StartRemotely(executorUri As String, Optional browser As String, Optional version As String, Optional platform As String)
    Call mDriver.StartRemotely(executorUri, browser, version, platform)
End Sub

'******************************************************************************
'* [概  要] SwitchToAlert メソッド
'* [詳  細] Switch the focus to an alert on the page.
'*          Selenium.WebDriver のメンバー
'*
'* @param timeout   xxx
'* @param raise   xxx
'* @return Selenium.Alert オブジェクト
'******************************************************************************
Public Function SwitchToAlert(Optional timeout As Long = -1, Optional raise As Boolean = True) As Selenium.Alert
    Set SwitchToAlert = mDriver.SwitchToAlert(timeout, raise)
End Function

'******************************************************************************
'* [概  要] SwitchToDefaultContent メソッド
'* [詳  細] Select either the first frame on the page or the main document when a page contains iFrames.
'*          Selenium.WebDriver のメンバー
'*
'******************************************************************************
Public Sub SwitchToDefaultContent()
    Call mDriver.SwitchToDefaultContent
End Sub

'******************************************************************************
'* [概  要] SwitchToFrame メソッド
'* [詳  細] Switche the focus to the specified frame, by index, name or WebElement.
'*          Selenium.WebDriver のメンバー
'*
'* @param identifier   xxx
'* @param timeout   xxx
'* @param raise   xxx
'* @return Boolean
'******************************************************************************
Public Function SwitchToFrame(identifier, Optional timeout As Long = -1, Optional raise As Boolean = True) As Boolean
    SwitchToFrame = mDriver.SwitchToFrame(identifier, timeout, raise)
End Function

'******************************************************************************
'* [概  要] SwitchToNextWindow メソッド
'* [詳  細] Switche the focus to the next window.
'*          Selenium.WebDriver のメンバー
'*
'* @param timeout   xxx
'* @param raise   xxx
'* @return Selenium.Window オブジェクト
'******************************************************************************
Public Function SwitchToNextWindow(Optional timeout As Long = -1, Optional raise As Boolean = True) As Selenium.Window
    Set SwitchToNextWindow = mDriver.SwitchToNextWindow(timeout, raise)
End Function

'******************************************************************************
'* [概  要] SwitchToParentFrame メソッド
'* [詳  細] Switche the focus to the parent frame.
'*          Selenium.WebDriver のメンバー
'*
'******************************************************************************
Public Sub SwitchToParentFrame()
    Call mDriver.SwitchToParentFrame
End Sub

'******************************************************************************
'* [概  要] SwitchToPreviousWindow メソッド
'* [詳  細] Switche the focus to the previous window.
'*          Selenium.WebDriver のメンバー
'*
'* @return Selenium.Window オブジェクト
'******************************************************************************
Public Function SwitchToPreviousWindow() As Selenium.Window
    Set SwitchToPreviousWindow = mDriver.SwitchToPreviousWindow
End Function

'******************************************************************************
'* [概  要] SwitchToWindowByName メソッド
'* [詳  細] Switche the focus to the specified window.
'*          Selenium.WebDriver のメンバー
'*
'* @param name   xxx
'* @param timeout   xxx
'* @param raise   xxx
'* @return Selenium.Window オブジェクト
'******************************************************************************
Public Function SwitchToWindowByName(name As String, Optional timeout As Long = -1, Optional raise As Boolean = True) As Selenium.Window
    Set SwitchToWindowByName = mDriver.SwitchToWindowByName(name, timeout, raise)
End Function

'******************************************************************************
'* [概  要] SwitchToWindowByTitle メソッド
'* [詳  細] Switche the focus to the specified window.
'*          Selenium.WebDriver のメンバー
'*
'* @param Title   xxx
'* @param timeout   xxx
'* @param raise   xxx
'* @return Selenium.Window オブジェクト
'******************************************************************************
Public Function SwitchToWindowByTitle(Title As String, Optional timeout As Long = -1, Optional raise As Boolean = True) As Selenium.Window
    Set SwitchToWindowByTitle = mDriver.SwitchToWindowByTitle(Title, timeout, raise)
End Function

'******************************************************************************
'* [概  要] TakeScreenshot メソッド
'* [詳  細] Capture a screenshot
'*          Selenium.WebDriver のメンバー
'*
'* @param delayms   xxx
'* @return Selenium.Image オブジェクト
'******************************************************************************
Public Function TakeScreenshot(Optional delayms As Long = 100) As Selenium.Image
    Set TakeScreenshot = mDriver.TakeScreenshot(delayms)
End Function

'******************************************************************************
'* [概  要] WaitUntil メソッド（Until メソッド）
'* [詳  細] Waits for a procedure to return true
'*          Selenium.WebDriver のメンバー
'*
'* @param procedure   xxx
'* @param argument   xxx
'* @param timeout   xxx
'* @return Variant
'******************************************************************************
Public Function WaitUntil(procedure, Optional argument, Optional timeout As Long = -1)
    Set WaitUntil = mDriver.WaitUntil(procedure, argument, timeout)
End Function

'******************************************************************************
'* [概  要] Wait メソッド
'* [詳  細] Waits the specified time in millisecond before executing the next command
'*          Selenium.WebDriver のメンバー
'*
'* @param timems   xxx
'******************************************************************************
Public Sub Wait(timems As Long)
    Call mDriver.Wait(timems)
End Sub

'******************************************************************************
'* [概  要] WaitForScript メソッド
'* [詳  細] Waits for the Javascript engine to return true or not null
'*          Selenium.WebDriver のメンバー
'*
'* @param script   xxx
'* @param arguments   xxx
'* @param timeout   xxx
'* @return Variant
'******************************************************************************
Public Function WaitForScript(script As String, Optional arguments, Optional timeout As Long = -1)
    WaitForScript = mDriver.WaitForScript(script, arguments, timeout)
End Function

'******************************************************************************
'* [概  要] WaitNotElement メソッド
'* [詳  細] Waits for an element to disappear from the page
'*          Selenium.WebDriver のメンバー
'*
'* @param By   xxx
'* @param timeout   xxx
'******************************************************************************
Public Sub WaitNotElement(By As By, Optional timeout As Long = -1)
    Call mDriver.WaitNotElement(By, timeout)
End Sub
'******************************************************************************
'* [概  要] AddCookie メソッド
'* [詳  細] Adds a cookie to the current page.
'*          Selenium.Manage のメンバー
'*
'* @param name   xxx
'* @param value   xxx
'* @param Domain   xxx
'* @param path   xxx
'* @param Expiry   xxx
'* @param Secure   xxx
'* @param httpOnly   xxx
'******************************************************************************
Public Sub AddCookie(name As String, value As String, Optional Domain As String, Optional path As String, Optional Expiry As String, Optional Secure As Boolean = False, Optional httpOnly As Boolean = False)
    Call mDriver.Manage.AddCookie(name, value, Domain, path, Expiry, Secure, httpOnly)
End Sub

    
'******************************************************************************
'* [概  要] DeleteAllCookies メソッド
'* [詳  細] Deletes all cookies from the page.
'*          Selenium.Manage のメンバー
'*
'******************************************************************************
Public Sub DeleteAllCookies()
    Call mDriver.Manage.DeleteAllCookies
End Sub

'******************************************************************************
'* [概  要] DeleteCookieByName メソッド
'* [詳  細] Deletes the cookie with the specified name from the page.
'*          Selenium.Manage のメンバー
'*
'* @param name   xxx
'******************************************************************************
Public Sub DeleteCookieByName(name As String)
    Call mDriver.Manage.DeleteCookieByName(name)
End Sub

'******************************************************************************
'* [概  要] FindCookieByName メソッド
'* [詳  細] Gets a cookie with the specified name.
'*          Selenium.Manage のメンバー
'*
'* @param name   xxx
'* @return Selenium.Cookie オブジェクト
'******************************************************************************
Public Function FindCookieByName(name As String) As Cookie
    Set FindCookieByName = mDriver.Manage.FindCookieByName(name)
End Function

'******************************************************************************
'* [概  要] Location メソッド
'* [詳  細] Get the current geo location.
'*          Selenium.Manage のメンバー
'*
'* @return Selenium.Dictionary オブジェクト
'******************************************************************************
Public Function Location() As Dictionary
    Set Location = mDriver.Manage.Location
End Function

'******************************************************************************
'* [概  要] SetLocation メソッド
'* [詳  細] Set the current geo location.
'*          Selenium.Manage のメンバー
'*
'* @param latitude   xxx
'* @param longitude   xxx
'* @param altitude   xxx
'******************************************************************************
Public Sub SetLocation(latitude As Double, longitude As Double, altitude As Double)
    Call mDriver.Manage.SetLocation(latitude, longitude, altitude)
End Sub


'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------

'******************************************************************************
'* [概  要] ChromePreCheck メソッド
'* [詳  細] ChromeDriverが最新かチェックし、古い場合はダウンロードを促す。
'*
'* @return Boolean チェック結果
'******************************************************************************
Public Function ChromePreCheck() As Boolean
    Debug.Print "--------------------------------------------------------------------------------"
    Debug.Print " Chrome PreCheck"
    Debug.Print "--------------------------------------------------------------------------------"
    Dim chromeVersion As String: chromeVersion = GetChromeVersion
    Debug.Print "Chrome Version       : " & chromeVersion
    Debug.Print "ChromeDriver Version : " & GetChromeDriverVersion
    Debug.Print "Driver適合チェック   : " & IsChromeDriverMatching
    Debug.Print "--------------------------------------------------------------------------------"
    Debug.Print

    If Not IsChromeDriverMatching() Then
        Dim rc As VbMsgBoxResult
        rc = MsgBox("Chrome と ChromeDriver のバージョンが適合しません。" & vbNewLine & "適合するChromeDriverをダウンロードしますか？", vbYesNo + vbQuestion)
        If rc = vbYes Then
            Dim driverLatestVersion As String: driverLatestVersion = GetChromeDriverLatestVersion(chromeVersion)
            Debug.Print "ChromeDriver最新バージョン : " & driverLatestVersion
            If DownloadChromeDriverZipFile(driverLatestVersion, ThisWorkbook.path & "\chromedriver_win32.zip") Then
                Call UnZipFile(ThisWorkbook.path & "\chromedriver_win32.zip", ThisWorkbook.path)
                Call CopyChromeDriver(ThisWorkbook.path & "\chromedriver_win32", "\chromedriver.exe")
                Debug.Print "ChromeDriverを更新しました。"
                Debug.Print
            Else
                Debug.Print "[ERROR] ダウンロードに失敗しました。"
                Call MsgBox("ダウンロードに失敗しました。", vbOK + vbExclamation)
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    ChromePreCheck = True

End Function

'******************************************************************************
'* [概  要] GetChromeVersion メソッド
'* [詳  細] Chromeバージョンを取得する。
'*
'* @return String       Chromeバージョン
'******************************************************************************
Public Function GetChromeVersion() As String
    Dim stdRegProv As Object: Set stdRegProv = CreateStdRegProv()
    
    ' chrome.exeのインストール先をレジストリより取得
    Dim chromePath As String
    stdRegProv.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\WOW6432Node\Clients\StartMenuInternet\Google Chrome\shell\open\command", "", chromePath
    If chromePath = "" Then
        stdRegProv.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Clients\StartMenuInternet\Google Chrome\shell\open\command", "", chromePath
    End If
    If chromePath = "" Then
        Exit Function
    End If
    
    '' chrome.exeのバージョンをWmiを利用して取得
    'Dim chromeWmiFiles As Object
    'Set chromeWmiFiles = GetWmiServer().ExecQuery("Select * From CIM_DataFile Where Name = '" & Replace(Replace(chromePath, """", ""), "\", "\\") & "'")
    '
    'If chromeWmiFiles.count = 0 Then
    '    Exit Function
    'End If
    'Dim versionFullString As String: versionFullString = chromeWmiFiles.ItemIndex(0).version
    
    Dim versionFullString As String: versionFullString = mFso.GetFileVersion(Replace(chromePath, """", ""))
    
    ' メジャーバージョン.マイナーバージョン.ビルドバージョンのみに編集（マイナー番号（リビジョン）を除去）して返却
    GetChromeVersion = Mid(versionFullString, 1, InStrRev(versionFullString, ".") - 1)
End Function

'******************************************************************************
'* [概  要] GetChromeDriverLatestVersion メソッド
'* [詳  細] Chromeバージョンに対応する最新のChromeDriverバージョンを取得する。
'*
'* @param chromeVersion Chromeバージョン
'* @return String       ChromeDriverの最新バージョン
'******************************************************************************
Public Function GetChromeDriverLatestVersion(chromeVersion As String) As String
    Dim winHttp: Set winHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    winHttp.Open "GET", URL_CHROME_DRIVER_LATEST_RELEASE_VERSION & chromeVersion, False
    winHttp.Send
    Dim resBody As Variant
    With CreateObject("ADODB.Stream")
        .Type = 1 ' adTypeBinary
        .Open
        .Write winHttp.ResponseBody
        .Position = 0
        .Type = 2 ' adTypeText
        .Charset = "UTF-8"
        resBody = .ReadText(-1) ' adReadAll
        .Close
    End With
    GetChromeDriverLatestVersion = resBody
End Function

'******************************************************************************
'* [概  要] DownloadChromeDriverZipFile メソッド
'* [詳  細] ChromeDriverのzipファイルをダウンロードする。
'*
'* @param driverVersion ダウンロード対象のChromeDriverバージョン
'* @param saveFilePath  保存先ファイルパス
'* @return Boolean 実行結果（OK:True, NG:False）
'******************************************************************************
Public Function DownloadChromeDriverZipFile(driverVersion As String, saveFilePath As String) As Boolean
    Const adTypeBinary = 1
    Const adSaveCreateOverWrite = 2

    Dim winHttp: Set winHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    winHttp.Open "GET", Replace(URL_CHROME_DRIVER_ZIP, "{driver_version}", driverVersion), False
    winHttp.Send
    
    If winHttp.Status >= 200 And winHttp.Status < 400 Then
        With CreateObject("ADODB.Stream")
            .Type = 1 ' adTypeBinary
            .Open
            .Write winHttp.ResponseBody
            .SaveToFile saveFilePath, 2 ' adSaveCreateOverWrite
            .Close
        End With
        DownloadChromeDriverZipFile = True
    Else
        DownloadChromeDriverZipFile = False
    End If
End Function

'******************************************************************************
'* [概  要] GetChromeDriverMainVersion メソッド
'* [詳  細] ChromeDriverのMainバージョンを取得する。
'*          （ChromeDriverのバージョンから最後のリビジョン番号を除去したもの。）
'*
'* @return String ChromeDriverのMainバージョン
'******************************************************************************
Public Function GetChromeDriverMainVersion() As String
    Dim chromeDriverVersion As String: chromeDriverVersion = GetChromeDriverVersion()
    If chromeDriverVersion <> "" Then
        GetChromeDriverMainVersion = Mid(chromeDriverVersion, 1, InStrRev(chromeDriverVersion, ".") - 1)
    End If
End Function

'******************************************************************************
'* [概  要] GetChromeDriverVersion メソッド
'* [詳  細] ChromeDriverのバージョンを取得する。
'*
'* @return String ChromeDriverのバージョン
'******************************************************************************
Public Function GetChromeDriverVersion() As String
    Dim result As String: result = Commander.GetCommandResult(Environ("LOCALAPPDATA") & "\SeleniumBasic\chromedriver.exe --version")
    Dim resList: resList = Split(result, " ")
    If UBound(resList) >= 1 Then GetChromeDriverVersion = resList(1)
End Function

'******************************************************************************
'* [概  要] IsChromeDriverMatching メソッド
'* [詳  細] ChromeDriverがChromeのバージョンに適合するかチェックする。
'*
'* @return Boolean チェック結果（OK:True, NG:False）
'******************************************************************************
Public Function IsChromeDriverMatching() As Boolean
    IsChromeDriverMatching = (GetChromeDriverMainVersion() = GetChromeVersion())
End Function

'******************************************************************************
'* [概  要] CopyChromeDriver メソッド
'* [詳  細] ChromeDriverを所定の位置に移動し、移動元のフォルダを削除する。
'*          所定の位置に元々あったChromeDriverは、chromebackupフォルダに退避。
'*
'* @param srcDir      コピー元フォルダ（コピー後削除される）
'* @param srcFileName コピー元ChromeDriverファイル名
'******************************************************************************
Public Sub CopyChromeDriver(srcDir As String, srcFileName As String)
    Dim chromeDriverFilePath As String
    chromeDriverFilePath = Environ("LOCALAPPDATA") & "\SeleniumBasic\chromedriver.exe"
    
    If mFso.FileExists(chromeDriverFilePath) Then
        Dim backupPath As String
        backupPath = Environ("LOCALAPPDATA") & "\SeleniumBasic\chromebackup\"
        Call CreateFolderTree(backupPath)
        
        mFso.MoveFile chromeDriverFilePath, backupPath & "chromedriver_" & Format(Now, "yyyymmddhhnnss") & ".exe"
    End If
    
     mFso.MoveFile mFso.BuildPath(srcDir, srcFileName), chromeDriverFilePath
     mFso.DeleteFolder srcDir, True
End Sub

'******************************************************************************
'* [概  要] SelectAlert メソッド
'* [詳  細] アラートポップアップのボタンを選択する。
'* [参  考] https://rabbitfoot.xyz/selenium-alertpush/
'*
'* @param Accept OKかどうか（True:Accept, False:Dismiss）
'*
'******************************************************************************
Public Sub SelectAlert(Optional Accept As Boolean = True)
    Dim sAlert As Selenium.Alert
    Set sAlert = mDriver.SwitchToAlert(raise:=False)
    
    If Not sAlert Is Nothing Then
        If Accept Then
            sAlert.Accept
            Exit Sub
        End If
        sAlert.Dismiss
    End If
End Sub

'******************************************************************************
'* [概  要] LoginBasicAuthDialog メソッド
'* [詳  細] Basic認証ダイアログに認証情報を入力してログインする。
'* [参  考] SeleniumBasic Examples.xlsm
'*
'* @param id       ユーザID
'* @param password パスワード
'******************************************************************************
Public Sub LoginBasicAuthDialog(id As String, Password As String)
    Dim sAlert As Selenium.Alert
    Set sAlert = mDriver.SwitchToAlert(raise:=False)

    If Not sAlert Is Nothing Then
        sAlert.SetCredentials id, Password
        sAlert.Accept
    End If
End Sub

'******************************************************************************
'* [概  要] DownloadByLinkText メソッド
'* [詳  細] 指定したテキストに合致するリンクから直接ファイルをダウンロードする。
'* [参  考] SeleniumBasic Examples.xlsm
'*
'* @param text         リンクテキスト
'* @param saveFilePath 保存先ファイルパス
'******************************************************************************
Public Sub DownloadStaticByLinkText(text As String, saveFilePath As String)
    Dim elem As WebElement
    Set elem = mDriver.FindElementByLinkText(text)
    Download_StaticLink elem, saveFilePath
End Sub

'******************************************************************************
'* [概  要] DownloadClickByLinkText メソッド
'* [詳  細] 指定したテキストに合致するリンクをクリックしファイルをダウンロードする。
'* [参  考] SeleniumBasic Examples.xlsm
'* [備  考] ブラウザ起動前に SetPreferenceForDownload メソッドを呼び出す
'*          （ダウンロードフォルダを指定しておく）必要あり。
'*
'* @param text      リンクテキスト
'* @param dlDirPath ダウンロードフォルダのパス
'* @param dlFileExt ダウンロードファイル拡張子
'* @return String   ダウンロードファイルパス
'******************************************************************************
Public Function DownloadClickByLinkText(text As String, dlDirPath As String, dlFileExt As String) As String
    ' Init the file waiter
    WaitNewFile dlDirPath & "\*." & dlFileExt
    
    mDriver.FindElementByLinkText(text).Click
  
    ' Waits for a new file
    DownloadClickByLinkText = WaitNewFile()
End Function

'******************************************************************************
'* [概  要] ListLinks メソッド
'* [詳  細] リンク（aタグのhref属性）のリストを取得する。
'* [参  考] SeleniumBasic Examples.xlsm
'*
'* @return Selenium.List   ダウンロードファイルパス
'******************************************************************************
Public Function ListLinks() As Selenium.List
    Dim links As List
    Set links = mDriver.FindElementsByTag("a").Attribute("href")
    links.Distinct
    links.Sort
    Set ListLinks = links
End Function

'******************************************************************************
'* [概  要] FindElementByAttribute メソッド
'* [詳  細] WebElementsのうち、指定したAttributeの値と最初に一致するWebElementを
'*          取得する。
'*          要素数が多いと処理時間がかかるので、FindElementByXPathの使用を推奨。
'*
'* @param elems          対象のWebElements
'* @param attributeName  検索する属性名
'* @param attributeValue 検索する属性値
'* @return WebElement    最初に一致したWebElement
'******************************************************************************
Public Function FindElementByAttribute(elems As WebElements, attributeName As String, attributeValue As String) As WebElement
    Dim elem As Selenium.WebElement
    For Each elem In elems
        If elem.Attribute(attributeName) = attributeValue Then
            Set FindElementByAttribute = elem
            Exit Function
        End If
    Next
End Function

'******************************************************************************
'* [概  要] FindElementByAttributeUseXPath メソッド
'* [詳  細] 指定したTag、Attributeの値と一致するWebElementを
'*          取得する。
'*
'* @param targetElement   検索対象とするルートのWebElement
'*                        （指定しない場合はドキュメントルートから検索）
'* @param tagName         対象のタグ名（全てのタグは*指定）
'* @param attributeName   検索する属性名（textも指定可能(子要素テキストも参照)）
'* @param attributeValue  検索する属性値
'* @param useContains     部分一致検索を有効にするか
'* @param andOr　　　　　 属性2の追加条件（"and"または"or"を指定）
'* @param attributeName2  検索する属性名2（textも指定可能(子要素テキストも参照)）
'* @param attributeValue2 検索する属性値2
'* @param useContains2    部分一致検索を有効にするか（属性2指定時）
'* @return WebElement     一致したWebElement
'******************************************************************************
Public Function FindElementByAttributeUseXPath( _
                    targetElement As Selenium.WebElement, tagName As String, _
                    attributeName As String, attributeValue As String, Optional useContains As Boolean = False, _
                    Optional andOr As String, Optional attributeName2 As String, Optional attributeValue2 As String, _
                    Optional useContains2 As Boolean = False _
                ) As WebElement
    Dim xPathString As String
    xPathString = MakeXPathString(tagName, attributeName, attributeValue, useContains, andOr, attributeName2, attributeValue2, useContains2)

    If targetElement Is Nothing Then
        Set FindElementByAttributeUseXPath = mDriver.FindElementByXPath("//" & xPathString)
    Else
        Set FindElementByAttributeUseXPath = targetElement.FindElementByXPath(xPathString)
    End If
End Function

'******************************************************************************
'* [概  要] FindElementsByAttributeUseXPath メソッド
'* [詳  細] 指定したTag、Attributeの値と一致するWebElementsを
'*          取得する。
'*
'* @param targetElement   検索対象とするルートのWebElement
'*                        （指定しない場合はドキュメントルートから検索）
'* @param tagName         対象のタグ名（全てのタグは*指定）
'* @param attributeName   検索する属性名（textも指定可能(子要素テキストも参照)）
'* @param attributeValue  検索する属性値
'* @param useContains     部分一致検索を有効にするか
'* @param andOr　　　　　 属性2の追加条件（"and"または"or"を指定）
'* @param attributeName2  検索する属性名2（textも指定可能(子要素テキストも参照)）
'* @param attributeValue2 検索する属性値2
'* @param useContains2    部分一致検索を有効にするか（属性2指定時）
'* @return WebElements     一致したWebElements
'******************************************************************************
Public Function FindElementsByAttributeUseXPath( _
                    targetElement As Selenium.WebElement, tagName As String, _
                    attributeName As String, attributeValue As String, Optional useContains As Boolean = False, _
                    Optional andOr As String, Optional attributeName2 As String, Optional attributeValue2 As String, _
                    Optional useContains2 As Boolean = False _
                ) As WebElements
    Dim xPathString As String
    xPathString = MakeXPathString(tagName, attributeName, attributeValue, useContains, andOr, attributeName2, attributeValue2, useContains2)

    If targetElement Is Nothing Then
        Set FindElementsByAttributeUseXPath = mDriver.FindElementsByXPath("//" & xPathString)
    Else
        Set FindElementsByAttributeUseXPath = targetElement.FindElementsByXPath(xPathString)
    End If
End Function

'******************************************************************************
'* [概  要] MakeXPathString メソッド
'* [詳  細] XPathの条件文字列を作成する。
'*
'* @param tagName         対象のタグ名（全てのタグは*指定）
'* @param attributeName   検索する属性名（textも指定可能(子要素テキストも参照)）
'* @param attributeValue  検索する属性値
'* @param useContains     部分一致検索を有効にするか
'* @param andOr　　　　　 属性2の追加条件（"and"または"or"を指定）
'* @param attributeName2  検索する属性名2（textも指定可能(子要素テキストも参照)）
'* @param attributeValue2 検索する属性値2
'* @param useContains2    部分一致検索を有効にするか（属性2指定時）
'* @return String         XPathの条件文字列
'******************************************************************************
Private Function MakeXPathString(tagName As String, attributeName As String, attributeValue As String, Optional useContains As Boolean = False, _
        Optional andOr As String, Optional attributeName2 As String, Optional attributeValue2 As String, _
        Optional useContains2 As Boolean = False) As String
    Dim xPathString As String
    xPathString = tagName & "[{searchValue}]"
    
    Dim searchValue As String
    searchValue = MakeXPathStringPart(attributeName, attributeValue, useContains)
    If attributeName2 <> "" Then
        searchValue = searchValue & " " & andOr & " " & MakeXPathStringPart(attributeName2, attributeValue2, useContains2)
    End If
    xPathString = Replace(xPathString, "{searchValue}", searchValue)
    MakeXPathString = xPathString
End Function

'******************************************************************************
'* [概  要] MakeXPathStringPart メソッド
'* [詳  細] XPathの条件文字列の部品を作成する。
'*
'* @param attributeName   検索する属性名（textも指定可能(子要素テキストも参照)）
'* @param attributeValue  検索する属性値
'* @param useContains     部分一致検索を有効にするか
'* @return String         XPathの条件文字列の部品
'******************************************************************************
Private Function MakeXPathStringPart(attributeName As String, attributeValue As String, Optional useContains As Boolean = False) As String
    Dim searchValue As String
    Dim sign As String: sign = "="
    If useContains Then
        searchValue = "contains({expression})"
        sign = ","
    Else
        searchValue = "{expression}"
    End If
    
    If attributeName = "text" Then
        searchValue = Replace(searchValue, "{expression}", "." & sign & "'" & attributeValue & "'")
    Else
        searchValue = Replace(searchValue, "{expression}", "@" & attributeName & sign & "'" & attributeValue & "'")
    End If
    MakeXPathStringPart = searchValue
End Function

'******************************************************************************
'* [概  要] FindElementByText メソッド
'* [詳  細] WebElementsのうち、指定したTextの値と最初に一致するWebElementを
'*          取得する。
'*          要素数が多いと処理時間がかかるので、FindElementByXPathの使用を推奨。
'*
'* @param elems       対象のWebElements
'* @param textValue   検索するテキスト
'* @return WebElement 最初に一致したWebElement
'******************************************************************************
Public Function FindElementByText(elems As WebElements, textValue As String) As WebElement

    Dim elem As Selenium.WebElement
    For Each elem In elems
        If elem.text = textValue Then
            Set FindElementByText = elem
            Exit Function
        End If
    Next
End Function

'******************************************************************************
'* [概  要] WaitElement メソッド
'* [詳  細] 画面要素が表示されるまで待つ。
'*
'* @param elem       対象のWebElement
'* @param waitMsec   1回の待ち時間
'* @param retryCount リトライ回数
'******************************************************************************
Public Sub WaitElement(elem As WebElement, Optional waitMsec As Long = 1500, Optional retryCount As Long = 3)
    Dim cnt As Long: cnt = 0
    Do
        If cnt = retryCount Then Err.raise 9999, "SeleniumExDriver#WaitElement", "表示待ちタイムアウト"
        Sleep waitMsec
        cnt = cnt + 1
    Loop Until elem.IsDisplayed And elem.IsPresent

End Sub

'******************************************************************************
'* [概  要] GetElementAttributes メソッド
'* [詳  細] 指定したエレメントが持つ属性を取得する。
'* [参  考] SeleniumBasic Examples.xlsm
'*
'* @param elem     対象のWebElement
'* @return Selenium.Dictionary Attributesセット
'******************************************************************************
Public Function GetElementAttributes(elem As WebElement) As Selenium.Dictionary
  Set GetElementAttributes = elem.ExecuteScript(JS_LIST_ATTRIBUTES)
End Function

'******************************************************************************
'* [概  要] GetPerformanceTiming メソッド
'* [詳  細] 処理性能を測るページ応答時間（responseEnd - navigationStart）、
'*          ページロード時間（loadEventEnd - responseEnd）を取得する。
'* [参  考] SeleniumBasic Examples.xlsm
'*          https://developer.mozilla.org/en/docs/Web/API/Navigation_timing_API
'*          https://developer.mozilla.org/en-US/docs/Web/Performance/Navigation_and_resource_timings
'*          http://www.html5rocks.com/en/tutorials/webperformance/basics
'*          http://www.w3.org/TR/navigation-timing
'*
'* @return Variant ページ応答時間/ページロード時間のVariant配列リスト（2要素）
'******************************************************************************
Public Function GetPerformanceTiming()
    GetPerformanceTiming = mDriver.ExecuteScript(JS_GET_TIMINGS).Values
End Function

'******************************************************************************
'* [概  要] StartChrome メソッド
'* [詳  細] Chromeを起動する。
'*
'******************************************************************************
Public Sub StartChrome(Optional baseUrl As String)
    mDriver.Start "chrome", baseUrl
End Sub

'******************************************************************************
'* [概  要] StartInternetExplorer メソッド
'* [詳  細] Internet Explorerを起動する。
'*
'******************************************************************************
Public Sub StartInternetExplorer(Optional baseUrl As String)
    mDriver.Start "ie", baseUrl
End Sub

'******************************************************************************
'* [概  要] StartFirefox メソッド
'* [詳  細] Firefoxを起動する。
'*
'******************************************************************************
Public Sub StartFirefox(Optional baseUrl As String)
    mDriver.Start "firefox", baseUrl
End Sub

'******************************************************************************
'* [概  要] StartEdge メソッド
'* [詳  細] Microsoft Edgeを起動する。
'*
'******************************************************************************
Public Sub StartEdge(Optional baseUrl As String)
    mDriver.Start "edge", baseUrl
End Sub

'******************************************************************************
'* [概  要] WindowMaximize メソッド
'* [詳  細] Windowを最大化する。
'*
'******************************************************************************
Public Sub WindowMaximize(Optional baseUrl As String)
    mDriver.Window.Maximize
End Sub

'******************************************************************************
'* [概  要] InputText メソッド
'* [詳  細] 要素（HTML部品）にテキストを入力。
'*
'* @param elm WebElement
'* @param txt 入力するテキスト
'******************************************************************************
Public Sub InputText(elm As Selenium.WebElement, txt As String)
    elm.Clear
    elm.SendKeys txt
    elm.SendKeys mKeys.Enter
End Sub

'******************************************************************************
'* [概  要] PulldownSelectByText メソッド
'* [詳  細] プルダウンのリストからテキストを指定して選択する。
'*
'* @param elm WebElement
'* @param txt 選択対象のテキスト
'******************************************************************************
Public Sub PulldownSelectByText(elm As Selenium.WebElement, txt As String)
    elm.AsSelect.SelectByText txt
End Sub

'******************************************************************************
'* [概  要] PulldownSelectByIndex メソッド
'* [詳  細] プルダウンのリストからインデックスを指定して選択する。
'*
'* @param elm WebElement
'* @param ind 選択対象のインデックス
'******************************************************************************
Public Sub PulldownSelectByIndex(elm As Selenium.WebElement, ind As Integer)
    elm.AsSelect.SelectByIndex ind
End Sub

'******************************************************************************
'* [概  要] GetDomObject メソッド
'* [詳  細]ページソースからHTML DOMオブジェクトを取得する。
'* [参  考] SeleniumBasic Examples.xlsm
'*
'* @return Object HTML DOMオブジェクト
'******************************************************************************
Public Function GetDomObject() As Object
  Dim html As Object
  Set html = CreateObject("htmlfile")
  html.Open
  html.Write mDriver.PageSource()
  html.Close
  
  Set GetDomObject = html
End Function

'******************************************************************************
'* [概  要] SetPreferenceForDownload メソッド
'* [詳  細] ダウンロード用の設定を行う（Chrome）。
'* [参  考] https://stackoverflow.com/questions/32157107/launching-chrome-or-ie-via-selenium-vba
'*
'* @param downloadDirPath ダウンロードを行うディレクトリパス
'* TODO:設定値用のクラス作成
'******************************************************************************
Public Sub SetPreferenceForDownload(Optional downloadDirPath As String)
    If downloadDirPath = "" Then downloadDirPath = ThisWorkbook.path
    mDriver.SetPreference "download.default_directory", downloadDirPath
    mDriver.SetPreference "download.directory_upgrade", True
    mDriver.SetPreference "plugins.plugins_disabled", Array("Chrome PDF Viewer")
    mDriver.SetPreference "download.prompt_for_download", False
    mDriver.SetPreference "download.extensions_to_open", ""
    mDriver.SetPreference "--disable-popup-blocking", ""
    mDriver.SetPreference "--enable-panels", ""

End Sub

'******************************************************************************
'* [概  要] SetTimeoutsImplicitWait メソッド
'* [詳  細] Seleniumの読み込み待機時間を指定する。
'*
'******************************************************************************
Public Sub SetTimeoutsImplicitWait(msec As Long)
    mDriver.Timeouts.ImplicitWait = msec
End Sub

'******************************************************************************
'* [概  要] SetArgumentMaximized メソッド
'* [詳  細] ウィンドウを最大化して表示する初期パラメータを指定する。
'*
'******************************************************************************
Public Sub SetArgumentMaximized()
    mDriver.AddArgument "--start-maximized"
End Sub

'******************************************************************************
'* [概  要] SetArgumentHeadless メソッド
'* [詳  細] ヘッドレスモードを指定する。
'* [参  考] https://vba-labo.rs-techdev.com/archives/1245
'*
'******************************************************************************
Public Sub SetArgumentHeadless()
    mDriver.AddArgument "headless"
End Sub

'******************************************************************************
'* [概  要] TakeDesktopScreenShot メソッド
'* [詳  細] デスクトップのスクリーンショットを取得する。
'* [参  考] SeleniumBasic Examples.xlsm
'*
'* @return Selenium.Image スクリーンショット画像オブジェクト
'******************************************************************************
Public Function TakeDesktopScreenShot() As Selenium.Image
  Set TakeDesktopScreenShot = Utils.TakeScreenshot()
End Function

'******************************************************************************
'* [概  要] SetWindowSize メソッド
'* [詳  細] ウィンドウのサイズを指定する。
'*
'******************************************************************************
Public Sub SetWindowSize(x As Long, y As Long)
    mDriver.Window.SetSize x, y
End Sub

'******************************************************************************
'* [概  要] ScrollElementCenter メソッド
'* [詳  細] 指定した要素が画面中央の位置になるようにスクロールする。
'* [参  考] SeleniumBasic Examples.xlsm
'*
'* @param elem WebElement
'******************************************************************************
Public Sub ScrollElementCenter(elem As Selenium.WebElement)
    elem.ExecuteScript "this.scrollIntoView(true);" & _
                       "window.scrollBy(0, -(window.innerHeight - this.clientHeight) / 2);"
End Sub

'******************************************************************************
'* [概  要] GetScrollHeight メソッド
'* [詳  細] スクロールの位置（縦位置）を取得する。
'* [参  考] https://vba-labo.rs-techdev.com/archives/1245
'*
'* @return Long スクロールの位置（縦位置）
'******************************************************************************
Public Function GetScrollHeight() As Long
    GetScrollHeight = mDriver.ExecuteScript("return document.body.scrollHeight;")
End Function
    
'******************************************************************************
'* [概  要] GetScrollWidth メソッド
'* [詳  細] スクロールの位置（横位置）を取得する。
'* [参  考] https://vba-labo.rs-techdev.com/archives/1245
'*
'* @return Long スクロールの位置（横位置）
'******************************************************************************
Public Function GetScrollWidth() As Long
    GetScrollWidth = mDriver.ExecuteScript("return document.body.scrollWidth;")
End Function

'******************************************************************************
'* [概  要] SetWindowWholeSize メソッド
'* [詳  細] ウィンドウのサイズをスクロール可能な最大サイズで指定する。
'* [参  考] https://vba-labo.rs-techdev.com/archives/1245
'*
'******************************************************************************
Public Sub SetWindowWholeSize()
    mDriver.SetSize GetScrollWidth(), GetScrollHeight()
End Sub
 
'******************************************************************************
'* [概  要] ScrollNextPageBySpace メソッド
'* [詳  細] スペースキーを送りページをスクロールする。
'* [参  考] https://www.seleniumqref.com/api/java/element_set/Java_special_sendKeys.html
'*
'******************************************************************************
Public Sub ScrollNextPageBySpace()
    mDriver.SendKeys Me.Keys.Space
End Sub

'******************************************************************************
'* [概  要] ScrollStart メソッド
'* [詳  細] ページの開始位置までスクロールする。
'*
'******************************************************************************
Public Sub ScrollStart()
    mDriver.ExecuteScript "window.scrollTo(0, 0);"
End Sub

'******************************************************************************
'* [概  要] ScrollEnd メソッド
'* [詳  細] ページの終了位置までスクロールする。
'*
'******************************************************************************
Public Sub ScrollEnd()
    mDriver.ExecuteScript "window.scrollTo(0, document.body.scrollHeight);"
End Sub

'******************************************************************************
'* [概  要] GetWindowXPosition メソッド
'* [詳  細] ウィンドウ横位置を取得する。
'*
'* @return Long ウィンドウ横位置
'******************************************************************************
Public Function GetWindowXPosition() As Long
    GetWindowXPosition = mDriver.ExecuteScript("return window.pageXOffset;")
End Function

'******************************************************************************
'* [概  要] GetWindowYPosition メソッド
'* [詳  細] ウィンドウ縦位置を取得する。
'*
'* @return Long ウィンドウ縦位置
'******************************************************************************
Public Function GetWindowYPosition() As Long
    GetWindowYPosition = mDriver.ExecuteScript("return window.pageYOffset;")
End Function

'******************************************************************************
'* [概  要] SendJWPNewSession メソッド
'* [詳  細] JsonWireProtocolを使用し、 NewSession する。
'* [参  考] https://w3c.github.io/webdriver/#new-session
'*
'* @return Selenium.List
'******************************************************************************
Public Function SendJWPNewSession() As Selenium.List
  Dim method As String: method = "POST"
  Dim uri As String: uri = "/session"

  Set SendJWPNewSession = mDriver.Send(method, uri)
End Function

'******************************************************************************
'* [概  要] SendJWPDeleteSession メソッド
'* [詳  細] JsonWireProtocolを使用し、 DeleteSession する。
'* [参  考] https://w3c.github.io/webdriver/#delete-session
'*
'* @param sessionId セッションID
'* @return Selenium.List
'******************************************************************************
Public Function SendJWPDeleteSession(sessionId As String) As Selenium.List
  Dim method As String: method = "DELETE"
  Dim uri As String: uri = "/session/{sessionId}"
  uri = Replace(uri, "{sessionId}", sessionId)

  Set SendJWPDeleteSession = mDriver.Send(method, uri)
End Function

'******************************************************************************
'* [概  要] SendJWPStatus メソッド
'* [詳  細] JsonWireProtocolを使用し、 Status する。
'* [参  考] https://w3c.github.io/webdriver/#status
'*
'* @return Selenium.List
'******************************************************************************
Public Function SendJWPStatus() As Selenium.List
  Dim method As String: method = "GET"
  Dim uri As String: uri = "/status"

  Set SendJWPStatus = mDriver.Send(method, uri)
End Function

'******************************************************************************
'* [概  要] SendJWPGetTimeouts メソッド
'* [詳  細] JsonWireProtocolを使用し、 GetTimeouts する。
'* [参  考] https://w3c.github.io/webdriver/#get-timeouts
'*
'* @param sessionId セッションID
'* @return Selenium.List
'******************************************************************************
Public Function SendJWPGetTimeouts(sessionId As String) As Selenium.List
  Dim method As String: method = "GET"
  Dim uri As String: uri = "/session/{sessionId}/timeouts"
  uri = Replace(uri, "{sessionId}", sessionId)
  
  Set SendJWPGetTimeouts = mDriver.Send(method, uri)
End Function


'******************************************************************************
'* [概  要] SendJWPGetHandle メソッド
'* [詳  細] JsonWireProtocolを使用し、Handleリストを取得する。
'* [参  考] SeleniumBasic Examples.xlsm
'*          https://w3c.github.io/webdriver/
'*
'* @return Selenium.List
'******************************************************************************
Public Function SendJWPGetHandle() As Selenium.List
  Dim method As String: method = "GET"
  Dim uri As String: uri = "/window_handles"

  Set SendJWPGetHandle = mDriver.Send(method, uri)
End Function

'******************************************************************************
'* [概  要] SendJWPGetLinkList メソッド
'* [詳  細] JsonWireProtocolを使用し、Linkリストを取得する。
'* [参  考] SeleniumBasic Examples.xlsm
'*          https://w3c.github.io/webdriver/
'*
'* @return Selenium.List
'******************************************************************************
Public Function SendJWPGetLinkList() As Selenium.List
  Dim method As String: method = "POST"
  Dim uri As String: uri = "/elements"

  Set SendJWPGetLinkList = mDriver.Send(method, uri, "using", "css selector", "value", "a")
End Function

'******************************************************************************
'* [概  要] CountTagElements メソッド
'* [詳  細] タグの件数をカウントする。
'*
'* @param tagname   タグ名
'* @return Long 指定したタグがページ内に存在する件数
'******************************************************************************
Public Function CountTagElements(tagName As String) As Long
    CountTagElements = mDriver.FindElementsByTag(tagName).count
End Function

'******************************************************************************
'* [概  要] ShowElements メソッド
'* [詳  細] 指定した条件またはWebElementsについて、詳細情報をイミディエイト
'*          ウィンドウに表示する。
'*
'* @param condOrElements 条件またはWebElements
'* @param searchType     検索タイプ
'* @param maxShowCount   最大表示件数
'* @param attributeName  属性名（検索タイプが SEARCH_ATTRIBUTE の場合に有効）
'******************************************************************************
Public Sub ShowElements(condOrElements As Variant, Optional searchType As SearchElemTypeEnum = SEARCH_TAG, Optional maxShowCount As Long = 0, Optional attributeName As String, Optional actualAttribute As Boolean = False)
    On Error Resume Next

    Dim elems As WebElements
    Dim searchCond As String
    If TypeName(condOrElements) = "String" Then
        Select Case searchType
            Case SearchElemTypeEnum.SEARCH_CLASS
                Set elems = mDriver.FindElementsByClass(condOrElements)
                searchCond = "[" & condOrElements & "] Class"
            Case SearchElemTypeEnum.SEARCH_CSS
                Set elems = mDriver.FindElementsByCss(condOrElements)
                searchCond = "[" & condOrElements & "] Css"
            Case SearchElemTypeEnum.SEARCH_ID
                Set elems = mDriver.FindElementsById(condOrElements)
                searchCond = "[" & condOrElements & "] Id"
            Case SearchElemTypeEnum.SEARCH_LINKTEXT
                Set elems = mDriver.FindElementsByLinkText(condOrElements)
                searchCond = "[" & condOrElements & "] LinkText"
            Case SearchElemTypeEnum.SEARCH_NAME
                Set elems = mDriver.FindElementsByName(condOrElements)
                searchCond = "[" & condOrElements & "] Name"
            Case SearchElemTypeEnum.SEARCH_PARTIALLINKTEXT
                Set elems = mDriver.FindElementsByPartialLinkText(condOrElements)
                searchCond = "[" & condOrElements & "] PartialLinkText"
            Case SearchElemTypeEnum.SEARCH_TAG
                Set elems = mDriver.FindElementsByTag(condOrElements)
                searchCond = "[" & condOrElements & "] Tag"
            Case SearchElemTypeEnum.SEARCH_XPATH
                Set elems = mDriver.FindElementsByXPath(condOrElements)
                searchCond = "[" & condOrElements & "] XPath"
            Case SearchElemTypeEnum.SEARCH_TEXT
                Set elems = FindElementsByAttributeUseXPath(Nothing, "*", "text", CStr(condOrElements))
                searchCond = "[" & condOrElements & "] Text"
            Case SearchElemTypeEnum.SEARCH_ATTRIBUTE
                Set elems = FindElementsByAttributeUseXPath(Nothing, "*", attributeName, CStr(condOrElements))
                searchCond = "[" & condOrElements & "] Attribute(" & attributeName & ")"
        End Select

    ElseIf TypeName(condOrElements) = "WebElements" Then
        Set elems = condOrElements
        searchCond = "[WebElements]"
    Else
        Exit Sub
    End If
    
    Debug.Print "--------------------------------------------------------------------------------"
    Debug.Print " WebElement リスト情報"
    Debug.Print "--------------------------------------------------------------------------------"
    Debug.Print "Web URL      : " & Me.url
    Debug.Print "検索条件     : " & searchCond
    If maxShowCount <> 0 Then Debug.Print "最大表示件数 : " & maxShowCount
    Debug.Print "検索件数     : " & elems.count
    Debug.Print
    
    Dim elem As WebElement, i As Long: i = 1
    For Each elem In elems
        Debug.Print "................................................................................"
        Debug.Print "No.", i
        Debug.Print elem.tagName, elem.text, elem.value
        
        Dim v, val
        If actualAttribute Then
            For Each v In GetElementAttributes(elem)
                Debug.Print v.key, v.value
            Next
        Else
            For Each v In GetTagAttributeList(elem)
                val = elem.Attribute(CStr(v))
                If val <> "" Then
                    Debug.Print CStr(v), val
                End If
            Next
        End If
        
        i = i + 1
        Debug.Print
        
        If maxShowCount <> 0 And maxShowCount < i Then
            Exit For
        End If
    Next
    Debug.Print "--------------------------------------------------------------------------------"
    Debug.Print
End Sub

'******************************************************************************
'* [概  要] ShowPerformanceTiming メソッド
'* [詳  細] Webページの性能測定用。表示時間（ページ応答時間、ページロード時間）
'*          をイミディエイトウィンドウに表示する。
'*
'******************************************************************************
Public Sub ShowPerformanceTiming()
    Debug.Print "--------------------------------------------------------------------------------"
    Debug.Print " Performance Timing"
    Debug.Print "--------------------------------------------------------------------------------"
    Dim perfList As Variant: perfList = GetPerformanceTiming
    Debug.Print "Web URL : " & Me.url
    Debug.Print
    Debug.Print "ページ応答時間  （responseEnd  - navigationStart） : " & Format(perfList(0), "#,###") & " ms"
    Debug.Print "ページロード時間（loadEventEnd - responseEnd）     : " & Format(perfList(1), "#,###") & " ms"
    Debug.Print "--------------------------------------------------------------------------------"
    Debug.Print
End Sub

'******************************************************************************
'* [概  要] ShowCapabilities メソッド
'* [詳  細] Capabilities の設定をイミディエイトウィンドウに表示する。
'*
'******************************************************************************
Public Sub ShowCapabilities()
    Debug.Print "--------------------------------------------------------------------------------"
    Debug.Print " Capabilities"
    Debug.Print "--------------------------------------------------------------------------------"
    Dim v As Selenium.DictionaryItem
    Dim s As String
    Dim maxLen As Long
    For Each v In Me.Capabilities
        If Len(v.key) > maxLen Then maxLen = Len(v.key)
    Next
    
    For Each v In Me.Capabilities
        If TypeName(v.value) = "Dictionary" Then
            s = "{"
            Dim vv As Selenium.DictionaryItem
            For Each vv In v.value
                s = s & vv.key & ": " & vv.value & ", "
            Next
            If v.value.count = 0 Then
                s = "(Nothing)"
            Else
                s = Mid(s, 1, Len(s) - 2) & "}"
            End If
            Debug.Print v.key & Space(maxLen - Len(v.key)) & " => " & s
        ElseIf IsArray(v.value) Then
            Debug.Print v.key & Space(maxLen - Len(v.key)) & " => " & "[" & Join(v.value, ", ") & "]"
        Else
            Debug.Print v.key & Space(maxLen - Len(v.key)) & " => " & v.value
        End If
    Next
    Debug.Print "--------------------------------------------------------------------------------"
    Debug.Print
End Sub

'******************************************************************************
'* [概  要] ShowTagCount メソッド
'* [詳  細] 指定したタグの件数をイミディエイトウィンドウに表示する。
'*
'* @param tagName タグ名
'* @param hideZero 0件の場合は表示しないかどうか
'******************************************************************************
Public Sub ShowTagCount(tagName As String, Optional hideZero As Boolean = False)
    Dim count As Long: count = CountTagElements(tagName)
    If Not (hideZero And count = 0) Then
        Debug.Print tagName, CountTagElements(tagName) & " 件"
    End If
End Sub

'******************************************************************************
'* [概  要] ShowTagCounts
'* [詳  細] ページ内で使用されている各タグの件数をイミディエイトウィンドウに
'*          表示する。
'*
'* @param tagNameList タグ名リスト
'* @param hideZero 0件の場合は表示しないかどうか
'******************************************************************************
Public Sub ShowTagCounts(Optional tagNameList As Variant, Optional hideZero As Boolean = True)
    Debug.Print "--------------------------------------------------------------------------------"
    Debug.Print " Tag 件数"
    Debug.Print "--------------------------------------------------------------------------------"
    
    If IsMissing(tagNameList) Then
        tagNameList = TagList
    End If
    Dim v
    For Each v In tagNameList
        ShowTagCount CStr(v), hideZero
    Next
    Debug.Print "--------------------------------------------------------------------------------"
    Debug.Print
End Sub
Public Property Get TagList() As Variant
    TagList = Split("html,head,title,base,link,style,meta,body,article,section,nav,aside,h1,h2,h3,h4,h5,h6,header,footer,address,p,hr,pre,blockquote,ol,ul,li,dl,dt,dd,figure,figcaption,main,div,a,em,strong,small,s,cite,q,dfn,abbr,code,var,samp,kbd,data,sub,sup,time,i,b,u,mark,ruby,rb,rt,rtc,rp,bdi,bdo,span,br,wbr,ins,del,img,picture,iframe,embed,object,param,video,audio,track,source,map,area,table,caption,colgroup,col,tbody,thead,tfoot,tr,td,th,form,fieldset,legend,label,input,select,option,optgroup,textarea,button,datalist,keygen,output,progress,meter,script,noscript,template,canvas,details,summary,menu,menuitem", ",")
End Property


'******************************************************************************
'* [概  要] GetTagAttributeList メソッド
'* [詳  細] 対象のタグ要素に対応する属性のリストを返却する。
'*
'* @param elem WebElement
'* @return Variant 属性リスト
'******************************************************************************
Public Function GetTagAttributeList(elem As WebElement) As Variant
    Select Case LCase(elem.tagName)
        Case "a"
            GetTagAttributeList = Split("id class href target hreflang type rel download rev name style", " ")
            Exit Function
        
        Case "input"
            GetTagAttributeList = Split("id class type name value disabled form autofocus style", " ")
            Exit Function
        
        Case "label"
            GetTagAttributeList = Split("id class for form style", " ")
            Exit Function

        Case "select"
            GetTagAttributeList = Split("id class name size multiple disabled form autofocus required style", " ")
            Exit Function

        Case "option"
            GetTagAttributeList = Split("id class value label selected disabled style", " ")
            Exit Function

        Case "optgroup"
            GetTagAttributeList = Split("id class label disabled style", " ")
            Exit Function

        Case "textarea"
            GetTagAttributeList = Split("id class name cols rows maxlength minlength wrap form autocomplete autofocus placeholder dirname inputmode readonly disabled required style", " ")
            Exit Function

        Case "button"
            GetTagAttributeList = Split("id class type name value disabled form autofocus menu formaction formmethod formenctype formtarget formnovalidate style", " ")
            Exit Function
    
        Case "base"
            GetTagAttributeList = Split("href target style", " ")
            Exit Function
            
        Case "link"
            GetTagAttributeList = Split("rel href hreflang type media sizes crossorigin rev title style", " ")
            Exit Function
            
        Case "style"
            GetTagAttributeList = Split("type media nonce title style", " ")
            Exit Function
        
        Case "meta"
            GetTagAttributeList = Split("charset http-equiv name property content style", " ")
            Exit Function
        
        Case "blockquote"
            GetTagAttributeList = Split("id class cite style", " ")
            Exit Function
        
        Case "ol"
            GetTagAttributeList = Split("id class type start reversed style", " ")
            Exit Function
        
        Case "li"
            GetTagAttributeList = Split("id class value style", " ")
            Exit Function
        
        Case "div"
            GetTagAttributeList = Split("id class lang style", " ")
            Exit Function
        
        Case "q"
            GetTagAttributeList = Split("id class cite style", " ")
            Exit Function
        
        Case "dfn"
            GetTagAttributeList = Split("id class title style", " ")
            Exit Function
        
        Case "abbr"
            GetTagAttributeList = Split("id class title style", " ")
            Exit Function
        
        Case "data"
            GetTagAttributeList = Split("id class value style", " ")
            Exit Function
        
        Case "time"
            GetTagAttributeList = Split("id class datetime style", " ")
            Exit Function
        
        Case "bdi"
            GetTagAttributeList = Split("id class dir style", " ")
            Exit Function
        
        Case "bdo"
            GetTagAttributeList = Split("id class dir style", " ")
            Exit Function
        
        Case "ins"
            GetTagAttributeList = Split("id class cite datetime style", " ")
            Exit Function
        
        Case "del"
            GetTagAttributeList = Split("id class cite datetime style", " ")
            Exit Function
        
        Case "img"
            GetTagAttributeList = Split("id class src alt width height usemap ismap crossorigin srcset sizes style", " ")
            Exit Function
        
        Case "iframe"
            GetTagAttributeList = Split("id class src width height name srcdoc sandbox allowfullscreen style", " ")
            Exit Function
        
        Case "embed"
            GetTagAttributeList = Split("id class src type width height style", " ")
            Exit Function
        
        Case "object"
            GetTagAttributeList = Split("id class data type name width height typemustmatch form style", " ")
            Exit Function
        
        Case "param"
            GetTagAttributeList = Split("id class name value style", " ")
            Exit Function
        
        Case "video"
            GetTagAttributeList = Split("id class src poster preload width height autoplay loop muted controls crossorigin style", " ")
            Exit Function
        
        Case "audio"
            GetTagAttributeList = Split("id class src preload autoplay loop muted controls crossorigin style", " ")
            Exit Function
        
        Case "track"
            GetTagAttributeList = Split("id class src kind srclang label default style", " ")
            Exit Function
        
        Case "source"
            GetTagAttributeList = Split("id class src type style", " ")
            Exit Function
        
        Case "map"
            GetTagAttributeList = Split("id class name style", " ")
            Exit Function
        
        Case "area"
            GetTagAttributeList = Split("id class shape coords href alt rel download target style", " ")
            Exit Function
        
        Case "table"
            GetTagAttributeList = Split("id class border style", " ")
            Exit Function
        
        Case "colgroup"
            GetTagAttributeList = Split("id class span style", " ")
            Exit Function
        
        Case "col"
            GetTagAttributeList = Split("id class span style", " ")
            Exit Function
        
        Case "td"
            GetTagAttributeList = Split("id class colspan rowspan headers style", " ")
            Exit Function
        
        Case "th"
            GetTagAttributeList = Split("id class colspan rowspan abbr headers scope style", " ")
            Exit Function
        
        Case "form"
            GetTagAttributeList = Split("id class name action method enctype accept-charset target autocomplete novalidate style", " ")
            Exit Function
        
        Case "fieldset"
            GetTagAttributeList = Split("id class form name disabled style", " ")
            Exit Function

        Case "datalist"
            GetTagAttributeList = Split("id class style", " ")
            Exit Function

        Case "output"
            GetTagAttributeList = Split("id class name for form style", " ")
            Exit Function

        Case "progress"
            GetTagAttributeList = Split("id class value max style", " ")
            Exit Function

        Case "meter"
            GetTagAttributeList = Split("id class value min max low high optimum style", " ")
            Exit Function

        Case "script"
            GetTagAttributeList = Split("id class src type crossorigin nonce charset async defer language", " ")
            Exit Function

        Case "canvas"
            GetTagAttributeList = Split("id class width height style", " ")
            Exit Function

        Case "details"
            GetTagAttributeList = Split("id class open style", " ")
            Exit Function
        
        Case "menu"
            GetTagAttributeList = Split("id class type label style", " ")
            Exit Function

        Case "menuitem"
            GetTagAttributeList = Split("id class type label icon radiogroup checked disabled default title", " ")
            Exit Function
       
        Case Else
            'あまり使用されない共通属性は返却しないこととした
            'GetTagAttributeList = Split("id class style accesskey autocapitalize contenteditable contextmenu dir draggable dropzone hidden lang spellcheck tabindex title translate", " ")
            GetTagAttributeList = Split("id class lang style", " ")
            Exit Function
    End Select
End Function

''
' Saves the file pointed by the href attribute : <a href="/doc.pdf">Document</a>
' @element {WebElement} Web element with the href link
' @save_as {String}  Path were the file is to be saved
''
Private Sub Download_StaticLink(element As WebElement, save_as As String)
    ' Extract the data to build the request (link, user-agent, language, cookie)
    Dim info As Selenium.Dictionary
    Set info = element.ExecuteScript("return {" & _
                                                "link: this.href," & _
                                                "agent: navigator.userAgent," & _
                                                "lang: navigator.userLanguage," & _
                                                "cookie: document.cookie };")
  
    ' Send the request
    Static xhr As Object
    If xhr Is Nothing Then Set xhr = CreateObject("Msxml2.ServerXMLHTTP.6.0")
    xhr.Open "GET", info("link")
    xhr.setRequestHeader "User-Agent", info("agent")
    xhr.setRequestHeader "Accept-Language", info("lang")
    xhr.setRequestHeader "Cookie", info("cookie")
    xhr.Send
    If (xhr.Status \ 100) - 2 Then Err.raise 5, , xhr.Status & " " & xhr.StatusText
  
    ' Save the response to a file
    Static bin As Object
    If bin Is Nothing Then Set bin = CreateObject("ADODB.Stream")
    If Len(Dir$(save_as)) Then Kill save_as
    bin.Open
    bin.Type = 1
    bin.Write xhr.ResponseBody
    bin.Position = 0
    bin.SaveToFile save_as
    bin.Close
End Sub

''
' Waits for a new file to be created in a folder
' @folder {String}  Folder where the file will be created
' Usage:
'   WaitNewFile "C:\download\*.pdf"
'   ' The new file is created here
'   filename = WaitNewFile()
''
Private Function WaitNewFile(Optional target As String) As String
  Static files As Collection, filter$
  Dim file$, file_path$, i&
  If Len(target) Then
    ' Initialize the list of files and return
    filter = target
    Set files = New Collection
    file = Dir(filter, vbNormal)
    Do While Len(file)
      files.Add Empty, file
      file = Dir
    Loop
    Exit Function
  End If
  
  ' Waits for a file that is not in the list
  On Error GoTo WaitReady
  Do
    file = Dir(filter, vbNormal)
    Do While Len(file)
      files.Item file
      file = Dir
    Loop
    For i = 0 To 3000: DoEvents: Next
  Loop
  
WaitReady:
  ' Waits for the size to be superior to 0 and try to rename it
  file_path = Left$(filter, InStrRev(filter, "\")) & file
  Do
    If FileLen(file_path) Then
      On Error Resume Next
      Name file_path As file_path
      If Err = 0 Then Exit Do
    End If
    For i = 0 To 3000: DoEvents: Next
  Loop
  files.Add Empty, file
  WaitNewFile = file_path
End Function

''
' Drop data on an element that implements HTML5 drop
' @target: web element that will receive the data
' @dataType: type of data
' @data: data to drop on the target
''
Public Sub DropDataHTML5(target As WebElement, dataType As String, data)
  Const JS_DropText As String = _
    "var t=this,m=arguments[0],v=arguments[1],d={dropEffect:'',types:[m],getData:funct" & _
    "ion(k){return v;}},f=function(e,k){var u=document.createEvent('Event');u.initEven" & _
    "t(k,1,1);u.dataTransfer=d;e.dispatchEvent(u);};f(t,'dragenter');f(t,'dragover');f" & _
    "(t,'drop');"
  
  target.ExecuteScript JS_DropText, Array(dataType, data)
End Sub

''
' Perform a drag and drop on elements that implement HTML5 drag an drop
' @source web element to drag
' @target: web element to drop on
''
Public Sub DragAndDropHTML5(source As WebElement, target As WebElement)
  Const JS_DnD As String = _
    "var s=this,t=arguments[0],d={dropEffect:'',types:[],setData:function(k,v){this[k]=" & _
    "v;types.append(k);},getData:function(k){return this[k]}},f=function(e,k){var u=doc" & _
    "ument.createEvent('Event');u.initEvent(k,1,1);u.dataTransfer=d;e.dispatchEvent(u);" & _
    "};f(s,'dragstart');f(t,'dragenter');f(t,'dragover');f(t,'drop');f(s,'dragend');"

  source.ExecuteScript JS_DnD, target
End Sub

''
' Perform a drag and drop on elements that implement HTML5 drag an drop
' @target: web element that will receive the file
' @filepath: file path
''
Public Sub DropFileHTML5(target As WebElement, filePath As String)
  Const JS_NewInput = _
    "var e=document.createElement('input');e.type='file';" & _
    "document.body.appendChild(e);return e;"
  Const JS_DropFile = _
    "var s=this,t=arguments[0],d={dropEffect:'',files:s.files},f=function(e,k){u=docum" & _
    "ent.createEvent('Event');u.initEvent(k,1,1);u.dataTransfer=d;e.dispatchEvent(u);}" & _
    ";f(t,'dragenter');f(t,'dragover');f(t,'drop');"
  
  Dim source As WebElement
  Set source = target.ExecuteScript(JS_NewInput)
  source.SendKeys filePath
  source.ExecuteScript JS_DropFile, target
End Sub

''
' Finds all the scrollable ancestor and scroll them vertically by the provided amout of pixels
' @element {WebElement}  Web element in a scrollable container or window
' @y {Long} Amount of pixels to vertically scroll
''
Public Function ScrollVertically(element As WebElement, y As Long)
  Const JS_SCROLL_Y = _
    "var y = arguments[0];" & _
    "for (var e=this; e; e=e.parentElement) {" & _
    "  var yy = e.scrollHeight - e.clientHeight;" & _
    "  if (yy === 0) continue;" & _
    "  yy = y < 0 ? Math.max(y, -e.scrollTop) : Math.min(y, yy - e.scrollTop);" & _
    "  if(yy === 0) continue;" & _
    "  e.scrollTop += yy;" & _
    "  if ((y -= yy) == 0) return;" & _
    "}" & _
    "window.scrollBy(0, y);"
    
  element.ExecuteScript JS_SCROLL_Y, y
End Function

''
' Scrolls an element in the center of the view
' @element {WebElement}  Web element in a scrollable container
''
Public Function ScrollIntoViewCenter(element As WebElement)
  Const JS_SCROLL_CENTER = _
    "this.scrollIntoView(true);" & _
    "var y = (window.innerHeight - this.offsetHeight) / 2;" & _
    "if (y < 1) return;" & _
    "for (var e=this; e; e=e.parentElement) {" & _
    "  if (e.scrollTop == 0) continue;" & _
    "  var yy = Math.min(e.scrollTop, y);" & _
    "  e.scrollTop -= yy;" & _
    "  if ((y -= yy) < 1) return;" & _
    "}" & _
    "window.scrollBy(0, -y);"
  
  element.ExecuteScript JS_SCROLL_CENTER
End Function

'******************************************************************************
'* [概  要] UnZipFile
'* [詳  細] zipファイルを解凍する。
'*
'* @param zipFilePath zipファイルパス
'* @param targetFolderPath 解凍先フォルダパス
'*                         （解凍先フォルダにzipファイルのベース名と同じ名称
'*                           のフォルダを作成）
'* @return Boolean 実行結果
'******************************************************************************
Private Function UnZipFile(zipFilePath As String, targetFolderPath As String) As Boolean

    If Not mFso.FileExists(zipFilePath) Then
        UnZipFile = False
        Exit Function
    End If
    Dim saveFolderPath As String
    saveFolderPath = mFso.BuildPath(targetFolderPath, mFso.GetBaseName(Mid(zipFilePath, 1, InStrRev(zipFilePath, ".") - 1)))
    Call CreateFolderTree(saveFolderPath)
    
    Dim shellObj As Object: Set shellObj = CreateObject("Shell.Application")
    Dim targetFolder As Object: Set targetFolder = shellObj.Namespace(mFso.GetAbsolutePathName(saveFolderPath))
    Dim srcFolder As Object: Set srcFolder = shellObj.Namespace(mFso.GetAbsolutePathName(zipFilePath))
    Dim cnt As Long: cnt = srcFolder.items().count
    Dim retryCnt As Long: retryCnt = 0
    
    targetFolder.CopyHere srcFolder.items, &H4 Or &H10
    Do While targetFolder.items().count < cnt
        If retryCnt > 10 Then
            Debug.Print "RetryOver"
            Exit Function
        End If
        Application.Wait Now + TimeSerial(0, 0, 1)
        retryCnt = retryCnt + 1
    Loop
    UnZipFile = True
End Function

'******************************************************************************
'* [概  要] CreateFolderTree メソッド
'* [詳  細] フォルダパスを作成する。（存在しない場合は親フォルダ含めて作成）
'* [参  考] Ariawase 0.9.0 IO.bas
'*
'* @param folderPath 作成するフォルダパス
'******************************************************************************
Private Sub CreateFolderTree(ByVal folderPath As String)
    If Not mFso.DriveExists(mFso.GetDriveName(folderPath)) Then Err.raise 5
    CreateFolderTreeImpl folderPath
End Sub

Private Sub CreateFolderTreeImpl(ByVal folderPath As String)
    If mFso.FolderExists(folderPath) Then GoTo Escape
    CreateFolderTreeImpl mFso.GetParentFolderName(folderPath)
    mFso.CreateFolder folderPath
Escape:
End Sub

'******************************************************************************
'* [概  要] CreateStdRegProv メソッド
'* [詳  細] StdRegProvオブジェクトを取得する。
'* [参  考] Ariawase 0.9.0 Core.bas
'*
'* @return Object StdRegProvオブジェクト
'******************************************************************************
Private Function CreateStdRegProv() As Object
    Dim wmi As Object: Set wmi = CreateObject("WbemScripting.SWbemLocator")
    Dim wmiSrv As Object: Set wmiSrv = wmi.ConnectServer(, "root\default")
    Set CreateStdRegProv = wmiSrv.Get("StdRegProv")
End Function

'******************************************************************************
'* [概  要] GetWmiServer メソッド
'* [詳  細] WmiServerオブジェクトを取得する。
'*
'* @return Object WmiServerオブジェクト
'******************************************************************************
Private Function GetWmiServer() As Object
    Dim wmi As Object: Set wmi = CreateObject("WbemScripting.SWbemLocator")
    Set GetWmiServer = wmi.ConnectServer()
End Function

