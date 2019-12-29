VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDOExMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] CDO.Messageクラスのラップ・拡張クラス
'* [詳  細] CDO.MessageクラスのWrapperとしての機能を提供する他、
'*          CDO.Messageクラスを使用したユーティリティを提供する。
'*
'* [参  考] <xx>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' CDO.Messageクラスについて。
' <xx>
' Defines an object used to manage a message.
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As CDO.Message
Private mOrigin As Object

'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* オリジナルObject プロパティ
'*-----------------------------------------------------------------------------
Property Get Origin() As Object
    Set Origin = mOrigin
End Property
Property Set Origin(arg As Object)
    Set mOrigin = arg
End Property

'*-----------------------------------------------------------------------------
'* Attachments プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The Object 's Attachments collection.
'*-----------------------------------------------------------------------------
Property Get Attachments() As CDOExIBodyParts
    Dim wrap As CDOExIBodyParts: Set wrap = New CDOExIBodyParts
    Set wrap.Origin = mOrigin.Attachments()
    Set Attachments = wrap
End Property

'*-----------------------------------------------------------------------------
'* AutoGenerateTextBody プロパティ
'*
'* [補  足]
'* ・Indicates whether a text/plain alternate representation should automatically be generated from the text/html part of the message body.
'*-----------------------------------------------------------------------------
Property Get AutoGenerateTextBody() As Boolean
    AutoGenerateTextBody = mOrigin.AutoGenerateTextBody
End Property
Property Let AutoGenerateTextBody(arg As Boolean)
    mOrigin.AutoGenerateTextBody = arg
End Property
    
'*-----------------------------------------------------------------------------
'* BCC プロパティ
'*
'* [補  足]
'* ・The Message 's hidden carbon copy (BCC header) recipients.
'*-----------------------------------------------------------------------------
Property Get BCC() As String
    BCC = mOrigin.BCC
End Property
Property Let BCC(arg As String)
    mOrigin.BCC = arg
End Property

'*-----------------------------------------------------------------------------
'* BodyPart プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The Object 's IBodyPart interface.
'*-----------------------------------------------------------------------------
Property Get BodyPart() As CDOExIBodyPart
    Dim wrap As CDOExIBodyPart: Set wrap = New CDOExIBodyPart
    Set wrap.Origin = mOrigin.BodyPart()
    Set BodyPart = wrap
End Property

'*-----------------------------------------------------------------------------
'* CC プロパティ
'*
'* [補  足]
'* ・The Message 's secondary (CC header) recipients.
'*-----------------------------------------------------------------------------
Property Get CC() As String
    CC = mOrigin.CC
End Property
Property Let CC(arg As String)
    mOrigin.CC = arg
End Property
    
'*-----------------------------------------------------------------------------
'* Configuration プロパティ
'*
'* [補  足]
'* ・The Object 's Configuration object.
'*-----------------------------------------------------------------------------
Property Get Configuration() As CDOExConfiguration
    Dim wrap As CDOExConfiguration: Set wrap = New CDOExConfiguration
    Set wrap.Origin = mOrigin.Configuration()
    Set Configuration = wrap
End Property
Property Set Configuration(arg As CDOExConfiguration)
    Set mOrigin.Configuration = arg.Origin
End Property

'*-----------------------------------------------------------------------------
'* DataSource プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The Object 's IDataSource interface.
'*-----------------------------------------------------------------------------
Property Get DataSource() As CDOExIDataSource
    Dim wrap As CDOExIDataSource: Set wrap = New CDOExIDataSource
    Set wrap.Origin = mOrigin.DataSource()
    Set DataSource = wrap
End Property

'*-----------------------------------------------------------------------------
'* DSNOptions プロパティ
'*
'* [補  足]
'* ・The delivery status notification (DSN) options for the message.
'*-----------------------------------------------------------------------------
Property Get DSNOptions() As CdoDSNOptions
    DSNOptions = mOrigin.DSNOptions
End Property
Property Let DSNOptions(arg As CdoDSNOptions)
    mOrigin.DSNOptions = arg
End Property
   
'*-----------------------------------------------------------------------------
'* EnvelopeFields プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The transport envelope Fields collection for the message (transport event sinks only).
'*-----------------------------------------------------------------------------
Property Get EnvelopeFields() As ADODBExFields
    Dim wrap As ADODBExFields: Set wrap = New ADODBExFields
    Set wrap.Origin = mOrigin.EnvelopeFields()
    Set EnvelopeFields = wrap
End Property

'*-----------------------------------------------------------------------------
'* Fields プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The Object 's Fields collection.
'*-----------------------------------------------------------------------------
Property Get Fields() As ADODBExFields
    Dim wrap As ADODBExFields: Set wrap = New ADODBExFields
    Set wrap.Origin = mOrigin.Fields()
    Set Fields = wrap
End Property

'*-----------------------------------------------------------------------------
'* FollowUpTo プロパティ
'*
'* [補  足]
'* ・The Message 's follow-up recipients.
'*-----------------------------------------------------------------------------
Property Get FollowUpTo() As String
    FollowUpTo = mOrigin.FollowUpTo
End Property
Property Let FollowUpTo(arg As String)
    mOrigin.FollowUpTo = arg
End Property

'*-----------------------------------------------------------------------------
'* From プロパティ
'*
'* [補  足]
'* ・The Message 's principle (From header) authors.
'*-----------------------------------------------------------------------------
Property Get From() As String
    From = mOrigin.From
End Property
Property Let From(arg As String)
    mOrigin.From = arg
End Property

'*-----------------------------------------------------------------------------
'* HTMLBody プロパティ
'*
'* [補  足]
'* ・The text/html portion of the message body.
'*-----------------------------------------------------------------------------
Property Get HTMLBody() As String
    HTMLBody = mOrigin.HTMLBody
End Property
Property Let HTMLBody(arg As String)
    mOrigin.HTMLBody = arg
End Property

'*-----------------------------------------------------------------------------
'* HTMLBodyPart プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Returns the BodyPart object (IBodyPart interface) containing the text/html portion of the message body.
'*-----------------------------------------------------------------------------
Property Get HTMLBodyPart() As CDOExIBodyPart
    Dim wrap As CDOExIBodyPart: Set wrap = New CDOExIBodyPart
    Set wrap.Origin = mOrigin.HTMLBodyPart()
    Set HTMLBodyPart = wrap
End Property

'*-----------------------------------------------------------------------------
'* Keywords プロパティ
'*
'* [補  足]
'* ・The Message 's keywords.
'*-----------------------------------------------------------------------------
Property Get Keywords() As String
    Keywords = mOrigin.Keywords
End Property
Property Let Keywords(arg As String)
    mOrigin.Keywords = arg
End Property

'*-----------------------------------------------------------------------------
'* MDNRequested プロパティ
'*
'* [補  足]
'* ・Indicates whether a mail delivery notification (MDN) should be sent when the message is received.
'*-----------------------------------------------------------------------------
Property Get MDNRequested() As Boolean
    MDNRequested = mOrigin.MDNRequested
End Property
Property Let MDNRequested(arg As Boolean)
    mOrigin.MDNRequested = arg
End Property

'*-----------------------------------------------------------------------------
'* MimeFormatted プロパティ
'*
'* [補  足]
'* ・Indicates whether the message is to be serialized using the MIME (True) or UUENCODE (False) format.
'*-----------------------------------------------------------------------------
Property Get MimeFormatted() As Boolean
    MimeFormatted = mOrigin.MimeFormatted
End Property
Property Let MimeFormatted(arg As Boolean)
    mOrigin.MimeFormatted = arg
End Property

'*-----------------------------------------------------------------------------
'* Newsgroups プロパティ
'*
'* [補  足]
'* ・The Message 's newsgroup (Newsgroups header) recipients.
'*-----------------------------------------------------------------------------
Property Get Newsgroups() As String
    Newsgroups = mOrigin.Newsgroups
End Property
Property Let Newsgroups(arg As String)
    mOrigin.Newsgroups = arg
End Property

'*-----------------------------------------------------------------------------
'* Organization プロパティ
'*
'* [補  足]
'* ・The Sender 's organization name.
'*-----------------------------------------------------------------------------
Property Get Organization() As String
    Organization = mOrigin.Organization
End Property
Property Let Organization(arg As String)
    mOrigin.Organization = arg
End Property

'*-----------------------------------------------------------------------------
'* ReceivedTime プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The date and time the message was received.
'*-----------------------------------------------------------------------------
Property Get ReceivedTime() As Date
    ReceivedTime = mOrigin.ReceivedTime
End Property

'*-----------------------------------------------------------------------------
'* ReplyTo プロパティ
'*
'* [補  足]
'* ・The email addresses (Reply-To header) to which to reply.
'*-----------------------------------------------------------------------------
Property Get ReplyTo() As String
    ReplyTo = mOrigin.ReplyTo
End Property
Property Let ReplyTo(arg As String)
    mOrigin.ReplyTo = arg
End Property

'*-----------------------------------------------------------------------------
'* Sender プロパティ
'*
'* [補  足]
'* ・The Message 's actual sender.
'*-----------------------------------------------------------------------------
Property Get Sender() As String
    Sender = mOrigin.Sender
End Property
Property Let Sender(arg As String)
    mOrigin.Sender = arg
End Property

'*-----------------------------------------------------------------------------
'* SentOn プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The date and time the message was sent.
'*-----------------------------------------------------------------------------
Property Get SentOn() As Date
    SentOn = mOrigin.SentOn
End Property

'*-----------------------------------------------------------------------------
'* Subject プロパティ
'*
'* [補  足]
'* ・The Message 's subject (Subject header).
'*-----------------------------------------------------------------------------
Property Get Subject() As String
    Subject = mOrigin.Subject
End Property
Property Let Subject(arg As String)
    mOrigin.Subject = arg
End Property
    
'*-----------------------------------------------------------------------------
'* TextBody プロパティ
'*
'* [補  足]
'* ・The text/plain portion of the message body.
'*-----------------------------------------------------------------------------
Property Get TextBody() As String
    TextBody = mOrigin.TextBody
End Property
Property Let TextBody(arg As String)
    mOrigin.TextBody = arg
End Property

'*-----------------------------------------------------------------------------
'* TextBodyPart プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Returns the BodyPart object (IBodyPart interface) containing the text/plain part of the message body.
'*-----------------------------------------------------------------------------
Property Get TextBodyPart() As IBodyPart
    Dim wrap As CDOExIBodyPart: Set wrap = New CDOExIBodyPart
    Set wrap.Origin = mOrigin.TextBodyPart()
    Set TextBodyPart = wrap
End Property

'*-----------------------------------------------------------------------------
'* To_ プロパティ
'*
'* [補  足]
'* ・The Message 's principle (To header) recipients.
'*-----------------------------------------------------------------------------
Property Get To_() As String
    To_ = mOrigin.To
End Property
Property Let To_(arg As String)
    mOrigin.To = arg
End Property

    
'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("CDO.Message")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] AddAttachment メソッド
'* [詳  細] AddAttachment のラッパーメソッド。
'*          Adds an attachment (BodyPart) to the message.
'*
'* @param Msg
'* @return コレクションの要素数
'*
'******************************************************************************
Function AddAttachment(URL As String, Optional UserName As String, Optional Password As String) As CDOExIBodyPart
    Dim wrap As CDOExIBodyPart: Set wrap = New CDOExIBodyPart
    Set wrap.Origin = mOrigin.AddAttachment(URL, UserName, Password)
    Set AddAttachment = wrap
End Function

'******************************************************************************
'* [概  要] AddRelatedBodyPart メソッド
'* [詳  細] AddRelatedBodyPart のラッパーメソッド。
'*          Adds a BodyPart object with content referenced within the text/html portion of the message body.
'*
'* @param URL
'* @param Reference
'* @param ReferenceType
'* @param UserName
'* @param Password
'* @return IBodyPart
'*
'******************************************************************************
Function AddRelatedBodyPart(URL As String, Reference As String, ReferenceType As CdoReferenceType, Optional UserName As String, Optional Password As String) As CDOExIBodyPart
    Dim wrap As CDOExIBodyPart: Set wrap = New CDOExIBodyPart
    Set wrap.Origin = mOrigin.AddRelatedBodyPart(URL, Reference, ReferenceType, UserName, Password)
    Set AddRelatedBodyPart = wrap
End Function

'******************************************************************************
'* [概  要] CreateMHTMLBody メソッド
'* [詳  細] CreateMHTMLBody のラッパーメソッド。
'*          Creates an MHTML-formatted message body using the resource(s) at the specified URL.
'*
'* @param URL
'* @param Flags
'* @param UserName
'* @param Password
'*
'******************************************************************************
Sub CreateMHTMLBody(URL As String, Optional Flags As CdoMHTMLFlags = cdoSuppressNone, Optional UserName As String, Optional Password As String)
    Call mOrigin.CreateMHTMLBody(URL, Flags, UserName, Password)
End Sub

'******************************************************************************
'* [概  要] Forward メソッド
'* [詳  細] Forward のラッパーメソッド。
'*          Returns a Message object used to forward a message.
'*
'* @return Message
'*
'******************************************************************************
Function Forward() As CDOExMessage
    Dim wrap As CDOExMessage: Set wrap = New CDOExMessage
    Set wrap.Origin = mOrigin.Forward()
    Set Forward = wrap
End Function

'******************************************************************************
'* [概  要] GetInterface メソッド
'* [詳  細] GetInterface のラッパーメソッド。
'*          Returns a specified interface on this object; provided for script languages.
'*
'* @param Interface
'* @return Object
'*
'******************************************************************************
Function GetInterface(Interface As String) As Object
    Set GetInterface = mOrigin.GetInterface(Interface)
End Function

'******************************************************************************
'* [概  要] GetStream メソッド
'* [詳  細] GetStream のラッパーメソッド。
'*          Returns an ADO Stream object containing the message in serialized,
'*          RFC 822 format.
'*          The message body is encoded using either MIME or UUENCODE as
'*          specified by the MIMEFormatted property.
'*
'* @return stream
'*
'******************************************************************************
Function GetStream() As ADODBExStream
    Dim wrap As ADODBExStream: Set wrap = New ADODBExStream
    Set wrap.Origin = mOrigin.GetStream()
    Set GetStream = wrap
End Function

'******************************************************************************
'* [概  要] Post メソッド
'* [詳  細] Post のラッパーメソッド。
'*          Posts the message using the method specified in the associated Configuration object.
'*
'******************************************************************************
Sub Post()
    Call mOrigin.Post
End Sub


'******************************************************************************
'* [概  要] PostReply メソッド
'* [詳  細] PostReply のラッパーメソッド。
'*          Returns a Message object used to post a reply to the message.
'*
'* @return Message
'*
'******************************************************************************
Function PostReply() As CDOExMessage
    Dim wrap As CDOExMessage: Set wrap = New CDOExMessage
    Set wrap.Origin = mOrigin.PostReply()
    Set PostReply = wrap
End Function

'******************************************************************************
'* [概  要] Reply メソッド
'* [詳  細] Reply のラッパーメソッド。
'*          Returns a Message object used to reply to the message.
'*
'* @return Message
'*
'******************************************************************************
Function Reply() As CDOExMessage
    Dim wrap As CDOExMessage: Set wrap = New CDOExMessage
    Set wrap.Origin = mOrigin.Reply()
    Set Reply = wrap
End Function

'******************************************************************************
'* [概  要] ReplyAll メソッド
'* [詳  細] ReplyAll のラッパーメソッド。
'*          Returns a Message object used to post a reply to all recipients of the message.
'*
'* @return Message
'*
'******************************************************************************
Function ReplyAll() As Message
    Dim wrap As CDOExMessage: Set wrap = New CDOExMessage
    Set wrap.Origin = mOrigin.ReplyAll()
    Set ReplyAll = wrap
End Function

'******************************************************************************
'* [概  要] Send メソッド
'* [詳  細] Send のラッパーメソッド。
'*          Sends the message using the method specified in the associated Configuration object.
'*
'******************************************************************************
Sub Send()
    Call mOrigin.Send
End Sub

'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------
