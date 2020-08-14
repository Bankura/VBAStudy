VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDOExIBodyPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] CDO.IBodyPartクラスのラップ・拡張クラス
'* [詳  細] CDO.IBodyPartクラスのWrapperとしての機能を提供する他、
'*          CDO.IBodyPartクラスを使用したユーティリティを提供する。
'*
'* [参  考] <xx>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' CDO.IBodyPartクラスについて。
' <xx>
' Defines methods, properties, and collections used to manage a message body part.
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As CDO.IBodyPart
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
'* Count プロパティ（読み取り専用）
'*
'* [補  足]
'* ・The Object 's BodyParts collection.
'*-----------------------------------------------------------------------------
Property Get BodyParts() As CDOExIBodyParts
    Dim wrap As CDOExIBodyParts: Set wrap = New CDOExIBodyParts
    Set wrap.Origin = mOrigin.BodyParts()
    Set BodyParts = wrap
End Property

'*-----------------------------------------------------------------------------
'* Charset プロパティ
'*
'* [補  足]
'* The character set of the body part's text content (not applicable for non-text content types).
'*-----------------------------------------------------------------------------
Property Get CharSet() As String
    CharSet = mOrigin.CharSet
End Property
Property Let CharSet(arg As String)
    mOrigin.CharSet = arg
End Property

'*-----------------------------------------------------------------------------
'* ContentClass プロパティ
'*
'* [補  足]
'* The body part's content class.
'*-----------------------------------------------------------------------------
Property Get ContentClass() As String
    ContentClass = mOrigin.ContentClass
End Property
Property Let ContentClass(arg As String)
    mOrigin.ContentClass = arg
End Property

'*-----------------------------------------------------------------------------
'* ContentClassName プロパティ
'*
'* [補  足]
'* Deprecated. Do not use.
'*-----------------------------------------------------------------------------
Property Get ContentClassName() As String
    ContentClassName = mOrigin.ContentClassName
End Property
Property Let ContentClassName(arg As String)
    mOrigin.ContentClassName = arg
End Property

'*-----------------------------------------------------------------------------
'* ContentMediaType プロパティ
'*
'* [補  足]
'* The content media type portion of the body part's content type.
'*-----------------------------------------------------------------------------
Property Get ContentMediaType() As String
    ContentMediaType = mOrigin.ContentMediaType
End Property
Property Let ContentMediaType(arg As String)
    mOrigin.ContentMediaType = arg
End Property

'*-----------------------------------------------------------------------------
'* ContentTransferEncoding プロパティ
'*
'* [補  足]
'* The method used to encode the body part content. For example, quoted-printable or base64.
'*-----------------------------------------------------------------------------
Property Get ContentTransferEncoding() As String
    ContentTransferEncoding = mOrigin.ContentTransferEncoding
End Property
Property Let ContentTransferEncoding(arg As String)
    mOrigin.ContentTransferEncoding = arg
End Property

'*-----------------------------------------------------------------------------
'* DataSource プロパティ（読み取り専用）
'*
'* [補  足]
'* The Object 's IDataSource interface.
'*-----------------------------------------------------------------------------
Property Get DataSource() As CDOExIDataSource
    Dim wrap As CDOExIDataSource: Set wrap = New CDOExIDataSource
    Set wrap.Origin = mOrigin.DataSource()
    Set DataSource = wrap
End Property

'*-----------------------------------------------------------------------------
'* Fields プロパティ（読み取り専用）
'*
'* [補  足]
'* The Object 's Fields collection.
'*-----------------------------------------------------------------------------
Property Get Fields() As ADODBExFields
    Dim wrap As ADODBExFields: Set wrap = New ADODBExFields
    Set wrap.Origin = mOrigin.Fields()
    Set Fields = wrap
End Property
    
'*-----------------------------------------------------------------------------
'* FileName プロパティ（読み取り専用）
'*
'* [補  足]
'* The value of the filename parameter for the content-disposition MIME header.
'*-----------------------------------------------------------------------------
Property Get FileName() As String
    FileName = mOrigin.FileName
End Property

'*-----------------------------------------------------------------------------
'* Parent プロパティ（読み取り専用）
'*
'* [補  足]
'* The Object 's parent object in the body part hierarchy.
'*-----------------------------------------------------------------------------
Property Get Parent() As CDOExIBodyPart
    Dim wrap As CDOExIBodyPart: Set wrap = New CDOExIBodyPart
    Set wrap.Origin = mOrigin.Parent()
    Set Parent = wrap
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("CDO.IBodyPart")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] AddBodyPart メソッド
'* [詳  細] AddBodyPart のラッパーメソッド。
'*          Adds a body part to the object's BodyParts collection.
'*
'* @param Index xx
'* @return IBodyPart
'******************************************************************************
Function AddBodyPart(Optional Index As Long = -1) As CDOExIBodyPart
    Dim wrap As CDOExIBodyPart: Set wrap = New CDOExIBodyPart
    Set wrap.Origin = mOrigin.AddBodyPart(Index)
    Set AddBodyPart = wrap
End Function

'******************************************************************************
'* [概  要] GetDecodedContentStream メソッド
'* [詳  細] GetDecodedContentStream のラッパーメソッド。
'*          Returns a Stream object containing the body part content in decoded format.
'*
'* @return stream
'******************************************************************************
Function GetDecodedContentStream() As ADODBExStream
    Dim wrap As ADODBExStream: Set wrap = New ADODBExStream
    Set wrap.Origin = mOrigin.GetDecodedContentStream()
    Set GetDecodedContentStream = wrap
End Function

'******************************************************************************
'* [概  要] GetEncodedContentStream メソッド
'* [詳  細] GetEncodedContentStream のラッパーメソッド。
'*          Returns a Stream object containing the body part content in encoded format. The encoding method is specified in the ContentTransferEncoding property.
'*
'* @return stream
'******************************************************************************
Function GetEncodedContentStream() As ADODBExStream
    Dim wrap As ADODBExStream: Set wrap = New ADODBExStream
    Set wrap.Origin = mOrigin.GetEncodedContentStream()
    Set GetEncodedContentStream = wrap
End Function

    
'******************************************************************************
'* [概  要] GetFieldParameter メソッド
'* [詳  細] GetFieldParameter のラッパーメソッド。
'*          Returns the specified parameter from the body part's specified header field.
'*
'* @param FieldName xx
'* @param Parameter xx
'* @return String
'******************************************************************************
Function GetFieldParameter(FieldName As String, Parameter As String) As String
    GetFieldParameter = mOrigin.GetFieldParameter(FieldName, Parameter)
End Function

'******************************************************************************
'* [概  要] GetInterface メソッド
'* [詳  細] GetInterface のラッパーメソッド。
'*          Returns a specified interface on this object; provided for script languages.
'*
'* @param Interface xx
'* @return Object
'******************************************************************************
Function GetInterface(Interface As String) As Object
    Set GetInterface = mOrigin.GetInterface(Interface)
End Function

    
'******************************************************************************
'* [概  要] GetStream メソッド
'* [詳  細] GetStream のラッパーメソッド。
'*          Returns an ADO Stream object containing the body part in serialized, MIME encoded format.
'*
'* @return stream
'******************************************************************************
Function GetStream() As ADODBExStream
    Dim wrap As ADODBExStream: Set wrap = New ADODBExStream
    Set wrap.Origin = mOrigin.GetStream()
    Set GetStream = wrap
End Function

'******************************************************************************
'* [概  要] SaveToFile メソッド
'* [詳  細] SaveToFile のラッパーメソッド。
'*          Saves the body part content to the specified file.
'*
'* @param FileName xx
'******************************************************************************
Sub SaveToFile(FileName As String)
    Call mOrigin.SaveToFile(FileName)
End Sub

'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------

