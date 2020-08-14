VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADODBExField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] ADODB.Fieldクラスのラップ・拡張クラス
'* [詳  細] ADODB.FieldクラスのWrapperとしての機能を提供する他、
'*          ADODB.Fieldクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' ADODB.Fieldクラスについて。
' <https://docs.microsoft.com/en-us/previous-versions/ms805115%28v%3dmsdn.10%29>
' <https://docs.microsoft.com/ja-jp/sql/ado/reference/ado-api/connection-object-ado?view=sql-server-2017>
' <https://docs.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/ado-api-reference>
'
' Represents a column of data with a common data type. Each Field object
' corresponds to a column in the Recordset.
'
' 共通のデータ型のデータの列を表す。
' 各Fieldオブジェクトは、Recordsetの列に対応している。
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As ADODB.Field
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
'* ActualSize プロパティ（読み取り専用）
'*
'* [補  足]
'* ・フィールドの値の実際の長さを返却する。
'* ・Fieldオブジェクトの値の長さを判断できない場合は、adUnknown が返却される。
'*-----------------------------------------------------------------------------
Property Get ActualSize() As Long
    ActualSize = mOrigin.ActualSize
End Property

'*-----------------------------------------------------------------------------
'* Attributes プロパティ
'*
'* [補  足]
'* ・1つまたは複数のFieldAttributeEnum値の合計を返却する。
'* ・通常は値の取得のみ可能。
'* ・RecordのFieldsコレクションに追加された新しいFieldオブジェクトの場合は、
'*   FieldのValueプロパティを指定し、FieldsコレクションのUpdateメソッドを呼び
'*   出すことによって新しいFieldがデータプロバイダーによって正常に追加された
'*   直後にのみ、Attributesの値の取得と設定の両方が可能。
'* ・複数の属性を設定する場合は、該当する定数の合計を使用できる。
'*   プロパティの値を､互換性のない定数を含む合計に設定すると､エラーが発生する｡
'* ・このプロパティは､クライアント側のConnectionオブジェクトでは使用できない｡
'*-----------------------------------------------------------------------------
Property Get Attributes() As Long
    Attributes = mOrigin.Attributes
End Property
Property Let Attributes(arg As Long)
    mOrigin.Attributes = arg
End Property

'*-----------------------------------------------------------------------------
'* DataFormat プロパティ
'*
'* [補  足]
'* ・MSDN等に掲載されていないプロパティ。
'* ・実態はStdDataFormatオブジェクト。StdDataFormatオブジェクトはMicrosoft Data
'*   Formatting Object Libraryを参照設定に追加すると確認できるとのこと。
'* ・以下のように使用する。
'*   Set dft = New StdFormat.StdDataFormat
'*   dft.Type = fmtCustom
'*   dft.Format = "Currency"
'*   Set rs.Fields("Price").DataFormat = dft
'*   Debug.Print rs!Price '通貨型のフォーマットで表示
'*-----------------------------------------------------------------------------
Property Get DataFormat() As Variant
    Set DataFormat = mOrigin.DataFormat
End Property
Property Set DataFormat(arg As Variant)
    Set mOrigin.DataFormat = arg
End Property

'*-----------------------------------------------------------------------------
'* DefinedSize プロパティ
'*
'* [補  足]
'* ・Fieldオブジェクトのデータ容量（フィールドの定義サイズ（バイト））を返却する。
'*-----------------------------------------------------------------------------
Property Get DefinedSize() As Long
    DefinedSize = mOrigin.DefinedSize
End Property
Property Let DefinedSize(arg As Long)
    mOrigin.DefinedSize = arg
End Property

    
'*-----------------------------------------------------------------------------
'* Name プロパティ
'*
'* [補  足]
'* ・オブジェクトの名前を返却する。
'* ・通常は値の取得のみ可能。
'* ・RecordのFieldsコレクションに追加された新しいFieldオブジェクトの場合は、
'*   FieldのValueプロパティを指定し、FieldsコレクションのUpdateメソッドを呼び
'*   出すことによって新しいFieldがデータプロバイダーによって正常に追加された
'*   直後にのみ、Nameの値の取得と設定の両方が可能。
'*-----------------------------------------------------------------------------
Property Get Name() As String
    Name = mOrigin.Name
End Property
Property Let Name(arg As String)
    mOrigin.Name = arg
End Property

'*-----------------------------------------------------------------------------
'* NumericScale プロパティ
'*
'* [補  足]
'* ・Field オブジェクトの数値の小数点以下の桁数を返却する。
'* ・通常は値の取得のみ可能。
'* ・RecordのFieldsコレクションに追加された新しいFieldオブジェクトの場合は、
'*   FieldのValueプロパティを指定し、FieldsコレクションのUpdateメソッドを呼び
'*   出すことによって新しいFieldがデータプロバイダーによって正常に追加された
'*   直後にのみ、NumericScaleの値の取得と設定の両方が可能。
'*-----------------------------------------------------------------------------
Property Get NumericScale() As Byte
    NumericScale = mOrigin.NumericScale
End Property
Property Let NumericScale(arg As Byte)
    mOrigin.NumericScale = arg
End Property

'*-----------------------------------------------------------------------------
'* OriginalValue プロパティ（読み取り専用）
'*
'* [補  足]
'* ・変更が行われる前のレコードに存在していたFieldの値を返却する。
'*-----------------------------------------------------------------------------
Property Get OriginalValue() As Variant
    OriginalValue = mOrigin.OriginalValue
End Property

'*-----------------------------------------------------------------------------
'* Precision プロパティ（読み取り専用）
'*
'* [補  足]
'* ・数値型の Field オブジェクトの精度（最大桁数を示すバイト型の値）を返却する。
'* ・通常は値の取得のみ可能。
'* ・RecordのFieldsコレクションに追加された新しいFieldオブジェクトの場合は、
'*   FieldのValueプロパティを指定し、FieldsコレクションのUpdateメソッドを呼び
'*   出すことによって新しいFieldがデータプロバイダーによって正常に追加された
'*   直後にのみ、Precisionの値の取得と設定の両方が可能。
'*-----------------------------------------------------------------------------
Property Get Precision() As Byte
    Precision = mOrigin.Precision
End Property
Property Let Precision(arg As Byte)
    mOrigin.Precision = arg
End Property

'*-----------------------------------------------------------------------------
'* Properties プロパティ（読み取り専用）
'*
'* [補  足]
'* ・PropertiesオブジェクトをラップしたADODBExPropertiesオブジェクトを返却する。
'* ・元はRecordsetオブジェクトに設定されているPropertiesオブジェクトを返却。
'*-----------------------------------------------------------------------------
Property Get Properties() As ADODBExProperties
    Dim wrap As ADODBExProperties: Set wrap = New ADODBExProperties
    Set wrap.Origin = mOrigin.Properties
    Set Properties = wrap
End Property

'*-----------------------------------------------------------------------------
'* Status プロパティ（読み取り専用）
'*
'* [補  足]
'* ・Fieldオブジェクトのステータスの値を返却する。
'* ・既定値はadFieldOK。
'* ・更新が失敗した場合､エラーが返され､Statusプロパティは操作とエラーステータス
'*   コードの組み合わせ値を示す｡ 各フィールドのStatusプロパティを使用して､フィー
'*   ルドが追加､変更､または削除されなかった理由を判断可能｡
'* ・フィールドの追加、変更、または削除に関する問題は、このプロパティを通じて
'*   報告される。たとえば、ユーザーがフィールドを削除すると、Fieldsコレクション
'*   で削除のマークが付けられる。 ユーザーが権限を持たないフィールドを削除しよう
'*   としたためにUpdate呼び出しがエラーを返した場合、このフィールドのステータス
'*   はadFieldPermissionDeniedまたはadFieldPendingDeleteになる。
'*-----------------------------------------------------------------------------
Property Get Status() As FieldStatusEnum
    Status = mOrigin.Status
End Property

'*-----------------------------------------------------------------------------
'* Type_ プロパティ（読み取り専用）（Typeが予約語のためType_に変更）
'*
'* [補  足]
'* ・オブジェクトのデータ型を返却する。
'* ・Typeが予約語で使用できないため、Type_に変更。
'* ・通常は値の取得のみ可能。
'* ・RecordのFieldsコレクションに追加された新しいFieldオブジェクトの場合は、
'*   FieldのValueプロパティを指定し、FieldsコレクションのUpdateメソッドを呼び
'*   出すことによって新しいFieldがデータプロバイダーによって正常に追加された
'*   直後にのみ、Typeの値の取得と設定の両方が可能。
'*-----------------------------------------------------------------------------
Property Get Type_() As DataTypeEnum
    Type_ = mOrigin.Type
End Property
Property Let Type_(arg As DataTypeEnum)
    mOrigin.Type_ = arg
End Property

'*-----------------------------------------------------------------------------
'* UnderlyingValue プロパティ（読み取り専用）
'*
'* [補  足]
'* ・データベース内のFieldオブジェクトの現在の値を返却する。
'*-----------------------------------------------------------------------------
Property Get UnderlyingValue() As Variant
    UnderlyingValue = mOrigin.UnderlyingValue
End Property

'*-----------------------------------------------------------------------------
'* Value プロパティ（読み取り専用:既定メンバー）
'*
'* [補  足]
'* ・Fieldオブジェクトの現在の値を返却する。
'* ・レコードのFieldsコレクションに追加される新しいFieldオブジェクトの場合､
'*   他のプロパティを設定する前にまずValueプロパティを設定し、更新を実行する
'*   必要がある｡
'*-----------------------------------------------------------------------------
Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Value = mOrigin.Value
End Property
Property Let Value(arg As Variant)
    mOrigin.Value = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("ADODB.Field")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] AppendChunkメソッド
'* [詳  細] AppendChunkのラッパーメソッド。
'*          大きなサイズの文字列またはバイナリデータを格納する際に使用。
'*          ParameterオブジェクトのAttributesプロパティのadParamLongビットが
'*          True に設定されていると、そのパラメーターに対してAppendChunkメソ
'*          ッドを使用可能。
'*
'* @param Data オブジェクトに追加するデータ
'******************************************************************************
Sub AppendChunk(data)
    Call mOrigin.AppendChunk(data)
End Sub

'******************************************************************************
'* [概  要] GetChunkメソッド
'* [詳  細] GetChunkのラッパーメソッド。
'*          大きなサイズの文字列またはバイナリデータを取得する際に使用。
'*
'* @Length データサイズを指定
'******************************************************************************
Sub GetChunk(Length As Long)
    Call mOrigin.GetChunk(Length)
End Sub


