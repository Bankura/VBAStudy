VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADODBExParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] ADODB.Parameterクラスのラップ・拡張クラス
'* [詳  細] ADODB.ParameterクラスのWrapperとしての機能を提供する他、
'*          ADODB.Parameterクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' ADODB.Parameterクラスについて。
' <https://docs.microsoft.com/en-us/previous-versions/ms805115%28v%3dmsdn.10%29>
' <https://docs.microsoft.com/ja-jp/sql/ado/reference/ado-api/connection-object-ado?view=sql-server-2017>
' <https://docs.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/ado-api-reference>
'
' Represents a parameter or argument associated with a Command object, based on
' a parameterized query or stored procedure.
'
' パラメーター化されたクエリまたはストアドプロシージャに基づいて、Commandオブジ
' ェクトに関連付けられたパラメーターまたは引数を表す。
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As ADODB.Parameter
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
'* Attributes プロパティ
'*
'* [補  足]
'* ・オブジェクトの 1 つまたは複数の属性を示す。
'* ・Attributes プロパティは値の取得および設定が可能で、その値は 1 つまたは複数
'*   の ParameterAttributesEnum 値の合計になる。既定値は adParamSigned 。
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
'* Direction プロパティ
'*
'* [補  足]
'* ・Parameter が、入力パラメーター、出力パラメーター、または入出力両方のパラ
'*   メーターを表しているか､あるいは､ストアド プロシージャからの戻り値であるか
'*   を示す。
'* ・Direction プロパティは、プロシージャとのパラメーターのやり取りの方法を指
'*   定するために使う。
'*   Direction プロパティは読み取り/書き込み可能になっている。
'*   これにより、パラメーター情報を取得するために ADO がそれ以上プロバイダー
'*   を呼び出さないようにする場合に、この情報を設定したり、この情報を返さない
'*   プロバイダーを操作したりできる。
'* ・プロバイダーの中には、ストアド プロシージャのパラメーターの入出力の方向
'*   を確認できないものがある。
'*   その場合は、クエリを実行する前に Direction プロパティを設定する必要が
'*   ある。
'*-----------------------------------------------------------------------------
Property Get Direction() As ParameterDirectionEnum
    Direction = mOrigin.Direction
End Property
Property Let Direction(arg As ParameterDirectionEnum)
    mOrigin.Direction = arg
End Property

'*-----------------------------------------------------------------------------
'* Name プロパティ
'*
'* [補  足]
'* ・オブジェクトの名前を示す値を設定または取得する。
'* ・Parameters コレクションに追加されていない Parameter オブジェクトでは、
'*   Name プロパティは読み取り/書き込み可能。
'*   追加された Parameter オブジェクトとその他のオブジェクトでは、 Name プロパ
'*   ティは読み取り専用。名前はコレクション内で一意でなくてもかまわない。
'* ・オブジェクトの Name は、序数参照で取得でき、その後は、その名前で直接オブ
'*   ジェクトを参照できる。たとえば、rstmain プロパティ (20) を使用する。
'*   Name は更新可能になり、後でこのプロパティを参照すると、更新の更新が可能
'*   になり、このプロパティを rstmain. プロパティ ("更新可能性") と呼ぶことが
'*   できる。
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
'* ・数値の桁数を示す。数値の小数点以下の桁数を示すバイト型 (Byte) の値を設定
'*   または取得する。
'*-----------------------------------------------------------------------------
Property Get NumericScale() As Byte
    NumericScale = mOrigin.NumericScale
End Property
Property Let NumericScale(arg As Byte)
    mOrigin.NumericScale = arg
End Property

'*-----------------------------------------------------------------------------
'* Precision プロパティ
'*
'* [補  足]
'* ・数値の精度を示す。値を表すために使用する最大桁数を示すバイト型 (Byte) の
'*   値を設定または取得する。
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
'* ・Commandに設定されているPropertiesオブジェクトを取得する。
'*-----------------------------------------------------------------------------
Property Get Properties() As ADODBExProperties
    Dim wrap As ADODBExProperties: Set wrap = New ADODBExProperties
    Set wrap.Origin = mOrigin.Properties
    Set Properties = wrap
End Property

'*-----------------------------------------------------------------------------
'* Size プロパティ
'*
'* [補  足]
'* ・Parameter オブジェクトの最大サイズをバイト数または文字数で示す。
'* ・Parameter オブジェクトの Value プロパティで設定または取得できる値の最大サ
'*   イズを調べるには、Size プロパティを使用する。
'* ・Parameter オブジェクトとして可変長データ型 (たとえば、adVarChar などのすべ
'*   ての文字列型 (String)) を指定した場合、Parameters コレクションにそのオブジ
'*   ェクトを追加する前に、オブジェクトの Size プロパティを設定する必要があり、
'*   この設定を行わないとエラーが発生する。
'* ・Parameter オブジェクトが Command オブジェクトの Parameters コレクションに
'*   既に追加されている場合に、そのデータ型を可変長データ型に変更した場合は、
'*   Command オブジェクトを実行する前に Parameter オブジェクトの Size プロパテ
'*   ィを設定する必要があり、この設定を行わないとエラーが発生する。
'* ・Refresh メソッドを使用してプロバイダーからパラメーター情報を取得したとき
'*   に、可変長データ型の Parameter オブジェクトが返された場合、可能な最大サイ
'*   ズに基づいてパラメーターにメモリが割り当てられるが、これが原因で実行時に
'*   エラーが発生することがある。
'*   エラーを回避するには、コマンドを実行する前に、明示的にこれらのパラメーター
'*   の Size プロパティを設定すること。
'* ・Size プロパティは､値の取得および設定が可能｡
'*-----------------------------------------------------------------------------
Property Get Size() As Long
    Size = mOrigin.Size
End Property
Property Let Size(arg As Long)
    mOrigin.Size = arg
End Property

'*-----------------------------------------------------------------------------
'* Type プロパティ（Typeが予約語のためType_に変更）
'*
'* [補  足]
'* ・操作の種類またはデータ型を示す。
'*-----------------------------------------------------------------------------
Property Get Type_() As DataTypeEnum
    Type_ = mOrigin.Type
End Property
Property Let Type_(arg As DataTypeEnum)
    mOrigin.Type = arg
End Property

'*-----------------------------------------------------------------------------
'* Value プロパティ（既定メンバー）
'*
'* [補  足]
'* ・オブジェクトに割り当てられた値を示す。
'* ・Value プロパティはさまざまな要因によって値の取得と設定が可能な場合と、
'*   値の取得のみが可能な場合がある。
'* ・ADO では、 Value プロパティを使用してロングバイナリデータを設定および
'*   取得可能。
'* ・[!メモ] Parameter オブジェクトの場合、ADO はプロバイダーから一度だけ
'*   Value プロパティを取得する。
'*   コマンドに含まれる Parameter の Value プロパティが空で、このコマンドから
'*   Recordset を作成する場合は、 Value プロパティを取得する前に、 Recordset
'*   を閉じる必要がある。
'*   このようにしないと、プロバイダーによっては、 Value プロパティが空になり、
'*   正しい値が格納されない場合がある。
'*-----------------------------------------------------------------------------
Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    If IsObject(mOrigin.Value) Then
        Set Value = mOrigin.Value
    Else
        Value = mOrigin.Value
    End If
End Property
Property Let Value(arg As Variant)
    If IsObject(arg) Then
        Set mOrigin.Value = arg
    Else
        mOrigin.Value = arg
    End If
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("ADODB.Parameter")
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
'* @param Val オブジェクトに追加するデータ
'******************************************************************************
Sub AppendChunk(Val)
    Call mOrigin.AppendChunk(Val)
End Sub
