VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADODBExFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] ADODB.Fieldsクラスのラップ・拡張クラス
'* [詳  細] ADODB.FieldsクラスのWrapperとしての機能を提供する他、
'*          ADODB.Fieldsクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' ADODB.Parametersクラスについて。
' <https://docs.microsoft.com/en-us/previous-versions/ms805115%28v%3dmsdn.10%29>
' <https://docs.microsoft.com/ja-jp/sql/ado/reference/ado-api/connection-object-ado?view=sql-server-2017>
' <https://docs.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/ado-api-reference>
'
' Contains all the Field objects of a Recordset object.
'
' RecordsetオブジェクトのすべてのFieldオブジェクトが含まれる。
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As ADODB.Fields
Private mOrigin As Object
Private mKeys As Collection

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
    Set mKeys = New Collection
    Dim myField As Object
    For Each myField In mOrigin
        Dim wrap As ADODBExField: Set wrap = New ADODBExField
        Set wrap.Origin = myField
        mKeys.Add wrap, myField.Name
    Next myField
End Property

'*-----------------------------------------------------------------------------
'* Count プロパティ（読み取り専用）
'*
'* [補  足]
'* ・コレクション内のオブジェクト数を示す。
'* ・Count プロパティは､特定のコレクション内のオブジェクトの数を調べるために
'*   使う｡
'* ・コレクションのメンバーは 0 から順に番号が割り当てられるため、ループを使う
'*   場合は常に 0 から始めて、 Count プロパティより 1 小さい値で終わらせる必要
'*   がある。
'*   Microsoft Visual Basic で Count プロパティをチェックせずにコレクションの
'*   メンバーをループ処理するには、 For Each...Next コマンドを使う。
'* ・Count が 0 の場合、コレクションにはオブジェクトが含まれていないことを意味
'*   する。
'*-----------------------------------------------------------------------------
Property Get Count() As Long
    Count = mOrigin.Count
End Property

'*-----------------------------------------------------------------------------
'* Item プロパティ（読み取り専用:既定メンバー）
'*
'* [補  足]
'* ・コレクションの特定のメンバーをその名前または序数で示す。
'* ・Indexには、コレクション内のオブジェクトの名前または序数に評価される
'*   バリアント型 ( Variant ) の式を指定。
'* ・Item プロパティは、コレクション内の特定のオブジェクトを返すために使う。
'*   コレクション内で Item が Index 引数に対応するオブジェクトを見つけられない
'*   場合は、エラーが発生する。
'*   また、コレクションの中には名前付きオブジェクトをサポートしていないものも
'*   ある。このようなコレクションでは、序数参照を使う必要がある。
'* ・Item プロパティはすべてのコレクションの既定プロパティなので､次のいずれの
'*   構文形式でも同じ結果が得られる｡
'*     collection.Item (Index)
'*     collection (Index)
'*-----------------------------------------------------------------------------
Property Get Item(Index) As ADODBExField
Attribute Item.VB_UserMemId = 0
    Dim wrap As ADODBExField: Set wrap = New ADODBExField
    Set wrap.Origin = mOrigin.Item(Index)
    Set Item = wrap
End Property


'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("ADODB.Fields")
    'Set mKeys = New Collection
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
    Set mKeys = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Appendメソッド
'* [詳  細] Appendのラッパーメソッド。
'*          コレクションにオブジェクトを追加する。
'*          コレクションに追加する前に、新しい Field オブジェクトを作成可能。
'*
'* @param Name        新しいFieldオブジェクトの名前。Fieldsに含まれる
'*                    他のオブジェクトとは異なる名前にする必要がある。
'* @param Type_       新しいフィールドのデータ型。既定値は adEmpty 。
'*                    adIDispatch 、 adIUnknown 、 adVariant の各データ型は
'*                    ADO ではサポートされていないので、 Recordset に新しい
'*                    フィールドを追加するときに、これらのデータ型を使用する
'*                    ことはできない。
'* @param DefinedSize 省略可能。新しいフィールドの定義されたサイズ。
'*                    このパラメーターの既定値は、Type によって決まる。
'*                    DefinedSize が 255 バイトより大きいフィールドは、
'*                    可変長列として扱われる (既定のDefinedSizeは指定されない)。
'* @param Attrib      省略可能。新しいフィールドの属性を指定。
'*                    既定値は adFldDefault 。値を指定しないと、Type に基づく
'*                    属性が設定される｡
'* @param FieldValue  省略可能。新しいフィールドの値。
'*                    値を指定しないと、フィールドは Null値で追加される。
'*
'******************************************************************************
Sub Append(Name As String, _
           Type_ As DataTypeEnum, _
           Optional DefinedSize As Long, _
           Optional Attrib As FieldAttributeEnum = adFldUnspecified, _
           Optional FieldValue)

    If IsMissing(FieldValue) Then
        Call mOrigin.Append(Name, Type_, DefinedSize, Attrib)
    Else
        Call mOrigin.Append(Name, Type_, DefinedSize, Attrib, FieldValue)
    End If
    
    mKeys.Add mOrigin.Item(Name), Name
    
End Sub

'******************************************************************************
'* [概  要] CancelUpdateメソッド
'* [詳  細] CancelUpdateのラッパーメソッド。
'*          Update メソッドを呼び出す前に行った、Recordset オブジェクトの
'*          カレント行や新規行に対する変更、または Record オブジェクトの
'*          Fields コレクションに対する変更を、すべてキャンセルする。
'*
'******************************************************************************
Sub CancelUpdate()
    Call mOrigin.CancelUpdate
End Sub

'******************************************************************************
'* [概  要] Deleteメソッド
'* [詳  細] Deleteのラッパーメソッド。
'*          Fields コレクションからオブジェクトを削除する。
'*
'* @param Index 削除するFieldオブジェクトを指定する値。
'*              このパラメーターには、 Field オブジェクトの名前または Field
'*              オブジェクト自体のインデックスを使用できる。
'*
'******************************************************************************
Sub Delete(Index)
    Dim keyName: keyName = mOrigin.Item(Index).Name
    Call mOrigin.Delete(Index)

    Dim i As Long
    For i = 1 To mKeys.Count
        If mKeys(i) = keyName Then
            mKeys.Remove i
            Exit Sub
        End If
    Next
End Sub

'******************************************************************************
'* [概  要] Refreshメソッド
'* [詳  細] Refreshのラッパーメソッド。
'*          コレクションのオブジェクトを更新する。
'*          Fields コレクションに対して Refresh メソッドを使用しても、
'*          目に見える効果はない。基になっているデータベース構造から
'*          変更を取得するには、Requery メソッドを使用するか、または Recordset
'*          オブジェクトがブックマークをサポートしない場合は MoveFirst メソッド
'*          を使用する必要がある。
'*
'******************************************************************************
Sub Refresh()
    Call mOrigin.Refresh
End Sub

'******************************************************************************
'* [概  要] Resyncメソッド
'* [詳  細] Resyncのラッパーメソッド。
'*          現在の Recordset オブジェクト、または Record オブジェクトの Fields
'*          コレクションのデータを､基になるデータベースのデータで更新する。
'*          Resync メソッドは、Record オブジェクトの Fields コレクションの値を、
'*          基になるデータソースと再同期させる場合に使用する。
'*          Count プロパティは、このメソッドによる影響を受けない。
'*          ResyncValues を adResyncAllValues (既定値) に設定すると、コレクショ
'*          ンに含まれる Field オブジェクトのプロパティ UnderlyingValue、Value、
'*          および OriginalValue が同期化さる。
'*          ResyncValues を adResyncUnderlyingValues に設定すると、
'*          UnderlyingValue プロパティだけが同期化される。
'*          呼び出し時の各 Field オブジェクトの Status プロパティの値も、
'*          Resync の動作に影響を与える。
'*          Status の値が adFieldPendingUnknown または adFieldPendingInsert
'*          である Field オブジェクトに対しては、 Resync は何も行わない。
'*          Status の値が adFieldPendingChange または adFieldPendingDelete
'*          である場合は、 Resync はデータソースにまだ存在しているフィールドの
'*          データ値を同期化する。
'*          Resyncが呼び出されたときにエラーが発生しない限り、ResyncはFieldオブ
'*          ジェクトのStatus値を変更しない。
'*          たとえば、フィールドが存在しなくなった場合、プロバイダーは
'*          adFieldDoesNotExistなどのFieldオブジェクトの適切なStatus値を返す。
'*          返されるStatus値は、Statusプロパティの値内で論理的に組み合わせる
'*          ことができる。
'*
'* @param ResyncValues 省略可能。基になる値を上書きするかどうかを指定。
'*
'******************************************************************************
Sub Resync(Optional ResyncValues As ResyncEnum = adResyncAllValues)
    Call mOrigin.Resync(ResyncValues)
End Sub

'******************************************************************************
'* [概  要] Updateメソッド
'* [詳  細] Updateのラッパーメソッド。
'*          Record オブジェクトの Fields コレクションに加えた変更を保存する。
'*
'******************************************************************************
Sub Update()
    Call mOrigin.Update
End Sub

'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------

'******************************************************************************
'* [概  要] NewEnum メソッド
'* [詳  細] For Each 構文でイテレーターを返すためのメソッド。
'*
'* @return イテレーター
'******************************************************************************
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mKeys.[_NewEnum]
End Function
