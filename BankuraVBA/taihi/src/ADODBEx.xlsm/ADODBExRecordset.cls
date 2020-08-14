VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADODBExRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] ADODB.Recordsetクラスのラップ・拡張クラス
'* [詳  細] ADODB.RecordsetクラスのWrapperとしての機能を提供する他、
'*          ADODB.Recordsetクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' ADODB.Recordsetクラスについて。
' <https://docs.microsoft.com/en-us/previous-versions/ms805115%28v%3dmsdn.10%29>
' <https://docs.microsoft.com/ja-jp/sql/ado/reference/ado-api/connection-object-ado?view=sql-server-2017>
' <https://docs.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/ado-api-reference>
'
' Represents the entire set of records from a base table or the results of an
' executed command.
' All Recordset objects consist of records (rows) and fields (columns).
'
' ベーステーブルのレコードセット全体、または実行されたコマンドの結果を表す。
' すべてのRecordsetオブジェクトは、レコード（行）とフィールド（列）で構成され
' る。
'
' <イベント情報>
' Event WillChangeField(cFields As Long, Fields, adStatus As EventStatusEnum, pRecordset As Recordset)
' Event WillChangeRecord(adReason As EventReasonEnum, cRecords As Long, adStatus As EventStatusEnum, pRecordset As Recordset)
' Event WillChangeRecordset(adReason As EventReasonEnum, adStatus As EventStatusEnum, pRecordset As Recordset)
' Event WillMove(adReason As EventReasonEnum, adStatus As EventStatusEnum, pRecordset As Recordset)
' Event EndOfRecordset(fMoreData As Boolean, adStatus As EventStatusEnum, pRecordset As Recordset)
' Event FetchComplete(pError As Error, adStatus As EventStatusEnum, pRecordset As Recordset)
' Event FetchProgress(Progress As Long, MaxProgress As Long, adStatus As EventStatusEnum, pRecordset As Recordset)
' Event FieldChangeComplete(cFields As Long, Fields, pError As Error, adStatus As EventStatusEnum, pRecordset As Recordset)
' Event MoveComplete(adReason As EventReasonEnum, pError As Error, adStatus As EventStatusEnum, pRecordset As Recordset)
' Event RecordChangeComplete(adReason As EventReasonEnum, cRecords As Long, pError As Error, adStatus As EventStatusEnum, pRecordset As Recordset)
' Event RecordsetChangeComplete(adReason As EventReasonEnum, pError As Error, adStatus As EventStatusEnum, pRecordset As Recordset)
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As ADODB.Recordset
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
    If Not mOrigin Is Nothing Then
        If mOrigin.State <> adStateClosed Then
            Call mOrigin.Close
        End If
    End If
    Set mOrigin = arg
End Property

'*-----------------------------------------------------------------------------
'* AbsolutePage プロパティ
'*
'* [補  足]
'* ・カレントレコードのページを設定。
'*-----------------------------------------------------------------------------
Property Get AbsolutePage() As PositionEnum
    AbsolutePage = mOrigin.AbsolutePage
End Property
Property Let AbsolutePage(arg As PositionEnum)
    mOrigin.AbsolutePage = arg
End Property

'*-----------------------------------------------------------------------------
'* AbsolutePosition プロパティ
'*
'* [補  足]
'* ・カレントレコード位置を設定。
'*-----------------------------------------------------------------------------
Property Get AbsolutePosition() As PositionEnum
    AbsolutePosition = mOrigin.AbsolutePosition
End Property
Property Let AbsolutePosition(arg As PositionEnum)
    mOrigin.AbsolutePosition = arg
End Property

'*-----------------------------------------------------------------------------
'* ActiveCommand プロパティ（読み取り専用）
'*
'* [補  足]
'* ・CommandオブジェクトをラップしたADODBExCommandオブジェクトを返却する。
'* ・元はRecordsetオブジェクトを作成したCommandオブジェクトを返却。
'*-----------------------------------------------------------------------------
Property Get ActiveCommand() As Object
    Set ActiveCommand = mOrigin.ActiveCommand
    If TypeName(mOrigin.ActiveCommand) = "Command" Then
        Dim wrap As ADODBExCommand: Set wrap = New ADODBExCommand
        Set wrap.Origin = mOrigin.ActiveCommand
        Set ActiveCommand = wrap
    Else
        Set ActiveCommand = mOrigin.ActiveCommand
    End If
End Property

'*-----------------------------------------------------------------------------
'* ActiveConnection プロパティ
'*
'*   ※オブジェクトブラウザでCommandのActiveConnection プロパティを確認すると
'*     型はConnectionとなっているが、接続情報文字列（String）も受け取ることが
'*     可能のため、Variant型としています。
'*
'* [補  足]
'* ・ADODBExConnection/Connectionオブジェクト、または接続情報文字列（String）
'*   を設定する。
'* ・Nothingを設定すると、Commandオブジェクトに現在設定されているConnection
'*   が解放される。
'* ・接続が閉じている場合には接続の定義が格納された文字列型 ( String ) の値
'*   を、接続が開いている場合には現在の Connection オブジェクトが格納された
'*   バリアント型 ( Variant ) の値を設定または取得する。
'*   既定は、Null オブジェクト参照。
'*   ConnectionString プロパティの説明を参照すること。
'* ・開いている Recordset オブジェクト、および Source プロパティが有効な
'*   Command オブジェクトに設定されている Recordset オブジェクトについては
'*   ActiveConnection プロパティは値の取得のみ可能。
'*   それ以外の場合は、値の取得および設定が可能。
'* ・このプロパティは、有効な Connection オブジェクトまたは有効な接続文字列
'*   に設定できる。
'*   この場合、プロバイダーが、この定義を使用して新しい Connection オブジェクト
'*   を作成し、接続を開く。
'*   さらに、プログラマが Connection オブジェクトにアクセスして拡張エラー情報
'*   を取得したり、その他のコマンドを実行したりできるように、プロバイダーが
'*   このプロパティを新しい Connection オブジェクトに設定する場合もある。
'* ・Open メソッドの ActiveConnection 引数を使用して Recordset オブジェクトを
'*   開いた場合、ActiveConnection プロパティは引数の値を継承する。
'* ・Recordset オブジェクトの Source プロパティを有効な Command オブジェクト
'*   変数に設定すると、Recordset の ActiveConnection プロパティは、Command
'*   オブジェクトの ActiveConnection プロパティの値を継承する。
'* ・リモートデータサービスの使用法: クライアント側の Recordset オブジェクト
'*   で使用されている場合、このプロパティは、接続文字列または (Microsoft visual
'*   basic または visual basic, Scripting Edition では) Nothingにのみ設定
'*   できる。
'*-----------------------------------------------------------------------------
Property Get ActiveConnection() As Variant
    If IsObject(mOrigin.ActiveConnection) Then
        Dim wrap As ADODBExConnection: Set wrap = New ADODBExConnection
        Set wrap.Origin = mOrigin.ActiveConnection
        Set ActiveConnection = wrap
    Else
        ActiveConnection = mOrigin.ActiveConnection
    End If
End Property
Property Let ActiveConnection(arg As Variant)
    If IsObject(arg) Then
        If TypeName(arg) = "ADODBExConnection" Then
            Set mOrigin.ActiveConnection = arg.Origin
        Else
            Set mOrigin.ActiveConnection = arg
        End If
    Else
        If TypeName(arg) = "Nothing" Then
            Set mOrigin.ActiveConnection = Nothing
        Else
            mOrigin.ActiveConnection = arg
        End If
    End If
End Property

'*-----------------------------------------------------------------------------
'* BOF プロパティ（読み取り専用）
'*
'* [補  足]
'* ・カレントレコード位置が最初のレコードより前ならTrueを返却する。
'*-----------------------------------------------------------------------------
Property Get BOF() As Boolean
    BOF = mOrigin.BOF
End Property

'*-----------------------------------------------------------------------------
'* Bookmark プロパティ
'*
'* [補  足]
'* ・カレントレコードを一意に識別するブックマークを設定する。
'*-----------------------------------------------------------------------------
Property Get Bookmark() As Variant
    If IsObject(mOrigin.Bookmark) Then
        Set Bookmark = mOrigin.Bookmark
    Else
        Bookmark = mOrigin.Bookmark
    End If
End Property
Property Let Bookmark(arg As Variant)
    If IsObject(arg) Then
        Set mOrigin.Bookmark = arg
    Else
        mOrigin.Bookmark = arg
    End If
End Property

'*-----------------------------------------------------------------------------
'* CacheSize プロパティ
'*
'* [補  足]
'* ・キャッシュするレコード数を設定する。
'*-----------------------------------------------------------------------------
Property Get CacheSize() As Long
    CacheSize = mOrigin.CacheSize
End Property
Property Let CacheSize(arg As Long)
    mOrigin.CacheSize = arg
End Property

'*-----------------------------------------------------------------------------
'* CursorLocation プロパティ
'*
'* [補  足]
'* ・カーソルサービスを設定する。
'*-----------------------------------------------------------------------------
Property Get CursorLocation() As CursorLocationEnum
    CursorLocation = mOrigin.CursorLocation
End Property
Property Let CursorLocation(arg As CursorLocationEnum)
    mOrigin.CursorLocation = arg
End Property

'*-----------------------------------------------------------------------------
'* CursorType プロパティ
'*
'* [補  足]
'* ・カーソルタイプを設定する。
'*-----------------------------------------------------------------------------
Property Get CursorType() As CursorTypeEnum
    CursorType = mOrigin.CursorType
End Property
Property Let CursorType(arg As CursorTypeEnum)
    mOrigin.CursorType = arg
End Property

'*-----------------------------------------------------------------------------
'* DataMember プロパティ
'*
'* [補  足]
'* ・DataSourceプロパティによって参照されるオブジェクトから取得する
'*   データメンバーの名前を示す。
'* ・DataMemberプロパティとDataSourceプロパティは組み合わせて使用する必要
'*   がある。
'*-----------------------------------------------------------------------------
Property Get DataMember() As String
    DataMember = mOrigin.DataMember
End Property
Property Let DataMember(arg As String)
    mOrigin.DataMember = arg
End Property

'*-----------------------------------------------------------------------------
'* DataSource プロパティ
'*
'* [補  足]
'* ・Recordsetオブジェクトとして表されるデータを持つオブジェクト。
'* ・DataMemberプロパティとDataSourceプロパティは組み合わせて使用する必要
'*   がある。
'*-----------------------------------------------------------------------------
Property Get DataSource() As Variant
    Set DataSource = mOrigin.DataSource
End Property
Property Set DataSource(arg As Variant)
    Set mOrigin.DataSource = arg
End Property

'*-----------------------------------------------------------------------------
'* EditMode プロパティ（読み取り専用）
'*
'* [補  足]
'* ・カレントコードの編集状況を返却する。
'*-----------------------------------------------------------------------------
Property Get EditMode() As EditModeEnum
    EditMode = mOrigin.EditMode
End Property

'*-----------------------------------------------------------------------------
'* EOF プロパティ（読み取り専用）
'*
'* [補  足]
'* ・カレントレコード位置が最後のレコードより後ならTrueを返却する。
'*-----------------------------------------------------------------------------
Property Get EOF() As Boolean
    EOF = mOrigin.EOF
End Property

'*-----------------------------------------------------------------------------
'* Fields プロパティ（読み取り専用:既定メンバー）
'*
'* [補  足]
'* ・１つ以上の列（フィールド）を表すオブジェクトを返却する。
'*-----------------------------------------------------------------------------
Property Get Fields() As ADODBExFields
Attribute Fields.VB_UserMemId = 0
    Dim wrap As ADODBExFields: Set wrap = New ADODBExFields
    Set wrap.Origin = mOrigin.Fields
    Set Fields = wrap
End Property

'*-----------------------------------------------------------------------------
'* Filter プロパティ
'*
'* [補  足]
'* ・データへのフィルタを設定する。
'*-----------------------------------------------------------------------------
Property Get Filter() As Variant
    If IsObject(mOrigin.Filter) Then
        Set Filter = mOrigin.Filter
    Else
        Filter = mOrigin.Filter
    End If
End Property
Property Let Filter(arg As Variant)
    If IsObject(arg) Then
        Set mOrigin.Filter = arg
    Else
        mOrigin.Filter = arg
    End If
End Property

'*-----------------------------------------------------------------------------
'* Index プロパティ
'*
'* [補  足]
'* ・インデックス名を設定する。
'*-----------------------------------------------------------------------------
Property Get Index() As String
    Index = mOrigin.Index
End Property
Property Let Index(arg As String)
    mOrigin.Index = arg
End Property

'*-----------------------------------------------------------------------------
'* LockType プロパティ
'*
'* [補  足]
'* ・レコードに対するロックの種類を設定する。
'*-----------------------------------------------------------------------------
Property Get LockType() As LockTypeEnum
    LockType = mOrigin.LockType
End Property
Property Let LockType(arg As LockTypeEnum)
    mOrigin.LockType = arg
End Property

'*-----------------------------------------------------------------------------
'* MarshalOptions プロパティ
'*
'* [補  足]
'* ・どのレコードがサーバーにマーシャリングされるかを示す。
'*-----------------------------------------------------------------------------
Property Get MarshalOptions() As MarshalOptionsEnum
    MarshalOptions = mOrigin.MarshalOptions
End Property
Property Let MarshalOptions(arg As MarshalOptionsEnum)
    mOrigin.MarshalOptions = arg
End Property

'*-----------------------------------------------------------------------------
'* MaxRecords プロパティ
'*
'* [補  足]
'* ・最大レコード数を設定する。
'*-----------------------------------------------------------------------------
Property Get MaxRecords() As Long
    MaxRecords = mOrigin.MaxRecords
End Property
Property Let MaxRecords(arg As Long)
    mOrigin.MaxRecords = arg
End Property

'*-----------------------------------------------------------------------------
'* PageCount プロパティ（読み取り専用）
'*
'* [補  足]
'* ・データのページ数を設定する。
'*-----------------------------------------------------------------------------
Property Get PageCount() As Long
    PageCount = mOrigin.PageCount
End Property

'*-----------------------------------------------------------------------------
'* PageSize プロパティ
'*
'* [補  足]
'* ・１ページあたりのレコード数を設定する。
'*-----------------------------------------------------------------------------
Property Get PageSize() As Long
    PageSize = mOrigin.PageSize
End Property
Property Let PageSize(arg As Long)
    mOrigin.PageSize = arg
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
'* RecordCount プロパティ（読み取り専用）
'*
'* [補  足]
'* ・オブジェクト内のレコード数を返却する。
'*-----------------------------------------------------------------------------
Property Get RecordCount() As Long
    RecordCount = mOrigin.RecordCount
End Property

'*-----------------------------------------------------------------------------
'* Sort プロパティ
'*
'* [補  足]
'* ・オブジェクト内のレコードのソートを設定する。
'*-----------------------------------------------------------------------------
Property Get Sort() As String
    Sort = mOrigin.Sort
End Property
Property Let Sort(arg As String)
    mOrigin.Sort = arg
End Property

'*-----------------------------------------------------------------------------
'* Source プロパティ
'*
'* [補  足]
'* ・データソースを設定する。
'*-----------------------------------------------------------------------------
Property Get Source() As Variant
    If IsObject(mOrigin.Source) Then
        Set Source = mOrigin.Source
    Else
        Source = mOrigin.Source
    End If
End Property
Property Let Source(arg As Variant)
    If IsObject(arg) Then
        Set mOrigin.Source = arg
    Else
        mOrigin.Source = arg
    End If
End Property

'*-----------------------------------------------------------------------------
'* State プロパティ（読み取り専用）
'*
'* [補  足]
'* ・オブジェクトの状態が開いているか、閉じているかを示す。
'* ・非同期メソッドを実行する対象になるすべてのオブジェクトについて､オブジェク
'*   トの状態が接続、実行、取得のいずれであるかを示す｡
'* ・ObjectStateEnum の値になる長整数型 ( Long ) の値を返す。
'*   既定値は adStateClosed 。
'* ・State プロパティを使用して､特定のオブジェクトの現在の状態をいつでも調べる
'*   ことができる｡
'* ・オブジェクトの State プロパティは、値の組み合わせになる場合がある。
'*   たとえば、ステートメントが実行中である場合、このプロパティの値は
'*   adStateOpen と adStateExecuting の組み合わせになる。
'* ・State プロパティは値の取得のみ可能｡
'*-----------------------------------------------------------------------------
Property Get State() As Long
    State = mOrigin.State
End Property

'*-----------------------------------------------------------------------------
'* Status プロパティ（読み取り専用）
'*
'* [補  足]
'* ・カレントレコードの処理状態を返却する。
'*-----------------------------------------------------------------------------
Property Get Status() As Long
    Status = mOrigin.Status
End Property

'*-----------------------------------------------------------------------------
'* StayInSync プロパティ
'*
'* [補  足]
'* ・親レコード位置に同期させるかを設定する。
'*-----------------------------------------------------------------------------
Property Get StayInSync() As Boolean
    StayInSync = mOrigin.StayInSync
End Property
Property Let StayInSync(arg As Boolean)
    mOrigin.StayInSync = arg
End Property

'*-----------------------------------------------------------------------------
'* Collect プロパティ
'*
'* [補  足]
'* ・隠し項目。
'*-----------------------------------------------------------------------------
Property Get Collect(Index)
    StayInSync = mOrigin.Collect(Index)
End Property
Property Let Collect(Index, arg)
    mOrigin.Collect(Index) = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("ADODB.Recordset")
End Sub
Private Sub Class_Terminate()
    If Not mOrigin Is Nothing Then
        If mOrigin.State <> adStateClosed Then
            Call mOrigin.Close
        End If
    End If
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] AddNewメソッド
'* [詳  細] AddNewのラッパーメソッド。
'*          レコードセットに新しいレコードを追加する。
'*          レコードセットを読み取り専用以外のロックタイプでOpenする必要がある。
'*          Field、Valueに値を設定し、UpdateメソッドやUpdateBatchメソッドを実行
'*          することでデータベースに書き込みが可能。
'*
'*          AddNew メソッドは、新しいレコードの作成と初期化に使用する。
'*          現在の Recordset オブジェクトにレコードを追加できるかどうかを確認
'*          するには、 Supports メソッドに adAddNew ( CursorOptionEnum 値) を
'*          指定して使用する。
'*          AddNew メソッドを呼び出した後は、新しいレコードがカレント レコード
'*          になり、 Update メソッドを呼び出した後もカレント レコードのままに
'*          なる。新しいレコードは Recordset の最後に追加されるので、Update
'*          に続いて MoveNext を呼び出すと、 Recordset の末尾を超えて移動し、
'*          EOF が True になる。
'*          Recordset オブジェクトがブックマークをサポートしていない場合は、
'*          別のレコードに移動すると、新しいレコードにアクセスできなくなること
'*          がある。
'*          カーソルの種類によっては、新規レコードにアクセスするために、 Requery
'*          メソッドを呼び出すことが必要になる場合がある。
'*          カレント レコードの編集中または新規レコードの追加中に AddNew メソ
'*          ッドを呼び出すと、ADO は Update メソッドを呼び出してすべての変更を
'*          保存した後、新しいレコードを作成する。
'*          AddNew メソッドの動作は、Recordset オブジェクトの更新モード、およ
'*          び Fieldlist 引数と Values 引数を渡すかどうかによって異なる。
'*          即時更新モード( updateメソッドを呼び出した後、プロバイダーが基にな
'*          るデータソースに変更を書き込む場合) では、引数を指定せずにAddNew
'*          メソッドを呼び出すと、 EditModeプロパティがadEditAddに設定される
'*           (editmodeenum値)。
'*          プロバイダーは、フィールド値の変更をローカルにキャッシュする。
'*          Updateメソッドを呼び出すと、新しいレコードがデータベースにポスト
'*          され、EditModeプロパティがadEditNone（EditModeEnum値）にリセット
'*          される。
'*          Fieldlist引数とValues引数を渡すと、ADO によって新しいレコードがす
'*          ぐにデータベースにポストされる(更新呼び出しは必要ない)。
'*          EditModeプロパティの値は変更されない (adEditNone)。
'*          バッチ更新モード(プロバイダーが複数の変更をキャッシュし、
'*          UpdateBatchメソッドを呼び出したときに、基になるデータソースにその
'*          変更を書き込む場合)、引数を指定せずにAddNewメソッドを呼び出すと、
'*          EditModeが設定される。
'*          プロパティをadEditAddする。プロバイダーは、フィールド値の変更
'*          をローカルにキャッシュする。
'*          Updateメソッドを呼び出すと、新しいレコードが現在のRecordsetに追加
'*          され、EditModeプロパティがadEditNoneにリセットされるが、
'*          UpdateBatchメソッドを呼び出すまで、プロバイダーは基になるデータベ
'*          ースに変更を送信しない。
'*          Fieldlist引数とValues引数を渡すと、ADO はキャッシュに格納されてい
'*          るプロバイダーに新しいレコードを送信する。
'*          新しいレコードを基になるデータベースにポストするには、 UpdateBatch
'*          メソッドを呼び出す必要がある。
'*
'* @param FieldList 省略可能。 新しいレコード内の単一フィールドの名前であるか、
'*                  複数フィールドの名前または順序の配列。
'* @param Values    省略可能。 新しいレコードのフィールドの値の1つまたは複数
'*                  の配列。 Fieldlistが配列の場合は、値も同じ数のメンバーの
'*                  配列である必要がある。それ以外の場合は、エラーが発生する。
'*                  また、一方の配列におけるフィールド名の順序と、もう一方の
'*                  配列におけるフィールド値の順序は、一致している必要がある。
'******************************************************************************
Sub AddNew(Optional FieldList, Optional Values)
    If IsMissing(FieldList) Then
        If IsMissing(Values) Then
            Call mOrigin.AddNew
        Else
            Call mOrigin.AddNew(, Values)
        End If
    ElseIf IsMissing(Values) Then
        Call mOrigin.AddNew(FieldList)
    Else
        Call mOrigin.AddNew(FieldList, Values)
    End If
End Sub

'******************************************************************************
'* [概  要] Cancelメソッド
'* [詳  細] Cancelのラッパーメソッド。
'*          保留中の非同期メソッド呼び出しの実行を取り消す。
'*          非同期メソッド呼び出し (つまり、adAsyncConnect、adAsyncExecute、
'*          または adAsyncFetch オプションを指定して呼び出されたメソッド) の
'*          実行を中止する。
'*          RecordSetの場合、Openが最後の非同期呼び出しが中止されるメソッド｡
'*
'******************************************************************************
Sub Cancel()
    Call mOrigin.Cancel
End Sub

'******************************************************************************
'* [概  要] CancelBatchメソッド
'* [詳  細] CancelBatchのラッパーメソッド。
'*          保留中のバッチ更新をキャンセルする｡
'*          CancelBatch メソッドは、バッチ更新モードの Recordset で保留中の更新
'*          をすべて取り消す場合に使用する。
'*          即時更新モードの Recordset では、 adAffectCurrent を指定しないで
'*          CancelBatch を呼び出すと、エラーが発生する。
'*          カレント レコードの編集中または新規レコードの追加中に CancelBatch
'*          メソッドを呼び出すと、最初に CancelUpdate メソッドが呼び出されて、
'*          キャッシュされているすべての変更が取り消される。
'*          その後、 Recordset で保留中のすべての変更が取り消される。
'*          CancelBatch を呼び出した後は、カレント レコードを特定できない可能性
'*          がある。
'*          特に、新規レコードの追加処理中の場合は、その可能性が高くなる。
'*          このため、 CancelBatch の呼び出し後は、カレント レコードの位置を
'*          Recordset 内の既知の位置に設定することをお勧めする。
'*          たとえば、 MoveFirst メソッドを呼び出す。
'*          基になるデータとの競合 (たとえば、他のユーザーによってレコードが
'*          削除されている場合) が原因で保留中の更新を取り消すこどができない
'*          場合、プロバイダーは Errors コレクションに警告を返すが、
'*          プログラムの実行は停止しない。
'*          要求したすべてのレコードで競合が発生した場合にのみ、実行時エラーが
'*          発生する。
'*          競合しているレコードを特定するには、Filter プロパティ
'*          (adFilterAffectedRecords) と Status プロパティを使用する。
'*
'* @param AffectRecords 省略可能。CancelBatch メソッドで処理するレコードの数を
'*                      示す AffectEnum 値。
'******************************************************************************
Sub CancelBatch(Optional AffectRecords As AffectEnum = adAffectAll)
    Call mOrigin.CancelBatch(AffectRecords)
End Sub


'******************************************************************************
'* [概  要] CancelUpdateメソッド
'* [詳  細] CancelUpdateのラッパーメソッド。
'*          Update メソッドを呼び出す前に行った、Recordset オブジェクトのカレン
'*          ト行や新規行に対する変更、または Record オブジェクトの Fields コレ
'*          クションに対する変更を、すべてキャンセルする。
'*          CancelUpdate メソッドは、カレント行に対する変更を取り消す場合、また
'*          は新しく追加した行を破棄する場合に使用する。
'*          Update メソッドを呼び出した後では、カレント行に対する変更または新規
'*          行を取り消すことはできない。
'*          ただし、変更が RollbackTrans メソッドでロールバックできるトランザク
'*          ションの一部である場合、またはバッチ更新の一部である場合を除く。
'*          バッチ更新の場合は、 CancelUpdate メソッドまたは CancelBatch メソッ
'*          ドで Update を取り消すことができる。
'*          新しい行を追加している場合は、 CancelUpdate メソッドを呼び出すと、
'*          AddNew を呼び出す前にカレントであった行がカレント行になる。
'*          編集モードでカレント レコードから移動する (たとえば、Move、
'*          NextRecordset、または Close を使用して) 場合は、 CancelUpdate を使
'*          用して、保留中のすべての変更を取り消すこどができる。
'*          データ ソースの更新が成功しなかった場合 (たとえば、参照整合性違反の
'*          ために削除の試みが失敗し、 Delete の呼び出しの後で Recordset が編集
'*          モードのままになった場合) には、この処理が必要になることがある。
'*          CancelUpdate メソッドは、 Field オブジェクトの保留中の挿入または削除
'*          をすべてキャンセルし、既存のフィールドの保留中の更新をキャンセルして
'*          元の値に戻す。
'*          Fields コレクションのすべてのフィールドの Status プロパティは、
'*          adFieldOK に設定される。
'*
'******************************************************************************
Sub CancelUpdate()
    Call mOrigin.CancelUpdate
End Sub

'******************************************************************************
'* [概  要] Cloneメソッド
'* [詳  細] Cloneのラッパーメソッド。
'*          既存の Recordset オブジェクトから Recordset オブジェクトの複製を
'*          作成する｡
'*          必要に応じて､複製を読み取り専用に指定できる｡
'*
'*          Clone メソッドは、Recordset オブジェクトの複製を複数作成する場合に
'*          使用する。
'*          特に、レコードセットの複数のカレント レコードを保持したい場合には、
'*          このメソッドを使用する。
'*          Clone メソッドを使用した方が、元の Recordset オブジェクトと同じ定義
'*          で新しいオブジェクトを作成して開くよりも効率的である。
'*          元の Recordset で Filter プロパティが設定されていても、複製には適用
'*          されない。
'*          結果をフィルターするには、新しい Recordset の Filter プロパティを設
'*          定する必要がある。
'*          既存の Filter 値をコピーする最も簡単な方法は、値を直接代入すること
'*          である｡
'*          1 つの Recordset オブジェクトに加えた変更は、カーソルの種類に関係
'*          なく、すべての複製で表示できる。
'*          ただし、複製元の Recordset で Requery を実行した後は、複製は複製元
'*          の Recordset とは同期しなくなる。
'*          複製元の Recordset を閉じても、その複製は開いたままである。
'*          また、1 つの複製を閉じても、複製元または他の複製は開いたままに
'*          なる。
'*          複製できるのは、ブックマークをサポートする Recordset オブジェクト
'*          だけである。ブックマークの値は交換可能。
'*          つまり、1 つの Recordset オブジェクトのブックマーク参照は、すべての
'*          複製の同じレコードを参照する。
'*          発生する Recordset イベントの中には、すべての Recordset 複製におい
'*          て発生するものがある。
'*          ただし、複製された Recordset によってカレント レコードが異なる可能性
'*          があるため、複製ではイベントが無効である場合もある。
'*          たとえば、あるフィールドの値を変更すると、変更された Recordset と
'*          すべての複製で、WillChangeField イベントが発生する。
'*          複製された (変更されていない) Recordset の WillChangeField イベント
'*          の Fields パラメーターは、その複製のカレント レコードのフィールドを
'*          参照するだけであり、変更が発生した元の Recordset のカレントレコード
'*          とは異なる可能性がある。
'*
'* @param LockType 省略可能。複製元 Recordset のロックの種類にするか、または読み
'*                 取り専用 Recordset にするかを指定する、 LockTypeEnum の値。
'*                 有効な値は、 adLockUnspecified または adLockReadOnly 。
'* @return ADODBExRecordset（Recordset）オブジェクト
'******************************************************************************
Function Clone(Optional LockType As LockTypeEnum = adLockUnspecified) As ADODBExRecordset
    Dim wrap As ADODBExRecordset: Set wrap = New ADODBExRecordset
    Set wrap.Origin = mOrigin.Clone(LockType)
    Set Clone = wrap
End Function

'******************************************************************************
'* [概  要] CloseRsメソッド
'* [詳  細] Closeのラッパーメソッド。
'*          開いているオブジェクトおよびそれに関連するすべてのオブジェクトを
'*          閉じる。
'*
'*          関連するすべてのシステム リソースを解放する場合に使用する。
'*          オブジェクトを閉じてもメモリからは削除されず、プロパティ設定を変更
'*          してもう一度開くことができる。
'*          オブジェクトをメモリから完全に削除するには、オブジェクトを閉じた後
'*          に、オブジェクト変数をNothingに設定する。
'*          オブジェクトを閉じると、関連するデータ、およびデータに対するその
'*          オブジェクトからの排他アクセスが、すべて解放される。
'*          後で Open メソッドを呼び出し、同じ属性で、または属性を変更して、
'*          オブジェクトを再度開くことができる。
'*          Recordset オブジェクトが閉じている間に､有効なカーソルを必要とする
'*          メソッドを呼び出すと､エラーが発生する｡
'*          即時更新モードで編集を行っているときに Close メソッドを呼び出すと、
'*          エラーが発生する。
'*          Update メソッドまたは CancelUpdate メソッドを先に呼び出す必要が
'*          ある。
'*          バッチ更新モードの実行中に Recordset オブジェクトを閉じると、
'*          最後に UpdateBatch を呼び出した後で行われた変更がすべて失われる。
'*          Clone メソッドを使用して、開いている Recordset オブジェクトの複製
'*          を作成した場合、元のオブジェクトまたはいずれかの複製を閉じても、
'*          他の複製には影響しない。
'******************************************************************************
Sub CloseRs()
    Call mOrigin.Close
End Sub

'******************************************************************************
'* [概  要] CompareBookmarksメソッド
'* [詳  細] CompareBookmarksのラッパーメソッド。
'*          2 つのブックマークを比較して､相対的な位置を示す値を返す。
'*
'*          ブックマークは、同じ Recordset オブジェクト、または Recordset オブ
'*          ジェクトとその 複製に適用する必要がある。
'*          異なる Recordset オブジェクトのブックマークを比較した場合、同じソー
'*          スまたはコマンドから作成されたブックマークであっても、信頼できる結
'*          果は得られない。
'*          また、基になるプロバイダーがブックマークの比較をサポートしていない
'*          Recordset オブジェクトの場合も、ブックマークを比較できない。
'*          ブックマークは、 Recordset オブジェクトの行を一意に識別する。
'*          カレント行のブックマークを取得するには、カレント行の Bookmark プロ
'*          パティを使用する。
'*          ブックマークのデータ型はプロバイダー固有であり、ADO ではバリアント
'*          (Variant) 型として公開される。
'*          たとえば、SQL Server のブックマークは型 DBTYPE_R8 (Double) である。
'*          ADO では、このデータ型は、サブタイプが倍精度浮動小数点型のバリアン
'*          ト型として表される。
'*          ブックマークを比較するとき、ADO はどのような種類の強制も試まない。
'*          値は、比較が行われるプロバイダーにそのまま渡される。
'*          CompareBookmarks メソッドに渡されるブックマークが異なる型の変数に
'*          格納されている場合は、"引数が間違った型、許容範囲外、または競合して
'*          います。" という型不一致エラーが発生する。
'*          無効なブックマークや、形式が正しくないブックマークは、エラーの原因に
'*          なる｡
'*
'* @param Bookmark1 最初の行のブックマーク
'* @param Bookmark2 2 番目の行のブックマーク
'* @return ブックマークで表される 2 つのレコードの相対的な行位置を示す
'*          CompareEnum 値
'******************************************************************************
Function CompareBookmarks(Bookmark1, Bookmark2) As CompareEnum
    CompareBookmarks = mOrigin.CompareBookmarks(Bookmark1, Bookmark2)
End Function

'******************************************************************************
'* [概  要] Deleteメソッド
'* [詳  細] Deleteのラッパーメソッド。
'*          カレント レコードまたはレコードのグループを削除する。
'*
'*          Delete メソッドを使用すると、 Recordset オブジェクト内のカレント
'*          レコードまたはレコードのグループは削除の対象としてマークされる。
'*          レコードを削除できない Recordset オブジェクトの場合は、エラーが発生
'*          する。
'*          即時更新モードでは、削除は直ちにデータベースに反映される。
'*          データベースの整合性違反などにより削除を実行できない場合、レコードは
'*          Update を呼び出した後も編集モードのままになる。
'*          そのため、 Close、Move、NextRecordset などによりカレント レコードから
'*          移動する前に、 CancelUpdate を使用して更新を取り消す必要がある。
'*          バッチ更新モードでは、キャッシュ内の削除されるレコードにマークはさ
'*          れるが、実際の削除は UpdateBatch メソッドを呼び出すまで行われない。
'*          削除されたレコードを参照するには、 Filter プロパティを使用する。
'*          削除されたレコードからフィールド値を取得すると、エラーが発生する。
'*          カレント レコードを削除すると、別のレコードに移動するまで、削除され
'*          たレコードがカレント レコードのままになる。
'*          削除されたレコードから移動した後は、そのレコードにアクセスできなく
'*          なる。
'*          トランザクションで削除レコードをネストしている場合は、RollbackTrans
'*          メソッドを使用して、削除されたレコードを復元することができる。
'*          バッチ更新モードの場合は、保留中の削除を CancelBatch メソッドで取り
'*          消すことができる。
'*          基になるデータとの競合 (レコードが別のユーザーによって既に削除され
'*          ているなど) が原因でレコードの削除が失敗した場合は、プロバイダー
'*          から Errors コレクションに警告が返されますが、プログラムの実行は
'*          停止されない。実行時エラーが発生するのは、要求したすべてのレコー
'*          ドで競合が発生した場合のみ。
'*          Unique Table ダイナミック プロパティが設定されていて、 Recordset
'*          が複数のテーブルに対して JOIN 操作を実行した結果である場合、
'*          Delete メソッドは、 Unique Table プロパティで指定されたテーブル
'*          からのみ行を削除できる。
'*
'* @param AffectRecords Deleteメソッドで操作するレコードの数を決める値。
'*                      adAffectAll と adAffectAllChapters は、 Delete では
'*                      無効な引数。
'******************************************************************************
Sub Delete(Optional AffectRecords As AffectEnum = adAffectCurrent)
    Call mOrigin.Delete(AffectRecords)
End Sub

'******************************************************************************
'* [概  要] Findメソッド
'* [詳  細] Findのラッパーメソッド。
'*          Recordset から、指定した条件を満たす行を検索する。
'*          必要に応じて、検索の方向、開始行、および開始行からのオフセットを
'*          指定できる。
'*          条件が一致すると、カレント行の位置は、検出されたレコードに設定され、
'*          条件を満たす行がない場合は、 Recordset の最後 (または最初) に設定
'*          される。
'*
'*          criteria には、列の名前を 1 つだけ指定できる。
'*          このメソッドでは、複数列の検索はサポートしている。
'*          条件の比較演算子には、">" (より大きい)、"<" (より小さい)、"=" (等
'*          しい)、">=" (より大きいまたは等しい)、<"=" (以下)、"<>" (等しくな
'*          い)、または "like" (パターンマッチング) のいずれかを指定できる。
'*          抽出条件の値には、文字列、浮動小数点数、または日付を指定できる。
'*          文字列型 (String) の値は、#一重引用符または "" (シャープ記号) マー
'*          クで区切られる (たとえば、"state = #'#WA '" または "state = WA")。
'*          日付の値は#、_ > ##"" (シャープ記号) マークで区切られており、タイ
'*          ムスタンプを示すのに時間、分、および秒を含めることができるが、
'*          ミリ秒またはエラーが発生することはない。
'*          比較演算子に "like"を使用する場合、文字列値にアスタリスク (*) を
'*          含めると、1 つまたは複数の文字、または部分文字列を検索することが
'*          できる。たとえば、「state like 'M*'」と指定すると、Maine や
'*          Massachusetts が該当する。
'*          また、文字列の先頭と末尾にアスタリスクを使用して、その間に含まれる
'*          部分文字列を検索対象として指定することもできる。
'*          たとえば、「state like '*as*'」と指定すると、Alaska、Arkansas、
'*          および Massachusetts が該当する。
'*          上の例のように、アスタリスクは、検索文字列の末尾のみに使用するか、
'*          または検索文字列の先頭と末尾の両方で使用することができる。
'*          ただし、先頭のみのワイルドカード ('*str') または文字列中のワイルド
'*          カード ('s*r') としてアスタリスクを使用することはできない。
'*          この場合はエラーが発生する。
'*          [!メモ] Find メソッドを呼び出す前にカレント行の位置が設定されて
'*          いない場合は、エラーが発生する。
'*          Find メソッドを呼び出す前に、 MoveFirst などの、行の位置を設定する
'*          メソッドを呼び出す必要がある。
'*          [!メモ] レコードセットに対して Find メソッドを呼び出す場合で、
'*          レコードセット内の現在の位置が最後のレコードまたはファイルの最
'*          後 (EOF) の場合、何も検索されない。
'*          あらかじめ MoveFirst メソッドを呼び出して、現在の位置またはカー
'*          ソルをレコードセットの先頭に設定する必要がある。
'*
'* @param Criteria        検索に使用する列の名前、比較演算子、および値を指定す
'*                        るステートメントを含む文字列型 ( String ) の値を指定
'*                        する。
'* @param SkipRecords     省略可能。 既定値が0で、検索を開始するために現在の行
'*                        または開始ブックマークからの行のオフセットを指定する
'*                        値。デフォルトでは、検索は現在の行から開始される。
'* @param SearchDirection 省略可能。現在の行から検索を開始するか、検索の方向で
'*                        次に利用可能な行から検索を開始するかを指定する値。
'*                        値がadSearchForwardの場合、レコードセットの最後で失
'*                        敗した検索が停止する。
'*                        値がadSearchBackwardの場合、失敗した検索はレコード
'*                        セットの先頭で停止する。
'* @param Start           省略可能。検索の開始位置として使用する、バリアント
'*                        型 ( Variant ) のブックマークを指定。
'******************************************************************************
Sub Find(Criteria As String, _
         Optional SkipRecords As Long, _
         Optional SearchDirection As SearchDirectionEnum = adSearchForward, _
         Optional Start)
    If IsMissing(Start) Then
        Call mOrigin.Find(Criteria, SkipRecords, SearchDirection)
    Else
        Call mOrigin.Find(Criteria, SkipRecords, SearchDirection, Start)
    End If
End Sub

'******************************************************************************
'* [概  要] GetRowsメソッド
'* [詳  細] GetRowsのラッパーメソッド。
'*          Recordset オブジェクトの複数のレコードを配列に取り込む。
'*
'*          GetRows メソッドでは、Recordset のレコードを 2 次元配列にコピー
'*          する。
'*          1 番目の添え字でフィールドを指定し、2 番目の添え字でレコード番号を
'*          指定する。
'*          GetRows メソッドからデータが返されると、array 変数は自動的に適切な
'*          サイズに調整される。
'*          Rows 引数の値を指定しないと、GetRows メソッドは Recordset オブジェ
'*          クトのすべてのレコードを自動的に取得する。
'*          要求したレコード数が使用可能なレコード数より多い場合、GetRows は使
'*          用可能なレコード数だけを返す。
'*          Recordset オブジェクトがブックマークをサポートしている場合は、Start
'*          引数にレコードの Bookmark プロパティの値を渡して、GetRows メソッド
'*          の取得開始レコードを指定できる。
'*          GetRows の呼び出しで返されるフィールドを制限するには、Fields 引数に
'*          単一のフィールド名または番号、あるいは複数のフィールド名または番号
'*          の配列を渡す。
'*          GetRows を呼び出すと、まだ読み込まれていない次のレコードがカレント
'*          レコードになるか、それ以上レコードがない場合は EOF プロパティが
'*          True に設定される。
'*
'* @param Rows   省略可能。取得するレコード数。
'*               adGetRowsRest (-1)を指定した場合、指定位置から残りのレコード
'*               を取得｡
'* @param Start  省略可能。GetRows 操作を開始するレコードへのブックマークとして
'*               評価される文字列型 ( String ) の値またはバリアント型 ( Variant )
'*               を指定。 以下のBookmarkEnum 値を使用することも可能。
'*                 adBookmarkCurrent(0) … カレントレコード
'*                 adBookmarkFirst(1)   … 最初のレコード
'*                 adBookmarkLast(2)    … 最後のレコード
'* @param Fields 省略可能。取得するフィールド名または配列。
'* @return 2次元配列データ（Recordsetオブジェクト内データ）
'******************************************************************************
Function GetRows(Optional Rows As Long = -1, _
                 Optional Start, _
                 Optional Fields)
    If IsMissing(Start) Then
        If IsMissing(Fields) Then
            GetRows = mOrigin.GetRows(Rows)
        Else
            GetRows = mOrigin.GetRows(Rows, , Fields)
        End If
    ElseIf IsMissing(Fields) Then
        GetRows = mOrigin.GetRows(Rows, Start)
    Else
        GetRows = mOrigin.GetRows(Rows, Start, Fields)
    End If
End Function

'******************************************************************************
'* [概  要] GetStringメソッド
'* [詳  細] GetStringのラッパーメソッド。
'*          Recordset を文字列として返却する｡
'*
'*          文字列には行データが保存され、スキーマ データは保存されない。
'*          したがって、この文字列を使用して Recordset を再度開くことはできない。
'*          このメソッドは、RDO の GetClipString メソッドと同等です。
'*
'* [補  足] 使用例： object.GeString(, 5, ",", vbCrLf)
'*
'* @param StringFormat    文字列を変換する方法。
'*                        ColumnDelimiter、RowDelimiterおよびNullExprの各パラメ
'*                        ータは、StringFormatにadClipStringを指定した場合にの
'*                        み使用可能。
'*                        StringFormatEnumには、adClipString(2)以外定義されてい
'*                        ないので、実質adClipString一択であり、省略して問題な
'*                        いと思われる。
'* @param NumRows         取得レコード数。
'*                        指定しない場合、またはRecordsetの行数の合計より大きい
'*                        値を指定した場合は、Recordsetのすべての行が変換される。
'* @param ColumnDelimeter フィールド区切り文字。
'*                        指定しなければ、タブが使用される。
'* @param RowDelimeter    レコード区切り文字。
'*                        指定しなければ、復帰改行(CR)が使用される。
'* @param NullExpr        Null値に入る文字。
'*                        指定しなければ、空文字列が使用される。
'* @return 文字列データ（Recordsetオブジェクト内データ）
'******************************************************************************
Function GetString(Optional StringFormat As StringFormatEnum = adClipString, _
                   Optional NumRows As Long = -1, _
                   Optional ColumnDelimeter As String, _
                   Optional RowDelimeter As String, _
                   Optional NullExpr As String) As String
    GetString = mOrigin.GetString(StringFormat, NumRows, ColumnDelimeter, RowDelimeter, NullExpr)
End Function

'******************************************************************************
'* [概  要] Moveメソッド
'* [詳  細] Moveのラッパーメソッド。
'*          Recordset オブジェクトでカレント レコードの位置を移動する。
'*
'*          Move メソッドは、すべての Recordset オブジェクトでサポートされている。
'*          NumRecords 引数がゼロより大きい場合、現在のレコードの位置は前方
'*          (Recordset の末尾方向) に移動する。
'*          NumRecords がゼロより小さい場合、現在のレコードの位置は後方 (Recordset
'*          の先頭方向) に移動する。
'*          moveメソッドを呼び出して、カレントレコードの位置を最初のレコードの
'*          前に移動すると、ADO は現在のレコードを recordset の最初のレコードの
'*          前の位置 (BOFがTrue) に設定する。
'*          BOF プロパティが既に True の場合、後方へ移動しようとすると、エラー
'*          が発生する。
'*          Move メソッドを呼び出してカレント レコードの位置を最後のレコードの
'*          後に移動しようとすると、カレント レコードがレコードセットの最後の
'*          レコードの後に設定され、EOF が True になる。
'*          EOF プロパティが既に True の場合、前方へ移動しようとすると、エラー
'*          が発生する。
'*          空の Recordset オブジェクトから Move メソッドを呼び出すと、エラーが
'*          発生する。
'*          Start 引数を指定した場合、Recordset オブジェクトではブックマークが
'*          サポートされていると見なされ、このブックマークを持つレコードが移動
'*          の基準となる。
'*          指定しない場合は、カレントレコードが移動の基準となる。
'*          CacheSize プロパティを使用してプロバイダーからのレコードをローカル
'*          にキャッシュしている場合、NumRecords 引数を渡して、現在キャッシュ
'*          されているレコードグループの範囲外に現在のレコード位置を移動すると
'*          移動先のレコードから始まる新規レコード グループが強制的に取得される。
'*          CacheSize プロパティが、新規に取得されるグループのサイズを決定し、
'*          移動先のレコードが最初に取得されるレコードになる。
'*          Recordset オブジェクトが前方スクロールのみ可能な場合でも、移動先が
'*          現在キャッシュされているレコードセットの範囲内であれば、NumRecords
'*          引数に 0 より小さい値を指定できる。
'*          Move メソッドを呼び出して、カレント レコードの位置を、キャッシュ
'*          されている最初のレコードより前のレコードに移動しようとすると、
'*          エラーが発生する。
'*          このように、前方スクロールのみをサポートするプロバイダーで、完全
'*          スクロールをサポートするレコード キャッシュを使用することができる。
'*          キャッシュされたレコードはメモリに読み込まれるため、必要以上のレ
'*          コードのキャッシュは避けること。
'*          前方スクロールのみ可能な Recordset オブジェクトでも、この方法で
'*          後方への移動をサポートできるが、前方スクロールのみ可能な Recordset
'*          オブジェクトで MovePrevious メソッドを呼び出すと、エラーが発生する。
'*          [!メモ] 前方スクロールのみ可能な Recordset で後方への移動がサポー
'*          トされているかどうかは、プロバイダーによる。
'*          カレント レコードが Recordset の最後のレコードの後にある場合、後
'*          方への Move を実行しても現在の位置が正しく示されないことがある。
'*
'* @param NumRecords カレントレコードの位置を移動するレコード数を指定。
'* @param Start      省略可能。ブックマークとして評価される文字列型 の値
'*                   またはバリアント型を指定。 BookmarkEnum 値を使用する
'*                   ことも可能。
'******************************************************************************
Sub Move(NumRecords As Long, _
         Optional Start)
    If IsMissing(Start) Then
        Call mOrigin.Move(NumRecords)
    Else
        Call mOrigin.Move(NumRecords, Start)
    End If
End Sub

'******************************************************************************
'* [概  要] MoveFirstメソッド
'* [詳  細] MoveFirstのラッパーメソッド。
'*          指定された Recordset オブジェクトの最初のレコードに移動して、その
'*          レコードをカレントレコードにする。
'*
'*          MoveFirst メソッドを使用して、カレント レコードの位置を Recordset
'*          の最初のレコードに移動する。
'*          Recordset が空 ( BOF と EOF の両方が True) の場合、 MoveFirst を
'*          呼び出すとエラーが発生する。
'*          前方スクロールのみ可能な Recordset で両方向スクロールのサポートが
'*          必要な場合、 CacheSize プロパティを使用して後方へのカーソル移動を
'*          サポートするレコード キャッシュを作成し、 Move メソッドを使用して
'*          移動することができる。
'*          キャッシュされたレコードはメモリに読み込まれるため、必要以上のレコ
'*          ードのキャッシュは避けること。
'*          前方スクロールのみ可能な Recordset オブジェクトで MoveFirst メソッ
'*          ドを呼び出すことはできるが、その場合、 Recordset オブジェクトを
'*          生成するコマンドがプロバイダーで再度実行されることがある。
'*
'******************************************************************************
Sub MoveFirst()
    Call mOrigin.MoveFirst
End Sub

'******************************************************************************
'* [概  要] MoveLastメソッド
'* [詳  細] MoveLastのラッパーメソッド。
'*          指定された Recordset オブジェクトの最後のレコードに移動して、その
'*          レコードをカレントレコードにする。
'*
'*          MoveLast メソッドを使用して、カレント レコードの位置を Recordset
'*          の最後のレコードに移動する。
'*          Recordset オブジェクトはブックマークまたは後方へのカーソル移動を
'*          サポートしている必要があり、サポートしていない場合、このメソッドを
'*          呼び出すとエラーが発生する。
'*          Recordset が空 ( BOF と EOF の両方が True) の場合、MoveLast を呼び
'*          出すとエラーが発生する。
'*
'******************************************************************************
Sub MoveLast()
    Call mOrigin.MoveLast
End Sub

'******************************************************************************
'* [概  要] MoveNextメソッド
'* [詳  細] MoveNextのラッパーメソッド。
'*          指定された Recordset オブジェクトの次のレコードに移動して、その
'*          レコードをカレントレコードにする。
'*
'*          MoveNext メソッドを使用して、カレント レコードの位置を 1 レコード
'*          前方 (Recordset の終端方向) に移動する。
'*          カレント レコードが最後のレコードの場合に MoveNext メソッドを呼び
'*          出すと、カレント レコードが Recordset の最後のレコードの後に設定
'*          され、EOF が True になる。
'*          EOF プロパティが既に True の場合、前方へ移動しようとするとエラーが
'*          発生する。
'*          フィルターまたはソートが実行された Recordset でカレント レコードの
'*          データを変更すると、その位置も変更されることがある。
'*          その場合、 MoveNext メソッドは正常に動作するが、その位置が、元の
'*          位置からではなく、変更後の位置から 1 レコード先に移動されることに
'*          注意する必要がある。
'*
'******************************************************************************
Sub MoveNext()
    Call mOrigin.MoveNext
End Sub

'******************************************************************************
'* [概  要] MovePreviousメソッド
'* [詳  細] MovePreviousのラッパーメソッド。
'*          指定された Recordset オブジェクトの前のレコードに移動して、その
'*          レコードをカレントレコードにする。
'*
'*          MovePrevious メソッドを使用して、カレント レコードの位置を 1 レコ
'*          ード後方 (Recordset の始端方向) に移動する。
'*          Recordset オブジェクトはブックマークまたは後方へのカーソル移動を
'*          サポートしている必要があり、サポートしていない場合、このメソッドを
'*          呼び出すとエラーが発生する。
'*          カレント レコードが最初のレコードの場合に MovePrevious メソッドを
'*          呼び出すと、カレント レコードは Recordset の最初のレコードの前に
'*          設定され、BOF が True になる。
'*          BOF プロパティが既に True の場合、後方へ移動しようとすると、エラー
'*          が発生する。
'*          Recordset オブジェクトがブックマークまたは後方へのカーソル移動を
'*          サポートしていない場合、MovePrevious メソッドを実行すると、エラー
'*          が発生する。
'*
'******************************************************************************
Sub MovePrevious()
    Call mOrigin.MovePrevious
End Sub

'******************************************************************************
'* [概  要] NextRecordsetメソッド
'* [詳  細] NextRecordsetのラッパーメソッド。
'*          現在の Recordset オブジェクトをクリアし、一連のコマンド操作を実行し
'*          て次の Recordset を返却する。
'*
'*          NextRecordset メソッドでは、複合コマンドステートメントの次のコマン
'*          ドの結果、または複数の結果を返すストアドプロシージャの結果を返す。
'*          複合コマンドステートメントに基づいてRecordsetオブジェクトを開いた
'*          場合 (たとえば、"select * from table1; select * from table2"
'*          コマンドに対してExecuteメソッドまたはrecordsetのOpenメソッドを
'*          使用する)、最初のコマンドのみが実行され、結果がrecordsetに返され
'*          る。ステートメントの次のコマンドの結果にアクセスするには、
'*          NextRecordset メソッドを呼び出す。
'*          その他の結果があり、複合ステートメントを含むRecordsetがプロセス境界
'*          を越えて切断またはマーシャリングされていない限り、 NextRecordsetメ
'*          ソッドは引き続きrecordsetオブジェクトを返す。
'*          行を返すコマンドが正常に実行されてもレコードが返されない場合は、返
'*          されたRecordsetオブジェクトが空になる。
'*          この場合は、 BOFプロパティとEOFプロパティが両方ともTrueであることを
'*          確認してテストする。
'*          行を返さないコマンドが正常に実行された場合は、返されたrecordsetオブ
'*          ジェクトが閉じられ、 recordsetのStateプロパティをテストすることによ
'*          って確認できる。
'*          結果がなければ、 recordsetはNothingに設定される。
'*          切断された Recordset オブジェクトでは、 ActiveConnection が Nothing
'*          (Microsoft Visual Basic) または NULL (その他の言語) に設定されてい
'*          るため、 NextRecordset メソッドは使用できない。
'*          即時更新モードで編集を行っているときに NextRecordset メソッドを呼び
'*          出すと、エラーが発生する。
'*          Update メソッドまたは CancelUpdate メソッドを先に呼び出す必要がある。
'*          複合ステートメントの複数のコマンドにパラメーターを渡すには、
'*          Parameters コレクションを使用するか、または元の Open メソッドまたは
'*          Execute メソッドの呼び出しで配列を渡す。
'*          この場合、コレクションまたは配列のパラメーターの並び順は、一連の
'*          コマンドの並び順と同じである必要がある。
'*          出力パラメーターの値を読み込むには、すべての結果の読み込みが終了
'*          している必要がある。
'*          複合ステートメント内の各コマンドをいつ実行するかは、OLE DB プロバイ
'*          ダーが決定する。
'*          たとえば、Microsoft OLE DB Provider for SQL Server は、複合ステート
'*          メントを受け取ったときにすべてのコマンドをバッチで実行する。
'*          結果の Recordsets は、 NextRecordset を呼び出したときに返される。
'*          ただし、それ以外のプロバイダーでは、ステートメント内の次のコマンド
'*          が NextRecordset の呼び出しの後にしか実行されないこともある。
'*          そのようなプロバイダーの場合、ステートメントのすべてのコマンドを
'*          実行する前に Recordset オブジェクトを明示的に閉じると、残りのコマ
'*          ンドは実行されない。
'*
'* @param RecordsAffected 省略可能。長整数型 ( Long ) の変数を指定。
'*                        プロバイダーは、操作の影響を受けたレコード数をここに
'*                        返却する。このパラメーターは、操作の影響を受けたレコ
'*                        ード数のみを返す。Recordset を作成するために使用された
'*                        Select ステートメントからのレコード数を返すことはない｡
'* @return ADODBExRecordset（Recordset）オブジェクト
'******************************************************************************
Function NextRecordset(Optional RecordsAffected) As ADODBExRecordset
    Dim wrap As ADODBExRecordset: Set wrap = New ADODBExRecordset
    Set wrap.Origin = mOrigin.NextRecordset(RecordsAffected)
    Set NextRecordset = wrap
End Function

'******************************************************************************
'* [概  要] OpenRsメソッド
'* [詳  細] Openのラッパーメソッド。
'*          カーソルを開く｡
'*
'*          ADO の Recordset の既定のカーソルは、サーバー側にある読み取り専用の
'*          前方スクロール カーソルである。
'*          Recordset オブジェクトで Open メソッドを使用すると、ベース テーブル
'*          からのレコード、クエリの結果、または以前に保存された Recordset を表
'*          すカーソルを開くことができる。
'*          省略可能なsource引数を使用して、 Commandオブジェクト変数、SQL ステー
'*          トメント、ストアドプロシージャ、テーブル名、URL、完全なファイルパス
'*          名のいずれかを使用して、データソースを指定する。
'*          Sourceがファイルパス名の場合は、相対パスで指定できる。
'*          Open メソッドの Source 引数を使用して、レコードを返さないアクション
'*          クエリを実行するのは、お勧めできない。
'*          これは、呼び出しが成功したかどうかを容易に判断できないためである。
'*          このようなクエリによって返された Recordset は閉じられる。
'*          SQL INSERT ステートメントのように、レコードを返さないクエリを実行
'*          する場合は、代わりに Command オブジェクトの Execute メソッドまたは
'*          Connection オブジェクトの Execute メソッドを呼び出すこと。
'*          ActiveConnection 引数は ActiveConnection プロパティに対応し、
'*          Recordset オブジェクトを開く接続を指定する。この引数に接続の定義
'*          を指定すると、ADO により、指定されたパラメーターを使用して新しい接続
'*          が開かれる。
'*          クライアント側カーソル (cursor location = adUseClient) を使用して
'*          Recordsetを開いた後、このプロパティの値を変更して、別のプロバイダー
'*          に更新を送信することができる。
'*          また、このプロパティを Nothing (Microsoft Visual Basic の場合) また
'*          は NULL に設定すると、その Recordset を任意のプロバイダーから切断す
'*          ることができる。
'*          ただし、サーバー側カーソルの ActiveConnection を変更すると、エラー
'*          が発生する。
'*          Recordset オブジェクトのプロパティに直接対応するその他の引数 (Source、
'*          CursorType、および LockType) と、プロパティの関係は次のとおり。
'*            ・Recordset オブジェクトを開く前のプロパティは､読み取り / 書き込み
'*              可能｡
'*            ・Open メソッドの実行時に、対応する引数を指定しないと、プロパティ
'*              の設定値が使用される。引数を指定した場合は、対応するプロパティの
'*              設定値が上書きされて引数の値に更新される。
'*            ・Recordset オブジェクトを開いた後､プロパティは読み取り専用になる｡
'*          Recordset オブジェクトの Source プロパティに有効な Command オブジェ
'*          クトが設定されている場合、その Recordset オブジェクトが開いていなく
'*          ても、ActiveConnection プロパティは読み取り専用になる。
'*          Source 引数に Command オブジェクトを指定する場合、同時に
'*          ActiveConnection 引数を指定すると、エラーが発生する。Command オブ
'*          ジェクトの ActiveConnection プロパティは、あらかじめ有効な Connection
'*          オブジェクトまたは接続文字列に設定しておく必要がある。
'*          Source 引数に Command オブジェクト以外のソースを指定する場合は、
'*          Options 引数を使用して Source 引数の評価を最適化することができる。
'*          Options 引数が定義されていない場合は、パフォーマンスが低下する可能
'*          性がある。
'*          これは、引数が SQL ステートメント、ストアド プロシージャ、URL、テー
'*          ブル名のいずれであるかを判断するために、プロバイダーを呼び出す必要が
'*          あるためである。
'*          使用する Source の種類がわかっている場合は、Options 引数を設定する
'*          ことにより、該当するコードに直接ジャンプすることができる。
'*          Options 引数が Source の種類と一致しない場合は、エラーが発生する。
'*          Source 引数に Stream オブジェクトを指定する場合は、他の引数は指定
'*          しないこと。
'*          そのようにすると、エラーが発生する。
'*          Stream から Recordset を開く場合は、ActiveConnection の情報は保持
'*          されない。
'*          Recordset に接続が関連付けられていない場合、Options 引数の既定値は
'*          adCmdFile である。
'*          代表的な例では、永続的に保存された Recordset オブジェクトの場合が
'*          これに該当する。
'*          データ ソースがレコードを返さない場合、プロバイダーは BOF プロパテ
'*          ィと EOF プロパティを True に設定し、カレント レコードの位置が未定義
'*          になる。
'*          カーソルの種類によっては、この空の Recordset オブジェクトに新しい
'*          データを追加することができる。
'*          開いている Recordset に対する操作が完了したら、Close メソッドを使用
'*          して、関連するシステム リソースをすべて解放する。
'*          オブジェクトを閉じてもメモリからは削除されないので、そのオブジェク
'*          トのプロパティを変更したり、Open メソッドを使用してもう一度開いたり
'*          することができる。
'*          オブジェクトをメモリから完全に削除するには、オブジェクト変数に
'*          Nothing を設定する。
'*          ActiveConnectionプロパティを設定する前に、オペランドなしでOpenを
'*          呼び出して、 recordset fieldsコレクションにフィールドを追加する
'*          ことによって作成されたrecordsetのインスタンスを作成する。
'*          CursorLocation プロパティを adUseClient に設定している場合は、非
'*          同期に行を取得できる方法が 2 つある。
'*          推奨される方法は、Options を adAsyncFetch に設定することである。
'*          もう 1 つの方法として、 Properties コレクションで "非同期行セット
'*          処理" ダイナミック プロパティを使用することもできるが、
'*          Options パラメーターを adAsyncFetch に設定していない場合は、関連
'*          して取得されたイベントが失われる可能性がある。
'*          MS Remote プロバイダーのバックグラウンド フェッチは、Open メソッ
'*          ドの Options パラメーターを通してのみサポートされる。
'*          [!メモ] http スキームを使用している URL は、Microsoft OLE DB
'*          Provider for Internet Publishing を自動的に呼び出す。
'*
'* @param Source           省略可能。有効な Command オブジェクト、SQL ステート
'*                         メント、テーブル名、ストアド プロシージャの呼び出し、
'*                         URL、または永続的に保存された Recordset を格納して
'*                         いるファイル名や Stream オブジェクト名に評価される
'*                         バリアント型 ( Variant ) の値を指定。
'* @param ActiveConnection 省略可能。有効な Connection オブジェクト
'*                         または ConnectionString パラメーターの値を指定。
'* @param CursorType       省略可能。プロバイダーが Recordset を開くときに使用
'*                         するカーソルの種類を決める CursorTypeEnum 値を指定。
'* @param LockType         省略可能。プロバイダーが Recordset を開くときに使用
'*                         するロック (同時作用) の種類を決めるための値を指定。
'* @param Options          省略可能。Source 引数が Command オブジェクト以外の
'*                         ソースを表す場合にプロバイダーがこの引数を評価する
'*                         方法、または以前に保存されていたファイルから
'*                         Recordset を復元する必要があることを示す値を指定。
'*                         1 つまたは複数の CommandTypeEnum 値または
'*                         ExecuteOptionEnum 値を指定できる。
'*                         これらの値は、ビット単位の AND 演算子で組み合わせ
'*                         て使用することができる。
'*                         永続化された Recordset を含む Stream から Recordset
'*                         を開く場合、ExecuteOptionEnum 値の
'*                         adAsyncFetchNonBlocking を使用しても無効である。
'*                         この場合、フェッチは同期であり、ブロッキングが行わ
'*                         れる。
'*                         Open メソッドを使用する場合は、ExecuteOpenEnum に
'*                         adExecuteNoRecords または adExecuteStream を指定
'*                         しないでこと。
'******************************************************************************
Sub OpenRs(Optional Source, _
           Optional ActiveConnection, _
           Optional CursorType As CursorTypeEnum = adOpenUnspecified, _
           Optional LockType As LockTypeEnum = adLockUnspecified, _
           Optional Options As Long = -1)
    If IsMissing(Source) Then
        If IsMissing(ActiveConnection) Then
            Call mOrigin.Open(, , CursorType, LockType, Options)
        Else
            If TypeName(ActiveConnection) = "ADODBExConnection" Then
                Call mOrigin.Open(, ActiveConnection.Origin, CursorType, LockType, Options)
            Else
                Call mOrigin.Open(, ActiveConnection, CursorType, LockType, Options)
            End If
        End If
    ElseIf IsMissing(ActiveConnection) Then
        Call mOrigin.Open(Source, , CursorType, LockType, Options)
    Else
        If TypeName(ActiveConnection) = "ADODBExConnection" Then
            Call mOrigin.Open(Source, ActiveConnection.Origin, CursorType, LockType, Options)
        Else
            Call mOrigin.Open(Source, ActiveConnection, CursorType, LockType, Options)
        End If
    End If
End Sub

'******************************************************************************
'* [概  要] Requeryメソッド
'* [詳  細] Requeryのラッパーメソッド。
'*          オブジェクトの基になるクエリを再実行して､Recordset オブジェクトの
'*          データを更新する｡
'*
'*          Requery メソッドは、元のコマンドを再実行してデータをもう一度取得する
'*          ことにより、データ ソースで Recordset オブジェクトの内容全体を更新
'*          する場合に使用する。
'*          このメソッドの呼び出しは、Close メソッドと Open メソッドを連続して
'*          呼び出すことと同じである。
'*          カレント レコードを編集しているとき、または新規レコードを追加してい
'*          るときにこのメソッドを呼び出すと、エラーが発生する。
'*          Recordset オブジェクトを開いている間は、カーソルの属性を定義するプロ
'*          パティ (CursorType、LockType、MaxRecords など) は、読み取り専用にな
'*          る。したがって、Requery メソッドでは、現在のカーソルの更新のみを行う
'*          ことができる。
'*          カーソルのプロパティを変更して結果を確認するには、Close メソッドを
'*          使用して、プロパティを再度読み取り/書き込み可能にする必要がある。
'*          その後、目的のプロパティの設定を変更し、Open メソッドを呼び出して
'*          カーソルを再度開く。
'*          オプションがadasyncexecuteに設定されている場合､この操作は非同期で
'*          実行され､ RecordsetChangeCompleteイベントが終了すると発行される｡
'*          Requery では、ExecuteOpenEnum の値 adExecuteNoRecords または
'*          adExecuteStream は使用できない。
'*
'* @param Options 省略可能。この操作に対する ExecuteOptionEnum 値と
'*                CommandTypeEnum 値を含む､ビットマスクを指定。
'******************************************************************************
Sub Requery(Optional Options As Long = -1)
    Call mOrigin.Requery(Options)
End Sub


'******************************************************************************
'* [概  要] Resyncメソッド
'* [詳  細] Resyncのラッパーメソッド。
'*          現在の Recordset オブジェクト、または Record オブジェクトの Fields
'*          コレクションのデータを､基になるデータベースのデータで更新する。
'*          Resync メソッドは、現在の Recordset のレコードを基になるデータベース
'*          と再同期させる場合に使用する。
'*          このメソッドは、静的または前方専用のカーソルを使用していて、基になる
'*          データベースの変更を確認する場合に役立つ。
'*          CursorLocation プロパティを adUseClient に設定した場合には、 Resync
'*          は、読み取り専用ではない Recordset オブジェクトに対してのみ使用でき
'*          る。
'*          Requery メソッドとは異なり、 Resync メソッドは、 Recordset オブジェ
'*          クトの基になっているコマンドを再実行しない。
'*          基になるデータベースの新しいレコードを参照することはできない。
'*          基になるデータとの競合 (たとえば、他のユーザーがレコードを削除した
'*          場合) が原因で再同期に失敗した場合、プロバイダーは Errors コレクショ
'*          ンに警告を返し、実行時エラーが発生する。
'*          競合しているレコードを特定するには、Filter プロパティ
'*          (adFilterConflictingRecords) と Status プロパティを使用する。
'*          ダイナミック プロパティ Unique Table および Resync Command が設定
'*          されていて、 Recordset が複数のテーブルに対する JOIN 操作の実行結果
'*          である場合、 Resync メソッドは、 Unique Table プロパティで指定され
'*          ているテーブルに対してのみ、 Resync Command プロパティで指定されて
'*          いるコマンドを実行する。
'*
'* @param AffectRecords 省略可能。Resyncメソッドで操作するレコードの数。
'* @param ResyncValues  省略可能。基になる値を上書きするかどうかを指定。
'*
'******************************************************************************
Sub Resync(Optional AffectRecords As AffectEnum = adAffectAll, _
           Optional ResyncValues As ResyncEnum = adResyncAllValues)
    Call mOrigin.Resync(AffectRecords, ResyncValues)
End Sub

'******************************************************************************
'* [概  要] Saveメソッド
'* [詳  細] Saveのラッパーメソッド。
'*          Recordset をファイルまたは Stream オブジェクトに保存する。
'*
'*          Save メソッドは、開いている Recordset でのみ呼び出すことができる。
'*          保存した Recordset を Destination から復元するには、Open メソッド
'*          を使用する。
'*          Filter プロパティが Recordset に対して有効な場合は、フィルターで
'*          アクセスできる行のみが保存される。
'*          Recordset が階層の場合は、親 Recordset を含めて、現在の子 Recordset
'*          とその子が保存される。
'*          子 Recordset の Save メソッドを呼び出すと、子と、そのすべての子は
'*          保存されるが、親は保存されない。
'*          Recordset メソッドを初めて保存するとき、Destination は指定しても
'*          しなくてもかまわない。
'*          Destination を省略すると、Recordset の Source プロパティの値に設定
'*          された名前を使用して、新規ファイルが作成される。
'*          初めて保存した後に、続けて Save を呼び出すときは、Destination を
'*          指定しないこと。指定すると、実行時エラーが発生する。
'*          Save メソッドを続けて呼び出すときに新しい Destination を指定すると、
'*          Recordset は新しい保存先に保存される。
'*          ただし、その場合、新しい保存先と元の保存先の両方が開いた状態になる。
'*          Save メソッドは、Recordset または Destination を閉じない。
'*          したがって、Recordset の操作を続行し、最新の変更を保存することが
'*          できる。Recordset を閉じるまで、Destination は開いたままになる。
'*          セキュリティ上の理由により、 Save メソッドは、Microsoft Internet
'*          Explorer で実行されるスクリプトからは、低レベルおよびカスタムの
'*          セキュリティ設定の使用時のみ許可される。
'*          非同期の Recordset のフェッチ、実行、または更新の操作中に Save
'*          メソッドが呼び出された場合、 Save メソッドは、非同期操作が完了する
'*          まで待機する。
'*          レコードは、 Recordset の最初の行から順に保存される。
'*          Save メソッドが終了すると、カレント行の位置は Recordset の最初の
'*          行になる。
'*          最良の結果を得るには、 Save メソッドの実行時に CursorLocation プロ
'*          パティを adUseClient に設定する。
'*          Recordset オブジェクトを保存するために必要なすべての機能が、ご使用
'*          のプロバイダーによってサポートされていない場合は、Cursor Service
'*          によりその機能が提供される。
'*          Recordset の CursorLocation プロパティが adUseServer に設定された
'*          ままの場合、その Recordset の更新機能が制限される。
'*          プロバイダーの機能によって異なるが、通常は、単一テーブルの更新、
'*          挿入、および削除のみが許可される。
'*          この設定では、 Resync メソッドも利用できない。
'*          [!メモ] ADO では、 Fields の種類が adVariant 、 adIDispatch 、また
'*          は adIUnknown に設定された Recordset の保存はサポートされていない
'*          ため、予期しない結果が生じることがある。
'*          条件文字列の形式 (たとえば、OrderDate > ' 12/31/1999 ') のフィル
'*          ターのみが、永続化されたRecordsetの内容に影響する。
'*          Bookmarks の配列で作成されたフィルター、または FilterGroupEnum の
'*          値を使用して作成されたフィルターは、永続化された Recordset の内容
'*          に影響しない。
'*          これらの規則は、クライアント側カーソルまたはサーバー側カーソルを
'*          使用して作成された Recordsets に当てはまる。
'*          Destination パラメーターには、OLE DB IStream インターフェイスを
'*          サポートするすべてのオブジェクトを指定できるので、Recordset を
'*          ASP Response オブジェクトに直接保存することができる。
'*          次の Visual Basic コードに示すように、XML 形式の Recordset を
'*          MSXML DOM オブジェクトのインスタンスに保存することもできる。
'*          [!メモ] 階層 Recordset (データ シェイプ) を XML 形式で保存するとき
'*          は、2 つの制限事項がある。階層 Recordset に保留中の更新が含まれて
'*          いる場合は、XML で保存できない。
'*          また、パラメーター化された階層 Recordset を保存することはできない。
'*          XML 形式で保存した Recordset は、UTF-8 形式を使用して保存される。
'*          このようなファイルを ADO Stream に読み込むと、Stream オブジェクト
'*          はストリームの Charset プロパティが UTF-8 形式用の適切な値に設定
'*          されていない限り、ストリームから Recordset を開こうとしない。
'*
'* @param Destination   省略可能。Recordset の保存先であるファイルの完全なパス
'*                      名を表すバリアント型の値、または Stream オブジェクトへ
'*                      の参照を指定。
'* @param PersistFormat 省略可能。Recordset の保存形式 (XML または ADTG) を
'*                      PersistFormatEnum 値で指定。
'******************************************************************************
Sub Save(Optional Destination, _
         Optional PersistFormat As PersistFormatEnum = adPersistADTG)
    If IsMissing(Destination) Then
        Call mOrigin.Save(, PersistFormat)
    Else
        Call mOrigin.Save(Destination, PersistFormat)
    End If
End Sub

'******************************************************************************
'* [概  要] SeekRsメソッド
'* [詳  細] Seekのラッパーメソッド。
'*          Recordset のインデックスを検索して､指定された値と一致する行をすば
'*          やく探し､その行をカレント行にする｡
'*
'*          基になるプロバイダーが Recordset オブジェクトのインデックスをサポー
'*          トしている場合、Index プロパティと共に Seek メソッドを使用する。
'*          基になるプロバイダーが Seek をサポートしているかどうかを判別するに
'*          は、Supports(adSeek) メソッドを使用する。
'*          プロバイダーがインデックスをサポートしているかどうかを判別するには、
'*          Supports(adIndex) メソッドを使用する。
'*          たとえば、OLE DB Provider for Microsoft Jet は Seek および Index を
'*          サポートしている。
'*          Seek メソッドで目的の行が見つからない場合、エラーは発生せず、行は
'*          Recordset の末尾に配置される。
'*          このメソッドを実行する前に、 Index プロパティを必要なインデックスに
'*          設定すること。
'*          このメソッドは、サーバー側のカーソルでのみサポートされる。 Recordset
'*          オブジェクトの CursorLocation プロパティの値が adUseClient になって
'*          いる場合は、Seek はサポートされない。
'*          このメソッドは、 Recordset オブジェクトが CommandTypeEnum の値を
'*          adCmdTableDirect にして開かれた場合にのみ使用できる。
'*
'* @param KeyValues  バリアント型 ( Variant ) の値の配列。
'*                   インデックスは 1 つまたは複数の列から成るため、
'*                   対応する各列と比較する値をこの配列に格納する。
'* @param SeekOption インデックスの各列とそれに対応する KeyValues の比較に使用
'*                   する比較の種類を指定する SeekEnum 値。
'******************************************************************************
Sub SeekRs(Optional KeyValues, _
           Optional SeekOption As SeekEnum = adSeekFirstEQ)
    If IsMissing(KeyValues) Then
        Call mOrigin.Seek(, SeekOption)
    Else
        Call mOrigin.Seek(KeyValues, SeekOption)
    End If
End Sub

'******************************************************************************
'* [概  要] Supportsメソッド
'* [詳  細] Supportsのラッパーメソッド。
'*          指定された Recordset オブジェクトが特定の種類の機能をサポートして
'*          いるかどうかを調べる。
'*
'*          Supports メソッドを使用すると、Recordset オブジェクトでサポートされ
'*          ている機能の種類を調べることができる。
'*          CursorOptions で指定した定数に対応する機能が Recordset オブジェクト
'*          でサポートされていると、Supports メソッドは True を返す。
'*          それ以外の場合は、False を返す。
'*          [!メモ] 指定した機能について Supports メソッドから True が返されて
'*          も、その機能がどのような状況でもそのプロバイダーで使用できるという
'*          保証はない。 Supports メソッドは、一定の条件が満たされていることを
'*          前提にしたうえで、指定された機能をプロバイダーがサポートできるかど
'*          うかを判別した結果を単に返すだけである。
'*          たとえば、カーソルが複数のテーブルの結合に基づいているために更新不
'*          可能な列があったとしても、 Supports は Recordset オブジェクトが更
'*          新をサポートしていると判別する場合がある。
'*
'* @param CursorOptions 1つまたは複数の CursorOptionEnum 値から成る長整数
'*                      型 ( Long ) の式を指定
'* @return プロバイダーが CursorOptions 引数で指定されたすべての機能をサポート
'*         しているかどうかを示すブール型 (Boolean) の値
'******************************************************************************
Function Supports(CursorOptions As CursorOptionEnum) As Boolean
    Supports = mOrigin.Supports(CursorOptions)
End Function

'******************************************************************************
'* [概  要] Updateメソッド
'* [詳  細] Updateのラッパーメソッド。
'*          Recordset オブジェクトのカレント行、または Record オブジェクトの
'*          Fields コレクションに加えた変更を保存する。
'*
'*          Update メソッドを使用すると、 AddNew メソッドを呼び出したか、または
'*          既存レコードのいずれかのフィールドの値を変更した時点以降に Recordset
'*          オブジェクトのカレント レコードに加えた変更がすべて保存される。
'*          Recordset オブジェクトが更新をサポートしている必要がある。
'*          フィールドの値を設定するには､次のいずれかを行う｡
'*            ・Field オブジェクトの Value プロパティに値を代入してから、 Update
'*              メソッドを呼び出す。
'*            ・フィールド名と値を Update 呼び出しで引数として渡す。
'*            ・フィールド名の配列と値の配列を Update 呼び出しで渡す。
'*          フィールドと値の配列を使用する場合は､両方の配列の要素の数を同じにす
'*          る必要がある｡また、フィールド名の順序がフィールドの値の順序と一致し
'*          ている必要がある｡
'*          フィールドと値の数と順序が一致していないと､エラーが発生する｡
'*          Recordset オブジェクトが一括更新をサポートしている場合は、1 つまたは
'*          複数のレコードへの複数の変更をローカルにキャッシュしてから、
'*          UpdateBatch メソッドを呼び出すことができる。
'*          カレント レコードの編集中、または新しいレコードの追加中に UpdateBatch
'*          メソッドを呼び出すと、ADO によって自動的に Update メソッドが呼び出さ
'*          れてカレント レコードへの保留中の変更が保存されてから、変更内容がま
'*          とめてプロバイダーに転送される。
'*          Update メソッドを呼び出すより前に追加中または編集中のレコードから
'*          移動すると、ADO によって自動的に Update が呼び出されて変更内容が
'*          保存される。
'*          カレントレコードに加えた変更を取り消す場合や、新しく追加したレコー
'*          ドを破棄する場合は、 CancelUpdate メソッドを呼び出す必要がある。
'*          カレントレコードは、 Update メソッドを呼び出した後もカレントのまま
'*          である。
'*
'* @param Fields 省略可能。変更する単一のフィールドの名前を表すバリアント型の
'*               値、または複数のフィールドの名前か順序を表すバリアント型の
'*               配列を指定｡
'* @param Values 省略可能。新しいレコードでの単一のフィールドの値を表すバリア
'*               ント型の値、または複数のフィールドの値を表すバリアント型の
'*               配列を指定。
'******************************************************************************
Sub Update(Optional Fields, _
           Optional Values)
    If IsMissing(Fields) Then
        If IsMissing(Values) Then
            Call mOrigin.Update
        Else
            Call mOrigin.Update(, Values)
        End If
    ElseIf IsMissing(Values) Then
        Call mOrigin.Update(Fields)
    Else
        Call mOrigin.Update(Fields, Values)
    End If
End Sub
           
'******************************************************************************
'* [概  要] UpdateBatchメソッド
'* [詳  細] UpdateBatchのラッパーメソッド。
'*          保留中の一括更新をすべてディスクに書き込む｡
'*
'*          Recordset オブジェクトを一括更新モードで変更している場合に
'*          UpdateBatch メソッドを使用すると、Recordset オブジェクトに加えられ
'*          た変更を基になるデータベースに転送できる。
'*          Recordset オブジェクトが一括更新をサポートしている場合は、
'*          UpdateBatch メソッドを呼び出すまでの間、1 つまたは複数のレコード
'*          への複数の変更をローカルにキャッシュしておくことができる。
'*          カレントレコードの編集中、または新しいレコードの追加中に
'*          UpdateBatch メソッドを呼び出すと、ADO によって自動的に Update
'*          メソッドが呼び出されてカレント レコードへの保留中の変更が保存されて
'*          から、変更内容がまとめてプロバイダーに転送される。
'*          一括更新は、キーセットまたは静的カーソルのみで使用すること。
'*          [!メモ] パラメーターの値として adAffectGroup を指定すると、現在の
'*          Recordset 内に表示可能なレコードがなかった場合 (一致するレコードの
'*          ないフィルターがかかっているなど) にはエラーが発生する。
'*          基になるデータとの競合が原因で一部またはすべてのレコードの変更の
'*          転送に失敗した場合 (たとえば、レコードが他のユーザーによって既に
'*          削除されていた場合)、プロバイダーから Errors コレクションに警告が
'*          返され、実行時エラーが発生する。
'*          Filter プロパティ (adFilterAffectedRecords) と Status プロパティ
'*          を使用すると、競合のあるレコードを突き止めることができる。
'*          保留になっているすべての一括更新を取り消すには､CancelBatch メソッ
'*          ドを使用する｡
'*          動的プロパティの Unique Table と Update Resync が設定されていて、
'*          Recordset が複数のテーブルに対して JOIN 操作を実行した結果になって
'*          いる場合は、 UpdateBatch メソッドの実行の後に、 Update Resync プロ
'*          パティの設定に応じて暗黙的に Resync メソッドが実行される。
'*          個々の一括更新がデータ ソースに対して実行される順序は、必ずしも
'*          それらがローカル Recordset に実行された順序と同じになるわけでは
'*          ない。更新の順序は、プロバイダーに依存する。
'*          挿入または更新に対する外部キー制約など、相互に関連のある更新を
'*          コーディングする場合は、このことを考慮すること。
'*
'* @param AffectRecords 省略可能。UpdateBatch メソッドで処理するレコードの数
'*                      を示す AffectEnum 値を指定。
'******************************************************************************
Sub UpdateBatch(Optional AffectRecords As AffectEnum = adAffectAll)
    Call mOrigin.UpdateBatch(AffectRecords)
End Sub


'******************************************************************************
'* [概  要] xCloneメソッド
'* [詳  細] 隠し項目_xCloneのラッパーメソッド。
'*
'******************************************************************************
Function xClone() As ADODBExRecordset
    Dim wrap As ADODBExRecordset: Set wrap = New ADODBExRecordset
    Set wrap.Origin = mOrigin.[_xClone]
    Set xClone = wrap
End Function


'******************************************************************************
'* [概  要] xResyncメソッド
'* [詳  細] 隠し項目_xResyncのラッパーメソッド。
'*
'* @param AffectRecords 省略可能
'*
'******************************************************************************
Sub xResync(Optional AffectRecords As AffectEnum = adAffectAll)
    Call mOrigin.[_xResync](AffectRecords)
End Sub

'******************************************************************************
'* [概  要] xSaveメソッド
'* [詳  細] 隠し項目_xSaveのラッパーメソッド。
'*
'* @param FileName 省略可能
'* @param PersistFormat 省略可能
'*
'******************************************************************************
Sub xSave(Optional FileName As String, Optional PersistFormat As PersistFormatEnum = adPersistADTG)
    If IsMissing(FileName) Then
        Call mOrigin.[_xSave](, PersistFormat)
    Else
        Call mOrigin.[_xSave](FileName, PersistFormat)
    End If
End Sub

    
