VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] StringEx クラス
'* [詳  細] Stringの拡張クラス。
'*
'* [参  考] https://excel-ubara.com/excelvba4/EXCEL271.html
'*          http://dev-clips.com/clip/vba/stringbuilder-for-vba/
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* Enum定義
'******************************************************************************


'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mBuffer As String
Private mCapacity As Long
Private mLength As Long

'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* Value プロパティ
'*-----------------------------------------------------------------------------
Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    Value = Me.ToString
End Property
Public Property Let Value(ByVal arg As String)
    mBuffer = String$(mCapacity, vbNullChar)
    mLength = 0
    Me.Append arg
End Property

'*-----------------------------------------------------------------------------
'* Length プロパティ
'*-----------------------------------------------------------------------------
Public Property Get Length() As Long
    Length = mLength
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    mCapacity = 32768
    mBuffer = String$(mCapacity, vbNullChar)
    mLength = 0
End Sub
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Init メソッド
'* [詳  細] 初期化用メソッド｡
'*
'* @param val 初期文字列
'* @param cap 初期確保領域（バイト）
'* @return StringExオブジェクト
'*
'******************************************************************************
Public Function Init(ByVal val, Optional cap As Long = 32768) As StringEx
    mCapacity = cap
    Me.Value = val
    Set Init = Me
End Function

'******************************************************************************
'* [概  要] Append メソッド
'* [詳  細] 文字列を追加する｡
'*
'* @param val 追加文字列
'* @return StringExオブジェクト
'*
'******************************************************************************
Public Function Append(ByVal val As String) As StringEx
    If mLength + Len(val) > Len(mBuffer) Then
        mBuffer = mBuffer & String$(CLng(Len(mBuffer) * 2) + Len(val), vbNullChar)
    End If
    Mid(mBuffer, mLength + 1) = val
    mLength = mLength + Len(val)
    Set Append = Me
End Function

'******************************************************************************
'* [概  要] AppendLine メソッド
'* [詳  細] 文字列を行追加する｡
'*
'* @param val 追加文字列
'* @param lineSeparator 改行コード
'* @return StringExオブジェクト
'*
'******************************************************************************
Public Function AppendLine(ByVal val As String, Optional ByVal LineSeparator As String = vbNewLine) As StringEx
    If mLength > 0 Then
        If Not StringUtils.Contains(Mid$(mBuffer, mLength - 1, 2), LineSeparator) Then
            Append LineSeparator
        End If
    End If
    Set AppendLine = Append(val).Append(LineSeparator)
End Function

'******************************************************************************
'* [概  要] CodePointCount
'* [詳  細] サロゲートペアを考慮した文字数を取得｡
'*
'* @return Long 文字数
'*
'******************************************************************************
Public Function CodePointCount() As Long
     CodePointCount = StringUtils.CheckLength(Me.ToString)
End Function


'******************************************************************************
'* [概  要] ToString メソッド
'* [詳  細] 確保領域内に設定されている文字列を返却する｡
'*
'* @return 文字列
'*
'******************************************************************************
Public Function ToString() As String
    ToString = Left$(mBuffer, mLength)
End Function


'******************************************************************************
'* [概  要] Insert メソッド
'* [詳  細] 文字列を指定位置に挿入する｡
'*
'* @param val 追加文字列
'* @param position 位置
'* @return StringExオブジェクト
'*
'******************************************************************************
Public Function Insert(ByRef val As String, ByVal Position As Long) As StringEx

    If mLength < Position Then
        Set Insert = Me.Append(val)
        Exit Function
    ElseIf Position < 1 Then
        Position = 1
    End If

    Dim tmpLen As Long: tmpLen = mLength
    mLength = mLength + Len(val)

    If mLength > mCapacity Then
        Do While mCapacity < mLength
            mCapacity = mCapacity * 2
        Loop
    End If
    Mid(mBuffer, Position + Len(val)) = Mid$(mBuffer, Position, tmpLen)
    Mid(mBuffer, Position) = val

    Set Insert = Me
End Function

'******************************************************************************
'* [概  要] Trim
'* [詳  細] 左右の半角スペースをトリム｡
'*
'* @return StringEx トリム済みStringExオブジェクト
'*
'******************************************************************************
Public Function Trim() As StringEx
     Set Trim = Me.Init(VBA.Trim$(Me.ToString))
End Function

'******************************************************************************
'* [概  要] TrimZ
'* [詳  細] 全角も含む左右のスペースをトリム｡
'*
'* @return StringEx トリム済みStringExオブジェクト
'*
'******************************************************************************
Public Function TrimZ() As StringEx
     Set TrimZ = Me.Init(StringUtils.TrimZ(Me.ToString))
End Function

'******************************************************************************
'* [概  要] StartsWith
'* [詳  細] 指定した文字列で始まるかチェック｡
'*
'* @param startStr 開始文字列
'* @return Boolean チェック結果（該当:True）
'*
'******************************************************************************
Public Function StartsWith(ByVal startStr As String) As Boolean
    StartsWith = StringUtils.StartsWith(Me.ToString, startStr)
End Function

'******************************************************************************
'* [概  要] EndsWith
'* [詳  細] 指定した文字列で終るかチェック｡
'*
'* @param endStr 終了文字列
'* @return Boolean チェック結果（該当:True）
'*
'******************************************************************************
Public Function EndsWith(ByVal endStr As String) As Boolean
    EndsWith = StringUtils.EndsWith(Me.ToString, endStr)
End Function

'******************************************************************************
'* [概  要] Escape
'* [詳  細] 文字列のエスケープを行う｡
'*
'* @return StringEx エスケープ後文字列
'*
'******************************************************************************
Public Function Escape() As StringEx
    Set Escape = Me.Init(StringUtils.Escape(Me.ToString))
End Function

'******************************************************************************
'* [概  要] Unescape
'* [詳  細] 文字列のアンエスケープを行う｡
'*
'* @return String アンエスケープ後文字列
'*
'******************************************************************************
Public Function Unescape() As StringEx
    Set Unescape = Me.Init(StringUtils.Unescape(Me.ToString))
End Function

'******************************************************************************
'* [概  要] PlaceHolder
'* [詳  細] 文字列のプレースホルダー{x}に値を埋め込む｡
'*
'* @param p 埋込文字列の配列
'* @return StringEx 埋込後文字列
'*
'******************************************************************************
Public Function PlaceHolder(ParamArray p() As Variant) As StringEx
    Dim tmp: tmp = p
    Set PlaceHolder = Me.Init(StringUtils.PlaceHolder(Me.ToString, tmp))
End Function

'******************************************************************************
'* [概  要] FromUTF8
'* [詳  細] UTF8バイト配列からStringEx型に変換｡
'*
'* @param bytBuf() Byte型配列(UTF8)
'* @return StringEx 文字列オブジェクト
'*
'******************************************************************************
Function FromUTF8(ByRef bytBuf() As Byte) As StringEx
    Set FromUTF8 = Me.Init(StringUtils.FromUTF8(bytBuf))
End Function

'******************************************************************************
'* [概  要] ToUTF8
'* [詳  細] Byte型配列(UTF8)に変換｡
'*
'* @return Byte() Byte型配列(UTF8)
'*
'******************************************************************************
Function ToUTF8() As Byte()
    ToUTF8 = StringUtils.ToUTF8(Me.ToString)
End Function

'******************************************************************************
'* [概  要] ToHiragana
'* [詳  細] カタカナからひらがなに変換｡
'*
'* @return StringEx ひらがな文字列に変換したStringExオブジェクト
'*
'******************************************************************************
Public Function ToHiragana(Optional ByVal flag As Boolean = False) As StringEx
    Set ToHiragana = Me.Init(StringUtils.ToHiragana(Me.ToString, flag))
End Function

'******************************************************************************
'* [概  要] ToKatakana
'* [詳  細] ひらがなからカタカナに変換｡
'*
'* @return StringEx カタカナ文字列に変換したStringExオブジェクト
'*
'******************************************************************************
Public Function ToKatakana(Optional ByVal flag As Boolean = False) As StringEx
    Set ToKatakana = Me.Init(StringUtils.ToKatakana(Me.ToString, flag))
End Function

'******************************************************************************
'* [概  要] ToKatakana
'* [詳  細] ひらがなからカタカナに変換｡
'*
'* @return StringEx カタカナ文字列に変換したStringExオブジェクト
'*
'******************************************************************************
Public Function Replace(ByVal exp As String, ByVal rep As String) As StringEx
    Set Replace = Me.Init(VBA.Replace(Me.ToString, exp, rep))
End Function

'******************************************************************************
'* [概  要] Matches
'* [詳  細] 正規表現にマッチする文字列かチェック｡
'*
'* @return Boolean チェック結果（該当:True）
'*
'******************************************************************************
Public Function matches(ByVal regex As String) As Boolean
    matches = ValidateUtils.ValidatePattern(Me.ToString, regex)
End Function

'******************************************************************************
'* [概  要] GetBytes
'* [詳  細] 文字列をバイト配列に変換｡
'*
'* @param encode エンコードのキャラセット
'* @return Byte() バイト配列
'*
'******************************************************************************
Public Function GetBytes(Optional encode As String) As Byte()
    If encode = "" Then
       GetBytes = Me.ToString
       Exit Function
    End If
    
    With IO.CreateADODBStream(adTypeText, encode)
        .mode = adModeReadWrite
        .Open
        .WriteText Me.ToString
        .Position = 0
        .Type = adTypeBinary
        GetBytes = .Read
        .Close
    End With
End Function

'******************************************************************************
'* [概  要] AppendIfMissing
'* [詳  細] 末尾に接尾辞が存在しない場合は、Stringの末尾に接尾辞を追加｡
'*
'* @param suffix 文字列の最後に追加する接尾辞
'* @param suffixes 有効なターミネーターである追加のサフィックス接尾辞
'* @return StringEx 接尾辞が追加された場合は新しい文字列、それ以外の場合は同じ文字列
'*
'******************************************************************************
Public Function AppendIfMissing(suffix As String, ParamArray suffixes()) As StringEx
    If Not ValidateUtils.IsParamArrayMissing(suffixes) Then
        Dim tmp: tmp = suffixes ' ParamArrayを別関数に渡すための対応
        Set AppendIfMissing = Me.Init(StringUtils.AppendIfMissing(Me.ToString, suffix, tmp))
    Else
        Set AppendIfMissing = Me.Init(StringUtils.AppendIfMissing(Me.ToString, suffix))
    End If
End Function

'******************************************************************************
'* [概  要] AppendIfMissingIgnoreCase
'* [詳  細] 末尾に接尾辞が存在しない場合は、Stringの末尾に接尾辞を追加｡
'*          接尾辞の存在チェックは大文字小文字を考慮しない。
'*
'* @param suffix 文字列の最後に追加する接尾辞
'* @param suffixes 有効なターミネーターである追加のサフィックス接尾辞
'* @return StringEx 接尾辞が追加された場合は新しい文字列、それ以外の場合は同じ文字列
'*
'******************************************************************************
Public Function AppendIfMissingIgnoreCase(str As String, suffix As String, ParamArray suffixes()) As StringEx
    If Not ValidateUtils.IsParamArrayMissing(suffixes) Then
        Dim tmp: tmp = suffixes ' ParamArrayを別関数に渡すための対応
        Set AppendIfMissingIgnoreCase = Me.Init(StringUtils.AppendIfMissingIgnoreCase(Me.ToString, suffix, tmp))
    Else
        Set AppendIfMissingIgnoreCase = Me.Init(StringUtils.AppendIfMissingIgnoreCase(Me.ToString, suffix))
    End If
End Function

'******************************************************************************
'* [概  要] Capitalize
'* [詳  細] 文字列の最初の文字を大文字に変換する｡
'*
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function Capitalize() As StringEx
    Set Capitalize = Me.Init(StringUtils.Capitalize(Me.ToString))
End Function

'******************************************************************************
'* [概  要] Center
'* [詳  細] 文字列をパディングして中央寄せにする｡
'*          指定されたパディング文字があれば、文字列に埋め込む値として使用する。
'*          デフォルトのパディング文字は半角スペース。
'*          サイズが文字列の長さより小さい場合、元の文字列が返される。
'*          空文字列は空文字列を返す。負のサイズはゼロとして扱われる。
'*
'* @param size 新しい文字列のサイズ
'* @param padStr パディング文字
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function Center(size As Long, Optional padStr As String = " ") As StringEx
    Set Center = Me.Init(StringUtils.Center(Me.ToString, size, padStr))
End Function

'******************************************************************************
'* [概  要] Chomp
'* [詳  細] 文字列の末尾にある場合は改行を1つ削除する。
'*          それ以外の場合はそのままにする｡
'*          改行は「\n」、「\r」、または「\r\n」｡
'*
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function Chomp() As StringEx
    Set Chomp = Me.Init(StringUtils.Chomp(Me.ToString))
End Function

'******************************************************************************
'* [概  要] Chop
'* [詳  細] 文字列から最後の文字を削除｡
'*          文字列が\r\nで終わっている場合は、両方を削除する。
'*
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function Chop() As StringEx
    Set Chop = Me.Init(StringUtils.Chop(Me.ToString))
End Function

'******************************************************************************
'* [概  要] Compare
'* [詳  細] 指定した文字列と比較する｡
'*
'* @param str 比較する文字列
'* @return Long チェック結果
'*         int = 0     ：Me.ToString = str
'*         int < 0(-1) ：Me.ToString  < str
'*         int > 0( 1) ：Me.ToString  > str
'*
'******************************************************************************
Public Function Compare(str As String) As Long
    Compare = StringUtils.Compare(Me.ToString, str)
End Function

'******************************************************************************
'* [概  要] CompareIgnoreCase
'* [詳  細] 指定した文字列と比較する｡大文字小文字を考慮しない。
'*
'* @param str 比較する文字列
'* @return Long チェック結果
'*         int = 0     ：Me.ToString = str
'*         int < 0(-1) ：Me.ToString  < str
'*         int > 0( 1) ：Me.ToString  > str
'*
'******************************************************************************
Public Function CompareIgnoreCase(str As String) As Long
    CompareIgnoreCase = StringUtils.CompareIgnoreCase(Me.ToString, str)
End Function

'******************************************************************************
'* [概  要] Contains
'* [詳  細] 指定した文字列が含まれているかどうかをチェック｡
'*          空文字の文字列はFalseを返す｡
'*          空文字の検索文字はFalseを返す｡
'*
'* @param searchStr 検索文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function Contains(searchStr As String) As Boolean
    Contains = StringUtils.Contains(Me.ToString, searchStr)
End Function

'******************************************************************************
'* [概  要] ContainsAll
'* [詳  細] 指定した文字列のすべてが含まれているかチェック｡
'*          空文字の文字列はFalseを返す｡
'*          空文字または長さがゼロの検索配列はFalseを返す｡
'*
'* @param searchStr 検索文字配列（複数指定）
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function ContainsAll(ParamArray searchStr()) As Boolean
    If ValidateUtils.IsParamArrayMissing(searchStr) Then
        ContainsAll = False
        Exit Function
    End If
    Dim tmp: tmp = searchStr ' ParamArrayを別関数に渡すための対応
    ContainsAll = StringUtils.ContainsAll(Me.ToString, searchStr)
End Function

'******************************************************************************
'* [概  要] ContainsAny
'* [詳  細] 指定した文字列のどれかが含まれているかどうかをチェック｡
'*          空文字の文字列はFalseを返す｡
'*          空文字または長さがゼロの検索配列はFalseを返す｡
'*
'* @param searchStr 検索文字配列（複数指定）
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function ContainsAny(ParamArray searchStr()) As Boolean
    If ValidateUtils.IsParamArrayMissing(searchStr) Then
        ContainsAny = False
        Exit Function
    End If
    Dim tmp: tmp = searchStr ' ParamArrayを別関数に渡すための対応
    ContainsAny = StringUtils.ContainsAny(Me.ToString, searchStr)
End Function

'******************************************************************************
'* [概  要] ContainsIgnoreCase
'* [詳  細] 指定した文字列が含まれているかどうかをチェック｡
'*          大文字小文字を考慮しない。
'*          空文字の文字列はFalseを返す｡
'*          空文字の検索文字はFalseを返す｡
'*
'* @param searchStr 検索文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function ContainsIgnoreCase(searchStr As String) As Boolean
    ContainsIgnoreCase = StringUtils.ContainsIgnoreCase(Me.ToString, searchStr)
End Function

'******************************************************************************
'* [概  要] ContainsNone
'* [詳  細] 指定した文字列が含まれていないことをチェック｡
'*          空文字の文字列はTrueを返す｡
'*          空文字の検索文字はTrueを返す｡
'*
'* @param searchStr 検索文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function ContainsNone(searchStr As String) As Boolean
    ContainsNone = StringUtils.ContainsNone(Me.ToString, searchStr)
End Function

'******************************************************************************
'* [概  要] ContainsOnly
'* [詳  細] 特定の文字のみが含まれているかどうかをチェック｡
'*          空の文字列（length（）= 0）は常にTrueを返す。
'*
'* @param searchStr 検索文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function ContainsOnly(str As String, searchStr As String) As Boolean
    ContainsOnly = StringUtils.ContainsOnly(Me.ToString, searchStr)
End Function

'******************************************************************************
'* [概  要] ContainsWhitespace
'* [詳  細] 空白文字が含まれているかどうかをチェック｡
'*
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function ContainsWhitespace() As Boolean
    ContainsWhitespace = StringUtils.ContainsWhitespace(Me.ToString)
End Function

'******************************************************************************
'* [概  要] CountMatches
'* [詳  細] 文字が指定された文字列に出現する回数をカウントする｡
'*
'* @param searchStr 検索文字列
'* @return Long 出現回数
'*
'******************************************************************************
Public Function CountMatches(str As String, searchStr As String) As Long
    CountMatches = StringUtils.CountMatches(Me.ToString, searchStr)
End Function

'******************************************************************************
'* [概  要] DefaultIfBlank
'* [詳  細] 文字列がスペースか空（""）の場合は、指定したデフォルト値を返す｡
'*
'* @param defaultStr デフォルト文字列
'* @return StringEx チェック済文字列
'*
'******************************************************************************
Public Function DefaultIfBlank(defaultStr As String) As StringEx
    Set DefaultIfBlank = Me.Init(StringUtils.DefaultIfBlank(Me.ToString, defaultStr))
End Function

'******************************************************************************
'* [概  要] DefaultIfEmpty
'* [詳  細] 文字列が空（""）の場合は、指定したデフォルト値を返す｡
'*
'* @param defaultStr デフォルト文字列
'* @return String チェック済文字列
'*
'******************************************************************************
Public Function DefaultIfEmpty(defaultStr As String) As StringEx
    Set DefaultIfEmpty = Me.Init(StringUtils.DefaultIfEmpty(Me.ToString, defaultStr))
End Function

'******************************************************************************
'* [概  要] DeleteWhitespace
'* [詳  細] 文字列からすべての半角スペースを削除｡
'*
'* @return String 削除後文字列
'*
'******************************************************************************
Public Function DeleteWhitespace() As StringEx
    Set DeleteWhitespace = Me.Init(StringUtils.DeleteWhitespace(Me.ToString))
End Function

'******************************************************************************
'* [概  要] Difference
'* [詳  細] 2つの文字列を比較し、それらが異なる部分を返す。
'*          より正確には、最初の文字列とは異なる場所から、2番目の文字列の残りを返す。
'*          「abc」と「ab」の違いは空の文字列であり、「c」ではない。
'*
'* @param str 比較する文字列
'* @return StringEx 異なる部分の文字列。等しい場合は空の文字列を返す
'*
'******************************************************************************
Public Function Difference(str As String) As StringEx
    Set Difference = Me.Init(StringUtils.Difference(Me.ToString, str))
End Function

'******************************************************************************
'* [概  要] EndsWithAny
'* [詳  細] 指定した文字列のセットのいづれかで終るかチェック｡
'*
'* @param endStr 終了文字列（複数指定）
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function EndsWithAny(ParamArray endStr() As Variant) As Boolean
    If Not ValidateUtils.IsParamArrayMissing(endStr) Then
        Dim tmp: tmp = endStr ' ParamArrayを別関数に渡すための対応
        EndsWithAny = StringUtils.EndsWithAny(Me.ToString, tmp)
    Else
        EndsWithAny = StringUtils.EndsWithAny(Me.ToString)
    End If
End Function

'******************************************************************************
'* [概  要] EndsWithIgnoreCase
'* [詳  細] 指定した文字列で終るかチェック｡
'*          大文字小文字を考慮しない。
'*
'* @param endStr 終了文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function EndsWithIgnoreCase(endStr As String) As Boolean
     EndsWithIgnoreCase = StringUtils.EndsWithIgnoreCase(Me.ToString, endStr)
End Function

'******************************************************************************
'* [概  要] Equals
'* [詳  細] 指定した文字列と等しいかチェック｡
'*
'* @param str 文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function Equals(str As String) As Boolean
     Equals = StringUtils.Equals(Me.ToString, str)
End Function

'******************************************************************************
'* [概  要] EqualsIgnoreCase
'* [詳  細] 指定した文字列と等しいかチェック｡
'*          大文字小文字を考慮しない。
'*
'* @param str 文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function EqualsIgnoreCase(str As String) As Boolean
     EqualsIgnoreCase = StringUtils.EqualsIgnoreCase(Me.ToString, str)
End Function

'******************************************************************************
'* [概  要] EqualsAny
'* [詳  細] 文字列が指定した文字列のいづれかに等しいかチェック｡
'*
'* @param searchStrings 比較文字列（複数指定）
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function EqualsAny(ParamArray searchStrings() As Variant) As Boolean
    If Not ValidateUtils.IsParamArrayMissing(searchStrings) Then
        Dim tmp: tmp = searchStrings ' ParamArrayを別関数に渡すための対応
        EqualsAny = StringUtils.EqualsAny(Me.ToString, searchStrings)
    Else
        EqualsAny = StringUtils.EqualsAny(Me.ToString)
    End If
End Function

'******************************************************************************
'* [概  要] EqualsAnyIgnoreCase
'* [詳  細] 文字列が指定した文字列のいづれかに等しいかチェック｡
'*          大文字小文字を考慮しない。
'*
'* @param searchStrings 比較文字列（複数指定）
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function EqualsAnyIgnoreCase(ParamArray searchStrings() As Variant) As Boolean
    If Not ValidateUtils.IsParamArrayMissing(searchStrings) Then
        Dim tmp: tmp = searchStrings ' ParamArrayを別関数に渡すための対応
        EqualsAnyIgnoreCase = StringUtils.EqualsAnyIgnoreCase(Me.ToString, searchStrings)
    Else
        EqualsAnyIgnoreCase = StringUtils.EqualsAnyIgnoreCase(Me.ToString)
    End If
End Function

'******************************************************************************
'* [概  要] GetDigits
'* [詳  細] 文字列strに数字が含まれているかどうかをチェックし、
'*          含まれている場合はstr内のすべての数字を連結して文字列として返す｡
'*          数字が見つからない場合は、空の（""）文字列が返される。
'*
'* @param str 文字列
'* @return StringEx 数字文字列
'*
'******************************************************************************
Public Function GetDigits() As StringEx
    Set GetDigits = Me.Init(StringUtils.GetDigits(Me.ToString))
End Function

'******************************************************************************
'* [概  要] GetIfBlank
'* [詳  細] 文字列がスペースか空（""）の場合は、指定した関数の結果を返す｡
'*
'* @param defaultSupplier デフォルト文字列を返すサプライヤー（Funcオブジェクト）
'* @return StringEx チェック済文字列
'*
'******************************************************************************
Public Function GetIfBlank(defaultSupplier As Func) As StringEx
    Set GetIfBlank = Me.Init(StringUtils.GetIfBlank(Me.ToString, defaultSupplier))
End Function

'******************************************************************************
'* [概  要] GetIfEmpty
'* [詳  細] 渡された文字列が空（""）の場合は、指定した関数の結果を返す｡
'*
'* @param defaultSupplier デフォルト文字列を返すサプライヤー（Funcオブジェクト）
'* @return StringEx チェック済文字列
'*
'******************************************************************************
Public Function GetIfEmpty(defaultSupplier As Func) As StringEx
    Set GetIfEmpty = Me.Init(StringUtils.GetIfEmpty(Me.ToString, defaultSupplier))
End Function

'******************************************************************************
'* [概  要] IndexOf
'* [詳  細] 文字列内の最初のインデックス（文字位置）を検索する｡
'*
'* @param searchStr 検索文字列
'* @param startPos 開始位置。負の値はゼロとして扱われる。
'* @return Long 文字列の最初のインデックス。一致がない場合、空文字列入力の場合は-1
'*
'******************************************************************************
Public Function IndexOf(searchStr As String, Optional ByVal startPos As Long = 0) As Long
     IndexOf = StringUtils.IndexOf(Me.ToString, searchStr, startPos)
End Function

'******************************************************************************
'* [概  要] IndexOfAny
'* [詳  細] 指定された文字セット内の任意の文字の最初のインデックス（文字位置）を検索する｡
'*
'* @param searchStrs 検索文字列（複数指定）
'* @return Long 文字列の最初のインデックス。一致がない場合、空配列入力の場合は-1
'*
'******************************************************************************
Public Function IndexOfAny(ParamArray searchStrs() As Variant) As Long
    If Not ValidateUtils.IsParamArrayMissing(searchStrs) Then
        Dim tmp: tmp = searchStrs ' ParamArrayを別関数に渡すための対応
        IndexOfAny = StringUtils.IndexOfAny(Me.ToString, searchStrs)
    Else
        IndexOfAny = StringUtils.IndexOfAny(Me.ToString)
    End If
End Function

'******************************************************************************
'* [概  要] IndexOfAnyBut
'* [詳  細] 指定された文字セットにない任意の文字の最初のインデックス（文字位置）
'*          を検索する｡
'*
'* @param searchStrs 検索文字（１文字。複数指定）
'* @return Long 文字列の最初のインデックス。一致がない場合、空文字列入力の場合は-1
'*
'******************************************************************************
Public Function IndexOfAnyBut(ParamArray searchStrs() As Variant) As Long
    If Not ValidateUtils.IsParamArrayMissing(searchStrs) Then
        Dim tmp: tmp = searchStrs ' ParamArrayを別関数に渡すための対応
        IndexOfAnyBut = StringUtils.IndexOfAnyBut(Me.ToString, searchStrs)
    Else
        IndexOfAnyBut = StringUtils.IndexOfAnyBut(Me.ToString)
    End If
End Function

'******************************************************************************
'* [概  要] IndexOfDifference
'* [詳  細] 指定文字列と比較し、文字が異なり始めるインデックスを返す｡
'*
'* @param str 文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IndexOfDifference(str As String) As Long
    IndexOfDifference = StringUtils.IndexOfDifference(Me.ToString, str)
End Function

'******************************************************************************
'* [概  要] IndexOfIgnoreCase
'* [詳  細] 文字列内の最初のインデックス（文字位置）を検索する｡
'*          大文字小文字を考慮しない。
'*
'* @param searchStr 検索文字列
'* @param startPos 開始位置。負の値はゼロとして扱われる。
'* @return Long 文字列の最初のインデックス。一致がない場合、空文字列入力の場合は-1
'*
'******************************************************************************
Public Function IndexOfIgnoreCase(searchStr As String, Optional startPos As Long = 0) As Long
    IndexOfIgnoreCase = StringUtils.IndexOfIgnoreCase(Me.ToString, searchStr, startPos)
End Function

'******************************************************************************
'* [概  要] IsLowerCase
'* [詳  細] 小文字のみかどうかをチェック｡
'*
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsLowerCase() As Boolean
    IsLowerCase = StringUtils.IsLowerCase(Me.ToString)
End Function

'******************************************************************************
'* [概  要] IsUpperCase
'* [詳  細] 大文字のみかどうかをチェック｡
'*
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsUpperCase() As Boolean
    IsUpperCase = StringUtils.IsUpperCase(Me.ToString)
End Function

'******************************************************************************
'* [概  要] IsAlpha
'* [詳  細] 文字列が英字かどうかをチェック｡
'*
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsAlpha() As Boolean
    IsAlpha = StringUtils.IsAlpha(Me.ToString)
End Function


'******************************************************************************
'* [概  要] IsAlphanumeric
'* [詳  細] 文字列が英数字かどうかをチェック｡
'*
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsAlphanumeric() As Boolean
    IsAlphanumeric = StringUtils.IsAlphanumeric(Me.ToString)
End Function


'******************************************************************************
'* [概  要] IsAlphanumericSpace
'* [詳  細] 文字列が英数字とスペースかどうかをチェック｡
'*
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsAlphanumericSpace() As Boolean
    IsAlphanumericSpace = StringUtils.IsAlphanumericSpace(Me.ToString)
End Function

'******************************************************************************
'* [概  要] IsAlphaSpace
'* [詳  細] 文字列が英字とスペースかどうかをチェック｡
'*
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsAlphaSpace() As Boolean
    IsAlphaSpace = StringUtils.IsAlphaSpace(Me.ToString)
End Function

'******************************************************************************
'* [概  要] IsAsciiPrintable
'* [詳  細] 文字列にASCII印刷可能文字のみが含まれているかどうかをチェック｡
'*
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsAsciiPrintable() As Boolean
    IsAsciiPrintable = StringUtils.IsAsciiPrintable(Me.ToString)
End Function

'******************************************************************************
'* [概  要] IsBlank
'* [詳  細] 文字列が空（""）、スペースのみかどうかをチェック｡
'*
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsBlank() As Boolean
    IsBlank = StringUtils.IsBlank(Me.ToString)
End Function

'******************************************************************************
'* [概  要] IsEmpty
'* [詳  細] 文字列が空（""）かどうかをチェック｡
'*
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsEmpty() As Boolean
    IsEmpty = StringUtils.IsEmpty(Me.ToString)
End Function

'******************************************************************************
'* [概  要] IsMixedCase
'* [詳  細] 文字列が大文字と小文字の両方が混在しているかどうかをチェック｡
'*
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsMixedCase() As Boolean
    IsMixedCase = StringUtils.IsMixedCase(Me.ToString)
End Function
'******************************************************************************
'* [概  要] IsNotBlank
'* [詳  細] 文字列が空（""）、スペースのみでないかどうかをチェック｡
'*
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsNotBlank() As Boolean
    IsNotBlank = StringUtils.IsNotBlank(Me.ToString)
End Function

'******************************************************************************
'* [概  要] IsNotEmpty
'* [詳  細] 文字列が空（""）のみでないかどうかをチェック｡
'*
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsNotEmpty() As Boolean
    IsNotEmpty = StringUtils.IsNotEmpty(Me.ToString)
End Function

'******************************************************************************
'* [概  要] IsNumeric
'* [詳  細] 文字列が数字かどうかをチェック｡
'*          小数は、Falseを返す。
'*
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsNumeric() As Boolean
    IsNumeric = StringUtils.IsNumeric(Me.ToString)
End Function

'******************************************************************************
'* [概  要] IsNumericSpace
'* [詳  細] 文字列が数字またはスペースかどうかをチェック｡
'*
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsNumericSpace() As Boolean
    IsNumericSpace = StringUtils.IsNumericSpace(Me.ToString)
End Function

'******************************************************************************
'* [概  要] LastIndexOf
'* [詳  細] 文字列内の最後のインデックスを検索｡
'*
'* @param searchStr 検索文字列
'* @param startPos 開始位置。負の値はゼロとして扱われる。
'* @return Long 検索文字列の最後のインデックス。一致しない場合、-1
'*
'******************************************************************************
Public Function LastIndexOf(searchStr As String, Optional ByVal startPos) As Long
    LastIndexOf = StringUtils.LastIndexOf(Me.ToString, searchStr, startPos)
End Function

'******************************************************************************
'* [概  要] LastIndexOfIgnoreCase
'* [詳  細] 文字列内の最後のインデックスを検索｡
'*          接尾辞の存在チェックは大文字小文字を考慮しない。
'*
'* @param searchStr 検索文字列
'* @param startPos 開始位置。負の値はゼロとして扱われる。
'* @return Long 検索文字列の最後のインデックス。一致しない場合、-1
'*
'******************************************************************************
Public Function LastIndexOfIgnoreCase(searchStr As String, Optional ByVal startPos) As Long
    LastIndexOfIgnoreCase = StringUtils.LastIndexOfIgnoreCase(Me.ToString, searchStr, startPos)
End Function


'******************************************************************************
'* [概  要] LeftPad
'* [詳  細] 文字列を指定したパディング文字列で左詰めする｡
'*          デフォルトのパディング文字列は半角スペース。
'*
'* @param size パディングするサイズ
'* @param padStr パディング文字列
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function LeftPad(size As Long, Optional padStr As String = " ") As StringEx
    Set LeftPad = Me.Init(StringUtils.LeftPad(Me.ToString, size, padStr))
End Function


'******************************************************************************
'* [概  要] LowerCase
'* [詳  細] 文字列を小文字に変換する｡
'*
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function LowerCase() As StringEx
    Set LowerCase = Me.Init(StringUtils.LowerCase(Me.ToString))
End Function


'******************************************************************************
'* [概  要] Narrow
'* [詳  細] 全角文字を半角文字に変換する｡
'*
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function Narrow() As StringEx
    Set Narrow = Me.Init(StringUtils.Narrow(Me.ToString))
End Function

'******************************************************************************
'* [概  要] NarrowNumber
'* [詳  細] 全角数字を半角数字に変換する｡
'*
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function NarrowNumber() As StringEx
    Set NarrowNumber = Me.Init(StringUtils.NarrowNumber(Me.ToString))
End Function

'******************************************************************************
'* [概  要] PrependIfMissing
'* [詳  細] 文字列がどの接頭辞（プレフィックス）でも始まっていない場合は、
'*          文字列の先頭に接頭辞を付加する｡
'*
'* @param prefix  文字列の先頭に追加する接頭辞
'* @param prefixes() 追加のプレフィックス
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function PrependIfMissing(prefix As String, ParamArray prefixes() As Variant) As StringEx
    If Not ValidateUtils.IsParamArrayMissing(prefixes) Then
        Dim tmp: tmp = prefixes
        Set PrependIfMissing = Me.Init(StringUtils.PrependIfMissing(Me.ToString, prefix, tmp))
    Else
        Set PrependIfMissing = Me.Init(StringUtils.PrependIfMissing(Me.ToString, prefix))
    End If
End Function


'******************************************************************************
'* [概  要] PrependIfMissingIgnoreCase
'* [詳  細] 文字列がどの接頭辞（プレフィックス）でも始まっていない場合は、
'*          文字列の先頭に接頭辞を付加する｡
'*          大文字小文字を考慮しない。
'*
'* @param prefix  文字列の先頭に追加する接頭辞
'* @param prefixes() 追加のプレフィックス
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function PrependIfMissingIgnoreCase(prefix As String, ParamArray prefixes() As Variant) As StringEx
    If Not ValidateUtils.IsParamArrayMissing(prefixes) Then
         Dim tmp: tmp = prefixes
        Set PrependIfMissingIgnoreCase = Me.Init(StringUtils.PrependIfMissingIgnoreCase(Me.ToString, prefix, tmp))
    Else
        Set PrependIfMissingIgnoreCase = Me.Init(StringUtils.PrependIfMissingIgnoreCase(Me.ToString, prefix))
    End If
End Function


'******************************************************************************
'* [概  要] Proper
'* [詳  細] 先頭文字のみを大文字にする｡
'*
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function Proper() As StringEx
    Set Proper = Me.Init(StringUtils.Proper(Me.ToString))
End Function

'******************************************************************************
'* [概  要] Remove
'* [詳  細] ソース文字列内から指定した文字に一致する全ての文字を削除｡
'*
'* @param removeStr 削除する文字列
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function Remove(removeStr As String) As StringEx
    Set Remove = Me.Init(StringUtils.Remove(Me.ToString, removeStr))
End Function

'******************************************************************************
'* [概  要] RemoveEnd
'* [詳  細] ソース文字列の末尾にある場合にのみ部分文字列を削除し、
'*          それ以外の場合はソース文字列を返す。
'*
'* @param removeStr 削除する文字列
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function RemoveEnd(removeStr As String) As StringEx
    Set RemoveEnd = Me.Init(StringUtils.RemoveEnd(Me.ToString, removeStr))
End Function


'******************************************************************************
'* [概  要] RemoveEndIgnoreCase
'* [詳  細] ソース文字列の末尾にある場合にのみ部分文字列を削除し、
'*          それ以外の場合はソース文字列を返す。
'*          大文字小文字を考慮しない。
'*
'* @param removeStr 削除する文字列
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function RemoveEndIgnoreCase(removeStr As String) As StringEx
    Set RemoveEndIgnoreCase = Me.Init(StringUtils.RemoveEndIgnoreCase(Me.ToString, removeStr))
End Function

'******************************************************************************
'* [概  要] RemoveIgnoreCase
'* [詳  細] ソース文字列内から指定した文字に一致する全ての文字を削除｡
'*          大文字小文字を考慮しない。
'*
'* @param removeStr 削除する文字列
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function RemoveIgnoreCase(removeStr As String) As StringEx
    Set RemoveIgnoreCase = Me.Init(StringUtils.RemoveIgnoreCase(Me.ToString, removeStr))
End Function

'******************************************************************************
'* [概  要] RemoveStart
'* [詳  細] ソース文字列の先頭にある場合にのみ部分文字列を削除し、
'*          それ以外の場合はソース文字列を返す｡
'*
'* @param removeStr 削除する文字列
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function RemoveStart(removeStr As String) As StringEx
    Set RemoveStart = Me.Init(StringUtils.RemoveStart(Me.ToString, removeStr))
End Function

'******************************************************************************
'* [概  要] RemoveStartIgnoreCase
'* [詳  細] ソース文字列の先頭にある場合にのみ部分文字列を削除し、
'*          それ以外の場合はソース文字列を返す｡
'*          大文字小文字を考慮しない。
'*
'* @param removeStr 削除する文字列
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function RemoveStartIgnoreCase(removeStr As String) As StringEx
    Set RemoveStartIgnoreCase = Me.Init(StringUtils.RemoveStartIgnoreCase(Me.ToString, removeStr))
End Function


'******************************************************************************
'* [概  要] Repeat
'* [詳  細] 文字列を繰返し回数分繰返し、新しい文字列を形成する。
'*          挿入文字列を指定している場合、繰返し毎に毎回挿入される。
'*
'* @param repeatCount 繰返し回数
'* @param separator 挿入文字列
'* @return StringEx 元の文字列が繰り返された新しい文字列
'*
'******************************************************************************
Public Function Repeat(repeatCount As Long, Optional separator As String) As StringEx
    Set Repeat = Me.Init(StringUtils.Repeat(Me.ToString, repeatCount, separator))
End Function

'******************************************************************************
'* [概  要] ReplaceEach
'* [詳  細] 検索文字列リストで指定した文字列に合致するすべての文字列を
'*          置換文字列リストの文字列で置き換える｡
'*          配列の長さが同じでない場合は置換しない。
'*            ※JavaApacheCommonsLangのStringUtilsとは異なり単純に繰返し
'*              置換を行う。
'*
'* @param searchList 検索文字列リスト（Variant配列）
'* @param replacementList 置換文字列リスト（Variant配列）
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function ReplaceEach(searchList As Variant, replacementList As Variant) As StringEx
    Set ReplaceEach = Me.Init(StringUtils.ReplaceEach(Me.ToString, searchList, replacementList))
End Function

'******************************************************************************
'* [概  要] ReplaceIgnoreCase
'* [詳  細] 指定した検索文字列に合致するすべての文字列を
'*          置換文字列で置き換える｡
'*          大文字小文字を考慮しない。
'*
'* @param searchString 検索文字列
'* @param replacement 置換文字列
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function ReplaceIgnoreCase(searchString As String, replacement As String) As StringEx
    Set ReplaceIgnoreCase = Me.Init(StringUtils.ReplaceIgnoreCase(Me.ToString, searchString, replacement))
End Function

'******************************************************************************
'* [概  要] Reverse
'* [詳  細] 文字列を反転させる｡
'*
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function reverse() As StringEx
    Set reverse = Me.Init(StringUtils.reverse(Me.ToString))
End Function


'******************************************************************************
'* [概  要] RightPad
'* [詳  細] 文字列を指定したパディング文字列で右詰めする｡
'*          デフォルトのパディング文字列は半角スペース。
'*
'* @param size パディングするサイズ
'* @param padStr パディング文字列
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function RightPad(size As Long, Optional padStr As String = " ") As StringEx
    Set RightPad = Me.Init(StringUtils.RightPad(Me.ToString, size, padStr))
End Function


'******************************************************************************
'* [概  要] Rotate
'* [詳  細] 文字列を循環シフトさせる｡
'*          シフト > 0の場合、右循環シフト（例：ABCDEF => FABCDE）
'*          シフト < 0の場合、左循環シフト（例：ABCDEF => BCDEFA）
'*
'* @param shiftCount シフトする回数（正：右シフト、負：左シフト）
'* @return String 変換後文字列
'*
'******************************************************************************
Public Function Rotate(shiftCount As Long) As StringEx
    Set Rotate = Me.Init(StringUtils.Rotate(Me.ToString, shiftCount))
End Function

'******************************************************************************
'* [概  要] StartsWithAny
'* [詳  細] 指定した文字列のいずれかで始まるかチェック｡
'*
'* @param searchStrings 開始文字列（複数指定）
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function StartsWithAny(ParamArray searchStrings() As Variant) As Boolean
    StartsWithAny = StringUtils.StartsWithAny(Me.ToString, searchStrings)
End Function

'******************************************************************************
'* [概  要] StartsWithIgnoreCase
'* [詳  細] 指定した文字列で始まるかチェック｡
'*          接尾辞の存在チェックは大文字小文字を考慮しない。
'*
'* @param prefix 開始文字列
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function StartsWithIgnoreCase(prefix As String) As Boolean
    StartsWithIgnoreCase = StringUtils.StartsWithIgnoreCase(Me.ToString, prefix)
End Function

'******************************************************************************
'* [概  要] Substring
'* [詳  細] 指定された文字列から部分文字列を取得する｡
'*          負の開始位置を使用して、文字列の末尾から開始/終了位置を指定可能。
'*          返される部分文字列は、開始位置の文字で始まり、終了位置の前で終わる。
'*          位置はゼロベースでカウントされる。つまり、文字列の先頭から開始する
'*          には、start = 0を使用する。
'*          開始する位置の指定が終了する位置と逆転していた場合、 ""が返される。
'*
'* @param startPos 開始する位置
'* @param endPos 終了する位置
'* @return StringEx 開始位置から終了位置までの部分文字列
'*
'******************************************************************************
Public Function Substring(startPos As Long, Optional endPos) As StringEx
    Set Substring = Me.Init(StringUtils.Substring(Me.ToString, startPos, endPos))
End Function


'******************************************************************************
'* [概  要] SubstringAfter
'* [詳  細] セパレータが最初に出現した後の部分文字列を取得する。
'*
'* @param separator 検索する文字列
'* @return StringEx セパレータが最初に出現した後の部分文字列
'*
'******************************************************************************
Public Function SubstringAfter(separator As String) As StringEx
    Set SubstringAfter = Me.Init(StringUtils.SubstringAfter(Me.ToString, separator))
End Function


'******************************************************************************
'* [概  要] SubstringAfterLast
'* [詳  細] セパレータが最後に出現した後の部分文字列を取得する｡
'*
'* @param separator 検索する文字列
'* @return StringEx セパレータが最後に出現した後の部分文字列
'*
'******************************************************************************
Public Function SubstringAfterLast(separator As String) As StringEx
    Set SubstringAfterLast = Me.Init(StringUtils.SubstringAfterLast(Me.ToString, separator))
End Function


'******************************************************************************
'* [概  要] SubstringBefore
'* [詳  細]セパレータが最初に出現する前の部分文字列を取得する｡
'*
'* @param separator 検索する文字列
'* @return StringEx セパレータが最初に出現する前の部分文字列
'*
'******************************************************************************
Public Function SubstringBefore(separator As String) As StringEx
    Set SubstringBefore = Me.Init(StringUtils.SubstringBefore(Me.ToString, separator))
End Function

'******************************************************************************
'* [概  要] SubstringBeforeLast
'* [詳  細] セパレータが最後に出現する前の部分文字列を取得する｡
'*
'* @param separator 検索する文字列
'* @return StringEx セパレータが最後に出現する前の部分文字列
'*
'******************************************************************************
Public Function SubstringBeforeLast(separator As String) As StringEx
    Set SubstringBeforeLast = Me.Init(StringUtils.SubstringBeforeLast(Me.ToString, separator))
End Function

'******************************************************************************
'* [概  要] SubstringBetween
'* [詳  細] ２つの文字列の間にネストされている文字列を取得する｡
'*
'* @param openStr 部分文字列の開始を識別する文字列
'* @param closeStr 部分文字列の終わりを識別する文字列
'* @return StringEx 部分文字列
'*
'******************************************************************************
Public Function SubstringBetween(tagOrOpenStr As String, Optional closeStr As String = "") As StringEx
    Set SubstringBetween = Me.Init(StringUtils.SubstringBetween(Me.ToString, tagOrOpenStr, closeStr))
End Function


'******************************************************************************
'* [概  要] ToEncodedString
'* [詳  細] 指定された文字エンコーディングを使用して、バイト配列を文字列に
'*          変換する｡
'*
'* @param bytes() 読込むバイト配列
'* @param charset 使用するエンコーディング
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function ToEncodedString(Bytes() As Byte, CharSet As String) As StringEx
    Set ToEncodedString = Me.Init(StringUtils.ToEncodedString(Bytes(), CharSet))
End Function

'******************************************************************************
'* [概  要] Truncate
'* [詳  細] 文字列を切り捨てる｡
'*
'* @param maxWidth 結果文字列の最大長
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function Truncate(maxWidth As Long) As StringEx
    Set Truncate = Me.Init(StringUtils.Truncate(Me.ToString, maxWidth))
End Function

'******************************************************************************
'* [概  要] Uncapitalize
'* [詳  細] 文字列の最初の文字を小文字に変更する｡
'*
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function Uncapitalize() As StringEx
    Set Uncapitalize = Me.Init(StringUtils.Uncapitalize(Me.ToString))
End Function

'******************************************************************************
'* [概  要] Unwrap
'* [詳  細] 指定された文字列をアンラップする｡
'*
'* @param wrapToken アンラップ文字列
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function Unwrap(wrapToken As String) As StringEx
    Set Unwrap = Me.Init(StringUtils.Unwrap(Me.ToString, wrapToken))
End Function

'******************************************************************************
'* [概  要] UpperCase
'* [詳  細] 文字列を大文字に変換する｡
'*
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function UpperCase() As StringEx
    Set UpperCase = Me.Init(StringUtils.UpperCase(Me.ToString))
End Function

'******************************************************************************
'* [概  要] Wide
'* [詳  細] 半角文字を全角文字に変換する｡
'*
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function Wide() As StringEx
    Set Wide = Me.Init(StringUtils.Wide(Me.ToString))
End Function

'******************************************************************************
'* [概  要] Wrap
'* [詳  細] 指定された文字列をラップする｡
'*
'* @param wrapWith ラップ文字列
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function Wrap(wrapWith As String) As StringEx
    Set Wrap = Me.Init(StringUtils.Wrap(Me.ToString, wrapWith))
End Function

'******************************************************************************
'* [概  要] WrapIfMissing
'* [詳  細] 文字列が指定された文字列の先頭または末尾にない場合、
'*          文字列をラップする｡
'*
'* @param wrapWith ラップ文字列
'* @return StringEx 変換後文字列
'*
'******************************************************************************
Public Function WrapIfMissing(wrapWith As String) As StringEx
    Set WrapIfMissing = Me.Init(StringUtils.WrapIfMissing(Me.ToString, wrapWith))
End Function


'******************************************************************************
'* [概  要] ToCharArray
'* [詳  細] 文字列を１文字ずつの配列に変換する｡
'*
'* @return Variant 変換後文字配列
'*
'******************************************************************************
Public Function ToCharArray() As Variant
    ToCharArray = StringUtils.ToCharArray(Me.ToString)
End Function

'******************************************************************************
'* [概  要] ToCharArrayReverse
'* [詳  細] 文字列を並び順を逆にして１文字ずつの配列に変換する｡
'*
'* @return Variant 変換後文字配列
'*
'******************************************************************************
Public Function ToCharArrayReverse() As Variant
    ToCharArrayReverse = StringUtils.ToCharArrayReverse(Me.ToString)
End Function

'******************************************************************************
'* [概  要] CharAt
'* [詳  細] 文字列から指定したインデックス位置の文字を取得する｡
'*
'* @param idx インデックス
'* @return StringEx 文字
'*
'******************************************************************************
Public Function CharAt(idx As Long) As StringEx
    Set CharAt = Me.Init(StringUtils.CharAt(Me.ToString, idx))
End Function


'*******************************************************************************
'* [概  要] LeftPad0 左ゼロ埋め処理
'* [詳  細] 指定した桁数になる範囲で左ゼロ埋め編集を行う。
'*
'* @param size 桁数
'* @return StringEx 左ゼロ埋めした値
'*
'*******************************************************************************
Public Function LeftPad0(ByVal size As Long) As StringEx
    Set LeftPad0 = Me.Init(StringUtils.LeftPad0(Me.ToString, size))
End Function

'*******************************************************************************
'* [概  要] RightPad0 右ゼロ埋め処理
'* [詳  細] 指定した桁数になる範囲で右ゼロ埋め編集を行う。
'*
'* @param size 桁数
'* @return StringEx 右ゼロ埋めした値
'*
'*******************************************************************************
Public Function RightPad0(ByVal size As Long) As StringEx
    Set RightPad0 = Me.Init(StringUtils.RightPad0(Me.ToString, size))
End Function

'*******************************************************************************
'* [概  要] EscapeSqlQuot シングルクォーテーションエスケープ処理
'* [詳  細] SQLに埋め込む文字列にシングルクォーテーションがある場合にエスケープ
'*          を行うための処理。
'*
'* @return StringEx エスケープ後文字列
'*
'*******************************************************************************
Public Function EscapeSqlQuot() As StringEx
    Set EscapeSqlQuot = Me.Init(StringUtils.EscapeSqlQuot(Me.ToString))
End Function

'*******************************************************************************
'* [概  要] SplitByLength
'* [詳  細] 指定した文字数で文字列を分割する。
'*
'* @param str Variant型データ
'* @param cutLength 文字数
'* @return Variant Variant文字配列
'*
'*******************************************************************************
Public Function SplitByLength(ByVal cutLength) As Variant
    SplitByLength = StringUtils.SplitByLength(Me.ToString, cutLength)
End Function

'*******************************************************************************
'* [概  要] EditTextByFunc
'* [詳  細] FuncオブジェクトまたはFunction名を指定して、テキストを編集する。
'*
'* @param fun FuncオブジェクトまたはFunction名
'* @return StringEx 結果文字列
'*
'*******************************************************************************
Public Function EditTextByFunc(ByRef fun As Variant) As StringEx
    Set EditTextByFunc = Me.Init(StringUtils.EditTextByFunc(Me.ToString, fun))
End Function

'******************************************************************************
'* [概  要] SplitCsvRow CSV行分割処理。
'* [詳  細] CSV文字列をカンマごとに分割してVariant配列に格納し、返却する。
'*          項目両端の「"」を削除する。
'*
'* @param quotExists ダブルクォート存在有無（True：あり）。省略可能。
'* @param delimiter デリミタ。省略可能。
'* @return 分割後Variant配列
'******************************************************************************
Public Function SplitCsvRow(Optional ByVal QuotExists As Boolean = False, Optional Delimiter As String = ",") As Variant
    SplitCsvRow = StringUtils.SplitCsvRow(Me.ToString, QuotExists, Delimiter)
End Function

'******************************************************************************
'* [概  要] RemoveLineSeparator
'* [詳  細] 改行コード除去処理。
'*
'* @param val 文字列
'* @return StringEx 編集後文字列
'******************************************************************************
Public Function RemoveLineSeparator() As StringEx
    Set RemoveLineSeparator = Me.Init(StringUtils.RemoveLineSeparator(Me.ToString))
End Function

'******************************************************************************
'* [概  要] EscapeDoubleQuote
'* [詳  細] ダブルクォーテーションエスケープ処理。
'*
'* @return StringEx 編集後文字列
'******************************************************************************
Public Function EscapeDoubleQuote() As StringEx
    Set EscapeDoubleQuote = Me.Init(StringUtils.EscapeDoubleQuote(Me.ToString))
End Function

