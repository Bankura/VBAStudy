VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] File/Folder操作ユーティリティクラス
'* [詳  細] ファイル/フォルダ操作に関するユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'*
'* [参  考] https://dobon.net/vb/dotnet/file/index.html
'*          https://popozure.info/20190515/14201
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

'******************************************************************************
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 定数定義
'******************************************************************************
Private Const FO_DELETE = &H3               ' 削除
Private Const FOF_ALLOWUNDO = &H40          ' ごみ箱に送る
Private Const FOF_NOCONFIRMATION = &H10&    '確認なし
Private Const FOF_NOERRORUI = &H400&        'エラーのダイアログを表示しない

'******************************************************************************
'* 構造体定義
'******************************************************************************
Private Type SHFILEOPSTRUCT
        hWnd As LongPtr
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAnyOperationsAborted As Long
        hNameMappings As LongPtr
        lpszProgressTitle As String
End Type

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mWinApi As WinAPI ' TODO:廃止予定

'******************************************************************************
'* プロパティ定義
'******************************************************************************

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mWinApi = Base.GetWinAPI()
End Sub
Private Sub Class_Terminate()
    Set mWinApi = Nothing
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] ディレクトリパス分離符付与処理。
'* [詳  細] ディレクトリパスの末尾に分離符（\）がなければ付与を行う。
'*
'* @param strDirPath ディレクトリパス
'* @return 分離符付きディレクトリパス
'******************************************************************************
Public Function AddPathSeparator(strDirPath As String) As String
    If Right(strDirPath, 1) <> "\" Then
        AddPathSeparator = strDirPath & "\"
    Else
        AddPathSeparator = strDirPath
    End If
End Function

'******************************************************************************
'* [概  要] FileExists メソッド
'* [詳  細] ファイルが存在するかどうか判定｡
'*
'* @param targetFilePath ファイルパス
'* @return 判定結果（True:存在）
'*
'******************************************************************************
Public Function FileExists(ByVal targetFilePath As String) As Boolean
    FileExists = IO.FileExists(targetFilePath)
End Function

'******************************************************************************
'* [概  要] DirExists
'* [詳  細] フォルダ存在チェック
'*
'* @param strDirPath チェック対象のフォルダパス
'* @return Boolean 処理結果（True:正常 False：異常）
'******************************************************************************
Public Function DirExists(strDirPath As String) As Boolean
    On Error GoTo ErrorHandler

    'フォルダが存在しない場合は作成するか確認する
    If Dir(strDirPath, vbDirectory) = "" Then
        Dim ret As Long: ret = MsgBox("以下のフォルダは存在しません。作成しますか？" & _
                                   vbCrLf & vbCrLf & strDirPath, vbOKCancel, "フォルダ作成")
        If ret = vbOK Then
            Dim dirRet As LongPtr: dirRet = mWinApi.SHCreateDirectoryEx(0, strDirPath, 0)
            If dirRet = 0 Then
                DirExists = True
                Exit Function
            Else
                DirExists = False
                MsgBox "フォルダを作成できませんでした。"
                Exit Function
            End If
        End If
        DirExists = False
    Else
        DirExists = True
    End If

    Exit Function
    
ErrorHandler:
    DirExists = False
End Function

'******************************************************************************
'* [概  要] SaveToFileWithoutBom
'* [詳  細] ストリームからBOMを取り除いてファイルに保存する。
'*          AriawaseのIOモジュールラッパーメソッド。
'*
'* @param strm ADODB.Streamオブジェクト
'* @param fpath ファイルパス
'* @param opSave 保存オプション（SaveOptionsEnum列挙体の値）
'*
'******************************************************************************
Public Sub SaveToFileWithoutBom( _
    ByVal strm As Object, ByVal fPath As String, ByVal opSave As SaveOptionsEnum _
    )
    Call IO.SaveToFileWithoutBom(strm, fPath, opSave)
End Sub

'******************************************************************************
'* [概  要] RemoveBom
'* [詳  細] ファイルからBOMを取り除いて保存する。
'*          AriawaseのIOモジュールラッパーメソッド。
'*
'* @param fpath ファイルパス
'* @param chrset 文字コード
'* @param linsep 改行コード
'*
'******************************************************************************
Public Sub RemoveBom( _
    ByVal fPath As String, ByVal chrset As String, ByVal linsep As LineSeparatorsEnum _
    )
    Call IO.RemoveBom(fPath, chrset, linsep)
End Sub

'******************************************************************************
'* [概  要] ChangeCharset
'* [詳  細] ストリームの文字コードを変更する。
'*          AriawaseのIOモジュールラッパーメソッド。
'*
'* @param strm ADODB.Streamオブジェクト
'* @param chrset 変更先文字コード
'* @return Object 処理結果（True:正常 False：異常）
'*
'******************************************************************************
Public Function ChangeCharset(ByVal strm As Object, ByVal chrset As String) As Object
    ChangeCharset = IO.ChangeCharset(strm, chrset)
End Function

'******************************************************************************
'* [概  要] ChangeFileCharset
'* [詳  細] ファイルの文字コードを変更して保存する。
'*          AriawaseのIOモジュールラッパーメソッド。
'*
'* @param fpath ファイルパス
'* @param crrChrset 文字コード
'* @param chgChrset 変更後文字コード
'*
'******************************************************************************
Public Sub ChangeFileCharset( _
    ByVal fPath As String, ByVal crrChrset As String, ByVal chgChrset As String _
    )
    Call IO.ChangeFileCharset(fPath, crrChrset, chgChrset)
End Sub

'******************************************************************************
'* [概  要] ChangeLineSeparator
'* [詳  細] ストリームの改行コードを変更する。
'*          AriawaseのIOモジュールラッパーメソッド。
'*
'* @param strm ADODB.Streamオブジェクト
'* @param linsep 変更先改行コード
'* @return Object ADODB.Streamオブジェクト（改行コード変更後）
'*
'******************************************************************************
Public Function ChangeLineSeparator( _
    ByVal strm As Object, ByVal linsep As LineSeparatorsEnum _
    ) As Object
    ChangeLineSeparator = IO.ChangeLineSeparator(strm, linsep)
End Function

'******************************************************************************
'* [概  要] ChangeFileLineSeparator
'* [詳  細] ファイルの改行コードを変更して保存する。
'*          AriawaseのIOモジュールラッパーメソッド。
'*
'* @param fpath ファイルパス
'* @param chrset 文字コード
'* @param crrLinsep 改行コード
'* @param chgLinsep 変更後改行コード
'*
'******************************************************************************
Public Sub ChangeFileLineSeparator( _
    ByVal fPath As String, ByVal chrset As String, _
    ByVal crrLinsep As LineSeparatorsEnum, ByVal chgLinsep As LineSeparatorsEnum _
    )
    Call IO.ChangeFileLineSeparator(fPath, chrset, crrLinsep, chgLinsep)
End Sub

'******************************************************************************
'* [概  要] IsPathRooted
'* [詳  細] 絶対パスかどうかをチェックする。
'*          AriawaseのIOモジュールラッパーメソッド。
'*
'* @param fpath ファイル/フォルダパス
'* @return Boolean チェック結果（True:該当）
'*
'******************************************************************************
Public Function IsPathRooted(ByVal fPath As String) As Boolean
    IsPathRooted = IO.IsPathRooted(fPath)
End Function

'******************************************************************************
'* [概  要] GetSpecialFolder
'* [詳  細] Windowsの特殊ディレクトリパスを取得する。
'*          AriawaseのIOモジュールラッパーメソッド。
'*
'* @param spFolder IO.SpFolderEnum列挙体の値
'*                 WshShell.SpecialFoldersに指定可能な文字列
'* @return String 特殊ディレクトリパス
'* @seealso https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/getspecialfolder-method
'* @seealso https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364490(v=msdn.10)
'*
'******************************************************************************
Public Function GetSpecialFolder(ByVal spFolder As Variant) As String
    GetSpecialFolder = IO.GetSpecialFolder(spFolder)
End Function

'******************************************************************************
'* [概  要] GetTempFilePath
'* [詳  細] 一意となるTempファイルパスを生成する。
'*          AriawaseのIOモジュールラッパーメソッド。
'*
'* @param tdir Tempファイル作成ディレクトリ（省略時は標準Tempフォルダ）
'* @param extName Tempファイル拡張子（省略時は.tmp）
'* @param willCreateFile Trueの場合ファイル作成も行う
'* @return String Tempファイルパス
'*
'******************************************************************************
Public Function GetTempFilePath( _
    Optional ByVal tdir As String, Optional extName As String = ".tmp", _
    Optional willCreateFile As Boolean = False) As String
    GetTempFilePath = IO.GetTempFilePath(tdir, extName)
    If willCreateFile Then
        Call IO.fso.CreateTextFile(GetTempFilePath)
    End If
End Function

'******************************************************************************
'* [概  要] GetUniqueFileName
'* [詳  細] 指定したファイルパスのファイルが存在する場合、ファイルベース名末尾
'*          にデリミタ + 数字を加えて、一意となるファイル名を生成する。
'*          AriawaseのIOモジュールラッパーメソッド。
'*
'* @param fpath ファイル名
'* @param delim デリミタ
'* @return String 一意となるファイル名
'*
'******************************************************************************
Public Function GetUniqueFileName( _
    ByVal fPath As String, Optional delim As String = "_" _
    ) As String
    GetUniqueFileName = IO.GetUniqueFileName(fPath, delim)
End Function

'******************************************************************************
'* [概  要] GetAllFolders
'* [詳  細] 指定したパス配下のすべてのフォルダ（サブフォルダ内含め）の
'*          フォルダパス文字列を取得し、Variant配列として返却する。
'*          AriawaseのIOモジュールラッパーメソッド。
'*
'* @param folderPath フォルダパス
'* @return Variant フォルダパスのリスト（1次元配列）
'*
'******************************************************************************
Public Function GetAllFolders(ByVal folderPath As String) As Variant
    GetAllFolders = IO.GetAllFolders(folderPath)
End Function

'******************************************************************************
'* [概  要] GetAllFiles
'* [詳  細] 指定したパス配下のすべてのファイル（サブフォルダ内含め）の
'*          ファイルパス文字列を取得し、Variant配列として返却する。
'*          AriawaseのIOモジュールラッパーメソッド。
'*
'* @param folderPath フォルダパス
'* @return Variant ファイルパスのリスト（1次元配列）
'*
'******************************************************************************
Public Function GetAllFiles(ByVal folderPath As String) As Variant
    GetAllFiles = IO.GetAllFiles(folderPath)
End Function

'******************************************************************************
'* [概  要] CreateFolderTree
'* [詳  細] 指定したパスの階層毎に存在しないフォルダがあれば作成する。
'*          AriawaseのIOモジュールラッパーメソッド。
'*
'* @param folderPath フォルダパス
'*
'******************************************************************************
Public Sub CreateFolderTree(ByVal folderPath As String)
    Call IO.CreateFolderTree(folderPath)
End Sub

'******************************************************************************
'* [概  要] GetFileTime
'* [詳  細] FileTime を取得する
'*
'* @param dtSetting 設定日時
'* @return FileTime ファイル日時
'*
'******************************************************************************
Private Function GetFileTime(ByVal dtSetting As Date) As FILETIME
    Dim tSystemTime As SystemTime

    With tSystemTime
        .wYear = Year(dtSetting)
        .wMonth = Month(dtSetting)
        .wDayOfWeek = Weekday(dtSetting)
        .wDay = Day(dtSetting)
        .wHour = hour(dtSetting)
        .wMinute = Minute(dtSetting)
        .wSecond = second(dtSetting)
    End With

    Dim tLocalTime As FILETIME
    Call mWinApi.SystemTimeToFileTime(tSystemTime, tLocalTime)

    Dim tFileTime As FILETIME
    Call mWinApi.LocalFileTimeToFileTime(tLocalTime, tFileTime)

    GetFileTime = tFileTime
End Function

'******************************************************************************
'* [概  要] GetFileHandle
'* [詳  細] ファイルのハンドルを取得する
'*
'* @param stFilePath 対象ファイルのパス
'* @return LongPtr ハンドル番号
'*
'******************************************************************************
Private Function GetFileHandle(ByVal stFilePath As String) As LongPtr
    GetFileHandle = mWinApi.CreateFile( _
                        stFilePath, GENERIC_READ Or GENERIC_WRITE, _
                        FILE_SHARE_READ, 0, OPEN_EXISTING, _
                        FILE_ATTRIBUTE_NORMAL, 0 _
                    )
End Function

'******************************************************************************
'* [概  要] SetFileTimeProc
'* [詳  細] ファイルの日時を指定した日付と時間に設定する
'*
'* @param stFilePath 対象ファイルのパス
'* @param dtCreateTime 設定する日付と時間
'* @param onCreationTime 作成日時に設定するか
'* @param onLastWriteTime 更新日時に設定するか
'* @param onLastAccessTime アクセス日時に設定するか
'*
'******************************************************************************
Private Sub SetFileTimeProc(ByVal stFilePath As String, ByVal dtCreateTime As Date, _
                            ByVal onCreationTime As Boolean, _
                            ByVal onLastWriteTime As Boolean, _
                            ByVal onLastAccessTime As Boolean)
    ' FileTime を取得する
    Dim tFileTime As FILETIME
    tFileTime = GetFileTime(dtCreateTime)

    ' ファイルのハンドルを取得する
    Dim cFileHandle As LongPtr
    cFileHandle = GetFileHandle(stFilePath)

    ' ファイルのハンドルが取得できた場合のみ「作成日時」を更新する
    If cFileHandle >= 0 Then
        Dim tNullable As FILETIME
        Dim tCreationTime As FILETIME
        Dim tLastAccessTime As FILETIME
        Dim tLastWriteTime As FILETIME
        
        If onCreationTime Then tCreationTime = tFileTime Else tCreationTime = tNullable
        If onLastAccessTime Then tLastAccessTime = tFileTime Else tLastAccessTime = tNullable
        If onLastWriteTime Then tLastWriteTime = tFileTime Else tLastWriteTime = tNullable

        Call mWinApi.SetFileTime(cFileHandle, _
                                 tCreationTime, _
                                 tLastAccessTime, _
                                 tLastWriteTime _
                        )
        Call mWinApi.CloseHandle(cFileHandle)
    End If
End Sub

'******************************************************************************
'* [概  要] SetCreationTime
'* [詳  細] ファイルの作成日時を指定した日付と時間に設定する
'*
'* @param stFilePath 対象ファイルのパス
'* @param dtCreateTime 作成日時に設定する日付と時間
'*
'******************************************************************************
Public Sub SetCreationTime(ByVal stFilePath As String, ByVal dtCreateTime As Date)
    Call SetFileTimeProc(stFilePath, dtCreateTime, True, False, False)
End Sub

'******************************************************************************
'* [概  要] SetLastWriteTime
'* [詳  細] ファイルの更新日時を指定した日付と時間に設定する
'*
'* @param stFilePath 対象ファイルのパス
'* @param dtUpdateTime 更新日時に設定する日付と時
'*
'******************************************************************************
Public Sub SetLastWriteTime(ByVal stFilePath As String, ByVal dtUpdateTime As Date)
    Call SetFileTimeProc(stFilePath, dtUpdateTime, False, True, False)
End Sub

'******************************************************************************
'* [概  要] SetLastAccessTime
'* [詳  細] ファイルのアクセス日時を指定した日付と時間に設定する
'*
'* @param stFilePath 対象ファイルのパス
'* @param dtAccessTime アクセス日時に設定する日付と時間
'*
'******************************************************************************
Public Sub SetLastAccessTime(ByVal stFilePath As String, ByVal dtAccessTime As Date)
    Call SetFileTimeProc(stFilePath, dtAccessTime, False, False, True)
End Sub

'******************************************************************************
'* [概  要] SetAllTime
'* [詳  細] ファイルの作成日時・更新日時・アクセス日時を指定した日付と時間に
'          設定する
'*
'* @param stFilePath 対象ファイルのパス
'* @param dtSetTime 日時に設定する日付と時間
'*
'******************************************************************************
Public Sub SetAllTime(ByVal stFilePath As String, ByVal dtSetTime As Date)
    Call SetFileTimeProc(stFilePath, dtSetTime, True, True, True)
End Sub

'******************************************************************************
'* [概  要] ファイルエンコード一括変換処理。
'* [詳  細] 指定したフォルダ内のファイルのエンコードを一括変換する。
'*
'* @param targetFolderName 対象となるフォルダのフルパス
'* @param srcEncode 変更元エンコード
'* @param destEncode 変更先エンコード
'* @param bomInclude BOM有無（省略可。規定はFalse:BOM無）
'******************************************************************************
Public Sub ChangeFilesEncode(targetFolderName As String, srcEncode As String, destEncode As String, _
                            Optional bomInclude As Boolean = False)
    Dim oFolder, oFiles, oFile
    Set oFolder = IO.GetFolder(targetFolderName)
    Set oFiles = oFolder.Files
    
    For Each oFile In oFiles
        Dim filePath As String
        filePath = fso.GetAbsolutePathName(AddPathSeparator(targetFolderName) & oFile.name)
         
        ' ファイルエンコード変換
        Call ChangeFileEncode(filePath, srcEncode, destEncode, bomInclude)
    Next
End Sub

'******************************************************************************
'* [概  要] ファイルエンコード変換処理。
'* [詳  細] 指定したファイルのエンコードを変換する。
'*
'* @param filePath 対象となるファイルのフルパス
'* @param srcEncode 変更元エンコード
'* @param destEncode 変更先エンコード
'* @param bomInclude BOM有無（省略可。規定はFalse:BOM無）
'******************************************************************************
Public Sub ChangeFileEncode(filePath As String, srcEncode As String, destEncode As String, _
                            Optional bomInclude As Boolean = False)
    Dim adoStream1 As Object, adoStream2 As Object
    Set adoStream1 = IO.CreateADODBStream(adTypeText, srcEncode)
    Set adoStream2 = IO.CreateADODBStream(adTypeText, destEncode)
         
    ' 変更元ファイルStream読込
    With adoStream1
        .Open
        .LoadFromFile filePath
    End With
     
    ' 変更先ファイルStream読込
    adoStream2.Open
     
    ' エンコード変換
    adoStream1.copyto adoStream2
    If Not bomInclude Then ExcludeBom adoStream2  ' BOM対応
    adoStream2.SaveToFile filePath, adSaveCreateOverWrite 'ファイル上書指定
     
    ' Streamクローズ
    adoStream2.Close
    adoStream1.Close
End Sub

'******************************************************************************
'* [概  要] ファイル読込・書き込み処理。
'* [詳  細] 指定した読込ファイルのデータを別ファイルに書き込む。
'* [参  考] 大容量データの読み込みについては以下のサイトを参考にした。
'*          <https://mussyu1204.myhome.cx/wordpress/it/?p=720>
'*
'* @param srcFilePath 読込ファイルのフルパス
'* @param srcEncode 読込元エンコード
'* @param srcSep 読込元改行コード
'* @param destFilePath 書込ファイルのフルパス
'* @param destEncode 書込先エンコード
'* @param destSep 書込先改行コード
'* @param func 行編集処理用関数名。
'*             以下のように引数に文字列、戻り値に文字列を返す関数名を指定。
'*             funcName(row As String) As String
'*             指定しない（空文字）場合は、行編集は行わない。
'*             Funcオブジェクトも指定可能。
'*
'* @param chunkSize チャンクサイズ。このサイズを超える読込データの場合は、
'*                  チャンクサイズごとに分割して処理を行う。
'* @param bomInclude BOM有無（省略可。規定はFalse:BOM無）
'******************************************************************************
Public Sub ReadAndWrite(srcFilePath As String, srcEncode As String, srcSep As LineSeparatorsEnum, _
                        destFilePath As String, destEncode As String, destSep As LineSeparatorsEnum, _
                        Optional Func As Variant, _
                        Optional chunkSize As Long = 2048, _
                        Optional bomInclude As Boolean = False)

    Dim inStream As Object, outStream As Object
    Set inStream = IO.CreateADODBStream(adTypeText, srcEncode, srcSep)
    Set outStream = IO.CreateADODBStream(adTypeText, destEncode, destSep)

    With inStream
        .Open
        .LoadFromFile srcFilePath
    End With
     
    outStream.Open

    Dim lines As Variant, lastLine As String
    
    ' 読込データのサイズが指定サイズより大きい場合は分割処理（高速化）実施
    If inStream.size > chunkSize Then
        Do Until inStream.EOS
            Dim tmp As String: tmp = inStream.ReadText(chunkSize)
            lines = Split(tmp, vbLf)
             
            Dim lineCnt As Long: lineCnt = UBound(lines)
            lines(0) = lastLine + lines(0)
            
            Dim i As Long
            For i = 0 To (lineCnt - 1)
                lines(i) = Replace(lines(i), vbCr, "")
                
                lines(i) = StringUtils.EditTextByFunc(lines(i), Func)
                
                outStream.WriteText CStr(lines(i)), adWriteLine
            Next
             
            lastLine = lines(lineCnt)
        Loop
        If lastLine <> "" Then
            outStream.WriteText lastLine, adWriteLine
        End If
    Else
        Do Until inStream.EOS
            Dim tmpLine As String: tmpLine = inStream.ReadText(adReadLine)
            
            tmpLine = StringUtils.EditTextByFunc(tmpLine, Func)
            
            outStream.WriteText tmpLine, adWriteChar
        Loop
    End If

    ' ファイル保存
    outStream.SaveToFile destFilePath, adSaveCreateOverWrite
     
    inStream.Close
    outStream.Close
End Sub

'******************************************************************************
'* [概  要] BOM除外処理。
'* [詳  細] UTF-8テキストのBOMを除外する。
'* [備　考] 呼び出す前にStreamがOpenされていること。
'*          WriteText呼出し直後、またはファイル保存前に実行すること。
'*
'* @param streamObj Open済みStreamオブジェクト
'* @return BOM除外データ
'******************************************************************************
Public Function ExcludeBom(streamObj As Object) As Byte()
    Dim tmpCSet As String:  tmpCSet = streamObj.charSet
    With streamObj
        .Position = 0
        .Type = adTypeBinary
        .Position = IO.BomSize(tmpCSet)  'BOM飛ばし
        ExcludeBom = .Read
        .Position = 0
        .Write ExcludeBom
        .SetEOS
    End With
End Function

'******************************************************************************
'* [概  要] テキストファイル読込処理。
'* [詳  細] テキストファイルの読込を行う（一括読込）。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @param cSet キャラセット（文字コード指定）
'*             省略時は自動判別（"_autodetect_all"）
'*               ※使用可能なCharset例：
'*                 utf-8、shift_jis、euc-jp、ISO-2022-JP、unicode（既定）
'*                 「HKEY_CLASSES_ROOT\MIME\Database\Charset」参照。
'* @return 読込テキストデータ
'******************************************************************************
Public Function ReadTextFile(filePath As String, Optional cSet As String = "_autodetect_all") As String
    Dim data As String
    With IO.CreateADODBStream(adTypeText, cSet)
        .Mode = adModeReadWrite
        .Open
        .LoadFromFile filePath 'ファイル指定
        .Position = 0
        data = .ReadText()
        .Close
    End With

    ReadTextFile = data
End Function

'******************************************************************************
'* [概  要] UTF-8テキストファイル読込処理。
'* [詳  細] UTF-8テキストファイルの読込を行う（一括読込）。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @return 読込テキストデータ
'******************************************************************************
Public Function ReadUTF8TextFile(filePath As String) As String
    ReadUTF8TextFile = ReadTextFile(filePath, "UTF-8")
End Function

'******************************************************************************
'* [概  要] テキストファイル読込処理。
'* [詳  細] テキストファイルを行単位で読込み、Variant配列に設定して返却する。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @param cSet キャラセット（文字コード指定）
'*             省略時は自動判別（"_autodetect_all"）
'*               ※使用可能なCharset例：
'*                 utf-8、shift_jis、euc-jp、ISO-2022-JP、unicode（既定）
'*                 「HKEY_CLASSES_ROOT\MIME\Database\Charset」参照。
'* @param separator 改行コード（省略可。規定はCRLF）
'* @return 読込テキストデータ（Variant配列）
'******************************************************************************
Public Function ReadTextFileToVArray(filePath As String, _
                                     Optional cSet As String = "_autodetect_all", _
                                     Optional separator As LineSeparatorsEnum = LineSeparatorsEnum.adCRLF) As Variant

    Dim vArr(), lineCount As Long, cnt As Long
    Dim stream As Object: Set stream = IO.CreateADODBStream(adTypeText, cSet, separator)
    With stream
        .Mode = adModeReadWrite
        .Open
        .LoadFromFile filePath 'ファイル指定
        lineCount = CountLine(stream)
        ReDim Preserve vArr(lineCount - 1)
        .Position = 0
        Do While Not .EOS
            vArr(cnt) = .ReadText(adReadLine)
            cnt = cnt + 1
        Loop
        .Close
    End With

    ReadTextFileToVArray = vArr
End Function

'******************************************************************************
'* [概  要] ReadCsvFile
'* [詳  細] CSVファイルを行単位で読込み、Variant配列に設定して返却する。
'*
'* @param filePath CSVファイルパス
'* @param headerExists ヘッダ行の有無
'* @param quotExists ""で括っているか
'* @param cSet 文字コード
'* @param separator 改行コード
'* @return 読込テキストデータ（Variant2次元配列）
'******************************************************************************
Public Function ReadCsvFile(ByVal filePath As String, _
                            Optional ByVal headerExists As Boolean = True, _
                            Optional ByVal quotExists As Boolean = False, _
                            Optional ByVal cSet As String = "UTF-8", _
                            Optional ByVal separator As LineSeparatorsEnum = adCRLF) As Variant

    Dim lSize As Long, lCount As Long, i As Long, vRow, vArr()
    Dim stream As Object: Set stream = IO.CreateADODBStream(adTypeText, cSet, separator)
    With stream
        .Mode = adModeReadWrite
        .Open
        .LoadFromFile filePath
        lSize = Me.CountLine(stream)
        
        ' 空ファイルは空Variant配列で返却
        If (headerExists And lSize <= 1) Or (Not headerExists And lSize = 0) Then
            .Close
            Exit Function
        End If
        
        ' 1行目処理
        .Position = 0
        vRow = StringUtils.SplitCsvRow(.ReadText(adReadLine), quotExists)
        If Not headerExists Then
            ReDim vArr(0 To lSize - 1, 0 To UBound(vRow))
            For i = LBound(vRow) To UBound(vRow)
                vArr(lCount, i) = vRow(i)
            Next
            lCount = lCount + 1
        Else
            ReDim vArr(0 To lSize - 2, 0 To UBound(vRow))
        End If
        Dim colCount As Long: colCount = UBound(vRow)

        ' 2行目以降処理
        Do While Not .EOS
            vRow = StringUtils.SplitCsvRow(.ReadText(adReadLine), quotExists)
            If UBound(vRow) <> colCount Then
                .Close
                ReadCsvFile = Empty
                Exit Function
            End If
            
            For i = LBound(vRow) To UBound(vRow)
                vArr(lCount, i) = vRow(i)
            Next
            lCount = lCount + 1
        Loop
        .Close
    End With

    ReadCsvFile = vArr
End Function

'******************************************************************************
'* [概  要] SelectCsvCol
'* [詳  細] CSVファイルを行単位で読込み、指定したIndexの項目をVariant配列に
'*          設定して返却する。
'*
'* @param filePath CSVファイルパス
'* @param colIndexes 取得する列項目のIndex配列（数値、0始まり）
'* @param headerExists ヘッダ行の有無
'* @param quotExists ""で括っているか
'* @param cSet 文字コード
'* @param separator 改行コード
'* @return 読込テキストデータ（Variant2次元配列）
'******************************************************************************
Public Function SelectCsvCol(ByVal filePath As String, _
                            ByVal colIndexes As Variant, _
                            Optional ByVal headerExists As Boolean = True, _
                            Optional ByVal quotExists As Boolean = False, _
                            Optional ByVal cSet As String = "UTF-8", _
                            Optional ByVal separator As LineSeparatorsEnum = adCRLF) As Variant

    Dim lSize As Long, lCount As Long, i As Long, vRow, vArr()
    Dim stream As Object: Set stream = IO.CreateADODBStream(adTypeText, cSet, separator)
    With stream
        .Mode = adModeReadWrite
        .Open
        .LoadFromFile filePath
        lSize = Me.CountLine(stream)
        
        ' 空ファイルは空Variant配列で返却
        If (headerExists And lSize <= 1) Or (Not headerExists And lSize = 0) Then
            .Close
            Exit Function
        End If
        
        If Not IsArray(colIndexes) Then
            If ValidateUtils.IsString(colIndexes) Or MathUtils.IsInt(colIndexes) Then
                colIndexes = Array(colIndexes)
            Else
                .Close
                Exit Function
            End If
        End If
        
        ' 1行目処理
        .Position = 0
        vRow = StringUtils.SplitCsvRow(.ReadText(adReadLine), quotExists)
        If Not headerExists Then
            ReDim vArr(0 To lSize - 1, 0 To UBound(colIndexes))
            For i = LBound(colIndexes) To UBound(colIndexes)
                vArr(lCount, i) = vRow(colIndexes(i))
            Next
            lCount = lCount + 1
        Else
            ReDim vArr(0 To lSize - 2, 0 To UBound(colIndexes))
        End If
        Dim colCount As Long: colCount = UBound(vRow)

        ' 2行目以降処理
        Do While Not .EOS
            vRow = StringUtils.SplitCsvRow(.ReadText(adReadLine), quotExists)
            If UBound(vRow) <> colCount Then
                .Close
                SelectCsvCol = Empty
                Exit Function
            End If
            
            For i = LBound(colIndexes) To UBound(colIndexes)
                vArr(lCount, i) = vRow(colIndexes(i))
            Next
            lCount = lCount + 1
        Loop
        .Close
    End With

    SelectCsvCol = vArr
End Function

'******************************************************************************
'* [概  要] SelectCsvColByRange
'* [詳  細] CSVファイルを行単位で読込み、範囲指定したIndexの項目をVariant配列に
'*          設定して返却する。
'*
'* @param filePath CSVファイルパス
'* @param startIndex 取得する列項目の開始Index（数値、0始まり）
'* @param endIndex 取得する列項目の終了Index（数値、0始まり）
'* @param headerExists ヘッダ行の有無
'* @param quotExists ""で括っているか
'* @param cSet 文字コード
'* @param separator 改行コード
'* @return 読込テキストデータ（Variant2次元配列）
'******************************************************************************
Public Function SelectCsvColByRange(ByVal filePath As String, _
                            ByVal startIndex As Long, _
                            ByVal endIndex As Long, _
                            Optional ByVal headerExists As Boolean = True, _
                            Optional ByVal quotExists As Boolean = False, _
                            Optional ByVal cSet As String = "UTF-8", _
                            Optional ByVal separator As LineSeparatorsEnum = adCRLF) As Variant
    Dim colIndexes: colIndexes = ArrayUtils.Range(startIndex, endIndex)
    SelectCsvColByRange = SelectCsvCol(filePath, colIndexes, headerExists, quotExists, cSet, separator)
End Function

'******************************************************************************
'* [概  要] テキストの行数取得処理。
'* [詳  細] テキストの行数を取得する。
'* [備　考] 呼び出す前にStreamがOpenされていること。
'*
'* @param stream Open済みStreamオブジェクト
'* @return 行数
'******************************************************************************
Public Function CountLine(stream As Object) As Long
    Dim lineCount As Long
    With stream
        .Position = 0
        Do While Not .EOS
            .SkipLine
            lineCount = lineCount + 1
        Loop
    End With

    CountLine = lineCount
End Function
  
'******************************************************************************
'* [概  要] テキストファイル行数取得処理。
'* [詳  細] テキストファイルを読込み、行数を返却する。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @param cSet キャラセット（文字コード指定）
'*             省略時は自動判別（"_autodetect_all"）
'*               ※使用可能なCharset例：
'*                 utf-8、shift_jis、euc-jp、ISO-2022-JP、unicode（既定）
'*                 「HKEY_CLASSES_ROOT\MIME\Database\Charset」参照。
'* @param separator 改行コード（省略可。規定はCRLF）
'* @return 行数
'******************************************************************************
Public Function GetTextFileLineCount(filePath As String, _
                                      Optional cSet As String = "_autodetect_all", _
                                      Optional separator As LineSeparatorsEnum = adCRLF) As Long
    ' ファイルが存在しない場合は-1返却
    If (IO.fso.FileExists(filePath) = False) Then
        GetTextFileLineCount = -1
        Exit Function
    End If
    If Not separator = adCR Then
        ' 追記モードで開き行数をカウント
        GetTextFileLineCount = IO.fso.OpenTextFile(filePath, ForAppending).line - 1
        Exit Function
    End If
    
    Dim lineCount As Long
    Dim stream As Object: Set stream = IO.CreateADODBStream(adTypeText, cSet, separator)
    With stream
        .Mode = adModeReadWrite
        .Open
        .LoadFromFile filePath 'ファイル指定
        lineCount = .CountLine(stream)
        .Close
    End With

    GetTextFileLineCount = lineCount
End Function

'******************************************************************************
'* [概  要] テキストファイル書込処理。
'* [詳  細] テキストファイルの書込を行う。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @param data 書込テキストデータ
'* @param cSet キャラセット（文字コード指定）
'*               ※使用可能なCharset例：
'*                 utf-8、shift_jis、euc-jp、ISO-2022-JP、unicode（既定）
'*                 「HKEY_CLASSES_ROOT\MIME\Database\Charset」参照。
'* @param willAppend 追記有無（省略可。規定はFalse:追記なし）
'* @param bomInclude BOM有無（省略可。規定はFalse:BOM無）
'******************************************************************************
Public Sub WriteTextFile(filePath As String, data As String, cSet As String, _
                         Optional willAppend As Boolean = False, _
                         Optional bomInclude As Boolean = False)
                         
    Dim stream As Object: Set stream = IO.CreateADODBStream(adTypeText, cSet)
    With stream
        .Mode = adModeReadWrite
        .Open
        If willAppend Then
            AppendMode stream, filePath
        End If
        .WriteText data, adWriteChar
        If Not bomInclude Then ExcludeBom stream  ' BOM対応
            
        .SaveToFile filePath, adSaveCreateOverWrite 'ファイル上書指定
        .Close
    End With
End Sub

'******************************************************************************
'* [概  要] UTF-8テキストファイル書込処理。
'* [詳  細] UTF-8テキストファイルの書込を行う。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @param data 書込テキストデータ
'* @param willAppend 追記有無（省略可。規定はFalse:追記なし）
'* @param bomInclude BOM有無（省略可。規定はFalse:BOM無）
'******************************************************************************
Public Sub WriteUTF8TextFile(filePath As String, data As String, _
                             Optional willAppend As Boolean = False, _
                             Optional bomInclude As Boolean = False)
     Call WriteTextFile(filePath, data, "UTF-8", willAppend, bomInclude)
End Sub

'******************************************************************************
'* [概  要] テキストファイル書込処理。
'* [詳  細] テキストファイルに配列データを行単位で書込を行う。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @param vArray 書込配列データ（Variant）
'* @param cSet キャラセット（文字コード指定）
'*               ※使用可能なCharset例：
'*                 utf-8、shift_jis、euc-jp、ISO-2022-JP、unicode（既定）
'*                 「HKEY_CLASSES_ROOT\MIME\Database\Charset」参照。
'* @param separator 改行コード（省略可。規定はCRLF）
'* @param willAppend 追記有無（省略可。規定はFalse:追記なし）
'* @param bomInclude BOM有無（省略可。規定はFalse:BOM無）
'******************************************************************************
Public Sub WriteTextFileFromVArray(filePath As String, vArray As Variant, cSet As String, _
                                   Optional separator As LineSeparatorsEnum = adCRLF, _
                                   Optional willAppend As Boolean = False, _
                                   Optional bomInclude As Boolean = False)
    Dim tmp
    Dim stream As Object: Set stream = IO.CreateADODBStream(adTypeText, cSet, separator)
    With stream
        .Mode = adModeReadWrite
        .Open
        If willAppend Then
            AppendMode stream, filePath
        End If
        For Each tmp In vArray
            .WriteText CStr(tmp), adWriteLine
        Next
        If Not bomInclude Then ExcludeBom stream  ' BOM対応
        .SaveToFile filePath, adSaveCreateOverWrite 'ファイル上書指定
        .Close
    End With
End Sub

'******************************************************************************
'* [概  要] ファイルアペンド処理。
'* [詳  細] ファイル追記するための処理を行う。
'*          ファイルの存在チェックを行い、存在する場合は対象ファイルの
'*          ポインタを終端へ移動する。
'* [備　考] 呼び出す前にStreamがOpenされていること。
'*
'* @param filePath 書込み対象となるテキストファイルのフルパス
'******************************************************************************
Public Sub AppendMode(stream As Object, filePath As String)
    If IO.fso.FileExists(filePath) Then
        With stream
            .LoadFromFile filePath
            .Position = .size 'ポインタを終端へ
        End With
    End If
End Sub

'******************************************************************************
'* [概  要] CSVファイル書込処理。
'* [詳  細] テキストファイルに配列データを行単位で書込を行う。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @param vDataArray 書込2次元配列データ（Variant）
'* @param vHeaderArray ヘッダ行1次元配列データ（Variant）（省略時はヘッダなし）
'* @param withQuotWrap 項目を""で括るか
'* @param cSet キャラセット（文字コード指定）
'*               ※使用可能なCharset例：
'*                 utf-8、shift_jis、euc-jp、ISO-2022-JP、unicode
'*                 「HKEY_CLASSES_ROOT\MIME\Database\Charset」参照。
'* @param separator 改行コード（省略可。規定はCRLF）
'* @param willAppend 追記有無（省略可。規定はFalse:追記なし）
'* @param bomInclude BOM有無（省略可。規定はFalse:BOM無）
'******************************************************************************
Public Sub WriteCsvFileFromVArray(ByVal filePath As String, _
                                  ByRef vDataArray As Variant, _
                                  Optional ByRef vHeaderArray As Variant, _
                                  Optional ByVal withQuotWrap As Boolean = False, _
                                  Optional cSet As String = "UTF-8", _
                                  Optional separator As LineSeparatorsEnum = adCRLF, _
                                  Optional willAppend As Boolean = False, _
                                  Optional bomInclude As Boolean = False)
    Dim stream As Object: Set stream = IO.CreateADODBStream(adTypeText, cSet, separator)
    With stream
        .Mode = adModeReadWrite
        .Open
        If willAppend Then
            AppendMode stream, filePath
        End If
        
        Dim recText As String
        'ヘッダ行出力
        If Not IsMissing(vHeaderArray) And IsArray(vHeaderArray) Then
            recText = EditCsvRecordInArray(vHeaderArray, withQuotWrap)
            .WriteText recText, adWriteLine
        End If
        'データ行出力
        Dim i As Long, j As Long
        For i = LBound(vDataArray, 1) To UBound(vDataArray, 1)
            recText = EditCsvRecordIn2DArray(vDataArray, i, withQuotWrap)
            .WriteText recText, adWriteLine
        Next

        If Not bomInclude Then ExcludeBom stream  ' BOM対応
        .SaveToFile filePath, adSaveCreateOverWrite 'ファイル上書指定
        .Close
    End With
End Sub

'******************************************************************************
'* [概  要] CSVレコード編集処理。
'* [詳  細] レコード情報をCSV文字列に編集する。
'*
'* @param vArray レコード情報（Variant配列）
'* @param withQuotWrap 項目を""で括るか
'* @return CSV文字列（ヘッダ）
'******************************************************************************
Private Function EditCsvRecordInArray(vArray As Variant, Optional ByVal withQuotWrap As Boolean = False) As String
    Dim arrx As ArrayEx: Set arrx = New ArrayEx
    Dim v
    For Each v In vArray
        If withQuotWrap Then
            arrx.Add StringUtils.EscapeDoubleQuote(v)
        Else
            arrx.Add CStr(v)
        End If
    Next
    If withQuotWrap Then
        EditCsvRecordInArray = arrx.Join(",", """", True)
        Exit Function
    End If
    EditCsvRecordInArray = arrx.Join(",", , True)
End Function

'******************************************************************************
'* [概  要] CSVレコード編集処理。
'* [詳  細] レコード情報をCSV文字列に編集する。
'*
'* @param v2DArray レコード情報（Variant2次元配列）
'* @param rowIndex 行インデックス
'* @param withQuotWrap 項目を""で括るか
'* @return CSV文字列
'******************************************************************************
Private Function EditCsvRecordIn2DArray(v2DArray As Variant, ByVal rowIndex As Long, Optional ByVal withQuotWrap As Boolean = False) As String
    Dim arrx As ArrayEx: Set arrx = New ArrayEx
    Dim i As Long
    For i = LBound(v2DArray, 2) To UBound(v2DArray, 2)
        If withQuotWrap Then
            arrx.Add StringUtils.EscapeDoubleQuote(v2DArray(rowIndex, i))
        Else
            arrx.Add CStr(v2DArray(rowIndex, i))
        End If
    Next
    If withQuotWrap Then
        EditCsvRecordIn2DArray = arrx.Join(",", """", True)
        Exit Function
    End If
    EditCsvRecordIn2DArray = arrx.Join(",", , True)
End Function

'******************************************************************************
'* [概  要] BOM有無チェック処理。
'* [詳  細] テキストファイルにBOMがあるかチェックを行う。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @param charSet 文字コード
'* @param チェック結果（true:BOM有、false:BOM無）
'******************************************************************************
Function IsBomIncluded(filePath As String, charSet As String) As Boolean
    If LCase(charSet) <> "utf-8" And _
       LCase(charSet) <> "utf-16" And LCase(charSet) <> "unicode" Then
        IsBomIncluded = False
        Exit Function
    End If
    
    Dim buf() As Byte
    Dim suspected As String
    Dim v As Variant

    Dim stream As Object: Set stream = IO.CreateADODBStream(adTypeBinary)
    With stream
        .Open
        .LoadFromFile filePath
        buf = .Read(IO.BomSize(charSet))
        .Close
    End With

    For Each v In buf
        suspected = suspected & Hex(v)
    Next
    
    If (LCase(charSet) = "utf-8" And suspected = "EFBBBF") Or _
       ((LCase(charSet) = "utf-16" Or LCase(charSet) = "unicode") And _
        (suspected = "FFFE" Or suspected = "FEFF")) Then
        IsBomIncluded = True
    Else
        IsBomIncluded = False
    End If
End Function

'******************************************************************************
'* [概  要] BOM有無チェック処理。
'* [詳  細] UTF-8テキストファイルにBOMがあるかチェックを行う。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @param チェック結果（true:BOM有、false:BOM無）
'******************************************************************************
Function IsUTF8BomIncluded(filePath As String) As Boolean
    IsUTF8BomIncluded = IsBomIncluded(filePath, "utf-8")
End Function

'******************************************************************************
'* [概  要] ファイルダンプ処理。
'* [詳  細] ファイルを読込み、16進ダンプ文字データを返却する。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @return 16進ダンプ文字データ
'******************************************************************************
Public Function ReadFileToDump(filePath As String) As String
    Dim ret As StringEx, buf As String, i As Long
    Set ret = New StringEx
    Dim stream As Object: Set stream = IO.CreateADODBStream(adTypeBinary)
    With stream
        .Mode = adModeReadWrite
        .Open
        .LoadFromFile filePath 'ファイル指定
        Do While Not .EOS
            buf = .Read(16)
 
            For i = 1 To LenB(buf)
                Dim tmp As String
                tmp = MidB(buf, i, 1)
                tmp = AscB(tmp)
                tmp = Hex(tmp)
                tmp = UCase(tmp)
                tmp = Right("0" & tmp, 2)
                Call ret.Append(tmp).Append(" ")
            Next
             Call ret.Append(vbCrLf)
        Loop
        .Close
    End With

    ReadFileToDump = ret.ToString
End Function

'******************************************************************************
'* [概  要] CreateNullCharFile
'* [詳  細] 指定バイト数分NULL文字を書き込んだファイル作成する。
'*
'* @param destDirPath 出力先ファイルパス
'* @param fileSize 作成ファイルのバイト数
'* @param bufSize バッファサイズ（一度に書き込むサイズ）
'* @return Boolean 実行結果（True:成功）
'******************************************************************************
Public Function CreateNullCharFile(ByVal filePath As String, _
                                   Optional ByVal fileSize As LongPtr = 1024, _
                                   Optional ByVal bufSize As Long = 1024) As Boolean
    If fileSize < 0 Then
        CreateNullCharFile = False
        Exit Function
    End If
    If Not IsValidFilePath(filePath) Then
        CreateNullCharFile = False
        Exit Function
    End If
    Dim tStream As Object
    Set tStream = IO.CreateTextFile(filePath)
    If fileSize = 0 Then
        Call tStream.Close
        CreateNullCharFile = True
        Exit Function
    End If
    
    Dim wCnt As LongPtr: wCnt = fileSize / CLngPtr(bufSize)
    Dim surplus As LongPtr: surplus = fileSize Mod CLngPtr(bufSize)
    
    On Error GoTo ErrorHandler
    Dim i As LongPtr
    For i = 1 To wCnt
        Call tStream.Write(String(bufSize, vbNullChar))
    Next
    If surplus <> 0 Then Call tStream.Write(String(CLng(surplus), vbNullChar))
    
    Call tStream.Close
    CreateNullCharFile = True
    Exit Function
ErrorHandler:
    Call tStream.Close
    CreateNullCharFile = False
End Function

'******************************************************************************
'* [概  要] CreateNullCharFiles
'* [詳  細] 指定フォルダ内に、指定バイト数分NULL文字を書き込んだファイルを
'*          指定数分作成する。
'*
'* @param destDirPath 出力先フォルダパス
'* @param baseFileName ファイルベース名（この名前に連番を付けたファイル名
'*                     のファイルを作成する）
'* @param fileCount 作成ファイル数
'* @param fileSize 作成ファイルのバイト数
'* @param bufSize バッファサイズ（一度に書き込むサイズ）
'* @return Boolean 実行結果（True:成功）
'******************************************************************************
Public Function CreateNullCharFiles(ByVal destDirPath As String, _
                                    ByVal baseFileName As String, _
                                    ByVal fileCount As Long, _
                                    Optional ByVal fileSize As LongPtr = 1024, _
                                    Optional ByVal bufSize As Long = 1024) As Boolean
    If Not IO.FolderExists(destDirPath) Then
        CreateNullCharFiles = False
        Exit Function
    End If
    If fileCount < 1 Then
        CreateNullCharFiles = False
        Exit Function
    End If
    
    Dim v, tmpPath As String, srcPath As String
    For Each v In ArrayUtils.Range(1, fileCount)
        tmpPath = IO.BuildPath(destDirPath, baseFileName & "_" & StringUtils.LeftPad0(v, Len(CStr(fileCount))) & ".txt")
        
        If v = 1 Then
            If Not CreateNullCharFile(tmpPath, fileSize, bufSize) Then
                CreateNullCharFiles = False
                Exit Function
            End If
            srcPath = tmpPath
        Else
            Call IO.CopyFile(srcPath, tmpPath)
        End If
    Next
    CreateNullCharFiles = True
End Function

'******************************************************************************
'* [概  要] CreateFilesFromTemplete
'* [詳  細] 指定したテンプレートとなるファイルを指定フォルダ内に、指定数分
'*          コピーする。
'*
'* @param srcFilePath テンプレートファイルパス（コピー元）
'* @param destDirPath 出力先フォルダパス
'* @param baseFileName ファイルベース名（この名前に連番を付けたファイル名
'*                     のファイルを作成する）
'* @param fileCount 作成ファイル数
'* @return Boolean 実行結果（True:成功）
'******************************************************************************
Public Function CreateFilesFromTemplete(ByVal srcFilePath As String, _
                                        ByVal destDirPath As String, _
                                        ByVal baseFileName As String, _
                                        ByVal fileCount As Long) As Boolean
    If Not IO.FileExists(srcFilePath) Then
        CreateFilesFromTemplete = False
        Exit Function
    End If
    If Not IO.FolderExists(destDirPath) Then
        CreateFilesFromTemplete = False
        Exit Function
    End If
    If fileCount < 1 Then
        CreateFilesFromTemplete = False
        Exit Function
    End If

    Dim v, tmpPath As String
    For Each v In ArrayUtils.Range(1, fileCount)
        tmpPath = IO.BuildPath(destDirPath, baseFileName & "_" & _
                  StringUtils.LeftPad0(v, Len(CStr(fileCount))) & "." & IO.GetExtensionName(srcFilePath))
        Call IO.CopyFile(srcFilePath, tmpPath)
    Next
    CreateFilesFromTemplete = True
End Function

'******************************************************************************
'* [概  要] DeleteFiles
'* [詳  細] 指定したフォルダ内のファイルについて、ファイル名が検索文字列に
'*          部分一致するファイルを削除する。
'*          検索文字列未指定時は、フォルダ内の全てのファイルを削除する。
'*          サブフォルダ内のファイルには適用されない。
'*
'* @param dirPath 対象フォルダパス
'* @param filterStr 検索文字列
'* @return Boolean 実行結果（True:成功）
'******************************************************************************
Public Function DeleteFiles(ByVal dirPath As String, Optional ByVal filterStr As String) As Boolean
    If Not IO.FolderExists(dirPath) Then
        DeleteFiles = False
        Exit Function
    End If

    On Error GoTo ErrorHandler
    Dim f As Object
    For Each f In IO.GetFolder(dirPath).Files
        If filterStr = "" Or (f.name Like "*" & filterStr & "*") Then
            f.Delete
        End If
    Next
    DeleteFiles = True
    Exit Function
ErrorHandler:
    DeleteFiles = False
End Function

'******************************************************************************
'* [概  要] RenameFiles
'* [詳  細] 指定した複数ファイルを指定したファイルパス名でリネームする。
'*
'* @param srcFilePathArray 対象ファイルパス配列
'* @param destFilePathArray 変更先ファイルパス配列
'* @return Boolean 実行結果（True:成功）
'******************************************************************************
Public Function RenameFiles(ByVal srcFilePathArray As Variant, ByVal destFilePathArray As Variant) As Boolean
    If Not IsArray(srcFilePathArray) Or Not IsArray(destFilePathArray) Then
        RenameFiles = False
        Exit Function
    End If
    If LBound(srcFilePathArray) <> LBound(destFilePathArray) Or _
       UBound(srcFilePathArray) <> UBound(destFilePathArray) Then
        RenameFiles = False
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    Dim i As Long
    For i = LBound(srcFilePathArray) To UBound(srcFilePathArray)
        If IO.FileExists(srcFilePathArray(i)) And IO.FolderExists(IO.GetParentFolderName(destFilePathArray(i))) Then
            IO.GetFile(srcFilePathArray(i)).Move destFilePathArray(i)
        Else
            Debug.Print "Skip: src=[" & srcFilePathArray(i) & "] dest=[" & destFilePathArray(i) & "]"
        End If
    Next
    RenameFiles = True
    Exit Function
    
ErrorHandler:
    RenameFiles = False
End Function

'******************************************************************************
'* [概  要] ReplaceFileName
'* [詳  細] 指定したファイルについて、ファイル名の検索文字列に
'*          一致する箇所を指定した文字列に置換し、リネームする。
'*
'* @param filePath 対象ファイルパス
'* @param searchStr 検索文字列
'* @param replaceStr 置換文字列
'* @return Boolean 実行結果（True:成功）
'******************************************************************************
Public Function ReplaceFileName(ByVal filePath As String, ByVal searchStr As String, ByVal replaceStr As String) As Boolean
    If Not IO.FileExists(filePath) Then
        ReplaceFileName = False
        Exit Function
    End If
            
    Dim dirPath As String:  dirPath = IO.GetParentFolderName(filePath)
    Dim fileName As String: fileName = IO.GetBaseName(filePath)
    Dim extention As String:  extention = IO.GetExtensionName(filePath)
    
    Dim destFileName As String
    destFileName = IO.BuildPath(dirPath, Replace(fileName, searchStr, replaceStr) & "." & extention)
    
    IO.GetFile(filePath).Move destFileName
    
    ReplaceFileName = True
End Function

'******************************************************************************
'* [概  要] ReplaceFileNames
'* [詳  細] 指定したフォルダ内のファイルについて、ファイル名の検索文字列に
'*          一致する箇所を指定した文字列に置換し、リネームする。
'*          サブフォルダ内のファイルには適用されない。
'*
'* @param dirPath 対象フォルダパス
'* @param searchStr 検索文字列
'* @param replaceStr 置換文字列
'* @return Boolean 実行結果（True:成功）
'******************************************************************************
Public Function ReplaceFileNames(ByVal dirPath As String, ByVal searchStr As String, ByVal replaceStr As String) As Boolean
    If Not IO.FolderExists(dirPath) Then
        ReplaceFileNames = False
        Exit Function
    End If

    On Error GoTo ErrorHandler
    Dim f As Object
    For Each f In IO.GetFolder(dirPath).Files
        Call ReplaceFileName(f.Path, searchStr, replaceStr)
    Next
    ReplaceFileNames = True
    Exit Function
    
ErrorHandler:
    ReplaceFileNames = False
End Function

'******************************************************************************
'* [概  要] ConcatFile
'* [詳  細] ２つのファイルを連結して１ファイルに出力する。
'*
'* @param destFilePath 出力先ファイルパス
'* @param filePath1 連結先ファイルパス
'* @param filePath2 連結ファイルパス
'* @param bufSize 一度に読み込むバッファサイズ
'* @return Boolean 実行結果（True:成功）
'******************************************************************************
Public Function ConcatFile(ByVal destFilePath As String, ByVal filePath1 As String, ByVal filePath2 As String, _
                            Optional ByVal bufSize As Long = 1048576) As Boolean
    If Not IO.FileExists(filePath1) Or Not IO.FileExists(filePath2) Then
        ConcatFile = False
        Exit Function
    End If
    If Not IsValidFilePath(destFilePath) Then
        ConcatFile = False
        Exit Function
    End If

    ' 1つ目のファイルをコピー
    Call IO.CopyFile(filePath1, destFilePath)

    ' 出力先ファイルをStreamとしてOpen（追記指定）
    Dim outStream As Object
    Set outStream = IO.CreateADODBStream(adTypeBinary)
    outStream.Open
    AppendMode outStream, destFilePath

    ' 書き込み処理
    Call WriteByteStreamFromFile(outStream, filePath2, bufSize)
    
    ' ファイル保存
    outStream.SaveToFile destFilePath, adSaveCreateOverWrite
    outStream.Close
    
    ConcatFile = True
End Function

'******************************************************************************
'* [概  要] ConcatFiles
'* [詳  細] 指定した複数ファイルを連結して１ファイルに出力する。
'*
'* @param destFilePath 出力先ファイルパス
'* @param filePathArray 入力ファイルパス配列
'* @param bufSize 一度に読み込むバッファサイズ
'* @return Boolean 実行結果（True:成功）
'******************************************************************************
Public Function ConcatFiles(ByVal destFilePath As String, ByVal filePathArray As Variant, _
                            Optional ByVal bufSize As Long = 1048576) As Boolean
    If Not IsArray(filePathArray) Then
        ConcatFiles = False
        Exit Function
    End If
    If ArrayUtils.GetLength(filePathArray) < 2 Then
        ConcatFiles = False
        Exit Function
    End If
    Dim v
    For Each v In filePathArray
        If Not IO.FileExists(v) Then
            ConcatFiles = False
            Exit Function
        End If
    Next
    If Not IsValidFilePath(destFilePath) Then
        ConcatFiles = False
        Exit Function
    End If

    ' 出力先ファイルをStreamとしてOpen（追記指定）
    Dim outStream As Object
    Set outStream = IO.CreateADODBStream(adTypeBinary)
    outStream.Open
    AppendMode outStream, destFilePath

    ' 書き込み処理
    For Each v In filePathArray
        Call WriteByteStreamFromFile(outStream, v, bufSize)
    Next
        
    ' ファイル保存
    outStream.SaveToFile destFilePath, adSaveCreateOverWrite
    outStream.Close
    
    ConcatFiles = True
End Function

'******************************************************************************
'* [概  要] WriteByteStreamFromFile
'* [詳  細] 入力ファイルのデータを、Streamに追加で書き込む。
'* [備　考] 呼び出す前に出力先StreamがOpenされていること。
'*
'* @param outStream 出力先Stream
'* @param inFilePath 入力ファイルパス
'* @param bufSize 一度に読み込むバッファサイズ
'******************************************************************************
Private Sub WriteByteStreamFromFile(ByVal outStream As Object, ByVal inFilePath As String, _
                                    Optional ByVal bufSize As Long = 1048576)
    Dim inStream As Object
    Set inStream = IO.CreateADODBStream(adTypeBinary)

    With inStream
        .Open
        .LoadFromFile inFilePath
    End With
    
    ' 書き込み処理
    Do Until inStream.EOS
        Dim tmpBuf As Variant: tmpBuf = inStream.Read(bufSize)
        outStream.Write tmpBuf
        outStream.Flush
    Loop
    inStream.Close
End Sub

'******************************************************************************
'* [概  要] IsValidFilePath
'* [詳  細] 指定したファイルパスが、ファイルパスとして妥当な内容か判定する。
'*
'* @param fPath ファイルパス
'* @param pathSeparator パス区切り文字
'* @return Boolean 判定結果（True:OK）
'******************************************************************************
Public Function IsValidFilePath(ByVal fPath As String, Optional ByVal pathSeparator As String = "\") As Boolean
    If Len(fPath) > 260 Then
        IsValidFilePath = False
        Exit Function
    End If

    ' ディレクトリパス妥当性チェック
    Dim pathElements As Variant: pathElements = Split(fPath, pathSeparator)
    Dim i As Long
    For i = LBound(pathElements) To UBound(pathElements) - 1
        If ValidateUtils.ValidatePattern(pathElements(i), _
        "[\x00-\x1f<>""|]|^(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9]|CLOCK\$)$", True) Then
            IsValidFilePath = False
            Exit Function
        End If
    Next
    
    ' ファイル名妥当性チェック
    Dim fileName As String: fileName = ArrayUtils.GetLast(pathElements)
    If ValidateUtils.ValidatePattern(fileName, _
        "[\x00-\x1f<>:""/\\|?*]|^(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9]|CLOCK\$)(\.|$)|[\. ]$", True) Then
        IsValidFilePath = False
        Exit Function
    End If
    IsValidFilePath = True
End Function

'******************************************************************************
'* [概  要] IsValidFolderPath
'* [詳  細] 指定したフォルダパスが、フォルダパスとして妥当な内容か判定する。
'*
'* @param fPath ファイルパス
'* @param pathSeparator パス区切り文字
'* @return Boolean 判定結果（True:OK）
'******************************************************************************
Public Function IsValidFolderPath(ByVal fPath As String, Optional ByVal pathSeparator As String = "\") As Boolean
    If Len(fPath) > 260 Then
        IsValidFolderPath = False
        Exit Function
    End If

    ' ディレクトリパス妥当性チェック
    Dim pathElements As Variant: pathElements = Split(fPath, pathSeparator)
    Dim i As Long
    For i = LBound(pathElements) To UBound(pathElements)
        If ValidateUtils.ValidatePattern(pathElements(i), _
        "[\x00-\x1f<>""|]|^(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9]|CLOCK\$)$", True) Then
            IsValidFolderPath = False
            Exit Function
        End If
    Next
    IsValidFolderPath = True
End Function

'******************************************************************************
'* [概  要] IsFolder
'* [詳  細] 指定したパスが、フォルダか判定する。
'*
'* @param fPath パス
'* @return Boolean 判定結果（True:OK）
'******************************************************************************
Public Function IsFolder(ByVal fPath As String) As Boolean
    IsFolder = ((GetAttr(fPath) And vbDirectory) > 0)
End Function

'******************************************************************************
'* [概  要] ChangeExtension
'* [詳  細] 指定したファイルパスから、拡張子を変更したファイルパスを取得する。
'*
'* @param fPath ファイルパス
'* @param extension 拡張子
'* @return String 変更後ファイルパス
'******************************************************************************
Public Function ChangeExtension(ByVal fPath As String, ByVal extension As String) As String
    If InStrRev(fPath, ".") = 0 Then
        If extension = "" Then
            ChangeExtension = fPath
        Else
            ChangeExtension = fPath & "." & StringUtils.RemoveStart(extension, ".")
        End If
        Exit Function
    End If
    If extension = "" Then
        ChangeExtension = Mid(fPath, 1, InStrRev(fPath, ".") - 1)
        Exit Function
    End If
    ChangeExtension = Mid(fPath, 1, InStrRev(fPath, ".")) & StringUtils.RemoveStart(extension, ".")
End Function

'******************************************************************************
'* [概  要] GetSubFoldersCount
'* [詳  細] 指定したフォルダ内のサブフォルダを取得する。
'*
'* @param dirPath フォルダパス
'* @return Long サブフォルダ数
'******************************************************************************
Public Function GetSubFoldersCount(ByVal dirPath As String) As Long
    GetSubFoldersCount = IO.GetFolder(dirPath).SubFolders.Count
End Function

'******************************************************************************
'* [概  要] GetFilesCount
'* [詳  細] 指定したフォルダ内のファイル数を取得する。
'*          サブフォルダ内のファイル数は含まれない。
'*
'* @param dirPath フォルダパス
'* @return Long ファイル数
'******************************************************************************
Public Function GetFilesCount(ByVal dirPath As String) As Long
    GetFilesCount = IO.GetFolder(dirPath).Files.Count
End Function

'******************************************************************************
'* [概  要] Grep
'* [詳  細] テキストファイルをGrepする。
'*
'* @param dirPath 対象となるテキストファイルの格納されているフォルダフルパス
'* @param searchPattern 検索パターン（正規表現）
'* @param wildCards ファイルワイルドカード指定（省略可）
'* @param igCase 大文字小文字の区別有無（省略可）
'* @param cSet 文字コート指定（省略可）
'* @param showConsole イミディエイト出力有無（省略可）
'* @param separator 改行コード指定（省略可）
'* @return Variant Grep結果（Variant配列）
'******************************************************************************
Public Function Grep(ByVal dirPath As String, _
                     ByVal searchPattern As String, _
                     Optional ByVal wildCards As String = "*.txt", _
                     Optional ByVal igCase As Boolean = False, _
                     Optional ByVal cSet As String = "auto", _
                     Optional ByVal showConsole As Boolean = True, _
                     Optional ByVal separator As Long) As Variant
    If Not IO.FolderExists(dirPath) Then
        Grep = ArrayUtils.EmptyVariantArray
        Exit Function
    End If
    If searchPattern = "" Then
        Grep = ArrayUtils.EmptyVariantArray
        Exit Function
    End If
    
    Dim reg As Object: Set reg = Base.GetRegExp
    With reg
        .Pattern = searchPattern
        .IgnoreCase = igCase
        .Global = True
    End With
    
    Dim arr As ArrayEx: Set arr = New ArrayEx
    Dim folder As Object: Set folder = IO.GetFolder(dirPath)
    Call GrepImpl(folder, reg, wildCards, arr, cSet, showConsole, separator)
    Grep = arr.ToArray
End Function

'******************************************************************************
'* [概  要] GrepImpl
'* [詳  細] GrepAndReplaceのメイン処理。
'*          テキストファイルをGrepして一致する箇所を置換する。
'*
'* @param folder Scripting.Folderオブジェクト
'* @param reg RegExpオブジェクト
'* @param wildCards ファイルワイルドカード指定（省略可）
'* @param arr ArrayExオブジェクト（OUT）
'* @param cSet 文字コート指定（省略可）
'* @param showConsole イミディエイト出力有無（省略可）
'* @param separator 改行コード指定（省略可）
'******************************************************************************
Private Sub GrepImpl(ByVal folder As Object, _
                     ByVal reg As Object, _
                     ByVal wildCards As String, _
                     ByRef arr As ArrayEx, _
                     Optional ByVal cSet As String = "auto", _
                     Optional ByVal showConsole As Boolean = True, _
                     Optional ByVal separator As Long)
    Dim f As Object
    For Each f In folder.Files
        If f.name Like wildCards Then
            
            Dim tmpSep As LineSeparatorsEnum: tmpSep = LineSeparatorsEnum.adCRLF
            If separator = LineSeparatorsEnum.adCRLF Or _
               separator = LineSeparatorsEnum.adCR Or _
               separator = LineSeparatorsEnum.adLF Then
               tmpSep = separator
            End If
            
            Dim tmpCSet As String: tmpCSet = cSet
            If tmpCSet = "auto" Then
                tmpCSet = GetFileCharset(f.Path)
            End If
            
            If tmpCSet = "" Then
                Call GrepTextFile(f.Path, reg, arr, , showConsole, tmpSep)
            ElseIf tmpCSet <> "BINARY" Then
                If separator = 0 Then
                    tmpSep = GetFileLineSeparator(f.Path)
                End If
                Call GrepTextFile(f.Path, reg, arr, tmpCSet, showConsole, tmpSep)
            Else
                Debug.Print tmpCSet, f.Path
            End If
        End If
    Next
    Dim sfolder As Object
    For Each sfolder In folder.SubFolders
        Call GrepImpl(sfolder, reg, wildCards, arr, cSet, showConsole, separator)
    Next
End Sub

'******************************************************************************
'* [概  要] GrepTextFile
'* [詳  細] テキストファイルを行単位で読込み、Grep結果を返却する。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @param reg RegExpオブジェクト
'* @param arr ArrayExオブジェクト（OUT）
'* @param cSet キャラセット（文字コード指定）
'*             省略時は自動判別（"_autodetect_all"）
'*               ※使用可能なCharset例：
'*                 utf-8、shift_jis、euc-jp、ISO-2022-JP、unicode
'*                 「HKEY_CLASSES_ROOT\MIME\Database\Charset」参照。
'* @param showConsole イミディエイト出力有無（省略可）
'* @param separator 改行コード（省略可。規定はCRLF）
'******************************************************************************
Private Sub GrepTextFile(ByVal filePath As String, _
                         ByVal reg As Object, _
                         ByRef arr As ArrayEx, _
                         Optional ByVal cSet As String = "_autodetect_all", _
                         Optional ByVal showConsole As Boolean = True, _
                         Optional ByVal separator As LineSeparatorsEnum = LineSeparatorsEnum.adCRLF)
    Dim txt As String, cnt As Long: cnt = 1
    Dim stream As Object: Set stream = IO.CreateADODBStream(adTypeText, cSet, separator)

    With stream
        .Mode = adModeReadWrite
        .Open
        .LoadFromFile filePath

        Do While Not .EOS
            Dim tmpCSet As String: tmpCSet = cSet
            txt = TryReadText(stream, filePath, tmpCSet, separator)
            If reg.Test(txt) Then
                If showConsole Then Debug.Print filePath & "(" & cnt & ")  [" & tmpCSet & "]: " & txt
                arr.Add Array(filePath, cnt, txt)
            End If
            cnt = cnt + 1
        Loop
        .Close
    End With
End Sub

'******************************************************************************
'* [概  要] TryReadText
'* [詳  細] テキストファイルのファイルストリームから行文字列を読み込む。
'*          文字コード違いにより読み込みに失敗した場合は、文字コードを
'*          変更して読込を試みる。
'* [備  考] 文字コードが違っていても読み込まれる場合が多いため、確実に
'*          正しい文字コードで読み込まれることは一切保証できない。
'*          あわよくばの精神で使用されたし。
'*
'* @param stream ファイルストリーム
'* @param filePath ファイルパス
'* @param cSet 文字コード
'* @param separator 改行コード指定
'* @return String 読込行文字列
'******************************************************************************
Private Function TryReadText(ByVal stream As Object, _
                     ByVal filePath As String, _
                     ByRef cSet As String, _
                     ByVal separator As LineSeparatorsEnum) As String
    On Error Resume Next
    TryReadText = stream.ReadText(adReadLine)
    If err.Number <> 0 Then
        err.Clear
        Select Case LCase(cSet)
            Case "iso-2022-jp"
                cSet = "euc-jp"
            Case "euc-jp"
                cSet = "unicode"
            Case "unicode"
                cSet = "utf-8"
            Case "utf-8"
                cSet = "shift_jis"
            Case "shift_jis", "shift-jis"
                cSet = "_autodetect_all"
            Case Else
                Exit Function
        End Select
        With stream
            .Close
            .Mode = adModeReadWrite
            .charSet = cSet
            .lineSeparator = separator
            .Type = adTypeText
            .Open
            .LoadFromFile filePath
        End With
        TryReadText = TryReadText(stream, filePath, cSet, separator)
    End If
End Function

'******************************************************************************
'* [概  要] GetFileLineSeparator
'* [詳  細] テキストファイルの改行コードを判定する。
'*
'* @param fileName 対象となるテキストファイルのファイルパス
'* @return LineSeparatorsEnum 改行コード
'******************************************************************************
Public Function GetFileLineSeparator(ByVal fileName As String) As LineSeparatorsEnum
    GetFileLineSeparator = adCRLF

    Dim fNo As Long, fLen As Long, buf() As Byte
    fLen = FileLen(fileName)
    If fLen > 2048 Then
        fLen = 2048
    End If
    ReDim buf(fLen - 1)

    fNo = FreeFile()
    Open fileName For Binary As fNo
    Get fNo, , buf
    Close #fNo
    If (err.Number <> 0) Then
        Exit Function
    End If
    On Error Resume Next
    Dim i As Long, cntCrLf As Long, cntLf As Long, cntCr As Long
    For i = 0 To fLen - 1
        Dim b1 As Byte
        b1 = buf(i)
        
        If b1 = &HA Then
            cntLf = cntLf + 1
        End If
        If b1 = &HD Then
            cntCr = cntCr + 1
        End If
    Next
    If cntLf > 0 And cntCr > 0 Then
        If cntLf > cntCr Then
            If cntCr / cntLf > 0.6 Then
                GetFileLineSeparator = adCRLF
                Exit Function
            End If
        ElseIf cntLf < cntCr Then
            If cntLf / cntCr > 0.6 Then
                GetFileLineSeparator = adCRLF
                Exit Function
            End If
        End If
    End If
    
    If cntLf > cntCr Then
        GetFileLineSeparator = adLF
    ElseIf cntLf < cntCr Then
        GetFileLineSeparator = adCR
    End If
End Function

'******************************************************************************
'* [概  要] GetFileCharset
'* [詳  細] テキストファイルの文字コードを判定する。
'*
'* @param fileName 対象となるテキストファイルのファイルパス
'* @param bomInclude BOMの有無判定（OUT）
'* @return String 文字コード
'*
'*　[参  考] https://popozure.info/20190515/14201
'*           https://dobon.net/vb/dotnet/string/detectcode.html
'******************************************************************************
Public Function GetFileCharset(ByVal fileName As String, Optional ByRef bomInclude As Boolean = False) As String
    Const bEscape As Byte = &H1B
    Const bAt As Byte = &H40
    Const bDollar As Byte = &H24
    Const bAnd As Byte = &H26
    Const bOpen As Byte = &H28 ''('
    Const bB As Byte = &H42
    Const bD As Byte = &H44
    Const bJ As Byte = &H4A
    Const bI As Byte = &H49
    
    Dim i                   As Long
    Dim hdlFile             As Long
    Dim lngFileLen          As Long
    Dim bytFile()           As Byte
    Dim b1                  As Byte
    Dim b2                  As Byte
    Dim b3                  As Byte
    Dim b4                  As Byte
    Dim lngSJIS             As Long
    Dim lngUTF8             As Long
    Dim lngEUC              As Long
    On Error Resume Next
    'ファイル読み込み
    lngFileLen = FileLen(fileName)
    If lngFileLen > 1024 Then
        lngFileLen = 1024 - 1
    End If
    ReDim bytFile(lngFileLen)

    hdlFile = FreeFile()
    Open fileName For Binary As #hdlFile
    Get #hdlFile, , bytFile
    Close #hdlFile
    If (err.Number <> 0) Then
        Exit Function
    End If
    
    'BOMによる判断
    If (bytFile(0) = &HEF And bytFile(1) = &HBB And bytFile(2) = &HBF) Then
        'GetFileCharset = "UTF-8 BOM"
        bomInclude = True
        GetFileCharset = "UTF-8"
        Exit Function
    ElseIf (bytFile(0) = &HFF And bytFile(1) = &HFE) Then
        'GetFileCharset = "UTF-16 LE BOM"
        bomInclude = True
        GetFileCharset = "Unicode"
        Exit Function
    ElseIf (bytFile(0) = &HFE And bytFile(1) = &HFF) Then
        'GetFileCharset = "UTF-16 BE BOM"
        bomInclude = True
        GetFileCharset = "Unicode"
        Exit Function
    End If
    
    'BINARY
    For i = 0 To lngFileLen - 1
        b1 = bytFile(i)
        If b1 = &H0 And i < lngFileLen - 1 And bytFile(i + 1) <= &H7F Then
            'smells like raw unicode
            GetFileCharset = "Unicode"
            Exit Function
        End If
        If (b1 >= &H0 And b1 <= &H8) Or (b1 >= &HA And b1 <= &H9) Or _
           (b1 >= &HB And b1 <= &HC) Or (b1 >= &HE And b1 <= &H19) Or _
           (b1 >= &H1C And b1 <= &H1F) Or (b1 = &H7F) Then
           
            ' unicode判定（変換後「?」が1割以下ならunicode）
            Dim v, badEncCnt As Long, tmpCnt As Long
            badEncCnt = 0: tmpCnt = 0
            For Each v In StringUtils.ToCharArray(StringUtils.ToEncodedString(bytFile, "unicode"))
                If Asc(v) = Asc("?") Then badEncCnt = badEncCnt + 1
                tmpCnt = tmpCnt + 1
            Next
            If (badEncCnt / tmpCnt) < 0.1 Then
                GetFileCharset = "Unicode"
                Exit Function
            End If
            
            ' UTF-16BE判定（変換後「?」が1割以下ならUTF-16BE）
            badEncCnt = 0: tmpCnt = 0
            For Each v In StringUtils.ToCharArray(StringUtils.ToEncodedString(bytFile, "UTF-16BE"))
                If Asc(v) = Asc("?") Then badEncCnt = badEncCnt + 1
                tmpCnt = tmpCnt + 1
            Next
            If (badEncCnt / tmpCnt) < 0.1 Then
                GetFileCharset = "UTF-16BE"
                Exit Function
            End If
            GetFileCharset = "BINARY"
            Exit Function
        End If
    Next i
    
'    間違った判定になりやすいためコメントアウト
'    'not Japanese
'    Dim notJapanese As Boolean: notJapanese = True
'    For i = 0 To lngFileLen - 1
'        b1 = bytFile(i)
'        If b1 = bEscape Or &H80 <= b1 Then
'            notJapanese = False
'            Exit For
'        End If
'    Next
'    If notJapanese Then
'        GetFileCharset = "ASCII"
'        Exit Function
'    End If

    For i = 0 To lngFileLen - 3
        b1 = bytFile(i)
        b2 = bytFile(i + 1)
        b3 = bytFile(i + 2)

        If b1 = bEscape Then
            If b2 = bDollar And b3 = bAt Then
                'JIS_0208 1978
                'JIS
                GetFileCharset = "iso-2022-jp"
                Exit Function
            ElseIf b2 = bDollar And b3 = bB Then
                'JIS_0208 1983
                'JIS
                GetFileCharset = "iso-2022-jp"
                Exit Function
            ElseIf b2 = bOpen And (b3 = bB Or b3 = bJ) Then
                'JIS_ASC
                'JIS
                GetFileCharset = "iso-2022-jp"
                Exit Function
            ElseIf b2 = bOpen And b3 = bI Then
                'JIS_KANA
                'JIS
                GetFileCharset = "iso-2022-jp"
                Exit Function
            End If
            If i < lngFileLen - 3 Then
                b4 = bytFile(i + 3)
                If b2 = bDollar And b3 = bOpen And b4 = bD Then
                    'JIS_0212
                    'JIS
                    GetFileCharset = "iso-2022-jp"
                    Exit Function
                End If
                If i < lngFileLen - 5 And _
                    b2 = bAnd And b3 = bAt And b4 = bEscape And _
                    bytFile(i + 4) = bDollar And bytFile(i + 5) = bB Then
                    'JIS_0208 1990
                    'JIS
                    GetFileCharset = "iso-2022-jp"
                    Exit Function
                End If
            End If
        End If
    Next
    
    'SJIS
    For i = 0 To lngFileLen - 1
        b1 = bytFile(i)
        If (b1 = &H9) Or (b1 = &HA) Or (b1 = &HD) Or (b1 >= &H20 And b1 <= &H7E) Or (b1 >= &HB0 And b1 <= &HDF) Then
            lngSJIS = lngSJIS + 1
        Else
            If (i < lngFileLen - 2) Then
                b2 = bytFile(i + 1)
                If ((b1 >= &H81 And b1 <= &H9F) Or (b1 >= &HE0 And b1 <= &HFC)) And _
                   ((b2 >= &H40 And b2 <= &H7E) Or (b2 >= &H80 And b2 <= &HFC)) Then
                   lngSJIS = lngSJIS + 2
                   i = i + 1
                End If
            End If
        End If
    Next i

    'UTF-8
    For i = 0 To lngFileLen - 1
        b1 = bytFile(i)
        If (b1 = &H9) Or (b1 = &HA) Or (b1 = &HD) Or (b1 >= &H20 And b1 <= &H7E) Then
            lngUTF8 = lngUTF8 + 1
        Else
            If (i < lngFileLen - 2) Then
                b2 = bytFile(i + 1)
                If (b1 >= &HC2 And b1 <= &HDF) And (b2 >= &H80 And b2 <= &HBF) Then
                   lngUTF8 = lngUTF8 + 2
                   i = i + 1
                Else
                    If (i < lngFileLen - 3) Then
                        b3 = bytFile(i + 2)
                        If (b1 >= &HE0 And b1 <= &HEF) And (b2 >= &H80 And b2 <= &HBF) And (b3 >= &H80 And b3 <= &HBF) Then
                            lngUTF8 = lngUTF8 + 3
                            i = i + 2
                        Else
                            If (i < lngFileLen - 4) Then
                                b4 = bytFile(i + 3)
                                If (b1 >= &HF0 And b1 <= &HF7) And (b2 >= &H80 And b2 <= &HBF) And (b3 >= &H80 And b3 <= &HBF) And (b4 >= &H80 And b3 <= &HBF) Then
                                    lngUTF8 = lngUTF8 + 4
                                    i = i + 3
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i

    'EUC-JP
    For i = 0 To lngFileLen - 1
        b1 = bytFile(i)
        If (b1 = &H7) Or (b1 = 10) Or (b1 = 13) Or (b1 >= &H20 And b1 <= &H7E) Then
            lngEUC = lngEUC + 1
        Else
            If (i < lngFileLen - 2) Then
                b2 = bytFile(i + 1)
                If ((b1 >= &HA1 And b1 <= &HFE) And _
                   (b2 >= &HA1 And b2 <= &HFE)) Or _
                   ((b1 = &H8E) And (b2 >= &HA1 And b2 <= &HDF)) Then
                   lngEUC = lngEUC + 2
                   i = i + 1
                End If
            End If
        End If
    Next i

    ' 文字コード出現順位による判断
    If (lngSJIS <= lngUTF8) And (lngEUC <= lngUTF8) Then
        GetFileCharset = "UTF-8"
        Exit Function
    End If
    If (lngUTF8 <= lngSJIS) And (lngEUC <= lngSJIS) Then
        GetFileCharset = "Shift_JIS"
        Exit Function
    End If
    If (lngUTF8 <= lngEUC) And (lngSJIS <= lngEUC) Then
        GetFileCharset = "EUC-JP"
        Exit Function
    End If
    GetFileCharset = "_autodetect_all"
End Function



'******************************************************************************
'* [概  要] GrepAndReplace
'* [詳  細] テキストファイルをGrepして一致する箇所を置換する。
'*
'* @param dirPath 対象となるテキストファイルの格納されているフォルダフルパス
'* @param searchPattern 検索パターン（正規表現）
'* @param replaceString 置換文字列
'* @param wildCards ファイルワイルドカード指定（省略可）
'* @param igCase 大文字小文字の区別有無（省略可）
'* @param cSet 文字コート指定（省略可）
'* @param showConsole イミディエイト出力有無（省略可）
'* @param separator 改行コード指定（省略可）
'* @return Variant Grep結果（Variant配列）
'******************************************************************************
Public Function GrepAndReplace(ByVal dirPath As String, _
                     ByVal searchPattern As String, _
                     ByVal replaceString As String, _
                     Optional ByVal wildCards As String = "*.txt", _
                     Optional ByVal igCase As Boolean = False, _
                     Optional ByVal cSet As String = "auto", _
                     Optional ByVal showConsole As Boolean = True, _
                     Optional ByVal separator As Long) As Variant
    If Not IO.FolderExists(dirPath) Then
        GrepAndReplace = ArrayUtils.EmptyVariantArray
        Exit Function
    End If
    If searchPattern = "" Then
        GrepAndReplace = ArrayUtils.EmptyVariantArray
        Exit Function
    End If
    
    Dim reg As Object: Set reg = Base.GetRegExp
    With reg
        .Pattern = searchPattern
        .IgnoreCase = igCase
        .Global = True
    End With
    
    Dim arr As ArrayEx: Set arr = New ArrayEx
    Dim folder As Object: Set folder = IO.GetFolder(dirPath)
    Call GrepAndReplaceImpl(folder, reg, replaceString, wildCards, arr, cSet, showConsole, separator)
    GrepAndReplace = arr.ToArray
End Function

'******************************************************************************
'* [概  要] GrepAndReplaceImpl
'* [詳  細] GrepAndReplaceのメイン処理。
'*          テキストファイルをGrepして一致する箇所を置換する。
'*
'* @param folder Scripting.Folderオブジェクト
'* @param reg RegExpオブジェクト
'* @param replaceString 置換文字列
'* @param wildCards ファイルワイルドカード指定（省略可）
'* @param arr ArrayExオブジェクト（OUT）
'* @param cSet 文字コート指定（省略可）
'* @param showConsole イミディエイト出力有無（省略可）
'* @param separator 改行コード指定（省略可）
'******************************************************************************
Private Sub GrepAndReplaceImpl(ByVal folder As Object, _
                     ByVal reg As Object, _
                     ByVal replaceString As String, _
                     ByVal wildCards As String, _
                     ByRef arr As ArrayEx, _
                     Optional ByVal cSet As String = "auto", _
                     Optional ByVal showConsole As Boolean = True, _
                     Optional ByVal separator As Long)
    Dim f As Object
    For Each f In folder.Files
        If f.name Like wildCards Then
            Dim bomInclude As Boolean: bomInclude = False
            Dim tmpSep As LineSeparatorsEnum: tmpSep = LineSeparatorsEnum.adCRLF
            If separator = LineSeparatorsEnum.adCRLF Or _
               separator = LineSeparatorsEnum.adCR Or _
               separator = LineSeparatorsEnum.adLF Then
               tmpSep = separator
            End If
            
            Dim tmpCSet As String: tmpCSet = cSet
            If tmpCSet = "auto" Then
                tmpCSet = GetFileCharset(f.Path, bomInclude)
            Else
                bomInclude = IsBomIncluded(f.Path, tmpCSet)
            End If
            
            If tmpCSet = "" Then
                Call GrepAndReplaceTextFile(f.Path, reg, replaceString, arr, , showConsole, True, tmpSep)
            ElseIf tmpCSet <> "BINARY" Then
                If separator = 0 Then
                    tmpSep = GetFileLineSeparator(f.Path)
                End If
                Call GrepAndReplaceTextFile(f.Path, reg, replaceString, arr, tmpCSet, showConsole, bomInclude, tmpSep)
            Else
                Debug.Print tmpCSet, f.Path
            End If
        End If
    Next
    Dim sfolder As Object
    For Each sfolder In folder.SubFolders
        Call GrepAndReplaceImpl(sfolder, reg, replaceString, wildCards, arr, cSet, showConsole, separator)
    Next
End Sub

'******************************************************************************
'* [概  要] GrepAndReplaceTextFile
'* [詳  細] テキストファイルを行単位で読込み、Grep/Replace結果を返却する。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @param reg RegExpオブジェクト
'* @param replaceString 置換文字列
'* @param arr ArrayExオブジェクト（OUT）
'* @param cSet 文字コート指定（省略可）
'* @param showConsole イミディエイト出力有無（省略可）
'* @param bomInclude BOM有無（省略可）
'* @param separator 改行コード（省略可。規定はCRLF）
'******************************************************************************
Private Sub GrepAndReplaceTextFile(ByVal filePath As String, _
                                   ByVal reg As Object, _
                                   ByVal replaceString As String, _
                                   ByRef arr As ArrayEx, _
                                   Optional ByVal cSet As String = "_autodetect_all", _
                                   Optional ByVal showConsole As Boolean = True, _
                                   Optional ByVal bomInclude As Boolean = True, _
                                   Optional ByVal separator As LineSeparatorsEnum = LineSeparatorsEnum.adCRLF)

    Dim txt As String, cnt As Long: cnt = 1
    Dim wtxt As String, hitcnt As Long: hitcnt = 0
    Dim istream As Object: Set istream = IO.CreateADODBStream(adTypeText, cSet, separator)
    Dim ostream As Object: Set ostream = IO.CreateADODBStream(adTypeText, cSet, separator)
    
    With ostream
        .Mode = adModeReadWrite
        .Open
    End With
    
    With istream
        .Mode = adModeReadWrite
        .Open
        .LoadFromFile filePath

        Do While Not .EOS
            Dim tmpCSet As String: tmpCSet = cSet
            txt = .ReadText(adReadLine)
            If reg.Test(txt) Then
                wtxt = reg.Replace(txt, replaceString)
                If showConsole Then Debug.Print filePath & "(" & cnt & ")  [" & tmpCSet & "]: " & txt & " ⇒ " & wtxt
                arr.Add Array(filePath, cnt, txt, wtxt)
                hitcnt = hitcnt + 1
            Else
                wtxt = txt
            End If
            If LCase(tmpCSet) = "utf-16be" Then
                Select Case separator
                    Case LineSeparatorsEnum.adCR
                        ostream.WriteText wtxt & Chr(13), adWriteChar
                    Case LineSeparatorsEnum.adLF
                        ostream.WriteText wtxt & Chr(10), adWriteChar
                    Case LineSeparatorsEnum.adCRLF
                        ostream.WriteText wtxt & Chr(13) & Chr(10), adWriteChar
                End Select
            Else
                ostream.WriteText wtxt, adWriteLine
            End If
            cnt = cnt + 1
        Loop
        .Close
    End With
    
    If hitcnt > 0 Then
        If Not bomInclude Then ExcludeBom ostream  ' BOM対応
        With ostream
            .SaveToFile filePath, adSaveCreateOverWrite 'ファイル上書指定
            .Close
        End With
    Else
        ostream.Close
    End If
End Sub

'******************************************************************************
'* [概  要] SearchFile
'* [詳  細] フォルダ内のファイルを検索する。
'*
'* @param dirPath 対象となるファイルの格納されているフォルダフルパス
'* @param searchFileName 検索ファイル名パターン（正規表現）
'* @param wildCards ファイルワイルドカード指定（省略可）
'* @param igCase 大文字小文字の区別有無（省略可）
'* @return Variant 検索結果（１次元配列：ファイルパス）
'******************************************************************************
Public Function SearchFile(ByVal dirPath As String, _
                           ByVal searchFileName As String, _
                           Optional ByVal wildCards As String = "*.txt", _
                           Optional ByVal igCase As Boolean = False) As Variant
    If Not IO.FolderExists(dirPath) Then
        SearchFile = ArrayUtils.EmptyVariantArray
        Exit Function
    End If
    If searchFileName = "" Then
        SearchFile = ArrayUtils.EmptyVariantArray
        Exit Function
    End If
    
    Dim reg As Object: Set reg = Base.GetRegExp
    With reg
        .Pattern = searchFileName
        .IgnoreCase = igCase
        .Global = True
    End With
    
    Dim arr As ArrayEx: Set arr = New ArrayEx
    Dim folder As Object: Set folder = IO.GetFolder(dirPath)
    Call SearchFileImpl(folder, reg, wildCards, arr)
    SearchFile = arr.ToArray
End Function

'******************************************************************************
'* [概  要] SearchFileImpl
'* [詳  細] フォルダ内の正規表現の指定条件に一致するファイルパスを
'*          配列（ArrayEx）に格納する。
'*
'* @param folder Scripting.Folderオブジェクト
'* @param reg RegExpオブジェクト
'* @param wildCards ファイルワイルドカード指定
'* @param arr ArrayExオブジェクト（OUT）
'******************************************************************************
Private Sub SearchFileImpl(ByVal folder As Object, _
                           ByVal reg As Object, _
                           ByVal wildCards As String, _
                           ByRef arr As ArrayEx)
    Dim f As Object
    For Each f In folder.Files
        If f.name Like wildCards Then
            If reg.Test(f.name) Then
                arr.Add f.Path
            End If
        End If
    Next
    Dim sfolder As Object
    For Each sfolder In folder.SubFolders
        Call SearchFileImpl(sfolder, reg, wildCards, arr)
    Next
End Sub

'******************************************************************************
'* [概  要] MoveTrashBox
'* [詳  細] ゴミ箱に移動する。
'*
'* @param fPath パス
'* @return Boolean 実行結果
'******************************************************************************
Public Function MoveTrashBox(ByVal fPath As String) As Boolean
    Dim sh As SHFILEOPSTRUCT, re As Long
    With sh
        .hWnd = Application.hWnd
        .wFunc = FO_DELETE
        .pFrom = fPath
        .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION + FOF_NOERRORUI
    End With
    re = SHFileOperation(sh)
    If re = 0 Then MoveTrashBox = True
End Function

'******************************************************************************
'* [概  要] GetUniqueFolderName
'* [詳  細] 一意となるフォルダパスを作成する。
'*          指定したフォルダパスが存在する場合は、フォルダ名末尾にデリミタと
'*          番号を付与して現状存在しないフォルダパスとなるまで番号をインクリ
'*          メントする。
'*
'* @param fPath 元となるフォルダパス
'* @param delim パス
'* @return String 一意となるフォルダパス
'******************************************************************************
Public Function GetUniqueFolderName(ByVal fPath As String, Optional delim As String = "_") As String
    Dim d As String: d = fso.GetParentFolderName(fPath)
    Dim b As String: b = fso.GetBaseName(fPath) & delim
    Dim n As Long: n = 0
    Do While IO.FolderExists(fPath)
        n = n + 1
        fPath = fso.BuildPath(d, b & n)
    Loop
    GetUniqueFolderName = fPath
End Function

'******************************************************************************
'* [概  要] CreateZipFile
'* [詳  細] zipファイルを作成する。
'* [参  考] https://qiita.com/kou_tana77/items/72346c69107fabf99335
'*          http://scripting.cocolog-nifty.com/blog/2007/11/vbazip_a144.html
'* [補  足] Microsoft非推奨の方法であるため、Powershell経由でZIP圧縮する方法
'*          をとるのが望ましい。
'*
'* @param filePath 保存先zipファイルパス
'* @param targetFiles()  対象ファイルパス（複数指定可）
'* @return Boolean 実行結果
'******************************************************************************
Public Function CreateZipFile(filePath As String, ParamArray targetFiles() As Variant) As Boolean
    If IsMissing(targetFiles) Or UCase(IO.GetExtensionName(filePath)) <> "ZIP" Then
        CreateZipFile = False
        Exit Function
    End If
    If IO.FileExists(filePath) Then
        IO.DeleteFile filePath
        IO.CreateTextFile(filePath, False).Write "PK" & Chr(5) & Chr(6) & String(18, 0)
    Else
        IO.CreateTextFile(filePath, False).Write "PK" & Chr(5) & Chr(6) & String(18, 0)
    End If
    Dim shellObj As Object: Set shellObj = Base.GetShell
    Dim zipFolder As Object: Set zipFolder = shellObj.Namespace(IO.GetAbsolutePathName(filePath))
    Dim targetFilePath As Variant
    For Each targetFilePath In targetFiles
        Dim fileName As String: fileName = fso.GetFileName(targetFilePath)
        Dim sFolderItem As Object
        Set sFolderItem = shellObj.Namespace(IO.GetParentFolderName(IO.GetAbsolutePathName(targetFilePath))).ParseName(fileName)
        Dim skip As Boolean: skip = False
        If sFolderItem Is Nothing Then
            Debug.Print "Nothing: " & targetFilePath
            skip = True
        End If
        If IsFolder(targetFilePath) Then
            If shellObj.Namespace(IO.GetAbsolutePathName(targetFilePath)).items().Count > 0 Then
                ' 空フォルダをスキップ対象とする
                skip = True
            End If
        End If
        If Not skip Then
            Dim retryCnt0 As Long: retryCnt0 = 0
            Do
                Dim zFolderItem As Object: Set zFolderItem = zipFolder.ParseName(fileName)
                If zFolderItem Is Nothing Then
                    Dim cnt As Long: cnt = zipFolder.items().Count
                    zipFolder.CopyHere sFolderItem, &H4 Or &H10
                    Dim retryCnt As Long: retryCnt = 0
                    Do While zipFolder.items().Count <= cnt
                        If retryCnt > 10 Then
                            Debug.Print "RetryOver: " & targetFilePath
                            Exit Function
                        End If
                        Application.Wait Now + TimeSerial(0, 0, 1)
                        retryCnt = retryCnt + 1
                    Loop
                    Exit Do
                Else
                    ' ファイルがあれば置き換え
                    Call zFolderItem.InvokeVerb("delete")
                End If
            Loop
        End If
    Next
    CreateZipFile = True
End Function

'******************************************************************************
'* [概  要] UnZipFile
'* [詳  細] zipファイルを解凍する。
'* [補  足] Microsoft非推奨の方法であるため、Powershell経由でZIP解凍する方法
'*          をとるのが望ましい。
'*
'* @param zipFilePath zipファイルパス
'* @param targetFolderPath 解凍先フォルダパス
'*                         （解凍先フォルダにzipファイルのベース名と同じ名称
'*                           のフォルダを作成）
'* @return Boolean 実行結果
'******************************************************************************
Public Function UnZipFile(zipFilePath As String, targetFolderPath As String) As Boolean
    If Not FileExists(zipFilePath) Then
        UnZipFile = False
        Exit Function
    End If
    Dim saveFolderPath As String
    saveFolderPath = GetUniqueFolderName(IO.BuildPath(targetFolderPath, IO.GetBaseName(ChangeExtension(zipFilePath, ""))))
    Call CreateFolderTree(saveFolderPath)
    
    Dim shellObj As Object: Set shellObj = Base.GetShell
    Dim targetFolder As Object: Set targetFolder = shellObj.Namespace(IO.GetAbsolutePathName(saveFolderPath))
    Dim srcFolder As Object: Set srcFolder = shellObj.Namespace(IO.GetAbsolutePathName(zipFilePath))
    Dim cnt As Long: cnt = srcFolder.items().Count
    Dim retryCnt As Long: retryCnt = 0
    
    targetFolder.CopyHere srcFolder.items, &H4 Or &H10
    Do While targetFolder.items().Count < cnt
        If retryCnt > 10 Then
            Debug.Print "RetryOver"
            Exit Function
        End If
        Application.Wait Now + TimeSerial(0, 0, 1)
        retryCnt = retryCnt + 1
    Loop
    UnZipFile = True
End Function
