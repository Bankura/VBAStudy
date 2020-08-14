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
'* [参  考]
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* Enum定義
'******************************************************************************

'******************************************************************************
'* 定数定義
'******************************************************************************

'******************************************************************************
'* 構造体定義
'******************************************************************************

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mWinApi As WinAPI

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
    FileExists = IO.fso.FileExists(targetFilePath)
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
'* [概  要] ReadCsv
'* [詳  細] CSVファイル読み込み関数（改行コードがCRLF）
'*
'* @param strCsvFilePath CSVファイルパス
'* @param objOutWorksheet データ貼付先シートオブジェクト
'* @param lngStartRow シート貼付開始行
'* @param lngStartCol シート貼付開始列
'* @return Long 読込行数
'*
'* [備  考] ■廃止予定（UTF-8に対応していないため）
'*
'******************************************************************************
Public Function ReadCsv(strCsvFilePath As String, objOutWorksheet As Worksheet, _
                        lngStartRow As Long, lngStartCol As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim buf As String, tmp As Variant, n As Long, i As Long
    n = 0
  
    Open strCsvFilePath For Input As #1
    Do Until EOF(1)
        Line Input #1, buf
        tmp = Split(buf, ",")
        objOutWorksheet.Cells(lngStartRow + n, lngStartCol).Resize(1, UBound(tmp) + lngStartCol - 1).value = tmp
        n = n + 1
    Loop

    Close #1
    ReadCsv = n
    Exit Function
    
ErrorHandler:
    ReadCsv = -1
End Function

'******************************************************************************
'* [概  要] ReadCsvLf
'* [詳  細] CSVファイル読み込み関数（改行コードがLF）
'*
'* @param strCsvFilePath CSVファイルパス
'* @param objOutWorksheet データ貼付先シートオブジェクト
'* @param lngStartRow シート貼付開始行
'* @param lngStartCol シート貼付開始列
'* @return Long 読込行数
'*
'* [備  考] ■廃止予定（UTF-8に対応していないため）
'*
'******************************************************************************
Public Function ReadCsvLf(strCsvFilePath As String, objOutWorksheet As Worksheet, _
                          lngStartRow As Long, lngStartCol As Long) As Long
                          
    On Error GoTo ErrorHandler
    Dim buf() As Byte
    Dim tmp As Variant, tmp2 As Variant, n As Long, i As Long
  
    Open strCsvFilePath For Binary As #1
        ReDim buf(1 To LOF(1))
        Get #1, , buf
    Close #1
  
    tmp = Split(StrConv(buf, vbUnicode), vbLf)
    For n = 0 To UBound(tmp)
        If tmp(n) = "" Then
            n = n - 1
            Exit For
        End If
        tmp2 = Split(tmp(n), ",")
        objOutWorksheet.Cells(lngStartRow + n, lngStartCol).Resize(1, UBound(tmp2) + lngStartCol - 1).value = tmp2
    Next n
  
    ReadCsvLf = n + 1

    Exit Function
    
ErrorHandler:
    ReadCsvLf = -1
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
    ByVal strm As Object, ByVal fpath As String, ByVal opSave As SaveOptionsEnum _
    )
    Call IO.SaveToFileWithoutBom(strm, fpath, opSave)
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
    ByVal fpath As String, ByVal chrset As String, ByVal linsep As LineSeparatorsEnum _
    )
    Call IO.RemoveBom(fpath, chrset, linsep)
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
    ByVal fpath As String, ByVal crrChrset As String, ByVal chgChrset As String _
    )
    Call IO.ChangeFileCharset(fpath, crrChrset, chgChrset)
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
    ByVal fpath As String, ByVal chrset As String, _
    ByVal crrLinsep As LineSeparatorsEnum, ByVal chgLinsep As LineSeparatorsEnum _
    )
    Call IO.ChangeFileLineSeparator(fpath, chrset, crrLinsep, chgLinsep)
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
Public Function IsPathRooted(ByVal fpath As String) As Boolean
    IsPathRooted = IO.IsPathRooted(fpath)
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
    ByVal fpath As String, Optional delim As String = "_" _
    ) As String
    GetUniqueFileName = IO.GetUniqueFileName(fpath, delim)
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
    Call SetFileTimeProc(stFilePath, dtUpdateTime, False, False, True)
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
    Call SetFileTimeProc(stFilePath, dtAccessTime, False, True, False)
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
    Dim fso, oFolder, oFiles, oFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFolder = fso.GetFolder(targetFolderName)
    Set oFiles = oFolder.Files
    
    For Each oFile In oFiles
        Dim filePath As String
        filePath = fso.GetAbsolutePathName(AddPathSeparator(targetFolderName) & oFile.Name)
         
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
    adoStream1.CopyTo adoStream2
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
    With streamObj
        .Position = 0
        .Type_ = adTypeBinary
        .Position = IO.BomSize(streamObj.charset)  'BOM飛ばし
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

    Dim varr(), lineCount As Long, cnt As Long
    Dim stream As Object: Set stream = IO.CreateADODBStream(adTypeText, cSet, separator)
    With stream
        .Mode = adModeReadWrite
        .Open
        .LoadFromFile filePath 'ファイル指定
        lineCount = CountLine(stream)
        ReDim Preserve varr(lineCount - 1)
        .Position = 0
        Do While Not .EOS
            varr(cnt) = .ReadText(adReadLine)
            cnt = cnt + 1
        Loop
        .Close
    End With

    ReadTextFileToVArray = varr
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
            .AppendMode stream, filePath
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
            .AppendMode stream, filePath
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
'* [概  要] BOM有無チェック処理。
'* [詳  細] UTF-8テキストファイルにBOMがあるかチェックを行う。
'*
'* @param filePath 対象となるテキストファイルのフルパス
'* @param チェック結果（true:BOM有、false:BOM無）
'******************************************************************************
Function IsUTF8BomIncluded(filePath As String) As Boolean
    Dim buf() As Byte
    Dim suspected As String
    Dim v As Variant

    Dim stream As Object: Set stream = IO.CreateADODBStream(adTypeBinary)
    With stream
        .Open
        .LoadFromFile filePath
        buf = .Read(3) 'BOMにあたる先頭3バイトを取得
        .Close
    End With

    For Each v In buf
        suspected = suspected & Hex(v)
    Next
    
    If suspected = "EFBBBF" Then
        IsUTF8BomIncluded = True
    Else
        IsUTF8BomIncluded = False
    End If
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
        .Type_ = adTypeBinary
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
