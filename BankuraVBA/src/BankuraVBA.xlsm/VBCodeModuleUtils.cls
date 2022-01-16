VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBCodeModuleUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] VBA Code/Module ユーティリティクラス
'* [詳  細] VBA Code/Moduleに関するユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'*
'* [備  考] オプションから[セキュリティセンターの設定]>[マクロの設定]を選択し、
'*          [VBAプロジェクトオブジェクトモデルへのアクセスを信頼する]をオンに
'*          設定すること。
'*
'*          「Microsoft Visual Basic for Applications Extensibility」を参照設定
'*          すると、編集時に便利。
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

'******************************************************************************
'* 定数定義
'******************************************************************************

'******************************************************************************
'* 内部変数定義
'******************************************************************************

'******************************************************************************
'* プロパティ定義
'******************************************************************************

'*-----------------------------------------------------------------------------
'* VBProjectsCount プロパティ（読み取り専用）
'*
'* [補  足]
'* ・現在開いているVBProjectの個数を返す。
'*-----------------------------------------------------------------------------
Property Get VBProjectsCount() As Long
    VBProjectsCount = Application.VBE.VBProjects.Count
End Property

'*-----------------------------------------------------------------------------
'* VBProjectNames プロパティ（読み取り専用）
'*
'* [補  足]
'* ・現在開いているVBProjectの名前リストを返す。
'*-----------------------------------------------------------------------------
Property Get VBProjectNames() As Variant
    Dim v, arrex As ArrayEx: Set arrex = New ArrayEx
    For Each v In Application.VBE.VBProjects
        arrex.Add v.Name
    Next
    VBProjectNames = arrex.ToArray
End Property

'*-----------------------------------------------------------------------------
'* ThisVBProjectName プロパティ（読み取り専用）
'*
'* [補  足]
'* ・現在使用しているVBProjectの名前を返す。
'*-----------------------------------------------------------------------------
Property Get ThisVBProjectName() As String
    ThisVBProjectName = ThisWorkbook.VBProject.Name
End Property

'*-----------------------------------------------------------------------------
'* ThisVBProjectFilename プロパティ（読み取り専用）
'*
'* [補  足]
'* ・現在使用しているVBProjectのファイル名を返す。
'*-----------------------------------------------------------------------------
Property Get ThisVBProjectFilename() As String
    ThisVBProjectFilename = ThisWorkbook.VBProject.FileName
End Property

'*-----------------------------------------------------------------------------
'* ThisVBProjectProtection プロパティ（読み取り専用）
'*
'* [補  足]
'* ・現在使用しているVBProjectの保護の設定状態を返す。
'*-----------------------------------------------------------------------------
Property Get ThisVBProjectProtection() As Long
    ThisVBProjectProtection = ThisWorkbook.VBProject.Protection
End Property

'*-----------------------------------------------------------------------------
'* ThisVBComponentsCount プロパティ（読み取り専用）
'*
'* [補  足]
'* ・現在使用しているVBProjectのコンポーネント個数のを返す。
'*-----------------------------------------------------------------------------
Property Get ThisVBComponentsCount() As Long
    ThisVBComponentsCount = ThisWorkbook.VBProject.VBComponents.Count
End Property


'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
End Sub
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] ExportComponent
'* [詳  細] コンポーネントをファイルにエクスポートする｡
'*
'* @param moduleName     エクスポートするモジュール名
'* @param exportFilePath エクスポート先のファイル名（パス）
'* @param book           対象のブック（未指定時はThisWorbook）
'*
'******************************************************************************
Public Sub ExportComponent(ByVal moduleName As String, ByVal exportFilePath As String, Optional book As Workbook)
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    
    mybook.VBProject.VBComponents(moduleName).Export exportFilePath
End Sub

'******************************************************************************
'* [概  要] ImportComponent
'* [詳  細] コンポーネントをインポートする｡
'*
'* @param importFilePath インポートするファイル名（パス）
'* @param book           対象のブック（未指定時はThisWorbook）
'*
'******************************************************************************
Public Sub ImportComponent(ByVal importFilePath As String, Optional book As Workbook)
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    
    mybook.VBProject.VBComponents.Import importFilePath
End Sub

'******************************************************************************
'* [概  要] AddClass
'* [詳  細] クラスを追加する｡
'*
'* @param className クラス名
'* @param book      対象のブック（未指定時はThisWorbook）
'* @return Object(Of VBComponent)
'*
'******************************************************************************
Public Function AddClass(ByVal className As String, Optional book As Workbook) As Object
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If

    Dim vbc As Object: Set vbc = mybook.VBProject.VBComponents.Add(2)
    vbc.Name = className
    
    Set AddClass = vbc
End Function

'******************************************************************************
'* [概  要] AddModule
'* [詳  細] モジュールを追加する｡
'*
'* @param moduleName モジュール名
'* @param book       対象のブック（未指定時はThisWorbook）
'* @return Object(Of VBComponent)
'*
'******************************************************************************
Public Function AddModule(ByVal moduleName As String, Optional book As Workbook) As Object
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If

    Dim vbc As Object: Set vbc = mybook.VBProject.VBComponents.Add(1)
    vbc.Name = moduleName
    
    Set AddModule = vbc
End Function

'******************************************************************************
'* [概  要] AddForm
'* [詳  細] フォームを追加する｡
'*
'* @param formName フォーム名
'* @param book     対象のブック（未指定時はThisWorbook）
'* @return Object(Of VBComponent)
'*
'******************************************************************************
Public Function AddForm(ByVal formName As String, Optional book As Workbook) As Object
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    
    Dim vbc As Object: Set vbc = mybook.VBProject.VBComponents.Add(3)
    vbc.Name = formName
        
    Set AddForm = vbc
End Function

'******************************************************************************
'* [概  要] Remove
'* [詳  細] 指定したコンポーネントを削除する｡
'*
'* @param componentName コンポーネント名
'* @param book          対象のブック（未指定時はThisWorbook）
'*
'******************************************************************************
Public Sub Remove(ByVal componentName As String, Optional book As Workbook)
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    With mybook.VBProject.VBComponents
        .Remove .Item(componentName)
    End With
End Sub


'******************************************************************************
'* [概  要] RenameComponent
'* [詳  細] 指定したコンポーネントの名称を変更する｡
'*
'* @param componentName コンポーネント名
'* @param newCpName     変更後の名称
'* @param book          対象のブック（未指定時はThisWorbook）
'*
'******************************************************************************
Public Sub RenameComponent(ByVal componentName As String, ByVal newCpName As String, Optional book As Workbook)
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    
    If VBComponentExists(componentName, mybook) Then
        mybook.VBProject.VBComponents(componentName).Name = newCpName
    End If
End Sub

'******************************************************************************
'* [概  要] CountOfLines
'* [詳  細] 指定したコンポーネントの行数を取得する｡
'*
'* @param componentName コンポーネント名
'* @param book           対象のブック（未指定時はThisWorbook）
'* @return Long 行数
'*
'******************************************************************************
Public Function CountOfLines(ByVal componentName As String, Optional book As Workbook) As Long
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    CountOfLines = mybook.VBProject.VBComponents(componentName).CodeModule.CountOfLines
End Function

'******************************************************************************
'* [概  要] GetComponentCodes
'* [詳  細] 指定したコンポーネントのコードを取得する｡
'*
'* @param componentName コンポーネント名
'* @param startline     取得開始行
'* @param count         行数
'* @param book          対象のブック（未指定時はThisWorbook）
'* @return String コード
'*
'******************************************************************************
Public Function GetComponentCodes(ByVal componentName As String, startline As Long, Count As Long, Optional book As Workbook) As String
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    GetComponentCodes = mybook.VBProject.VBComponents(componentName).CodeModule.lines(startline, Count)
End Function

'******************************************************************************
'* [概  要] GetComponentAllCodes
'* [詳  細] 指定したコンポーネントの全てのコードを取得する｡
'*
'* @param componentName コンポーネント名
'* @param book          対象のブック（未指定時はThisWorbook）
'* @return String コード
'*
'******************************************************************************
Public Function GetComponentAllCodes(ByVal componentName As String, Optional book As Workbook) As String
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    GetComponentAllCodes = GetComponentCodes(componentName, 1, CountOfLines(componentName, mybook), mybook)
End Function

'******************************************************************************
'* [概  要] CountOfEmptyLines
'* [詳  細] 指定したコンポーネントの空行数を取得する｡
'*
'* @param componentName コンポーネント名
'* @param book          対象のブック（未指定時はThisWorbook）
'* @return Long 空行数
'*
'******************************************************************************
Public Function CountOfEmptyLines(ByVal componentName As String, Optional book As Workbook) As Long
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If

    Dim v, tmp As String
    Dim emptyLineCount As Long
    For Each v In Split(GetComponentAllCodes(componentName, mybook), vbNewLine)
       tmp = Trim(v)
       If tmp = "" Then
           emptyLineCount = emptyLineCount + 1
       End If
    Next
    CountOfEmptyLines = emptyLineCount
End Function

'******************************************************************************
'* [概  要] CountOfCommentLines
'* [詳  細] 指定したコンポーネントのコメント行数を取得する｡
'*
'* @param componentName コンポーネント名
'* @param book          対象のブック（未指定時はThisWorbook）
'* @return Long コメント行数
'*
'******************************************************************************
Public Function CountOfCommentLines(ByVal componentName As String, Optional book As Workbook) As Long
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If

    Dim v, tmp As String
    Dim commentLineCount As Long
    For Each v In Split(GetComponentAllCodes(componentName, mybook), vbNewLine)
       tmp = Trim(v)
       If StringUtils.StartsWithAny(LCase(tmp), "'", "rem") Then
           commentLineCount = commentLineCount + 1
       End If
    Next
    CountOfCommentLines = commentLineCount
End Function

'******************************************************************************
'* [概  要] CountOfLogicalLines
'* [詳  細] 指定したコンポーネントの空行・コメント行数を除いた行数（LLOC）を
'*          取得する｡
'*
'* @param componentName コンポーネント名
'* @param book          対象のブック（未指定時はThisWorbook）
'* @return Long 空行数
'*
'******************************************************************************
Public Function CountOfLogicalLines(ByVal componentName As String, Optional book As Workbook) As Long
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    CountOfLogicalLines = CountOfLines(componentName, mybook) - CountOfEmptyLines(componentName, mybook) - CountOfCommentLines(componentName, mybook)
End Function


'******************************************************************************
'* [概  要] CountOfDeclarationLines
'* [詳  細] 指定したコンポーネントの宣言セクション部分の行数を取得する｡
'*
'* @param componentName コンポーネント名
'* @param book          対象のブック（未指定時はThisWorbook）
'* @return Long 宣言セクション部分の行数
'*
'******************************************************************************
Public Function CountOfDeclarationLines(ByVal componentName As String, Optional book As Workbook) As Long
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    CountOfDeclarationLines = mybook.VBProject.VBComponents(componentName).CodeModule.CountOfDeclarationLines
End Function

'******************************************************************************
'* [概  要] GetProcBodyLine
'* [詳  細] 指定したプロシージャの先頭行を取得する｡
'*          先頭行はプロシージャの宣言行を指す。
'*          Propertyの場合、同名の宣言があるため、以下の優先順で行数を返す。
'*           Let > Set > Get
'*
'* @param componentName コンポーネント名
'* @param book          対象のブック（未指定時はThisWorbook）
'* @return Long プロシージャの先頭行
'*
'******************************************************************************
Public Function GetProcBodyLine(ByVal componentName As String, ByVal procName As String, Optional book As Workbook) As Long
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    Dim i As Long
    On Error Resume Next
    For i = 0 To 3
        GetProcBodyLine = mybook.VBProject.VBComponents(componentName).CodeModule.ProcBodyLine(procName, ByVal i)
        If Err.Number = 0 Then Exit Function
        Err.Clear
    Next
End Function

'******************************************************************************
'* [概  要] GetProcCountLines
'* [詳  細] 指定したプロシージャの行数を取得する｡
'*          Propertyの場合、同名の宣言があるため、以下の優先順で行数を返す。
'*           Let > Set > Get
'*
'* @param componentName コンポーネント名
'* @param procName      プロシージャ名
'* @param book          対象のブック（未指定時はThisWorbook）
'* @return String プロシージャの行数
'*
'******************************************************************************
Public Function GetProcCountLines(ByVal componentName As String, ByVal procName As String, Optional book As Workbook) As Long
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    Dim i As Long
    On Error Resume Next
    For i = 0 To 3
        GetProcCountLines = mybook.VBProject.VBComponents(componentName).CodeModule.ProcCountLines(procName, ByVal i)
        If Err.Number = 0 Then Exit Function
        Err.Clear
    Next
End Function

'******************************************************************************
'* [概  要] GetProcNameOfLine
'* [詳  細] 指定した行を含むプロシージャの名前を取得する｡
'*
'* @param componentName コンポーネント名
'* @param procName      プロシージャ名
'* @param book          対象のブック（未指定時はThisWorbook）
'* @return String プロシージャの名前
'*
'******************************************************************************
Public Function GetProcNameOfLine(ByVal componentName As String, ByVal procLine As Long, Optional book As Workbook) As String
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    GetProcNameOfLine = mybook.VBProject.VBComponents(componentName).CodeModule.ProcOfLine(procLine, 0)
End Function

'******************************************************************************
'* [概  要] GetProcStartLine
'* [詳  細] 指定したプロシージャの開始行を取得する｡
'*          プロシージャの宣言行ではなく、直前のプロシージャの終了行の次行を
'*          取得するため、宣言行の前に空白行を開けていると、その行が開始行と
'*          なることに注意。
'*          Propertyの場合、同名の宣言があるため、以下の優先順で行数を返す。
'*           Let > Set > Get
'*
'* @param componentName コンポーネント名
'* @param procName      プロシージャ名
'* @param book          対象のブック（未指定時はThisWorbook）
'* @return Long 開始行
'*
'******************************************************************************
Public Function GetProcStartLine(ByVal componentName As String, ByVal procName As String, Optional book As Workbook) As Long
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If

    Dim i As Long
    On Error Resume Next
    For i = 0 To 3
        GetProcStartLine = mybook.VBProject.VBComponents(componentName).CodeModule.ProcStartLine(procName, ByVal i)
        If Err.Number = 0 Then Exit Function
        Err.Clear
    Next
End Function

'******************************************************************************
'* [概  要] GetProcCodes
'* [詳  細] 指定したプロシージャのコードを取得する｡
'*
'* @param componentName コンポーネント名
'* @param procName      プロシージャ名
'* @param book          対象のブック（未指定時はThisWorbook）
'* @return String コード
'*
'******************************************************************************
Public Function GetProcCodes(ByVal componentName As String, ByVal procName As String, Optional book As Workbook) As String
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    GetProcCodes = GetComponentCodes(componentName, _
                                     GetProcStartLine(componentName, procName, mybook), _
                                     GetProcCountLines(componentName, procName, mybook), mybook)
End Function

'******************************************************************************
'* [概  要] AddFromFile
'* [詳  細] 指定したファイルの内容をコードモジュールに挿入する｡
'*
'* @param componentName コンポーネント名
'* @param fileName      ファイル名
'* @param book          対象のブック（未指定時はThisWorbook）
'*
'******************************************************************************
Public Sub AddFromFile(ByVal componentName As String, ByVal FileName As String, Optional book As Workbook)
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    mybook.VBProject.VBComponents(componentName).CodeModule.AddFromFile FileName
End Sub

'******************************************************************************
'* [概  要] AddFromString
'* [詳  細] 指定した文字列をコードモジュールに挿入する｡
'*
'* @param componentName コンポーネント名
'* @param codeString    コード文字列
'* @param book          対象のブック（未指定時はThisWorbook）
'*
'******************************************************************************
Public Sub AddFromString(ByVal componentName As String, ByVal codeString As String, Optional book As Workbook)
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    mybook.VBProject.VBComponents(componentName).CodeModule.AddFromString codeString
End Sub

'******************************************************************************
'* [概  要] InsertLines
'* [詳  細] 指定した行に指定した文字列を挿入する｡
'*
'* @param componentName コンポーネント名
'* @param lineNo        行
'* @param codeString    コード文字列
'* @param book          対象のブック（未指定時はThisWorbook）
'*
'******************************************************************************
Public Sub InsertLines(ByVal componentName As String, ByVal lineNo As Long, ByVal codeString As String, Optional book As Workbook)
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    mybook.VBProject.VBComponents(componentName).CodeModule.InsertLines lineNo, codeString
End Sub

'******************************************************************************
'* [概  要] DeleteLines
'* [詳  細] 指定した行を削除する｡
'*
'* @param componentName コンポーネント名
'* @param lineNo        行
'* @param delCount      削除する行数
'* @param book          対象のブック（未指定時はThisWorbook）
'*
'******************************************************************************
Public Sub DeleteLines(ByVal componentName As String, ByVal lineNo As Long, Optional ByVal delCount As Long, Optional book As Workbook)
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    mybook.VBProject.VBComponents(componentName).CodeModule.DeleteLines lineNo, delCount
End Sub

'******************************************************************************
'* [概  要] ReplaceLine
'* [詳  細] 指定した行を指定した文字列で置換する｡
'*
'* @param componentName コンポーネント名
'* @param lineNo        行
'* @param codeString    コード文字列
'* @param book          対象のブック（未指定時はThisWorbook）
'*
'******************************************************************************
Public Sub ReplaceLine(ByVal componentName As String, ByVal lineNo As Long, ByVal codeString As String, Optional book As Workbook)
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    mybook.VBProject.VBComponents(componentName).CodeModule.ReplaceLine lineNo, codeString
End Sub

'******************************************************************************
'* [概  要] GetProcNames
'* [詳  細] 指定したコンポーネントのプロシージャ名のリストを取得する｡
'*
'* @param componentName コンポーネント名
'* @param book          対象のブック（未指定時はThisWorbook）
'* @return Variant プロシージャ名のリスト（1次元配列）
'*
'******************************************************************************
Public Function GetProcNames(ByVal componentName As String, Optional book As Workbook) As Variant
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If

    Dim arrex As ArrayEx: Set arrex = New ArrayEx
    Dim tmp As String, lineNo As Long
    With mybook.VBProject.VBComponents(componentName).CodeModule
        For lineNo = 1 To .CountOfLines
            If tmp <> .ProcOfLine(lineNo, 0) Then
                tmp = .ProcOfLine(lineNo, 0)
                arrex.Add tmp
            End If
        Next
    End With
    GetProcNames = arrex.ToArray
End Function

'******************************************************************************
'* [概  要] CountOfProc
'* [詳  細] 指定したコンポーネントのプロシージャ数を取得する｡
'*
'* @param componentName コンポーネント名
'* @param book          対象のブック（未指定時はThisWorbook）
'* @return Long プロシージャ数
'*
'******************************************************************************
Public Function CountOfProc(ByVal componentName As String, Optional book As Workbook) As Long
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If

    CountOfProc = ArrayUtils.GetLength(GetProcNames(componentName, mybook))
End Function

'******************************************************************************
'* [概  要] ExportComponents
'* [詳  細] 対象ブックに含まれるコンポーネントをフォルダにエクスポートする｡
'*
'* @param moduleName       エクスポートするモジュール名
'* @param exportFolderPath エクスポート先のフォルダ名（パス）
'* @param book             対象のブック（未指定時はThisWorbook）
'*
'******************************************************************************
Public Sub ExportComponents(ByVal exportFolderPath As String, Optional book As Workbook)
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    
    With mybook.VBProject
        Dim v
        For Each v In .VBComponents
            If v.Type = 1 Then
                v.Export IO.BuildPath(exportFolderPath, v.Name & ".bas")
            ElseIf v.Type = 2 Then
                v.Export IO.BuildPath(exportFolderPath, v.Name & ".cls")
            ElseIf v.Type = 3 Then
                v.Export IO.BuildPath(exportFolderPath, v.Name & ".frm")
            ElseIf v.Type = 11 Then
                v.Export IO.BuildPath(exportFolderPath, v.Name & ".frx")
            ElseIf v.Type = 100 Then
                v.Export IO.BuildPath(exportFolderPath, v.Name & ".dcm")
            End If
        Next
    End With
End Sub

'******************************************************************************
'* [概  要] ImportComponents
'* [詳  細] 指定フォルダ内のコンポーネントをインポートする｡
'*
'* @param importFolderPath インポート元のフォルダ名（パス）
'* @param book             対象のブック（未指定時はThisWorbook）
'*
'******************************************************************************
Public Sub ImportComponents(ByVal importFolderPath As String, Optional book As Workbook)
    If Not IO.FolderExists(importFolderPath) Then
        Exit Sub
    End If
    
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    
    Dim folder As Object: Set folder = IO.GetFolder(importFolderPath)
    Dim f As Object
    For Each f In folder.Files
        If IsVbModuleFile(f.Name) Then
            ' DELETE/INSERT
            mybook.VBProject.VBComponents.Remove mybook.VBProject.VBComponents(IO.GetBaseName(f.Path))
            mybook.VBProject.VBComponents.Import f.Path
        End If
    Next
End Sub

'******************************************************************************
'* [概  要] AddFromFileToBooks
'* [詳  細] 指定したファイルの内容を標準モジュールとしてフォルダ内の
'*          ブック（xlsm）に挿入する｡
'*
'* @param booksFolderPath エクスポート先のフォルダ名（パス）
'* @param moduleName      追加する標準モジュール名
'* @param insFilePath     ファイルパス
'*
'******************************************************************************
Public Sub AddFromFileToBooks(ByVal booksFolderPath As String, ByVal moduleName As String, ByVal insFilePath As String)
    If Not IO.FolderExists(booksFolderPath) Then
        Exit Sub
    End If
    If Not IO.FileExists(insFilePath) Then
        Exit Sub
    End If
    
    Dim displayAlertsState As Boolean: displayAlertsState = Application.DisplayAlerts
    If Application.DisplayAlerts Then Application.DisplayAlerts = False
    
    Dim folder As Object: Set folder = IO.GetFolder(booksFolderPath)
    Dim f As Object, mybook As Workbook
    For Each f In folder.Files
        If IO.GetExtensionName(f.Name) = "xlsm" Then
            Set mybook = Workbooks.Open(f.Path)
            
            AddModule(moduleName, mybook).AddFromFile insFilePath

            mybook.Save
            mybook.Close
        End If
    Next
    
    Application.DisplayAlerts = displayAlertsState
End Sub

'******************************************************************************
'* [概  要] GetComponentsInfo
'* [詳  細] 対象ブックに含まれる全コンポーネントの情報を取得する｡
'*          1コンポーネント毎に以下の項目を出力する。
'*          ＜項目＞
'*            0: コンポーネント名
'*            1: コンポーネントの種類
'*            2: プロシージャ数
'*            3: 全ステップ数
'*            4: 実ステップ数（コメント数、空行数を全行数から除いた行数）
'*            5: 空行数
'*            6: コメント行数
'*
'* @param book 対象のブック（未指定時はThisWorbook）
'* @return Array2DEx 全コンポーネントの情報
'*
'******************************************************************************
Public Function GetComponentsInfo(Optional book As Workbook) As Array2DEx
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    
    Dim arr2dex As Array2DEx: Set arr2dex = New Array2DEx
    With mybook.VBProject
        Dim row As ArrayEx
        Dim v As Object
        For Each v In .VBComponents
            Set row = New ArrayEx
            row.Add v.Name
            row.Add GetComponentTypeName(v.Type)
            row.Add CountOfProc(v.Name, mybook)
            row.Add CountOfLines(v.Name, mybook)
            row.Add CountOfLogicalLines(v.Name, mybook)
            row.Add CountOfEmptyLines(v.Name, mybook)
            row.Add CountOfCommentLines(v.Name, mybook)
            
            arr2dex.AddRow row
        Next
    End With
    Set GetComponentsInfo = arr2dex
End Function

'******************************************************************************
'* [概  要] IsVbModuleFile
'* [詳  細] 指定したファイル名がVBのコンポーネントのファイルに該当するか判定する｡
'*
'* @param targetFileName ファイル名
'* @return Boolean 判定結果（True：該当）
'*
'******************************************************************************
Public Function IsVbModuleFile(ByVal targetFileName As String) As Boolean
    Dim ext As String: ext = LCase(IO.GetExtensionName(targetFileName))
    
    IsVbModuleFile = ext = "bas" Or _
                     ext = "cls" Or _
                     ext = "frm" Or _
                     ext = "frx" Or _
                     ext = "dcm"
End Function

'******************************************************************************
'* [概  要] VBComponentExists
'* [詳  細] 指定したコンポーネントが存在するか判定する｡
'*
'* @param componentName コンポーネント名
'* @param book          対象のブック（未指定時はThisWorbook）
'* @return Boolean 判定結果（True：該当）
'*
'******************************************************************************
Public Function VBComponentExists(ByVal componentName As String, Optional book As Workbook) As Boolean
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    
    With mybook.VBProject
        Dim v
        For Each v In .VBComponents
            If v.Name = componentName Then
                VBComponentExists = True
                Exit Function
            End If
        Next

    End With
    VBComponentExists = False
End Function

'******************************************************************************
'* [概  要] GetProcKind
'* [詳  細] 指定したプロシージャの種類（番号）を取得する｡
'*
'* @param componentName コンポーネント名
'* @param procName      プロシージャ名
'* @param book          対象のブック（未指定時はThisWorbook）
'* @return String プロシージャの種類（番号）
'*
'******************************************************************************
Public Function GetProcKind(ByVal componentName As String, ByVal procName As String, Optional book As Workbook) As Long
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    Dim procKind As Long, tmp As String
    tmp = mybook.VBProject.VBComponents(componentName).CodeModule.ProcOfLine(GetProcBodyLine(componentName, procName, mybook), procKind)
    GetProcKind = procKind
End Function

'******************************************************************************
'* [概  要] GetProcKindName
'* [詳  細] 指定したプロシージャの種類（名称）を取得する｡
'*
'* @param componentName コンポーネント名
'* @param procName      プロシージャ名
'* @param book          対象のブック（未指定時はThisWorbook）
'* @return String プロシージャの種類（名称）
'*
'******************************************************************************
Public Function GetProcKindName(ByVal componentName As String, ByVal procName As String, Optional book As Workbook) As String
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    Dim procKind As Long

    GetProcKindName = GetProcKindNameByNum(GetProcKind(componentName, procName, mybook))
End Function

'******************************************************************************
'* [概  要] GetReferences
'* [詳  細] 参照設定のリストを取得する｡
'*
'* @param book     対象のブック（未指定時はThisWorbook）
'* @return Array2DEx 参照設定のリスト
'*
'******************************************************************************
Public Function GetReferences(Optional book As Workbook) As Array2DEx
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    
    Dim arr2dex As Array2DEx: Set arr2dex = New Array2DEx
    Dim row As ArrayEx
    Dim ref As Object
    For Each ref In mybook.VBProject.References
        Set row = New ArrayEx
        row.Add ref.Name
        row.Add ref.Description
        row.Add ref.FullPath
        arr2dex.AddRow row
    Next

    Set GetReferences = arr2dex
End Function

'******************************************************************************
'* [概  要] ClearImmediateWindow
'* [詳  細] イミディエイトウィンドウをクリアする｡
'*
'******************************************************************************
Public Sub ClearImmediateWindow()
    If Application.VBE.MainWindow.Visible And _
       Application.VBE.Windows("イミディエイト").Visible Then
        SendKeys "^g", True
        SendKeys "^a", True
        SendKeys "{Del}", True
        SendKeys "{F7}", True
        UXUtils.ChangeNumLockOn
        'UXUtils.WaitFor 1
    End If
End Sub

'******************************************************************************
'* [概  要] GetComponentTypeName
'* [詳  細] コンポーネントの種類を番号から判定し、取得する｡
'*
'* @param typeNum コンポーネントの種類（番号）
'* @return String コンポーネントの種類（名称）
'*
'******************************************************************************
Private Function GetComponentTypeName(ByVal typeNum As Long) As String
    Select Case typeNum
        Case 1
            GetComponentTypeName = "標準モジュール"
        Case 2
            GetComponentTypeName = "クラスモジュール"
        Case 3
            GetComponentTypeName = "MicrosoftForm"
        Case 11
            GetComponentTypeName = "ActiveXデザイナ"
        Case 100
            GetComponentTypeName = "Documentモジュール" '「ThisWorkbook」「Sheet1」等
        Case Else
            GetComponentTypeName = "不明"
    End Select
End Function

'******************************************************************************
'* [概  要] GetProcKindNameByNum
'* [詳  細] プロシージャの種類を番号から判定し、取得する｡
'*
'* @param procKindNum プロシージャの種類（番号）
'* @return String     プロシージャの種類（名称）
'*
'******************************************************************************
Private Function GetProcKindNameByNum(ByVal procKindNum As Long) As String
    Select Case procKindNum
        Case 0
            GetProcKindNameByNum = "Sub/Function"
        Case 1
            GetProcKindNameByNum = "Property Let"
        Case 2
            GetProcKindNameByNum = "Property Set"
        Case 3
            GetProcKindNameByNum = "Property Get"
    End Select
End Function
