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
'* [�@�\��] VBA Code/Module ���[�e�B���e�B�N���X
'* [��  ��] VBA Code/Module�Ɋւ��郆�[�e�B���e�B��񋟂���B
'*          Static�N���X�iAttribute VB_PredeclaredId = True�j�Ƃ���B
'*
'* [��  �l] �I�v�V��������[�Z�L�����e�B�Z���^�[�̐ݒ�]>[�}�N���̐ݒ�]��I�����A
'*          [VBA�v���W�F�N�g�I�u�W�F�N�g���f���ւ̃A�N�Z�X��M������]���I����
'*          �ݒ肷�邱�ƁB
'*
'*          �uMicrosoft Visual Basic for Applications Extensibility�v���Q�Ɛݒ�
'*          ����ƁA�ҏW���ɕ֗��B
'*
'* @author Bankura
'* Copyright (c) 2021 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI��`
'******************************************************************************
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'******************************************************************************
'* Enum��`
'******************************************************************************

'******************************************************************************
'* �萔��`
'******************************************************************************

'******************************************************************************
'* �����ϐ���`
'******************************************************************************

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************

'*-----------------------------------------------------------------------------
'* VBProjectsCount �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �E���݊J���Ă���VBProject�̌���Ԃ��B
'*-----------------------------------------------------------------------------
Property Get VBProjectsCount() As Long
    VBProjectsCount = Application.VBE.VBProjects.Count
End Property

'*-----------------------------------------------------------------------------
'* VBProjectNames �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �E���݊J���Ă���VBProject�̖��O���X�g��Ԃ��B
'*-----------------------------------------------------------------------------
Property Get VBProjectNames() As Variant
    Dim v, arrex As ArrayEx: Set arrex = New ArrayEx
    For Each v In Application.VBE.VBProjects
        arrex.Add v.Name
    Next
    VBProjectNames = arrex.ToArray
End Property

'*-----------------------------------------------------------------------------
'* ThisVBProjectName �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �E���ݎg�p���Ă���VBProject�̖��O��Ԃ��B
'*-----------------------------------------------------------------------------
Property Get ThisVBProjectName() As String
    ThisVBProjectName = ThisWorkbook.VBProject.Name
End Property

'*-----------------------------------------------------------------------------
'* ThisVBProjectFilename �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �E���ݎg�p���Ă���VBProject�̃t�@�C������Ԃ��B
'*-----------------------------------------------------------------------------
Property Get ThisVBProjectFilename() As String
    ThisVBProjectFilename = ThisWorkbook.VBProject.fileName
End Property

'*-----------------------------------------------------------------------------
'* ThisVBProjectProtection �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �E���ݎg�p���Ă���VBProject�̕ی�̐ݒ��Ԃ�Ԃ��B
'*-----------------------------------------------------------------------------
Property Get ThisVBProjectProtection() As Long
    ThisVBProjectProtection = ThisWorkbook.VBProject.Protection
End Property

'*-----------------------------------------------------------------------------
'* ThisVBComponentsCount �v���p�e�B�i�ǂݎ���p�j
'*
'* [��  ��]
'* �E���ݎg�p���Ă���VBProject�̃R���|�[�l���g���̂�Ԃ��B
'*-----------------------------------------------------------------------------
Property Get ThisVBComponentsCount() As Long
    ThisVBComponentsCount = ThisWorkbook.VBProject.VBComponents.Count
End Property


'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
End Sub
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] ExportComponent
'* [��  ��] �R���|�[�l���g���t�@�C���ɃG�N�X�|�[�g����
'*
'* @param moduleName     �G�N�X�|�[�g���郂�W���[����
'* @param exportFilePath �G�N�X�|�[�g��̃t�@�C�����i�p�X�j
'* @param book           �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
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
'* [�T  �v] ImportComponent
'* [��  ��] �R���|�[�l���g���C���|�[�g����
'*
'* @param importFilePath �C���|�[�g����t�@�C�����i�p�X�j
'* @param book           �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
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
'* [�T  �v] AddClass
'* [��  ��] �N���X��ǉ�����
'*
'* @param className �N���X��
'* @param book      �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
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
'* [�T  �v] AddModule
'* [��  ��] ���W���[����ǉ�����
'*
'* @param moduleName ���W���[����
'* @param book       �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
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
'* [�T  �v] AddForm
'* [��  ��] �t�H�[����ǉ�����
'*
'* @param formName �t�H�[����
'* @param book     �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
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
'* [�T  �v] Remove
'* [��  ��] �w�肵���R���|�[�l���g���폜����
'*
'* @param componentName �R���|�[�l���g��
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
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
'* [�T  �v] RenameComponent
'* [��  ��] �w�肵���R���|�[�l���g�̖��̂�ύX����
'*
'* @param componentName �R���|�[�l���g��
'* @param newCpName     �ύX��̖���
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
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
'* [�T  �v] CountOfLines
'* [��  ��] �w�肵���R���|�[�l���g�̍s�����擾����
'*
'* @param componentName �R���|�[�l���g��
'* @param book           �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
'* @return Long �s��
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
'* [�T  �v] GetComponentCodes
'* [��  ��] �w�肵���R���|�[�l���g�̃R�[�h���擾����
'*
'* @param componentName �R���|�[�l���g��
'* @param startline     �擾�J�n�s
'* @param count         �s��
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
'* @return String �R�[�h
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
'* [�T  �v] GetComponentAllCodes
'* [��  ��] �w�肵���R���|�[�l���g�̑S�ẴR�[�h���擾����
'*
'* @param componentName �R���|�[�l���g��
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
'* @return String �R�[�h
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
'* [�T  �v] CountOfEmptyLines
'* [��  ��] �w�肵���R���|�[�l���g�̋�s�����擾����
'*
'* @param componentName �R���|�[�l���g��
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
'* @return Long ��s��
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
'* [�T  �v] CountOfCommentLines
'* [��  ��] �w�肵���R���|�[�l���g�̃R�����g�s�����擾����
'*
'* @param componentName �R���|�[�l���g��
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
'* @return Long �R�����g�s��
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
'* [�T  �v] CountOfLogicalLines
'* [��  ��] �w�肵���R���|�[�l���g�̋�s�E�R�����g�s�����������s���iLLOC�j��
'*          �擾����
'*
'* @param componentName �R���|�[�l���g��
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
'* @return Long ��s��
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
'* [�T  �v] CountOfDeclarationLines
'* [��  ��] �w�肵���R���|�[�l���g�̐錾�Z�N�V���������̍s�����擾����
'*
'* @param componentName �R���|�[�l���g��
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
'* @return Long �錾�Z�N�V���������̍s��
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
'* [�T  �v] GetProcBodyLine
'* [��  ��] �w�肵���v���V�[�W���̐擪�s���擾����
'*          �擪�s�̓v���V�[�W���̐錾�s���w���B
'*          Property�̏ꍇ�A�����̐錾�����邽�߁A�ȉ��̗D�揇�ōs����Ԃ��B
'*           Let > Set > Get
'*
'* @param componentName �R���|�[�l���g��
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
'* @return Long �v���V�[�W���̐擪�s
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
'* [�T  �v] GetProcCountLines
'* [��  ��] �w�肵���v���V�[�W���̍s�����擾����
'*          Property�̏ꍇ�A�����̐錾�����邽�߁A�ȉ��̗D�揇�ōs����Ԃ��B
'*           Let > Set > Get
'*
'* @param componentName �R���|�[�l���g��
'* @param procName      �v���V�[�W����
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
'* @return String �v���V�[�W���̍s��
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
'* [�T  �v] GetProcNameOfLine
'* [��  ��] �w�肵���s���܂ރv���V�[�W���̖��O���擾����
'*
'* @param componentName �R���|�[�l���g��
'* @param procName      �v���V�[�W����
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
'* @return String �v���V�[�W���̖��O
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
'* [�T  �v] GetProcStartLine
'* [��  ��] �w�肵���v���V�[�W���̊J�n�s���擾����
'*          �v���V�[�W���̐錾�s�ł͂Ȃ��A���O�̃v���V�[�W���̏I���s�̎��s��
'*          �擾���邽�߁A�錾�s�̑O�ɋ󔒍s���J���Ă���ƁA���̍s���J�n�s��
'*          �Ȃ邱�Ƃɒ��ӁB
'*          Property�̏ꍇ�A�����̐錾�����邽�߁A�ȉ��̗D�揇�ōs����Ԃ��B
'*           Let > Set > Get
'*
'* @param componentName �R���|�[�l���g��
'* @param procName      �v���V�[�W����
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
'* @return Long �J�n�s
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
'* [�T  �v] GetProcCodes
'* [��  ��] �w�肵���v���V�[�W���̃R�[�h���擾����
'*
'* @param componentName �R���|�[�l���g��
'* @param procName      �v���V�[�W����
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
'* @return String �R�[�h
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
'* [�T  �v] AddFromFile
'* [��  ��] �w�肵���t�@�C���̓��e���R�[�h���W���[���ɑ}������
'*
'* @param componentName �R���|�[�l���g��
'* @param fileName      �t�@�C����
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
'*
'******************************************************************************
Public Sub AddFromFile(ByVal componentName As String, ByVal fileName As String, Optional book As Workbook)
    Dim mybook As Workbook
    If book Is Nothing Then
        Set mybook = ThisWorkbook
    Else
        Set mybook = book
    End If
    mybook.VBProject.VBComponents(componentName).CodeModule.AddFromFile fileName
End Sub

'******************************************************************************
'* [�T  �v] AddFromString
'* [��  ��] �w�肵����������R�[�h���W���[���ɑ}������
'*
'* @param componentName �R���|�[�l���g��
'* @param codeString    �R�[�h������
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
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
'* [�T  �v] InsertLines
'* [��  ��] �w�肵���s�Ɏw�肵���������}������
'*
'* @param componentName �R���|�[�l���g��
'* @param lineNo        �s
'* @param codeString    �R�[�h������
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
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
'* [�T  �v] DeleteLines
'* [��  ��] �w�肵���s���폜����
'*
'* @param componentName �R���|�[�l���g��
'* @param lineNo        �s
'* @param delCount      �폜����s��
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
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
'* [�T  �v] ReplaceLine
'* [��  ��] �w�肵���s���w�肵��������Œu������
'*
'* @param componentName �R���|�[�l���g��
'* @param lineNo        �s
'* @param codeString    �R�[�h������
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
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
'* [�T  �v] GetProcNames
'* [��  ��] �w�肵���R���|�[�l���g�̃v���V�[�W�����̃��X�g���擾����
'*
'* @param componentName �R���|�[�l���g��
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
'* @return Variant �v���V�[�W�����̃��X�g�i1�����z��j
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
'* [�T  �v] CountOfProc
'* [��  ��] �w�肵���R���|�[�l���g�̃v���V�[�W�������擾����
'*
'* @param componentName �R���|�[�l���g��
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
'* @return Long �v���V�[�W����
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
'* [�T  �v] ExportComponents
'* [��  ��] �Ώۃu�b�N�Ɋ܂܂��R���|�[�l���g���t�H���_�ɃG�N�X�|�[�g����
'*
'* @param moduleName       �G�N�X�|�[�g���郂�W���[����
'* @param exportFolderPath �G�N�X�|�[�g��̃t�H���_���i�p�X�j
'* @param book             �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
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
'* [�T  �v] ImportComponents
'* [��  ��] �w��t�H���_���̃R���|�[�l���g���C���|�[�g����
'*
'* @param importFolderPath �C���|�[�g���̃t�H���_���i�p�X�j
'* @param book             �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
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
'* [�T  �v] AddFromFileToBooks
'* [��  ��] �w�肵���t�@�C���̓��e��W�����W���[���Ƃ��ăt�H���_����
'*          �u�b�N�ixlsm�j�ɑ}������
'*
'* @param booksFolderPath �G�N�X�|�[�g��̃t�H���_���i�p�X�j
'* @param moduleName      �ǉ�����W�����W���[����
'* @param insFilePath     �t�@�C���p�X
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
'* [�T  �v] GetComponentsInfo
'* [��  ��] �Ώۃu�b�N�Ɋ܂܂��S�R���|�[�l���g�̏����擾����
'*          1�R���|�[�l���g���Ɉȉ��̍��ڂ��o�͂���B
'*          �����ځ�
'*            0: �R���|�[�l���g��
'*            1: �R���|�[�l���g�̎��
'*            2: �v���V�[�W����
'*            3: �S�X�e�b�v��
'*            4: ���X�e�b�v���i�R�����g���A��s����S�s�����珜�����s���j
'*            5: ��s��
'*            6: �R�����g�s��
'*
'* @param book �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
'* @return Array2DEx �S�R���|�[�l���g�̏��
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
'* [�T  �v] IsVbModuleFile
'* [��  ��] �w�肵���t�@�C������VB�̃R���|�[�l���g�̃t�@�C���ɊY�����邩���肷��
'*
'* @param targetFileName �t�@�C����
'* @return Boolean ���茋�ʁiTrue�F�Y���j
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
'* [�T  �v] VBComponentExists
'* [��  ��] �w�肵���R���|�[�l���g�����݂��邩���肷��
'*
'* @param componentName �R���|�[�l���g��
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
'* @return Boolean ���茋�ʁiTrue�F�Y���j
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
'* [�T  �v] GetProcKind
'* [��  ��] �w�肵���v���V�[�W���̎�ށi�ԍ��j���擾����
'*
'* @param componentName �R���|�[�l���g��
'* @param procName      �v���V�[�W����
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
'* @return String �v���V�[�W���̎�ށi�ԍ��j
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
'* [�T  �v] GetProcKindName
'* [��  ��] �w�肵���v���V�[�W���̎�ށi���́j���擾����
'*
'* @param componentName �R���|�[�l���g��
'* @param procName      �v���V�[�W����
'* @param book          �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
'* @return String �v���V�[�W���̎�ށi���́j
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
'* [�T  �v] GetReferences
'* [��  ��] �Q�Ɛݒ�̃��X�g���擾����
'*
'* @param book     �Ώۂ̃u�b�N�i���w�莞��ThisWorbook�j
'* @return Array2DEx �Q�Ɛݒ�̃��X�g
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
'* [�T  �v] ClearImmediateWindow
'* [��  ��] �C�~�f�B�G�C�g�E�B���h�E���N���A����
'*
'******************************************************************************
Public Sub ClearImmediateWindow()
    If Application.VBE.MainWindow.Visible And _
       Application.VBE.Windows("�C�~�f�B�G�C�g").Visible Then
        SendKeys "^g", True
        SendKeys "^a", True
        SendKeys "{Del}", True
        SendKeys "{F7}", True
        UXUtils.ChangeNumLockOn
        'UXUtils.WaitFor 1
    End If
End Sub

'******************************************************************************
'* [�T  �v] GetComponentTypeName
'* [��  ��] �R���|�[�l���g�̎�ނ�ԍ����画�肵�A�擾����
'*
'* @param typeNum �R���|�[�l���g�̎�ށi�ԍ��j
'* @return String �R���|�[�l���g�̎�ށi���́j
'*
'******************************************************************************
Private Function GetComponentTypeName(ByVal typeNum As Long) As String
    Select Case typeNum
        Case 1
            GetComponentTypeName = "�W�����W���[��"
        Case 2
            GetComponentTypeName = "�N���X���W���[��"
        Case 3
            GetComponentTypeName = "MicrosoftForm"
        Case 11
            GetComponentTypeName = "ActiveX�f�U�C�i"
        Case 100
            GetComponentTypeName = "Document���W���[��" '�uThisWorkbook�v�uSheet1�v��
        Case Else
            GetComponentTypeName = "�s��"
    End Select
End Function

'******************************************************************************
'* [�T  �v] GetProcKindNameByNum
'* [��  ��] �v���V�[�W���̎�ނ�ԍ����画�肵�A�擾����
'*
'* @param procKindNum �v���V�[�W���̎�ށi�ԍ��j
'* @return String     �v���V�[�W���̎�ށi���́j
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