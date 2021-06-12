VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PsqlCommander"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] PSQL�R�}���h���s�N���X
'* [��  ��] PostgeSQL�̃R�}���h���C���c�[��PSQL�̃R�}���h���s�Ɋւ��鏈��
'*          ��񋟂���B
'*
'* [�Q  �l]�@https://kenpg.bitbucket.io/blog/201507/24.html
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI��`
'******************************************************************************

'******************************************************************************
'* Enum��`
'******************************************************************************

'******************************************************************************
'* �萔��`
'******************************************************************************

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
Private mCommander As Commander
Private mPsqlPath As String
Private mDbHost As String
Private mDbPort As String
Private mDbName As String
Private mDbUserName As String
Private mDbPassword As String
Private mClEncode As String


'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************

'*-----------------------------------------------------------------------------
'* PsqlPath �v���p�e�B
'*
'* psql�R�}���h�̃t���p�X
'*-----------------------------------------------------------------------------
Public Property Get PsqlPath() As String
    PsqlPath = mPsqlPath
End Property
Public Property Let PsqlPath(arg As String)
    mPsqlPath = arg
End Property

'*-----------------------------------------------------------------------------
'* DbHost �v���p�e�B
'*
'* �A�N�Z�X����DB�̃z�X�g�i��Flocalhost�j
'*-----------------------------------------------------------------------------
Public Property Get DbHost() As String
    DbHost = mDbHost
End Property
Public Property Let DbHost(arg As String)
    mDbHost = arg
End Property

'*-----------------------------------------------------------------------------
'* DbPort �v���p�e�B
'*
'* �A�N�Z�X����DB�̃|�[�g�ԍ�
'*-----------------------------------------------------------------------------
Public Property Get DbPort() As String
    DbPort = mDbPort
End Property
Public Property Let DbPort(arg As String)
    mDbPort = arg
End Property

'*-----------------------------------------------------------------------------
'* DbName �v���p�e�B
'*
'* �A�N�Z�X����DB��
'*-----------------------------------------------------------------------------
Public Property Get dbName() As String
    dbName = mDbName
End Property
Public Property Let dbName(arg As String)
    mDbName = arg
End Property

'*-----------------------------------------------------------------------------
'* DbUserName �v���p�e�B
'*
'* DB���[�U��
'*-----------------------------------------------------------------------------
Public Property Get DbUserName() As String
    DbUserName = mDbUserName
End Property
Public Property Let DbUserName(arg As String)
    mDbUserName = arg
End Property

'*-----------------------------------------------------------------------------
'* DbPassword �v���p�e�B
'*
'* DB�p�X���[�h
'*-----------------------------------------------------------------------------
Public Property Get DbPassword() As String
    DbPassword = mDbPassword
End Property
Public Property Let DbPassword(arg As String)
    mDbPassword = arg
End Property

'*-----------------------------------------------------------------------------
'* ClEncode �v���p�e�B
'*
'* Client�G���R�[�h
'*-----------------------------------------------------------------------------
Public Property Get clEncode() As String
    clEncode = mClEncode
End Property
Public Property Let clEncode(arg As String)
    mClEncode = arg
End Property


'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    Set mCommander = New Commander
    mClEncode = "SJIS"
End Sub

'******************************************************************************
'* ���\�b�h��`
'******************************************************************************
'******************************************************************************
'* [�T  �v] GetCommandResultAsTextStream
'* [��  ��] �R�}���h�����s���A���s�������ʂ��ꎞ�t�@�C���ɕۑ�����B
'*          �ꎞ�t�@�C��������s���ʂ�TextStream�Ŏ擾����B
'*
'* @param commandString �R�}���h������
'* @param myTempPath �ꎞ�t�@�C���p�X�i�W���̃p�X�ȊO�ɂ������ꍇ�Ɏw��j
'* @return Object(Of Scripting.TextStream)
'******************************************************************************
Public Function GetCommandResultAsTextStream(commandString, Optional myTempPath) As Object
    Call SetPsqlBaseCommand
    Set GetCommandResultAsTextStream = mCommander.GetCommandResultAsTextStream(commandString, myTempPath)
End Function

'******************************************************************************
'* [�T  �v] GetCommandResult
'* [��  ��] �R�}���h�����s���A���s�������ʕ�������擾����B
'*
'* @param commandString �R�}���h������
'* @return String �R�}���h���s����
'******************************************************************************
Public Function GetCommandResult(commandString) As String
    Call SetPsqlBaseCommand
    GetCommandResult = mCommander.GetCommandResult(commandString)
End Function

'******************************************************************************
'* [�T  �v] GetCommandResultAsArray
'* [��  ��] �R�}���h�����s���A���s�������ʂ��s���ɕ������������z��Ŏ擾����B
'*
'* @param commandString �R�}���h������
'* @return String() �R�}���h���s����
'******************************************************************************
Public Function GetCommandResultAsArray(commandString) As String()
    Call SetPsqlBaseCommand
    GetCommandResultAsArray = mCommander.GetCommandResultAsArray(commandString)
End Function

'******************************************************************************
'* [�T  �v] SetPsqlBaseCommand
'* [��  ��] Psql�̃R�}���h�E���ϐ���ݒ肷��B
'*          ���O�Ɋe�v���p�e�B��ݒ肵�Ă����K�v������B
'*
'******************************************************************************
Private Sub SetPsqlBaseCommand()
    mCommander.BaseCommand = mPsqlPath & " -h " & mDbHost & " -p " & mDbPort & _
                                         " -d " & mDbName & " -U " & mDbUserName & _
                                         " -c ""{0}"" -A " & "-o {1}"
    With Core.Wsh
        ' �p�X���[�h�����ϐ��ɐݒ�
        .Environment("Process").Item("PGPASSWORD") = mDbPassword
        ' Client(psql)�̃G���R�[�h�ݒ�F�f�t�H���g�uSJIS�v
        .Environment("Process").Item("PGCLIENTENCODING") = mClEncode
    End With
End Sub

'******************************************************************************
'* [�T  �v] psql�R�}���h�����SQL�����s����B
'* [��  ��] psql�R�}���h����āASQL�����s���A���s���ʂ�ԋp����B
'*          ���[�J������psql���C���X�g�[������Ă���K�v������B
'*          psql��postgreSQL�C���X�g�[�����ɍ��킹�ăC���X�g�[�������B
'*          ���s���ʂ͈ꎞ�t�@�C���Ɉ�x�o�͌�A�ǂݏo�����s�����߁A
'*          ��ʃf�[�^�ɂ��Ή��\�i����ꊇ�Ǐo���̂��߁AGiB���x������
'*          �Ή��͓����������Ȃ��j�B
'*
'* @param sql SQL��
'* @return ���s���ʃf�[�^�i2�����z��j
'*
'******************************************************************************
Public Function ExecPsql(sql As String) As Variant
    Dim resultData: resultData = GetCommandResultAsArray(sql)
    ExecPsql = EditQueryResult(resultData)
End Function

'******************************************************************************
'* [�T  �v] psql�R�}���h�����SQL�����s����B
'* [��  ��] psql�R�}���h����āASQL�����s���A���s���ʂ�ԋp����B
'*          ���[�J������psql���C���X�g�[������Ă���K�v������B
'*          psql��postgreSQL�C���X�g�[�����ɍ��킹�ăC���X�g�[�������B
'*
'* [��  ��] WScript.Shell��Exec���\�b�h�̕W���o�͋y�уG���[�o�͂�
'*          �o�b�t�@�T�C�Y��4KB�̂��߁A��ʂɃf�[�^���o�͂����ꍇ�o�b�t�@��
'*          ��t�ɂȂ菈�����t���[�Y����B
'*          ���̂��߁AExec�I����҂����ɁA�W���o�͂���s���̓ǂݏo�����s�����Ƃ�
'*          �o�b�t�@�������ς��ɂȂ�Ȃ��悤�ɂ��Ă��邪�A
'*          �W���o�͂�1�s���̃f�[�^�ʂ�4KB�𒴂���ꍇ�͑Ή��ł����t���[�Y����B
'*
'* @param sql SQL��
'* @return CommandResult���
'*
'******************************************************************************
Public Function ExecPsqlToCommandResult(sql As String) As CommandResult
    Dim cmd As String, oExec
    
    ' �R�}���h�g����
    cmd = mPsqlPath & " -h " & mDbHost & " -p " & mDbPort & " -d " & mDbName & " -U " & mDbUserName & " -c """ & sql & """ -A"
    Debug.Print cmd
    
    ' Wscript.Shell�I�u�W�F�N�g����
    With Core.Wsh
        ' �p�X���[�h�����ϐ��ɐݒ�
        .Environment("Process").Item("PGPASSWORD") = mDbPassword
    
        ' Client(psql)�̃G���R�[�h�ݒ�F�f�t�H���g�uSJIS�v
        .Environment("Process").Item("PGCLIENTENCODING") = mClEncode
    End With

    ' �R�}���h���s�E���ʐݒ�
    Dim res As CommandResult: Set res = mCommander.Exec(cmd)
    
    ' �R�}���h���s������
    If res.ExitCd <> 0 Then
        Set ExecPsqlToCommandResult = res
        Exit Function
    End If

    If ArrayUtils.IsEmptyArray(res.StdOutList) Then
        Set ExecPsqlToCommandResult = res
        Exit Function
    End If
    
    ' ���s���ʕҏW
    Dim resultEditData: resultEditData = EditQueryResult(res.StdOutList)

    ' �ԋp�l�ݒ�
    res.StdOutList = resultEditData
    res.RowMax = ArrayUtils.GetLength(resultEditData)
    res.ColMax = ArrayUtils.GetLength(resultEditData, 2)
    Set ExecPsqlToCommandResult = res
End Function

'******************************************************************************
'* [�T  �v] SQL���s���ʕҏW�����B
'* [��  ��] psql�R�}���h�����SQL���s���ʂ�2�����z��ɕҏW����B
'*
'* @param resultData ���s���ʃf�[�^�i1�����z��j
'* @return �ҏW�ςݎ��s���ʃf�[�^�i2�����z��j
'*
'******************************************************************************
Private Function EditQueryResult(resultData)
    ' ���펞����
    Dim lRowMax As Long: lRowMax = 0
    Dim lColMax As Long: lColMax = 0
    
    If ArrayUtils.IsEmptyArray(resultData) Then
        EditQueryResult = ArrayUtils.EmptyVariant2DArray
        Exit Function
    End If
    lRowMax = ArrayUtils.GetLength(resultData)
    
    ' �񐔌v��
    Dim vTmpCols As Variant: vTmpCols = Split(resultData(0), "|")
    lColMax = UBound(vTmpCols) - LBound(vTmpCols) + 1

    '�f�[�^�ݒ�
    Dim vArray()
    ReDim vArray(0 To lRowMax - 1, 0 To lColMax - 1)
    
    Dim i As Long, cols As Variant
    Dim row
    For Each row In resultData
        'Debug.Print "Line: " & row
        
        ' SQL���s���ʂɗ��؂�(|)�Ȃ��̃p�^�[��
        If StringUtils.ContainsNone(row, "|") Then
            ' 1��ڈȊO�͋󖄂�
            Call ArrayUtils.SetDefaultTo2DArrayRow(vArray, i, "")
            vArray(i, 0) = row
        ' ���؂�(|)����
        Else
            cols = Split(row, "|")
            Call ArrayUtils.Copy1DArrayTo2DArrayRow(vArray, i, cols)
        End If
        i = i + 1
    Next
    EditQueryResult = vArray
End Function
