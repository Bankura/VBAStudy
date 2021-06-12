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
Public Property Get DbName() As String
    DbName = mDbName
End Property
Public Property Let DbName(arg As String)
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
Public Property Get ClEncode() As String
    ClEncode = mClEncode
End Property
Public Property Let ClEncode(arg As String)
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
'* [�T  �v] psql�R�}���h�����SQL�����s����B
'* [��  ��] psql�R�}���h����āASQL�����s���A���s���ʂ�ԋp����B
'*          ���[�J������psql���C���X�g�[������Ă���K�v������B
'*          psql��postgreSQL�C���X�g�[�����ɍ��킹�ăC���X�g�[�������B
'*
'* @param sql SQL��
'* @return CommandResult���
'*
'******************************************************************************
Public Function ExecPsql(sql As String) As CommandResult
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
        Set ExecPsql = res
        Exit Function
    End If

    ' ���펞����
    Dim lRowMax As Long: lRowMax = 0
    Dim lColMax As Long: lColMax = 0
    
    If ArrayUtils.IsEmptyArray(res.StdOutList) Then
        Set ExecPsql = res
        Exit Function
    End If
    ' �񐔌v��
    Dim vTmpCols As Variant: vTmpCols = Split(res.StdOutList(0), "|")
    lColMax = UBound(vTmpCols) - LBound(vTmpCols) + 1

    '�f�[�^�ݒ�
    Dim vArray()
    ReDim vArray(0 To lRowMax - 1, 0 To lColMax - 1)
    
    Dim i As Long, cols As Variant
    Dim row
    For Each row In res.StdOutList
        Debug.Print "Line: " & row
        
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

    res.StdOutList = vArray
    res.RowMax = lRowMax
    res.ColMax = lColMax
    Set ExecPsql = res
End Function