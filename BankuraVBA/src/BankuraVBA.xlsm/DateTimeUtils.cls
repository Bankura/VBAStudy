VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTimeUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] ���t�E���ԑ��샆�[�e�B���e�B�N���X
'* [��  ��] ���t�E���ԑ���Ɋւ��郆�[�e�B���e�B��񋟂���B
'*          Static�N���X�iAttribute VB_PredeclaredId = True�j�Ƃ���B
'* [�Q  �l]
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
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

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************


'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
End Sub
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] BeginOfMonth
'* [��  ��] �w�肵�����t�̑����錎�̊J�n�����擾����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param dt �Ώۓ��t
'* @return Date �����̓��t
'******************************************************************************
Public Function BeginOfMonth(ByVal dt As Date) As Date
    BeginOfMonth = Core.BeginOfMonth(dt)
End Function

'******************************************************************************
'* [�T  �v] EndOfMonth
'* [��  ��] �w�肵�����t�̑����錎�̍ŏI�����擾����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param dt �Ώۓ��t
'* @return Date �����̓��t
'******************************************************************************
Public Function EndOfMonth(ByVal dt As Date) As Date
    EndOfMonth = Core.EndOfMonth(dt)
End Function

'******************************************************************************
'* [�T  �v] BeginOfWeek
'* [��  ��] �w�肵�����t�̑�����T�̊J�n�����擾����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param dt �Ώۓ��t
'* @param fstDayOfWeek �T�̎n�܂�̗j��
'* @return Date �T���̓��t
'******************************************************************************
Public Function BeginOfWeek(ByVal dt As Date, Optional fstDayOfWeek As VbDayOfWeek = vbSunday) As Date
    BeginOfWeek = Core.BeginOfWeek(dt, fstDayOfWeek)
End Function

'******************************************************************************
'* [�T  �v] EndOfWeek
'* [��  ��] �w�肵�����t�̑�����T�̍ŏI�����擾����B
'*          Ariawase��Core���W���[�����b�p�[���\�b�h�B
'*
'* @param dt �Ώۓ��t
'* @param fstDayOfWeek �T�̎n�܂�̗j��
'* @return Date �T���̓��t
'******************************************************************************
Public Function EndOfWeek(ByVal dt As Date, Optional fstDayOfWeek As VbDayOfWeek = vbSunday) As Date
    EndOfWeek = Core.EndOfWeek(dt, fstDayOfWeek)
End Function

'******************************************************************************
'* [�T  �v] IsLeapYear
'* [��  ��] �[�N���ǂ����`�F�b�N����B
'*
'* @param yearValue �Ώ۔N
'* @return Boolean �`�F�b�N���ʁiTrue�F�[�N�j
'******************************************************************************
Public Function IsLeapYear(yearValue As Integer) As Boolean
    If ((yearValue Mod 4) = 0 And (yearValue Mod 100) <> 0 Or (yearValue Mod 400) = 0) Then
        IsLeapYear = True
        Exit Function
    End If
    IsLeapYear = False
End Function