VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CryptoUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] �Í����E�������[�e�B���e�B�N���X
'* [��  ��] E�Í����E�����Ɋւ��郆�[�e�B���e�B��񋟂���B
'*          Static�N���X�iAttribute VB_PredeclaredId = True�j�Ƃ���B
'* [��  ��] .Net Framework 3.5�̃I�u�W�F�N�g���g�p���邽�߁A
'*          .Net Framework 3.5���L���ɂȂ��Ă��Ȃ��ꍇ�́A�I�[�g���[�V����
'*          �G���[����������B
'*          <�L�����菇>
'*            �E�R���g���[���p�l�����v���O�����Ƌ@�\��Windows�@�\�̗L������
'*              ���������J��
'*            �E.Net Framework 3.5 ���`�F�b�N
'*            �EWindows Update �� .Net Framework 3.5 ����
'*
'* [�Q  �l] https://itbyari.wordpress.com/2015/12/08/excel-vba-%E6%9A%97%E5%8F%B7%E5%8C%96%E3%81%A8%E5%BE%A9%E5%8F%B7%E5%8C%96%EF%BC%88tripledes%EF%BC%89/
'*          https://qiita.com/nobu-maple/items/ac0094a226e1d7d25ec3
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
Public Enum StringConversionMethodEnum
    Base64
    HexString
End Enum

'******************************************************************************
'* �萔��`
'******************************************************************************

'******************************************************************************
'* �����ϐ���`
'******************************************************************************
Private mUTF8Encoding As Object
Private mTripleDESCryptoSp As Object
Private mSHA256 As Object
Private mInitializationVector As String
Private mTripleDesKey As String

'******************************************************************************
'* �v���p�e�B��`
'******************************************************************************
'*-----------------------------------------------------------------------------
'* InitializationVector �v���p�e�B
'* �����x�N�g���B�K��8�������w�肷��B��F"12345678"
'*-----------------------------------------------------------------------------
Public Property Get InitializationVector() As String
    InitializationVector = mInitializationVector
End Property
Public Property Let InitializationVector(arg As String)
    mInitializationVector = arg
End Property

'*-----------------------------------------------------------------------------
'* TripleDesKey �v���p�e�B
'* TripleDes�Í��p���ʌ��B�K��16�������w�肷��B��F"bankurarakusitai"
'*-----------------------------------------------------------------------------
Public Property Get TripleDesKey() As String
    TripleDesKey = mTripleDesKey
End Property
Public Property Let TripleDesKey(arg As String)
    mTripleDesKey = arg
End Property

'******************************************************************************
'* �R���X�g���N�^�E�f�X�g���N�^
'******************************************************************************
Private Sub Class_Initialize()
    Set mUTF8Encoding = CreateObject("System.Text.UTF8Encoding")
    Set mTripleDESCryptoSp = CreateObject("System.Security.Cryptography.TripleDESCryptoServiceProvider")
    Set mSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
     
End Sub
'******************************************************************************
'* ���\�b�h��`
'******************************************************************************

'******************************************************************************
'* [�T  �v] EncryptStringTripleDES
'* [��  ��] TripleDES�ɂ��Í����iTripleDES�Í�����BASE64�������j�B
'*
'* @param plainString �Í����Ώە���
'* @param stringConvMethod �ϊ����@�i�uBase64�v���uHexString�v(16�i��������)�j
'* @return �Í����i����I���j or Null�i�ُ�I��)
'******************************************************************************
Public Function EncryptStringTripleDES(plainString As String, _
    Optional stringConvMethod As StringConversionMethodEnum = Base64) As String
    
    On Error GoTo ErrorHandler

    ' ����������˕����o�C�g�f�[�^
    Dim plainByteData() As Byte
    plainByteData = mUTF8Encoding.GetBytes_4(plainString)
    
    ' �����o�C�g�f�[�^�ˈÍ��o�C�g�f�[�^
    Call Set3DESKeyAndIV2CryptoSp
    Dim encryptedByteData() As Byte
    encryptedByteData = Enc3DESTransformFinalBlock(plainByteData, 0, UBound(plainByteData) + 1)
    
    If stringConvMethod = Base64 Then
        ' �Í��o�C�g�f�[�^��BASE64����������
        EncryptStringTripleDES = EncodeDecodeUtils.BytesToBase64(encryptedByteData)
    Else
        ' �o�C�i����16�i��������ɕϊ�
        EncryptStringTripleDES = BinaryUtils.ByteArray2HexStrings(encryptedByteData, "")
    End If

    Exit Function
ErrorHandler:
    EncryptStringTripleDES = Null
    Debug.Print "TripleDES�ɂ��Í����Ɏ��s���܂����B"
End Function

'******************************************************************************
'* [�T  �v] DecryptStringTripleDES
'* [��  ��] TripleDES�ɂ�镜�����iBASE64��������DES�������j
'*
'* @param encryptedString �Í���
'* @param stringConvMethod �ϊ����@�i�uBase64�v���uHexString�v(16�i��������)�j
'* @return �����i����I���j or Null�i�ُ�I��)
'******************************************************************************
Public Function DecryptStringTripleDES(encryptedString As String, _
    Optional stringConvMethod As StringConversionMethodEnum = Base64) As String
    
    On Error GoTo ErrorHandler

    Dim encryptedByteData() As Byte
    If stringConvMethod = Base64 Then
        ' BASE64�����������DES�Í��o�C�g�f�[�^
        encryptedByteData = EncodeDecodeUtils.Base64toBytes(encryptedString)
    Else
        ' 16�i���������DES�Í��o�C�g�f�[�^
        encryptedByteData = BinaryUtils.HexStrings2ByteArray(encryptedString, "")
    End If
    
    ' DES�Í��o�C�g�f�[�^�˕����o�C�g�f�[�^
    Call Set3DESKeyAndIV2CryptoSp
    Dim plainByteData() As Byte
    plainByteData = Desc3DESTransformFinalBlock(encryptedByteData, 0, UBound(encryptedByteData) + 1)
            
    ' �����o�C�g�f�[�^�˕���������
    Dim plainString As String
    plainString = mUTF8Encoding.GetString(plainByteData)

    ' �߂�l�ݒ�
    DecryptStringTripleDES = plainString

    Exit Function
ErrorHandler:
    DecryptStringTripleDES = Null
    Debug.Print "TripleDES�ɂ�镜�����Ɏ��s���܂����B"
End Function

'******************************************************************************
'* [�T  �v] Set3DESKeyAndIV2CryptoSp
'* [��  ��] TripleDESCryptoServiceProvider��Key��InitializationVector��
'*          �ݒ肷��B
'*
'******************************************************************************
Private Sub Set3DESKeyAndIV2CryptoSp()
    mTripleDESCryptoSp.Key = mUTF8Encoding.GetBytes_4(mTripleDesKey)
    mTripleDESCryptoSp.IV = mUTF8Encoding.GetBytes_4(mInitializationVector)
End Sub

'******************************************************************************
'* [�T  �v] Enc3DESTransformFinalBlock
'* [��  ��] �w�肵���o�C�g�z��̎w�肵���̈��TripleDES�ɂ��Í�������B
'* [�Q  �l] https://docs.microsoft.com/ja-jp/dotnet/api/system.security.cryptography.icryptotransform.transformfinalblock
'*
'* @param inputBuffer �ϊ����v�Z����Ώۂ̓��̓f�[�^
'* @param inputOffset �o�C�g�z����̃f�[�^�̎g�p�J�n�ʒu�������I�t�Z�b�g
'* @param inputCount �o�C�g�z����Ńf�[�^�Ƃ��Ďg�p�����o�C�g�̐�
'* @return �ϊ�
'******************************************************************************
Public Function Enc3DESTransformFinalBlock(inputBuffer() As Byte, inputOffset As Long, inputCount As Long) As Byte()
    Enc3DESTransformFinalBlock = mTripleDESCryptoSp.CreateEncryptor() _
                                 .TransformFinalBlock(inputBuffer, inputOffset, inputCount)
End Function

'******************************************************************************
'* [�T  �v] Desc3DESTransformFinalBlock
'* [��  ��] �w�肵���o�C�g�z��̎w�肵���̈��TripleDES�ɂ�蕜������B
'* [�Q  �l] https://docs.microsoft.com/ja-jp/dotnet/api/system.security.cryptography.icryptotransform.transformfinalblock
'*
'* @param inputBuffer �ϊ����v�Z����Ώۂ̓��̓f�[�^
'* @param inputOffset �o�C�g�z����̃f�[�^�̎g�p�J�n�ʒu�������I�t�Z�b�g
'* @param inputCount �o�C�g�z����Ńf�[�^�Ƃ��Ďg�p�����o�C�g�̐�
'* @return �ϊ�
'******************************************************************************
Public Function Desc3DESTransformFinalBlock(inputBuffer() As Byte, inputOffset As Long, inputCount As Long) As Byte()
    Desc3DESTransformFinalBlock = mTripleDESCryptoSp.CreateDecryptor() _
                                  .TransformFinalBlock(inputBuffer, inputOffset, inputCount)
End Function


'******************************************************************************
'* [�T  �v] EncryptStringSHA256
'* [��  ��] SHA256�ɂ��Í����iSHA256�Í�����BASE64�������j�B
'*
'* @param plainString �Í����Ώە���
'* @param stringConvMethod �ϊ����@�i�uBase64�v���uHexString�v(16�i��������)�j
'* @return �Í����i����I���j or Null�i�ُ�I��)
'******************************************************************************
Function EncryptStringSHA256(plainString As String, _
                            Optional stringConvMethod As StringConversionMethodEnum = Base64) As String
    On Error GoTo ErrorHandler

    ' ����������˕����o�C�g�f�[�^�i�������UTF8�ɃG���R�[�h���o�C�g�z��ɕϊ��j
    Dim plainByteData() As Byte
    plainByteData = mUTF8Encoding.GetBytes_4(plainString)

    '�����o�C�g�f�[�^�ˈÍ��o�C�g�f�[�^�i�n�b�V���l�j
    Dim encryptedByteData() As Byte
    encryptedByteData = mSHA256.ComputeHash_2((plainByteData))

    If stringConvMethod = Base64 Then
        ' �Í��o�C�g�f�[�^��BASE64����������
        EncryptStringSHA256 = EncodeDecodeUtils.BytesToBase64(encryptedByteData)
    Else
        ' �o�C�i����16�i��������ɕϊ�
        EncryptStringSHA256 = BinaryUtils.ByteArray2HexStrings(encryptedByteData, "")
    End If

    Exit Function
ErrorHandler:
    EncryptStringSHA256 = Null
    Debug.Print "SHA256�ɂ��Í����Ɏ��s���܂����B"
End Function



