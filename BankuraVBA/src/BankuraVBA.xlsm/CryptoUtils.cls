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
'* [機能名] 暗号化・復号ユーティリティクラス
'* [詳  細] E暗号化・復号に関するユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'* [注  意] .Net Framework 3.5のオブジェクトを使用するため、
'*          .Net Framework 3.5が有効になっていない場合は、オートメーション
'*          エラーが発生する。
'*          <有効化手順>
'*            ・コントロールパネル＞プログラムと機能＞Windows機能の有効化と
'*              無効化を開く
'*            ・.Net Framework 3.5 をチェック
'*            ・Windows Update で .Net Framework 3.5 入手
'*
'* [参  考] https://itbyari.wordpress.com/2015/12/08/excel-vba-%E6%9A%97%E5%8F%B7%E5%8C%96%E3%81%A8%E5%BE%A9%E5%8F%B7%E5%8C%96%EF%BC%88tripledes%EF%BC%89/
'*          https://qiita.com/nobu-maple/items/ac0094a226e1d7d25ec3
'*
'* @author Bankura
'* Copyright (c) 2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************

'******************************************************************************
'* Enum定義
'******************************************************************************
Public Enum StringConversionMethodEnum
    Base64
    HexString
End Enum

'******************************************************************************
'* 定数定義
'******************************************************************************

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mUTF8Encoding As Object
Private mTripleDESCryptoSp As Object
Private mSHA256 As Object
Private mInitializationVector As String
Private mTripleDesKey As String

'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* InitializationVector プロパティ
'* 初期ベクトル。必ず8文字分指定する。例："12345678"
'*-----------------------------------------------------------------------------
Public Property Get InitializationVector() As String
    InitializationVector = mInitializationVector
End Property
Public Property Let InitializationVector(arg As String)
    mInitializationVector = arg
End Property

'*-----------------------------------------------------------------------------
'* TripleDesKey プロパティ
'* 暗号用共通鍵。必ず16文字分指定する。例："bankurarakusitai"
'*-----------------------------------------------------------------------------
Public Property Get TripleDesKey() As String
    TripleDesKey = mTripleDesKey
End Property
Public Property Let TripleDesKey(arg As String)
    mTripleDesKey = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mUTF8Encoding = CreateObject("System.Text.UTF8Encoding")
    Set mTripleDESCryptoSp = CreateObject("System.Security.Cryptography.TripleDESCryptoServiceProvider")
    Set mSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
     
End Sub
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] EncryptStringTripleDES
'* [詳  細] TripleDESによる暗号化（TripleDES暗号化⇒BASE64符号化）。
'*
'* @param plainString 暗号化対象平文
'* @param stringConvMethod 変換方法（「Base64」か「HexString」(16進数文字列)）
'* @return 暗号文（正常終了） or Null（異常終了)
'******************************************************************************
Public Function EncryptStringTripleDES(plainString As String, _
    Optional stringConvMethod As StringConversionMethodEnum = Base64) As String
    
    On Error GoTo ErrorHandler

    ' 平文文字列⇒平文バイトデータ
    Dim plainByteData() As Byte
    plainByteData = mUTF8Encoding.GetBytes_4(plainString)
    
    ' 平文バイトデータ⇒暗号バイトデータ
    Call Set3DESKeyAndIV2CryptoSp
    Dim encryptedByteData() As Byte
    encryptedByteData = Enc3DESTransformFinalBlock(plainByteData, 0, UBound(plainByteData) + 1)
    
    If stringConvMethod = Base64 Then
        ' 暗号バイトデータ⇒BASE64符号文字列
        EncryptStringTripleDES = EncodeDecodeUtils.BytesToBase64(encryptedByteData)
    Else
        ' バイナリを16進数文字列に変換
        EncryptStringTripleDES = BinaryUtils.ByteArray2HexStrings(encryptedByteData, "")
    End If

    Exit Function
ErrorHandler:
    EncryptStringTripleDES = Null
    Debug.Print "TripleDESによる暗号化に失敗しました。"
End Function

'******************************************************************************
'* [概  要] DecryptStringTripleDES
'* [詳  細] TripleDESによる復号化（BASE64復号化⇒DES復号化）
'*
'* @param encryptedString 暗号文
'* @param stringConvMethod 変換方法（「Base64」か「HexString」(16進数文字列)）
'* @return 平文（正常終了） or Null（異常終了)
'******************************************************************************
Public Function DecryptStringTripleDES(encryptedString As String, _
    Optional stringConvMethod As StringConversionMethodEnum = Base64) As String
    
    On Error GoTo ErrorHandler

    Dim encryptedByteData() As Byte
    If stringConvMethod = Base64 Then
        ' BASE64符号文字列⇒DES暗号バイトデータ
        encryptedByteData = EncodeDecodeUtils.Base64toBytes(encryptedString)
    Else
        ' 16進数文字列⇒DES暗号バイトデータ
        encryptedByteData = BinaryUtils.HexStrings2ByteArray(encryptedString, "")
    End If
    
    ' DES暗号バイトデータ⇒平文バイトデータ
    Call Set3DESKeyAndIV2CryptoSp
    Dim plainByteData() As Byte
    plainByteData = Desc3DESTransformFinalBlock(encryptedByteData, 0, UBound(encryptedByteData) + 1)
            
    ' 平文バイトデータ⇒平文文字列化
    Dim plainString As String
    plainString = mUTF8Encoding.GetString(plainByteData)

    ' 戻り値設定
    DecryptStringTripleDES = plainString

    Exit Function
ErrorHandler:
    DecryptStringTripleDES = Null
    Debug.Print "TripleDESによる復号化に失敗しました。"
End Function

'******************************************************************************
'* [概  要] Set3DESKeyAndIV2CryptoSp
'* [詳  細] TripleDESCryptoServiceProviderにKeyとInitializationVectorを
'*          設定する。
'*
'******************************************************************************
Private Sub Set3DESKeyAndIV2CryptoSp()
    mTripleDESCryptoSp.Key = mUTF8Encoding.GetBytes_4(mTripleDesKey)
    mTripleDESCryptoSp.IV = mUTF8Encoding.GetBytes_4(mInitializationVector)
End Sub

'******************************************************************************
'* [概  要] Enc3DESTransformFinalBlock
'* [詳  細] 指定したバイト配列の指定した領域をTripleDESにより暗号化する。
'* [参  考] https://docs.microsoft.com/ja-jp/dotnet/api/system.security.cryptography.icryptotransform.transformfinalblock
'*
'* @param inputBuffer 変換を計算する対象の入力データ
'* @param inputOffset バイト配列内のデータの使用開始位置を示すオフセット
'* @param inputCount バイト配列内でデータとして使用されるバイトの数
'* @return 変換
'******************************************************************************
Public Function Enc3DESTransformFinalBlock(inputBuffer() As Byte, inputOffset As Long, inputCount As Long) As Byte()
    Enc3DESTransformFinalBlock = mTripleDESCryptoSp.CreateEncryptor() _
                                 .TransformFinalBlock(inputBuffer, inputOffset, inputCount)
End Function

'******************************************************************************
'* [概  要] Desc3DESTransformFinalBlock
'* [詳  細] 指定したバイト配列の指定した領域をTripleDESにより復号する。
'* [参  考] https://docs.microsoft.com/ja-jp/dotnet/api/system.security.cryptography.icryptotransform.transformfinalblock
'*
'* @param inputBuffer 変換を計算する対象の入力データ
'* @param inputOffset バイト配列内のデータの使用開始位置を示すオフセット
'* @param inputCount バイト配列内でデータとして使用されるバイトの数
'* @return 変換
'******************************************************************************
Public Function Desc3DESTransformFinalBlock(inputBuffer() As Byte, inputOffset As Long, inputCount As Long) As Byte()
    Desc3DESTransformFinalBlock = mTripleDESCryptoSp.CreateDecryptor() _
                                  .TransformFinalBlock(inputBuffer, inputOffset, inputCount)
End Function


'******************************************************************************
'* [概  要] EncryptStringSHA256
'* [詳  細] SHA256による暗号化（SHA256暗号化⇒BASE64符号化）。
'*
'* @param plainString 暗号化対象平文
'* @param stringConvMethod 変換方法（「Base64」か「HexString」(16進数文字列)）
'* @return 暗号文（正常終了） or Null（異常終了)
'******************************************************************************
Function EncryptStringSHA256(plainString As String, _
                            Optional stringConvMethod As StringConversionMethodEnum = Base64) As String
    On Error GoTo ErrorHandler

    ' 平文文字列⇒平文バイトデータ（文字列をUTF8にエンコードしバイト配列に変換）
    Dim plainByteData() As Byte
    plainByteData = mUTF8Encoding.GetBytes_4(plainString)

    '平文バイトデータ⇒暗号バイトデータ（ハッシュ値）
    Dim encryptedByteData() As Byte
    encryptedByteData = mSHA256.ComputeHash_2((plainByteData))

    If stringConvMethod = Base64 Then
        ' 暗号バイトデータ⇒BASE64符号文字列
        EncryptStringSHA256 = EncodeDecodeUtils.BytesToBase64(encryptedByteData)
    Else
        ' バイナリを16進数文字列に変換
        EncryptStringSHA256 = BinaryUtils.ByteArray2HexStrings(encryptedByteData, "")
    End If

    Exit Function
ErrorHandler:
    EncryptStringSHA256 = Null
    Debug.Print "SHA256による暗号化に失敗しました。"
End Function




