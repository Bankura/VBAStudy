VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NetworkDriveUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] ネットワークドライブユーティリティクラス
'* [詳  細] ネットワークドライブに関するユーティリティを提供する。
'*          Staticクラス（Attribute VB_PredeclaredId = True）とする。
'* [参  考]
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************

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


'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] ConnectNetworkDrive
'* [詳  細] ネットワークドライブに接続する。
'*
'* @param strNdName ネットワークドライブ名
'* @param strUncPath 接続先パス
'* @param strUserId  ユーザID
'* @param strPass    パスワード
'* @return Boolean 処理結果（True:正常 False：異常）
'******************************************************************************
Public Function ConnectNetworkDrive(strNdName As String, strUncPath As String, _
                                    strUserId As String, strPass As String) As Boolean
    On Error GoTo ErrorHandler
    Base.GetWshNetwork.MapNetworkDrive strNdName, strUncPath, False, strUserId, strPass
    
    ConnectNetworkDrive = True
    Exit Function
ErrorHandler:
    ConnectNetworkDrive = False
    
End Function

'******************************************************************************
'* [概  要] DisconnectNetworkDrive
'* [詳  細] ネットワークドライブから切断する。
'*
'* @param strNdName ネットワークドライブ名
'* @return Boolean 処理結果（True:正常 False：異常）
'******************************************************************************
Public Function DisconnectNetworkDrive(strNdName As String) As Boolean
    On Error GoTo ErrorHandler
    Base.GetWshNetwork.RemoveNetworkDrive strNdName, True, True
    DisconnectNetworkDrive = True
    Exit Function
ErrorHandler:
    DisconnectNetworkDrive = False
    
End Function

'******************************************************************************
'* [概  要] GetUnusedNetworkDriveName
'* [詳  細] 使用されていないネットワークドライブ名を取得する。
'*
'* @return String ネットワークドライブ名
'******************************************************************************
Public Function GetUnusedNetworkDriveName() As String
    Dim i As Long
    Dim objNetWork As Object
    Dim colDrives As Object
    Set objNetWork = Base.GetWshNetwork
    
    Dim vNdArr() As Variant
    Dim strTmpNd As Variant
    Dim flg As Boolean
    flg = True
    vNdArr = Array("E:", "F:", "G:", "H:", "I:", "J:", "K:", "L:", "M:", "N:", "O:", "P:", "Q:", "R:", "S:", "T:", "U:", "V:", "W:", "X:", "Y:", "Z:")
    
    Set colDrives = objNetWork.EnumNetworkDrives
    For Each strTmpNd In vNdArr
        flg = True
        For i = 0 To colDrives.Count - 1 Step 2
            If colDrives.Item(i) = strTmpNd Then
                flg = False
                Exit For
            End If
        Next i
        If flg Then
            Exit For
        End If
    Next
  
    If Not flg Then
        strTmpNd = ""
    End If
    Set colDrives = Nothing
    Set objNetWork = Nothing
    GetUnusedNetworkDriveName = strTmpNd
End Function



