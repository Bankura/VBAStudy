VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IniFileAccessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] INIファイルアクセスクラス
'* [詳  細] INIファイルにアクセスする処理を提供する。
'*
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
Private mIniFilePath As String
Private mIniStringMaxSize As Long
Private mWinApi As WinAPI

'******************************************************************************
'* プロパティ定義
'******************************************************************************


'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    mIniStringMaxSize = 4096
    Set mWinApi = Base.GetWinAPI()
End Sub
Private Sub Class_Terminate()
    Set mWinApi = Nothing
End Sub

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Init メソッド
'* [詳  細] 初期化用メソッド｡
'*
'* @param iniFilePath INIファイルパス
'* @param iniStringMaxSize 文字項目最大サイズ
'* @return IniFileAccessorオブジェクト
'*
'******************************************************************************
Public Function Init(iniFilePath As String, Optional iniStringMaxSize As Long = 4096) As IniFileAccessor
    mIniFilePath = iniFilePath
    mIniStringMaxSize = iniStringMaxSize
    Set Init = Me
End Function

'******************************************************************************
'* [概  要] INIファイル項目設定値取得処理
'* [詳  細] 指定されたINIファイルパス/セクション/項目の値を取得する。
'*          WindowsAPIのGetPrivateProfileStringを使用。
'*
'* @param sectionName セクション名
'* @param itmName 項目名
'* @return String INIファイル項目設定値（NG時は空文字列返却）
'******************************************************************************
Public Function GetStringValue(sectionName As String, itmName As String) As String
                                 
    Dim retString As String: retString = Space$(mIniStringMaxSize)
    On Error Resume Next
    Dim ret As Long
    ret = mWinApi.GetPrivateProfileString(sectionName, itmName, vbNullString, _
                                          retString, Len(retString), mIniFilePath)
    On Error GoTo 0
    If ret = 0 Then
        GetStringValue = ""
        Exit Function
    End If

    GetStringValue = Left$(retString, InStr(retString, Chr$(0)) - 1)
End Function

