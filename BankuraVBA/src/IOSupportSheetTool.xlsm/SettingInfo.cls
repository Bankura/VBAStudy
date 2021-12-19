VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SettingInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] 設定情報クラス
'* [詳  細] 設定情報を保持するためのクラス。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* 内部定数定義
'******************************************************************************
Private Const SETTING_SHEET_NAME As String = "setting"
Private Const SETTING_SH_START_ROW As Long = 4
Private Const SETTING_SH_START_COL As Long = 4

'******************************************************************************
'* 内部変数定義
'******************************************************************************
Private mMasters As Object 'Dictionary
Private mSetting As Object 'Dictionary
Private mSettingSheetName As String
Private mSettingSheetStartRow As Long
Private mSettingSheetStartCol As Long

'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* SettingSheetName プロパティ
'*-----------------------------------------------------------------------------
Public Property Get SettingSheetName() As String
    SettingSheetName = mSettingSheetName
End Property
Public Property Let SettingSheetName(ByVal arg As String)
    mSettingSheetName = arg
End Property

'*-----------------------------------------------------------------------------
'* SettingSheetStartRow プロパティ
'*-----------------------------------------------------------------------------
Public Property Get SettingSheetStartRow() As Long
    SettingSheetStartRow = mSettingSheetStartRow
End Property
Public Property Let SettingSheetStartRow(ByVal arg As Long)
    mSettingSheetStartRow = arg
End Property

'*-----------------------------------------------------------------------------
'* SettingSheetStartCol プロパティ
'*-----------------------------------------------------------------------------
Public Property Get SettingSheetStartCol() As Long
    SettingSheetStartCol = mSettingSheetStartCol
End Property
Public Property Let SettingSheetStartCol(ByVal arg As Long)
    mSettingSheetStartCol = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mSetting = CreateObject("Scripting.Dictionary")
    Set mMasters = CreateObject("Scripting.Dictionary")
End Sub
Private Sub Class_Terminate()
    Set mMasters = Nothing
    Set mSetting = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] Init メソッド
'* [詳  細] 初期化用メソッド｡
'*
'* @param sSheetName 設定シート名
'* @param sSheetStartRow 初期確保領域（バイト）
'* @param sSheetStartCol 初期確保領域（バイト）
'* @return SettingInfoオブジェクト
'*
'******************************************************************************
Public Function Init(ByVal sSheetName As String, ByVal sSheetStartRow As Long, ByVal sSheetStartCol As Long) As SettingInfo
    mSettingSheetName = sSheetName
    mSettingSheetStartRow = sSheetStartRow
    mSettingSheetStartCol = sSheetStartCol
    Call GetSettingInfo
    Call GetMasterInfo
    Set Init = Me
End Function

'******************************************************************************
'* [概  要] 設定情報取得処理。
'* [詳  細] 設定情報を設定シートより取得し、保持します｡
'*
'******************************************************************************
Private Sub GetSettingInfo()
    Dim wb As Workbook, ws As Worksheet
    Set wb = Application.ThisWorkbook
    Set ws = wb.Sheets(mSettingSheetName)
    Set mSetting = GetMasterDataFromSheet(ws, mSettingSheetStartRow, mSettingSheetStartCol, 2)
End Sub

'******************************************************************************
'* [概  要] マスタ情報一括取得処理。
'* [詳  細] マスタ情報を一括取得し、保持します｡
'*
'******************************************************************************
Private Sub GetMasterInfo()
    Dim mstHeadLetters As String: mstHeadLetters = GetSettingValue("MasterHeadLetters")
    Dim startRow As Long: startRow = GetSettingValue("MasterStartRowNo")
    Dim startCol As Long: startCol = GetSettingValue("MasterStartColNo")
    
    Dim reg As Object
    Set reg = GetRegExp()
    With reg
        .Pattern = "^" & mstHeadLetters & ".*"
        .IgnoreCase = True
        .Global = True
    End With
    
    Dim wb As Workbook, ws As Worksheet
    Set wb = Application.ThisWorkbook
    For Each ws In wb.Worksheets
        If reg.test(ws.Name) Then
            Dim master As Object
            Set master = GetMasterDataFromSheet(ws, startRow, startCol, 2)
            Call mMasters.Add(ws.Name, master)
        End If
    Next
End Sub

'******************************************************************************
'* [概  要] マスタ情報取得処理。
'* [詳  細] マスタ情報を取得し、内部ディクショナリに保持します｡
'*
'* @param dataSheet ワークシート
'* @param lStartRow データ開始行番号
'* @param lStartCol データ開始列番号
'* @param itemCount 項目列数
'*
'******************************************************************************
Private Function GetMasterDataFromSheet(dataSheet As Worksheet, lStartRow As Long, lStartCol As Long, Optional colCount As Long) As Object
    Dim vArr: vArr = XlWorkSheetUtils.GetVariantDataFromSheet(dataSheet, lStartRow, lStartCol, colCount)
    If IsEmpty(vArr) Then
        Exit Function
    End If
    Dim master As Object: Set master = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(vArr, 1) To UBound(vArr, 1)
        Dim code As String: code = vArr(i, 1)
        Dim val As String: val = vArr(i, 2)
        Call master.Add(code, val)
    Next
    Set GetMasterDataFromSheet = master
End Function

'******************************************************************************
'* [概  要] 設定値取得処理。
'* [詳  細] 設定情報から設定値を取得します｡
'*
'* @param key キー値
'* @return 設定値
'*
'******************************************************************************
Public Function GetSettingValue(ByVal Key As String) As String
    If mSetting.Exists(Key) Then
        GetSettingValue = mSetting.Item(Key)
    Else
        GetSettingValue = ""
    End If
End Function

'******************************************************************************
'* [概  要] マスタ取得処理。
'* [詳  細] マスタ情報からマスタオブジェクトを取得します｡
'*
'* @param masterName キー値（マスタシート名）
'* @return マスタオブジェクト
'*
'******************************************************************************
Public Function GetMaster(ByVal masterName As String) As Object
    If mMasters.Exists(masterName) Then
        Set GetMaster = mMasters.Item(masterName)
    Else
        Set GetMaster = Nothing
    End If
End Function

'******************************************************************************
'* [概  要] マスタ値取得処理。
'* [詳  細] マスタ情報からマスタ名・コード値（キー）を指定して値を取得します｡
'*
'* @param masterName マスタシート名
'* @param code キー値（コード）
'* @return 値
'*
'******************************************************************************
Public Function GetMasterValueByCode(ByVal masterName As String, ByVal code As String) As String
    Dim master As Object
    If mMasters.Exists(masterName) Then
        Set master = mMasters.Item(masterName)
        If master.Exists(code) Then
            GetMasterValueByCode = master.Item(code)
            Exit Function
        End If
    End If
    GetMasterValueByCode = ""
End Function

'******************************************************************************
'* [概  要] マスタコード値取得処理。
'* [詳  細] マスタ情報からマスタ名・値を指定してコード値を取得します｡
'*          前提として、値が一意（重複値なし）である必要があります。
'*
'* @param masterName マスタシート名
'* @param val キー値（コード）
'* @return コード値
'*
'******************************************************************************
Public Function GetMasterCodeByValue(ByVal masterName As String, ByVal val As String) As String
    Dim master As Object, vKey
    If mMasters.Exists(masterName) Then
        Set master = mMasters.Item(masterName)
        For Each vKey In master.keys
            If master.Item(vKey) = val Then
                GetMasterCodeByValue = CStr(vKey)
                Exit Function
            End If
        Next
    End If
    GetMasterCodeByValue = ""
End Function
