VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScriptingExDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Scripting.Driveクラスのラップ・拡張クラス
'* [詳  細] Scripting.DriveクラスのWrapperとしての機能を提供する他、
'*          Scripting.Driveクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' Scripting.Driveクラスについて。
' <https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/drive-object>
' 特定のディスク ドライブまたはネットワーク共有のプロパティにアクセスできます。
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As Scripting.Drive
Private mOrigin As Object
    
'******************************************************************************
'* プロパティ定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* オリジナルObject プロパティ
'*-----------------------------------------------------------------------------
Property Get Origin() As Object
    Set Origin = mOrigin
End Property
Property Set Origin(arg As Object)
    Set mOrigin = arg
End Property

'*-----------------------------------------------------------------------------
'* AvailableSpace プロパティ（読み取り専用）
'*
'* [補  足]
'* ・使用可能なディスク容量（バイト単位）を取得します。
'* ・通常、FreeSpace プロパティ が返す値と同じ
'*   （OS が、ディスククォータ機能をサポートしている場合は、
'*     違う値になる可能性がある）。
'*-----------------------------------------------------------------------------
Property Get AvailableSpace() As Variant
    AvailableSpace = mOrigin.AvailableSpace
End Property


'*-----------------------------------------------------------------------------
'* DriveLetter プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ドライブ名を取得します（C や D など）。
'*-----------------------------------------------------------------------------
Property Get DriveLetter() As String
    DriveLetter = mOrigin.DriveLetter
End Property


'*-----------------------------------------------------------------------------
'* DriveType プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ドライブの種類。
'*-----------------------------------------------------------------------------
Property Get DriveType() As DriveTypeConst
    DriveType = mOrigin.DriveType
End Property

'*-----------------------------------------------------------------------------
'* FileSystem プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ファイルシステムの種類（FAT、NTFS、CDFS など）。
'*-----------------------------------------------------------------------------
Property Get FileSystem() As String
    FileSystem = mOrigin.FileSystem
End Property

'*-----------------------------------------------------------------------------
'* FreeSpace プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ドライブの空き容量（バイト単位）を取得する。
'* ・通常、AvailableSpace プロパティ が返す値と同じ
'*   （OS が、ディスククォータ機能をサポートしている場合は、
'*     違う値になる可能性がある）。
'*-----------------------------------------------------------------------------
Property Get FreeSpace() As Variant
    FreeSpace = mOrigin.FreeSpace
End Property

'*-----------------------------------------------------------------------------
'* IsReady プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ディスクが使用可能かどうか調べます。
'* ・CD-ROM や DVD ディスクなどの場合は、メディアが挿入されているかどうかを
'*   確認することができます。
'* ・ドライブの準備ができている場合は、True 、できていない場合は、False。
'*-----------------------------------------------------------------------------
Property Get IsReady() As Boolean
    IsReady = mOrigin.IsReady
End Property

'*-----------------------------------------------------------------------------
'* Path プロパティ（読み取り専用:既定メンバー）
'*
'* [補  足]
'* ・ドライブのパス。
'*-----------------------------------------------------------------------------
Property Get Path() As String
Attribute Path.VB_UserMemId = 0
    Path = mOrigin.Path
End Property

'*-----------------------------------------------------------------------------
'* RootFolder プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ルートフォルダ。
'*-----------------------------------------------------------------------------
Property Get RootFolder() As ScriptingExFolder
    Dim wrap As ScriptingExFolder: Set wrap = New ScriptingExFolder
    Set wrap.Origin = mOrigin.RootFolder()
    Set RootFolder = wrap
End Property

'*-----------------------------------------------------------------------------
'* SerialNumber プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ディスクのシリアル番号。
'* ・CD-ROMドライブなどでは、挿入されているメディアのシリアルナンバー。
'*-----------------------------------------------------------------------------
Property Get SerialNumber() As Long
    SerialNumber = mOrigin.SerialNumber
End Property

'*-----------------------------------------------------------------------------
'* ShareName プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ドライブのネットワーク共有名。
'* ・指定したドライブがネットワークドライブでない場合は、長さ 0 の文字列（””）
'*-----------------------------------------------------------------------------
Property Get ShareName() As String
    ShareName = mOrigin.ShareName
End Property

'*-----------------------------------------------------------------------------
'* TotalSize プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ドライブの全体の容量（バイト単位）を返す。
'*-----------------------------------------------------------------------------
Property Get TotalSize() As Variant
    TotalSize = mOrigin.TotalSize
End Property

'*-----------------------------------------------------------------------------
'* VolumeName プロパティ
'*
'* [補  足]
'* ・ドライブのボリューム名。
'*-----------------------------------------------------------------------------
Property Get VolumeName() As String
    VolumeName = mOrigin.VolumeName
End Property
Property Let VolumeName(arg As String)
    mOrigin.VolumeName = arg
End Property

'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    'Set mOrigin = CreateObject("Scripting.Drive")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************
'なし


