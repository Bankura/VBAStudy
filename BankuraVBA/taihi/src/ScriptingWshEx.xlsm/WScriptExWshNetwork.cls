VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WScriptExWshNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WScript.WshNetworkクラスのラップ・拡張クラス
'* [詳  細] WScript.WshNetworkクラスのWrapperとしての機能を提供する他、
'*          WScript.WshNetworkクラスを使用したユーティリティを提供する。
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'------------------------------------------------------------------------------
' [補  足]
' WScript.WshNetworkクラスについて。
' コンピュータが接続されているネットワーク上の共有リソースへのアクセスを提供します。
' <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364454(v=msdn.10)>
'
'------------------------------------------------------------------------------

'******************************************************************************
'* 内部変数定義
'******************************************************************************
'Private mOrigin As IWshRuntimeLibrary.WshNetwork
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
'* ComputerName プロパティ（読み取り専用）
'*
'* [補  足]
'* ・コンピュータ名の文字列表現を提供します。
'*-----------------------------------------------------------------------------
Property Get ComputerName() As String
    ComputerName = mOrigin.ComputerName
End Property

'*-----------------------------------------------------------------------------
'* UserDomain プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ユーザー ドメイン名の文字列表現を提供します。
'*-----------------------------------------------------------------------------
Property Get UserDomain() As String
    UserDomain = mOrigin.UserDomain
End Property

'*-----------------------------------------------------------------------------
'* UserName プロパティ（読み取り専用）
'*
'* [補  足]
'* ・ユーザー名の文字列表現を提供します。
'*-----------------------------------------------------------------------------
Property Get UserName() As String
    UserName = mOrigin.UserName
End Property

'*-----------------------------------------------------------------------------
'* Organization プロパティ（読み取り専用）
'*
'* [補  足]
'* ・隠し項目。
'*-----------------------------------------------------------------------------
'Property Get Organization() As String
'    Organization = mOrigin.Organization
'End Property

'*-----------------------------------------------------------------------------
'* Site プロパティ（読み取り専用）
'*
'* [補  足]
'* ・隠し項目。
'*-----------------------------------------------------------------------------
'Property Get Site() As String
'    Site = mOrigin.Site
'End Property

'*-----------------------------------------------------------------------------
'* UserProfile プロパティ（読み取り専用）
'*
'* [補  足]
'* ・隠し項目。
'*-----------------------------------------------------------------------------
'Property Get UserProfile() As String
'    UserProfile = mOrigin.UserProfile
'End Property
    
'******************************************************************************
'* コンストラクタ・デストラクタ
'******************************************************************************
Private Sub Class_Initialize()
    Set mOrigin = CreateObject("WScript.Network")
End Sub
Private Sub Class_Terminate()
    Set mOrigin = Nothing
End Sub
 
'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] AddPrinterConnection メソッド
'* [詳  細] AddPrinterConnection のラッパーメソッド。
'*          リモートの MS-DOS ベースのプリンタ接続をコンピュータ システムに
'*          追加します。
'* [参  考] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364413%28v%3dmsdn.10%29>
'*
'* @param LocalName 接続プリンタに割り当てるローカル名を示す文字列値
'* @param RemoteName リモートプリンタ名を示す文字列値
'* @param UpdateProfile 省略可能。プリンタのマッピング情報を現在のユーザー
'*                      プロファイルに格納するかどうかを示すブール値。
'*                      TRUE を指定すると、この割り当てがユーザープロファイル
'*                      に保存されます。既定値は FALSE です。
'* @param UserName 省略可能。ユーザー名を示す文字列値。
'* @param Password 省略可能。ユーザーのパスワードを示す文字列値。
'*                 UserName および Password は、現在のユーザー以外のアカウント
'*                 情報を使ってリモート プリンタを割り当てる場合に使用します。
'*
'******************************************************************************
Sub AddPrinterConnection(LocalName As String, RemoteName As String, Optional UpdateProfile, Optional UserName, Optional Password)
    If IsMissing(UpdateProfile) Then
        If IsMissing(UserName) Then
             If IsMissing(Password) Then
                Call mOrigin.AddPrinterConnection(LocalName, RemoteName)
            Else
                Call mOrigin.AddPrinterConnection(LocalName, RemoteName, , , Password)
            End If
        ElseIf IsMissing(Password) Then
            Call mOrigin.AddPrinterConnection(LocalName, RemoteName, , UserName)
        Else
            Call mOrigin.AddPrinterConnection(LocalName, RemoteName, , UserName, Password)
        End If
    Else
        If IsMissing(UserName) Then
             If IsMissing(Password) Then
                Call mOrigin.AddPrinterConnection(LocalName, RemoteName, UpdateProfile)
            Else
                Call mOrigin.AddPrinterConnection(LocalName, RemoteName, UpdateProfile, , Password)
            End If
        ElseIf IsMissing(Password) Then
            Call mOrigin.AddPrinterConnection(LocalName, RemoteName, UpdateProfile, UserName)
        Else
            Call mOrigin.AddPrinterConnection(LocalName, RemoteName, UpdateProfile, UserName, Password)
        End If
    End If
End Sub

'******************************************************************************
'* [概  要] AddWindowsPrinterConnection メソッド
'* [詳  細] AddWindowsPrinterConnection のラッパーメソッド。
'*          Windows ベースのプリンタ接続をコンピュータ システムに追加します。
'* [参  考] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364395%28v%3dmsdn.10%29>
'*
'* @param PrinterName プリンタ接続へのパスを示す文字列値
'* @param DriverName ドライブ名を示す文字列値。
'*                    (Windows NT/Windows 2000 では無視されます)。
'* @param Port 省略可能。プリンタ接続のプリンタ ポートを指定する文字列値。
'*             (Windows NT/Windows 2000 では無視されます)。
'*
'******************************************************************************
Sub AddWindowsPrinterConnection(PrinterName As String, Optional DriverName As String, Optional Port As String = "LPT1")
    Call mOrigin.AddWindowsPrinterConnection(PrinterName, DriverName, Port)
End Sub

'******************************************************************************
'* [概  要] EnumNetworkDrives メソッド
'* [詳  細] EnumNetworkDrives のラッパーメソッド。
'*          ネットワーク ドライブの現在の割り当てを返します。
'* [参  考] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364403%28v%3dmsdn.10%29>
'*
'* @return WScriptExWshCollection ネットワークドライブのマッピング情報
'*         ネットワーク ドライブのローカル名と対応する UNC 名という 2 つの
'*         アイテムをペアにした配列。
'*         偶数番目のアイテムは論理ドライブのローカル名。
'*         奇数番目のアイテムは対応する UNC 共有名。
'*         コレクションのアイテムはインデックス ゼロ (0) から始まります。
'******************************************************************************
Function EnumNetworkDrives() As WScriptExWshCollection
    Dim wrap As WScriptExWshCollection: Set wrap = New WScriptExWshCollection
    Set wrap.Origin = mOrigin.EnumNetworkDrives()
    Set EnumNetworkDrives = wrap
End Function

'******************************************************************************
'* [概  要] EnumPrinterConnections メソッド
'* [詳  細] EnumPrinterConnections のラッパーメソッド。
'*          ネットワーク プリンタの現在の割り当てを返します。
'* [参  考] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364404%28v%3dmsdn.10%29>
'*
'* @return WScriptExWshCollection プリンタのコレクションへの変数参照
'*         ネットワーク プリンタのローカル名と対応する UNC 名という 2 つの
'*         アイテムをペアにした配列。
'*         偶数番目のアイテムはプリンタ ポート。
'*         奇数番目のアイテムは対応するネットワーク プリンタ。
'*         コレクションのアイテムはインデックス ゼロ (0) から始まります。
'******************************************************************************
Function EnumPrinterConnections() As WScriptExWshCollection
    Dim wrap As WScriptExWshCollection: Set wrap = New WScriptExWshCollection
    Set wrap.Origin = mOrigin.EnumPrinterConnections()
    Set EnumPrinterConnections = wrap
End Function

'******************************************************************************
'* [概  要] MapNetworkDrive メソッド
'* [詳  細] MapNetworkDrive のラッパーメソッド。
'*          共有ネットワーク ドライブをコンピュータ システムに追加します。
'* [参  考] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364409%28v%3dmsdn.10%29>
'*
'* @param LocalName 割り当てるドライブのローカル名を示す文字列値
'* @param RemoteName 共有ドライブの UNC 名 (\\xxx\yyy) を示す文字列値
'* @param UpdateProfile 省略可能。マッピング情報を現在のユーザー プロファイル
'*                      に格納するかどうかを示すブール値。
'*                      TRUE を指定すると、ユーザー プロファイルに割り当てが
'*                      保存されます。既定値は FALSE です。
'* @param UserName 省略可能。ユーザー名を示す文字列値。
'* @param Password 省略可能。ユーザー パスワードを示す文字列値。
'*                 UserNameとPasswordは、ネットワーク ドライブの割り当て時に
'*                 現在のユーザー以外のユーザーのアカウント情報を使用する場合、
'*                 指定する必要があります。
'*
'******************************************************************************
Sub MapNetworkDrive(LocalName As String, RemoteName As String, Optional UpdateProfile, Optional UserName, Optional Password)
    If IsMissing(UpdateProfile) Then
        If IsMissing(UserName) Then
             If IsMissing(Password) Then
                Call mOrigin.MapNetworkDrive(LocalName, RemoteName)
            Else
                Call mOrigin.MapNetworkDrive(LocalName, RemoteName, , , Password)
            End If
        ElseIf IsMissing(Password) Then
            Call mOrigin.MapNetworkDrive(LocalName, RemoteName, , UserName)
        Else
            Call mOrigin.MapNetworkDrive(LocalName, RemoteName, , UserName, Password)
        End If
    Else
        If IsMissing(UserName) Then
             If IsMissing(Password) Then
                Call mOrigin.MapNetworkDrive(LocalName, RemoteName, UpdateProfile)
            Else
                Call mOrigin.MapNetworkDrive(LocalName, RemoteName, UpdateProfile, , Password)
            End If
        ElseIf IsMissing(Password) Then
            Call mOrigin.MapNetworkDrive(LocalName, RemoteName, UpdateProfile, UserName)
        Else
            Call mOrigin.MapNetworkDrive(LocalName, RemoteName, UpdateProfile, UserName, Password)
        End If
    End If
End Sub

'******************************************************************************
'* [概  要] RemoveNetworkDrive メソッド
'* [詳  細] RemoveNetworkDrive のラッパーメソッド。
'*          コンピュータ システムの共有ネットワーク ドライブ接続を解除します。
'* [参  考] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364419%28v%3dmsdn.10%29>
'*
'* @param Name 接続を解除する割り当てドライブを示す文字列値。
'*             ドライブの割り当て方法によって、ローカル名かリモート名を
'*             指定します。
'* @param Force 省略可能。ドライブの割り当てを強制的に解除するかどうかを示す
'*              ブール値。TRUE を指定すると、リソースが使用されているかどうか
'*              に関係なく、接続が解除されます。
'* @param UpdateProfile 省略可能。ユーザー プロファイルから割り当てを削除する
'*                      かどうかを示す文字列値。TRUE を指定すると、ユーザー
'*                      プロファイルから割り当てが削除されます。
'*                      既定値は FALSE です。
'*
'******************************************************************************
Sub RemoveNetworkDrive(Name As String, Optional Force, Optional UpdateProfile)
    If IsMissing(Force) Then
        If IsMissing(UpdateProfile) Then
            Call mOrigin.RemoveNetworkDrive(Name)
        Else
            Call mOrigin.RemoveNetworkDrive(Name, , UpdateProfile)
        End If
    ElseIf IsMissing(UpdateProfile) Then
        Call mOrigin.RemoveNetworkDrive(Name, Force)
    Else
        Call mOrigin.RemoveNetworkDrive(Name, Force, UpdateProfile)
    End If
End Sub

'******************************************************************************
'* [概  要] RemovePrinterConnection メソッド
'* [詳  細] RemovePrinterConnection のラッパーメソッド。
'*          コンピュータ システムの共有ネットワーク プリンタ接続を解除します。
'* [参  考] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364420%28v%3dmsdn.10%29>
'*
'* @param Name プリンタ識別名を示す文字列値。
'*             UNC 名 (\\xxx\yyy の形式) またはローカル名 (たとえば LPT1)
'*             を指定できます。
'* @param Force 省略可能。プリンタの割り当てを強制的に解除するかどうかを示す
'*              ブール値。TRUE を指定すると (既定値は FALSE)、ユーザーが接続
'*              されているかどうかに関係なく、接続が解除されます。
'* @param UpdateProfile 省略可能。ブール値です。TRUE を指定すると (既定値は
'*                      FALSE)、ユーザー プロファイルに変更内容が保存されます。
'*
'******************************************************************************
Sub RemovePrinterConnection(Name As String, Optional Force, Optional UpdateProfile)
    If IsMissing(Force) Then
        If IsMissing(UpdateProfile) Then
            Call mOrigin.RemovePrinterConnection(Name)
        Else
            Call mOrigin.RemovePrinterConnection(Name, , UpdateProfile)
        End If
    ElseIf IsMissing(UpdateProfile) Then
        Call mOrigin.RemovePrinterConnection(Name, Force)
    Else
        Call mOrigin.RemovePrinterConnection(Name, Force, UpdateProfile)
    End If
End Sub

'******************************************************************************
'* [概  要] SetDefaultPrinter メソッド
'* [詳  細] SetDefaultPrinter のラッパーメソッド。
'*          指定されたリモート プリンタを既定のプリンタとして設定します。
'* [参  考] <https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364424%28v%3dmsdn.10%29>
'*
'* @param Name リモート プリンタの UNC 名を示す文字列値
'*
'******************************************************************************
Sub SetDefaultPrinter(Name As String)
    Call mOrigin.SetDefaultPrinter(Name)
End Sub

'*-----------------------------------------------------------------------------
'* 拡張メソッド
'*-----------------------------------------------------------------------------

'******************************************************************************
'* [概  要] GetUnusedNetworkDriveName メソッド
'* [詳  細] 使用されていないネットワークドライブ名を取得する。
'*
'* @return ネットワークドライブ名
'*
'******************************************************************************
Public Function GetUnusedNetworkDriveName() As String
    Dim i As Long
    Dim colDrives As WScriptExWshCollection
  
    Dim vNdArr() As Variant
    Dim strTmpNd As Variant
    Dim flg As Boolean
    flg = True
    vNdArr = Array("E:", "F:", "G:", "H:", "I:", "J:", "K:", "L:", "M:", "N:", "O:", "P:", "Q:", "R:", "S:", "T:", "U:", "V:", "W:", "X:", "Y:", "Z:")
    
    Set colDrives = Me.EnumNetworkDrives
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
    GetUnusedNetworkDriveName = strTmpNd
End Function

