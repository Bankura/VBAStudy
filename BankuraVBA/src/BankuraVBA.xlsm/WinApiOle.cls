VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiOle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [�@�\��] WindowAPI ole �N���X
'* [��  ��] WindowAPI�́uole32.dll�v�̋@�\��񋟂���B
'*
'* [�Q  �l]
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI��`
'******************************************************************************
Private Declare PtrSafe Function IIDFromString Lib "ole32.dll" (lpsz As Any, lpiid As Any) As Long

'******************************************************************************
'* [�T  �v] WinApi_IIDFromString
'* [��  ��] WindowsAPI�u IIDFromString �v���b�p�[���\�b�h�B
'*          Lib "ole32"
'*
'******************************************************************************
Friend Function WinApi_IIDFromString(lpsz As Byte, lpiid As LongPtr) As Long
    WinApi_IIDFromString = IIDFromString(lpsz, lpiid)
End Function
