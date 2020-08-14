Attribute VB_Name = "ScriptingEx"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Scriptingラップ・拡張モジュール
'* [詳  細] ScriptingのWrapperとしての機能を提供する他、Scriptingを使用した
'*          ユーティリティを提供する。
'*          ラップするScriptingライブラリは以下のものとする。
'*              [name] Microsoft Scripting Runtime
'*              [dll] C:\Windows\System32\scrrun.dll
'* [参  考]
'*  <https://docs.microsoft.com/ja-jp/office/vba/language/reference/objects-visual-basic-for-applications>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* Enum定義
'******************************************************************************

'*-----------------------------------------------------------------------------
'* CompareMethod
'*
'*-----------------------------------------------------------------------------
Public Enum CompareMethod
    BinaryCompare = 0
    DatabaseCompare = 2
    TextCompare = 1
End Enum

'*-----------------------------------------------------------------------------
'* DriveTypeConst
'* ドライブの種類を表す
'*-----------------------------------------------------------------------------
Public Enum DriveTypeConst
    CDRom = 4       'CD-ROM ドライブ
    Fixed = 2       'ハードディスク
    RamDisk = 5     'RAM ディスク
    Remote = 3      'ネットワークドライブ
    Removable = 1   'リムーバブルディスク
    UnknownType = 0 '不明
End Enum

'*-----------------------------------------------------------------------------
'* FileAttribute
'*
'*-----------------------------------------------------------------------------
Public Enum FileAttribute
    Alias = 1024
    Archive = 32
    Compressed = 2048
    Directory = 16
    Hidden = 2
    Normal = 0
    ReadOnly = 1
    System = 4
    Volume = 8
End Enum

'*-----------------------------------------------------------------------------
'* IOMode
'*
'*-----------------------------------------------------------------------------
Public Enum IOMode
    ForAppending = 8
    ForReading = 1
    ForWriting = 2
End Enum

'*-----------------------------------------------------------------------------
'* SpecialFolderConst
'*
'*-----------------------------------------------------------------------------
Public Enum SpecialFolderConst
    SystemFolder = 1
    TemporaryFolder = 2
    WindowsFolder = 0
End Enum

'*-----------------------------------------------------------------------------
'* StandardStreamTypes
'*
'*-----------------------------------------------------------------------------
Public Enum StandardStreamTypes
    StdErr = 2
    StdIn = 0
    StdOut = 1
End Enum

'*-----------------------------------------------------------------------------
'* Tristate
'*
'*-----------------------------------------------------------------------------
Public Enum Tristate
    TristateFalse = 0
    TristateMixed = -2
    TristateTrue = -1
    TristateUseDefault = -2
End Enum

'******************************************************************************
'* メソッド定義
'******************************************************************************


