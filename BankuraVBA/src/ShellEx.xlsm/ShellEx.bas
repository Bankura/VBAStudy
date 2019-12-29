Attribute VB_Name = "ShellEx"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] Shellラップ・拡張モジュール
'* [詳  細] ShellのWrapperとしての機能を提供する他、Scriptingを使用した
'*          ユーティリティを提供する。
'*          ラップするShellライブラリは以下のものとする。
'*              [name] Microsoft Shell Controls And Automation
'*              [dll] C:\Windows\SysWOW64\shell32.dll
'* [参  考]
'*  <xxxxxxxxxxxxxxxxxxxxxxx>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* Enum定義
'******************************************************************************

'*-----------------------------------------------------------------------------
'* Constants for Folder2.OfflineStatus
'*
'*-----------------------------------------------------------------------------
Public Enum OfflineFolderStatus
    OFS_DIRTYCACHE = 3  'Server is online with unmerged changes
    OFS_INACTIVE = -1   'Offline caching not available for this folder
    OFS_OFFLINE = 1     'Server is offline
    OFS_ONLINE = 0      'Server is online
    OFS_SERVERBACK = 2  'Server is offline but is reachable
End Enum

'*-----------------------------------------------------------------------------
'* Constants for ViewOptions
'*
'*-----------------------------------------------------------------------------
Public Enum ShellFolderViewOptions
    SFVVO_DESKTOPHTML = 512          'Is Desktop HTML enabled
    SFVVO_DOUBLECLICKINWEBVIEW = 128 'User needs to double click in web View
    SFVVO_SHOWALLOBJECTS = 1         'Show All Objects
    SFVVO_SHOWCOMPCOLOR = 8          'Color encode Compressed files
    SFVVO_SHOWSYSFILES = 32          'Show System Files
    SFVVO_WIN95CLASSIC = 64          'Use Windows 95 UI settings
End Enum

'*-----------------------------------------------------------------------------
'* Constants for Special Folders for open/Explore
'*
'*-----------------------------------------------------------------------------
Public Enum ShellSpecialFolderConstants
    ssfALTSTARTUP = 29            'Special Folder ALTSTARTUP
    ssfAPPDATA = 26               'Special Folder APPDATA
    ssfBITBUCKET = 10             'Special Folder BITBUCKET
    ssfCOMMONALTSTARTUP = 30      'Special Folder COMMON ALTSTARTUP
    ssfCOMMONAPPDATA = 35         'Special Folder COMMON APPDATA
    ssfCOMMONDESKTOPDIR = 25      'Special Folder COMMON DESKTOPDIR
    ssfCOMMONFAVORITES = 31       'Special Folder COMMON FAVORITES
    ssfCOMMONPROGRAMS = 23        'Special Folder COMMON PROGRAMS
    ssfCOMMONSTARTMENU = 22       'Special Folder COMMON STARTMENU
    ssfCOMMONSTARTUP = 24         'Special Folder COMMON STARTUP
    ssfCONTROLS = 3               'Special Folder CONTROLS
    ssfCOOKIES = 33               'Special Folder COOKIES
    ssfDESKTOP = 0                'Special Folder DESKTOP
    ssfDESKTOPDIRECTORY = 16      'Special Folder DESKTOPDIRECTORY
    ssfDRIVES = 17                'Special Folder DRIVES
    ssfFAVORITES = 6              'Special Folder FAVORITES
    ssfFONTS = 20                 'Special Folder FONTS
    ssfHISTORY = 34               'Special Folder HISTORY
    ssfINTERNETCACHE = 32         'Special Folder INTERNET CACHE
    ssfLOCALAPPDATA = 28          'Special Folder LOCAL APPDATA
    ssfMYPICTURES = 39            'Special Folder MYPICTURES
    ssfNETHOOD = 19               'Special Folder NETHOOD
    ssfNETWORK = 18               'Special Folder NETWORK
    ssfPERSONAL = 5               'Special Folder PERSONAL
    ssfPRINTERS = 4               'Special Folder PRINTERS
    ssfPRINTHOOD = 27             'Special Folder PRINTHOOD
    ssfPROFILE = 40               'Special Folder PROFILE
    ssfPROGRAMFILES = 38          'Special Folder PROGRAM FILES
    ssfPROGRAMFILESx86 = 48       'Special Folder PROGRAM FILESx86
    ssfPROGRAMS = 2               'Special Folder PROGRAMS
    ssfRECENT = 8                 'Special Folder RECENT
    ssfSENDTO = 9                 'Special Folder SENDTO
    ssfSTARTMENU = 11             'Special Folder STARTMENU
    ssfSTARTUP = 7                'Special Folder STARTUP
    ssfSYSTEM = 37                'Special Folder SYSTEM
    ssfSYSTEMx86 = 41             'Special Folder SYSTEMx86
    ssfTEMPLATES = 21             'Special Folder TEMPLATES
    ssfWINDOWS = 36               'Special Folder WINDOWS
End Enum


'******************************************************************************
'* メソッド定義
'******************************************************************************


