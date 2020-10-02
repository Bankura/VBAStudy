VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiWinspool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WindowAPI winspool クラス
'* [詳  細] WindowAPIのwinspool.drvの機能を提供する。
'*
'* [参  考] Win32API_PtrSafe.TXT
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal Flags As Long, ByVal Name As String, ByVal Level As Long, pPrinterEnum As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long

Private Declare PtrSafe Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As LongPtr, pDefault As PRINTER_DEFAULTS) As Long
Private Declare PtrSafe Function ResetPrinter Lib "winspool.drv" Alias "ResetPrinterA" (ByVal hPrinter As LongPtr, pDefault As PRINTER_DEFAULTS) As Long
Private Declare PtrSafe Function SetJob Lib "winspool.drv" Alias "SetJobA" (ByVal hPrinter As LongPtr, ByVal JobId As Long, ByVal Level As Long, pJob As Byte, ByVal Command As Long) As Long
Private Declare PtrSafe Function GetJob Lib "winspool.drv" Alias "GetJobA" (ByVal hPrinter As LongPtr, ByVal JobId As Long, ByVal Level As Long, pJob As Byte, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Private Declare PtrSafe Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" (ByVal hPrinter As LongPtr, ByVal FirstJob As Long, ByVal NoJobs As Long, ByVal Level As Long, pJob As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare PtrSafe Function AddPrinter Lib "winspool.drv" Alias "AddPrinterA" (ByVal pName As String, ByVal Level As Long, pPrinter As Any) As LongPtr

Private Declare PtrSafe Function AddPrinterDriver Lib "winspool.drv" Alias "AddPrinterDriverA" (ByVal pName As String, ByVal Level As Long, pDriverInfo As Any) As Long

Private Declare PtrSafe Function EnumPrinterDrivers Lib "winspool.drv" Alias "EnumPrinterDriversA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, pDriverInfo As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcRetruned As Long) As Long
Private Declare PtrSafe Function GetPrinterDriver Lib "winspool.drv" Alias "GetPrinterDriverA" (ByVal hPrinter As LongPtr, ByVal pEnvironment As String, ByVal Level As Long, pDriverInfo As Byte, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Private Declare PtrSafe Function GetPrinterDriverDirectory Lib "winspool.drv" Alias "GetPrinterDriverDirectoryA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, pDriverDirectory As Byte, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Private Declare PtrSafe Function DeletePrinterDriver Lib "winspool.drv" Alias "DeletePrinterDriverA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pDriverName As String) As Long

Private Declare PtrSafe Function AddPrintProcessor Lib "winspool.drv" Alias "AddPrintProcessorA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pPathName As String, ByVal pPrintProcessorName As String) As Long
Private Declare PtrSafe Function EnumPrintProcessors Lib "winspool.drv" Alias "EnumPrintProcessorsA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, pPrintProcessorInfo As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare PtrSafe Function GetPrintProcessorDirectory Lib "winspool.drv" Alias "GetPrintProcessorDirectoryA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, ByVal pPrintProcessorInfo As String, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Private Declare PtrSafe Function EnumPrintProcessorDatatypes Lib "winspool.drv" Alias "EnumPrintProcessorDatatypesA" (ByVal pName As String, ByVal pPrintProcessorName As String, ByVal Level As Long, pDatatypes As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcRetruned As Long) As Long
Private Declare PtrSafe Function DeletePrintProcessor Lib "winspool.drv" Alias "DeletePrintProcessorA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pPrintProcessorName As String) As Long

Private Declare PtrSafe Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As LongPtr, ByVal Level As Long, pDocInfo As Byte) As Long
Private Declare PtrSafe Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As LongPtr) As Long
Private Declare PtrSafe Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As LongPtr, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long
Private Declare PtrSafe Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As LongPtr) As Long
Private Declare PtrSafe Function AbortPrinter Lib "winspool.drv" (ByVal hPrinter As LongPtr) As Long
Private Declare PtrSafe Function ReadPrinter Lib "winspool.drv" (ByVal hPrinter As LongPtr, pBuf As Any, ByVal cdBuf As Long, pNoBytesRead As Long) As Long
Private Declare PtrSafe Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As LongPtr) As Long

Private Declare PtrSafe Function AddJob Lib "winspool.drv" Alias "AddJobA" (ByVal hPrinter As LongPtr, ByVal Level As Long, pData As Byte, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Private Declare PtrSafe Function ScheduleJob Lib "winspool.drv" (ByVal hPrinter As LongPtr, ByVal JobId As Long) As Long
Private Declare PtrSafe Function PrinterProperties Lib "winspool.drv" (ByVal hWnd As LongPtr, ByVal hPrinter As LongPtr) As Long
Private Declare PtrSafe Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As LongPtr, ByVal hPrinter As LongPtr, ByVal pDeviceName As String, pDevModeOutput As DEVMODE, pDevModeInput As DEVMODE, ByVal fMode As Long) As Long
Private Declare PtrSafe Function AdvancedDocumentProperties Lib "winspool.drv" Alias "AdvancedDocumentPropertiesA" (ByVal hWnd As LongPtr, ByVal hPrinter As LongPtr, ByVal pDeviceName As String, pDevModeOutput As DEVMODE, pDevModeInput As DEVMODE) As Long

Private Declare PtrSafe Function GetPrinterData Lib "winspool.drv" Alias "GetPrinterDataA" (ByVal hPrinter As LongPtr, ByVal pValueName As String, pType As Long, pData As Byte, ByVal nSize As Long, pcbNeeded As Long) As Long
Private Declare PtrSafe Function SetPrinterData Lib "winspool.drv" Alias "SetPrinterDataA" (ByVal hPrinter As LongPtr, ByVal pValueName As String, ByVal dwType As Long, pData As Byte, ByVal cbData As Long) As Long
Private Declare PtrSafe Function WaitForPrinterChange Lib "winspool.drv" (ByVal hPrinter As LongPtr, ByVal Flags As Long) As LongPtr

Private Declare PtrSafe Function PrinterMessageBox Lib "winspool.drv" Alias "PrinterMessageBoxA" (ByVal hPrinter As LongPtr, ByVal error As Long, ByVal hWnd As LongPtr, ByVal pText As String, ByVal pCaption As String, ByVal dwType As Long) As Long

Private Declare PtrSafe Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As LongPtr) As Long
Private Declare PtrSafe Function AddForm Lib "winspool.drv" Alias "AddFormA" (ByVal hPrinter As LongPtr, ByVal Level As Long, pForm As Byte) As Long
Private Declare PtrSafe Function DeleteForm Lib "winspool.drv" Alias "DeleteFormA" (ByVal hPrinter As LongPtr, ByVal pFormName As String) As Long
Private Declare PtrSafe Function GetForm Lib "winspool.drv" Alias "GetFormA" (ByVal hPrinter As LongPtr, ByVal pFormName As String, ByVal Level As Long, pForm As Byte, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Private Declare PtrSafe Function SetForm Lib "winspool.drv" Alias "SetFormA" (ByVal hPrinter As LongPtr, ByVal pFormName As String, ByVal Level As Long, pForm As Byte) As Long
Private Declare PtrSafe Function EnumForms Lib "winspool.drv" Alias "EnumFormsA" (ByVal hPrinter As LongPtr, ByVal Level As Long, pForm As Byte, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long

Private Declare PtrSafe Function EnumMonitors Lib "winspool.drv" Alias "EnumMonitorsA" (ByVal pName As String, ByVal Level As Long, pMonitors As Byte, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare PtrSafe Function AddMonitor Lib "winspool.drv" Alias "AddMonitorA" (ByVal pName As String, ByVal Level As Long, pMonitors As Byte) As Long
Private Declare PtrSafe Function DeleteMonitor Lib "winspool.drv" Alias "DeleteMonitorA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pMonitorName As String) As Long

Private Declare PtrSafe Function EnumPorts Lib "winspool.drv" Alias "EnumPortsA" (ByVal pName As String, ByVal Level As Long, ByVal lpbPorts As LongPtr, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare PtrSafe Function AddPort Lib "winspool.drv" Alias "AddPortA" (ByVal pName As String, ByVal hWnd As LongPtr, ByVal pMonitorName As String) As Long
Private Declare PtrSafe Function ConfigurePort Lib "winspool.drv" Alias "ConfigurePortA" (ByVal pName As String, ByVal hWnd As LongPtr, ByVal pPortName As String) As Long
Private Declare PtrSafe Function DeletePort Lib "winspool.drv" Alias "DeletePortA" (ByVal pName As String, ByVal hWnd As LongPtr, ByVal pPortName As String) As Long

Private Declare PtrSafe Function AddPrinterConnection Lib "winspool.drv" Alias "AddPrinterConnectionA" (ByVal pName As String) As Long

Private Declare PtrSafe Function DeletePrinterConnection Lib "winspool.drv" Alias "DeletePrinterConnectionA" (ByVal pName As String) As Long
Private Declare PtrSafe Function ConnectToPrinterDlg Lib "winspool.drv" (ByVal hWnd As LongPtr, ByVal Flags As Long) As LongPtr

Private Declare PtrSafe Function AddPrintProvidor Lib "winspool.drv" Alias "AddPrintProvidorA" (ByVal pName As String, ByVal Level As Long, pProvidorInfo As Byte) As Long
Private Declare PtrSafe Function DeletePrintProvidor Lib "winspool.drv" Alias "DeletePrintProvidorA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pPrintProvidorName As String) As Long

Private Declare PtrSafe Function DeletePrinter Lib "winspool.drv" (ByVal hPrinter As LongPtr) As Long
Private Declare PtrSafe Function FindClosePrinterChangeNotification Lib "winspool.drv" (ByVal hChange As LongPtr) As Long
Private Declare PtrSafe Function FindFirstPrinterChangeNotification Lib "winspool.drv" (ByVal hPrinter As LongPtr, ByVal fdwFlags As Long, ByVal fdwOptions As Long, ByVal pPrinterNotifyOptions As String) As LongPtr
Private Declare PtrSafe Function FindNextPrinterChangeNotification Lib "winspool.drv" (ByVal hChange As LongPtr, pdwChange As Long, ByVal pvReserved As String, ByVal ppPrinterNotifyInfo As LongPtr) As Long
Private Declare PtrSafe Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As LongPtr, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Private Declare PtrSafe Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As LongPtr, ByVal Level As Long, pPrinter As Byte, ByVal Command As Long) As Long


