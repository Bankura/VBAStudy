VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiGdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WindowAPI GDI クラス
'* [詳  細] WindowAPIのgdi32.dllの機能を提供する。
'*
'* [参  考] Win32API_PtrSafe.TXT
'*
'* @author Bankura
'* Copyright (c) 2019-2020 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


'******************************************************************************
'* WindowsAPI定義
'******************************************************************************
Private Declare PtrSafe Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long

Private Declare PtrSafe Function AnimatePalette Lib "gdi32" Alias "AnimatePaletteA" (ByVal hPalette As LongPtr, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteColors As PALETTEENTRY) As Long
Private Declare PtrSafe Function Arc Lib "gdi32" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function CancelDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function Chord Lib "gdi32" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare PtrSafe Function CloseMetaFile Lib "gdi32" (ByVal hMF As LongPtr) As LongPtr
Private Declare PtrSafe Function CombineRgn Lib "gdi32" (ByVal hDestRgn As LongPtr, ByVal hSrcRgn1 As LongPtr, ByVal hSrcRgn2 As LongPtr, ByVal nCombineMode As Long) As Long
Private Declare PtrSafe Function CopyMetaFile Lib "gdi32" Alias "CopyMetaFileA" (ByVal hMF As LongPtr, ByVal lpFileName As String) As LongPtr

Private Declare PtrSafe Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As LongPtr
Private Declare PtrSafe Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As LongPtr
Private Declare PtrSafe Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As LongPtr
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
Private Declare PtrSafe Function CreateDiscardableBitmap Lib "gdi32" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr

Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As LongPtr

Private Declare PtrSafe Function CreateDIBitmap Lib "gdi32" (ByVal hdc As LongPtr, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO, ByVal wUsage As Long) As LongPtr
Private Declare PtrSafe Function CreateDIBPatternBrush Lib "gdi32" (ByVal hPackedDIB As LongPtr, ByVal wUsage As Long) As LongPtr
Private Declare PtrSafe Function CreateDIBPatternBrushPt Lib "gdi32" (lpPackedDIB As Any, ByVal iUsage As Long) As LongPtr
Private Declare PtrSafe Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As LongPtr
Private Declare PtrSafe Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As LongPtr

Private Declare PtrSafe Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As LongPtr
Private Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal e As Long, ByVal o As Long, ByVal w As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal q As Long, ByVal PAF As Long, ByVal f As String) As LongPtr

Private Declare PtrSafe Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As LongPtr
Private Declare PtrSafe Function CreateIC Lib "gdi32" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As LongPtr

Private Declare PtrSafe Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As LongPtr

Private Declare PtrSafe Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As LongPtr
Private Declare PtrSafe Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As LongPtr
Private Declare PtrSafe Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As LongPtr
Private Declare PtrSafe Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As LongPtr
Private Declare PtrSafe Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As LongPtr
Private Declare PtrSafe Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As LongPtr
Private Declare PtrSafe Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As LongPtr
Private Declare PtrSafe Function CreateScalableFontResource Lib "gdi32" Alias "CreateScalableFontResourceA" (ByVal fHidden As Long, ByVal lpszResourceFile As String, ByVal lpszFontFile As String, ByVal lpszCurrentPath As String) As Long
Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr

Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function DeleteMetaFile Lib "gdi32" (ByVal hMF As LongPtr) As Long

Private Declare PtrSafe Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As DEVMODE) As Long

Private Declare PtrSafe Function Ellipse Lib "gdi32" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare PtrSafe Function EqualRgn Lib "gdi32" (ByVal hSrcRgn1 As LongPtr, ByVal hSrcRgn2 As LongPtr) As Long
Private Declare PtrSafe Function Escape Lib "gdi32" (ByVal hdc As LongPtr, ByVal nEscape As Long, ByVal nCount As Long, ByVal lpInData As String, lpOutData As Any) As Long
Private Declare PtrSafe Function ExtEscape Lib "gdi32" (ByVal hdc As LongPtr, ByVal nEscape As Long, ByVal cbInput As Long, ByVal lpszInData As String, ByVal cbOutput As Long, ByVal lpszOutData As String) As Long
Private Declare PtrSafe Function DrawEscape Lib "gdi32" (ByVal hdc As LongPtr, ByVal nEscape As Long, ByVal cbInput As Long, ByVal lpszInData As String) As Long
Private Declare PtrSafe Function ExcludeClipRect Lib "gdi32" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare PtrSafe Function ExtCreateRegion Lib "gdi32" (lpXform As xform, ByVal nCount As Long, lpRgnData As RgnData) As LongPtr
Private Declare PtrSafe Function ExtFloodFill Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Private Declare PtrSafe Function FillRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr, ByVal hBrush As LongPtr) As Long
Private Declare PtrSafe Function FrameRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr, ByVal hBrush As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare PtrSafe Function FloodFill Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare PtrSafe Function GetROP2 Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetAspectRatioFilterEx Lib "gdi32" (ByVal hdc As LongPtr, lpAspectRatio As size) As Long
Private Declare PtrSafe Function GetBkColor Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetBkMode Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As LongPtr, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare PtrSafe Function GetBitmapDimensionEx Lib "gdi32" (ByVal hBitmap As LongPtr, lpDimension As size) As Long
Private Declare PtrSafe Function GetBoundsRect Lib "gdi32" (ByVal hdc As LongPtr, lprcBounds As RECT, ByVal Flags As Long) As Long

Private Declare PtrSafe Function GetCharWidth Lib "gdi32" Alias "GetCharWidthA" (ByVal hdc As LongPtr, ByVal wFirstChar As Long, ByVal wLastChar As Long, lpBuffer As Long) As Long

Private Declare PtrSafe Function GetCharWidth32 Lib "gdi32" Alias "GetCharWidth32A" (ByVal hdc As LongPtr, ByVal iFirstChar As Long, ByVal iLastChar As Long, lpBuffer As Long) As Long
Private Declare PtrSafe Function GetCharWidthFloat Lib "gdi32" Alias "GetCharWidthFloatA" (ByVal hdc As LongPtr, ByVal iFirstChar As Long, ByVal iLastChar As Long, pxBuffer As Double) As Long

Private Declare PtrSafe Function GetCharABCWidths Lib "gdi32" Alias "GetCharABCWidthsA" (ByVal hdc As LongPtr, ByVal uFirstChar As Long, ByVal uLastChar As Long, lpabc As ABC) As Long
Private Declare PtrSafe Function GetCharABCWidthsFloat Lib "gdi32" Alias "GetCharABCWidthsFloatA" (ByVal hdc As LongPtr, ByVal iFirstChar As Long, ByVal iLastChar As Long, lpABCF As ABCFLOAT) As Long

Private Declare PtrSafe Function GetClipBox Lib "gdi32" (ByVal hdc As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function GetClipRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr) As Long
Private Declare PtrSafe Function GetMetaRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr) As Long
Private Declare PtrSafe Function GetCurrentObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal uObjectType As Long) As LongPtr
Private Declare PtrSafe Function GetCurrentPositionEx Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDIBits Lib "gdi32" (ByVal aHDC As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare PtrSafe Function GetFontData Lib "gdi32" Alias "GetFontDataA" (ByVal hdc As LongPtr, ByVal dwTable As Long, ByVal dwOffset As Long, lpvBuffer As Any, ByVal cbData As Long) As Long
Private Declare PtrSafe Function GetGlyphOutline Lib "gdi32" Alias "GetGlyphOutlineA" (ByVal hdc As LongPtr, ByVal uChar As Long, ByVal fuFormat As Long, lpgm As GLYPHMETRICS, ByVal cbBuffer As Long, lpBuffer As Any, lpmat2 As MAT2) As Long
Private Declare PtrSafe Function GetGraphicsMode Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetMapMode Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetMetaFileBitsEx Lib "gdi32" (ByVal hMF As LongPtr, ByVal nSize As Long, lpvData As Any) As Long
Private Declare PtrSafe Function GetMetaFile Lib "gdi32" Alias "GetMetaFileA" (ByVal lpFileName As String) As LongPtr
Private Declare PtrSafe Function GetNearestColor Lib "gdi32" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
Private Declare PtrSafe Function GetNearestPaletteIndex Lib "gdi32" (ByVal hPalette As LongPtr, ByVal crColor As Long) As Long
Private Declare PtrSafe Function GetObjectType Lib "gdi32" (ByVal hgdiobj As LongPtr) As Long

Private Declare PtrSafe Function GetOutlineTextMetrics Lib "gdi32" Alias "GetOutlineTextMetricsA" (ByVal hdc As LongPtr, ByVal cbData As Long, lpotm As OUTLINETEXTMETRIC) As Long

Private Declare PtrSafe Function GetPaletteEntries Lib "gdi32" (ByVal hPalette As LongPtr, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function GetPolyFillMode Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetRasterizerCaps Lib "gdi32" (lpraststat As RASTERIZER_STATUS, ByVal cb As Long) As Long
Private Declare PtrSafe Function GetRegionData Lib "gdi32" Alias "GetRegionDataA" (ByVal hRgn As LongPtr, ByVal dwCount As Long, lpRgnData As RgnData) As Long
Private Declare PtrSafe Function GetRgnBox Lib "gdi32" (ByVal hRgn As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function GetStretchBltMode Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As LongPtr, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare PtrSafe Function GetSystemPaletteUse Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetTextCharacterExtra Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetTextAlign Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetTextColor Lib "gdi32" (ByVal hdc As LongPtr) As Long

Private Declare PtrSafe Function GetTextExtentPoint Lib "gdi32" Alias "GetTextExtentPointA" (ByVal hdc As LongPtr, ByVal lpszString As String, ByVal cbString As Long, lpSize As size) As Long
Private Declare PtrSafe Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As LongPtr, ByVal lpsz As String, ByVal cbString As Long, lpSize As size) As Long
Private Declare PtrSafe Function GetTextExtentExPoint Lib "gdi32" Alias "GetTextExtentExPointA" (ByVal hdc As LongPtr, ByVal lpszStr As String, ByVal cchString As Long, ByVal nMaxExtent As Long, lpnFit As Long, alpDx As Long, lpSize As size) As Long

Private Declare PtrSafe Function GetViewportExtEx Lib "gdi32" (ByVal hdc As LongPtr, lpSize As size) As Long
Private Declare PtrSafe Function GetViewportOrgEx Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function GetWindowExtEx Lib "gdi32" (ByVal hdc As LongPtr, lpSize As size) As Long
Private Declare PtrSafe Function GetWindowOrgEx Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI) As Long

Private Declare PtrSafe Function IntersectClipRect Lib "gdi32" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare PtrSafe Function InvertRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr) As Long
Private Declare PtrSafe Function LineTo Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function MaskBlt Lib "gdi32" (ByVal hDCDest As LongPtr, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As LongPtr, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal hbmMask As LongPtr, ByVal xMask As Long, ByVal yMask As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function PlgBlt Lib "gdi32" (ByVal hDCDest As LongPtr, lpPoint As POINTAPI, ByVal hdcSrc As LongPtr, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hbmMask As LongPtr, ByVal xMask As Long, ByVal yMask As Long) As Long

Private Declare PtrSafe Function OffsetClipRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function OffsetRgn Lib "gdi32" (ByVal hRgn As LongPtr, ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function PatBlt Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function Pie Lib "gdi32" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare PtrSafe Function PlayMetaFile Lib "gdi32" (ByVal hdc As LongPtr, ByVal hMF As LongPtr) As Long
Private Declare PtrSafe Function PaintRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr) As Long
Private Declare PtrSafe Function PolyPolygon Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
Private Declare PtrSafe Function PtInRegion Lib "gdi32" (ByVal hRgn As LongPtr, ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function PtVisible Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function RectVisible Lib "gdi32" (ByVal hdc As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function RectInRegion Lib "gdi32" (ByVal hRgn As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function Rectangle Lib "gdi32" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare PtrSafe Function RestoreDC Lib "gdi32" (ByVal hdc As LongPtr, ByVal nSavedDC As Long) As Long
Private Declare PtrSafe Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hdc As LongPtr, lpInitData As DEVMODE) As LongPtr
Private Declare PtrSafe Function RealizePalette Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long
Private Declare PtrSafe Function RoundRect Lib "gdi32" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare PtrSafe Function ResizePalette Lib "gdi32" (ByVal hPalette As LongPtr, ByVal nNumEntries As Long) As Long

Private Declare PtrSafe Function SaveDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function SelectClipRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr) As Long
Private Declare PtrSafe Function ExtSelectClipRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr, ByVal fnMode As Long) As Long
Private Declare PtrSafe Function SetMetaRgn Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function SelectPalette Lib "gdi32" (ByVal hdc As LongPtr, ByVal hPalette As LongPtr, ByVal bForceBackground As Long) As LongPtr
Private Declare PtrSafe Function SetBkColor Lib "gdi32" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
Private Declare PtrSafe Function SetBkMode Lib "gdi32" (ByVal hdc As LongPtr, ByVal nBkMode As Long) As Long
Private Declare PtrSafe Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As LongPtr, ByVal dwCount As Long, lpBits As Any) As Long

Private Declare PtrSafe Function SetBoundsRect Lib "gdi32" (ByVal hdc As LongPtr, lprcBounds As RECT, ByVal Flags As Long) As Long
Private Declare PtrSafe Function SetDIBits Lib "gdi32" (ByVal hdc As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare PtrSafe Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare PtrSafe Function SetMapperFlags Lib "gdi32" (ByVal hdc As LongPtr, ByVal dwFlag As Long) As Long
Private Declare PtrSafe Function SetGraphicsMode Lib "gdi32" (ByVal hdc As LongPtr, ByVal iMode As Long) As Long
Private Declare PtrSafe Function SetMapMode Lib "gdi32" (ByVal hdc As LongPtr, ByVal nMapMode As Long) As Long
Private Declare PtrSafe Function SetMetaFileBitsEx Lib "gdi32" (ByVal nSize As Long, lpData As Byte) As LongPtr
Private Declare PtrSafe Function SetPaletteEntries Lib "gdi32" (ByVal hPalette As LongPtr, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare PtrSafe Function SetPixel Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare PtrSafe Function SetPixelV Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare PtrSafe Function SetPolyFillMode Lib "gdi32" (ByVal hdc As LongPtr, ByVal nPolyFillMode As Long) As Long
Private Declare PtrSafe Function StretchBlt Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function SetRectRgn Lib "gdi32" (ByVal hRgn As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare PtrSafe Function StretchDIBits Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function SetROP2 Lib "gdi32" (ByVal hdc As LongPtr, ByVal nDrawMode As Long) As Long
Private Declare PtrSafe Function SetStretchBltMode Lib "gdi32" (ByVal hdc As LongPtr, ByVal nStretchMode As Long) As Long
Private Declare PtrSafe Function SetSystemPaletteUse Lib "gdi32" (ByVal hdc As LongPtr, ByVal wUsage As Long) As Long
Private Declare PtrSafe Function SetTextCharacterExtra Lib "gdi32" Alias "SetTextCharacterExtraA" (ByVal hdc As LongPtr, ByVal nCharExtra As Long) As Long
Private Declare PtrSafe Function SetTextColor Lib "gdi32" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
Private Declare PtrSafe Function SetTextAlign Lib "gdi32" (ByVal hdc As LongPtr, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function SetTextJustification Lib "gdi32" (ByVal hdc As LongPtr, ByVal nBreakExtra As Long, ByVal nBreakCount As Long) As Long
Private Declare PtrSafe Function UpdateColors Lib "gdi32" (ByVal hdc As LongPtr) As Long

Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr

Private Declare PtrSafe Function PlayMetaFileRecord Lib "gdi32" (ByVal hdc As LongPtr, lpHandletable As HANDLETABLE, lpMetaRecord As METARECORD, ByVal nHandles As Long) As Long

Private Declare PtrSafe Function CloseEnhMetaFile Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As LongPtr, ByVal lpszFile As String) As LongPtr
Private Declare PtrSafe Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileA" (ByVal hdcRef As LongPtr, ByVal lpFileName As String, lpRect As RECT, ByVal lpDescription As String) As LongPtr
Private Declare PtrSafe Function DeleteEnhMetaFile Lib "gdi32" (ByVal hemf As LongPtr) As Long
Private Declare PtrSafe Function EnumEnhMetaFile Lib "gdi32" (ByVal hdc As LongPtr, ByVal hemf As LongPtr, ByVal lpEnhMetaFunc As LongPtr, lpData As Any, lpRect As RECT) As Long
Private Declare PtrSafe Function GetEnhMetaFile Lib "gdi32" Alias "GetEnhMetaFileA" (ByVal lpszMetaFile As String) As LongPtr
Private Declare PtrSafe Function GetEnhMetaFileBits Lib "gdi32" (ByVal hemf As LongPtr, ByVal cbBuffer As Long, lpbBuffer As Byte) As Long
Private Declare PtrSafe Function GetEnhMetaFileDescription Lib "gdi32" Alias "GetEnhMetaFileDescriptionA" (ByVal hemf As LongPtr, ByVal cchBuffer As Long, ByVal lpszDescription As String) As Long
Private Declare PtrSafe Function GetEnhMetaFileHeader Lib "gdi32" (ByVal hemf As LongPtr, ByVal cbBuffer As Long, lpemh As ENHMETAHEADER) As Long
Private Declare PtrSafe Function GetEnhMetaFilePaletteEntries Lib "gdi32" (ByVal hemf As LongPtr, ByVal cEntries As Long, lppe As PALETTEENTRY) As Long
Private Declare PtrSafe Function GetWinMetaFileBits Lib "gdi32" (ByVal hemf As LongPtr, ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal fnMapMode As Long, ByVal hdcRef As LongPtr) As Long
Private Declare PtrSafe Function PlayEnhMetaFile Lib "gdi32" (ByVal hdc As LongPtr, ByVal hemf As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function PlayEnhMetaFileRecord Lib "gdi32" (ByVal hdc As LongPtr, lpHandletable As HANDLETABLE, lpEnhMetaRecord As ENHMETARECORD, ByVal nHandles As Long) As Long
Private Declare PtrSafe Function SetEnhMetaFileBits Lib "gdi32" (ByVal cbBuffer As Long, lpData As Byte) As LongPtr
Private Declare PtrSafe Function SetWinMetaFileBits Lib "gdi32" (ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal hdcRef As LongPtr, lpmfp As METAFILEPICT) As LongPtr
Private Declare PtrSafe Function GdiComment Lib "gdi32" (ByVal hdc As LongPtr, ByVal cbSize As Long, lpData As Byte) As Long

Private Declare PtrSafe Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As LongPtr, lpMetrics As TEXTMETRIC) As Long

Private Declare PtrSafe Function AngleArc Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal dwRadius As Long, ByVal eStartAngle As Double, ByVal eSweepAngle As Double) As Long
Private Declare PtrSafe Function PolyPolyline Lib "gdi32" (ByVal hdc As LongPtr, lppt As POINTAPI, lpdwPolyPoints As Long, ByVal cCount As Long) As Long
Private Declare PtrSafe Function GetWorldTransform Lib "gdi32" (ByVal hdc As LongPtr, lpXform As xform) As Long
Private Declare PtrSafe Function SetWorldTransform Lib "gdi32" (ByVal hdc As LongPtr, lpXform As xform) As Long
Private Declare PtrSafe Function ModifyWorldTransform Lib "gdi32" (ByVal hdc As LongPtr, lpXform As xform, ByVal iMode As Long) As Long
Private Declare PtrSafe Function CombineTransform Lib "gdi32" (lpxformResult As xform, lpxform1 As xform, lpxform2 As xform) As Long

Private Declare PtrSafe Function SetColorAdjustment Lib "gdi32" (ByVal hdc As LongPtr, lpca As ColorAdjustment) As Long
Private Declare PtrSafe Function GetColorAdjustment Lib "gdi32" (ByVal hdc As LongPtr, lpca As ColorAdjustment) As Long
Private Declare PtrSafe Function CreateHalftonePalette Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr

Private Declare PtrSafe Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hdc As LongPtr, lpdi As DOCINFO) As Long
Private Declare PtrSafe Function StartPage Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function EndPage Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function EndDoc Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function AbortDoc Lib "gdi32" (ByVal hdc As LongPtr) As Long

Private Declare PtrSafe Function AbortPath Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function ArcTo Lib "gdi32" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare PtrSafe Function BeginPath Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function CloseFigure Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function EndPath Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function FillPath Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function FlattenPath Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetPath Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI, lpTypes As Byte, ByVal nSize As Long) As Long
Private Declare PtrSafe Function PathToRegion Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function PolyDraw Lib "gdi32" (ByVal hdc As LongPtr, lppt As POINTAPI, lpbTypes As Byte, ByVal cCount As Long) As Long
Private Declare PtrSafe Function SelectClipPath Lib "gdi32" (ByVal hdc As LongPtr, ByVal iMode As Long) As Long
Private Declare PtrSafe Function SetArcDirection Lib "gdi32" (ByVal hdc As LongPtr, ByVal ArcDirection As Long) As Long
Private Declare PtrSafe Function SetMiterLimit Lib "gdi32" (ByVal hdc As LongPtr, ByVal eNewLimit As Double, peOldLimit As Double) As Long
Private Declare PtrSafe Function StrokeAndFillPath Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function StrokePath Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function WidenPath Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, lplb As LOGBRUSH, ByVal dwStyleCount As Long, lpStyle As Long) As LongPtr
Private Declare PtrSafe Function GetMiterLimit Lib "gdi32" (ByVal hdc As LongPtr, peLimit As Double) As Long
Private Declare PtrSafe Function GetArcDirection Lib "gdi32" (ByVal hdc As LongPtr) As Long

Private Declare PtrSafe Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
Private Declare PtrSafe Function MoveToEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare PtrSafe Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Private Declare PtrSafe Function PolyTextOut Lib "gdi32" Alias "PolyTextOutA" (ByVal hdc As LongPtr, pptxt As POLYTEXT, ByVal cStrings As Long) As Long

Private Declare PtrSafe Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As LongPtr
Private Declare PtrSafe Function DPtoLP Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare PtrSafe Function LPtoDP Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare PtrSafe Function Polyline Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare PtrSafe Function Polygon Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Private Declare PtrSafe Function PolyBezier Lib "gdi32" (ByVal hdc As LongPtr, lppt As POINTAPI, ByVal cPoints As Long) As Long
Private Declare PtrSafe Function PolyBezierTo Lib "gdi32" (ByVal hdc As LongPtr, lppt As POINTAPI, ByVal cCount As Long) As Long
Private Declare PtrSafe Function PolylineTo Lib "gdi32" (ByVal hdc As LongPtr, lppt As POINTAPI, ByVal cCount As Long) As Long

Private Declare PtrSafe Function SetViewportExtEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpSize As size) As Long
Private Declare PtrSafe Function SetViewportOrgEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function SetWindowOrgEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function SetWindowExtEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpSize As size) As Long

Private Declare PtrSafe Function OffsetViewportOrgEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function OffsetWindowOrgEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function ScaleWindowExtEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal nXnum As Long, ByVal nXdenom As Long, ByVal nYnum As Long, ByVal nYdenom As Long, lpSize As size) As Long
Private Declare PtrSafe Function ScaleViewportExtEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal nXnum As Long, ByVal nXdenom As Long, ByVal nYnum As Long, ByVal nYdenom As Long, lpSize As size) As Long
Private Declare PtrSafe Function SetBitmapDimensionEx Lib "gdi32" (ByVal hbm As LongPtr, ByVal nX As Long, ByVal nY As Long, lpSize As size) As Long
Private Declare PtrSafe Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long

Private Declare PtrSafe Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As LongPtr, ByVal nCount As Long, ByVal lpFacename As String) As Long

Private Declare PtrSafe Function GetKerningPairs Lib "gdi32" Alias "GetKerningPairsA" (ByVal hdc As LongPtr, ByVal cPairs As Long, lpkrnpair As KERNINGPAIR) As Long

Private Declare PtrSafe Function GetDCOrgEx Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function UnrealizeObject Lib "gdi32" (ByVal hObject As LongPtr) As Long

Private Declare PtrSafe Function GdiFlush Lib "gdi32" () As Long
Private Declare PtrSafe Function GdiSetBatchLimit Lib "gdi32" (ByVal dwLimit As Long) As Long
Private Declare PtrSafe Function GdiGetBatchLimit Lib "gdi32" () As Long

Private Declare PtrSafe Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hdc As LongPtr, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" (ByVal hdc As LongPtr, lpLogFont As LOGFONT, ByVal lpEnumFontProc As LongPtr, ByVal lParam As LongPtr, ByVal dw As Long) As Long
Private Declare PtrSafe Function GetTextCharset Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetTextCharsetInfo Lib "gdi32" (ByVal hdc As LongPtr, lpSig As FONTSIGNATURE, ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function TranslateCharsetInfo Lib "gdi32" (lpSrc As Long, lpcs As CHARSETINFO, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function GetFontLanguageInfo Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetCharacterPlacement Lib "gdi32" Alias " GetCharacterPlacementA" (ByVal hdc As LongPtr, ByVal lpsz As String, ByVal n1 As Long, ByVal n2 As Long, lpGcpResults As GCP_RESULTS, ByVal dw As Long) As Long

Private Declare PtrSafe Function SetICMMode Lib "gdi32" (ByVal hdc As LongPtr, ByVal n As Long) As Long
Private Declare PtrSafe Function CheckColorsInGamut Lib "gdi32" (ByVal hdc As LongPtr, lpv As Any, lpv2 As Any, ByVal dw As Long) As Long
Private Declare PtrSafe Function GetLogColorSpace Lib "gdi32" Alias "GetLogColorSpaceA" (ByVal hcolorspace As LongPtr, lplogcolorspace As LOGCOLORSPACE, ByVal dw As Long) As Long
Private Declare PtrSafe Function GetColorSpace Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr

Private Declare PtrSafe Function CreateColorSpace Lib "gdi32" Alias "CreateColorSpaceA" (lplogcolorspace As LOGCOLORSPACE) As LongPtr

Private Declare PtrSafe Function SetColorSpace Lib "gdi32" (ByVal hdc As LongPtr, ByVal hcolorspace As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteColorSpace Lib "gdi32" (ByVal hcolorspace As LongPtr) As Long
Private Declare PtrSafe Function GetICMProfile Lib "gdi32" Alias "GetICMProfileA" (ByVal hdc As LongPtr, ByVal dw As LongPtr, ByVal lpStr As String) As Long
Private Declare PtrSafe Function SetICMProfile Lib "gdi32" Alias "SetICMProfileA" (ByVal hdc As LongPtr, ByVal lpStr As String) As Long
Private Declare PtrSafe Function GetDeviceGammaRamp Lib "gdi32" (ByVal hdc As LongPtr, lpv As Any) As Long
Private Declare PtrSafe Function SetDeviceGammaRamp Lib "gdi32" (ByVal hdc As LongPtr, lpv As Any) As Long
Private Declare PtrSafe Function ColorMatchToTarget Lib "gdi32" (ByVal hdc As LongPtr, ByVal hdc2 As LongPtr, ByVal dw As Long) As Long

Private Declare PtrSafe Function EnumICMProfiles Lib "gdi32" Alias "EnumICMProfilesA" (ByVal hdc As LongPtr, ByVal icmEnumProc As LongPtr, ByVal lParam As LongPtr) As Long

Private Declare PtrSafe Function ChoosePixelFormat Lib "gdi32" (ByVal hdc As LongPtr, pPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Private Declare PtrSafe Function CreateDIBSection Lib "gdi32" (ByVal hdc As LongPtr, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As LongPtr, ByVal handle As LongPtr, ByVal dw As Long) As LongPtr
Private Declare PtrSafe Function DescribePixelFormat Lib "gdi32" (ByVal hdc As LongPtr, ByVal n As Long, ByVal un As Long, lpPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Private Declare PtrSafe Function EnumFonts Lib "gdi32" Alias "EnumFontsA" (ByVal hdc As LongPtr, ByVal lpsz As String, ByVal lpFontEnumProc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function EnumMetaFile Lib "gdi32" (ByVal hdc As LongPtr, ByVal hMetafile As LongPtr, ByVal lpMFEnumProc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function EnumObjects Lib "gdi32" (ByVal hdc As LongPtr, ByVal n As Long, ByVal lpGOBJEnumProc As LongPtr, lpVoid As Any) As Long
Private Declare PtrSafe Function FixBrushOrgEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal n1 As Long, ByVal n2 As Long, lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function GetBrushOrgEx Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function GetDIBColorTable Lib "gdi32" (ByVal hdc As LongPtr, ByVal un1 As Long, ByVal un2 As Long, pRGBQuad As RGBQUAD) As Long
Private Declare PtrSafe Function GetPixelFormat Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function LineDDA Lib "gdi32" (ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal lpLineDDAProc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function SetAbortProc Lib "gdi32" (ByVal hdc As LongPtr, ByVal lpAbortProc As LongPtr) As Long
Private Declare PtrSafe Function SetDIBColorTable Lib "gdi32" (ByVal hdc As LongPtr, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As RGBQUAD) As Long
Private Declare PtrSafe Function SetPixelFormat Lib "gdi32" (ByVal hdc As LongPtr, ByVal n As Long, pcPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Private Declare PtrSafe Function SwapBuffers Lib "gdi32" (ByVal hdc As LongPtr) As Long

