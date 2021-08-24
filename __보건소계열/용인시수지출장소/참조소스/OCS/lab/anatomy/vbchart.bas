Attribute VB_Name = "vbChart"
' Function Prototypes
Declare Function chart_Create Lib "CHART2FX.VBX" (ByVal lType As Long, ByVal lStyle As Long, ByVal hWnd As Integer, x As Integer, y As Integer, w As Integer, h As Integer, nPoint As Integer, nSerie As Integer, wIdm As Integer, dwStyle As Long) As Integer
Declare Function chart_Send Lib "CHART2FX.VBX" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Declare Function chart_OpenData Lib "CHART2FX.VBX" (ByVal hWnd As Integer, ByVal wCode As Integer, dwSize As Long) As Long
Declare Function chart_CloseData Lib "CHART2FX.VBX" (ByVal hWnd As Integer, ByVal wCode As Integer) As Integer
Declare Function chart_SetValue Lib "CHART2FX.VBX" (ByVal hWnd As Integer, ByVal nSerie As Integer, ByVal nPoint As Integer, ByVal dValue As Double) As Long
Declare Function chart_SetIniValue Lib "CHART2FX.VBX" (ByVal hWnd As Integer, ByVal nSerie As Integer, ByVal nPoint As Integer, ByVal dValue As Double) As Long
Declare Function chart_SetXvalue Lib "CHART2FX.VBX" (ByVal hWnd As Integer, ByVal nSerie As Integer, ByVal nPoint As Integer, ByVal dValue As Double) As Long
Declare Function chart_SetConst Lib "CHART2FX.VBX" (ByVal hWnd As Integer, ByVal nIndex As Integer, ByVal dValue As Double) As Long
Declare Function chart_SetColor Lib "CHART2FX.VBX" (ByVal hWnd As Integer, ByVal nIndex As Integer, ByVal lColor As Long, ByVal bBack As Integer) As Long
Declare Sub chart_SetAdm Lib "CHART2FX.VBX" (ByVal hWnd As Integer, ByVal nIndex As Integer, ByVal dValue As Double)
Declare Function chart_Get Lib "CHART2FX.VBX" (ByVal hWnd As Integer, ByVal lType As Long, ByVal wCode As Integer) As Double
Declare Function chart_SetStripe Lib "CHART2FX.VBX" (ByVal hWnd As Integer, ByVal nIndex As Integer, ByVal dIni As Double, ByVal dEnd As Double, ByVal lColor As Long) As Long
Declare Function chart_SetStatusItem Lib "CHART2FX.VBX" (ByVal hWnd As Integer, ByVal n As Integer, ByVal Text As Integer, ByVal idm As Integer, ByVal Frame As Integer, ByVal w As Integer, ByVal min As Integer, ByVal desp As Integer, ByVal s As Long) As Long

' graph_OpenData CONSTANTS
Global Const COD_VALUES = 1
Global Const COD_CONSTANTS = 2
Global Const COD_COLORS = 3
Global Const COD_STRIPES = 4
Global Const COD_INIVALUES = 5
Global Const COD_XVALUES = 6
Global Const COD_STATUSITEMS = 7

' definiciones
Global Const COD_UNKNOWN = &HFFFF
Global Const COD_UNCHANGE = 0

' graph_SetAdm CONSTANTS
Global Const CSA_MIN = 0
Global Const CSA_MAX = 1
Global Const CSA_GAP = 2
Global Const CSA_SCALE = 3
Global Const CSA_YLEGGAP = 4
Global Const CSA_PIXXVALUE = 5
Global Const CSA_XMIN = 6
Global Const CSA_XMAX = 7


Function CHART_ML(wLow As Integer, wHi As Integer)
    CHART_ML = CLng(&H10000 * wHi) + wLow
End Function

