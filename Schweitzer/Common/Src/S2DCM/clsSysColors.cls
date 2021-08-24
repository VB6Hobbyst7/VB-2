VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysColors"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'** VB에서 제공하는 색상상수들...
Public DCM_Black As Long
Public DCM_White As Long
Public DCM_Yellow As Long
Public DCM_Red As Long
Public DCM_Green As Long
Public DCM_Blue As Long
Public DCM_Magenta As Long
Public DCM_Cyan As Long
Public DCM_MidGray As Long

'** 여기서부터는 User-Defined Color
'Public DCM_Grey As Long
Public DCM_Gray As Long
'Public DCM_LightGrey As Long
Public DCM_LightGray As Long
Public DCM_LightPink As Long
Public DCM_LightRed As Long
Public DCM_LightBlue As Long
Public DCM_MidBlue As Long
Public DCM_Brown As Long
Public DCN_Green As Long

Public DCM_Title_Green As Long
Public DCM_Title_Pink As Long
Public DCM_Title_Blue As Long

'하늘색 : &H00FEDECD&

Private Sub Class_Initialize()

    DCM_Black = vbBlack         '검정색
    DCM_White = vbWhite         '하얀색
    DCM_Yellow = vbYellow       '노란색
    DCM_Red = vbRed             '빨간색
    DCM_Green = vbGreen         '녹색
    DCM_Blue = vbBlue           '파란색
    DCM_Magenta = vbMagenta     '자홍색
    DCM_Cyan = vbCyan           '청록색
    
'    DCM_Grey = &H808080         '회색 --> 이건 쓰지말것. 나중에 지울것임.
    DCM_Gray = &H808080         '회색
    DCM_MidGray = &HC0C0C0
'    DCM_LightGrey = &HE0E0E0    '옅은회색
    DCM_LightGray = &HE0E0E0    '옅은회색 --> 이것두..스펠이 틀려서리...
    DCM_LightPink = &HF7F3F8    '옅은 분홍색
    DCM_LightRed = &H7477EF     '옅은 빨간색
    DCM_LightBlue = &HDF6A3E    '옅은 파란색
    DCM_MidBlue = &HB9602F      '옅은 파란색
    
    DCM_Brown = &H4A4189        '갈색
    DCM_Green = &H8000&         '연두색
    
    'Title Color
    DCM_Title_Green = &HCDD19E  '연두색비스무리.. ^^;
    DCM_Title_Pink = &HD8A9D6   '분홍색비스무리.. ^^;
    DCM_Title_Blue = &HF9A071   '파란색비스무리.. ^^;
End Sub
