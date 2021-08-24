Attribute VB_Name = "DisplaySet"
Option Explicit

Private Declare Function lstrcpy Lib _
"kernel32" Alias "lstrcpyA" _
(lpString1 As Any, lpString2 As Any) As Long

' lstrcpy에서 As String 대신 As Any를 사용한 이유는 문자열의
' 포인터를 얻기위한것이 아니라.. DEVMODE의 구조체에 대한
' 포인터를 얻기 위해서입니다.

Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32

Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Private Type CurrentDisplay
    curX As Long
    curY As Long
    curTmp As Long
End Type
Public CurMode As CurrentDisplay

Private Declare Function ChangeDisplaySettings _
    Lib "User32" Alias "ChangeDisplaySettingsA" _
    (ByVal lpDevMode As Long, ByVal dwflags As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "User32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Const HORZRES = 8 ' Horizontal width in pixels
Const VERTRES = 10 ' Vertical height in pixels

Public Function SetDisplayMode _
                (ByVal Width As Integer, _
                ByVal Height As Integer, _
                byvalColor As Integer) As Long

    Const DM_PELSWIDTH = &H80000
    Const DM_PELSHEIGHT = &H100000
    Const DM_BITSPERPEL = &H40000
    
    Dim NewDevMode As DEVMODE
    Dim pDevmode As Long ' < -- 수정 되었음
    
    With NewDevMode
        .dmSize = 122
        
        If Color = -1 Then  '---->> Color 변경없이
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
        Else
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
        End If
        
        '---->> Width,Height변경
        .dmPelsWidth = Width
        .dmPelsHeight = Height
        
        '----->> Color변경
        If Color <> -1 Then
            .dmBitsPerPel = Color
        End If
    
    End With
    
    '----->> DEVMODE Type을 갖는 변수의 포인터
    pDevmode = lstrcpy(NewDevMode, NewDevMode)
    ' ~~~~~~~~
    ' └─> NewDevMode의 구조체의 포인터를 할당받음..
    ' 즉, NewDevMode의 모든 정보를 pDevmode를 이용
    ' 함으로서 이용할 수 있다.
    
    '----->> DEVMODE Type의 설정된 값으로 DisplayMode변경
    SetDisplayMode = ChangeDisplaySettings(pDevmode, 0)
    ' ~~~~~~~~
    ' 메모리상에 있는 pDevmode와 연관된 모든 정보를
    ' ChangeDisplaySettings()에 넘겨줌니다.

End Function


Public Sub GetDisplayMode()
    Dim dcscreen
    
    dcscreen = GetDC(0)
    ' 수평 해상도
    CurMode.curX = GetDeviceCaps(dcscreen, HORZRES)
    ' 수직 해상도
    CurMode.curY = GetDeviceCaps(dcscreen, VERTRES)
    
    CurMode.curTmp = ReleaseDC(0, dcscreen)
End Sub


