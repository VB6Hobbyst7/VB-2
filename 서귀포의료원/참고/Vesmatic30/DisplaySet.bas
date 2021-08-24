Attribute VB_Name = "DisplaySet"
Option Explicit

Private Declare Function lstrcpy Lib _
"kernel32" Alias "lstrcpyA" _
(lpString1 As Any, lpString2 As Any) As Long

' lstrcpy���� As String ��� As Any�� ����� ������ ���ڿ���
' �����͸� ������Ѱ��� �ƴ϶�.. DEVMODE�� ����ü�� ����
' �����͸� ��� ���ؼ��Դϴ�.

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
    Dim pDevmode As Long ' < -- ���� �Ǿ���
    
    With NewDevMode
        .dmSize = 122
        
        If Color = -1 Then  '---->> Color �������
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
        Else
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
        End If
        
        '---->> Width,Height����
        .dmPelsWidth = Width
        .dmPelsHeight = Height
        
        '----->> Color����
        If Color <> -1 Then
            .dmBitsPerPel = Color
        End If
    
    End With
    
    '----->> DEVMODE Type�� ���� ������ ������
    pDevmode = lstrcpy(NewDevMode, NewDevMode)
    ' ~~~~~~~~
    ' ����> NewDevMode�� ����ü�� �����͸� �Ҵ����..
    ' ��, NewDevMode�� ��� ������ pDevmode�� �̿�
    ' �����μ� �̿��� �� �ִ�.
    
    '----->> DEVMODE Type�� ������ ������ DisplayMode����
    SetDisplayMode = ChangeDisplaySettings(pDevmode, 0)
    ' ~~~~~~~~
    ' �޸𸮻� �ִ� pDevmode�� ������ ��� ������
    ' ChangeDisplaySettings()�� �Ѱ��ܴϴ�.

End Function


Public Sub GetDisplayMode()
    Dim dcscreen
    
    dcscreen = GetDC(0)
    ' ���� �ػ�
    CurMode.curX = GetDeviceCaps(dcscreen, HORZRES)
    ' ���� �ػ�
    CurMode.curY = GetDeviceCaps(dcscreen, VERTRES)
    
    CurMode.curTmp = ReleaseDC(0, dcscreen)
End Sub


