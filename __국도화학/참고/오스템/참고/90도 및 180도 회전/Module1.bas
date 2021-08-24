Attribute VB_Name = "Module1"
Option Explicit

'구조체 선언
Public Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors() As RGBQUAD
End Type

'API 함수 선언
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

'상수 선언
Public Const DIB_RGB_COLORS = 0

'변수 선언
Public BITMAP_INFO As BITMAPINFO '비트맵 구조체
Public BITMAP_INPUT() As Byte 'Input 배열
Public BITMAP_OUTPUT() As Byte 'Ouput 배열

'비트맵 배열 준비하기
Public Sub SetBitmapArray_Input(wid As Long, hgt As Long)
    On Error Resume Next
    Dim bitCnt As Long '//비트수
    Dim BytesPerScanLine As Long
    
    '픽셀당 32비트(4바이트)로 설정
    bitCnt = 32
    
    '비트맵 구조체 만든다.
    With BITMAP_INFO.bmiHeader
      .biSize = Len(BITMAP_INFO.bmiHeader) '40바이트
      .biWidth = wid
      .biHeight = -hgt
      .biPlanes = 1
      .biBitCount = bitCnt
      .biCompression = DIB_RGB_COLORS
      BytesPerScanLine = (((.biWidth * .biBitCount) + bitCnt - 1) \ bitCnt) * 4 '나머지 연산이기 때문에 (a+b)\32 => a\32 + b\32, 직접 계산을 해보면 wid*4가 나온다.
      .biSizeImage = BytesPerScanLine * Abs(.biHeight)
    End With
    
    '배열을 재선언한다.
    ReDim BITMAP_INPUT(0 To (bitCnt / 8) - 1, 0 To wid - 1, 0 To hgt - 1)
End Sub

Public Sub SetBitmapArray_Output(wid As Long, hgt As Long)
    On Error Resume Next
    Dim bitCnt As Long '//비트수
    Dim BytesPerScanLine As Long
    
    '픽셀당 32비트(4바이트)로 설정
    bitCnt = 32
    
    '비트맵 구조체 만든다.
    With BITMAP_INFO.bmiHeader
      .biSize = Len(BITMAP_INFO.bmiHeader) '40바이트
      .biWidth = wid
      .biHeight = -hgt
      .biPlanes = 1
      .biBitCount = bitCnt
      .biCompression = DIB_RGB_COLORS
      BytesPerScanLine = (((.biWidth * .biBitCount) + bitCnt - 1) \ bitCnt) * 4 '나머지 연산이기 때문에 (a+b)\32 => a\32 + b\32, 직접 계산을 해보면 wid*4가 나온다.
      .biSizeImage = BytesPerScanLine * Abs(.biHeight)
    End With
    
    '배열을 재선언한다.
    ReDim BITMAP_OUTPUT(0 To (bitCnt / 8) - 1, 0 To wid - 1, 0 To hgt - 1)
End Sub

