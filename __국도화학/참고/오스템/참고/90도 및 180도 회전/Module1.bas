Attribute VB_Name = "Module1"
Option Explicit

'����ü ����
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

'API �Լ� ����
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

'��� ����
Public Const DIB_RGB_COLORS = 0

'���� ����
Public BITMAP_INFO As BITMAPINFO '��Ʈ�� ����ü
Public BITMAP_INPUT() As Byte 'Input �迭
Public BITMAP_OUTPUT() As Byte 'Ouput �迭

'��Ʈ�� �迭 �غ��ϱ�
Public Sub SetBitmapArray_Input(wid As Long, hgt As Long)
    On Error Resume Next
    Dim bitCnt As Long '//��Ʈ��
    Dim BytesPerScanLine As Long
    
    '�ȼ��� 32��Ʈ(4����Ʈ)�� ����
    bitCnt = 32
    
    '��Ʈ�� ����ü �����.
    With BITMAP_INFO.bmiHeader
      .biSize = Len(BITMAP_INFO.bmiHeader) '40����Ʈ
      .biWidth = wid
      .biHeight = -hgt
      .biPlanes = 1
      .biBitCount = bitCnt
      .biCompression = DIB_RGB_COLORS
      BytesPerScanLine = (((.biWidth * .biBitCount) + bitCnt - 1) \ bitCnt) * 4 '������ �����̱� ������ (a+b)\32 => a\32 + b\32, ���� ����� �غ��� wid*4�� ���´�.
      .biSizeImage = BytesPerScanLine * Abs(.biHeight)
    End With
    
    '�迭�� �缱���Ѵ�.
    ReDim BITMAP_INPUT(0 To (bitCnt / 8) - 1, 0 To wid - 1, 0 To hgt - 1)
End Sub

Public Sub SetBitmapArray_Output(wid As Long, hgt As Long)
    On Error Resume Next
    Dim bitCnt As Long '//��Ʈ��
    Dim BytesPerScanLine As Long
    
    '�ȼ��� 32��Ʈ(4����Ʈ)�� ����
    bitCnt = 32
    
    '��Ʈ�� ����ü �����.
    With BITMAP_INFO.bmiHeader
      .biSize = Len(BITMAP_INFO.bmiHeader) '40����Ʈ
      .biWidth = wid
      .biHeight = -hgt
      .biPlanes = 1
      .biBitCount = bitCnt
      .biCompression = DIB_RGB_COLORS
      BytesPerScanLine = (((.biWidth * .biBitCount) + bitCnt - 1) \ bitCnt) * 4 '������ �����̱� ������ (a+b)\32 => a\32 + b\32, ���� ����� �غ��� wid*4�� ���´�.
      .biSizeImage = BytesPerScanLine * Abs(.biHeight)
    End With
    
    '�迭�� �缱���Ѵ�.
    ReDim BITMAP_OUTPUT(0 To (bitCnt / 8) - 1, 0 To wid - 1, 0 To hgt - 1)
End Sub

