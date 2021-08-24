Attribute VB_Name = "LogoForm"
DefInt A-Z

#If Win32 Then
    Declare Function StretchBlt Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
#Else
    Declare Function StretchBlt Lib "GDI" (ByVal DhDC, ByVal Sx, ByVal Sy, ByVal Ss, ByVal Sh, ByVal DhDC, ByVal Dx, ByVal Dy, ByVal Dw, ByVal Dh, ByVal Rop As Long) As Integer
    Declare Function BitBlt Lib "GDI" (ByVal hdcDest, ByVal XDest, ByVal nYDest, ByVal nWidth, ByVal nHeight, ByVal hdcSrc, ByVal nXSrc, ByVal nYSrc, ByVal dwRop As Long) As Integer
#End If


Dim I, J, K         As Integer


Sub BlindVert2Pass(StripeWidth As Integer, msecdelay As Long, frm1 As Control, frm2 As Control)

    Dim Stripes As Integer

    Stripes = Int(frm1.ScaleWidth / StripeWidth + 0.5)
    
    For I = 0 To Stripes - 1 Step 2
        P1 = StripeWidth * I
        r% = BitBlt(frm2.hdc, P1, 0, StripeWidth, frm1.ScaleHeight, frm1.hdc, P1, 0, &HCC0020)
        Call delay(msecdelay)
    Next
    
    For I = Stripes - 1 To 0 Step -2
        P1 = StripeWidth * I
        r% = BitBlt(frm2.hdc, P1, 0, StripeWidth, frm1.ScaleHeight, frm1.hdc, P1, 0, &HCC0020)
        Call delay(msecdelay)
    Next

End Sub

Sub BlocksRandom(BlockSize As Integer, msecdelay As Long, frm1 As Control, frm2 As Control)
    Dim K As Integer
    Dim X As Integer, Y As Integer
    Dim XBlocks As Integer, YBlocks As Integer

    XBlocks = frm2.ScaleWidth / BlockSize
    YBlocks = frm2.ScaleHeight / BlockSize
    
    For K = 1 To 8
        tot = 0
        While tot < 1.5 * (XBlocks + 1) * (YBlocks + 1)
            tot = tot + 1
            X = Rnd(1) * XBlocks
            Y = Rnd(1) * YBlocks
            FromX = X * BlockSize
            FromY = Y * BlockSize
            r% = BitBlt(frm2.hdc, FromX, FromY, BlockSize, BlockSize, frm1.hdc, FromX, FromY, &HCC0020)
            Call delay(msecdelay)
        Wend
        BlockSize = BlockSize + BlockSize
        XBlocks = frm2.ScaleWidth / BlockSize
        YBlocks = frm2.ScaleHeight / BlockSize
    Next

End Sub

Sub delay(N As Long)

    Dim I As Double

    For I = 1 To N * 5
    Next

End Sub

Sub Diamond(grainx As Integer, grainy As Integer, msecdelay As Long, frm1 As Control, frm2 As Control)
    Dim Pa As Integer, Pb As Integer

    frm2.AutoRedraw = False
    XSize = frm1.ScaleWidth / grainx
    YSize = frm1.ScaleHeight / grainy
    TotReps = YSize
    If XSize > TotReps Then TotReps = XSize

    XStart = (XSize / 2) * grainx
    YStart = (YSize / 2) * grainy
    r% = BitBlt(frm2.hdc, XStart, YStart, grainx, grainy, frm1.hdc, XStart, YStart, &HCC0020)
    
    KRep = 1
    For J = 1 To TotReps
        X = 0: Y = -KRep
        For I = 1 To KRep
            XStart = (X + XSize / 2) * grainx
            YStart = (Y + YSize / 2) * grainy
            r% = BitBlt(frm2.hdc, XStart, YStart, grainx, grainy, frm1.hdc, XStart, YStart, &HCC0020)
            X = X + 1
            Y = Y + 1
        Next
LOOP2:
        X = KRep: Y = 0
        For I = 1 To KRep
            XStart = (X + XSize / 2) * grainx
            YStart = (Y + YSize / 2) * grainy
            r% = BitBlt(frm2.hdc, XStart, YStart, grainx, grainy, frm1.hdc, XStart, YStart, &HCC0020)
            X = X - 1
            Y = Y + 1
        Next

        X = 0: Y = KRep
        For I = 1 To KRep
            XStart = (X + XSize / 2) * grainx
            YStart = (Y + YSize / 2) * grainy
            r% = BitBlt(frm2.hdc, XStart, YStart, grainx, grainy, frm1.hdc, XStart, YStart, &HCC0020)
            X = X - 1
            Y = Y - 1
        Next

        X = -KRep: Y = 0
        For I = 1 To KRep
            XStart = (X + XSize / 2) * grainx
            YStart = (Y + YSize / 2) * grainy
            r% = BitBlt(frm2.hdc, XStart, YStart, grainx, grainy, frm1.hdc, XStart, YStart, &HCC0020)
            X = X + 1
            Y = Y - 1
        Next
        KRep = KRep + 1
        Call delay(msecdelay)
    Next

End Sub

Sub DropPaint(Stripes As Integer, msecdelay As Long, frm1 As Control, frm2 As Control)

    Dim Pb As Integer

    Pb = frm1.ScaleHeight / Stripes
    
    For I = Stripes To 0 Step -1
        P2 = Pb * I
        P1 = frm1.ScaleWidth
        For J = 0 To I
            P3 = Pb * J
            P4 = 0
            r% = BitBlt(frm2.hdc, 0, P3, frm1.ScaleWidth, Pb, frm1.hdc, 0, P2, &HCC0020)
            Call delay(msecdelay)
        Next
    Next

End Sub

Sub Get_Pic()
'    Static PicNum

'    If PicNum = 4 Then PicNum = 0
'    Select Case PicNum
'    Case 0
'        Form1.Picture1.Picture = LoadPicture("c:\down\logo\image1.bmp")
'    Case 1
'        Form1.Picture1.Picture = LoadPicture("c:\down\logo\image2.bmp")
'    Case 2
 '       Form1.Picture1.Picture = LoadPicture("c:\down\logo\image3.bmp")
'    Case 3
'        Form1.Picture1.Picture = LoadPicture("c:\down\logo\image4.bmp")
'    End Select
'    Form1.Picture2.Width = Form1.Picture1.Width
'    Form1.Picture2.Height = Form1.Picture1.Height
'    PicNum = PicNum + 1

End Sub

Sub VenetianBlindVert(Stripes As Integer, msecdelay As Long, frm1 As Control, frm2 As Control)
Dim StripeWidth As Integer

    StripeWidth = Int(frm2.ScaleWidth / Stripes)
    Wx = 0
    For I = 0 To StripeWidth - 1
        Wx = Wx + 1
        For J = 0 To Stripes - 1
            Px = StripeWidth * J
            Py = frm2.ScaleHeight
            r% = BitBlt(frm2.hdc, Px, 0, Wx, Py, frm1.hdc, Px, 0, &HCC0020)
        Next
        Call delay(msecdelay)
    Next

End Sub

