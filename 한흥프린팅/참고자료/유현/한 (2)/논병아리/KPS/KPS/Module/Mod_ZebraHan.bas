Attribute VB_Name = "Mod_ZebraHan"
Option Explicit
Global PrintCancel
Global GRAPCOUNT
Global IMAGEDATA() As String
Global MIMAGEDATA() As String
Global ScreenXPoint, ScreenYPoint
Global FONTDIRECTORY$
Dim Fontsize As Integer
Global gu As String
Sub Label_Print(CUST As String, POSX As Integer, POSY As Integer, Fot As Integer)
Dim Tmp$
Dim Ls_TempData As String
Dim X$, Y$, ft$, fs$, XS$, YS$
       
        '품명
        Tmp$ = CUST
        X$ = POSX '가로위치
        Y$ = POSY '세로위치
        
        Open App.Path + "\Setting\Font.ini" For Input As #2   ' 파일을 입력모드로 오픈한다.
        
            While Not EOF(2)
            
                Line Input #2, Ls_TempData
            
            Wend
        
        Close #2
        
      
        ft$ = Ls_TempData '폰트
        fs$ = Fot '폰트크기
        XS$ = 1 '가로확대
        YS$ = 1 '세로확대
        Call ZEBRA_WINFONT(X$, Y$, Tmp$, ft$, fs$, XS$, YS$, True, False, False, "0")
    
    Exit Sub
Label_Print_Error:
    MsgBox Error$(Err)
    Exit Sub
End Sub
Sub HANCOUNT(H$, ecnt, hcnt)
Dim hlen, i, KsFirst$
hlen = Len(H$)
i = 1
ecnt = 0: hcnt = 0
Do
  KsFirst$ = Mid$(H$, i, 1)
  If Asc(KsFirst$) > 0 Then   '영숫자
    i = i + 1
    ecnt = ecnt + 1
  Else
    i = i + 1
    hcnt = hcnt + 1
  End If
Loop While (i <= hlen)
End Sub


Sub ZEBRA_WINFONT(XX$, YY$, dat$, IFONT$, Fontsize$, XA$, YA$, FontBold, FontItalic, FontReverse, rot$)
'*********************************************************
'* 내용 : WINDOWS용 폰트를 프린터로 출력                 *
'* 변수 : XX$ - 가로위치                                 *
'*        YY$ - 세로위치                                 *
'*        DAT$ - 출력할 데이타                           *
'*        IFONT$ - FONT NAME                             *
'*        FONTSIZE$ - FONT SIZE                          *
'*        XA$ - 가로확대                                 *
'*        YA$ - 세로확대                                 *
'*        FONTBOLD - 진하게(TRUE or FALSE)               *
'*        FONTITALIC - 기울임 (TRUE or FALSE)            *
'*********************************************************
Dim X, Y
    ScreenFrm.Picture = LoadPicture("")
    If FontReverse Then
        ScreenFrm.BackColor = QBColor(0)
        ScreenFrm.ForeColor = QBColor(15)
    Else
        ScreenFrm.BackColor = QBColor(15)
        ScreenFrm.ForeColor = QBColor(0)
    End If
    ScreenFrm.Cls
    ScreenFrm.FontName = IFONT$
    ScreenFrm.Fontsize = Val(Fontsize$)
    ScreenFrm.FontBold = FontBold
    ScreenFrm.FontItalic = FontItalic
    ScreenFrm.Print dat$
    X = ScreenFrm.TextWidth(dat$)
    Y = ScreenFrm.TextHeight(dat$)
    If 0 <> X Mod 8 Then X = ((X \ 8) + 1) * 8
    DoEvents
    Call Screen_Print(XX$, YY$, XA$, YA$, X, Y, rot$)

End Sub
Sub Screen_Print(XX$, YY$, XA$, YA$, XS, YS, rot$)
'*********************************************************
'* 내용 : ScreenFrm의 화면의 내용을 그대로 출력한다.     *
'* 변수 : XX$ - 가로위치(프린터상)                       *
'*        YY$ - 세로위치(프린터상)                       *
'*        DAT$ - 출력할 데이타                           *
'*        XA$ - 가로확대                                 *
'*        YA$ - 세로확대                                 *
'*        XS - 가로크기(인쇄될 ScreenFrm의 화면크기)     *
'*        YS - 세로크기(인쇄될 ScreenFrm의 화면크기)     *
'*        ROT$ - 회전                                    *
'*********************************************************

ReDim IMAGEDATA(YS)
Dim i, j, Tmp1$, tmp2$, TMP3$, GrapName$, mtmp$
Dim size As Integer
Dim mi As Integer
Dim hexname As String


ReDim IMAGEDATA(YS) 'YS개의 요소를 할당합니다
ReDim MIMAGEDATA(YS)
     GRAPCOUNT = GRAPCOUNT + 1
    GrapName$ = Right$("00000000" + Trim(Str$(GRAPCOUNT)), 8)
    Frm_Main.Mcom.Output = "~DG" + GrapName$ + "," + Right$("00000" + Trim$(Str$((XS / 8 * YS))), 5) + "," + Right$("000" + Trim(Str$(XS / 8)), 3) + ","
            For i = 0 To YS - 1   '화면을 읽는다. - 검은색을 찾아낸다
                For j = 0 To XS - 1
                    If ScreenFrm.Point(j, i) = QBColor(0) Then
                        Tmp1$ = Tmp1$ + "1"
                    Else
                        Tmp1$ = Tmp1$ + "0"
                    End If

                    If ((j + 1) Mod 4) = 0 Then
                        IMAGEDATA(i) = IMAGEDATA(i) + BinaryToHex$(Tmp1$) '이진수로 구성된 데이타를 16진수로 변환
                        Tmp1$ = ""
                    End If
                Next j
               Frm_Main.Mcom.Output = IMAGEDATA(i)
            Next i
           Frm_Main.Mcom.Output = "^FO" + XX$ + "," + YY$ + "^XG" + GrapName$ + "," + XA$ + "," + YA$ + "^FS"
End Sub

Function BinaryToHex$(dat$)
'*********************************************************
'* 내용 : 2진수를 16진수로 변환                          *
'* 참고 : DAT$는 반드시 한자리의 16진수이다              *
'*********************************************************

  Select Case dat$
    Case "0000": BinaryToHex$ = "0"
    Case "0001": BinaryToHex$ = "1"
    Case "0010": BinaryToHex$ = "2"
    Case "0011": BinaryToHex$ = "3"
    Case "0100": BinaryToHex$ = "4"
    Case "0101": BinaryToHex$ = "5"
    Case "0110": BinaryToHex$ = "6"
    Case "0111": BinaryToHex$ = "7"
    Case "1000": BinaryToHex$ = "8"
    Case "1001": BinaryToHex$ = "9"
    Case "1010": BinaryToHex$ = "A"
    Case "1011": BinaryToHex$ = "B"
    Case "1100": BinaryToHex$ = "C"
    Case "1101": BinaryToHex$ = "D"
    Case "1110": BinaryToHex$ = "E"
    Case "1111": BinaryToHex$ = "F"
  End Select

End Function
