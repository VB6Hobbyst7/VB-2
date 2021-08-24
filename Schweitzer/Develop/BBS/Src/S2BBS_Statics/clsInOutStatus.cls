VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInOutStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private PrtLeft   As Long
Private Left1     As Long
Private Left2     As Long
Private Left3     As Long
Private PrtTop    As Long
Private Top1      As Long
'Private DBConn As DrDatabase                'DataBase 연결

Private LineSpace As Long

Private Sub print_BasePos()
    
    PrtLeft = 20
    PrtTop = 40
    LineSpace = 6
    
    Left1 = PrtLeft + LineSpace                         '혈액원찍는 자리
    Left2 = PrtLeft + LineSpace * 2                     '그룹찍히는 자리
    Left3 = PrtLeft + LineSpace * 6                     '일자찍히는 자리
    
    Top1 = PrtTop + LineSpace                           '초기값다음줄부터
    
End Sub

Public Sub Hearder_Line(ByVal RS As Recordset, ByVal center As String, ByVal ABO As String, ByVal Rh As String, _
                        ByVal Volume As String, ByVal PMonth As String, ByVal mode As String, _
                        ByVal objdic As clsDictionary)
    Dim Pmode   As String
    Dim Pmode1  As String
    Dim Pmode2  As String
    Dim PABO    As String
    Dim PVolume As String
    Dim PCenter As String
    Dim xP(30)  As Long
    Dim AddLine As Long
    Dim Right   As Long
    Dim AddTop  As Long
    Dim cnt     As Long
    Dim Div     As Long
    Dim Div1    As Long
    Dim Div2    As Long
    Dim ii      As Long
    Dim jj      As Long
    
    
    Select Case mode
        Case 1: Pmode = "입 고 계": Pmode1 = "입": Pmode2 = "고"
        Case 2: Pmode = "출 고 계": Pmode1 = "출": Pmode2 = "고"
        Case 3: Pmode = "반 환 계": Pmode1 = "반": Pmode2 = "환"
        Case 4: Pmode = "폐 기 계": Pmode1 = "폐": Pmode2 = "기"
    End Select
    
    '------------------------------
    'Printer Setting 및 화면 그리기
    '------------------------------
    
    cnt = RS.RecordCount
    print_BasePos
    P_PrtSet
    
    Printer.FontSize = 15: Printer.FontBold = True: Printer.FontUnderline = True
    Call P_FIX(PMonth & "월 혈액 " & Pmode1 & Pmode2 & " 현황", PrtLeft, PrtTop - 20, Printer.ScaleWidth - (PrtLeft * 2), "C", LineSpace, "C")
    Printer.FontSize = 10: Printer.FontBold = False: Printer.FontUnderline = False
    
    '-----------------------------------
    '제제찍어줄 박스그리기
    '-----------------------------------
    Printer.DrawWidth = 4
    Printer.Line (PrtLeft, PrtTop)-(Left3, Top1 + (LineSpace * (cnt * 3 + 1))), , B
    Printer.Line (PrtLeft, Top1)-(Left3, Top1)
    Printer.DrawWidth = 2
    
    '------------------------------------
    '입고종류별 제제 찍어주는 위치 그리기
    '------------------------------------
    For ii = 1 To (cnt * 3)
        AddTop = LineSpace * ii
        If ii Mod cnt = 0 Then
            Printer.DrawWidth = 4
            Printer.Line (Left1, Top1 + AddTop)-(Left3, Top1 + AddTop)
            Printer.DrawWidth = 2
            Div = Div + 1
            If Div = 1 Then
                Call P_FIX("혈", Left1, Top1 + (LineSpace * (ii \ cnt)), LineSpace, "C", LineSpace, "C")
                Call P_FIX("액", Left1, Top1 + (LineSpace * (ii \ cnt + 1)), LineSpace, "C", LineSpace, "C")
                Call P_FIX("원", Left1, Top1 + (LineSpace * (ii \ cnt + 2)), LineSpace, "C", LineSpace, "C")
                Div1 = ii
            ElseIf Div = 2 Then
                Call P_FIX("헌", Left1, Top1 + (LineSpace * (Div1 + 1)), LineSpace, "C", LineSpace, "C")
                Call P_FIX("혈", Left1, Top1 + (LineSpace * (Div1 + 2)), LineSpace, "C", LineSpace, "C")
                Div2 = ii
            ElseIf Div = 3 Then
                Call P_FIX("외", Left1, Top1 + (LineSpace * (Div2 + 1)), LineSpace, "C", LineSpace, "C")
                Call P_FIX("부", Left1, Top1 + (LineSpace * (Div2 + 2)), LineSpace, "C", LineSpace, "C")
            End If
        Else
            Printer.Line (Left2, Top1 + AddTop)-(Left3, Top1 + AddTop)
        End If
    Next
    
    Call P_FIX(Pmode1, PrtLeft, Top1 + (LineSpace * (Div1)), LineSpace, "C", LineSpace, "C")
    Call P_FIX(Pmode2, PrtLeft, Top1 + (LineSpace * (Div2 - 1)), LineSpace, "C", LineSpace, "C")
    
    Printer.DrawWidth = 4
    Printer.Line (Left1, Top1)-(Left1, Top1 + AddTop)
    Printer.Line (Left2, Top1)-(Left2, Top1 + AddTop)
    Printer.Line (PrtLeft, Top1 + AddTop)-(Left3, Top1 + AddTop)
    Call P_FIX(Pmode, PrtLeft, Top1 + AddTop, Left3 - PrtLeft, "C", LineSpace, "C")
    Printer.DrawWidth = 2
    
    '-------------------------------
    '날짜를 찍기 위해서 박스를 그리자
    '-------------------------------
    For ii = 1 To 31
        xP(ii - 1) = Left3 + LineSpace * ii
    Next
    Printer.Line (Left3, PrtTop)-(xP(30) + LineSpace * 2, Top1 + AddTop + LineSpace), , B
    
    '----------------------------
    '세로선 및 날짜 찍기
    '----------------------------
    For ii = LineSpace To LineSpace * 31 Step LineSpace
        Printer.Line (Left3 + ii, PrtTop)-(Left3 + ii, Top1 + AddTop + LineSpace)
        Call P_FIX(ii \ 6, Left3 + ii - LineSpace, PrtTop, LineSpace, "C", LineSpace, "C")
    Next
    Call P_FIX("합계", xP(30) + 3, PrtTop, LineSpace, "C", LineSpace, "C")
    
    '----------------------------
    '가로선
    '----------------------------
    For ii = 1 To cnt * 3 + 2
        Printer.Line (Left3, PrtTop + LineSpace * ii)-(xP(30) + LineSpace * 2, PrtTop + LineSpace * ii)
    Next
    
    '----------------------------
    '데이타 찍어주기....
    '----------------------------
    PCenter = IIf(center = "ALL", "모든센터", center)
    PABO = IIf(ABO = "ALL", "모든 혈액형", ABO & Rh)
    PVolume = IIf(Volume = "ALL", "모든 용량", Volume)
    
    Call P_FIX("혈액형: " & PABO & "       용량: " & PVolume, PrtLeft, PrtTop - LineSpace, xP(30), "L", LineSpace, "C")
    Call P_FIX(PCenter, PrtLeft, PrtTop, Left3 - PrtLeft, "C", LineSpace, "C")
    With RS
        For ii = 1 To 3
            Do Until .EOF
                Call P_FIX(.Fields("field1").Value & "", Left2, Top1 + AddLine, Left3 - Left2, "L", LineSpace, "C")
                jj = jj + 1
                AddLine = LineSpace * jj
                .MoveNext
            Loop
            .MoveFirst
        Next
    End With
    '----------------------------
    '수량출력하기....
    '----------------------------
    Dim total(31) As String
    
    With objdic
        .MoveFirst
        AddLine = 0
        Do Until .EOF
            Call P_FIX(IIf(.Fields("day1") = "0", "", .Fields("day1")), xP(0) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day2") = "0", "", .Fields("day2")), xP(1) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day3") = "0", "", .Fields("day3")), xP(2) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day4") = "0", "", .Fields("day4")), xP(3) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day5") = "0", "", .Fields("day5")), xP(4) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day6") = "0", "", .Fields("day6")), xP(5) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day7") = "0", "", .Fields("day7")), xP(6) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day8") = "0", "", .Fields("day8")), xP(7) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day9") = "0", "", .Fields("day9")), xP(8) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day10") = "0", "", .Fields("day10")), xP(9) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day11") = "0", "", .Fields("day11")), xP(10) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day12") = "0", "", .Fields("day12")), xP(11) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day13") = "0", "", .Fields("day13")), xP(12) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day14") = "0", "", .Fields("day14")), xP(13) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day15") = "0", "", .Fields("day15")), xP(14) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day16") = "0", "", .Fields("day16")), xP(15) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day17") = "0", "", .Fields("day17")), xP(16) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day18") = "0", "", .Fields("day18")), xP(17) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day19") = "0", "", .Fields("day19")), xP(18) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day20") = "0", "", .Fields("day20")), xP(19) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day21") = "0", "", .Fields("day21")), xP(20) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day22") = "0", "", .Fields("day22")), xP(21) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day23") = "0", "", .Fields("day23")), xP(22) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day24") = "0", "", .Fields("day24")), xP(23) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day25") = "0", "", .Fields("day25")), xP(24) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day26") = "0", "", .Fields("day26")), xP(25) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day27") = "0", "", .Fields("day27")), xP(26) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day28") = "0", "", .Fields("day28")), xP(27) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day29") = "0", "", .Fields("day29")), xP(28) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day30") = "0", "", .Fields("day30")), xP(29) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("day31") = "0", "", .Fields("day31")), xP(30) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            Call P_FIX(IIf(.Fields("tot") = "0", "", .Fields("tot")), xP(30) + 3, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
            
            total(0) = Val(total(0)) + Val(.Fields("day1"))
            total(1) = Val(total(1)) + Val(.Fields("day2"))
            total(2) = Val(total(2)) + Val(.Fields("day3"))
            total(3) = Val(total(3)) + Val(.Fields("day4"))
            total(4) = Val(total(4)) + Val(.Fields("day5"))
            
            total(5) = Val(total(5)) + Val(.Fields("day6"))
            total(6) = Val(total(6)) + Val(.Fields("day7"))
            total(7) = Val(total(7)) + Val(.Fields("day8"))
            total(8) = Val(total(8)) + Val(.Fields("day9"))
            total(9) = Val(total(9)) + Val(.Fields("day10"))
            
            total(10) = Val(total(10)) + Val(.Fields("day11"))
            total(11) = Val(total(11)) + Val(.Fields("day12"))
            total(12) = Val(total(12)) + Val(.Fields("day13"))
            total(13) = Val(total(13)) + Val(.Fields("day14"))
            total(14) = Val(total(14)) + Val(.Fields("day15"))
            
            total(15) = Val(total(15)) + Val(.Fields("day16"))
            total(16) = Val(total(16)) + Val(.Fields("day17"))
            total(17) = Val(total(17)) + Val(.Fields("day18"))
            total(18) = Val(total(18)) + Val(.Fields("day19"))
            total(19) = Val(total(19)) + Val(.Fields("day20"))
            
            total(20) = Val(total(20)) + Val(.Fields("day21"))
            total(21) = Val(total(21)) + Val(.Fields("day22"))
            total(22) = Val(total(22)) + Val(.Fields("day23"))
            total(23) = Val(total(23)) + Val(.Fields("day24"))
            total(24) = Val(total(24)) + Val(.Fields("day25"))
            
            total(25) = Val(total(25)) + Val(.Fields("day26"))
            total(26) = Val(total(26)) + Val(.Fields("day27"))
            total(27) = Val(total(27)) + Val(.Fields("day28"))
            total(28) = Val(total(28)) + Val(.Fields("day29"))
            total(29) = Val(total(29)) + Val(.Fields("day30"))
            
            total(30) = Val(total(30)) + Val(.Fields("day31"))
            total(31) = Val(total(31)) + Val(.Fields("tot"))
           
            AddLine = AddLine + LineSpace
            .MoveNext
        Loop
    End With
    
    For ii = 0 To 31
        If total(ii) = "0" Then total(ii) = ""
    Next ii
    
    For ii = 0 To 31
        Call P_FIX(total(0), xP(0) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(1), xP(1) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(2), xP(2) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(3), xP(3) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(4), xP(4) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(5), xP(5) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(6), xP(6) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(7), xP(7) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(8), xP(8) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(9), xP(9) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(10), xP(10) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(11), xP(11) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(12), xP(12) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(13), xP(13) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(14), xP(14) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(15), xP(15) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(16), xP(16) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(17), xP(17) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(18), xP(18) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(19), xP(19) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(20), xP(20) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(21), xP(21) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(22), xP(22) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(23), xP(23) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(24), xP(24) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(25), xP(25) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(26), xP(26) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(27), xP(27) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(28), xP(28) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(29), xP(29) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(30), xP(30) - LineSpace, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
        Call P_FIX(total(31), xP(30) + 3, Top1 + AddLine, LineSpace, "R", LineSpace, "C")
    Next
    
    Printer.EndDoc

End Sub

