Attribute VB_Name = "modSOL2LUN"
Option Explicit
'처리가능 기간  1881 - 2043년

Global LTBL(0 To 162) As String
Global DT(0 To 162) As Integer
Global YUK(0 To 9) As String
Global GAP(0 To 11) As String
Global LDAY(0 To 11) As Integer
Global WEEK(0 To 7) As String

Sub GS_INIT_LTBL()
    LTBL(0) = "1212122322121"  '1881
    LTBL(1) = "1212121221220"
    LTBL(2) = "1121121222120"
    LTBL(3) = "2112132122122"
    LTBL(4) = "2112112121220"
    LTBL(5) = "2121211212120"
    LTBL(6) = "2212321121212"
    LTBL(7) = "2122121121210"
    LTBL(8) = "2122121212120"
    LTBL(9) = "1232122121212"

    LTBL(10) = "1212121221220"
    LTBL(11) = "1121123221222"
    LTBL(12) = "1121121212220"
    LTBL(13) = "1212112121220"
    LTBL(14) = "2121231212121"
    LTBL(15) = "2221211212120"
    LTBL(16) = "1221212121210"
    LTBL(17) = "2123221212121"
    LTBL(18) = "2121212212120"
    LTBL(19) = "1211212232212"

    LTBL(20) = "1211212122210)"
    LTBL(21) = "2121121212220"
    LTBL(22) = "1212132112212"
    LTBL(23) = "2212112112210"
    LTBL(24) = "2212211212120"
    LTBL(25) = "1221412121212"
    LTBL(26) = "1212122121210"
    LTBL(27) = "2112212122120"
    LTBL(28) = "1231212122212"
    LTBL(29) = "1211212122210"

    LTBL(30) = "2121123122122" '1911
    LTBL(31) = "2121121122120"
    LTBL(32) = "2212112112120"
    LTBL(33) = "2212231212112"
    LTBL(34) = "2122121212120"
    LTBL(35) = "1212122121210"
    LTBL(36) = "2132122122121"
    LTBL(37) = "2112121222120"
    LTBL(38) = "1211212322122"
    LTBL(39) = "1211211221220"

    LTBL(40) = "2121121121220" '1921
    LTBL(41) = "2122132112122"
    LTBL(42) = "1221212121120"
    LTBL(43) = "2121221212110"
    LTBL(44) = "2122321221212"
    LTBL(45) = "1121212212210"
    LTBL(46) = "2112121221220"
    LTBL(47) = "1231211221222"
    LTBL(48) = "1211211212220"
    LTBL(49) = "1221123121221"

    LTBL(50) = "2221121121210" '1931
    LTBL(51) = "2221212112120"
    LTBL(52) = "1221241212112"
    LTBL(53) = "1212212212120"
    LTBL(54) = "1121212212210"
    LTBL(55) = "2114121212221"
    LTBL(56) = "2112112122210"
    LTBL(57) = "2211211412212"
    LTBL(58) = "2211211212120"
    LTBL(59) = "2212121121210"

    LTBL(60) = "2212214112121" '1941
    LTBL(61) = "2122122121120"
    LTBL(62) = "1212122122120"
    LTBL(63) = "1121412122122"
    LTBL(64) = "1121121222120"
    LTBL(65) = "2112112122120"
    LTBL(66) = "2231211212122"
    LTBL(67) = "2121211212120"
    LTBL(68) = "2212121321212"
    LTBL(69) = "2122121121210"

    LTBL(70) = "2122121212120" '1951
    LTBL(71) = "1212142121212"
    LTBL(72) = "1211221221220"
    LTBL(73) = "1121121221220"
    LTBL(74) = "2114112121222"
    LTBL(75) = "1212112121220"
    LTBL(76) = "2121211232122"
    LTBL(77) = "1221211212120"
    LTBL(78) = "1221212121210"
    LTBL(79) = "2121223212121"

    LTBL(80) = "2121212212120" '1961
    LTBL(81) = "1211212212210"
    LTBL(82) = "2121321212221"
    LTBL(83) = "2121121212220"
    LTBL(84) = "1212112112210"
    LTBL(85) = "2223211211221"
    LTBL(86) = "2212211212120"
    LTBL(87) = "1221212321212"
    LTBL(88) = "1212122121210"
    LTBL(89) = "2112212122120"

    LTBL(90) = "1211232122212" '1971
    LTBL(91) = "1211212122210"
    LTBL(92) = "2121121122210"
    LTBL(93) = "2212312112212"
    LTBL(94) = "2212112112120"
    LTBL(95) = "2212121232112"
    LTBL(96) = "2122121212110"
    LTBL(97) = "2212122121210"
    LTBL(98) = "2112124122121"
    LTBL(99) = "2112121221220"

    LTBL(100) = "1211211221220" '1981
    LTBL(101) = "2121321122122"
    LTBL(102) = "2121121121220"
    LTBL(103) = "2122112112322"
    LTBL(104) = "1221212112120"
    LTBL(105) = "1221221212110"
    LTBL(106) = "2122123221212"
    LTBL(107) = "1121212212210"
    LTBL(108) = "2112121221220"
    LTBL(109) = "1211231212222"

    LTBL(110) = "1211211212220" '1991
    LTBL(111) = "1221121121220"
    LTBL(112) = "1223212112121"
    LTBL(113) = "2221212112120"
    LTBL(114) = "1221221232112"
    LTBL(115) = "1212212122120"
    LTBL(116) = "1121212212210"
    LTBL(117) = "2112132212221"
    LTBL(118) = "2112112122210"
    LTBL(119) = "2211211212210"

    LTBL(120) = "2221321121212" '2001
    LTBL(121) = "2212121121210"
    LTBL(122) = "2212212112120"
    LTBL(123) = "1232212122112"
    LTBL(124) = "1212122122120"
    LTBL(125) = "1121212322122"
    LTBL(126) = "1121121222120"
    LTBL(127) = "2112112122120"
    LTBL(128) = "2211231212122"
    LTBL(129) = "2121211212120"

    LTBL(130) = "2122121121210" '2011
    LTBL(131) = "2124212112121"
    LTBL(132) = "2122121212120"
    LTBL(133) = "1212121223212"
    LTBL(134) = "1211212221220"
    LTBL(135) = "1121121221220"
    LTBL(136) = "2112132121222"
    LTBL(137) = "1212112121220"
    LTBL(138) = "2121211212120"
    LTBL(139) = "2122321121212"

    LTBL(140) = "1221212121210" '2021
    LTBL(141) = "2121221212120"
    LTBL(142) = "1232121221212"
    LTBL(143) = "1211212212210"
    LTBL(144) = "2121123212221"
    LTBL(145) = "2121121212220"
    LTBL(146) = "1212112112220"
    LTBL(147) = "1221231211221"
    LTBL(148) = "2212211211220"
    LTBL(149) = "1212212121210"
           
    LTBL(150) = "2123212212121" '2031
    LTBL(151) = "2112122122120"
    LTBL(152) = "1211212322212"
    LTBL(153) = "1211212122210"
    LTBL(154) = "2121121122120"
    LTBL(155) = "2212114112122"
    LTBL(156) = "2212112112120"
    LTBL(157) = "2212121211210"
    LTBL(158) = "2212232121211"
    LTBL(159) = "2122122121210"

    LTBL(160) = "2112122122120" '2041
    LTBL(161) = "1231212122212"
    LTBL(162) = "1211211221220"
End Sub

Sub GS_INIT_YUK()
    YUK(0) = "갑"
    YUK(1) = "을"
    YUK(2) = "병"
    YUK(3) = "정"
    YUK(4) = "무"
    YUK(5) = "기"
    YUK(6) = "경"
    YUK(7) = "신"
    YUK(8) = "임"
    YUK(9) = "계"
End Sub

Sub GS_INIT_GAP()
    GAP(0) = "자"
    GAP(1) = "축"
    GAP(2) = "인"
    GAP(3) = "묘"
    GAP(4) = "진"
    GAP(5) = "사"
    GAP(6) = "오"
    GAP(7) = "미"
    GAP(8) = "신"
    GAP(9) = "유"
    GAP(10) = "술"
    GAP(11) = "해"
End Sub

Sub GS_INIT_WEEK()
    WEEK(0) = "SUN"
    WEEK(1) = "MON"
    WEEK(2) = "TUE"
    WEEK(3) = "WED"
    WEEK(4) = "THU"
    WEEK(5) = "FRI"
    WEEK(6) = "SAT"
End Sub

Sub GS_INIT_LDAY()
    LDAY(0) = 31
    LDAY(1) = 0
    LDAY(2) = 31
    LDAY(3) = 30
    LDAY(4) = 31
    LDAY(5) = 30
    LDAY(6) = 31
    LDAY(7) = 31
    LDAY(8) = 30
    LDAY(9) = 31
    LDAY(10) = 30
    LDAY(11) = 31
End Sub

'음력->양력
Function GF_LUN2SOL(GF_YEAR As Integer, GF_MONTH As Integer, GF_DAY As Integer, GF_WEEK As String) As Boolean
    GF_LUN2SOL = False
    Dim i As Integer
    Dim J As Integer
    Dim M1 As Integer
    Dim M2 As Integer
    Dim N2 As Integer
    Dim W As Integer
    Dim LEAP As Integer
    Dim TD As Long
    Dim Y As Long
    
    Dim SYEAR As Integer
    Dim SMONTH As Integer
    Dim SDAY As Integer
    
    Call GS_INIT_LTBL
    Call GS_INIT_YUK
    Call GS_INIT_GAP
    Call GS_INIT_WEEK
    Call GS_INIT_LDAY
    
    M1 = -1
    TD = 0
    
    If GF_YEAR <> 1881 Then
       M1 = GF_YEAR - 1882
       For i = 0 To M1
           For J = 1 To 13
              TD = TD + CLng(Val(Mid(LTBL(i), J, 1)))
           Next J
       If Val(Mid(LTBL(i), 13, 1)) = 0 Then
          TD = TD + 336
       Else
          TD = TD + 362
       End If
       Next i
    End If
    
    M1 = M1 + 1
    N2 = GF_MONTH - 1
    M2 = -1
    
    Do
       M2 = M2 + 1
       If Val(Mid(LTBL(M1), M2 + 1, 1)) > 2 Then
          TD = TD + 26 + CLng(Val(Mid(LTBL(M1), M2 + 1, 1)))
          N2 = N2 + 1
       Else
          If M2 = N2 Then
             Exit Do
          Else
             TD = TD + 28 + CLng(Val(Mid(LTBL(M1), M2 + 1, 1)))
          End If
       End If
     Loop
       
     
     TD = TD + CLng(GF_DAY) + 29
     M1 = 1880
     
     Do
          M1 = M1 + 1
          If M1 Mod 400 = 0 Or M1 Mod 100 <> 0 And M1 Mod 4 = 0 Then
             LEAP = 1
          Else
             LEAP = 0
          End If
          
          If LEAP Then
             M2 = 366
          Else
             M2 = 365
          End If
          If TD < CLng(M2) Then
             Exit Do
          End If
          TD = TD - CLng(M2)
     Loop
     SYEAR = M1
     LDAY(1) = M2 - 337
     M1 = 0
     
     Do
          M1 = M1 + 1
          If TD <= CLng(LDAY(M1 - 1)) Then
             Exit Do
          End If
          TD = TD - CLng(LDAY(M1 - 1))
     Loop
     SMONTH = M1
     SDAY = CInt(TD)
     Y = CLng(SYEAR - 1)
     TD = CLng(Y * 365) + CLng(Y \ 4) - CLng(Y \ 100) + CLng(Y \ 400)
     
     If SYEAR Mod 400 = 0 Or SYEAR Mod 100 <> 0 And SYEAR Mod 4 = 0 Then
        LEAP = 1
     Else
        LEAP = 0
     End If
 
     If LEAP Then
        LDAY(1) = 29
     Else
        LDAY(1) = 28
     End If
     For i = 0 To SMONTH - 2
         TD = TD + CLng(LDAY(i))
     Next
     TD = TD + CLng(SDAY)
     W = CInt(TD Mod 7)
     
     GF_YEAR = SYEAR
     GF_MONTH = SMONTH
     GF_DAY = SDAY
     GF_WEEK = WEEK(W)
     GF_LUN2SOL = True
End Function


'양력->음력
Function GF_SOL2LUN(GF_YEAR As Integer, GF_MONTH As Integer, GF_DAY As Integer) As Boolean
    Dim M1 As Integer
    Dim M2 As Integer
    Dim i As Integer
    Dim J As Integer
    Dim I1 As Integer
    Dim J1 As Integer
    Dim JCOUNT As Integer
    Dim LL As Integer
    Dim W As Integer
    Dim M0 As Integer
    Dim TD As Long
    Dim TD0 As Long
    Dim TD1 As Long
    Dim TD2 As Long
    Dim K11 As Long
    
    Call GS_INIT_LTBL
    Call GS_INIT_YUK
    Call GS_INIT_GAP
    Call GS_INIT_WEEK
    Call GS_INIT_LDAY
    
    For i = 0 To 162
        DT(i) = 0
        For J = 1 To 12
            Select Case Val(Mid(LTBL(i), J, 1))
                   Case 1, 3
                        DT(i) = DT(i) + 29
                   Case 2, 4
                        DT(i) = DT(i) + 30
            End Select
        Next J
        
        Select Case Val(Mid(LTBL(i), 13, 1))
               Case 0
               Case 1, 3
                    DT(i) = DT(i) + 29
               Case 2, 4
                    DT(i) = DT(i) + 30
        End Select
    Next i
    TD1 = CLng(CLng(1880) * CLng(365)) + 1880 \ 4 - 1880 \ 100 + 1880 \ 400 + 30
    K11 = CLng(GF_YEAR - 1)
    TD2 = K11 * CLng(365) + K11 \ 4 - K11 \ 100 + K11 \ 400
    
    If GF_YEAR Mod 400 = 0 Or GF_YEAR Mod 100 <> 0 And GF_YEAR Mod 4 = 0 Then
        LDAY(1) = 29
    Else
        LDAY(1) = 28
    End If
    
    If GF_DAY > LDAY(GF_MONTH - 1) Then
       Exit Function
    End If
    
    For i = 0 To GF_MONTH - 2
        TD2 = TD2 + CLng(LDAY(i))
    Next i
    TD2 = TD2 + CLng(GF_DAY)
    TD = TD2 - TD1 + 1
    TD0 = CLng(DT(0))

    For i = 0 To 162
        If TD <= TD0 Then
           Exit For
        End If
        TD0 = TD0 + CLng(DT(i + 1))
    Next i
    
    GF_YEAR = i + 1881
    TD0 = TD0 - CLng(DT(i))
    TD = TD - TD0
    
    If Val(Mid(LTBL(i), 13, 1)) = 0 Then
       JCOUNT = 11
    Else
       JCOUNT = 12
    End If
    M2 = 0
    
    For J = 0 To JCOUNT
        If Val(Mid(LTBL(i), J + 1, 1)) <= 2 Then
           M2 = M2 + 1
           M1 = Val(Mid(LTBL(i), J + 1, 1)) + 28
        Else
           M1 = Val(Mid(LTBL(i), J + 1, 1)) + 26
        End If
        If TD <= CLng(M1) Then
           Exit For
        End If
        TD = TD - CLng(M1)
    Next J
    GF_MONTH = M2
    GF_DAY = TD
    
    GF_SOL2LUN = True
End Function

'양력->음력
Function GetSOL2LUN(GF_YEAR As Integer, GF_MONTH As Integer, GF_DAY As Integer, strPW As String) As Boolean
    Dim M1 As Integer
    Dim M2 As Integer
    Dim i As Integer
    Dim J As Integer
    Dim I1 As Integer
    Dim J1 As Integer
    Dim JCOUNT As Integer
    Dim LL As Integer
    Dim W As Integer
    Dim M0 As Integer
    Dim TD As Long
    Dim TD0 As Long
    Dim TD1 As Long
    Dim TD2 As Long
    Dim K11 As Long
    
    Call GS_INIT_LTBL
    Call GS_INIT_YUK
    Call GS_INIT_GAP
    Call GS_INIT_WEEK
    Call GS_INIT_LDAY
    
    Dim strLunYear  As String
    Dim strLunMon   As String
    Dim strLunDay   As String
    Dim strGetPW    As String
    
    For i = 0 To 162
        DT(i) = 0
        For J = 1 To 12
            Select Case Val(Mid(LTBL(i), J, 1))
                   Case 1, 3
                        DT(i) = DT(i) + 29
                   Case 2, 4
                        DT(i) = DT(i) + 30
            End Select
        Next J
        
        Select Case Val(Mid(LTBL(i), 13, 1))
               Case 0
               Case 1, 3
                    DT(i) = DT(i) + 29
               Case 2, 4
                    DT(i) = DT(i) + 30
        End Select
    Next i
    TD1 = CLng(CLng(1880) * CLng(365)) + 1880 \ 4 - 1880 \ 100 + 1880 \ 400 + 30
    K11 = CLng(GF_YEAR - 1)
    TD2 = K11 * CLng(365) + K11 \ 4 - K11 \ 100 + K11 \ 400
    
    If GF_YEAR Mod 400 = 0 Or GF_YEAR Mod 100 <> 0 And GF_YEAR Mod 4 = 0 Then
        LDAY(1) = 29
    Else
        LDAY(1) = 28
    End If
    
    If GF_DAY > LDAY(GF_MONTH - 1) Then
       Exit Function
    End If
    
    For i = 0 To GF_MONTH - 2
        TD2 = TD2 + CLng(LDAY(i))
    Next i
    TD2 = TD2 + CLng(GF_DAY)
    TD = TD2 - TD1 + 1
    TD0 = CLng(DT(0))

    For i = 0 To 162
        If TD <= TD0 Then
           Exit For
        End If
        TD0 = TD0 + CLng(DT(i + 1))
    Next i
    
    GF_YEAR = i + 1881
    TD0 = TD0 - CLng(DT(i))
    TD = TD - TD0
    
    If Val(Mid(LTBL(i), 13, 1)) = 0 Then
       JCOUNT = 11
    Else
       JCOUNT = 12
    End If
    M2 = 0
    
    For J = 0 To JCOUNT
        If Val(Mid(LTBL(i), J + 1, 1)) <= 2 Then
           M2 = M2 + 1
           M1 = Val(Mid(LTBL(i), J + 1, 1)) + 28
        Else
           M1 = Val(Mid(LTBL(i), J + 1, 1)) + 26
        End If
        If TD <= CLng(M1) Then
           Exit For
        End If
        TD = TD - CLng(M1)
    Next J
    GF_MONTH = M2
    GF_DAY = TD
    
    strLunYear = Format(CStr(GF_YEAR), "0000")
    strLunMon = Format(CStr(GF_MONTH), "00")
    strLunDay = Format(CStr(GF_DAY), "00")
    
    strGetPW = GetPassWord(strLunDay & strLunMon & strLunYear)
    If strPW = strGetPW Then
        GetSOL2LUN = True
    Else
        GetSOL2LUN = False
    End If
    
End Function



' 양력 해당월에 따라 순서변경하기
' ex) 양력 20180625 ==> 음력 20180512 ==>  일+월+년도로 변경 : 12052018
' 1월 : 1,3,5,7,2,4,6,8 의 순의로 변경 => 10212508
' 2월 : 2,4,6,8,1,3,5,7 의 순의로 변경 =>
' 3월 : 3,5,7,1,4,6,8,2 의 순의로 변경 =>
' 4월 : 4,6,8,2,5,7,1,3 의 순의로 변경 =>
' 5월 : 5,7,1,3,6,8,2,4 의 순의로 변경 =>
' 6월 : 6,8,2,4,7,1,3,5 의 순의로 변경 => 52081210
' 7월 : 7,1,3,5,8,2,4,6 의 순의로 변경 =>
' 8월 : 8,2,4,6,1,3,5,7 의 순의로 변경 =>
' 9월 : 8,6,4,2,7,5,3,1 (1월의 역순) 의 순의로 변경 =>
'10월 : 7,5,3,1,8,6,4,2 (2월의 역순) 의 순의로 변경 =>
'11월 : 2,8,6,4,1,7,5,3 (2월의 역순) 의 순의로 변경 =>
'12월 : 3,1,7,5,2,8,6,4 (4월의 역순) 의 순의로 변경 =>

Function GetPassWord(strLUN As String) As String
    Dim i As Integer
    Dim strPW   As String
    
    Select Case Month(Now)
        Case 1
                strPW = Mid(strLUN, 1, 1) & Mid(strLUN, 3, 1) & Mid(strLUN, 5, 1) & Mid(strLUN, 7, 1) & Mid(strLUN, 2, 1) & Mid(strLUN, 4, 1) & Mid(strLUN, 6, 1) & Mid(strLUN, 8, 1)
        Case 2
                strPW = Mid(strLUN, 2, 1) & Mid(strLUN, 4, 1) & Mid(strLUN, 6, 1) & Mid(strLUN, 8, 1) & Mid(strLUN, 1, 1) & Mid(strLUN, 3, 1) & Mid(strLUN, 5, 1) & Mid(strLUN, 7, 1)
        Case 3
                strPW = Mid(strLUN, 3, 1) & Mid(strLUN, 5, 1) & Mid(strLUN, 7, 1) & Mid(strLUN, 1, 1) & Mid(strLUN, 4, 1) & Mid(strLUN, 6, 1) & Mid(strLUN, 8, 1) & Mid(strLUN, 2, 1)
        Case 4
                strPW = Mid(strLUN, 4, 1) & Mid(strLUN, 6, 1) & Mid(strLUN, 8, 1) & Mid(strLUN, 2, 1) & Mid(strLUN, 5, 1) & Mid(strLUN, 7, 1) & Mid(strLUN, 1, 1) & Mid(strLUN, 3, 1)
        Case 5
                strPW = Mid(strLUN, 5, 1) & Mid(strLUN, 7, 1) & Mid(strLUN, 1, 1) & Mid(strLUN, 3, 1) & Mid(strLUN, 6, 1) & Mid(strLUN, 8, 1) & Mid(strLUN, 2, 1) & Mid(strLUN, 4, 1)
        Case 6
                strPW = Mid(strLUN, 6, 1) & Mid(strLUN, 8, 1) & Mid(strLUN, 2, 1) & Mid(strLUN, 4, 1) & Mid(strLUN, 7, 1) & Mid(strLUN, 1, 1) & Mid(strLUN, 3, 1) & Mid(strLUN, 5, 1)
        Case 7
                strPW = Mid(strLUN, 7, 1) & Mid(strLUN, 1, 1) & Mid(strLUN, 3, 1) & Mid(strLUN, 5, 1) & Mid(strLUN, 8, 1) & Mid(strLUN, 2, 1) & Mid(strLUN, 4, 1) & Mid(strLUN, 6, 1)
        Case 8
                strPW = Mid(strLUN, 8, 1) & Mid(strLUN, 2, 1) & Mid(strLUN, 4, 1) & Mid(strLUN, 6, 1) & Mid(strLUN, 1, 1) & Mid(strLUN, 3, 1) & Mid(strLUN, 5, 1) & Mid(strLUN, 7, 1)
        Case 9
                strPW = Mid(strLUN, 8, 1) & Mid(strLUN, 6, 1) & Mid(strLUN, 4, 1) & Mid(strLUN, 2, 1) & Mid(strLUN, 7, 1) & Mid(strLUN, 5, 1) & Mid(strLUN, 3, 1) & Mid(strLUN, 1, 1)
        Case 10
                strPW = Mid(strLUN, 7, 1) & Mid(strLUN, 5, 1) & Mid(strLUN, 3, 1) & Mid(strLUN, 1, 1) & Mid(strLUN, 8, 1) & Mid(strLUN, 6, 1) & Mid(strLUN, 4, 1) & Mid(strLUN, 2, 1)
        Case 10
                strPW = Mid(strLUN, 2, 1) & Mid(strLUN, 8, 1) & Mid(strLUN, 6, 1) & Mid(strLUN, 4, 1) & Mid(strLUN, 1, 1) & Mid(strLUN, 7, 1) & Mid(strLUN, 5, 1) & Mid(strLUN, 3, 1)
        Case 12
                strPW = Mid(strLUN, 3, 1) & Mid(strLUN, 1, 1) & Mid(strLUN, 7, 1) & Mid(strLUN, 5, 1) & Mid(strLUN, 2, 1) & Mid(strLUN, 8, 1) & Mid(strLUN, 6, 1) & Mid(strLUN, 4, 1)
    End Select
    
    GetPassWord = strPW
    
End Function
