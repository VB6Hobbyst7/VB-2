Attribute VB_Name = "BASAccount"
Option Explicit
                                      
Dim SQLDEF              As String

Global OBON(55)         As Integer  '외래본인부담율
Global IBON(55)         As Integer  '입원본인부담율
Global JOJE(20)         As Integer  '투약조제료금액
Global GISUL(9)         As Integer  '병원가산율
Global NIGHT(9)         As Integer  '야간/공휴 가산
Global NIGHT_22(9)      As Integer  '마취 야간/공휴가산
Global GAMEK(30)        As Integer  '감액율
Global GAMEK_JIN(30)    As Integer  '감액율 (진찰료)

Global OLD_GISUL(9)     As Integer  'OLD TABLE
Global OLD_NIGHT(9)     As Integer
Global OLD_NIGHT_22(9)  As Integer
    
Global GISUL_DATE       As String
Global NIGHT_DATE       As String
Global NGT22_DATE       As String

Global ROOM_GAM_AMT     As Currency '실료차 감액 기준금액
Sub BAS_GAMEK()
    Dim i, j, k     As Integer
    Dim strDate     As String

    SQLDEF = "FOR ALL  SELECT TO_CHAR(StartDate, 'YYYY-MM-DD') SDATE "
    SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
    SQLDEF = SQLDEF & " WHERE IDname = 'GAMEK'"
    SQLDEF = SQLDEF & " GROUP BY StartDate"
    SQLDEF = SQLDEF & " ORDER BY StartDate DESC"

    Result = dosql("open scope")
    Result = dosql(SQLDEF)
    
    If RowIndicator = 0 Then Exit Sub
    strDate = GlueGetString("SDATE", 0)

    
    SQLDEF = "FOR ALL  SELECT ArrayClass, RateValue"
    SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
    SQLDEF = SQLDEF & " WHERE IDname = 'GAMEK'"
    SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & strDate & "', 'yyyy-mm-dd') "

    Result = dosql(SQLDEF)
    For i = 0 To (RowIndicator - 1)
        j = GlueGetNumber("ArrayClass", i)
        k = GlueGetNumber("RateValue", i)
        GAMEK(j) = k
    Next i

    Result = dosql("close scope")

End Sub

Sub BAS_GAMEK_JIN()
    Dim i, j, k     As Integer
    Dim strDate     As String

    SQLDEF = "FOR ALL  SELECT TO_CHAR(StartDate, 'YYYY-MM-DD') SDATE "
    SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
    SQLDEF = SQLDEF & " WHERE IDname = 'GAMEK_JIN'"
    SQLDEF = SQLDEF & " GROUP BY StartDate"
    SQLDEF = SQLDEF & " ORDER BY StartDate DESC"

    Result = dosql("open scope")
    Result = dosql(SQLDEF)
    
    If RowIndicator = 0 Then Exit Sub
    strDate = GlueGetString("SDATE", 0)

    
    SQLDEF = "FOR ALL  SELECT ArrayClass, RateValue"
    SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
    SQLDEF = SQLDEF & " WHERE IDname = 'GAMEK_JIN'"
    SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & strDate & "', 'yyyy-mm-dd') "

    Result = dosql(SQLDEF)
    For i = 0 To (RowIndicator - 1)
        j = GlueGetNumber("ArrayClass", i)
        k = GlueGetNumber("RateValue", i)
        GAMEK_JIN(j) = k
    Next i

    Result = dosql("close scope")


End Sub

Sub Bas_GISUL()
    Dim i, j, k     As Integer
    Dim StrDate1    As String
    Dim StrDate2    As String

    GISUL_DATE = ""
    SQLDEF = "FOR ALL  SELECT TO_CHAR(StartDate, 'YYYY-MM-DD') SDATE"
    SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
    SQLDEF = SQLDEF & " WHERE IDname = 'GISUL'"
    SQLDEF = SQLDEF & "   AND StartDate < SYSDATE"
    SQLDEF = SQLDEF & " GROUP BY StartDate"
    SQLDEF = SQLDEF & " ORDER BY StartDate DESC"

    Result = dosql("open scope")
    Result = dosql(SQLDEF)
    
    If RowIndicator = 0 Then Exit Sub
    StrDate1 = GlueGetString("SDATE", 0)
    If RowIndicator > 1 Then StrDate2 = GlueGetString("SDATE", 1)
    
    SQLDEF = "FOR ALL  SELECT ArrayClass, RateValue"
    SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
    SQLDEF = SQLDEF & " WHERE IDname = 'GISUL'"
    SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & StrDate1 & "', 'yyyy-mm-dd') "

    Result = dosql(SQLDEF)
    For i = 0 To (RowIndicator - 1)
        j = GlueGetNumber("ArrayClass", i)
        k = GlueGetNumber("RateValue", i)
        GISUL(j) = k
    Next i

    If StrDate2 > "" Then
        GISUL_DATE = StrDate2
        SQLDEF = "FOR ALL  SELECT ArrayClass, RateValue"
        SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
        SQLDEF = SQLDEF & " WHERE IDname = 'GISUL'"
        SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & StrDate2 & "', 'yyyy-mm-dd') "
    
        Result = dosql(SQLDEF)
        For i = 0 To (RowIndicator - 1)
            j = GlueGetNumber("ArrayClass", i)
            k = GlueGetNumber("RateValue", i)
            OLD_GISUL(j) = k
        Next i
    End If

    Result = dosql("close scope")

End Sub

Sub BAS_IPD_BON()
    Dim i, j, k     As Integer
    Dim strDate     As String

    SQLDEF = "FOR ALL  SELECT TO_CHAR(StartDate, 'YYYY-MM-DD') SDATE"
    SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
    SQLDEF = SQLDEF & " WHERE IDname = 'IPD_BON'"
    SQLDEF = SQLDEF & "   AND StartDate < SYSDATE"
    SQLDEF = SQLDEF & " GROUP BY StartDate"
    SQLDEF = SQLDEF & " ORDER BY StartDate DESC"

    Result = dosql("open scope")
    Result = dosql(SQLDEF)
    
    If RowIndicator = 0 Then Exit Sub
    strDate = GlueGetString("SDATE", 0)

    
    SQLDEF = "FOR ALL  SELECT ArrayClass, RateValue"
    SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
    SQLDEF = SQLDEF & " WHERE IDname = 'IPD_BON'"
    SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & strDate & "', 'yyyy-mm-dd') "

    Result = dosql(SQLDEF)
    For i = 0 To (RowIndicator - 1)
        j = GlueGetNumber("ArrayClass", i)
        k = GlueGetNumber("RateValue", i)
        IBON(j) = k
    Next i

    Result = dosql("close scope")

End Sub

Sub BAS_JOJE()
    Dim i, j, k     As Integer
    Dim strDate     As String

    SQLDEF = "FOR ALL  SELECT TO_CHAR(StartDate, 'YYYY-MM-DD') SDATE "
    SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
    SQLDEF = SQLDEF & " WHERE IDname = 'JOJE'"
    SQLDEF = SQLDEF & " GROUP BY StartDate"
    SQLDEF = SQLDEF & " ORDER BY StartDate DESC"

    Result = dosql("open scope")
    Result = dosql(SQLDEF)
    
    If RowIndicator = 0 Then Exit Sub
    strDate = GlueGetString("SDATE", 0)

    
    SQLDEF = "FOR ALL  SELECT ArrayClass, RateValue"
    SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
    SQLDEF = SQLDEF & " WHERE IDname = 'JOJE'"
    SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & strDate & "', 'yyyy-mm-dd') "

    Result = dosql(SQLDEF)
    For i = 0 To (RowIndicator - 1)
        j = GlueGetNumber("ArrayClass", i)
        k = GlueGetNumber("RateValue", i)
        JOJE(j) = k
    Next i

    Result = dosql("close scope")


End Sub

Sub BAS_NIGHT()
    Dim i, j, k     As Integer
    Dim StrDate1    As String
    Dim StrDate2    As String

    NIGHT_DATE = ""
    SQLDEF = "FOR ALL  SELECT TO_CHAR(StartDate, 'YYYY-MM-DD') SDATE"
    SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
    SQLDEF = SQLDEF & " WHERE IDname = 'NIGHT'"
    SQLDEF = SQLDEF & "   AND StartDate < SYSDATE"
    SQLDEF = SQLDEF & " GROUP BY StartDate"
    SQLDEF = SQLDEF & " ORDER BY StartDate DESC"

    Result = dosql("open scope")
    Result = dosql(SQLDEF)
    
    If RowIndicator = 0 Then Exit Sub
    StrDate1 = GlueGetString("SDATE", 0)
    If RowIndicator > 1 Then StrDate2 = GlueGetString("SDATE", 1)
    
    SQLDEF = "FOR ALL  SELECT ArrayClass, RateValue"
    SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
    SQLDEF = SQLDEF & " WHERE IDname = 'NIGHT'"
    SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & StrDate1 & "', 'yyyy-mm-dd') "

    Result = dosql(SQLDEF)
    For i = 0 To (RowIndicator - 1)
        j = GlueGetNumber("ArrayClass", i)
        k = GlueGetNumber("RateValue", i)
        NIGHT(j) = k
    Next i

    If StrDate2 > "" Then
        NIGHT_DATE = StrDate2
        SQLDEF = "FOR ALL  SELECT ArrayClass, RateValue"
        SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
        SQLDEF = SQLDEF & " WHERE IDname = 'NIGHT'"
        SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & StrDate2 & "', 'yyyy-mm-dd') "
    
        Result = dosql(SQLDEF)
        For i = 0 To (RowIndicator - 1)
            j = GlueGetNumber("ArrayClass", i)
            k = GlueGetNumber("RateValue", i)
            OLD_NIGHT(j) = k
        Next i
    End If

    Result = dosql("close scope")


End Sub

Sub BAS_NIGHT_22()
    Dim i, j, k     As Integer
    Dim StrDate1    As String
    Dim StrDate2    As String

    NGT22_DATE = ""
    SQLDEF = "FOR ALL  SELECT TO_CHAR(StartDate, 'YYYY-MM-DD') SDATE"
    SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
    SQLDEF = SQLDEF & " WHERE IDname = 'NIGHT_22'"
    SQLDEF = SQLDEF & "   AND StartDate < SYSDATE"
    SQLDEF = SQLDEF & " GROUP BY StartDate"
    SQLDEF = SQLDEF & " ORDER BY StartDate DESC"

    Result = dosql("open scope")
    Result = dosql(SQLDEF)
    
    If RowIndicator = 0 Then Exit Sub
    StrDate1 = GlueGetString("SDATE", 0)
    If RowIndicator > 1 Then StrDate2 = GlueGetString("SDATE", 1)
    
    SQLDEF = "FOR ALL  SELECT ArrayClass, RateValue"
    SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
    SQLDEF = SQLDEF & " WHERE IDname = 'NIGHT_22'"
    SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & StrDate1 & "', 'yyyy-mm-dd') "

    Result = dosql(SQLDEF)
    For i = 0 To (RowIndicator - 1)
        j = GlueGetNumber("ArrayClass", i)
        k = GlueGetNumber("RateValue", i)
        NIGHT_22(j) = k
    Next i

    If StrDate2 > "" Then
        NGT22_DATE = StrDate2
        SQLDEF = "FOR ALL  SELECT ArrayClass, RateValue"
        SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
        SQLDEF = SQLDEF & " WHERE IDname = 'NIGHT_22'"
        SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & StrDate2 & "', 'yyyy-mm-dd') "
    
        Result = dosql(SQLDEF)
        For i = 0 To (RowIndicator - 1)
            j = GlueGetNumber("ArrayClass", i)
            k = GlueGetNumber("RateValue", i)
            OLD_NIGHT_22(j) = k
        Next i
    End If

    Result = dosql("close scope")

End Sub

Sub BAS_OPD_BON()
    Dim i, j, k     As Integer
    Dim strDate     As String

    SQLDEF = "FOR ALL  SELECT TO_CHAR(StartDate, 'YYYY-MM-DD') SDATE "
    SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
    SQLDEF = SQLDEF & " WHERE IDname = 'OPD_BON'"
    SQLDEF = SQLDEF & "   AND StartDate < SYSDATE"
    SQLDEF = SQLDEF & " GROUP BY StartDate"
    SQLDEF = SQLDEF & " ORDER BY StartDate DESC"

    Result = dosql("open scope")
    Result = dosql(SQLDEF)
    
    If RowIndicator = 0 Then Exit Sub
    strDate = GlueGetString("SDATE", 0)

    
    SQLDEF = "FOR ALL  SELECT ArrayClass, RateValue"
    SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
    SQLDEF = SQLDEF & " WHERE IDname = 'OPD_BON'"
    SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & strDate & "', 'yyyy-mm-dd') "
    Result = dosql(SQLDEF)
    For i = 0 To (RowIndicator - 1)
        j = GlueGetNumber("ArrayClass", i)
        k = GlueGetNumber("RateValue", i)
        OBON(j) = k
    Next i

    Result = dosql("close scope")
End Sub

Sub BAS_ROOM_GAMEK()
    
    SQLDEF = "FOR 1 SELECT RateValue FROM BAS_ACCOUNT "
    SQLDEF = SQLDEF & "   WHERE IDNAME = 'ROOM_GAMEK'"
    
    Result = dosql(SQLDEF)
    
    ROOM_GAM_AMT = GlueGetNumber("RateValue", 0)
    
End Sub

Function Valnum(strValue As String) As Integer

    Select Case strValue
        Case "0":   Valnum = 0
        Case "1":   Valnum = 1
        Case "2":   Valnum = 2
        Case "3":   Valnum = 3
        Case "4":   Valnum = 4
        Case "5":   Valnum = 5
        Case "6":   Valnum = 6
        Case "7":   Valnum = 7
        Case "8":   Valnum = 8
        Case "9":   Valnum = 9
        Case "A":   Valnum = 10
        Case "B":   Valnum = 11
        Case "C":   Valnum = 12
        Case "D":   Valnum = 13
        Case "E":   Valnum = 14
        Case "F":   Valnum = 15
        Case "G":   Valnum = 16
        Case "H":   Valnum = 17
        Case "I":   Valnum = 18
        Case "J":   Valnum = 19
        Case "K":   Valnum = 20
        Case "L":   Valnum = 21
        Case "M":   Valnum = 22
        Case "N":   Valnum = 23
        Case "O":   Valnum = 24
        Case "P":   Valnum = 25
        Case "Q":   Valnum = 26
        Case "R":   Valnum = 27
        Case "S":   Valnum = 28
        Case "T":   Valnum = 29
        Case "U":   Valnum = 30
        Case "V":   Valnum = 31
        Case "W":   Valnum = 32
        Case "X":   Valnum = 33
        Case "Y":   Valnum = 34
        Case "Z":   Valnum = 35
        Case Else:  Valnum = 0
    End Select

End Function

Function Valstr(nValue As Integer) As String

    Select Case nValue
        Case 1:     Valstr = "1"
        Case 2:     Valstr = "2"
        Case 3:     Valstr = "3"
        Case 4:     Valstr = "4"
        Case 5:     Valstr = "5"
        Case 6:     Valstr = "6"
        Case 7:     Valstr = "7"
        Case 8:     Valstr = "8"
        Case 9:     Valstr = "9"
        Case 10:    Valstr = "A"
        Case 11:    Valstr = "B"
        Case 12:    Valstr = "C"
        Case 13:    Valstr = "D"
        Case 14:    Valstr = "E"
        Case 15:    Valstr = "F"
        Case 16:    Valstr = "G"
        Case 17:    Valstr = "H"
        Case 18:    Valstr = "I"
        Case 19:    Valstr = "J"
        Case 20:    Valstr = "K"
        Case 21:    Valstr = "L"
        Case 22:    Valstr = "M"
        Case 23:    Valstr = "N"
        Case 24:    Valstr = "O"
        Case 25:    Valstr = "P"
        Case 26:    Valstr = "Q"
        Case 27:    Valstr = "R"
        Case 28:    Valstr = "S"
        Case 29:    Valstr = "T"
        Case 30:    Valstr = "U"
        Case 31:    Valstr = "V"
        Case 32:    Valstr = "W"
        Case 33:    Valstr = "X"
        Case 34:    Valstr = "Y"
        Case 35:    Valstr = "Z"
        Case Else:  Valstr = "0"
    End Select


End Function
Sub BAS_Bi()
    Dim i, j, k     As Integer
    Dim strDate     As String

    SQLDEF = "FOR ALL  SELECT TO_CHAR(StartDate, 'YYYY-MM-DD') SDATE "
    SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
    SQLDEF = SQLDEF & " WHERE IDname = 'OPD_BON'"
    SQLDEF = SQLDEF & "   AND StartDate < SYSDATE"
    SQLDEF = SQLDEF & " GROUP BY StartDate"
    SQLDEF = SQLDEF & " ORDER BY StartDate DESC"

    Result = dosql("open scope")
    Result = dosql(SQLDEF)
    
    If RowIndicator = 0 Then Exit Sub
    strDate = GlueGetString("SDATE", 0)

    
    SQLDEF = "FOR ALL  SELECT ArrayClass, RatetEXT"
    SQLDEF = SQLDEF & "  FROM BAS_ACCOUNT"
    SQLDEF = SQLDEF & " WHERE IDname = 'BI'"
    SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & strDate & "', 'yyyy-mm-dd') "
    Result = dosql(SQLDEF)
    For i = 0 To (RowIndicator - 1)
        j = GlueGetNumber("ArrayClass", i)
        GstrBis(j) = GlueGetString("RATETEXT", i)
    Next i

    Result = dosql("close scope")
End Sub
