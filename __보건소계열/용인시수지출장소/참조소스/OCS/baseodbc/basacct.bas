Attribute VB_Name = "BASAccount"
Option Explicit
                                      
Dim SQLDEF                  As String

Global OBON(65)             As Integer  '외래본인부담율
Global IBON(65)             As Integer  '입원본인부담율
Global JOJE(20)             As Integer  '투약조제료금액
Global GISUL(9)             As Integer  '병원가산율
Global NIGHT(9)             As Integer  '야간/공휴 가산
Global NIGHT_ILBAN(9)       As Integer  '야간/공휴 일반 가산
Global NIGHT_25(9)          As Integer  '마취 야간/공휴가산
Global GAMEK(30)            As Integer  '감액율
Global GAMEK_JIN(30)        As Integer  '감액율 (진찰료)

Global OLD_GISUL(9)         As Integer  'OLD TABLE
Global OLD_NIGHT(9)         As Integer
Global OLD_NIGHT_IL(9)      As Integer
Global OLD_NIGHT_25(9)      As Integer
    
Global GISUL_DATE           As String
Global NIGHT_DATE           As String
Global NGTIL_DATE           As String
Global NGT25_DATE           As String


Sub BAS_GAMEK()

    Dim I, J, K     As Integer
    Dim strDate     As String
    Dim rs100       As rdoResultset
    
    SQLDEF = " SELECT StartDate"
    SQLDEF = SQLDEF & "  FROM TW_MIS_PMPA.TWBAS_ACCTRULE"
    SQLDEF = SQLDEF & " WHERE IDname = 'GAMEK'"
    SQLDEF = SQLDEF & " GROUP BY StartDate"
    SQLDEF = SQLDEF & " ORDER BY StartDate DESC"

    Result = RdoOpenSet(rs100, SQLDEF)
    
    If Rowindicator = 0 Then Exit Sub
    strDate = RdoGetString(rs100, "StartDate", 0)
    
    rs100.Close: Set rs100 = Nothing
    
    SQLDEF = " SELECT  IArray, RateValue"
    SQLDEF = SQLDEF & "  FROM TW_MIS_PMPA.TWBAS_ACCTRULE"
    SQLDEF = SQLDEF & " WHERE IDname = 'GAMEK'"
    SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & Date_Format(strDate) & "','YYYY-MM-DD') "

    Result = RdoOpenSet(rs100, SQLDEF)
    
    For I = 0 To (Rowindicator - 1)
        J = RdoGetNumber(rs100, "IArray", I)
        K = RdoGetNumber(rs100, "RateValue", I)
        GAMEK(J) = K
    Next I

    rs100.Close: Set rs100 = Nothing

End Sub

Sub BAS_GAMEK_JIN()
    Dim I, J, K     As Integer
    Dim strDate     As String
    Dim rs100       As rdoResultset
    
    SQLDEF = " SELECT StartDate"
    SQLDEF = SQLDEF & "  FROM TWBAS_ACCTRULE"
    SQLDEF = SQLDEF & " WHERE IDname = 'GAMEK_JIN'"
    SQLDEF = SQLDEF & " GROUP BY StartDate"
    SQLDEF = SQLDEF & " ORDER BY StartDate DESC"

    Result = RdoOpenSet(rs100, SQLDEF)
    
    If Rowindicator = 0 Then Exit Sub
    strDate = RdoGetString(rs100, "StartDate", 0)

    rs100.Close: Set rs100 = Nothing
    
    SQLDEF = " SELECT IArray, RateValue"
    SQLDEF = SQLDEF & "  FROM TWBAS_ACCTRULE"
    SQLDEF = SQLDEF & " WHERE IDname = 'GAMEK_JIN'"
    SQLDEF = SQLDEF & "   AND StartDate = '" & strDate & "' "

    Result = RdoOpenSet(rs100, SQLDEF)
    
    For I = 0 To (Rowindicator - 1)
        J = RdoGetNumber(rs100, "IArray", I)
        K = RdoGetNumber(rs100, "RateValue", I)
        GAMEK_JIN(J) = K
    Next I

    rs100.Close: Set rs100 = Nothing


End Sub

Sub BAS_GISUL()
    Dim I, J, K     As Integer
    Dim strDate1    As String
    Dim strDate2    As String
    Dim rs100       As rdoResultset
    Dim rs200       As rdoResultset
    
    GISUL_DATE = ""
    SQLDEF = " SELECT TO_CHAR(StartDate, 'YYYY-MM-DD') SDATE"
    SQLDEF = SQLDEF & "  FROM TW_MIS_PMPA.TWBAS_ACCTRULE"
    SQLDEF = SQLDEF & " WHERE IDname = 'GISUL'"
    SQLDEF = SQLDEF & "   AND StartDate < SYSDATE"
    SQLDEF = SQLDEF & " GROUP BY StartDate"
    SQLDEF = SQLDEF & " ORDER BY StartDate DESC"

    Result = RdoOpenSet(rs100, SQLDEF)
    
    If Rowindicator = 0 Then Exit Sub
    strDate1 = RdoGetString(rs100, "SDATE", 0)
    If Rowindicator > 1 Then strDate2 = RdoGetString(rs100, "SDATE", 1)
    
    rs100.Close: Set rs100 = Nothing
    
    SQLDEF = " SELECT IArray, RateValue"
    SQLDEF = SQLDEF & "  FROM TW_MIS_PMPA.TWBAS_ACCTRULE"
    SQLDEF = SQLDEF & " WHERE IDname = 'GISUL'"
    SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & strDate1 & "','YYYY-MM-DD') "
    
    Result = RdoOpenSet(rs100, SQLDEF)
    
    For I = 0 To (Rowindicator - 1)
        J = RdoGetNumber(rs100, "IArray", I)
        K = RdoGetNumber(rs100, "RateValue", I)
        GISUL(J) = K
    Next I

    If strDate2 > "" Then
        GISUL_DATE = strDate1           '1998.6.1   son
       'GISUL_DATE = strDate2
        SQLDEF = " SELECT IArray, RateValue"
        SQLDEF = SQLDEF & "  FROM TW_MIS_PMPA.TWBAS_ACCTRULE"
        SQLDEF = SQLDEF & " WHERE IDname = 'GISUL'"
        SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & strDate2 & "','YYYY-MM-DD') "
    
        Result = RdoOpenSet(rs200, SQLDEF)
        
        For I = 0 To (Rowindicator - 1)
            J = RdoGetNumber(rs200, "IArray", I)
            K = RdoGetNumber(rs200, "RateValue", I)
            OLD_GISUL(J) = K
        Next I
        
        rs200.Close: Set rs200 = Nothing
    End If

    rs100.Close: Set rs100 = Nothing

End Sub

Sub BAS_IPD_BON()
    Dim I, J, K     As Integer
    Dim strDate     As String
    Dim rs100       As rdoResultset
    
    SQLDEF = " SELECT TO_CHAR(StartDate,'YYYY-MM-DD') StartDate "
    SQLDEF = SQLDEF & "  FROM TW_MIS_PMPA.TWBAS_ACCTRULE"
    SQLDEF = SQLDEF & " WHERE IDname = 'IPD_BON'"
    SQLDEF = SQLDEF & "   AND StartDate < SYSDATE"
    SQLDEF = SQLDEF & " GROUP BY StartDate"
    SQLDEF = SQLDEF & " ORDER BY StartDate DESC"

    Result = RdoOpenSet(rs100, SQLDEF)
    
    If Rowindicator = 0 Then Exit Sub
    strDate = RdoGetString(rs100, "StartDate", 0)

    rs100.Close: Set rs100 = Nothing
    
    SQLDEF = " SELECT IArray, RateValue"
    SQLDEF = SQLDEF & "  FROM TW_MIS_PMPA.TWBAS_ACCTRULE"
    SQLDEF = SQLDEF & " WHERE IDname = 'IPD_BON'"
    SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & strDate & "','YYYY-MM-DD') "

    Result = RdoOpenSet(rs100, SQLDEF)
    For I = 0 To (Rowindicator - 1)
        J = RdoGetNumber(rs100, "IArray", I)
        K = RdoGetNumber(rs100, "RateValue", I)
        IBON(J) = K
    Next I

    rs100.Close: Set rs100 = Nothing

End Sub

Sub BAS_JOJE()
    Dim I, J, K     As Integer
    Dim strDate     As String
    Dim rs100       As rdoResultset
    
    SQLDEF = " SELECT StartDate"
    SQLDEF = SQLDEF & "  FROM TWBAS_ACCTRULE"
    SQLDEF = SQLDEF & " WHERE IDname = 'JOJE'"
    SQLDEF = SQLDEF & " GROUP BY StartDate"
    SQLDEF = SQLDEF & " ORDER BY StartDate DESC"

    Result = RdoOpenSet(rs100, SQLDEF)
    
    If Rowindicator = 0 Then Exit Sub
    strDate = RdoGetString(rs100, "StartDate", 0)

    rs100.Close: Set rs100 = Nothing
    
    SQLDEF = " SELECT IArray, RateValue"
    SQLDEF = SQLDEF & "  FROM TWBAS_ACCTRULE"
    SQLDEF = SQLDEF & " WHERE IDname = 'JOJE'"
    SQLDEF = SQLDEF & "   AND StartDate = '" & strDate & "' "

    Result = RdoOpenSet(rs100, SQLDEF)
    
    For I = 0 To (Rowindicator - 1)
        J = RdoGetNumber(rs100, "IArray", I)
        K = RdoGetNumber(rs100, "RateValue", I)
        JOJE(J) = K
    Next I

    rs100.Close: Set rs100 = Nothing


End Sub

Sub BAS_NIGHT()
    
    Dim I, J, K         As Integer
    Dim strDate1        As String
    Dim strDate2        As String
    Dim rs100       As rdoResultset
    Dim rs200       As rdoResultset
    
    GoSub Read_NIGHT        '보험 심야가산
    GoSub Read_NGTIL        '일반 심야가산
    
    Exit Sub
    

'/-------------------------------------------------------------------------------------------/

Read_NIGHT:
    
    NIGHT_DATE = ""
    SQLDEF = " SELECT TO_CHAR(StartDate, 'YYYY-MM-DD') SDATE"
    SQLDEF = SQLDEF & "  FROM TW_MIS_PMPA.TWBAS_ACCTRULE"
    SQLDEF = SQLDEF & " WHERE IDname = 'NIGHT'"
    SQLDEF = SQLDEF & "   AND StartDate < SYSDATE"
    SQLDEF = SQLDEF & " GROUP BY StartDate"
    SQLDEF = SQLDEF & " ORDER BY StartDate DESC"

    Result = RdoOpenSet(rs100, SQLDEF)
    
    If Rowindicator = 0 Then Exit Sub
    strDate1 = RdoGetString(rs100, "SDATE", 0)
    If Rowindicator > 1 Then strDate2 = RdoGetString(rs100, "SDATE", 1)
    
    rs100.Close: Set rs100 = Nothing
    
    SQLDEF = " SELECT IArray, RateValue"
    SQLDEF = SQLDEF & "  FROM TW_MIS_PMPA.TWBAS_ACCTRULE"
    SQLDEF = SQLDEF & " WHERE IDname = 'NIGHT'"
    SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & strDate1 & "','YYYY-MM-DD') "

    Result = RdoOpenSet(rs100, SQLDEF)
    
    For I = 0 To (Rowindicator - 1)
        J = RdoGetNumber(rs100, "IArray", I)
        K = RdoGetNumber(rs100, "RateValue", I)
        NIGHT(J) = K
    Next I

    If strDate2 > "" Then
        NIGHT_DATE = strDate2
        SQLDEF = " SELECT IArray, RateValue"
        SQLDEF = SQLDEF & "  FROM TW_MIS_PMPA.TWBAS_ACCTRULE"
        SQLDEF = SQLDEF & " WHERE IDname = 'NIGHT'"
        SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & strDate2 & "','YYYY-MM-DD') "
    
        Result = RdoOpenSet(rs200, SQLDEF)
        For I = 0 To (Rowindicator - 1)
            J = RdoGetNumber(rs200, "IArray", I)
            K = RdoGetNumber(rs200, "RateValue", I)
            OLD_NIGHT(J) = K
        Next I
        
        rs200.Close: Set rs200 = Nothing
    End If

    rs100.Close: Set rs100 = Nothing

    Return
    
'/-------------------------------------------------------------------------------------------/

Read_NGTIL:
    
    NGTIL_DATE = ""
    SQLDEF = " SELECT TO_CHAR(StartDate, 'YYYY-MM-DD') SDATE"
    SQLDEF = SQLDEF & "  FROM TW_MIS_PMPA.TWBAS_ACCTRULE"
    SQLDEF = SQLDEF & " WHERE IDname = 'NIGHT_ILBAN'"
    SQLDEF = SQLDEF & "   AND StartDate < SYSDATE"
    SQLDEF = SQLDEF & " GROUP BY StartDate"
    SQLDEF = SQLDEF & " ORDER BY StartDate DESC"

    Result = RdoOpenSet(rs100, SQLDEF)
    
    If Rowindicator = 0 Then Exit Sub
    strDate1 = RdoGetString(rs100, "SDATE", 0)
    If Rowindicator > 1 Then strDate2 = RdoGetString(rs100, "SDATE", 1)
    
    rs100.Close: Set rs100 = Nothing
    
    SQLDEF = " SELECT IArray, RateValue"
    SQLDEF = SQLDEF & "  FROM TW_MIS_PMPA.TWBAS_ACCTRULE"
    SQLDEF = SQLDEF & " WHERE IDname = 'NIGHT_ILBAN'"
    SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & strDate1 & "','YYYY-MM-DD') "

    Result = RdoOpenSet(rs100, SQLDEF)
    
    For I = 0 To (Rowindicator - 1)
        J = RdoGetNumber(rs100, "IArray", I)
        K = RdoGetNumber(rs100, "RateValue", I)
        NIGHT_ILBAN(J) = K
    Next I

    If strDate2 > "" Then
        NIGHT_DATE = strDate2
        SQLDEF = " SELECT IArray, RateValue"
        SQLDEF = SQLDEF & "  FROM TW_MIS_PMPA.TWBAS_ACCTRULE"
        SQLDEF = SQLDEF & " WHERE IDname = 'NIGHT_ILBAN'"
        SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & strDate2 & "','YYYY-MM-DD') "
    
        Result = RdoOpenSet(rs200, SQLDEF)
        
        For I = 0 To (Rowindicator - 1)
            J = RdoGetNumber(rs200, "IArray", I)
            K = RdoGetNumber(rs200, "RateValue", I)
            OLD_NIGHT_IL(J) = K
        Next I
        
        rs200.Close: Set rs200 = Nothing
    End If

    rs100.Close: Set rs100 = Nothing

    Return
    
End Sub

Sub BAS_NIGHT_25()

    Dim I, J, K     As Integer
    Dim strDate1    As String
    Dim strDate2    As String
    Dim rs100       As rdoResultset
    Dim rs200       As rdoResultset
    
    NGT25_DATE = ""
    SQLDEF = " SELECT TO_CHAR(StartDate, 'YYYY-MM-DD') SDATE"
    SQLDEF = SQLDEF & "  FROM TW_MIS_PMPA.TWBAS_ACCTRULE"
    SQLDEF = SQLDEF & " WHERE IDname = 'NIGHT_25'"
    SQLDEF = SQLDEF & "   AND StartDate < SYSDATE"
    SQLDEF = SQLDEF & " GROUP BY StartDate"
    SQLDEF = SQLDEF & " ORDER BY StartDate DESC"

    Result = RdoOpenSet(rs100, SQLDEF)
    
    If Rowindicator = 0 Then Exit Sub
    strDate1 = RdoGetString(rs100, "SDATE", 0)
    If Rowindicator > 1 Then strDate2 = RdoGetString(rs100, "SDATE", 1)
    
    rs100.Close: Set rs100 = Nothing
    
    SQLDEF = " SELECT IArray, RateValue"
    SQLDEF = SQLDEF & "  FROM TW_MIS_PMPA.TWBAS_ACCTRULE"
    SQLDEF = SQLDEF & " WHERE IDname = 'NIGHT_25'"
    SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & strDate1 & "','YYYY-MM-DD') "

    Result = RdoOpenSet(rs100, SQLDEF)
    
    For I = 0 To (Rowindicator - 1)
        J = RdoGetNumber(rs100, "IArray", I)
        K = RdoGetNumber(rs100, "RateValue", I)
        NIGHT_25(J) = K
    Next I

    If strDate2 > "" Then
        NGT25_DATE = strDate2
        SQLDEF = " SELECT IArray, RateValue"
        SQLDEF = SQLDEF & "  FROM TW_MIS_PMPA.TWBAS_ACCTRULE"
        SQLDEF = SQLDEF & " WHERE IDname = 'NIGHT_25'"
        SQLDEF = SQLDEF & "   AND StartDate = TO_DATE('" & strDate2 & "','YYYY-MM-DD') "
    
        Result = RdoOpenSet(rs200, SQLDEF)
        For I = 0 To (Rowindicator - 1)
            J = RdoGetNumber(rs200, "IArray", I)
            K = RdoGetNumber(rs200, "RateValue", I)
            OLD_NIGHT_25(J) = K
        Next I
        
        rs200.Close: Set rs200 = Nothing
    End If

    rs100.Close: Set rs100 = Nothing

End Sub

Sub BAS_OPD_BON()
    Dim I, J, K     As Integer
    Dim strDate     As String
    Dim rs100       As rdoResultset
    
    SQLDEF = " SELECT StartDate"
    SQLDEF = SQLDEF & "  FROM TWBAS_ACCTRULE"
    SQLDEF = SQLDEF & " WHERE IDname = 'OPD_BON'"
    SQLDEF = SQLDEF & "   AND StartDate < SYSDATE"
    SQLDEF = SQLDEF & " GROUP BY StartDate"
    SQLDEF = SQLDEF & " ORDER BY StartDate DESC"

    Result = RdoOpenSet(rs100, SQLDEF)
    
    If Rowindicator = 0 Then Exit Sub
    strDate = RdoGetString(rs100, "StartDate", 0)

    rs100.Close: Set rs100 = Nothing
    
    SQLDEF = " SELECT IArray, RateValue"
    SQLDEF = SQLDEF & "  FROM TWBAS_ACCTRULE"
    SQLDEF = SQLDEF & " WHERE IDname = 'OPD_BON'"
    SQLDEF = SQLDEF & "   AND StartDate = '" & strDate & "' "

    Result = RdoOpenSet(rs100, SQLDEF)
    For I = 0 To (Rowindicator - 1)
        J = RdoGetNumber(rs100, "IArray", I)
        K = RdoGetNumber(rs100, "RateValue", I)
        OBON(J) = K
    Next I

    rs100.Close: Set rs100 = Nothing
    
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
        Case Else:  Valstr = "0"
    End Select


End Function

