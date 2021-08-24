VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCS0301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private recordFiled01 As String
Private recordFiled02 As String
Private recordFiled03 As String
Private recordFiled04 As String
Private recordFiled05 As String
Private recordFiled06 As String
Private recordFiled07 As String
Private recordFiled08 As String
Private recordFiled09 As String
Private recordFiled10 As String
Private recordFiled11 As String
Private recordFiled12 As String
Private recordFiled13 As String
Private recordFiled14 As String
Private recordFiled15 As String
Private recordFiled16 As String
Private recordFiled17 As String
Private recordFiled18 As String
Private recordFiled19 As String
Private recordFiled20 As String
Private recordFiled21 As String
Private recordFiled22 As String
Private recordFiled23 As String
Private recordFiled24 As String
Private recordFiled25 As String
Private recordFiled26 As String
Private recordFiled27 As String
Private recordFiled28 As String
Private recordFiled29 As String


Private count As Integer        'query 결과 레코드 개수

Dim AdoConn     As New ADODB.Connection

Public Property Let LetrecordFiled01(ByVal data As String)
    recordFiled01 = data
End Property

Public Property Get GetrecordField01() As String
    GetrecordField01 = recordFiled01
End Property

Public Property Let LetrecordFiled02(ByVal data As String)
    recordFiled02 = data
End Property

Public Property Get GetrecordFiled02() As String
    GetrecordFiled02 = recordFiled02
End Property

Public Property Let LetrecordFiled03(ByVal data As String)
    recordFiled03 = data
End Property

Public Property Get GetrecordFiled03() As String
    GetrecordFiled03 = recordFiled03
End Property

Public Property Let LetrecordFiled04(ByVal data As String)
    recordFiled04 = data
End Property

Public Property Get GetrecordFiled04() As String
    GetrecordFiled04 = recordFiled04
End Property

Public Property Let LetrecordFiled05(ByVal data As String)
    recordFiled05 = data
End Property

Public Property Get GetrecordFiled05() As String
    GetrecordFiled05 = recordFiled05
End Property

Public Property Let LetrecordFiled06(ByVal data As String)
    recordFiled06 = data
End Property

Public Property Get GetrecordFiled06() As String
    GetrecordFiled06 = recordFiled06
End Property

Public Property Let LetrecordFiled07(ByVal data As String)
    recordFiled07 = data
End Property

Public Property Get GetrecordFiled07() As String
    GetrecordFiled07 = recordFiled07
End Property

Public Property Let LetrecordFiled08(ByVal data As String)
    recordFiled08 = data
End Property

Public Property Get GetrecordFiled08() As String
    GetrecordFiled08 = recordFiled08
End Property

Public Property Let LetrecordFiled09(ByVal data As String)
    recordFiled09 = data
End Property

Public Property Get GetrecordFiled09() As String
    GetrecordFiled09 = recordFiled09
End Property

Public Property Let LetrecordFiled10(ByVal data As String)
    recordFiled10 = data
End Property

Public Property Get GetrecordFiled10() As String
    GetrecordFiled10 = recordFiled10
End Property

Public Property Let LetrecordFiled11(ByVal data As String)
    recordFiled11 = data
End Property

Public Property Get GetrecordFiled11() As String
    GetrecordFiled11 = recordFiled11
End Property

Public Property Let LetrecordFiled12(ByVal data As String)
    recordFiled12 = data
End Property

Public Property Get GetrecordFiled12() As String
    GetrecordFiled12 = recordFiled12
End Property

Public Property Let LetrecordFiled13(ByVal data As String)
    recordFiled13 = data
End Property

Public Property Get GetrecordFiled13() As String
    GetrecordFiled13 = recordFiled13
End Property

Public Property Let LetrecordFiled14(ByVal data As String)
    recordFiled14 = data
End Property

Public Property Get GetrecordFiled14() As String
    GetrecordFiled14 = recordFiled14
End Property

Public Property Let LetrecordFiled15(ByVal data As String)
    recordFiled15 = data
End Property

Public Property Get GetrecordFiled15() As String
    GetrecordFiled15 = recordFiled15
End Property

Public Property Let LetrecordFiled16(ByVal data As String)
    recordFiled16 = data
End Property

Public Property Get GetrecordFiled16() As String
    GetrecordFiled16 = recordFiled16
End Property

Public Property Let LetrecordFiled17(ByVal data As String)
    recordFiled17 = data
End Property

Public Property Get GetrecordFiled17() As String
    GetrecordFiled17 = recordFiled17
End Property

Public Property Let LetrecordFiled18(ByVal data As String)
    recordFiled18 = data
End Property

Public Property Get GetrecordFiled18() As String
    GetrecordFiled18 = recordFiled18
End Property

Public Property Let LetrecordFiled19(ByVal data As String)
    recordFiled19 = data
End Property

Public Property Get GetrecordFiled19() As String
    GetrecordFiled19 = recordFiled19
End Property

Public Property Let LetrecordFiled20(ByVal data As String)
    recordFiled20 = data
End Property

Public Property Get GetrecordFiled20() As String
    GetrecordFiled20 = recordFiled20
End Property

Public Property Let LetrecordFiled21(ByVal data As String)
    recordFiled21 = data
End Property

Public Property Get GetrecordFiled21() As String
    GetrecordFiled21 = recordFiled21
End Property

Public Property Let LetrecordFiled22(ByVal data As String)
    recordFiled22 = data
End Property

Public Property Get GetrecordFiled22() As String
    GetrecordFiled22 = recordFiled22
End Property

Public Property Let LetrecordFiled23(ByVal data As String)
    recordFiled23 = data
End Property

Public Property Get GetrecordFiled23() As String
    GetrecordFiled23 = recordFiled23
End Property

Public Property Let LetrecordFiled24(ByVal data As String)
    recordFiled24 = data
End Property

Public Property Get GetrecordFiled24() As String
    GetrecordFiled24 = recordFiled24
End Property

Public Property Let LetrecordFiled25(ByVal data As String)
    recordFiled25 = data
End Property

Public Property Get GetrecordFiled25() As String
    GetrecordFiled25 = recordFiled25
End Property

Public Property Let LetrecordFiled26(ByVal data As String)
    recordFiled26 = data
End Property

Public Property Get GetrecordFiled26() As String
    GetrecordFiled26 = recordFiled26
End Property

Public Property Let LetrecordFiled27(ByVal data As String)
    recordFiled27 = data
End Property

Public Property Get GetrecordFiled27() As String
    GetrecordFiled27 = recordFiled27
End Property

Public Property Let LetrecordFiled28(ByVal data As String)
    recordFiled28 = data
End Property

Public Property Get GetrecordFiled28() As String
    GetrecordFiled28 = recordFiled28
End Property

Public Property Let LetrecordFiled29(ByVal data As String)
    recordFiled29 = data
End Property

Public Property Get GetrecordFiled29() As String
    GetrecordFiled29 = recordFiled29
End Property

Public Property Let Letcount(ByVal data As Integer)
    count = data
End Property

Public Property Get Getcount() As Integer
    Getcount = count
End Property

Public Sub Get_SlipCD()
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Recordset
    Dim sSql As String
    Dim j As Integer
        
    sSql = "SELECT * FROM PART"

    AdoConn.CursorType = adOpenStatic
    AdoConn.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoConn.RecordCount = 0 Then
        AdoConn.Close
        count = 0
        Exit Sub
    Else

        count = AdoConn.RecordCount

        AdoConn.MoveFirst

        For j = 1 To count
            recordFiled01 = recordFiled01 & AdoConn("PARTCD") & "|"
            recordFiled02 = recordFiled02 & AdoConn("PARTGBN") & "|"
            recordFiled03 = recordFiled03 & AdoConn("SLIPNM") & "|"
            
            AdoConn.MoveNext
        Next

    End If

    AdoConn.Close

    Exit Sub
ErrHandler:
    MsgBox "Get_SlipCD(AS03.DLL) - " & Err.Description

End Sub

Public Sub Get_SlipNm(Cd As String)

    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
        
    sSql = "SELECT SLIPNM" _
        & " FROM PART" _
        & " WHERE PARTCD = '" & Left$(Cd, 1) & "'" _
        & " AND PARTGBN = '" & Right$(Cd, 2) & "'"
    
    AdoConn.CursorType = adOpenStatic
    AdoConn.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoConn.RecordCount = 0 Then
        AdoConn.Close
        count = 0
        Exit Sub
    Else
        
        count = 1
        
        AdoConn.MoveFirst
        
        recordFiled01 = AdoConn("SLIPNM")
        
    End If
    
    AdoConn.Close
    
    Exit Sub
ErrHandler:
    MsgBox "Get_SlipNm(AS03.DLL) - " & Err.Description

End Sub

Public Sub Get_Item(slipCd As String, refOption As Integer)
    
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Recordset
    Dim sSql As String
    Dim j As Integer
    Dim sTestNmCfg As String
    
    sTestNmCfg = fGetCurTestItemNmCfg
    
    If sTestNmCfg = "T" Then
        sSql = "SELECT TESTITEMNM, SPECIMENCD, TESTITEMSEQ, SUBMCD" _
            & " FROM TESTITEM" _
            & " WHERE PARTCD = '" & Left(slipCd, 1) & "'" _
            & " AND PARTGBN = '" & Right(slipCd, 2) & "'" _
            & " AND PANJUNGGBN = '" & refOption & "'"
    ElseIf sTestNmCfg = "P" Then
        sSql = "SELECT PRINTNM, SPECIMENCD, TESTITEMSEQ, SUBMCD" _
            & " FROM TESTITEM" _
            & " WHERE PARTCD = '" & Left(slipCd, 1) & "'" _
            & " AND PARTGBN = '" & Right(slipCd, 2) & "'" _
            & " AND PANJUNGGBN = '" & refOption & "'"
    End If
    
    AdoConn.CursorType = adOpenStatic
    AdoConn.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"

    If AdoConn.RecordCount = 0 Then
        AdoConn.Close
        count = 0
        Exit Sub
    Else

        count = AdoConn.RecordCount

        AdoConn.MoveFirst

        For j = 1 To count
            recordFiled01 = recordFiled01 & AdoConn(0) & "|"
            recordFiled02 = recordFiled02 & AdoConn(1) & "|"
            recordFiled03 = recordFiled03 & AdoConn(2) & "|"
            recordFiled04 = recordFiled04 & AdoConn(3) & "|"

            AdoConn.MoveNext
        Next

    End If

    AdoConn.Close

    Exit Sub
ErrHandler:
    MsgBox "Get_Item(AS03.DLL) - " & Err.Description

End Sub

Public Sub Get_Contents(slipCd As String, _
                        sDate As String, _
                        eDate As String, _
                        판정Gbn As Integer, _
                        조회Opt As Integer, _
                        이상Opt As String, _
                        code As String, _
                        arrayNum As Integer)

    On Error GoTo ErrHandler

    Dim AdoConn As New ADODB.Recordset
    Dim sSql As String
    Dim j As Integer
    Dim tmpStr$
    Dim sTestNmCfg$
    
    sTestNmCfg = fGetCurTestItemNmCfg
    
    If sTestNmCfg = "T" Then
        sSql = "SELECT ABNORMAL.LABDATE, ABNORMAL.LABSEQ, TESTITEMNM, ABNORMAL.RESULT," _
            & " ABNORMALCD, REGNO, NAME, AGE, SEX, TESTITEM.REFGBN, REFLOWM, REFLOWF," _
            & " LOWERLIMITM, LOWERLIMITF, GRAYLOWERM, GRAYLOWERF, REFHIGHM, REFHIGHF," _
            & " UPPERLIMITM, UPPERLIMITF, GRAYUPPERM, GRAYUPPERF, UNIT, TESTITEM.PANICGBN," _
            & " PANICLOW, PANICHIGH, TESTITEM.DELTAGBN, DELTAVAL, OTHERFLAG" _
            & " FROM H_ABNORMAL AS ABNORMAL, " _
            & " H_RESULT AS RESULT, TESTITEM" _
            & " WHERE TESTITEM.PARTCD = '" & Left(slipCd, 1) & "'" _
            & " AND TESTITEM.PARTGBN = '" & Right(slipCd, 2) & "'" _
            & " AND ("
    ElseIf sTestNmCfg = "P" Then
        sSql = "SELECT ABNORMAL.LABDATE, ABNORMAL.LABSEQ, PRINTNM, ABNORMAL.RESULT," _
            & " ABNORMALCD, REGNO, NAME, AGE, SEX, TESTITEM.REFGBN, REFLOWM, REFLOWF," _
            & " LOWERLIMITM, LOWERLIMITF, GRAYLOWERM, GRAYLOWERF, REFHIGHM, REFHIGHF," _
            & " UPPERLIMITM, UPPERLIMITF, GRAYUPPERM, GRAYUPPERF, UNIT, TESTITEM.PANICGBN," _
            & " PANICLOW, PANICHIGH, TESTITEM.DELTAGBN, DELTAVAL, OTHERFLAG" _
            & " FROM H_ABNORMAL AS ABNORMAL, " _
            & " H_RESULT AS RESULT, TESTITEM" _
            & " WHERE TESTITEM.PARTCD = '" & Left(slipCd, 1) & "'" _
            & " AND TESTITEM.PARTGBN = '" & Right(slipCd, 2) & "'" _
            & " AND ("
    End If
    
    For j = 1 To arrayNum
        tmpStr = GetByOne(code, code)
        sSql = sSql & " (TESTITEM.SPECIMENCD = '" & Left(tmpStr, 3) & "'"
        sSql = sSql & " AND TESTITEM.TESTITEMSEQ = '" & Mid(tmpStr, 4, 3) & "'"
        sSql = sSql & " AND TESTITEM.SUBMCD = '" & Right(tmpStr, 4) & "')"

        If j <> arrayNum Then
            sSql = sSql & " OR "
        Else
            sSql = sSql & ") "
        End If
    Next

    sSql = sSql & " AND TESTITEM.PANJUNGGBN = '" & 판정Gbn & "'" _
                & " AND ABNORMAL.LABDATE BETWEEN '" & sDate & "' AND '" & eDate & "'" _
                & " AND ABNORMAL.PARTGBN = TESTITEM.PARTGBN" _
                & " AND ABNORMAL.SPECIMENCD = TESTITEM.SPECIMENCD" _
                & " AND ABNORMAL.TESTITEMSEQ = TESTITEM.TESTITEMSEQ" _
                & " AND ABNORMAL.SUBMCD = TESTITEM.SUBMCD"

    '판정구분과 조회옵션에 따른 SQL문 조건절
    If 판정Gbn = 1 Then
        If 조회Opt = 1 Then
            sSql = sSql & AndOp(이상Opt)
        Else
            sSql = sSql & OrOp(이상Opt)
        End If
    ElseIf 판정Gbn = 2 Then
        sSql = sSql & " AND ABNORMALCD = '16'"
    Else
        sSql = sSql & " AND ABNORMALCD = '32'"
    End If

    sSql = sSql & " AND RESULT.LABDATE = ABNORMAL.LABDATE" _
                & " AND RESULT.PARTGBN = TESTITEM.PARTGBN" _
                & " AND RESULT.LABSEQ = ABNORMAL.LABSEQ" _
                & " AND RESULT.SPECIMENCD = TESTITEM.SPECIMENCD" _
                & " AND RESULT.TESTITEMSEQ = TESTITEM.TESTITEMSEQ" _
                & " AND RESULT.SUBMCD = TESTITEM.SUBMCD" _
                & " ORDER BY ABNORMAL.LABDATE, ABNORMAL.LABSEQ, TESTITEM.DISPLAYORDER"

    AdoConn.CursorType = adOpenStatic
    AdoConn.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoConn.RecordCount = 0 Then
        AdoConn.Close
        count = 0
        Exit Sub
    Else

        count = AdoConn.RecordCount

        AdoConn.MoveFirst

        For j = 1 To count
            recordFiled01 = recordFiled01 & AdoConn(0) & "|"
            recordFiled02 = recordFiled02 & AdoConn(1) & "|"
            recordFiled03 = recordFiled03 & AdoConn(2) & "|"
            recordFiled04 = recordFiled04 & AdoConn(3) & "|"
            recordFiled05 = recordFiled05 & AdoConn(4) & "|"
            recordFiled06 = recordFiled06 & AdoConn(5) & "|"
            recordFiled07 = recordFiled07 & AdoConn(6) & "|"
            recordFiled08 = recordFiled08 & AdoConn(7) & "|"
            recordFiled09 = recordFiled09 & AdoConn(8) & "|"
            recordFiled10 = recordFiled10 & AdoConn(9) & "|"
            recordFiled11 = recordFiled11 & AdoConn(10) & "|"
            recordFiled12 = recordFiled12 & AdoConn(11) & "|"
            recordFiled13 = recordFiled13 & AdoConn(12) & "|"
            recordFiled14 = recordFiled14 & AdoConn(13) & "|"
            recordFiled15 = recordFiled15 & AdoConn(14) & "|"
            recordFiled16 = recordFiled16 & AdoConn(15) & "|"
            recordFiled17 = recordFiled17 & AdoConn(16) & "|"
            recordFiled18 = recordFiled18 & AdoConn(17) & "|"
            recordFiled19 = recordFiled19 & AdoConn(18) & "|"
            recordFiled20 = recordFiled20 & AdoConn(19) & "|"
            recordFiled21 = recordFiled21 & AdoConn(20) & "|"
            recordFiled22 = recordFiled22 & AdoConn(21) & "|"
            recordFiled23 = recordFiled23 & AdoConn(22) & "|"
            recordFiled24 = recordFiled24 & AdoConn(23) & "|"
            recordFiled25 = recordFiled25 & AdoConn(24) & "|"
            recordFiled26 = recordFiled26 & AdoConn(25) & "|"
            recordFiled27 = recordFiled27 & AdoConn(26) & "|"
            recordFiled28 = recordFiled28 & AdoConn(27) & "|"
            recordFiled29 = recordFiled29 & AdoConn(28) & "|"

            AdoConn.MoveNext
        Next

    End If

    AdoConn.Close

    Exit Sub
ErrHandler:
    MsgBox "Get_Item(AS03.DLL) - " & Err.Description

End Sub

Private Function AndOp(code As String) As String
    Dim temp As Integer
    temp = CInt(Left(code, 1))
    temp = temp + CInt(Mid(code, 2, 1)) * 2
    temp = temp + CInt(Mid(code, 3, 1)) * 4
    temp = temp + CInt(Right(code, 1)) * 8
    AndOp = " AND ABNORMALCD = '" & temp & "'"
End Function

Private Function OrOp(code As String) As String
    Dim tmpStr As String
    Dim commaPos As Long
    
    tmpStr = " AND ABNORMALCD IN ("
    
    If Left(code, 1) = "1" Then
        tmpStr = tmpStr & "'1','5','9','13',"
    End If
    
    If Mid(code, 2, 1) = "1" Then
        tmpStr = tmpStr & "'2','6','10','14',"
    End If
    
    If Mid(code, 3, 1) = "1" Then
        tmpStr = tmpStr & "'4','5','6','12','13','14',"
    End If
    
    If Right(code, 1) = "1" Then
        tmpStr = tmpStr & "'8','9','10','12','13','14',"
    End If

    commaPos = InStrRev(tmpStr, ",")
    tmpStr = Mid(tmpStr, 1, commaPos - 1) & ")"
    
    OrOp = tmpStr

End Function
