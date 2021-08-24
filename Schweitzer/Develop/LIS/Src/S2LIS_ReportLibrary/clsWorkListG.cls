VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWorkListG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private iPageWidth As Integer
Private iPageHeight As Integer
Private iCurY As Integer

Private iWidthTestCD As Integer
Private iTestCount As Integer
    
Private pWorkCode As String
Private pWorkName As String
Private pWorkDate As String
Private pFromSeq As Integer
Private pToSeq As Integer
    
Const iCm = 567
Const iLineHeight = 10

Private iposSEQ%, iposWorkNo%, iposPtName%, iposPtID%, iposSAge%, _
    iposIO%, iposRcv, iposSF%, iposXTestCD%, iposSpccd%
Private aWSTestCD() As String


Public Property Get WorkCode() As String
    WorkCode = pWorkCode
End Property

Public Property Let WorkCode(ByVal vNewValue As String)
    pWorkCode = vNewValue
End Property

Public Property Get WorkName() As String
    WorkName = pWorkName
End Property

Public Property Let WorkName(ByVal vNewValue As String)
    pWorkName = vNewValue
End Property

Public Property Get WorkDate() As String
    WorkDate = pWorkDate
End Property

Public Property Let WorkDate(ByVal vNewValue As String)
    pWorkDate = vNewValue
End Property

Public Property Get FromSeq() As Integer
    FromSeq = pFromSeq
End Property

Public Property Let FromSeq(ByVal vNewValue As Integer)
    pFromSeq = vNewValue
End Property

Public Property Get ToSeq() As Integer
    ToSeq = pToSeq
End Property

Public Property Let ToSeq(ByVal vNewValue As Integer)
    pToSeq = vNewValue
End Property


Public Sub Print_Worksheet()
    
    Dim tmp As String
 
    If pWorkCode = "" Then Exit Sub
    
    Call GetWScode
    
    If iTestCount > 10 Then
        Call Reporting900
    Else
        Call Reporting
    End If

End Sub

Private Sub Reporting900()

    Dim i%, j%, K%, iCurXTestCD%, iCurYTestCD%

    Dim sSqlLab301 As String
    Dim sSqlLab201 As String

    Dim ssqlHis001 As String
    
    Dim rslab301 As Recordset
    Dim rslab201 As Recordset
    Dim rshis001 As Recordset
    Dim rscom099 As Recordset

    Dim sWorkno As String
    Dim S_Age As String
    
    
    Dim sWorkarea As String
    Dim sAccDt As String
    Dim iAccSeq As Integer


    '******* get workseq, labnum(workarea, accdt, accseq) ************************
    sSqlLab301 = " select a.workseq, a.workarea, a.accdt, a.accseq " & _
                    " from " & T_LAB301 & " a " & _
                    " where " & _
                                      DBW("a.workdt=", pWorkDate) & _
                            " and " & DBW("a.workcd=", pWorkCode) & _
                            " and " & DBW("a.workseq>=", pFromSeq) & _
                            " and " & DBW("a.workseq<=", pToSeq) & _
                    " order by workseq "
    Set rslab301 = New Recordset
    rslab301.Open sSqlLab301, DBConn
    
    If rslab301.EOF = True Then
        MsgBox " Worksheet 내역이 존재하지 않습니다.. "
        GoTo Nodata
    End If
    
    Call InitReport
    Call PrtHeader900
    Call prtPageNum
    Call prtTerm
    Call Print_WaterMark

    rslab301.MoveFirst
    
    For i = 1 To rslab301.RecordCount ' 301의 각 labnum 마다
        If iCurY > iPageHeight - 2 * iCm Then ' newPage일 경우
            Printer.NewPage
            Call PrtHeader900
            Call prtPageNum
            Call prtTerm
            Call Print_WaterMark
        End If
      
        Call ChangeLine(iCm / 2)
        ' workseq, workno출력
        
        sWorkarea = "" & rslab301.Fields("workarea").Value
        sAccDt = "" & rslab301.Fields("accdt").Value
        iAccSeq = "" & rslab301.Fields("accseq").Value
        
        
        '****** ptid, sex/age, deptcd, storecd, rcvdt, rcvtm, spccd 출력*******
        sSqlLab201 = " select a.ptid, a.sex, a.ageday, a.deptcd, a.wardid, a.hosilid, a.storecd, a.rcvdt, a.statfg, " & _
                            " a.rcvtm, a.spccd, b.field3" & _
                     " from " & T_LAB201 & " a, " & T_LAB032 & " b " & _
                     " where " & _
                                       DBW("a.workarea=", sWorkarea) & _
                             " and " & DBW("a.accdt=", sAccDt) & _
                             " and " & DBW("a.accseq=", iAccSeq) & _
                             " and " & DBW("b.cdindex =", LC3_Specimen) & _
                             " and " & DBJ("b.cdval1 =* a.spccd")
        Set rslab201 = Nothing
        Set rslab201 = New Recordset
        rslab201.Open sSqlLab201, DBConn
        
       '=================================================================================================
        '감염관리
        Dim sICSString As String
        
        sICSString = ICSPatientString(rslab201.Fields("ptid").Value & "", enICSNum.LIS_ALL)
        

        If sICSString <> "" Then
            Printer.FontBold = True
            Call WriteStr(iCurY, iposWorkNo, "Infection : " & sICSString, iCurY, 0)
            
            Printer.FontBold = False
            If Trim("" & rslab201.Fields("statfg").Value) = "1" Then
                Call WriteStr(iCurY, iposSEQ, "" & rslab301.Fields("workseq").Value & "**", iCurY, 0)
            Else
                Call WriteStr(iCurY, iposSEQ, "" & rslab301.Fields("workseq").Value, iCurY, 0)
            End If
            Call ChangeLine(iCm / 2)
        Else
            '응급여부 (**)로 표시..
            If Trim("" & rslab201.Fields("statfg").Value) = "1" Then
                Call WriteStr(iCurY, iposSEQ, "" & rslab301.Fields("workseq").Value & "**", iCurY, 0)
            Else
                Call WriteStr(iCurY, iposSEQ, "" & rslab301.Fields("workseq").Value, iCurY, 0)
            End If
        End If
        '=================================================================================================
        
        'Call WriteStr(iCurY, iposSEQ, rslab301.Fields("workseq").Value, iCurY, 0)
        sWorkno = DelFirst2Chr("" & rslab301.Fields("accdt").Value) & "-" & _
                 "" & rslab301.Fields("accseq").Value
        Call WriteStr(iCurY, iposWorkNo, sWorkno, iCurY, 0)
        
        
        Call WriteStr(iCurY, iposPtID, "" & rslab201.Fields("ptid").Value, iCurY, 0)
        
         '****** ptnm출력 ******************************************************
        ssqlHis001 = " select " & F_PTNM & "  as ptnm, " & F_SEX & " as sex from " & T_HIS001 & _
                     " where " & DBW(F_PTID & "=", "" & rslab201.Fields("ptid").Value)
                        
        Set rshis001 = Nothing
        Set rshis001 = New Recordset
        rshis001.Open ssqlHis001, DBConn
        
        S_Age = rshis001.Fields("sex").Value & ""
        
        If IsNumeric(S_Age) Then
            S_Age = Choose((Val(S_Age) Mod 2) + 1, "F", "M")
        End If
        
        S_Age = S_Age & "/" & ("" & rslab201.Fields("ageday").Value \ 365 + 1)
        
        
        Call WriteStr(iCurY, iposSAge, S_Age, iCurY, 0)
        
'        S_Age = "" & rslab201.Fields("sex").Value & "/" & _
'                ("" & rslab201.Fields("ageday").Value \ 365 + 1)
'        Call WriteStr(iCurY, iposSAge, S_Age, iCurY, 0)
        '*****---------------------------------------------------------------------
        
        If ("" & rslab201.Fields("wardid").Value) <> "" Then
            Call WriteStr(iCurY, iposIO, "" & rslab201.Fields("wardid").Value & "-" & rslab201.Fields("hosilid").Value, iCurY, 0)
        Else
            Call WriteStr(iCurY, iposIO, "" & rslab201.Fields("deptcd").Value, iCurY, 0)
        End If
        'Call WriteStr(iCurY, iposSF, "" & rslab201.Fields("storecd").Value, iCurY, 0)
        Call WriteStr(iCurY, iposRcv, DelFirst2Chr("" & rslab201.Fields("rcvdt").Value) & " " & CvtTmFormat("" & rslab201.Fields("rcvtm").Value), iCurY, 0)
        'Call WriteStr(iCurY, iposRcv + iCm + 1.5, CvtTmFormat(rslab201.Fields("rcvtm").Value), iCurY, 0)
        'Call WriteStr(iCurY, iposSpccd, "" & rslab201.Fields("spccd").Value, iCurY, 0)
        Call WriteStr(iCurY, iposSpccd, "" & rslab201.Fields("field3").Value, iCurY, 0)
        
        '-- 업무나열서 일련번호 가져오기=====================================
        Dim strRackMaxNo As String
        Dim strRackNo    As String
        Dim strRemark    As String
        Dim strSQL       As String
        Dim iSeq         As Long
        
        'strRackMaxNo = RackMaxSeq("" & rslab201.Fields("rcvdt").Value)
        '-- Remark ( WorkArea & AccDt & AccSeq)
        strRemark = sWorkarea & sAccDt & iAccSeq
        strSQL = " select seq from " & T_COM099 & _
                 "  where " & DBW("remark = ", strRemark)
' ----------------------------------------------------
' 2009.02.19 양성현 Query 문 조건대로면 Index Key가 없어서 너무 오래걸림
        strSQL = "select seq from " & T_COM099 & _
                 " where  noindex = 'L004' " & _
                 "  and DivCD1 = '" & Mid(strRemark, 1, 2) & _
                 "' and substr(divcd2,1,4) = '" & Mid(strRemark, 3, 4) & _
                 "' and substr(divcd2,1,4) =  substr(remark,3,4) and " & DBW("remark = ", strRemark)
' ----------------------------------------------------

        Set rscom099 = Nothing
        Set rscom099 = New Recordset
        rscom099.Open strSQL, DBConn
        
        If rscom099.EOF = False Then
            iSeq = rscom099.Fields("seq").Value & ""
        End If
        
        Set rscom099 = Nothing
        '=====================================================================
        
'        If strRackMaxNo <> "" Then
'            iSeq = CLng(strRackMaxNo) + 1
'        Else
'            iSeq = 1
'        End If
            
        '20131130 PSK 번호삭제
        'strRackNo = "(" & iSeq & ")"
        strRackNo = ""
        '20131130 PSK 번호삭제
        '****** ptnm출력 ******************************************************
        
'        ssqlHis001 = " select " & F_PTNM & "  as ptnm from " & T_HIS001 & _
'                        " where " & DBW(F_PTID & "=", "" & rslab201.Fields("ptid").Value)
'
'        Set rshis001 = OpenRecordSet(ssqlHis001)
        
        Call WriteStr(iCurY, iposPtName, "" & rshis001.Fields("ptnm").Value & strRackNo, iCurY, 0)
        
        
        '****** 해당 Test Code 괄호(    )  출력 *******************************
        Call DspBrackketTestCD(sWorkarea, sAccDt, iAccSeq)
        
        Set rshis001 = Nothing
        Set rslab201 = Nothing
        
        rslab301.MoveNext
    Next i
    Call Print_WaterMark
    Printer.EndDoc

Nodata:
    Set rslab301 = Nothing
    
End Sub


Private Sub GetWScode()

    Dim sSqlGetWSCode  As String
    Dim rsWorksheet As Recordset
    
    sSqlGetWSCode = " select  a.cdval1 as WorkCd, a.field1 as WorkNm, count(b.testcd) as TestCnt " & _
            " from " & _
                     T_LAB032 & " a, " & T_LAB008 & " b " & _
            " where " & _
                              DBW("a.cdindex=", LC3_WorkSheetName) & _
                    " and " & DBW("a.cdval1=", pWorkCode) & _
                    " and b.workcd = a.cdval1 " & _
            " group by a.cdval1, a.field1 "
    
    Set rsWorksheet = New Recordset
    rsWorksheet.Open sSqlGetWSCode, DBConn
    
    pWorkName = "" & rsWorksheet.Fields("WorkNm").Value
    iTestCount = Val("" & rsWorksheet.Fields("TestCnt").Value)
    
    Set rsWorksheet = Nothing
    
End Sub


Private Sub Reporting()

    Dim i%, j%, K%, iCurXTestCD%, iCurYTestCD%

    Dim sSqlLab301 As String
    Dim sSqlLab201 As String

    Dim ssqlHis001 As String
    
    Dim sWorkDate As String
    
    Dim rslab301 As Recordset
    Dim rslab201 As Recordset
    Dim rshis001 As Recordset
    Dim rscom099 As Recordset

    Dim sWorkno As String
    Dim S_Age As String
    
    
    Dim sWorkarea As String
    Dim sAccDt As String
    Dim iAccSeq As Integer
    
    Dim iHeaderPosY  As Integer
    Dim iHeaderPosY1 As Integer
    Dim iHeaderPosY2 As Integer


    
    '******* get workseq, labnum(workarea, accdt, accseq) ************************
    sSqlLab301 = " select workseq,workarea,accdt, accseq " & _
                 " from " & T_LAB301 & _
                 " where " & _
                                   DBW("workdt=", pWorkDate) & _
                         " and " & DBW("workcd=", pWorkCode) & _
                         " and " & DBW("workseq>=", pFromSeq) & _
                         " and " & DBW("workseq<=", pToSeq) & _
                 " order by workseq "
    Set rslab301 = New Recordset
    rslab301.Open sSqlLab301, DBConn
    
    If rslab301.EOF = True Then
        MsgBox " Worksheet 내역이 존재하지 않습니다.. "
        GoTo Nodata
    End If
    
    Call InitReport
    Call PrtHeader
    Call prtPageNum
    Call prtTerm
    Call Print_WaterMark

    rslab301.MoveFirst
    
    For i = 1 To rslab301.RecordCount ' 301의 각 labnum 마다
        
        If iCurY > iPageHeight - 4 * iCm Then ' newPage일 경우
            Printer.NewPage
            Call PrtHeader
            Call prtPageNum
            Call prtTerm
            Call Print_WaterMark
        End If
          
        Call ChangeLine(iCm / 2)
        
        
        ' workseq, workno출력
        Printer.FontSize = 9
        sWorkno = DelFirst2Chr("" & rslab301.Fields("accdt").Value) & "-" & _
                 "" & rslab301.Fields("accseq").Value
        sWorkarea = "" & rslab301.Fields("workarea").Value
        sAccDt = "" & rslab301.Fields("accdt").Value
        iAccSeq = "" & rslab301.Fields("accseq").Value
        
        
        '****** ptid, sex/age, deptcd, storecd, rcvdt, rcvtm, spccd 출력*******
        sSqlLab201 = " select a.ptid, a.sex, a.ageday, a.deptcd, a.wardid, a.hosilid, a.storecd, a.rcvdt, a.statfg, " & _
                            " a.rcvtm, a.spccd, b.field3" & _
                     " from " & T_LAB201 & " a, " & T_LAB032 & " b " & _
                     " where " & _
                                       DBW("a.workarea=", sWorkarea) & _
                             " and " & DBW("a.accdt=", sAccDt) & _
                             " and " & DBW("a.accseq=", iAccSeq) & _
                             " and " & DBW("b.cdindex=", LC3_Specimen) & _
                             " and  a.spccd = b.cdval1"
        
        Set rslab201 = Nothing
        Set rslab201 = New Recordset
        rslab201.Open sSqlLab201, DBConn
        
        '=================================================================================================
        '감염관리
        Dim sICSString As String
        
        sICSString = ICSPatientString(rslab201.Fields("ptid").Value & "", enICSNum.LIS_ALL)
        
        If sICSString <> "" Then
            iHeaderPosY = iCurY
            iHeaderPosY1 = iCurY + iCm / 2
            iHeaderPosY2 = iCurY + iCm
        Else
            iHeaderPosY1 = iCurY
            iHeaderPosY2 = iCurY + iCm / 2
        End If
        
        If sICSString <> "" Then
            Printer.FontBold = True
            Call WriteStr(iHeaderPosY, iposPtName, "Infection : " & sICSString, iHeaderPosY, 0)
            Printer.FontBold = False
            Call WriteStr(iHeaderPosY, iposSEQ, "" & rslab301.Fields("workseq").Value, iHeaderPosY, 0)
        Else
            Call WriteStr(iHeaderPosY1, iposSEQ, "" & rslab301.Fields("workseq").Value, iHeaderPosY1, 0)
        End If
        '=================================================================================================
        
        
        Call WriteStr(iHeaderPosY1, iposWorkNo, sWorkno, iHeaderPosY1, 0)
        
        
        If Not rslab201.EOF Then
        
        If Trim("" & rslab201.Fields("statfg").Value) = "1" Then Call WriteStr(iHeaderPosY2, iposSEQ, "**", iHeaderPosY2, 0)

        Call WriteStr(iHeaderPosY2, iposPtID, "" & rslab201.Fields("ptid").Value, iHeaderPosY2, 0)
        '****** ptnm출력 ******************************************************
        ssqlHis001 = " select " & F_PTNM & " as ptnm, " & F_SEX & " as sex from " & T_HIS001 & " where " & DBW(F_PTID & "=", "" & rslab201.Fields("ptid").Value)
                        
        Set rshis001 = Nothing
        Set rshis001 = New Recordset
        rshis001.Open ssqlHis001, DBConn
        
        S_Age = rshis001.Fields("sex").Value & ""
        
        If IsNumeric(S_Age) Then S_Age = Choose((Val(S_Age) Mod 2) + 1, "F", "M")
        
        S_Age = S_Age & "/" & ("" & rslab201.Fields("ageday").Value \ 365 + 1)
        Call WriteStr(iHeaderPosY2, iposSAge, S_Age, iHeaderPosY2, 0)
        '****************************************************************************
        'Call WriteStr(iHeaderPosY2, iposIO, "" & rslab201.Fields("deptcd").Value, iHeaderPosY2, 0)
        If ("" & rslab201.Fields("wardid").Value) <> "" Then
            Call WriteStr(iHeaderPosY2, iposIO, "" & rslab201.Fields("wardid").Value & "-" & rslab201.Fields("hosilid").Value, iHeaderPosY2, 0)
        Else
            Call WriteStr(iHeaderPosY2, iposIO, "" & rslab201.Fields("deptcd").Value, iHeaderPosY2, 0)
        End If
        'Call WriteStr(iHeaderPosY1, iposSF, "" & rslab201.Fields("storecd").Value, iHeaderPosY1, 0)
        Call WriteStr(iHeaderPosY1, iposRcv, DelFirst2Chr("" & rslab201.Fields("rcvdt").Value), iHeaderPosY1, 0)
        Call WriteStr(iHeaderPosY1, iposRcv + iCm + iCm / 6, CvtTmFormat("" & rslab201.Fields("rcvtm").Value), iHeaderPosY1, 0)
        Call WriteStr(iHeaderPosY1, iposSpccd, "" & rslab201.Fields("field3").Value, iHeaderPosY1, 0)
        'Call WriteStr(iHeaderPosY1, iposSpccd, "" & rslab201.Fields("spccd").Value, iHeaderPosY1, 0)
        
        '-- 업무나열서 일련번호 가져오기=====================================
        Dim strRackMaxNo As String
        Dim strRackNo    As String
        Dim strRemark    As String
        Dim strSQL       As String
        Dim iSeq         As Long
        
        'strRackMaxNo = RackMaxSeq("" & rslab201.Fields("rcvdt").Value)
        '-- Remark ( WorkArea & AccDt & AccSeq)
        strRemark = sWorkarea & sAccDt & iAccSeq
        strSQL = " select seq from " & T_COM099 & "  where " & DBW("remark = ", strRemark)
' ----------------------------------------------------
' 2009.02.19 양성현 Query 문 조건대로면 Index Key가 없어서 너무 오래걸림
        strSQL = "select seq from " & T_COM099 & _
                 " where  noindex = 'L004' " & _
                 "  and DivCD1 = '" & Mid(strRemark, 1, 2) & _
                 "' and substr(divcd2,1,4) = '" & Mid(strRemark, 3, 4) & _
                 "' and substr(divcd2,1,4) =  substr(remark,3,4) and " & DBW("remark = ", strRemark)
' ----------------------------------------------------

        Set rscom099 = Nothing
        Set rscom099 = New Recordset
        rscom099.Open strSQL, DBConn
        
        If rscom099.EOF = False Then iSeq = rscom099.Fields("seq").Value & ""
        
        Set rscom099 = Nothing
        '=====================================================================
        
'        If strRackMaxNo <> "" Then
'            iSeq = CLng(strRackMaxNo) + 1
'        Else
'            iSeq = 1
'        End If
        
        '20131130 PSK 번호삭제
        'strRackNo = "(" & iSeq & ")"
        strRackNo = ""
        '20131130 PSK 번호삭제
        
        '****** ptnm출력 ******************************************************
        
'        ssqlHis001 = " select " & F_PTNM & " as ptnm from " & T_HIS001 & _
                        " where " & DBW(F_PTID & "=", "" & rslab201.Fields("ptid").Value)
        
        
'        Set rshis001 = OpenRecordSet(ssqlHis001)
        
        Call WriteStr(iHeaderPosY2, iposPtName, "" & rshis001.Fields("ptnm").Value & strRackNo, iHeaderPosY2, 0)
        
        '-- 업무나열서 일련번호 Update ========================================
'        If SqlRackNoState(sWorkArea, sAccDt, iAccSeq) = True Then
'            Call SqlRackNoUpdate(sWorkArea, sAccDt, iAccSeq, iSeq)
'        Else
'            Call SqlRackNoInsert(sWorkArea, sAccDt, iAccSeq, "" & rslab201.Fields("rcvdt").Value, _
'                                  "" & rslab201.Fields("rcvtm").Value, iSeq)
'        End If
        '======================================================================
        Set rshis001 = Nothing
        
'        rshis001.RsClose
'        Set rshis001 = Nothing
        
        '****** 해당 Test Code 괄호(    )  출력 *******************************
        If sICSString <> "" Then
            iCurY = iCurY + iCm / 2
        End If
        Call DrawBrackets(sWorkarea, sAccDt, iAccSeq)
        End If
        
        Set rslab201 = Nothing
        
        'iSeq = iSeq + 1
        rslab301.MoveNext
    Next i
    Printer.EndDoc
    
Nodata:
    Set rslab301 = Nothing
    
End Sub

Private Function SqlRackNoState(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As Boolean
    Dim ObjAccession As clsLISSqlAccession
    Dim Rs           As Recordset
    Dim strSQL       As String
    
    Set ObjAccession = New clsLISSqlAccession
    
    strSQL = ObjAccession.SqlRackNoState(pWorkArea, pAccDt, pAccSeq)
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    If Rs.EOF = False Then
        SqlRackNoState = True
    Else
        SqlRackNoState = False
    End If
    
    Set ObjAccession = Nothing
End Function

Private Sub SqlRackNoUpdate(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                            ByVal pRackNo As String)
    Dim ObjAccession As clsLISSqlAccession
    Dim strSQL       As String
    
    Set ObjAccession = New clsLISSqlAccession
    
    strSQL = ObjAccession.SqlRackNoUpdate(pWorkArea, pAccDt, pAccSeq, pRackNo)
    
    DBConn.Execute (strSQL)
    
    Set ObjAccession = Nothing
End Sub

Private Sub SqlRackNoInsert(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                            ByVal pRcvDt As String, ByVal pRcvTm As String, ByVal pRackNo As String)
    Dim ObjAccession As clsLISSqlAccession
    Dim strSQL       As String
    
    Set ObjAccession = New clsLISSqlAccession
    
    strSQL = ObjAccession.SqlRackNoInsert(pWorkArea, pAccDt, pAccSeq, pRcvDt, pRcvTm, pRackNo)
    
    Call DBConn.Execute(strSQL)
    
    Set ObjAccession = Nothing
End Sub

Private Function RackMaxSeq(ByVal pRcvDt As String) As String
    Dim ObjAccession As clsLISSqlAccession
    Dim Rs           As Recordset
    Dim strSQL       As String
        
    Set ObjAccession = New clsLISSqlAccession
    
    strSQL = ObjAccession.SqlRackNoMax(pRcvDt, "")
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    If Rs.EOF = False Then
        RackMaxSeq = Rs.Fields("cnt").Value & ""
    End If
    
    Set Rs = Nothing
    Set ObjAccession = Nothing
End Function

Private Function CvtTmFormat(sStr As String) As String
    Dim Time As String
    Dim Hour As String
    Dim Min As String
    
    Time = DelLast2Chr(sStr)
    Hour = DelLast2Chr(Time)
    Min = DelFirst2Chr(Time)
    
    CvtTmFormat = Hour & ":" & Min
    
End Function

Private Function DelFirst2Chr(sStr As String) As String
    DelFirst2Chr = Trim(Mid(sStr, 3, Len(sStr) - 2))
End Function

Private Function DelLast2Chr(sStr As String) As String
    DelLast2Chr = Trim(Mid(sStr, 1, Len(sStr) - 2))
End Function

Private Sub InitReport()
    'Printer.PaperSize = vbPRPSA4
    iPageWidth = Printer.ScaleWidth
    iPageHeight = Printer.ScaleHeight - iCm
End Sub

Private Sub PrtHeader()
   
    Dim Title       As String
    Dim sWSNanme    As String
    Dim iHeaderPosY As Integer
    
    '/* 보고서 제목
    Title = "업 무 나 열 서"
    
    
    '        "업무나열서"
    ' ----------------------
    ' ----------------------
    
    Call prtTitle(Title, iCm / 4)
    Call DrawLine(8 * iCm, iCurY, 13 * iCm, iCurY, "dot", 1, iCm / 10)
    Call DrawLine(8 * iCm, iCurY, 13 * iCm, iCurY, "dot", 1, iCm / 10)
    'TITLE = "WorkSheet : " & lblWSName.Caption
    
    sWSNanme = "WorkSheet  :  " & pWorkName
    
    Call WriteStr(iCurY, iPageWidth / 2 - Printer.TextWidth(sWSNanme) / 2, _
                  sWSNanme, iCurY, iCm * 1.5)
'    Call prtTitle(TITLE, iCm / 4)
    ' -----------------------------------------------------------------------------
    Call DrawLine(iCm / 6, iCurY, iPageWidth - iCm / 8, iCurY, "solid", 2, iCm / 4)
    
    iposSEQ = iCm / 6                       '284
    iposWorkNo = iposSEQ + iCm / 1.4        '757
    iposRcv = iposWorkNo + 2.2 * iCm    'iposXTestCD + 8.5 * iCm        '10038.5
    iposSpccd = iposRcv + 2.2 * iCm         '11286
    iposSF = iposSpccd + 0.8 * iCm             '4765
    
    iposPtName = iposWorkNo       '1891
    iposPtID = iposRcv  'iposPtName + iCm + iCm / 1.2 '2894
    iposSAge = iposPtID + iCm + iCm / 1.5     '3744
    'iposIO = iposSAge + iCm                 '4311
    iposIO = iposSAge + iCm
    
    iposXTestCD = iposSF + iCm / 1.5     'iposWorkNo
    
    iHeaderPosY = iCurY
    
    Call WriteStr(iHeaderPosY, iposSEQ, "No", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposWorkNo, "접수번호", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposRcv, "검체도착시간", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposSpccd - iCm / 6, "검체", iHeaderPosY, 0)
    'Call WriteStr(iHeaderPosY, iposSF - iCm / 6, "S/F", iHeaderPosY, 0)
    
    iHeaderPosY = iHeaderPosY + iCm / 2
    
    'Call WriteStr(iHeaderPosY, iposPtName, "환자성명", iHeaderPosY, 0)
    'Call WriteStr(iHeaderPosY, iposPtID, "환자ID", iHeaderPosY, 0)
    'Call WriteStr(iHeaderPosY, iposSAge, "S/Age", iHeaderPosY, 0)
    'Call WriteStr(iHeaderPosY, iposIO, "I/O", iHeaderPosY, 0)
    
    'Call WriteStr(iCurY, iposXTestCD , "검사종목", iCurY, 0)
    Call DspWSTestCD(iHeaderPosY - iCm / 2, iposXTestCD)
    Call DrawLine(iCm / 6, iCurY + iCm / 2, iPageWidth - iCm / 8, iCurY + iCm / 2, "solid", 2, 0)
 
'    Call DspWSTestCD(iHeaderPosY - iCm / 2, iposXTestCD)
'    Call DrawLine(iCm / 6, iCurY, iPageWidth - iCm / 8, iCurY, "solid", 2, 0)
     
    '-------------------------------------------------------------------------
'    iposSEQ = iCm / 2                       '284
'    iposWorkNo = iposSEQ + iCm / 1.2        '757
'    iposPtName = iposWorkNo + 2 * iCm       '1891
'    iposPtID = iposPtName + iCm + iCm / 1.3 '2894
'    iposSAge = iposPtID + iCm + iCm / 2     '3744
'    iposIO = iposSAge + iCm                 '4311
'    iposSF = iposIO + 0.8 * iCm             '4765
'    iposRcv = iposSF + 0.8 * iCm  'iposXTestCD + 8.5 * iCm        '10038.5
'    iposSpccd = iposRcv + 2.2 * iCm         '11286
'
'    iposXTestCD = iposWorkNo
'
'    iHeaderPosY = iCurY
'
'    Call WriteStr(iHeaderPosY, iposSEQ, "SEQ", iHeaderPosY, 0)
'    Call WriteStr(iHeaderPosY, iposWorkNo, "Work No", iHeaderPosY, 0)
'    Call WriteStr(iHeaderPosY, iposPtName, "환자성명", iHeaderPosY, 0)
'    Call WriteStr(iHeaderPosY, iposPtID, "환자ID", iHeaderPosY, 0)
'    Call WriteStr(iHeaderPosY, iposSAge, "S/Age", iHeaderPosY, 0)
'    Call WriteStr(iHeaderPosY, iposIO, "I/O", iHeaderPosY, 0)
'    Call WriteStr(iHeaderPosY, iposSF, "S/F", iHeaderPosY, 0)
'    'Call WriteStr(iCurY, iposXTestCD , "검사종목", iCurY, 0)
'    Call WriteStr(iHeaderPosY, iposRcv, "검체도착시간", iHeaderPosY, 0)
'    Call WriteStr(iHeaderPosY, iposSpccd, "검체", iHeaderPosY, 0)
'    Call DspWSTestCD(iHeaderPosY + iCm / 2, iposXTestCD)
'    Call DrawLine(iCm / 2, iCurY, iPageWidth - iCm / 2, iCurY, "solid", 2, 0)
    
    
End Sub


Private Sub PrtHeader900()
   
    Dim Title As String
    Dim sWSNanme As String
    Dim iHeaderPosY As Integer
    
    '/* 보고서 제목
    Title = "업 무 나 열 서"
    
    '        "업무나열서"
    ' ----------------------
    ' ----------------------
    
    Call prtTitle(Title, iCm / 4)
    Call DrawLine(8 * iCm, iCurY, 13 * iCm, iCurY, "dot", 1, iCm / 10)
    Call DrawLine(8 * iCm, iCurY, 13 * iCm, iCurY, "dot", 1, iCm / 10)
    'TITLE = "WorkSheet : " & lblWSName.Caption
    
    sWSNanme = "WorkSheet  :  " & pWorkName
    
    Call WriteStr(iCurY, iPageWidth / 2 - Printer.TextWidth(sWSNanme) / 2, _
                  sWSNanme, iCurY, iCm * 1.5)
'    Call prtTitle(TITLE, iCm / 4)
    ' -----------------------------------------------------------------------------
    Call DrawLine(iCm / 2, iCurY, iPageWidth - iCm / 2, iCurY, "solid", 2, iCm / 4)
    
    iposSEQ = iCm / 2                       '284
    iposWorkNo = iposSEQ + iCm / 1.2        '757
    iposPtName = iposWorkNo + 2.2 * iCm       '1891
    iposPtID = iposPtName + iCm + iCm / 1.3 '2894
    iposSAge = iposPtID + iCm + iCm / 1.5     '3744
    iposIO = iposSAge + iCm                 '4311
    iposSF = iposIO + 0.8 * iCm             '4765
    iposRcv = iposSF + 0.8 * iCm  'iposXTestCD + 8.5 * iCm        '10038.5
    iposSpccd = iposRcv + 2.4 * iCm         '11286
    
    iposXTestCD = iposWorkNo
    
    iHeaderPosY = iCurY
    
    Call WriteStr(iHeaderPosY, iposSEQ, "No", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposWorkNo, "접수번호", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposPtName, "환자성명", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposPtID, "환자ID", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposSAge, "S/Age", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposIO, "I/O", iHeaderPosY, 0)
    'Call WriteStr(iHeaderPosY, iposSF, "S/F", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposRcv, "검체도착시간", iHeaderPosY, 0)
    Call WriteStr(iHeaderPosY, iposSpccd, "검체", iHeaderPosY, 0)
    'Call WriteStr(iHeaderPosY + iCm / 2, iposXTestCD, "검사종목", iCurY, iCm / 2)
    'Call DspWSTestCD(iHeaderPosY + iCm / 2, iposXTestCD)
    Call DrawLine(iCm / 2, iHeaderPosY + iCm / 2, iPageWidth - iCm / 2, iHeaderPosY + iCm / 2, "solid", 2, 0)
    
End Sub

Private Sub DspWSTestCD(ByVal iPosY As Integer, ByVal iPosX As Integer)
    Dim sSQL As String
    Dim rsWSTestCD As Recordset
    Dim i As Integer, iOldy As Integer, iOldX As Integer
    
    
    iOldy = iPosY
    iOldX = iPosX
    iWidthTestCD = 1.5 * iCm
    
    sSQL = " select a.testcd, b.abbrnm5, b.rptseq,a.workinfo from " & T_LAB008 & " a, " & T_LAB001 & " b " & _
            " where " & DBW("a.workcd=", pWorkCode) & _
            " and   b.testcd = a.testcd " & _
            " and   b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd=a.testcd) " & _
            " group by a.workinfo,a.testcd, b.abbrnm5, b.rptseq " & _
            " order by a.workinfo, b.rptseq, a.testcd"
            
    Set rsWSTestCD = New Recordset
    rsWSTestCD.Open sSQL, DBConn
    
    rsWSTestCD.MoveFirst
    ReDim aWSTestCD(1 To rsWSTestCD.RecordCount)
    
    For i = 1 To rsWSTestCD.RecordCount
        aWSTestCD(i) = "" & rsWSTestCD.Fields("testcd").Value
        Call WriteStr(iPosY, iPosX, "" & rsWSTestCD.Fields("testcd").Value, iCurY, 0)
        Call WriteStr(iPosY + iCm / 2, iPosX, "" & rsWSTestCD.Fields("abbrnm5").Value, iCurY, 0)
        
'        aWSTestCD(i) = "" & rsWSTestCD.Fields("testcd").Value
'        Call WriteStr(iPosY, iPosX, "" & rsWSTestCD.Fields("testcd").Value, iCurY, 0)
        rsWSTestCD.MoveNext
        iPosX = iPosX + iWidthTestCD
        'If i Mod 10 = 0 Then
        '    iPosY = iPosY + iCm / 2
        '    iPosX = iOldX
        'End If
        If i Mod 20 = 0 Then
            Call DrawLine(iOldX, iPosY + iCm / 2.5, iposRcv - iCm / 3, _
                            iPosY + iCm / 2.5, "dot", 1, 1)
            iPosY = iPosY + iCm / 2
            iPosX = iOldX
        End If
    Next i
    
    'If rsWSTestCD.RecordCount Mod 16 > 0 Then
    '  iCurY = iPosY + iCm
    'Else
      iCurY = iPosY + iCm / 2
   'End If
    
    Set rsWSTestCD = Nothing
    
End Sub

Private Sub prtPageNum()
    
    Dim oldX As Integer, oldY As Integer
    Dim sDate As String, sTime As String
    
    sDate = Format(Now, "YYYY/MM/DD")
    sTime = Format(Now, "HH:MM:SS")
    oldX = Printer.CurrentX
    oldY = Printer.CurrentY
    
    Printer.CurrentX = iPageWidth - 4 * iCm
    Printer.CurrentY = 0
    Printer.Print "P A G E  : " & Printer.Page
            
    Printer.CurrentX = iPageWidth - 4 * iCm
    Printer.CurrentY = Printer.TextHeight("P A G E") + iCm / 6
    Printer.Print "RUN-DATE : " & sDate
        
    Printer.CurrentX = iPageWidth - 4 * iCm
    Printer.CurrentY = Printer.TextHeight("P A G E") + iCm / 6 + _
                           Printer.TextHeight("RUN-DATE") + iCm / 6
    Printer.Print "RUN-TIME : " & sTime
        
    Printer.CurrentX = oldX
    Printer.CurrentY = oldY
    
End Sub

Private Sub prtTerm()
    Dim oldX As Integer, oldY As Integer
    Dim sSQLGetWAName As String
    Dim rsWAName As Recordset
    Dim sWorkDate As String
    
    
    sSQLGetWAName = " select distinct WS.workarea, WA.field1 as WAName" & _
                    " from " & T_LAB032 & " WA, " & T_LAB008 & " WS " & _
                    " where " & _
                          "         WA.cdval1 = WS.workarea  " & _
                          " and " & DBW("WA.cdindex=", LC3_WorkArea) & _
                          " and " & DBW("WS.workcd=", pWorkCode)

    Set rsWAName = New Recordset
    rsWAName.Open sSQLGetWAName, DBConn
    
    If rsWAName.EOF = True Then GoTo Nodata
        
    sWorkDate = Format(pWorkDate, CS_DateMask)
    oldX = Printer.CurrentX
    oldY = Printer.CurrentY
        
    Printer.CurrentX = iposSEQ  'iCm
    Printer.CurrentY = 1.3 * iCm
    Printer.Print "◈ 작 업 일  : " & sWorkDate
    
    Printer.CurrentX = iposSEQ  'iCm
    Printer.CurrentY = 1.3 * iCm + Printer.TextHeight("작업일 : ")
    Printer.Print "◈ WorkArea  : " & "" & rsWAName.Fields("workarea").Value & "   " & "" & rsWAName.Fields("WAName").Value
                                    
    Printer.CurrentX = iposSEQ  'iCm
    Printer.CurrentY = 1.3 * iCm + Printer.TextHeight("작업일 : ") + _
                       Printer.TextHeight("workArea : ")
    Printer.Print "◈ Work No   : " & pFromSeq & " - " & pToSeq
    
    Printer.CurrentX = oldX
    Printer.CurrentY = oldY

Nodata:
    Set rsWAName = Nothing
    
End Sub

Private Sub DrawLine(ByVal iStartX As Integer, ByVal iStartY As Integer, _
                    ByVal iEndX As Integer, ByVal iEndy As Integer, _
                    sLineStyle As String, iLinewidth As Integer, iSpace As Integer)

    Select Case sLineStyle
        Case "solid"
            Printer.DrawStyle = 0
        Case "dash"
            Printer.DrawStyle = 1
        Case "dot"
            Printer.DrawStyle = 2
        Case "dashdot"
            Printer.DrawStyle = 3
        Case "dashdotdot"
            Printer.DrawStyle = 4
    End Select
         
    Printer.DrawWidth = iLinewidth
    Printer.Line (iStartX, iStartY)-(iEndX, iEndy)
    iCurY = Printer.CurrentY + iSpace
End Sub

Private Sub WriteStr(ByVal Y As Integer, ByVal X As Integer, ByVal str As String, _
                    iNextY As Integer, iSpace As Integer)
    Printer.CurrentY = Y
    Printer.CurrentX = X
    iNextY = Printer.CurrentY + iSpace
    Printer.Print str
   
End Sub

Private Sub ChangeLine(iLineSpace As Integer)
    iCurY = iCurY + iLineSpace
    Printer.CurrentY = iCurY
    
End Sub

Private Sub prtTitle(Title As String, iSpace As Integer)

    Dim oldFontSize As Integer
    
    oldFontSize = Printer.FontSize
    Printer.FontSize = 14
    Printer.FontBold = True
    '/* Tile이 중앙으로 오도록 string길이에 따라 위치를 계산한다.
    
    Printer.CurrentY = 0
    Printer.CurrentX = iPageWidth / 2 - Printer.TextWidth(Title) / 2
    iCurY = Printer.CurrentY + Printer.TextHeight(Title) + iSpace
    
    Printer.Print Title
    Printer.FontSize = 9    'oldFontSize
    Printer.FontBold = False
    
End Sub

Private Sub DspBrackketTestCD(sWorkarea As String, sAccDt As String, iAccSeq As Integer)
    
    Dim sSqlLab302        As String
    Dim rslab302          As Recordset
    Dim aBracketsTestCD() As Integer
    Dim i%, j%
    Dim tmpiposXTestCD%
    Dim iXcnt%, iYcnt%, iOldy%
    Dim iWidthTestCD900%
    
    Call ChangeLine(iCm / 2)

    sSqlLab302 = " select a.testcd, c.rptseq, c.abbrnm5,b.workinfo " & _
                    " from " & T_LAB302 & " a, " & T_LAB008 & " b, " & T_LAB001 & " c " & _
                    " where " & _
                                       DBW("a.workarea=", sWorkarea) & _
                            " and  " & DBW("a.accdt=", sAccDt) & _
                            " and  " & DBW("a.accseq=", iAccSeq) & _
                            " and  " & DBW("a.workcd=", pWorkCode) & _
                            " and (" & DBW("a.vfydt=", "") & "  or  a.vfydt is null) " & _
                            " and  " & DBW("b.workcd=", pWorkCode) & _
                            " and b.testcd = a.testcd" & _
                            " and c.testcd = a.testcd" & _
                            " and (c.panelfg = ''  or  c.panelfg is null) " & _
                            " and c.applydt = (select max(applydt) from " & T_LAB001 & _
                                             " where testcd = c.testcd) " & _
                    " group by b.workinfo,a.testcd, c.rptseq, c.abbrnm5 "

    
    sSqlLab302 = sSqlLab302 & " UNION ALL " & _
                    " select a.testcd, c.rptseq, c.abbrnm5,b.workinfo " & _
                    " from " & T_LAB351 & " a, " & T_LAB008 & " b, " & T_LAB001 & " c " & _
                    " where " & _
                                      DBW("a.workarea=", sWorkarea) & _
                            " and " & DBW("a.accdt=", sAccDt) & _
                            " and " & DBW("a.accseq=", iAccSeq) & _
                            " and (" & DBW("a.vfydt=", "") & "  or  a.vfydt is null) " & _
                            " and " & DBW("b.workcd=", pWorkCode) & _
                            " and b.testcd = a.testcd" & _
                            " and c.testcd = a.testcd" & _
                            " and (c.panelfg = ''  or  c.panelfg is null) " & _
                            " and c.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = c.testcd) " & _
                    " group by b.workinfo,a.testcd, c.rptseq, c.abbrnm5  " & _
                    " order by workinfo,rptseq, testcd"
        
    Set rslab302 = New Recordset
    rslab302.Open sSqlLab302, DBConn
    
    If rslab302.EOF Then Exit Sub
    rslab302.MoveFirst
    
    tmpiposXTestCD = iposXTestCD
    iOldy = iCurY
    
    iWidthTestCD900 = iCm * 2.3
    
    For i = 1 To rslab302.RecordCount
        
        iYcnt = i \ 9   ' 몫
        iXcnt = i Mod 8 '나머지

        iCurY = iOldy + iYcnt * (iCm / 2)

        If iXcnt = 0 Then
            tmpiposXTestCD = iposXTestCD + iWidthTestCD900 * 7
        Else
            tmpiposXTestCD = iposXTestCD + iWidthTestCD900 * (iXcnt - 1)
        End If

        'Call WriteStr(iCurY, tmpiposXTestCD, Trim("" & rslab302.Fields("testcd").Value) & _
                     " (       )", iCurY, 0)
        '최근결과를 보여준다
        Dim Rs   As Recordset
        Dim sSQL As String
        Dim LastRst As String
        
        LastRst = ""
        sSQL = "select lastrst from " & T_LAB302 & " a" & _
               " where " & _
                          DBW("a.workarea=", sWorkarea) & _
               " and  " & DBW("a.accdt=", sAccDt) & _
               " and  " & DBW("a.accseq=", iAccSeq) & _
               " and  " & DBW("a.testcd=", "" & rslab302.Fields("testcd").Value)
        Set Rs = Nothing
        Set Rs = New Recordset
        Rs.Open sSQL, DBConn
        
        If Not Rs.EOF Then
            LastRst = Rs.Fields("lastrst").Value & ""
        End If
        
        If LastRst <> "" Then
            Call WriteStr(iCurY, tmpiposXTestCD, Trim("" & rslab302.Fields("abbrnm5").Value) & _
                         " (" & LastRst & ")", iCurY, 0)
        Else
            Call WriteStr(iCurY, tmpiposXTestCD, Trim("" & rslab302.Fields("abbrnm5").Value) & _
                         " (       )", iCurY, 0)
        End If
        Set Rs = Nothing
        
            
    
        rslab302.MoveNext
    Next i
        
    Set rslab302 = Nothing
    
    Call ChangeLine(iCm / 2)
    Call DrawLine(iCm / 2, iCurY, iPageWidth - iCm / 2, iCurY, "dot", 2, 0)

End Sub

Private Sub DrawBrackets(sWorkarea As String, sAccDt As String, iAccSeq As Integer)
    
    Dim sSqlLab302 As String
    Dim rslab302 As Recordset
    Dim aBracketsTestCD() As Integer
    Dim i%, j%
    Dim tmpiposXTestCD%
    Dim iXcnt%, iYcnt%, iOldy%
    Dim strRemark As String
    Dim sLastCd As String
    
    'Call ChangeLine(iCm / 2)
    sSqlLab302 = " select d.rptseq, a.testcd, b.mesg  " & _
                " from " & T_LAB302 & " a, " & T_LAB102 & " b, " & T_LAB008 & " c, " & T_LAB001 & " d " & _
                " where " & DBW("a.workarea=", sWorkarea) & _
                " and   " & DBW("a.accdt   =", sAccDt) & _
                " and   " & DBW("a.accseq  =", iAccSeq) & _
                " and   " & DBW("a.workcd=", pWorkCode) & _
                " and   b.ptid   = a.ptid " & _
                " and   b.orddt  = a.orddt " & _
                " and   b.ordno  = a.ordno " & _
                " and   b.ordseq = a.ordseq " & _
                " and   " & DBW("c.workcd=", pWorkCode) & _
                " and   c.testcd = a.testcd " & _
                " and   d.testcd = a.testcd  " & _
                " and  (d.panelfg = ''  or  d.panelfg is null) " & _
                " and   d.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = d.testcd) " & _
                " group by d.rptseq, a.testcd, b.mesg "
                    
    sSqlLab302 = sSqlLab302 & "union select d.rptseq, a.testcd, b.mesg " & _
                    " from  " & T_LAB351 & " a, " & T_LAB102 & " b, " & T_LAB008 & " c, " & T_LAB001 & " d " & _
                    " where " & DBW("a.workarea=", sWorkarea) & _
                    " and   " & DBW("a.accdt=", sAccDt) & _
                    " and   " & DBW("a.accseq = ", iAccSeq) & _
                    " and   b.ptid  = a.ptid  " & _
                    " and   b.orddt = a.orddt " & _
                    " and   b.ordno = a.ordno " & _
                    " and   b.ordseq = a.ordseq " & _
                    " and   " & DBW("c.workcd=", pWorkCode) & _
                    " and   c.testcd = a.testcd" & _
                    " and   d.testcd = a.testcd" & _
                    " and  (d.panelfg = ''  or  d.panelfg is null) " & _
                    " and   d.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = d.testcd) " & _
                    " group by d.rptseq, a.testcd, b.mesg  order by rptseq, testcd"

                    '"     order by 1 asc "
        
    Set rslab302 = New Recordset
    rslab302.Open sSqlLab302, DBConn
        
    If rslab302.EOF Then

        Call ChangeLine(iCm)
        Call DrawLine(iCm / 6, iCurY, iPageWidth - iCm / 8, iCurY, "dash", 2, 0)
        Exit Sub
    End If
    
    rslab302.MoveFirst
         
    strRemark = ""
    ReDim aBracketsTestCD(1 To rslab302.RecordCount)
    For i = 1 To rslab302.RecordCount
        For j = 1 To UBound(aWSTestCD)
            If Trim("" & rslab302.Fields("testcd").Value) = Trim(aWSTestCD(j)) Then
                aBracketsTestCD(i) = j
                If Trim("" & rslab302.Fields("mesg").Value) <> "" Then
                     strRemark = strRemark & Trim("" & rslab302.Fields("mesg").Value)
                     strRemark = Replace(strRemark, Chr(10), ",")
                     strRemark = Replace(strRemark, Chr(13), ",")
                End If
                Exit For
            End If
        Next j
        rslab302.MoveNext
    Next i
    

    tmpiposXTestCD = iposXTestCD
    iOldy = iCurY

    For i = 1 To UBound(aBracketsTestCD)
        If aBracketsTestCD(i) <> 0 Then
            '최근결과를 보여준다
            Dim Rs      As Recordset
            Dim sSQL    As String
            
            sLastCd = ""
            sSQL = "select lastrst from " & T_LAB302 & " a" & _
                   " where " & _
                              DBW("a.workarea=", sWorkarea) & _
                   " and  " & DBW("a.accdt=", sAccDt) & _
                   " and  " & DBW("a.accseq=", iAccSeq) & _
                   " and  " & DBW("a.testcd=", Trim(aWSTestCD(aBracketsTestCD(i))))
            Set Rs = Nothing
            Set Rs = New Recordset
            Rs.Open sSQL, DBConn
            
            If Not Rs.EOF Then
                sLastCd = Rs.Fields("lastrst").Value & ""
            End If
            Set Rs = Nothing
            iYcnt = aBracketsTestCD(i) \ 20   ' 몫
            iXcnt = aBracketsTestCD(i) Mod 20 '나머지
        
            iCurY = iOldy + iYcnt * (iCm / 2)
        
            If iXcnt = 0 Then
                tmpiposXTestCD = iposXTestCD + iWidthTestCD * 4
            Else
                tmpiposXTestCD = iposXTestCD + iWidthTestCD * (iXcnt - 1)
            End If
        
            If sLastCd <> "" Then
                Call WriteStr(iCurY, tmpiposXTestCD, "(" & sLastCd & ")", iCurY, 0)
            Else
                Call WriteStr(iCurY, tmpiposXTestCD, "(       )", iCurY, 0)
            End If
            
        End If
    Next i
    
    Set rslab302 = Nothing
    
    iOldy = iCurY
    Printer.FontSize = 8
    Call WriteStr(iCurY + iCm / 1.8, iposXTestCD + iCm, strRemark, iCurY, 0)
    iCurY = iOldy
    Printer.FontSize = 9
    Call ChangeLine(iCm)
    
    '================================================
    '관련검사 출력
    '================================================

    Dim rs1      As Recordset
    Dim sSQL1    As String
    Dim sResult  As String
    Dim sVfyDate As String
    
    sResult = ""
    For i = 1 To UBound(aWSTestCD)
        
        sSQL1 = SqlRelTest(sWorkarea, sAccDt, CStr(iAccSeq), aWSTestCD(i))
        Set rs1 = Nothing
        Set rs1 = New Recordset
        rs1.Open sSQL1, DBConn
        
        '** 수정 ----------------------------------------------------------------------------------------------
        '   관련검사 단일 --> 여러항목으로 변경 By M.G.Choi
        If rs1.BOF = False Then
            Do Until rs1.EOF = True
                
                If rs1.Fields("vfydt").Value & "" <> "" Then
                    sVfyDate = Format(rs1.Fields("vfydt").Value & "", "####/##/##")
                End If
                sResult = sResult & rs1.Fields("abbrnm5").Value & "" & "(" & rs1.Fields("rstcd").Value & "" & ")" & ":" & sVfyDate & ","
                
                rs1.MoveNext
            Loop
        End If
        '------------------------------------------------------------------------------------------------------
        
        '-- 원본 ----------------------------------------------------------------------------------------------
'        If Not rs1.EOF Then
'            sResult = sResult & rs1.Fields("abbrnm5").Value & "" & "(" & rs1.Fields("rstcd").Value & "" & "),"
'        End If
        '------------------------------------------------------------------------------------------------------
        
        Set rs1 = Nothing
    Next
    If sResult <> "" Then
        sResult = Mid(sResult, 1, Len(sResult) - 1)
        Call WriteStr(iCurY, 700, "▶ 관련검사결과 ◀ " & sResult, iCurY, 0)
        Call ChangeLine(300)
    End If
    
    
    Call DrawLine(iCm / 6, iCurY, iPageWidth - iCm / 8, iCurY, "dash", 2, 0)
    
End Sub

'%  47. LD, CPK 등 관련검사의 결과 조회
'%       - Calling from [ frm202AccDataEntry, frm204WSDataEntry ] :
Private Function SqlRelTest(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As String, ByVal TestCd As String) As String
      
    SqlRelTest = " select a.testcd, b.abbrnm5, a.rstcd, a.vfydt, a.vfytm " & _
                    " from   " & T_LAB302 & " a, " & T_LAB001 & " b, " & T_LAB302 & " d, " & T_LAB031 & " e " & _
                    " where  " & DBW("d.workarea", WorkArea, 2) & _
                    " and    " & DBW("d.accdt", AccDt, 2) & _
                    " and    " & DBW("d.accseq", AccSeq, 2) & _
                    " and    " & DBW("d.testcd", TestCd, 2) & _
                    " and    " & DBW("e.cdindex", LC2_RelTest, 2) & _
                    " and    e.cdval1 = d.testcd " & _
                    " and    a.ptid = d.ptid " & _
                    " and    a.testcd = e.cdval2 " & _
                    " and    a.vfydt <> ' '" & _
                    " and    a.vfydt" & FUNC_CONCAT & "a.vfytm = (select max(vfydt" & FUNC_CONCAT & "vfytm) from " & T_LAB302 & " where ptid = a.ptid and testcd = a.testcd) " & _
                    " and    b.testcd = a.testcd " & _
                    " and    b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = b.testcd) "
               
End Function
