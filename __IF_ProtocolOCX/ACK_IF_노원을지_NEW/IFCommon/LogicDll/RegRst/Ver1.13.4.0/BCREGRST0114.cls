VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BCREGRST0114"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private gbPanicDelta As Boolean

Public Function ConvLabnoToExpand$(ByVal sComp5$)
    ConvLabnoToExpand = Format(DateAdd("d", Val(sComp5), "1999-01-01"), "YYYYMMDD")
End Function

Public Function ConvLabnoToComp$(ByVal sYear8&)
    Dim sConvYear$
    
    sConvYear = Left(sYear8, 4) & "-" & Mid(sYear8, 5, 2) & "-" & Mid(sYear8, 7, 2)
    
    ConvLabnoToComp = Format(DateDiff("d", "1999-01-01", sConvYear), "00000")
End Function

'Public Function RegServer(ByVal iPersonCnt As Integer, ByVal sWDate As String, ByVal sTWSeq As String, _
'                ByVal sTJDate As String, ByVal sTJGbn As String, ByVal sTJNo As String, _
'                ByVal sTRegNo As String, ByVal sTName As String, ByVal sTSex As String, _
'                ByVal sTEmer As String, ByVal sTReRun As String, ByVal sTOther As String, _
'                ByVal sTItemCnt As String, ByVal sTServerCd As String, _
'                ByVal sTRst1 As String, ByVal sTRst2 As String, _
'                Optional ByVal iMode As Integer, Optional ByVal sUser As String) As String
'
'    On Error GoTo ErrHandler
'
'    Dim i%, j%, iITcnt%, k%, m%, iTRegCnt%
'    Dim sItemCnt$, sCServerCd$, sCRst1$, sCRst2$, sTmpServerCd$, sTmpRst1$, sTmpRst2$
'    Dim sTRegSeq$, sCWSeq$, sCJDate$, sCRegNo$, sCJNo$
'
'    Call GetTestItem
'    Call GetOrdRstCfg
'    Call GetTestCdSeq
'
'    Dim adoCn As New ADODB.Connection
'    Dim adoRs As New ADODB.Recordset
'    Dim sSql$
'    Dim sRtnVal$, sBuf$, sRef$, sPanic$, sDelta$, sRegRst$
'    Dim lngTransOK&
'    Dim lngTransOK1&
'    Dim lngTransOK2&
'
'    RegServer = "OK"
'    'sUser = GetUserId
'    '=== ....
'    sUser = "30036"
'
'    adoCn.Open gRstcfg.sPath, "medpack", "medpack000"
'
'    iTRegCnt = 0
'    sTRegSeq = ""
'
'    For i = 1 To iPersonCnt
'        sCWSeq = GetByOne(sTWSeq, sTWSeq)
'        sCJDate = GetByOne(sTJDate, sTJDate)
'        sCRegNo = GetByOne(sTRegNo, sTRegNo)
'        sCJNo = GetByOne(sTJNo, sTJNo)
'        sTmpServerCd = GetByOneUserSymbol(sTServerCd, sTServerCd, Chr$(3))
'        sTmpRst1 = GetByOneUserSymbol(sTRst1, sTRst1, Chr$(3))
'        sTmpRst2 = GetByOneUserSymbol(sTRst2, sTRst2, Chr$(3))
'        sItemCnt = GetByOne(sTItemCnt, sTItemCnt)
'
'        If Len(sCJNo) = 12 Then
'            adoCn.BeginTrans
'            lngTransOK1 = 0
'
'            iITcnt = 0
'
'            sCJNo = Left(sCJDate, 2) & sCJNo
'
'            For j = 1 To Val(sItemCnt)
'                sCServerCd = Left(GetByOne(sTmpServerCd, sTmpServerCd) & Space(10), 10)
'                sCRst1 = GetByOne(sTmpRst1, sTmpRst1)
'                sCRst2 = GetByOne(sTmpRst2, sTmpRst2)
'
'                '2002-06-21 KHS modified
'                '결과가 공백이면 SERVER등록하지 않는다.
'                If Trim(sCServerCd) <> "" And Trim(sCRst1) <> "" Then
'                    'Panic, Delta Check SP call
'                    sRef = "": sPanic = "": sDelta = "": sRegRst = ""
'
'                    sSql = ""
'                    sSql = sSql & " SELECT ORDER_REGISTER( "
'                    sSql = sSql & "                        '" & Left(sCJNo, 8) & "', "
'                    sSql = sSql & "                        '0" & Right(sCJNo, 4) & "', "
'                    sSql = sSql & "                        '" & Mid(sCServerCd, 1, 1) & "', "
'                    sSql = sSql & "                        '" & Mid(sCServerCd, 2, 1) & "', "
'                    sSql = sSql & "                        '" & Mid(sCServerCd, 3, 3) & "', "
'                    sSql = sSql & "                        '" & Mid(sCServerCd, 6, 3) & "', "
'                    sSql = sSql & "                        '" & Mid(sCServerCd, 9, 2) & "', "
'                    sSql = sSql & "                        '" & sCRst1 & "' "
'                    sSql = sSql & "                       ) "
'                    sSql = sSql & "   FROM SYS.DUAL "
'
'                    adoRs.CursorType = adOpenForwardOnly
'                    adoRs.Open sSql, ActiveConnection:=adoCn
'
'                    If adoRs.EOF = True Then
'                        sBuf = String(4, Chr(5))
'                    Else
'                        sBuf = Trim("" & adoRs(0)) & Chr(5)
'                    End If
'
'                    adoRs.Close
'
'                    sRef = GetByOneUserSymbol(sBuf, sBuf, Chr(5))
'                    sPanic = GetByOneUserSymbol(sBuf, sBuf, Chr(5))
'                    sDelta = GetByOneUserSymbol(sBuf, sBuf, Chr(5))
'                    sRegRst = GetByOneUserSymbol(sBuf, sBuf, Chr(5))
'
'                    sSql = ""
'
'                    If sPanic = "P" Or sDelta = "D" Then
'                            '이상자임시테이블
'                    sSql = sSql & "RESULT_REGISTER_TMP( "
'                        Else
'                            '정상테이블
'                    sSql = sSql & "RESULT_REGISTER( "
'                        End If
'                    sSql = sSql & "                 '" & Left(sCJNo, 8) & "', "
'                    sSql = sSql & "                 '" & Mid(sCServerCd, 1, 1) & "', "
'                    sSql = sSql & "                 '" & Mid(sCServerCd, 2, 1) & "', "
'                    sSql = sSql & "                 '0" & Right(sCJNo, 4) & "', "
'                    sSql = sSql & "                 '" & Mid(sCServerCd, 3, 3) & "', "
'                    sSql = sSql & "                 '" & Mid(sCServerCd, 6, 3) & "', "
'                    sSql = sSql & "                 '" & Mid(sCServerCd, 9, 2) & "', "
'                    sSql = sSql & "                 '" & sCRst1 & "', "
'                    sSql = sSql & "                 '" & sRef & "', "
'                    sSql = sSql & "                 '" & sPanic & "', "
'                    sSql = sSql & "                 '" & sDelta & "', "
'                    sSql = sSql & "                 '" & sCRegNo & "', "
'                    sSql = sSql & "                 '" & sUser & "' "
'                    sSql = sSql & "                ) "
'
'                    adoCn.Execute sSql
'
'                    lngTransOK1 = lngTransOK1 + 1
'
'                    iITcnt = iITcnt + 1
'                End If
'            Next
'
'            If Val(lngTransOK1) = Val(iITcnt) Then
'            '모두 UPDATE 된 경우
'                adoCn.CommitTrans
'                sTRegSeq = sTRegSeq & sCWSeq & "|"
'                iTRegCnt = iTRegCnt + 1
'            Else
'                adoCn.RollbackTrans
'            End If
'        End If
'    Next
'
'    adoCn.Close
'
'    Call EditRegState(iTRegCnt, sWDate, sTRegSeq)
'
'    Exit Function
'
'ErrHandler:
'    If adoRs.State <> 0 Then
'        adoRs.Close
'    End If
'
'    If adoCn.State <> 0 Then
'        adoCn.RollbackTrans
'        adoCn.Close
'    End If
'
'    RegServer = "NO"
'End Function

Public Function RegServer(ByVal iPersonCnt As Integer, ByVal sWDate As String, ByVal sTWSeq As String, _
                ByVal sTJDate As String, ByVal sTJGbn As String, ByVal sTJNo As String, _
                ByVal sRack As String, ByVal sPos As String, _
                ByVal sTRegNo As String, ByVal sTName As String, ByVal sTSex As String, _
                ByVal sTEmer As String, ByVal sTReRun As String, ByVal sTOther As String, _
                ByVal sTItemCnt As String, ByVal sTIFSeq As String, ByVal sTServerCd As String, _
                ByVal sTRst1 As String, ByVal sTRst2 As String, ByVal sTServerCd2 As String, _
                Optional ByVal iMode As Integer, Optional ByVal sUser As String) As String
    
    On Error GoTo ErrHandler
    
    Dim i%, j%, iITcnt%, k%, m%, iTRegCnt%
    Dim sItemCnt$, sCServerCd$, sCServerCd2$, sCRst1$, sCRst2$
    Dim sTmpServerCd$, sTmpServerCd2$, sTmpRst1$, sTmpRst2$
    Dim sTmpServerCd_Rst$, sTmpRst1_Rst$, sTmpRst2_Rst$
    Dim sTRegSeq$, sCWSeq$, sCJDate$, sCRegNo$, sCName$, sCJNo$
    Dim sUserID$, sMach$, sOther$, sSvrDate$, sSvrTime$
    Dim sCRack$, sCPos$, sCIFDOID$
    
    Call GetTestItem
    Call GetOrdRstCfg
    Call GetTestCdSeq
    
    Dim adoCn As New ADODB.Connection
    Dim adoRs As New ADODB.Recordset
    Dim sSql$
    Dim sRtnVal$, sBuf$, sGroupCd$, sGroupCdChk$ ', sRef$, sPanic$, sDelta$, sRegRst$
    
    Dim sRef() As String
    Dim sPanic() As String
    Dim sDelta() As String
    Dim sRegRst() As String
    Dim sInfo() As String
    
    Dim sData() As String
    
    Dim lngTransOK&
    Dim lngTransOK1&
    Dim lngTransOK2&
    
    RegServer = "OK"
    
    sData = Split(gsMachineNm & "-", "-")
    sMach = Trim(sData(0))
    
    'sUser = GetUserId
    '=== ....
    'sUser = "30036"
    
    '--- 서버 날짜,시간조회
    sSql = ""
    sSql = sSql & " SELECT TO_CHAR(SYSDATE, 'YYYYMMDD'), TO_CHAR(SYSDATE, 'HH24MI') "
    sSql = sSql & "   FROM DUAL "
    
    adoRs.CursorType = adOpenForwardOnly
    adoRs.Open sSql, "DSN=" & gRstcfg.sPath & ";UID=EJNUSER;PWD=EJNPASS;"
    
    If adoRs.EOF = True Then
        sSvrDate = Format(Now, "YYYYMMDD")
        sSvrTime = Format(Now, "HHNN")
    Else
        adoRs.MoveFirst
        sSvrDate = Trim(adoRs.Fields(0))
        sSvrTime = Trim(adoRs.Fields(1))
    End If
    
    adoRs.Close
    Set adoRs = Nothing
    
    adoCn.Open gRstcfg.sPath, "EJNUSER", "EJNPASS"
    
    iTRegCnt = 0
    sTRegSeq = ""
        
    For i = 1 To iPersonCnt
    
        sCWSeq = GetByOne(sTWSeq, sTWSeq)
        sCJDate = GetByOne(sTJDate, sTJDate)
        sCRegNo = GetByOne(sTRegNo, sTRegNo)
        sCName = GetByOne(sTName, sTName)
        sCJNo = GetByOne(sTJNo, sTJNo)
        sUserID = GetByOne(sTEmer, sTEmer)
'        sMach = GetByOne(sTReRun, sTReRun)
        sOther = GetByOne(sTOther, sTOther)
        sInfo = Split(sOther, Chr(6))
        
        '<09/01/08 lsm 추가
        sCRack = GetByOne(sRack, sRack)
        sCPos = GetByOne(sPos, sPos)
        sCIFDOID = Format(sCRack & sCPos, "00000000")
        '>
        
        sTmpServerCd = GetByOneUserSymbol(sTServerCd, sTServerCd, Chr$(3))
        sTmpServerCd2 = GetByOneUserSymbol(sTServerCd2, sTServerCd2, Chr$(3))
        sTmpRst1 = GetByOneUserSymbol(sTRst1, sTRst1, Chr$(3))
        sTmpRst2 = GetByOneUserSymbol(sTRst2, sTRst2, Chr$(3))
        
        sItemCnt = GetByOne(sTItemCnt, sTItemCnt)
        
        adoCn.BeginTrans
        lngTransOK1 = 0
        
        iITcnt = 0
        
        For j = 1 To Val(sItemCnt)
            sCServerCd = GetByOne(sTmpServerCd, sTmpServerCd)
            sCRst1 = GetByOne(sTmpRst1, sTmpRst1)
            sCRst2 = GetByOne(sTmpRst2, sTmpRst2)
                
            If Trim(sCServerCd) <> "" Then
                
                'ACRCPNWH(수치결과정보)
                sSql = ""
                sSql = sSql & " UPDATE ACRCPNWH "
                sSql = sSql & "    SET CPNWMACH = '" & sMach & "', "
                If IsNumeric(sCRst1) = True Then
                sSql = sSql & "        CPNWNVAL =  " & sCRst1 & " , "
                End If
                sSql = sSql & "        CPNWRSLT = '" & sCRst1 & "', "
                sSql = sSql & "        CPNWSTAT = '1' "
                sSql = sSql & "  WHERE CPNWDATE = '" & sInfo(0) & "' "
                sSql = sSql & "    AND CPNWSLIP = '" & sInfo(1) & "' "
                sSql = sSql & "    AND CPNWITEM = '" & sInfo(2) & "' "
                sSql = sSql & "    AND CPNWOITP = '" & sInfo(3) & "' "
                sSql = sSql & "    AND CPNWWKNO =  " & sInfo(4) & "  "
                sSql = sSql & "    AND CPNWSMS1 = '" & sInfo(5) & "' "
                sSql = sSql & "    AND CPNWHPFG = '" & sInfo(6) & "' "
                sSql = sSql & "    AND CPNWIDNO = '" & sCRegNo & "' "
                sSql = sSql & "    AND CPNWCODE = '" & sCServerCd & "' "
                sSql = sSql & "    AND CPNWSTAT != '7' "
                
                adoCn.Execute sSql, lngTransOK1
                
                lngTransOK1 = lngTransOK1 + 1
                
                iITcnt = iITcnt + 1
                
            End If
        Next j
        
'''        'ACRCPNWH(수치결과정보) - 대표코드도적용
'''        sSql = ""
'''        sSql = sSql & " UPDATE ACRCPNWH "
'''        sSql = sSql & "    SET CPNWMACH = '" & sMach & "', "
'''        sSql = sSql & "        CPNWSTAT = '7' "
'''        sSql = sSql & "  WHERE CPNWDATE = '" & sInfo(0) & "' "
'''        sSql = sSql & "    AND CPNWSLIP = '" & sInfo(1) & "' "
'''        sSql = sSql & "    AND CPNWITEM = '" & sInfo(2) & "' "
'''        sSql = sSql & "    AND CPNWOITP = '" & sInfo(3) & "' "
'''        sSql = sSql & "    AND CPNWWKNO =  " & sInfo(4) & "  "
'''        sSql = sSql & "    AND CPNWSMS1 = '" & sInfo(5) & "' "
'''        sSql = sSql & "    AND CPNWHPFG = '" & sInfo(6) & "' "
'''        sSql = sSql & "    AND CPNWIDNO = '" & sCRegNo & "' "
''''''        sSql = sSql & "    AND CPNWSTAT != '7' "
'''
'''        adoCn.Execute sSql, lngTransOK1

        'ACAWNIFH(임상병리장부번호)
        sSql = ""
        sSql = sSql & " UPDATE ACAWNIFH "
        sSql = sSql & "    SET WNIFRPDT = '" & sSvrDate & "', "
        sSql = sSql & "        WNIFRPTM = '" & sSvrTime & "', "
        sSql = sSql & "        WNIFSTAT = '1', "
        sSql = sSql & "        WNIFDOID = '" & sCIFDOID & "' "
        sSql = sSql & "  WHERE WNIFDATE = '" & sInfo(0) & "' "
        sSql = sSql & "    AND WNIFSLIP = '" & sInfo(1) & "' "
        sSql = sSql & "    AND WNIFITEM = '" & sInfo(2) & "' "
        sSql = sSql & "    AND WNIFOITP = '" & sInfo(3) & "' "
        sSql = sSql & "    AND WNIFWKNO =  " & sInfo(4) & "  "
        sSql = sSql & "    AND WNIF1SMS = '" & sInfo(5) & "' "
        sSql = sSql & "    AND WNIFHPFG = '" & sInfo(6) & "' "
        sSql = sSql & "    AND WNIFIDNO = '" & sCRegNo & "' "
        sSql = sSql & "    AND WNIFSTAT != '7' "

        adoCn.Execute sSql, lngTransOK1

'''        'OCSOPSLH(외래(/응급)처방내역)
'''        sSql = ""
'''        sSql = sSql & " UPDATE OCSOPSLH "
'''        sSql = sSql & "    SET OPSLSTAT = '7' "
'''        sSql = sSql & "  WHERE OPSLSMYR = '" & Mid(sCJNo, 1, 2) & "' "
'''        sSql = sSql & "    AND OPSLSMSN = '" & Mid(sCJNo, 3, 7) & "' "
'''        sSql = sSql & "    AND OPSLSMS1 = '" & Mid(sCJNo, 10, 1) & "' "
'''        sSql = sSql & "    AND OPSLIDNO = '" & sCRegNo & "' "
'''
'''        adoCn.Execute sSql, lngTransOK1
        
        adoCn.CommitTrans
        sTRegSeq = sTRegSeq & sCWSeq & "|"
        iTRegCnt = iTRegCnt + 1
        
    Next i
    
    adoCn.Close
        
    Call EditRegState(iTRegCnt, sWDate, sTRegSeq)
    
    Exit Function
    
ErrHandler:
    If adoRs.State <> 0 Then
        adoRs.Close
    End If
    
    If adoCn.State <> 0 Then
        adoCn.RollbackTrans
        adoCn.Close
    End If
    
    RegServer = "NO"
End Function

Public Function RegServer_GEM3000(ByVal iPersonCnt As Integer, ByVal sWDate As String, ByVal sTWSeq As String, _
                ByVal sTJDate As String, ByVal sTJGbn As String, ByVal sTJNo As String, _
                ByVal sTRegNo As String, ByVal sTName As String, ByVal sTSex As String, _
                ByVal sTEmer As String, ByVal sTReRun As String, ByVal sTOther As String, _
                ByVal sTItemCnt As String, ByVal sTServerCd As String, _
                ByVal sTRst1 As String, ByVal sTRst2 As String, ByVal sTServerCd2 As String, _
                Optional ByVal iMode As Integer, Optional ByVal sUser As String) As String
    
    On Error GoTo ErrHandler
    
    Dim i%, j%, iITcnt%, k%, m%, iTRegCnt%
    Dim sItemCnt$, sCServerCd$, sCServerCd2$, sCRst1$, sCRst2$
    Dim sTmpServerCd$, sTmpServerCd2$, sTmpRst1$, sTmpRst2$
    Dim sTmpServerCd_Rst$, sTmpRst1_Rst$, sTmpRst2_Rst$
    Dim sTRegSeq$, sCWSeq$, sCJDate$, sCRegNo$, sCName$, sCJNo$
    Dim sUserID$, sMach$, sOther$, sSvrDate$, sSvrTime$
    
    Call GetTestItem
    Call GetOrdRstCfg
    Call GetTestCdSeq
    
    Dim adoCn As New ADODB.Connection
    Dim adoRs As New ADODB.Recordset
    Dim sSql$
    Dim sRtnVal$, sBuf$, sGroupCd$, sGroupCdChk$ ', sRef$, sPanic$, sDelta$, sRegRst$
    
    Dim sRef() As String
    Dim sPanic() As String
    Dim sDelta() As String
    Dim sRegRst() As String
    Dim sInfo() As String
    
    Dim lngTransOK&
    Dim lngTransOK1&
    Dim lngTransOK2&
    
    RegServer_GEM3000 = "OK"
    'sUser = GetUserId
    '=== ....
    'sUser = "30036"
    
    '--- 서버 날짜,시간조회
    sSql = ""
    sSql = sSql & " SELECT TO_CHAR(SYSDATE, 'YYYYMMDD'), TO_CHAR(SYSDATE, 'HH24MI') "
    sSql = sSql & "   FROM DUAL "
    
    adoRs.CursorType = adOpenForwardOnly
    adoRs.Open sSql, "DSN=" & gRstcfg.sPath & ";UID=EJNUSER;PWD=EJNPASS;"
    
    If adoRs.EOF = True Then
        sSvrDate = Format(Now, "YYYYMMDD")
        sSvrTime = Format(Now, "HHNN")
    Else
        adoRs.MoveFirst
        sSvrDate = Trim(adoRs.Fields(0))
        sSvrTime = Trim(adoRs.Fields(1))
    End If
    
    adoRs.Close
    Set adoRs = Nothing
    
    adoCn.Open gRstcfg.sPath, "EJNUSER", "EJNPASS"
    
    iTRegCnt = 0
    sTRegSeq = ""
        
    For i = 1 To iPersonCnt
    
        sCWSeq = GetByOne(sTWSeq, sTWSeq)
        sCJDate = GetByOne(sTJDate, sTJDate)
        sCRegNo = GetByOne(sTRegNo, sTRegNo)
        sCName = GetByOne(sTName, sTName)
        sCJNo = GetByOne(sTJNo, sTJNo)
        sUserID = GetByOne(sTEmer, sTEmer)
        sMach = GetByOne(sTReRun, sTReRun)
        sOther = GetByOne(sTOther, sTOther)
        sInfo = Split(sOther, Chr(6))
        
        sTmpServerCd = GetByOneUserSymbol(sTServerCd, sTServerCd, Chr$(3))
        sTmpServerCd2 = GetByOneUserSymbol(sTServerCd2, sTServerCd2, Chr$(3))
        sTmpRst1 = GetByOneUserSymbol(sTRst1, sTRst1, Chr$(3))
        sTmpRst2 = GetByOneUserSymbol(sTRst2, sTRst2, Chr$(3))
        
        sItemCnt = GetByOne(sTItemCnt, sTItemCnt)
        
        adoCn.BeginTrans
        lngTransOK1 = 0
        
        iITcnt = 0
        
        For j = 1 To Val(sItemCnt)
            sCServerCd = GetByOne(sTmpServerCd, sTmpServerCd)
            sCRst1 = GetByOne(sTmpRst1, sTmpRst1)
            sCRst2 = GetByOne(sTmpRst2, sTmpRst2)
                
            If Trim(sCServerCd) <> "" Then
                
                'ACRCPNWH(수치결과정보)
                sSql = ""
                sSql = sSql & " UPDATE ACRCPNWH "
                sSql = sSql & "    SET CPNWMACH = '" & sMach & "', "
                If IsNumeric(sCRst1) = True Then
                sSql = sSql & "        CPNWNVAL =  " & sCRst1 & " , "
                End If
                sSql = sSql & "        CPNWRSLT = '" & sCRst1 & "', "
                sSql = sSql & "        CPNWSTAT = '7' "
                sSql = sSql & "  WHERE CPNWDATE = '" & sInfo(0) & "' "
                sSql = sSql & "    AND CPNWSLIP = '" & sInfo(1) & "' "
                sSql = sSql & "    AND CPNWITEM = '" & sInfo(2) & "' "
                sSql = sSql & "    AND CPNWOITP = '" & sInfo(3) & "' "
                sSql = sSql & "    AND CPNWWKNO =  " & sInfo(4) & "  "
                sSql = sSql & "    AND CPNWSMS1 = '" & sInfo(5) & "' "
                sSql = sSql & "    AND CPNWHPFG = '" & sInfo(6) & "' "
                sSql = sSql & "    AND CPNWIDNO = '" & sCRegNo & "' "
                sSql = sSql & "    AND CPNWCODE = '" & sCServerCd & "' "
'''                sSql = sSql & "    AND CPNWSTAT != '7' "
                
                adoCn.Execute sSql, lngTransOK1
                
                lngTransOK1 = lngTransOK1 + 1
                
                iITcnt = iITcnt + 1
                
            End If
        Next j
        
        'ACRCPNWH(수치결과정보) - 대표코드도적용
        sSql = ""
        sSql = sSql & " UPDATE ACRCPNWH "
        sSql = sSql & "    SET CPNWMACH = '" & sMach & "', "
        sSql = sSql & "        CPNWSTAT = '7' "
        sSql = sSql & "  WHERE CPNWDATE = '" & sInfo(0) & "' "
        sSql = sSql & "    AND CPNWSLIP = '" & sInfo(1) & "' "
        sSql = sSql & "    AND CPNWITEM = '" & sInfo(2) & "' "
        sSql = sSql & "    AND CPNWOITP = '" & sInfo(3) & "' "
        sSql = sSql & "    AND CPNWWKNO =  " & sInfo(4) & "  "
        sSql = sSql & "    AND CPNWSMS1 = '" & sInfo(5) & "' "
        sSql = sSql & "    AND CPNWHPFG = '" & sInfo(6) & "' "
        sSql = sSql & "    AND CPNWIDNO = '" & sCRegNo & "' "
'''        sSql = sSql & "    AND CPNWSTAT != '7' "
        
        adoCn.Execute sSql, lngTransOK1
        
        'ACAWNIFH(임상병리장부번호)
        sSql = ""
        sSql = sSql & " UPDATE ACAWNIFH "
        sSql = sSql & "    SET WNIFRPDT = '" & sSvrDate & "', "
        sSql = sSql & "        WNIFRPTM = '" & sSvrTime & "', "
        sSql = sSql & "        WNIFSTAT = '7' "
        sSql = sSql & "  WHERE WNIFDATE = '" & sInfo(0) & "' "
        sSql = sSql & "    AND WNIFSLIP = '" & sInfo(1) & "' "
        sSql = sSql & "    AND WNIFITEM = '" & sInfo(2) & "' "
        sSql = sSql & "    AND WNIFOITP = '" & sInfo(3) & "' "
        sSql = sSql & "    AND WNIFWKNO =  " & sInfo(4) & "  "
        sSql = sSql & "    AND WNIF1SMS = '" & sInfo(5) & "' "
        sSql = sSql & "    AND WNIFHPFG = '" & sInfo(6) & "' "
        sSql = sSql & "    AND WNIFIDNO = '" & sCRegNo & "' "
        
        adoCn.Execute sSql, lngTransOK1
        
        'OCSOPSLH(외래(/응급)처방내역)
        sSql = ""
        sSql = sSql & " UPDATE OCSOPSLH "
        sSql = sSql & "    SET OPSLSTAT = '7' "
        sSql = sSql & "  WHERE OPSLSMYR = '" & Mid(sCJNo, 1, 2) & "' "
        sSql = sSql & "    AND OPSLSMSN = '" & Mid(sCJNo, 3, 7) & "' "
        sSql = sSql & "    AND OPSLSMS1 = '" & Mid(sCJNo, 10, 1) & "' "
        sSql = sSql & "    AND OPSLIDNO = '" & sCRegNo & "' "
        
        adoCn.Execute sSql, lngTransOK1
        
        adoCn.CommitTrans
        sTRegSeq = sTRegSeq & sCWSeq & "|"
        iTRegCnt = iTRegCnt + 1
        
    Next i
    
    adoCn.Close
        
    Call EditRegState(iTRegCnt, sWDate, sTRegSeq)
    
    Exit Function
    
ErrHandler:
    If adoRs.State <> 0 Then
        adoRs.Close
    End If
    
    If adoCn.State <> 0 Then
        adoCn.RollbackTrans
        adoCn.Close
    End If
    
    RegServer_GEM3000 = "NO"
End Function

Public Sub SetMachineInfo(ByVal sMCd$, ByVal sMNm$)
    gsMachineCd = sMCd
    gsMachineNm = sMNm
End Sub
