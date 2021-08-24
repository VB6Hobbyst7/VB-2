VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BCREGRST0113"
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
                ByVal sTRegNo As String, ByVal sTName As String, ByVal sTSex As String, _
                ByVal sTEmer As String, ByVal sTReRun As String, ByVal sTOther As String, _
                ByVal sTItemCnt As String, ByVal sTServerCd As String, _
                ByVal sTRst1 As String, ByVal sTRst2 As String, _
                Optional ByVal iMode As Integer, Optional ByVal sUser As String) As String
    
    On Error GoTo ErrHandler
    
    Dim i%, j%, iITcnt%, k%, m%, iTRegCnt%
    Dim sItemCnt$, sCServerCd$, sCRst1$, sCRst2$
    Dim sTmpServerCd$, sTmpRst1$, sTmpRst2$
    Dim sTmpServerCd_Rst$, sTmpRst1_Rst$, sTmpRst2_Rst$
    Dim sTRegSeq$, sCWSeq$, sCJDate$, sCRegNo$, sCJNo$
    
    Call GetTestItem
    Call GetOrdRstCfg
    Call GetTestCdSeq
    
    Dim adoCn As New ADODB.Connection
    Dim adoRs As New ADODB.Recordset
    Dim sSql$
    Dim sRtnVal$, sBuf$ ', sRef$, sPanic$, sDelta$, sRegRst$
    
    Dim sRef() As String
    Dim sPanic() As String
    Dim sDelta() As String
    Dim sRegRst() As String
    
    Dim lngTransOK&
    Dim lngTransOK1&
    Dim lngTransOK2&
    
    RegServer = "OK"
    'sUser = GetUserId
    '=== ....
    'sUser = "30036"
    
    adoCn.Open gRstcfg.sPath, "medi", "medi"
    
    iTRegCnt = 0
    sTRegSeq = ""
        
    For i = 1 To iPersonCnt
        sCWSeq = GetByOne(sTWSeq, sTWSeq)
        sCJDate = GetByOne(sTJDate, sTJDate)
        sCRegNo = GetByOne(sTRegNo, sTRegNo)
        sCJNo = GetByOne(sTJNo, sTJNo)
        
        sTmpServerCd = GetByOneUserSymbol(sTServerCd, sTServerCd, Chr$(3))
        sTmpRst1 = GetByOneUserSymbol(sTRst1, sTRst1, Chr$(3))
        sTmpRst2 = GetByOneUserSymbol(sTRst2, sTRst2, Chr$(3))
        
        sItemCnt = GetByOne(sTItemCnt, sTItemCnt)
        
        If Len(sCJNo) = 9 Then
            adoCn.BeginTrans
            lngTransOK1 = 0
            
            iITcnt = 0
            
            sCJNo = Format(Now, "YYYY") & sCJNo
            
            gbPanicDelta = False
            
            For j = 1 To Val(sItemCnt)
                sCServerCd = GetByOne(sTmpServerCd, sTmpServerCd)
                sCRst1 = GetByOne(sTmpRst1, sTmpRst1)
                sCRst2 = GetByOne(sTmpRst2, sTmpRst2)
                    
                '결과가 공백이면 SERVER등록하지 않는다.
                If Trim(sCServerCd) <> "" And Trim(sCRst1) <> "" Then
                
                    sSql = ""
                    sSql = sSql & " INSERT INTO CPL0842 ( "
                    sSql = sSql & "             UPD_DATE,"
                    sSql = sSql & "             SPECIMEN_SER,"
                    sSql = sSql & "             HANGMOG_CODE,"
                    sSql = sSql & "             RESULT "
                    sSql = sSql & "             ) VALUES ( "
                    sSql = sSql & "             SYSDATE , "
                    sSql = sSql & "             '" & sCJNo & "', "
                    sSql = sSql & "             '" & sCServerCd & "', "
                    sSql = sSql & "             '" & sCRst1 & "' "
                    sSql = sSql & "             ) "
                    
                    adoCn.Execute sSql, lngTransOK1
                    
                    lngTransOK1 = lngTransOK1 + 1
                    
                    iITcnt = iITcnt + 1
                    
                End If
            Next
    
            '--- 프로파일 및 개별항목이 오더에 있을수도 있으므로 한번만 전체 update해준다.
            sSql = ""
            sSql = sSql & " UPDATE CPL0841 "
            sSql = sSql & "    SET END_FLAG     = 'Y' "
            sSql = sSql & "  WHERE SPECIMEN_SER = '" & sCJNo & "' "
            'sSql = sSql & "    AND HANGMOG_CODE = '" & sCServerCd & "' "
    
            adoCn.Execute sSql, lngTransOK1
            
            adoCn.CommitTrans
            sTRegSeq = sTRegSeq & sCWSeq & "|"
            iTRegCnt = iTRegCnt + 1
            
'            If Val(lngTransOK1) = Val(iITcnt) Then
'            '모두 UPDATE 된 경우
'                adoCn.CommitTrans
'                sTRegSeq = sTRegSeq & sCWSeq & "|"
'                iTRegCnt = iTRegCnt + 1
'            Else
'                adoCn.RollbackTrans
'            End If
        End If
    Next
    
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

Public Sub SetMachineInfo(ByVal sMCd$, ByVal sMNm$)
    gsMachineCd = sMCd
    gsMachineNm = sMNm
End Sub
