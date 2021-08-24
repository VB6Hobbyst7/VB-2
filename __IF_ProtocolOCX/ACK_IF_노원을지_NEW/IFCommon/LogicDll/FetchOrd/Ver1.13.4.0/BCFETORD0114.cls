VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BCFETORD0114"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCurItemCnt As Integer
Private mPersonCnt As Integer

Public Property Get CurItemCnt() As Integer
    CurItemCnt = mCurItemCnt
End Property

Public Property Get PersonCnt() As Integer
    PersonCnt = mPersonCnt
End Property

Public Function ConvLabnoToExpand$(ByVal sComp5$)
    ConvLabnoToExpand = Format(DateAdd("d", Val(sComp5), "1999-01-01"), "YYYYMMDD")
End Function

Public Function ConvLabnoToComp$(ByVal sYear8&)
    Dim sConvYear$
    
    sConvYear = Left(sYear8, 4) & "-" & Mid(sYear8, 5, 2) & "-" & Mid(sYear8, 7, 2)
    
    ConvLabnoToComp = Format(DateDiff("d", "1999-01-01", sConvYear), "00000")
End Function

Public Function FetchOrder(ByVal sMachineCd As String, ByVal ADOCNOrd As ADODB.Connection, ByVal sGbn As String, _
            ByVal sCWDate As String, ByVal sCJDate As String, Optional ByVal sCondition As String, Optional ByVal sCJNo As String) As String
            
    On Error GoTo ErrHandler
    
    Dim sOneRow$
    Dim sPtNo$, sPtNm$, sSex$, sTmp$, sOther$, sTSeq$
    Dim sOJNo$, sOPtNm$, sOSex$, sBuf$
    Dim iITcnt%, iLineCnt%, iPersonCnt%, i%, iPosInLine%, j%, iSeqExist%
    
    Call GetTestItem
    Call GetOrdRstCfg
    Call GetTestCdSeq
    
    Dim adoRs As New ADODB.Recordset
    Dim sSql$
    Dim sRtnVal$
    
    sSql = ""
    sSql = sSql & " SELECT A.WNIFDATE, A.WNIFSLIP, A.WNIFITEM, A.WNIFOITP, A.WNIFWKNO, "
    sSql = sSql & "        A.WNIF1SMS, A.WNIFHPFG, A.WNIFIDNO, A.WNIFNAME, B.CPNWCODE "
    sSql = sSql & "   FROM ACAWNIFH A, ACRCPNWH B "
    sSql = sSql & "  WHERE A.WNIFDATE = B.CPNWDATE "
    sSql = sSql & "    AND A.WNIFSLIP = B.CPNWSLIP "
    sSql = sSql & "    AND A.WNIFITEM = B.CPNWITEM "
    sSql = sSql & "    AND A.WNIFOITP = B.CPNWOITP "
    sSql = sSql & "    AND A.WNIFWKNO = B.CPNWWKNO "
    sSql = sSql & "    AND A.WNIF1SMS = B.CPNWSMS1 "
    sSql = sSql & "    AND A.WNIFHPFG = B.CPNWHPFG "
    sSql = sSql & "    AND A.WNIFSMYR = '" & Mid(sCondition, 1, 2) & "' "
    sSql = sSql & "    AND A.WNIFSMSN = '" & Mid(sCondition, 3, 7) & "' "
    sSql = sSql & "    AND A.WNIFSMS1 = '" & Mid(sCondition, 10, 1) & "' "
    sSql = sSql & "    AND A.WNIFSTAT != 'X' "
            
'    adoRs.CursorType = adOpenForwardOnly
'    adoRs.Open sSql, "DSN=" & gOrdCfg.sPath & ";UID=EJNUSER;PWD=EJNPASS;"

    adoRs.Open sSql, ADOCNOrd, adOpenForwardOnly
        
    If adoRs.EOF = True Then
        adoRs.Close
        Set adoRs = Nothing
        FetchOrder = ""
        Exit Function
    Else
        sRtnVal = adoRs.GetString(adClipString, -1, "|", Chr$(3))
    End If
    
    adoRs.Close
    
    Set adoRs = Nothing
    
    FetchOrder = ""
    
    sTmp = sRtnVal
    
    sOther = GetByOne(sTmp, sTmp) & Chr(6)
    sOther = sOther & GetByOne(sTmp, sTmp) & Chr(6)
    sOther = sOther & GetByOne(sTmp, sTmp) & Chr(6)
    sOther = sOther & GetByOne(sTmp, sTmp) & Chr(6)
    sOther = sOther & GetByOne(sTmp, sTmp) & Chr(6)
    sOther = sOther & GetByOne(sTmp, sTmp) & Chr(6)
    sOther = sOther & GetByOne(sTmp, sTmp) & Chr(6)
    sPtNo = GetByOne(sTmp, sTmp)
    sPtNm = GetByOne(sTmp, sTmp)
    sTSeq = ""
    iITcnt = 0
    
    Do
        sOneRow = GetByOneUserSymbol(sRtnVal, sRtnVal, Chr$(3))
        
        If sOneRow = "" Then
            Exit Do
        End If
        
        sOneRow = sOneRow & Chr$(124)
        
        Call GetByOne(sOneRow, sOneRow)
        Call GetByOne(sOneRow, sOneRow)
        Call GetByOne(sOneRow, sOneRow)
        Call GetByOne(sOneRow, sOneRow)
        Call GetByOne(sOneRow, sOneRow)
        Call GetByOne(sOneRow, sOneRow)
        Call GetByOne(sOneRow, sOneRow)
        Call GetByOne(sOneRow, sOneRow)
        Call GetByOne(sOneRow, sOneRow)
        
        sTmp = GetByOne(sOneRow, sOneRow)
        
        sTmp = ConvertIFItemInfo(1, sTmp)
        
        If sTmp = "" Then
        Else
            iITcnt = iITcnt + 1
            sTSeq = sTSeq & sTmp & Chr$(124)
        End If
    Loop
    
    FetchOrder = sPtNo & Chr$(124) & sPtNm & Chr$(124) & sOther & Chr$(124) & CStr(iITcnt) & Chr$(124) & sTSeq
    
    Exit Function
    
ErrHandler:
    If adoRs.State <> 0 Then
        If adoRs.EOF = True Then
            FetchOrder = ""
        End If
    End If
    ViewMsg "FetchOrder - " & Err.Description
End Function

Public Function FetchOrder_GEM3000(ByVal sMachineCd As String, ByVal sPatNo As String, ByVal sDocID As String, ByVal sUserID As String, ByVal sMach As String) As Variant
            
    On Error GoTo ErrHandler
    
'    Dim sOneRow$
'    Dim sPtNo$, sNm$, sSex$, sTmp$, sTSeq$
'    Dim sOJNo$, sOPtNm$, sOSex$, sBuf$
'    Dim iITcnt%, iLineCnt%, iPersonCnt%, i%, iPosInLine%, j%, iSeqExist%
'
'    Call GetTestItem
    Call GetOrdRstCfg
'    Call GetTestCdSeq
    
    Dim ADOCN   As New ADODB.Connection
    Dim ADOCMD  As New ADODB.Command
    
    FetchOrder_GEM3000 = ""
    
    ADOCN.Open gOrdCfg.sPath, "tw_hsp_ocs", "hospital"
    
    With ADOCMD
        .ActiveConnection = ADOCN
        .CommandText = "TW_HSP_OCS.TWEXAM_AUTO.OrderInsert"
        .CommandType = adCmdStoredProc
        
        .Parameters.Append .CreateParameter("Return", adVarChar, adParamReturnValue, 200)
        .Parameters.Append .CreateParameter("@v_PtNo", adVarChar, adParamInput, 8, sPatNo)
        .Parameters.Append .CreateParameter("@v_GuBun", adVarChar, adParamInput, 1, "8")
        .Parameters.Append .CreateParameter("@v_DrCode", adVarChar, adParamInput, 6, sDocID)
        .Parameters.Append .CreateParameter("@v_OpCode", adVarChar, adParamInput, 6, sUserID)
        .Parameters.Append .CreateParameter("@v_EQCode", adVarChar, adParamInput, 4, sMach)
        
        .Execute
                
        FetchOrder_GEM3000 = ADOCMD("Return").Value
        
        If Not ADOCMD Is Nothing Then Set ADOCMD = Nothing
    End With
    
    ADOCN.Close
    
    Set ADOCN = Nothing
    
    Exit Function
    
ErrHandler:
    If Not ADOCMD Is Nothing Then Set ADOCMD = Nothing
    If ADOCN.State <> 0 Then
        ADOCN.Close: Set ADOCN = Nothing
    End If
    
    ViewMsg "FetchOrder_GEM3000 - " & Err.Description
End Function

Public Function FetchOrder_ALL(ByVal sMachineCd As String, Optional ByVal sCondition As String, Optional ByVal sGenDate As String, Optional sGenApply As String) As String
            
    On Error GoTo ErrHandler
    
    Dim sOneRow$
    Dim sTmp$, sTSeq$
    Dim sJNo$, sOJNo$, sBuf$
    Dim iITcnt%, iLineCnt%, iPersonCnt%, i%, iPosInLine%, j%, iSeqExist%
    
    Call GetTestItem
    Call GetOrdRstCfg
    Call GetTestCdSeq
    
    Dim adoRs As New ADODB.Recordset
    Dim sSql$
    Dim sRtnVal$
    
    'DIMENSION 연동일때는 일반,응급구별하지 않는다
    If Trim(sCondition) = "D" Then
        sSql = ""
        sSql = sSql & " SELECT distinct specimen_ser, hangmog_code "
        sSql = sSql & "   FROM cpl0838 "
        sSql = sSql & "  WHERE upd_date >= SYSDATE-2 "
        sSql = sSql & "    AND upd_date <  SYSDATE+1 "
        sSql = sSql & "    AND end_flag = '" & Trim(sCondition) & "' "
    Else
        If Trim(sGenApply) = "1" Then   '모두적용
            sSql = ""
            sSql = sSql & " SELECT distinct specimen_ser, hangmog_code "
            sSql = sSql & "   FROM cpl0838 "
            sSql = sSql & "  WHERE upd_date >= TO_DATE('" & sGenDate & "', 'YYYY-MM-DD') "
            sSql = sSql & "    AND upd_date <  TO_DATE('" & sGenDate & "', 'YYYY-MM-DD') + 1 "
            sSql = sSql & "    AND end_flag = 'N' "
        Else                            '응급적용
            sSql = ""
            sSql = sSql & " SELECT distinct specimen_ser, hangmog_code "
            sSql = sSql & "   FROM cpl0838 "
            sSql = sSql & "  WHERE upd_date >= SYSDATE-2 "
            sSql = sSql & "    AND upd_date <  SYSDATE+1 "
            sSql = sSql & "    AND SUBSTR(hangmog_code,6,1) = 'E' "
            sSql = sSql & "    AND end_flag = 'N' "
        End If
    End If
        
    sSql = sSql & "  ORDER BY specimen_ser, hangmog_code "
    
    adoRs.CursorType = adOpenForwardOnly
    adoRs.Open sSql, "DSN=" & gOrdCfg.sPath & ";UID=medi;PWD=medi;"
        
    If adoRs.EOF = True Then
        adoRs.Close
        Set adoRs = Nothing
        FetchOrder_ALL = "0|"
        Exit Function
    Else
        sRtnVal = adoRs.GetString(adClipString, -1, "|", Chr$(3))
    End If
    
    adoRs.Close
    
    Set adoRs = Nothing
    
    FetchOrder_ALL = "0|"
    
    sBuf = ""
    sTSeq = ""
    iITcnt = 0
    
    Do
        sOneRow = GetByOneUserSymbol(sRtnVal, sRtnVal, Chr$(3))
        
        If sOneRow = "" Then
            'FetchOrder 구성
            If Len(sTSeq) <> 0 Then
                iPersonCnt = iPersonCnt + 1
                sBuf = sBuf & sOJNo & Chr(124) & CStr(iITcnt) & Chr(124) & sTSeq & Chr(3)
            End If
            
            sBuf = CStr(iPersonCnt) & Chr(124) & sBuf
            
            If iPersonCnt = 0 Then
                sBuf = "0|"
            End If
            
            Exit Do
        End If
        
        sOneRow = sOneRow & Chr$(124)
        
        sJNo = GetByOne(sOneRow, sOneRow)
        
        '20051-1234567 -> 1-1234567
        sJNo = Mid(sJNo, 5)
        
        If sJNo <> sOJNo Then
            'FetchOrder 구성
            If Len(sTSeq) <> 0 Then
                iPersonCnt = iPersonCnt + 1
                sBuf = sBuf & sOJNo & Chr(124) & CStr(iITcnt) & Chr(124) & sTSeq & Chr(3)
            End If
            
            sTSeq = ""
            iITcnt = 0
        End If
        
        sTmp = GetByOne(sOneRow, sOneRow)
        
        sTmp = ConvertIFItemInfo(1, sTmp)
        
        If sTmp = "" Then
        Else
            iITcnt = iITcnt + 1
            sTSeq = sTSeq & sTmp & Chr$(124)
        End If
        
        sOJNo = sJNo
    Loop
    
    FetchOrder_ALL = sBuf
    
    Exit Function
    
ErrHandler:
    If adoRs.State <> 0 Then
        If adoRs.EOF = True Then
            FetchOrder_ALL = "0|"
        End If
    End If
    FetchOrder_ALL = "0|"
    ViewMsg "FetchOrder_ALL - " & Err.Description
End Function

Public Sub SetMachineInfo(ByVal sMCd$, ByVal sMNm$)
    gsMachineCd = sMCd
    gsMachineNm = sMNm
End Sub
Public Function UpdateFlag(ByVal sCJNo As String, Optional ByVal iMode As Integer, Optional ByVal sUser As String) As String
    
    On Error GoTo ErrHandler
    
    Dim sSql$
    Dim lngTransOK1&
    Dim ADOCN As New ADODB.Connection
    
    UpdateFlag = "OK"
    
    ADOCN.Open "IFSERVER", "medi", "medi"
    
    sSql = sSql & " UPDATE CPL0838 "
    sSql = sSql & "    SET END_FLAG     = 'T' "
    sSql = sSql & "  WHERE SPECIMEN_SER = '" & Format(Now, "YYYY") & sCJNo & "' "
    
    ADOCN.Execute sSql, lngTransOK1
    ADOCN.Close
    
    Exit Function
    
ErrHandler:
    If ADOCN.State <> 0 Then
        ADOCN.Close
    End If
    UpdateFlag = "NO"
End Function


