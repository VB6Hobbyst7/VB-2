VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQcResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'접수내역(201) 변수들..
Public l201_spcyy    As String
Public l201_workarea As String
Public l201_accdt    As String
Public l201_accno    As String
Public l201_ctrlcd   As String
Public l201_ctrlnm   As String
Public l201_levelcd  As String
Public l201_lotno    As String
Public l201_eqpcd    As String
Public l201_eqpnm    As String
Public l201_makecd   As String
Public l201_remark   As String
Public l201_totcnt   As Integer
Public l201_inputcnt As Integer
Public l201_stscd    As String
Public l201_vfydt    As String
Public l201_vfytm    As String
Public l201_vfyid    As String
Public l201_footfg   As String
Public l201_rmkcd    As String
Public l201_qcfg     As String

Public l201_ptid     As String  '외부정도관리-회차
Public l201_deptcd   As String  '외부정도관리-분과
Public l201_entdt    As String  '외부정도관리-마감일
Public l201_spccd    As String  '외부정도관리-검체
Public l201_multifg  As String  '외부정도관리-검체상태
Public l201_wardid   As String  '외부정도관리-의뢰지병원
Public l201_buildcd  As String  '외부정도관리-의뢰기관
Public l201_rcvdt    As String  '외부정도관리-접수일
Public l201_rcvnm    As String  '외부정도관리-접수자
Public l201_vfynm    As String  '외부정도관리-검증전문의명

'Private mvarMyDb As DrDatabase
Private mCol026 As Collection
Private mTestCount As Integer
Private mvarErrMsg As String


'Database Connection
'Public Property Get MyDb() As DrDatabase
'    Set MyDb = mvarMyDb
'End Property
'Public Property Set MyDb(ByVal vNewValue As DrDatabase)
'    Set mvarMyDb = vNewValue
'End Property

'Test Count
Public Property Get TestCount() As Integer
    TestCount = mTestCount
End Property
Public Property Let TestCount(ByVal vNewValue As Integer)
    mTestCount = vNewValue
End Property

'Error Message
Public Property Get ErrMsg() As String
    ErrMsg = mvarErrMsg
End Property

Public Property Get Item(ByVal vntIndexKey As Variant) As Object
  Set Item = mCol026(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol026.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol026.Remove vntIndexKey
End Sub

'접수내역(201) 검색
Public Function getSlipTable(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccNo As String) As Boolean
    
    Dim strSql As String
    Dim Rs As Recordset

    
    On Error GoTo ErrorHandler
    
    getSlipTable = True
    
    strSql = "SELECT  a.spcyy, a.workarea, a.accdt, a.accseq, a.reqtotcnt, a.reqinputcnt, a.stscd,"
    strSql = strSql & " a.vfydt, a.vfytm, a.vfyid, a.footnotefg, a.rmkcd, a.qcfg, a.orgaccno, "
    strSql = strSql & " a.ptid, a.deptcd, a.spccd, a.multifg, a.wardid, a.buildcd, a.entdt, "
    strSql = strSql & " a.rcvdt, a.rcvid, b.empnm, c.empnm as doctnm "
    strSql = strSql & " FROM  " & T_LAB201 & " a, " & T_LAB015 & " b, " & T_LAB015 & " c "
    strSql = strSql & " WHERE " & DBW("a.workarea = ", pWorkArea)
    strSql = strSql & " AND   " & DBW("a.accdt    = ", pAccDt)
    strSql = strSql & " AND   " & DBW("a.accseq   = ", pAccNo)
    strSql = strSql & " AND " & DBJ("b.empid  =* a.rcvid")
    strSql = strSql & " AND " & DBJ("c.empid  =* a.vfyid")
    
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
    While (Not Rs.EOF)
        l201_spcyy = Trim("" & Rs.Fields("SpcYy").Value)
        l201_workarea = Trim("" & Rs.Fields("WorkArea").Value)
        l201_accdt = Trim("" & Rs.Fields("AccDt").Value)
        l201_accno = Trim("" & Rs.Fields("AccSeq").Value)
        l201_totcnt = Trim("" & Rs.Fields("ReqTotCnt").Value)
        l201_inputcnt = Trim("" & Rs.Fields("ReqInputCnt").Value)
        l201_stscd = Trim("" & Rs.Fields("StsCd").Value)
        l201_vfydt = Trim("" & Rs.Fields("VfyDt").Value)
        l201_vfytm = Trim("" & Rs.Fields("VfyTm").Value)
        l201_vfyid = Trim("" & Rs.Fields("VfyId").Value)
        l201_footfg = Trim("" & Rs.Fields("FootNoteFg").Value)
        l201_rmkcd = Trim("" & Rs.Fields("RmkCd").Value)
        l201_qcfg = Trim("" & Rs.Fields("QcFg").Value)
'        l201_ctrlcd = Trim("" & rs.Fields("CtrlCd").Value)
'        l201_ctrlnm = Trim("" & rs.Fields("CtrlNm").Value)
'        l201_levelcd = Trim("" & rs.Fields("LevelCd").Value)
'        l201_lotno = Trim("" & rs.Fields("OrgAccNo").Value)
        l201_ptid = Trim("" & Rs.Fields("PtId").Value)
        l201_deptcd = Trim("" & Rs.Fields("DeptCd").Value)
        l201_entdt = Trim("" & Rs.Fields("EntDt").Value)
'        l201_spccd = Trim("" & rs.Fields("SpcCd").Value)
        l201_spccd = Trim("" & Rs.Fields("OrgAccNo").Value)     '외부의뢰검체 특별처리....
        l201_multifg = Trim("" & Rs.Fields("MultiFg").Value)
        l201_wardid = Trim("" & Rs.Fields("WardId").Value)
        l201_buildcd = Trim("" & Rs.Fields("BuildCd").Value)
        l201_rcvdt = Trim("" & Rs.Fields("RcvDt").Value)
        l201_rcvnm = Trim("" & Rs.Fields("EmpNm").Value)
        l201_vfynm = Trim("" & Rs.Fields("DoctNm").Value)
        
'        FootNote = GetFootNote(l201_workarea, l201_accdt, l201_accno)
'        Remark = GetRemark(l201_rmkcd)
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    Exit Function
'
ErrorHandler:
    getSlipTable = False
    Set Rs = Nothing
    
End Function


Public Function getRstTable(ByVal pWorkArea As String, ByVal pAccDt As String, _
                            ByVal pAccNo As String, ByVal pPrgBar As Object, Optional ByVal pOutFg As Boolean = False) As Boolean

    Dim strSql As String
    Dim sKey As String
    Dim Rs As Recordset
    Dim rs1 As Recordset
    Dim obj026 As clsQcRstData
    
    On Error GoTo ErrorHandler
    
    getRstTable = True
    mTestCount = 0
    
    
    If Not pOutFg Then
        strSql = " SELECT a.workarea,a.accdt,a.accseq,a.testcd,a.rstval,a.rstcd,a.rstunit," & _
                        " a.rsttype,a.rstdiv,a.radiv,a.detailfg,a.vfydt,a.vfytm,a.vfyid,a.mfyfg," & _
                        " a.txtfg,a.autofg,a.eqpcd,a.method,c.rcvdt,c.rcvtm,a.ctrlcd,a.levelcd,a.lotno, " & _
                        " h.testnm, h.abbrnm5, h.abbrnm10, h.txttype, " & _
                        " i.avalval, i.meanval, i.sdval, i.refcd, i.cvval, i.minval, i.maxval, i.wmset, " & _
                        " j.field1 as methodnm "
        strSql = strSql & " FROM " & T_LAB026 & " a, " & T_LAB201 & " c, " & _
                                     T_LAB001 & " h, " & T_LAB024 & " i, " & T_LAB032 & " j "
        strSql = strSql & " WHERE " & _
                                            DBW("a.workarea=", pWorkArea) & _
                                  " AND " & DBW("a.accdt=", pAccDt) & _
                                  " AND " & DBW("a.accseq=", pAccNo) & _
                                  " AND h.testcd = a.testcd " & _
                                  " AND h.applydt = (select max(applydt) from " & T_LAB001 & _
                                                   " where testcd = h.testcd)" & _
                                  " AND c.workarea = a.workarea" & _
                                  " AND c.accdt = a.accdt" & _
                                  " AND c.accseq = a.accseq " & _
                                  " AND i.ctrlcd = a.ctrlcd" & _
                                  " AND i.levelcd = a.levelcd" & _
                                  " AND i.lotno = a.lotno" & _
                                  " AND i.testcd = a.testcd " & _
                                  " AND " & DBW("j.cdindex=", LC3_Method) & _
                           " ORDER BY h.rptseq, a.accdt, a.accseq"
'                                  " AND " & DBJ("j.cdval1 =* a.method")
    Else
        strSql = " SELECT a.workarea,a.accdt,a.accseq,a.testcd,a.rstval,a.rstcd,a.rstunit," & _
                        " a.rsttype,a.rstdiv,a.radiv,a.detailfg,a.vfydt,a.vfytm,a.vfyid,a.mfyfg," & _
                        " a.txtfg,a.autofg,a.eqpcd,a.method,c.rcvdt,c.rcvtm,a.ctrlcd,a.levelcd,a.lotno, " & _
                        " h.testnm, h.abbrnm5, h.abbrnm10, h.txttype, " & _
                        " i.avalval, 0 as meanval, 0 as sdval, '' as refcd, 0 as cvval, 0 as minval, 0 as maxval, '' as wmset, " & _
                        " j.field1 as methodnm "
        strSql = strSql & " FROM " & T_LAB026 & " a, " & T_LAB201 & " c, " & _
                                     T_LAB001 & " h, " & T_LAB004 & " i, " & T_LAB032 & " j "
        strSql = strSql & " WHERE " & _
                                            DBW("a.workarea=", pWorkArea) & _
                                  " AND " & DBW("a.accdt=", pAccDt) & _
                                  " AND " & DBW("a.accseq=", pAccNo) & _
                                  " AND h.testcd = a.testcd " & _
                                  " AND h.applydt = (select max(applydt) from " & T_LAB001 & _
                                                   " where testcd = h.testcd)" & _
                                  " AND c.workarea = a.workarea" & _
                                  " AND c.accdt = a.accdt" & _
                                  " AND c.accseq = a.accseq " & _
                                  " AND i.testcd = a.testcd " & _
                                  " AND i.seq = (select min(seq) from " & T_LAB004 & _
                                               " where testcd = i.testcd) " & _
                                  " AND " & DBW("j.cdindex=", LC3_Method) & _
                           " ORDER BY h.rptseq, a.accdt, a.accseq"
'                                  " AND " & DBJ("j.cdval1 =* a.method")
    End If
   '
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
    Set mCol026 = New Collection
    
    If Not Rs.EOF Then
        
        l201_ctrlcd = Trim("" & Rs.Fields("CtrlCd").Value)
        l201_levelcd = Trim("" & Rs.Fields("LevelCd").Value)
        l201_lotno = Trim("" & Rs.Fields("LotNo").Value)
        
        strSql = " SELECT a.ctrlnm, a.eqpcd, b.eqpnm, c.makecd, c.remark " & _
                 " FROM  " & T_LAB021 & " a, " & T_LAB006 & " b, " & T_LAB023 & " c " & _
                 " WHERE  " & _
                                    DBW("a.ctrlcd=", l201_ctrlcd) & _
                          " AND " & DBW("a.levelcd=", l201_levelcd) & _
                          " AND " & DBJ("b.eqpcd =* a.eqpcd") & _
                          " AND c.ctrlcd = a.ctrlcd " & _
                          " AND c.levelcd = a.levelcd " & _
                          " AND " & DBW("c.lotno=", l201_lotno)
                 
        Set rs1 = New Recordset
        rs1.Open strSql, DBConn

        If Not rs1.EOF Then
            l201_ctrlnm = Trim("" & rs1.Fields("CtrlNm").Value)
            l201_eqpcd = Trim("" & rs1.Fields("EqpCd").Value)
            l201_eqpnm = Trim("" & rs1.Fields("EqpNm").Value)
            l201_makecd = Trim("" & rs1.Fields("MakeCd").Value)
            l201_remark = Trim("" & rs1.Fields("Remark").Value)
            
            Set rs1 = Nothing
        End If
        
        pPrgBar.Min = 0
        pPrgBar.Max = Rs.RecordCount
        pPrgBar.Value = 0
        pPrgBar.Visible = True
        DoEvents
        
    End If
    
    While (Not Rs.EOF)
    
        pPrgBar.Value = pPrgBar.Value + 1
        DoEvents
        
        Set obj026 = New clsQcRstData
        With obj026
        
            .CtrlCd = l201_ctrlcd
            .LevelCd = l201_levelcd
            .Lotno = l201_lotno
        
'            Set .MyDb = MyDb
            .WorkArea = Trim("" & Rs.Fields("WorkArea").Value):    .AccDt = Trim("" & Rs.Fields("AccDt").Value)
            .AccSeq = Trim("" & Rs.Fields("AccSeq").Value):        .TestCd = Trim("" & Rs.Fields("TestCd").Value)
            .RstVal = Trim("" & Rs.Fields("RstVal").Value):        .RstCd = Trim("" & Rs.Fields("RstCd").Value)
            .RstUnit = Trim("" & Rs.Fields("RstUnit").Value):      .RstType = Trim("" & Rs.Fields("RstType").Value)
            .RstDiv = Trim("" & Rs.Fields("RstDiv").Value):        .RaDiv = Trim("" & Rs.Fields("RaDiv").Value)
            If .RaDiv = "R" Then
                .RaName = "Reject"
            ElseIf .RaDiv = "A" Then
                .RaName = "Accept"
            Else
                .RaName = ""
            End If
            '.RcvDt = Trim("" & rs.Fields("RcvDt").Value):          .RcvTm = Trim("" & rs.Fields("RcvTm").Value)
            .VfyDt = Trim("" & Rs.Fields("VfyDt").Value)
            .VfyTm = Trim("" & Rs.Fields("VfyTm").Value):          .VfyId = Trim("" & Rs.Fields("VfyId").Value)
            .DetailFg = Trim("" & Rs.Fields("DetailFg").Value)
            .MfyFg = Trim("" & Rs.Fields("MfyFg").Value):          .TxtFg = Trim("" & Rs.Fields("TxtFg").Value)
            .EqpCd = Trim("" & Rs.Fields("EqpCd").Value):          .Method = Trim("" & Rs.Fields("Method").Value)
            .AutoFg = Trim("" & Rs.Fields("AutoFg").Value):        .MethodNm = Trim("" & Rs.Fields("MethodNm").Value)
            If (.DetailFg = "") Or (.RstDiv = "*") Then
                .TestNm = Trim("" & Rs.Fields("TestNm").Value)
            Else
                .TestNm = Space(7) & Trim("" & Rs.Fields("TestNm").Value)
            End If
            .AbbrNm5 = Trim("" & Rs.Fields("AbbrNm5").Value)
            .AbbrNm10 = Trim("" & Rs.Fields("AbbrNm10").Value)
            '.TxtType = Trim("" & rs.Fields("TxtType").Value)
            
            '소견결과가 입력된 경우...
            If .TxtFg <> "" Then .GetRstTxt
            
            .MeanVal = Val("" & Rs.Fields("MeanVal").Value)
            .SdVal = Val("" & Rs.Fields("SdVal").Value)
            .AvalVal = Val("" & Rs.Fields("AvalVal").Value)
            .MinVal = Val("" & Rs.Fields("MinVal").Value)
            .MaxVal = Val("" & Rs.Fields("MaxVal").Value)
            .CvVal = Val("" & Rs.Fields("CvVal").Value)
            .WsSet = Trim("" & Rs.Fields("WmSet").Value)
            .RefCd = Trim("" & Rs.Fields("refcd").Value)
            
            Call .GetRefValues  '--1SD/2SD/3SD값 초기화
            
            sKey = .WorkArea & ":" & .AccDt & ":" & .AccSeq & ":" & .TestCd
            
        End With
        'Call Add(obj026, sKey)
        Call mCol026.Add(obj026, sKey)
        mTestCount = mTestCount + 1
        Rs.MoveNext
    Wend
    
    Call GetHistoryData(l201_ctrlcd, l201_levelcd, l201_lotno)
    
    Set Rs = Nothing
    
    pPrgBar.Visible = False
    
    Exit Function
    
ErrorHandler:
    'Resume Next
    pPrgBar.Visible = False
    getRstTable = False
    Set Rs = Nothing
    
End Function

Public Function GetFootNote(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccNo As String) As String
Dim strSql As String
Dim Rs As Recordset
   '
    GetFootNote = ""
    
    'FootNote내역
    strSql = " SELECT rsttxt FROM " & T_LAB304 & _
             " WHERE " & DBW("workarea=", WorkArea) & _
             " AND " & DBW("accdt=", AccDt) & _
             " AND " & DBW("accseq=", AccNo)
             
  '  " ORDER BY mfyseq"
   '
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
    Do Until Rs.EOF
        GetFootNote = Rs.Fields("rsttxt").Value & ""
       Rs.MoveNext
    Loop
   '
   Set Rs = Nothing
   
End Function


Private Function GetRemark(ByVal pRmkCd As String) As String
    Dim strSql As String
    Dim Rs As Recordset
   '
    GetRemark = ""
    If l201_rmkcd = "" Then Exit Function
    strSql = " SELECT text1 FROM " & T_LAB034 & _
             " WHERE " & DBW("cdindex=", LC4_Remark) & " and " & DBW("cdval1=", pRmkCd)
    '
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
    If Not Rs.EOF Then
        GetRemark = Trim("" & Rs.Fields("Text1").Value)
    End If
   
    Set Rs = Nothing
   '
End Function

Public Sub SetSpread(ssRst As Object, Optional ByVal QueryType As Integer = 1)
    
    Dim i As Integer
    Dim iCnt As Integer, strClip As String, sRef() As String
    
    strClip = ""
    iCnt = mCol026.Count
    
    ReDim sRef(1 To iCnt)
    ssRst.MaxRows = iCnt
    ssRst.ReDraw = False
    
    
    Select Case QueryType
    Case 1:
        For i = 1 To iCnt
            With mCol026.Item(i)
                ssRst.Row = i
                ssRst.Col = 1: ssRst.Value = .TestNm
                ssRst.Col = 2: ssRst.Value = Trim(.RstCd)
                ssRst.Col = 3: ssRst.Value = .RstUnit
                ssRst.Col = 4: ssRst.Value = .RaName
                ssRst.Col = 5: ssRst.Value = .EqpCd
                ssRst.Col = 6: ssRst.Value = .RstText
                ssRst.Col = 8: ssRst.Value = ""
                ssRst.Col = 9: ssRst.Value = ""
                ssRst.Col = 10: ssRst.Value = GetMean(.TestCd)
            End With
        Next i
    Case 2:
         For i = 1 To iCnt
            With mCol026.Item(i)
                ssRst.Row = i
                ssRst.Col = 1: ssRst.Value = .TestNm
                ssRst.Col = 2: ssRst.Value = Trim(.RstCd)
                ssRst.Col = 3: ssRst.Value = .RstUnit
                ssRst.Col = 4: ssRst.Value = .MethodNm
                ssRst.Col = 5: ssRst.Value = .EqpCd
                ssRst.Col = 6: ssRst.Value = .RstText
                ssRst.Col = 8: ssRst.Value = ""
                ssRst.Col = 9: ssRst.Value = ""
                ssRst.Col = 10: ssRst.Value = GetMean(.TestCd)
            End With
        Next i
    End Select
    
    Call SetColumn(ssRst, QueryType)
    
    ssRst.ReDraw = True
    
    On Error Resume Next
    ssRst.Row = 1: ssRst.Col = SCOL_RESULT
    ssRst.SetFocus
    ssRst.Action = 0    'ActionActiveCell
    
End Sub

Private Function GetMean(ByVal strTestcd As String) As String
    Dim Rs As Recordset
    Dim sSql As String

    sSql = " select meanval from " & T_LAB024 & _
           " where " & DBW("ctrlcd=", l201_ctrlcd) & _
           " and " & DBW("levelcd=", l201_levelcd) & _
           " and " & DBW("lotno=", l201_lotno) & _
           " and " & DBW("testcd=", strTestcd)
         
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Not Rs.EOF Then
        GetMean = Rs.Fields("meanval").Value & ""
    End If
    
    Set Rs = Nothing

End Function



Private Sub SetColumn(ssRst As Object, Optional ByVal intSpType As Integer = 1)

    Dim i As Integer
    Dim strTmp As String
'
'    SCOL_TESTNM = 1:   SCOL_RESULT = 2
'    SCOL_RSTUNIT = 3:  SCOL_RADIV = 4
'    SCOL_EQPNM = 9:    SCOL_TXT = 10
    
    With ssRst
        .Row = 1: .Row2 = .MaxRows
        
        '검사명
        .Col = SCOL_TESTNM:  .Col2 = SCOL_TESTNM: .BlockMode = True:  .ForeColor = &H864B24: .BlockMode = False
        '단위
        .Col = SCOL_RSTUNIT: .Col2 = SCOL_RSTUNIT: .BlockMode = True:  .ForeColor = &H80&:    .BlockMode = False
        
        If intSpType = 1 Then
            For i = 1 To .MaxRows
                .Row = i
                .Col = SCOL_RESULT    '상세대표항목 결과필드 disable
                If mCol026.Item(i).RstDiv = "*" Then .CellType = 5   'CellTypeStaticText
                
                .Col = SCOL_RADIV    'Accept/Reject
                If mCol026.Item(i).RaDiv = "R" Then
                    .ForeColor = &H7477EF '약간 붉은색  vbRed
                ElseIf mCol026.Item(i).RaDiv = "A" Then
                    .ForeColor = &HE48372 '약간 파란색  vbBlue
                End If
            Next
        End If
    
        ssRst.RowHeight(-1) = 12.5
    
    End With
    
    
    '

End Sub

Public Function ResultCheck(ssRst As Object, ByVal iRow As Integer, _
                            ByVal pValue As String, Optional ByVal pRuleFg As Boolean = True)

    With mCol026.Item(iRow)
        
        ResultCheck = .DoResultCheck(Trim(pValue), pRuleFg)
        
        Select Case ResultCheck
        Case 1 To 9:
            mvarErrMsg = "유효숫자 입력 오류. (소숫점이하 " & .AvalVal & "자리)"
            Call ssRst.SetText(SCOL_RESULT, iRow, "")
        Case 92:
            mvarErrMsg = "유효숫자 입력 오류. (정수형만 입력)"
            Call ssRst.SetText(SCOL_RESULT, iRow, "")
        Case 93:
            mvarErrMsg = "결과코드 입력 오류!"
            Call ssRst.SetText(SCOL_RESULT, iRow, "")
        Case 94:
            mvarErrMsg = "비율결과 입력 오류!"
            Call ssRst.SetText(SCOL_RESULT, iRow, "")
        Case 95:
            mvarErrMsg = "FREE결과 입력 오류! (20자리이내)"
            Call ssRst.SetText(SCOL_RESULT, iRow, "")
        Case 0:
            mvarErrMsg = ""
            If pRuleFg Then
                Call ssRst.SetText(SCOL_RADIV, iRow, .RaName)
                ssRst.Row = iRow: ssRst.Col = SCOL_RADIV
                If .RaDiv = "A" Then
                    ssRst.ForeColor = &HDF6A3E
                ElseIf .RaDiv = "R" Then
                    ssRst.ForeColor = &H7477EF
                End If
            End If
            Exit Function
        End Select
        
    End With
    
End Function


Public Function SaveResult(ssRst As Object, ByVal pEmpId As String) As Boolean
    
    Dim i As Integer
    Dim CurDetailFg As String
    Dim Ret As Boolean
    
    SaveResult = True
    
    For i = 1 To TestCount
        With mCol026.Item(i)
            
            '** 보류체크 추가 By M.G.Choi 2006.09.01
            ssRst.Row = i: ssRst.Col = 4
            
            If .DetailFg = "" Then
                If Trim(.RstCd) <> "" And ssRst.Value = "0" Then
                    .VfyDt = Format(GetSystemDate, CS_DateDbFormat)
                    .VfyTm = Format(GetSystemDate, CS_TimeDbFormat)
                    .VfyId = pEmpId
                Else
                    .VfyDt = ""
                    .VfyTm = ""
                    .VfyId = ""
                End If
            Else
                If .RstDiv = "*" Then   '상세검사 대표항목 처리
                    CurDetailFg = .DetailFg
                    Ret = DetailCheck(i, .DetailFg)
                End If
                If .DetailFg = CurDetailFg Then
                    If ssRst.Value = "1" Then
                        .VfyDt = ""
                        .VfyTm = ""
                        .VfyId = ""
                    Else
                        .VfyDt = IIf(Ret, Format(GetSystemDate, CS_DateDbFormat), "")
                        .VfyTm = IIf(Ret, Format(GetSystemDate, CS_TimeDbFormat), "")
                        .VfyId = IIf(Ret, pEmpId, "")
                    End If
                End If
            End If
        End With
    Next
    
    On Error GoTo Err_Trap
    
    DBConn.BeginTrans
    
    Ret = SaveLab026()
    If Not Ret Then GoTo Err_Trap
    
    '** 변경 ===================
    Ret = SaveLab201_New(pEmpId)
    '===========================
    
    '** 원본 ================
'    Ret = SaveLab201(pEmpId)
    '========================
    
    If Not Ret Then GoTo Err_Trap
    
    DBConn.CommitTrans

    Exit Function
    
Err_Trap:
    SaveResult = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function DetailCheck(ByVal iRow As Integer, ByVal DetailSeq As String) As Boolean

    Dim i As Integer
    Dim iReqTotCnt As Integer
    Dim iAltTotCnt As Integer
    Dim iReqInCnt As Integer
    Dim iAltInCnt As Integer
    
    DetailCheck = False
    
    iReqTotCnt = 0: iAltTotCnt = 0
    iReqInCnt = 0: iAltInCnt = 0
    
    For i = iRow + 1 To TestCount
        
        With mCol026.Item(i)
        
            If .DetailFg <> DetailSeq Then Exit For
            
            iReqTotCnt = IIf(.RstDiv = "R", iReqTotCnt + 1, iReqTotCnt)
            iAltTotCnt = IIf(.RstDiv = "A", iAltTotCnt + 1, iAltTotCnt)
            iReqInCnt = IIf((Trim(.RstCd) <> "") And (.RstDiv = "R"), iReqInCnt + 1, iReqInCnt)
            iAltInCnt = IIf((Trim(.RstCd) <> "") And (.RstDiv = "A"), iAltInCnt + 1, iAltInCnt)
            
        End With
    Next
    
    If (iReqTotCnt = iReqInCnt) And _
       ((iAltTotCnt = iAltInCnt) Or (iAltInCnt > 0)) Then
        DetailCheck = True
    Else
        DetailCheck = False
    End If

End Function

Private Function SaveLab026() As Boolean

    Dim i As Integer
    Dim Ret As Boolean
    
    On Error GoTo Err_Trap
        
    SaveLab026 = True
    For i = 1 To TestCount
'        Set mCol026.Item(i).MyDb = mvarMyDb
        Ret = mCol026.Item(i).UpdateResult
        If Not Ret Then GoTo Err_Trap
    Next
    Exit Function
        
Err_Trap:
    SaveLab026 = False
        
End Function

Private Function SaveLab201(ByVal pEmpId As String) As Boolean

    Dim SqlStmt As String
    Dim Ret As Boolean
    
    On Error GoTo Err_Trap
    
    SaveLab201 = True
    
    l201_inputcnt = GetInputCnt(1)
    
    If l201_totcnt = l201_inputcnt Then
        l201_vfydt = Format(GetSystemDate, CS_DateDbFormat)
        l201_vfytm = Format(GetSystemDate, CS_TimeDbFormat)
        l201_vfyid = l201_vfyid
        l201_stscd = StsCd_LIS_FinRst  '최종결과
    Else
        l201_vfydt = ""
        l201_vfytm = ""
        l201_vfyid = ""
        l201_stscd = StsCd_LIS_MidRst  '결과입력
    End If
        
    SqlStmt = " update " & T_LAB201 & " set " & _
                                                DBW("reqinputcnt=", l201_inputcnt, 1) & _
                                                DBW("rmkcd=", l201_rmkcd, 1) & _
                                                DBW("stscd=", l201_stscd, 1) & _
                                                DBW("vfydt=", l201_vfydt, 1) & _
                                                DBW("vfytm=", l201_vfytm, 1) & _
                                                DBW("vfyid=", l201_vfyid) & _
             " where " & _
                               DBW("workarea=", l201_workarea) & _
                     " and " & DBW("accdt=", l201_accdt) & _
                     " and " & DBW("accseq=", l201_accno)
    Call DBConn.Execute(SqlStmt)
    
    Exit Function
        
Err_Trap:
    SaveLab201 = False
        
End Function

'** 보류로 인한 프로시저 추가(변경) By M.G.Choi 2006.09.01
' * Flag => 0 : 결과등록, 1 : 보류
Private Function SaveLab201_New(ByVal pEmpId As String, _
                                Optional ByVal pFlag As String = "0") As Boolean

    Dim SqlStmt As String
    Dim Ret As Boolean
    
    On Error GoTo Err_Trap
    
    SaveLab201_New = True
    
    l201_inputcnt = GetInputCnt(1)
    
    If l201_totcnt = l201_inputcnt And pFlag = "0" Then
        l201_vfydt = Format(GetSystemDate, CS_DateDbFormat)
        l201_vfytm = Format(GetSystemDate, CS_TimeDbFormat)
        l201_vfyid = l201_vfyid
        l201_stscd = StsCd_LIS_FinRst  '최종결과
    Else
        l201_vfydt = ""
        l201_vfytm = ""
        l201_vfyid = ""
        l201_stscd = StsCd_LIS_MidRst  '결과입력
    End If
        
    SqlStmt = " update " & T_LAB201 & " set " & _
                                                DBW("reqinputcnt=", l201_inputcnt, 1) & _
                                                DBW("rmkcd=", l201_rmkcd, 1) & _
                                                DBW("stscd=", l201_stscd, 1) & _
                                                DBW("vfydt=", l201_vfydt, 1) & _
                                                DBW("vfytm=", l201_vfytm, 1) & _
                                                DBW("vfyid=", l201_vfyid) & _
             " where " & _
                               DBW("workarea=", l201_workarea) & _
                     " and " & DBW("accdt=", l201_accdt) & _
                     " and " & DBW("accseq=", l201_accno)
    Call DBConn.Execute(SqlStmt)
    
    Exit Function
        
Err_Trap:
    SaveLab201_New = False
        
End Function

Private Function GetInputCnt(ByVal iCase As Integer) As Integer

    Dim Rs As Recordset
    Dim SqlStmt As String
    
'    SqlStmt = " select count(*) as rstcnt from " & T_LAB026 & " " & _
'              " where " & _
'                        DBW("workarea=", l201_workarea) & _
'              " and " & DBW("accdt=", l201_accdt) & _
'              " and " & DBW("accseq=", l201_accno) & _
'              " and (" & DBW("rstdiv=", "*") & " or " & DBW("detailfg =", "") & ")" & _
'              " and (vfydt<>' ' or vfydt is not null) "
              
    SqlStmt = " select count(*) as rstcnt from " & T_LAB026 & " " & _
              " where " & _
                        DBW("workarea=", l201_workarea) & _
              " and " & DBW("accdt=", l201_accdt) & _
              " and " & DBW("accseq=", l201_accno) & _
              " and (vfydt<>' ' or vfydt is not null) "
              
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn

    If Not Rs.EOF Then
        GetInputCnt = Val("" & Rs.Fields("rstcnt").Value)
    Else
        GetInputCnt = 0
    End If
    
    Set Rs = Nothing
    
End Function



'Public Function GetRmkHelp() As ADODB.Recordset
'    Dim strSQL As String
'    Dim rs As New ADODB.Recordset
'   '
'    strSQL = "SELECT cdval1, text1 FROM lab034 WHERE cdindex = '" & LAB_Remark & "' "
'    '
'    rs.Open strSQL, gcConn.conn, adOpenForwardOnly
'    Set GetRmkHelp = rs
'
'End Function
'
'Public Function GetDataForMic(ByVal pL1 As String, ByVal pL2 As String, ByVal pL3 As String) As ADODB.Recordset
'
'    Dim bDbConn As Boolean
'    Dim sSql As String
'    Dim rsTmp As ADODB.Recordset
'
'    On Error GoTo ErrorHandler
'
'    sSql = "Select b.stscd,b.testcd,b.rsttype,b.ptid,f.ptnm,a.sex,a.ageday,d.spcnm,b.appendfg,b.mfyfg,a.footnotefg,a.rmkcd,b.rstcd " & _
'           "  From lab201 a, Lab026 b, fin001 f, lab007 d, lab009 e " & _
'           " Where b.workarea='" & pL1 & "' and b.accdt='" & pL2 & "' and b.accno='" & pL3 & "' " & _
'           "   and b.testdiv='2' " & _
'           "   and a.workarea=b.workarea and a.accdt=b.accdt and a.accno=b.accno " & _
'           "   and f.ptid=*b.ptid and d.spccd=*b.spccd " & _
'           "   and e.cdindex='C110' and e.cdval1=*b.testcd and e.cdval2=*b.rstcd"
'
'    Set rsTmp = New ADODB.Recordset
'    bDbConn = gcConn.DbConn
'
'    rsTmp.CursorLocation = adUseClient
'    rsTmp.CursorType = adOpenForwardOnly
'    rsTmp.Open sSql, gcConn.conn, adOpenForwardOnly
'
'    Set GetDataForMic = rsTmp
'
'    Exit Function
'
'ErrorHandler:
'
'    Set rsTmp = Nothing
'    Call gcError.ErrorLog(Err, sClassNm, "GetDataForMic")
'
'End Function
'
'Public Function GetDataForSpe(ByVal pL1 As String, ByVal pL2 As String, ByVal pL3 As String) As ADODB.Recordset
'
'    Dim bDbConn As Boolean
'    Dim sSql As String
'    Dim rsTmp As ADODB.Recordset
'
'    On Error GoTo ErrorHandler
'
'    sSql = "SELECT a.testcd, a.testnm, b.stscd, a.rsttype, b.mfyfg " & _
'           "  FROM lab001 a, Lab026 b " & _
'           " WHERE a.testcd = b.testcd and b.testdiv='1' " & _
'           "   AND b.workarea='" & pL1 & "' AND b.accdt='" & pL2 & "' AND b.accno=" & pL3
'
'    Set rsTmp = New ADODB.Recordset
'    bDbConn = gcConn.DbConn
'
'    rsTmp.CursorLocation = adUseClient
'    rsTmp.CursorType = adOpenForwardOnly
'    rsTmp.Open sSql, gcConn.conn, adOpenForwardOnly
'
'    Set GetDataForSpe = rsTmp
'
'    Exit Function
'
'ErrorHandler:
'
'    Set rsTmp = Nothing
'    Call gcError.ErrorLog(Err, sClassNm, "GetDataForSpe")
'
'End Function
'
'

Private Sub GetHistoryData(ByVal pCtrlCd As String, ByVal pLevelCd As String, ByVal pLotNo As String)
    
    Dim Rs As Recordset
    Dim SqlStmt As String
    Dim sTestCD As String
    Dim sKey As String
    Dim iCnt As Integer
    Dim tmpTestCd As String
    Dim strTmp As String
    
    Select Case ObjSysInfo.dbtype
        Case 2 '에스큐엘써버
            strTmp = " and   (a.workarea" & FUNC_CONCAT & "a.accdt" & FUNC_CONCAT & "a.accseq) <> (b.workarea" & FUNC_CONCAT & "b.accdt" & FUNC_CONCAT & "b.accseq) "
        Case 0 '오라클
            strTmp = " and   (a.workarea" & FUNC_CONCAT & "a.accdt" & FUNC_CONCAT & "a.accseq) <> (b.workarea" & FUNC_CONCAT & "b.accdt" & FUNC_CONCAT & "b.accseq) "
        Case 1 '사이베이스
            strTmp = " and   (a.workarea" & FUNC_CONCAT & "a.accdt" & FUNC_CONCAT & FUNC_CONVERT("CHAR", "a.accseq") & " <> (b.workarea" & FUNC_CONCAT & "b.accdt" & FUNC_CONCAT & FUNC_CONVERT("CHAR", "b.accseq")
    End Select
    
    SqlStmt = " select a.testcd, a.rstcd, a.vfydt, a.vfytm, a.vfydt" & FUNC_CONCAT & "a.vfytm as vfydttm,a.radiv " & _
              " from " & T_LAB026 & " a, " & T_LAB026 & " b " & _
              " where " & _
                       DBW("b.workarea=", l201_workarea) & _
              " and " & DBW("b.accdt=", l201_accdt) & _
              " and " & DBW("b.accseq=", l201_accno) & _
              " and a.ctrlcd = b.ctrlcd " & _
              " and a.levelcd = b.levelcd " & _
              " and a.lotno = b.lotno " & _
              " and a.testcd = b.testcd " & _
              " and (a.vfydt<>' ' or a.vfydt  is not null) " & _
              " " & _
              " " & _
              strTmp & _
              " order by a.testcd, a.vfydt desc, a.vfytm desc"
    Debug.Print SqlStmt
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    sTestCD = ""
    While (Not Rs.EOF)
        
        tmpTestCd = Rs.Fields("TestCd").Value
        sKey = l201_workarea & ":" & l201_accdt & ":" & l201_accno & ":" & tmpTestCd
        iCnt = mCol026.Item(sKey).HistoryCnt + 1
        
        If iCnt > 9 Then GoTo Skip   '수치결과가 아니면 Skip
        If Not IsNumeric(Rs.Fields("RstCd").Value) Then GoTo Skip   '수치결과가 아니면 Skip
'        If InStr(Rs.Fields("RstCd").Value, "+") > 0 Then GoTo Skip
'        If InStr(Rs.Fields("RstCd").Value, "-") > 0 Then GoTo Skip
        
        If mCol026.Item(sKey).VfyDt = "" Then
            mCol026.Item(sKey).HistoryCnt = iCnt
            'mCol026.Item(sKey).RstHistory(iCnt) = Val(Rs.Fields("RstCd").Value)
            mCol026.Item(sKey).RstHistory(iCnt) = Rs.Fields("RstCd").Value
            mCol026.Item(sKey).RstHistoryRaDiv(iCnt) = Rs.Fields("radiv").Value & ""
        Else
            If mCol026.Item(sKey).VfyDt & mCol026.Item(sKey).VfyTm >= Trim("" & Rs.Fields("VfyDtTm").Value) Then
                mCol026.Item(sKey).HistoryCnt = iCnt
                'mCol026.Item(sKey).RstHistory(iCnt) = Val(Rs.Fields("RstCd").Value)
                mCol026.Item(sKey).RstHistory(iCnt) = Rs.Fields("RstCd").Value
                mCol026.Item(sKey).RstHistoryRaDiv(iCnt) = Rs.Fields("radiv").Value & ""
            End If
        End If
Skip:
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
End Sub


Public Function GetHospNm(ByVal pHospCd As String) As String

    Dim Rs As Recordset
    Dim SqlStmt As String
    
    SqlStmt = " select field1, field2, field3 from " & T_LAB032 & _
              " where " & DBW("cdindex=", LC3_HospCd) & _
              " and " & DBW("cdval1=", pHospCd)
              
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Not Rs.EOF Then
        GetHospNm = Rs.Fields("Field1").Value & vbTab & Rs.Fields("Field2").Value & vbTab & Rs.Fields("Field3").Value
    Else
        GetHospNm = ""
    End If
    Set Rs = Nothing
    
End Function

Public Function GetMethodString() As String
    
    Dim Rs As Recordset
    Dim SqlStmt As String
    
    GetMethodString = ""
    
    SqlStmt = "select cdval1, field1 from " & T_LAB032 & " where " & DBW("cdindex=", LC3_Method)
    
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    While (Not Rs.EOF)
        GetMethodString = GetMethodString & Rs.Fields("CdVal1").Value & "-" & Rs.Fields("Field1").Value & vbTab
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
End Function


Public Function GetNextData(ByVal pWA As String, ByVal pAD As String, _
                            ByVal pAS As String, ByVal iOption As Integer, ByVal pOrganCd As String) As String

    Dim Rs As Recordset
    Dim SqlStmt As String
    
    If pAD Like "9*" Then
        pAD = "19" & pAD
    Else
        pAD = "20" & pAD
    End If
    
    If iOption = 0 Then 'Previous
        SqlStmt = " select max(accseq) as seq from " & T_LAB201 & _
                  " where " & DBW("workarea=", pWA) & _
                  " and " & DBW("accdt=", pAD) & _
                  " and " & DBW("buildcd=", pOrganCd) & _
                  " and " & DBW("accseq<", pAS)
    Else    'Next
        SqlStmt = " select min(accseq) as seq from " & T_LAB201 & _
                  " where " & DBW("workarea=", pWA) & _
                  " and " & DBW("accdt=", pAD) & _
                  " and " & DBW("buildcd=", pOrganCd) & _
                  " and " & DBW("accseq >", pAS)
    End If
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then
        GetNextData = ""
    Else
        GetNextData = Rs.Fields("Seq").Value & ""
    End If
    Set Rs = Nothing
End Function
