VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+--------------------------------------------------------------------------------------+
'|  1. Class 명  : clsOrders
'|  2. 기 능     : 처방내역(lab101, lab102)을 생성한다.
'|  3. 작성자    : 김미경
'|  4. 작성일    : 1999.06.01
'|  5. 작성일    : 2001.03.23
'|
'|  CopyRight(C) 1999 대련엠티에스
'+--------------------------------------------------------------------------------------+

Option Explicit

'%  클래스 clsOrder의 Data Attributes

Public Ptid As String       ' 환자ID
Public OrdDt As String      ' 처방일
Public OrdTm As String      ' 처방시간 --- 우일
Public BussDiv As String    ' 업무구분
Public BedInDt As String    ' 입원일
Public DeptCd As String     ' 진료과
Public OrdDoct As String    ' 처방의
Public MajDoct  As String   ' 주치의
Public EntId  As String     ' 입력자
Public EntDt As String      ' 입력일
Public EntTm  As String     ' 입력시간
Public DoneFg As String
Public OrdDiv As String     ' 처방구분
Public OrgAccNo As String   ' 원접수번호
Public SpOrdDiv As String   ' 특수처방구분
Public RepeatFg As String
Public ReceptNo As String   ' 영수증번호

'나중에 제거..
Public WardId As String
Public RoomId As String
Public HosilId As String


'공용변수 - 참조하는 프로젝트로부터 설정됨
'Public MyOraSE As Object     'OraSession
'Public dbConn As Object     'OraDatabase

Private OrdHeader As clsDictionary
Private OrdBody   As clsDictionary
Private LastOrdNo As Integer

Private objCollect As New clsLISCollectioin
Private objOrdSql  As New clsLISSqlOrder
Private objMySql   As New clsLISSqlStatement
Private Const DabindoCols = 14

Public Property Let BuildingCd(ByVal vData As String)
    gBuildingCd = vData
End Property

Public Property Let BuildingNm(ByVal vData As String)
    gBuildingNm = vData
End Property

Public Property Let BuildingNo(ByVal vData As Long)
    gBuildingNo = vData
End Property

'Public Sub SetDatabase(ByVal objDB As DrDatabase)
'    Set dbconn = objDB
'End Sub


'% Method 1 : 다빈도 처방항목 리스트를 Display하는 Method 이다.
Public Sub DabindoList(ByRef tblList As Object)
   
    Dim SqlStmt As String
    Dim tmpRs As Recordset
    Dim tmpStatFg As String
    Dim tmpTestFg As String
    Dim i As Integer, j As Integer
    Dim iRow As Integer
    Dim iCol As Integer
   
    SqlStmt = objOrdSql.SqlItemList(1)
       
    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, DBConn
    
    If tmpRs.EOF Then GoTo NoData
    'tmpRs.MoveFirst
   
    i = 0
    With tblList
        .ReDraw = False
        While (Not tmpRs.EOF)
            i = i + 1
            .Row = (i - 1) Mod .MaxRows + 1
            iCol = (((i - 1) \ .MaxRows) * DabindoCols)
            '.Col = iCol + 1: .Value = 0
         
'            tmpStatFg = medGetP("" & tmpRs.Fields("StatFlags").Value, 1, ";")  '건물별 응급가능 여부
'            tmpTestFg = medGetP("" & tmpRs.Fields("StatFlags").Value, 2, ";")  '건물별 검사가능 여부
         
            .Col = iCol + 2:  .Text = "" & tmpRs.Fields("TestNm").Value     ' 처방명
            .Col = iCol + 3:  .Text = "" & tmpRs.Fields("TestCd").Value     ' 처방코드
'            .Col = iCol + 4:  .Text = "" & tmpRs.Fields("SpcCd").Value     ' 검체코드
'            .Col = iCol + 5:  .Text = Mid(tmpStatFg, gBuildingNo, 1)        '** 응급여부(해당건물)
            '.Col = iCol + 5:  .Text = "" & tmpRs.Fields("StatFg").Value    ' 응급여부
'            .Col = iCol + 6:  .Text = "" & tmpRs.Fields("WorkArea").Value  ' Work Area
'            .Col = iCol + 7:  .Text = "" & tmpRs.Fields("StoreCd").Value   ' StoreCd
'            .Col = iCol + 8:  .Text = "" & tmpRs.Fields("RndFg").Value     ' 아침채혈여부
'            .Col = iCol + 9:  .Text = "" & tmpRs.Fields("TestDiv").Value   ' 검사구분
'            .Col = iCol + 10: .Text = "" & tmpRs.Fields("MultiFg").Value   ' 복수검체여부
'            .Col = iCol + 11: .Text = "" & tmpRs.Fields("SpcGrp").Value    ' 검체군
'            .Col = iCol + 12: .Text = "" & tmpRs.Fields("AbbrNm5").Value   ' 약어명
'            .Col = iCol + 13: .Text = "" & tmpRs.Fields("LabelCnt").Value  ' 라벨출력장수
'            .Col = iCol + 14: .Text = Mid(tmpTestFg, gBuildingNo, 1)        '** 검사가능여부(해당건물)
            tmpRs.MoveNext
        Wend
        .ReDraw = True
    End With
   
NoData:
    Set tmpRs = Nothing
      
End Sub


'% Method 2 : 전체 처방항목 리스트를 Display하는 Method 이다.
'Public Sub ItemList(ByRef lstList As Object, ByRef Rs As clsDictionary, Optional ByRef barStatus As Variant)
'
'    Dim i As Integer
'    Dim tmpTestCd As String
'    Dim tmpTestNm As String
'    Dim tmpStatFg As String
'    Dim tmpTestFg As String
'    Dim ColCnt As Integer
'
''    SqlStmt = objMySql.SqlItemList(2)
''    'Set tmpRs = OpenRecordSet(SqlStmt, , adOpenKeyset)
''    ColCnt = tmpRs.OpenCursor(, SqlStmt)
''
''    If ColCnt = 0 Then GoTo NoData
'    'tmpRs.MoveFirst
'
'    If Not IsMissing(barStatus) Then barStatus.Max = Rs.RecordCount + 1
'
'    DoEvents
'
'    With lstList
'        .Clear
'
'        medLockWindowUpdate (.hWnd)
'        Rs.MoveFirst
'
'        While (Not Rs.EOF)
'
'
'            tmpTestNm = Mid(Rs.Fields("TestNm"), 1, 40)
'            tmpTestNm = tmpTestNm & Space(40 - Len(tmpTestNm)) & vbTab  ' 검사명
'            tmpTestCd = Trim(Mid(Rs.Fields("TestCd"), 1, 9))
'            tmpTestCd = tmpTestCd & Space(9 - Len(tmpTestCd)) & vbTab   ' 검사코드
'
'            If Trim(tmpTestCd) <> "" Then
'
'                tmpStatFg = medGetP(Rs.Fields("StatFlags"), 1, ";") ' 건물별 응급가능 여부
'                tmpTestFg = medGetP(Rs.Fields("StatFlags"), 2, ";") ' 건물별 검사가능 여부
'
'    '            tmpStr = tmpStr & tmpRs.GetValue("SpcCd") & vbTab          ' 검체코드
'                'tmpStr = tmpStr & tmpRs.GetValue("StatFg") & vbTab         ' 응급여부
'                tmpStatFg = Mid(tmpStatFg, gBuildingNo, 1)                   '** 응급여부(해당건물)
'                Rs.Fields("statfg") = tmpStatFg
'    '            tmpStr = tmpStr & tmpRs.GetValue("WorkArea") & vbTab       ' WorkArea
'    '            tmpStr = tmpStr & tmpRs.GetValue("StoreCd") & vbTab        ' 보관구분
'    '            tmpStr = tmpStr & tmpRs.GetValue("RndFg") & vbTab          ' 아침채혈 여부
'    '            tmpStr = tmpStr & tmpRs.GetValue("TestDiv") & vbTab        ' 테스트구분
'    '            tmpStr = tmpStr & tmpRs.GetValue("MultiFg") & vbTab        ' 복수검체여부
'    '            tmpStr = tmpStr & tmpRs.GetValue("SpcGrp") & vbTab         ' 검체군
'    '            tmpStr = tmpStr & tmpRs.GetValue("AbbrNm5") & vbTab        ' 약어명
'    '            tmpStr = tmpStr & tmpRs.GetValue("LabelCnt") & vbTab       ' 라벨출력장수
'                tmpTestFg = Mid(tmpTestFg, gBuildingNo, 1)                   '** 검사가능여부(해당건물)
'                Rs.Fields("testfg") = tmpTestFg
'
'                .AddItem tmpTestNm & tmpTestCd & "1"  '검사명기준
'                .AddItem tmpTestCd & tmpTestNm & "2"  '검사코드기준
'
'            End If
'
'            If Not IsMissing(barStatus) Then barStatus.Value = barStatus.Value + 1
'
'            DoEvents
'            Rs.MoveNext
'        Wend
'        .Visible = False
'        medLockWindowUpdate (0&)
'
'    End With
'
''NoData:
'    'tmpRs.RsClose
''    tmpRs.CloseCursor
''    Set tmpRs = Nothing
'
'End Sub

Public Sub ItemList(ByRef lstList As Object, Optional ByRef barStatus As Variant)
   
    Dim i As Integer
    Dim tmpTestCd As String
    Dim tmpTestNm As String
    Dim tmpStatFg As String
    Dim tmpTestFg As String
    
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " SELECT a.testnm, a.abbrnm5, a.testcd, b.spccd, b.statfg, a.workarea, b.storecd, b.rndfg, " & _
            "        b.labelcnt, b.statflags, a.testdiv, c.field1 as MultiFg, c.field2 as SpcGrp, c.field5 as SpcNm, " & _
            "        d.field2 as LabDiv, e.field2 as LabRange, '1' InsurFg " & _
            " FROM " & T_LAB032 & " c, " & T_LAB032 & " d, " & T_LAB032 & " e, " & _
                       T_LAB004 & " b, " & T_LAB001 & " a " & _
            " WHERE  a.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
            "                     WHERE testcd = a.testcd ) " & _
            " AND   (a.detailfg = '' or a.detailfg is null) " & _
            " AND    a.testcd = b.testcd " & _
            " AND    b.seq = ( SELECT min(seq) FROM " & T_LAB004 & _
            "                  WHERE testcd = b.testcd ) " & _
            " AND   (b.expdt = '' or b.expdt is null)" & _
            " AND    b.applydt = ( SELECT max(applydt) FROM " & T_LAB004 & _
            "                      WHERE testcd = b.testcd AND spccd = b.spccd AND seq=b.seq) " & _
            " AND    c.cdindex = 'C215' " & _
            " AND    c.cdval1 = b.spccd  " & _
            " AND    d.cdindex = 'C213' " & _
            " AND    d.cdval1 = a.workarea " & _
            " AND    " & DBJ("e.cdindex = 'C217'") & _
            " AND    " & DBJ("e.cdval1 =* c.field2")
    
    Set Rs = New Recordset
    
    Rs.Open strSQL, DBConn
    
    If Not IsMissing(barStatus) Then barStatus.Max = Rs.RecordCount + 1
   
    DoEvents
   
    With lstList
        .Clear
        
        medLockWindowUpdate (.hWnd)
        Rs.MoveFirst
        
        While (Not Rs.EOF)
            
             
            tmpTestNm = Mid(Rs.Fields("TestNm").Value & "", 1, 40)
            tmpTestNm = tmpTestNm & Space(40 - Len(tmpTestNm)) & vbTab  ' 검사명
            tmpTestCd = Trim(Mid(Rs.Fields("TestCd").Value & "", 1, 9))
            tmpTestCd = tmpTestCd & Space(9 - Len(tmpTestCd)) & vbTab   ' 검사코드
             
            If Trim(tmpTestCd) <> "" Then
            
                tmpStatFg = medGetP(Rs.Fields("StatFlags").Value & "", 1, ";") ' 건물별 응급가능 여부
                tmpTestFg = medGetP(Rs.Fields("StatFlags").Value & "", 2, ";") ' 건물별 검사가능 여부
         
'                tmpStatFg = Mid(tmpStatFg, gBuildingNo, 1)                   '** 응급여부(해당건물)
'                Rs.Fields("statfg") = tmpStatFg
'                tmpTestFg = Mid(tmpTestFg, gBuildingNo, 1)                   '** 검사가능여부(해당건물)
'                Rs.Fields("testfg") = tmpTestFg
         
                .AddItem tmpTestNm & tmpTestCd & "1"  '검사명기준
                .AddItem tmpTestCd & tmpTestNm & "2"  '검사코드기준
                
            End If
         
            If Not IsMissing(barStatus) Then barStatus.Value = barStatus.Value + 1
         
            DoEvents
            Rs.MoveNext
        Wend
        .Visible = False
        medLockWindowUpdate (0&)
        
    End With
    
    Set Rs = Nothing
End Sub




'% 다빈도 처방 리스트를 Click하면 Ordersheet으로 그 내용을 전달한다.
Public Sub DabindoListClick(ByVal Col As Long, ByVal Row As Long, _
                            ByVal tblList As Object, ByRef tblOrder As Object, _
                            ByVal ColDate As String) ', ByVal rs As clsDictionary)
    
    Dim tmpValue As Variant
    Dim tmpOrdNm As Variant, tmpOrdCd As Variant
    Dim objSQL As clsLISSqlStatement
    Dim Rs As Recordset
    
    Set objSQL = New clsLISSqlStatement
    Set Rs = New Recordset

    With tblOrder
        If .DataRowCnt <= .MaxRows Then .MaxRows = .MaxRows + 1
        .Row = .DataRowCnt + 1
        .Col = enORDSHEET.tcTESTNM:     Call tblList.GetText(tblList.Col + 1, Row, tmpOrdNm): .Value = tmpOrdNm  ' 처방명
        .Col = enORDSHEET.tcTESTCD:     Call tblList.GetText(tblList.Col + 2, Row, tmpOrdCd): .Value = tmpOrdCd ' 처방코드
        
'        Call rs.KeyChange(tmpOrdCd)
        
        Rs.Open objSQL.GetItemInfo(tmpOrdCd), DBConn
        
        .Col = enORDSHEET.tcINSURFG:    .Value = Rs.Fields("insurfg").Value & ""   ' 급여구분
        .Col = enORDSHEET.tcSPCCD:      .Value = Rs.Fields("spccd").Value & ""     ' 검체코드
        .Col = enORDSHEET.tcREQDTTM:    .Value = ColDate
        .Col = enORDSHEET.tcSTATFG:     .Value = Rs.Fields("statfg").Value & ""    ' **응급여부(해당건물)
    '***건물정보 사용
        If P_ApplyBuildingInfo Then
            If .Value = "1" Then
                .Col = enORDSHEET.tcSTATCHK: .CellType = 10     'CellTypeCheckBox
                                             .TypeCheckCenter = True
            Else
                .Col = enORDSHEET.tcSTATCHK: .CellType = 5  'CellTypeStaticText
            End If
        Else
            .Col = enORDSHEET.tcSTATCHK: .CellType = 10     'CellTypeCheckBox
                                             .TypeCheckCenter = True
        End If
        .Col = enORDSHEET.tcWORKAREA:   .Value = Rs.Fields("workarea").Value & ""  ' WorkArea
        .Col = enORDSHEET.tcSTORECD:    .Value = Rs.Fields("storecd").Value & ""   ' 보관구분
        .Col = enORDSHEET.tcRNDFG:      .Value = Rs.Fields("rndfg").Value & ""     ' 아침채혈 여부
        .Col = enORDSHEET.tcTESTDIV:    .Value = Rs.Fields("testdiv").Value & ""   ' 테스트구분(0:일반,1:기타,2:미생물)
        .Col = enORDSHEET.tcMULTIFG:    .Value = Rs.Fields("multifg").Value & ""   ' 복수검체여부
        .Col = enORDSHEET.tcSPCGRP:     .Value = Rs.Fields("spcgrp").Value & ""    ' 검체군
        .Col = enORDSHEET.tcABBRNM:     .Value = Rs.Fields("abbrnm5").Value & ""   ' 약어명
        .Col = enORDSHEET.tcBARCNT:     .Value = Rs.Fields("labelcnt").Value & ""    ' 라벨출력장수
        Dim tmpTestFg As String
'        tmpStatFg = medGetP("" & tmpRs.Fields("StatFlags").Value, 1, ";")   '건물별 응급가능 여부
        tmpTestFg = medGetP("" & Rs.Fields("StatFlags").Value, 2, ";")   '건물별 검사가능 여부
        
        .Col = enORDSHEET.tcTESTFLAG:   .Value = tmpTestFg 'Rs.Fields("testfg").Value & ""    ' **검사가능여부(해당건물)
        
    '***건물정보 사용
        If P_ApplyBuildingInfo Then
            If .Value = "1" Then
                .Col = enORDSHEET.tcBUILDCD: .Value = gBuildingCd     ' ** 해당건물에서 일반검사 가능함
                .Col = enORDSHEET.tcBUILDNM: .Value = gBuildingNm
            Else
                .Col = enORDSHEET.tcBUILDCD: .Value = CentralLab      ' ** 해당건물에서 일반검사 불가능함 --> 중앙검사실로...
                .Col = enORDSHEET.tcBUILDNM: .Value = CentralLabNm
            End If
        Else
            .Col = enORDSHEET.tcBUILDCD: .Value = gBuildingCd     ' ** 해당건물에서 일반검사 가능함
            .Col = enORDSHEET.tcBUILDNM: .Value = LABName
        End If
        .Col = enORDSHEET.tcSPCABBR:    .Value = Rs.Fields("spcnm").Value & ""     ' 검체약어명
        .Col = enORDSHEET.tcLABDIV:     .Value = Rs.Fields("labdiv").Value & ""    ' 접수번호 부여기준
        .Col = enORDSHEET.tcLABRANGE:   .Value = Rs.Fields("labrange").Value & ""  ' 미생물 접수번호 범위
    End With
    
    Set Rs = Nothing
    Set objSQL = Nothing
End Sub
   

'% 처방코드를 기준으로 지정검체 리스트를  띄운다.
Public Sub SpcList(ByVal ParaTestCd As String, ByRef paraSpcList As Object)
   
    Dim SqlStmt As String
    Dim tmpRs As Recordset
    Dim tmpStr As String
    Dim tmpStatFg As String
    Dim tmpTestFg As String
   
    SqlStmt = objOrdSql.SqlSpecList(ParaTestCd)
   
    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, DBConn
    
    If tmpRs.EOF Then GoTo NoData:
    paraSpcList.Clear
   
    With tmpRs
        While (Not tmpRs.EOF)
            tmpStatFg = medGetP("" & .Fields("StatFlags").Value, 1, ";")    '건물별 응급가능 여부
            tmpTestFg = medGetP("" & .Fields("StatFlags").Value, 2, ";")    '건물별 검사가능 여부
   
            tmpStr = ""
            tmpStr = tmpStr & Trim(.Fields("SpcCd").Value) & vbTab          '검체코드
            tmpStr = tmpStr & Trim(.Fields("SpcNm").Value) & vbTab          '검체명
            'tmpStr = tmpStr & Trim(.Fields("StatFg").Value) & vbTab        '응급여부
            tmpStr = tmpStr & Mid(tmpStatFg, gBuildingNo, 1) & vbTab         '**응급여부(해당건물)
            tmpStr = tmpStr & Trim(.Fields("StoreCd").Value) & vbTab        '보관구분
            tmpStr = tmpStr & Trim(.Fields("MultiFg").Value) & vbTab        '복수검체여부
            tmpStr = tmpStr & Trim(.Fields("SpcGrp").Value) & vbTab         '검체군
            tmpStr = tmpStr & Trim(.Fields("LabelCnt").Value) & vbTab       '라벨출력장수
            tmpStr = tmpStr & Mid(tmpTestFg, gBuildingNo, 1) & vbTab         '**검사가능여부(해당건물)
            tmpStr = tmpStr & Trim(.Fields("SpcAbbr").Value) & vbTab        '검체약어명
            tmpStr = tmpStr & Trim(.Fields("LabDiv").Value) & vbTab         '접수번호 부여기준
            tmpStr = tmpStr & Trim(.Fields("LabRange").Value) & vbTab       '검체군별 접수번호 부여 범위
            tmpStr = tmpStr & Trim(.Fields("Seq").Value)                    '우선순위
            paraSpcList.AddItem tmpStr
            tmpRs.MoveNext
        Wend
    End With
    
NoData:
    Set tmpRs = Nothing
End Sub

'% 처방Header/Body 내역을 클래스로 Assign한다.
Public Sub MoveData(ByRef tblOrdSheet As Object)
   
    Dim i As Integer
    Dim HSeq As Integer
    Dim BSeq As Integer
    Dim SaveKeyString As String
    Dim KeyString As String
    Dim DataString As String
   
    Set OrdHeader = New clsDictionary
    OrdHeader.Clear
    OrdHeader.FieldInialize "hseq", "OrdNo,ReqDt,ReqTm,BodyCnt"
    OrdHeader.Sort = False
    
    Set OrdBody = New clsDictionary
    OrdBody.Clear
    OrdBody.FieldInialize "hseq,bseq", "OrdNo,OrdSeq,OrdCd,SpcCd,StoreCd,DcFg,DcDt,DcNo," & _
                                       "AttrCd,ExamDt,ExamTm,ExamDoct,StsCd,StatFg,InsDiv,DoneFg"
    OrdBody.Sort = False

    SaveKeyString = ""
    With tblOrdSheet
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = enORDSHEET.tcBUILDCD:  KeyString = .Value & vbTab                    '** Delivery Location
            .Col = enORDSHEET.tcWORKAREA: KeyString = KeyString & .Value & vbTab        'WorkArea
            .Col = enORDSHEET.tcSPCCD:    KeyString = KeyString & .Value & vbTab        '검체코드
            .Col = enORDSHEET.tcSTORECD:  KeyString = KeyString & .Value & vbTab        'StoreCd
            .Col = enORDSHEET.tcSTATFG:   KeyString = KeyString & 0 & vbTab 'CStr(Val(.Value)) & vbTab   '응급여부
            .Col = enORDSHEET.tcREQDTTM:  KeyString = KeyString & .Value                '희망채취시간
         
            'HEADER ----------------------------------------------------------------
            If SaveKeyString <> KeyString Then
                    
                HSeq = OrdHeader.RecordCount + 1    '처방 Header Seq
                SaveKeyString = KeyString
                    
                If Not OrdHeader.Exists(HSeq) Then
                    OrdHeader.AddNew HSeq, ""
                    
                    .Col = enORDSHEET.tcREQDTTM: OrdHeader.Fields("ReqDt") = Format(.Value, CS_DateDbFormat)
                                                 OrdHeader.Fields("ReqTm") = Format(.Value, CS_TimeDbFormat)
                    OrdHeader.Fields("BodyCnt") = 0
                Else
                    OrdHeader.KeyChange KeyString
                End If
            End If
            
            'BODY ----------------------------------------------------------------
            BSeq = OrdHeader.Fields("BodyCnt") + 1   '처방Body Seq
            
            KeyString = HSeq & COL_DIV & BSeq
            
            If Not OrdBody.Exists(KeyString) Then
                OrdBody.AddNew KeyString, ""
            Else
                OrdBody.KeyChange KeyString
            End If
            
            .Col = enORDSHEET.tcORDNO:  .Value = HSeq
            .Col = enORDSHEET.tcTESTCD:  OrdBody.Fields("OrdCd") = .Value    '처방코드
            .Col = enORDSHEET.tcSPCCD:   OrdBody.Fields("SpcCd") = .Value    '검체코드
            .Col = enORDSHEET.tcSTORECD: OrdBody.Fields("StoreCd") = .Value  '보관구분
            .Col = enORDSHEET.tcINSURFG: OrdBody.Fields("InsDiv") = .Value              '급여구분
            .Col = enORDSHEET.tcSTATCHK: OrdBody.Fields("StatFg") = CStr(Val(.Value))   '응급여부
            .Col = enORDSHEET.tcORDSEQ: .Value = BSeq
            
            OrdBody.Fields("DcFg") = ""         'DC 여부
            OrdBody.Fields("DcDt") = ""         'DC 일자
            OrdBody.Fields("DcNo") = "0"          'DC 번호
            OrdBody.Fields("AttrCd") = ""       '속성여부
            OrdBody.Fields("ExamDt") = ""       '실제 실시일자
            OrdBody.Fields("ExamTm") = ""       '실제 실시시간
            OrdBody.Fields("ExamDoct") = ""     '실제 실시의사
            OrdBody.Fields("StsCd") = enStsCd.StsCd_LIS_Order       'Status
            OrdBody.Fields("DoneFg") = enStsCd.StsCd_LIS_Order      'Status
            
            OrdHeader.Fields("BodyCnt") = BSeq
        Next
    End With

    OrdHeader.Sort = True
    OrdBody.Sort = True
    
End Sub

'% 처방Header/Body 내역의 Array Bound를 재조정한다.
Private Function NextSeq(ByVal TableFg As Integer, Optional ByVal lngIndex As Variant) As Long
'    Dim intBnd As Integer
'    Dim lngHSeq As Long
'    SELECT Case TableFg
'        Case 1:  'Header
'            intBnd = UBound(OrdHeader)
'            ReDim Preserve OrdHeader(intBnd + 1)
'            ReDim Preserve OrdHeader(intBnd + 1).OrdBody(0)
'        Case 2:  'Body
'            If Not IsMissing(lngIndex) Then lngHSeq = Val(lngIndex)
'            intBnd = UBound(OrdHeader(lngHSeq).OrdBody)
'            ReDim Preserve OrdHeader(lngHSeq).OrdBody(intBnd + 1)
'    End SELECT
'    NextSeq = intBnd + 1
End Function

'% 클래스의 처방Header/Body 내역을 DB로 저장한다.
Public Function SaveData(ByRef StartOrdNo As Integer, Optional ByRef ProgressBar As Object = Nothing) As Boolean
   
    Dim HCnt As Integer, BCnt As Integer
    Dim i As Integer, j As Integer
    Dim sqlStmtH As String, sqlStmtB As String
    Dim strOrdNo As String
    
    On Error GoTo Err_Trap
   
    LastOrdNo = GetLastNo(Ptid, OrdDt)
    
    'HCnt = UBound(OrdHeader)
   
    '**** Transaction Start *************'
    DBConn.BeginTrans
   
' 원소스
'    For i = 1 To BCnt 'OrdHeader.RecordCount
'        'BCnt = UBound(OrdHeader(I).OrdBody)
'
'        OrdHeader.KeyChange CStr(i)
'        OrdHeader.Fields("OrdNo") = LastOrdNo + i
'
'        sqlStmtH = objOrdSql.CreateSqlHeader(Me, OrdHeader)                '처방 Header Insert Sql문 생성
'        Call DBConn.Execute(sqlStmtH)       'Sql 실행
'
'        BCnt = Val(OrdHeader.Fields("BodyCnt"))
'        For j = 1 To BCnt
'            OrdBody.KeyChange CStr(i) & COL_DIV & CStr(j)
'            OrdBody.Fields("OrdNo") = OrdHeader.Fields("OrdNo")
'            sqlStmtB = objOrdSql.CreateSqlBody(Me, OrdBody)
'            Call DBConn.Execute(sqlStmtB)   'Sql 실행
'            If Not IsMissing(ProgressBar) Then ProgressBar.Value = ProgressBar.Value + 1
'            DoEvents
'        Next
'    Next
    
'    BCnt = Val(OrdHeader.Fields("BodyCnt"))
        
    For i = 1 To Val(OrdHeader.Fields("BodyCnt"))
        'BCnt = UBound(OrdHeader(I).OrdBody)
        
        OrdHeader.KeyChange CStr(i)
        OrdHeader.Fields("OrdNo") = LastOrdNo + i
        strOrdNo = OrdHeader.Fields("OrdNo")
        
        sqlStmtH = objOrdSql.CreateSqlHeader(Me, OrdHeader)                '처방 Header Insert Sql문 생성
        Call DBConn.Execute(sqlStmtH)       'Sql 실행
        
'        BCnt = Val(OrdHeader.Fields("BodyCnt"))
'        For j = 1 To BCnt
'            If i = j Then
'                OrdBody.KeyChange CStr(i) & COL_DIV & CStr(i)
'                OrdBody.Fields("OrdNo") = OrdHeader.Fields("OrdNo")
                OrdBody.KeyChange 1 & COL_DIV & i
                
                OrdBody.Fields("OrdNo") = OrdHeader.Fields("OrdNo")
                sqlStmtB = objOrdSql.CreateSqlBody(Me, OrdBody)
                Call DBConn.Execute(sqlStmtB)   'Sql 실행
                If Not IsMissing(ProgressBar) Then ProgressBar.Value = ProgressBar.Value + 1
                DoEvents
'            End If
'        Next
    Next
    StartOrdNo = LastOrdNo
'    StartOrdNo = strOrdNo
    
    '**** Transaction Commit ************'
    DBConn.CommitTrans
    SaveData = True
   
    Set OrdHeader = Nothing
    Set OrdBody = Nothing
    Exit Function
   
Err_Trap:
    DBConn.RollbackTrans
    
    Set OrdHeader = Nothing
    Set OrdBody = Nothing
    SaveData = False
    MsgBox Err.Description, vbExclamation
End Function

'% 가장 마지막 처방번호를 가져온다.
Public Function GetLastNo(ByVal paraPtId As String, ByVal paraOrdDt As String)
   
    Dim tmpDs As Recordset
    Dim SqlStmt As String
   
    SqlStmt = objOrdSql.SqlLastOrdNo(paraPtId, paraOrdDt)
    Set tmpDs = New Recordset
    tmpDs.Open SqlStmt, DBConn
    
    If tmpDs.EOF Then
        GetLastNo = 0
    Else
        GetLastNo = tmpDs.Fields("OrdNo").Value
    End If
    Set tmpDs = Nothing
End Function

   
Public Function CheckSameOrder(ByVal OrdTable As Object) As Integer

   Dim i As Integer, j As Integer
   Dim SaveCode As String
   Dim SaveSpc As String
   Dim SaveDate As String
   
   CheckSameOrder = 0
   With OrdTable
      For i = 1 To .DataRowCnt
         .Row = i
         .Col = enORDSHEET.tcTESTCD:    SaveCode = .Value
         .Col = enORDSHEET.tcREQDTTM:   SaveDate = .Value
         .Col = enORDSHEET.tcSPCCD:     SaveSpc = .Value
         For j = i + 1 To .DataRowCnt
            .Row = j
            .Col = enORDSHEET.tcTESTCD
            If .Value = SaveCode Then
               .Col = enORDSHEET.tcREQDTTM
               If .Value = SaveDate Then
                    .Col = enORDSHEET.tcSPCCD
                    If .Value = SaveSpc Then
                       CheckSameOrder = j
                       Exit Function
                    End If
               End If
            End If
         Next
      Next
   End With
            
End Function

Private Sub Class_Terminate()
   
    Set objMySql = Nothing
    Set OrdHeader = Nothing
    Set OrdBody = Nothing

End Sub
