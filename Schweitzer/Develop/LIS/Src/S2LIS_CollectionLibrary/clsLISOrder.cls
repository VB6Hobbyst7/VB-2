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
'|  1. Class ��  : clsOrders
'|  2. �� ��     : ó�泻��(lab101, lab102)�� �����Ѵ�.
'|  3. �ۼ���    : ��̰�
'|  4. �ۼ���    : 1999.06.01
'|  5. �ۼ���    : 2001.03.23
'|
'|  CopyRight(C) 1999 ��ÿ�Ƽ����
'+--------------------------------------------------------------------------------------+

Option Explicit

'%  Ŭ���� clsOrder�� Data Attributes

Public Ptid As String       ' ȯ��ID
Public OrdDt As String      ' ó����
Public OrdTm As String      ' ó��ð� --- ����
Public BussDiv As String    ' ��������
Public BedInDt As String    ' �Կ���
Public DeptCd As String     ' �����
Public OrdDoct As String    ' ó����
Public MajDoct  As String   ' ��ġ��
Public EntId  As String     ' �Է���
Public EntDt As String      ' �Է���
Public EntTm  As String     ' �Է½ð�
Public DoneFg As String
Public OrdDiv As String     ' ó�汸��
Public OrgAccNo As String   ' ��������ȣ
Public SpOrdDiv As String   ' Ư��ó�汸��
Public RepeatFg As String
Public ReceptNo As String   ' ��������ȣ

'���߿� ����..
Public WardId As String
Public RoomId As String
Public HosilId As String


'���뺯�� - �����ϴ� ������Ʈ�κ��� ������
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


'% Method 1 : �ٺ� ó���׸� ����Ʈ�� Display�ϴ� Method �̴�.
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
         
'            tmpStatFg = medGetP("" & tmpRs.Fields("StatFlags").Value, 1, ";")  '�ǹ��� ���ް��� ����
'            tmpTestFg = medGetP("" & tmpRs.Fields("StatFlags").Value, 2, ";")  '�ǹ��� �˻簡�� ����
         
            .Col = iCol + 2:  .Text = "" & tmpRs.Fields("TestNm").Value     ' ó���
            .Col = iCol + 3:  .Text = "" & tmpRs.Fields("TestCd").Value     ' ó���ڵ�
'            .Col = iCol + 4:  .Text = "" & tmpRs.Fields("SpcCd").Value     ' ��ü�ڵ�
'            .Col = iCol + 5:  .Text = Mid(tmpStatFg, gBuildingNo, 1)        '** ���޿���(�ش�ǹ�)
            '.Col = iCol + 5:  .Text = "" & tmpRs.Fields("StatFg").Value    ' ���޿���
'            .Col = iCol + 6:  .Text = "" & tmpRs.Fields("WorkArea").Value  ' Work Area
'            .Col = iCol + 7:  .Text = "" & tmpRs.Fields("StoreCd").Value   ' StoreCd
'            .Col = iCol + 8:  .Text = "" & tmpRs.Fields("RndFg").Value     ' ��ħä������
'            .Col = iCol + 9:  .Text = "" & tmpRs.Fields("TestDiv").Value   ' �˻籸��
'            .Col = iCol + 10: .Text = "" & tmpRs.Fields("MultiFg").Value   ' ������ü����
'            .Col = iCol + 11: .Text = "" & tmpRs.Fields("SpcGrp").Value    ' ��ü��
'            .Col = iCol + 12: .Text = "" & tmpRs.Fields("AbbrNm5").Value   ' ����
'            .Col = iCol + 13: .Text = "" & tmpRs.Fields("LabelCnt").Value  ' ��������
'            .Col = iCol + 14: .Text = Mid(tmpTestFg, gBuildingNo, 1)        '** �˻簡�ɿ���(�ش�ǹ�)
            tmpRs.MoveNext
        Wend
        .ReDraw = True
    End With
   
NoData:
    Set tmpRs = Nothing
      
End Sub


'% Method 2 : ��ü ó���׸� ����Ʈ�� Display�ϴ� Method �̴�.
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
'            tmpTestNm = tmpTestNm & Space(40 - Len(tmpTestNm)) & vbTab  ' �˻��
'            tmpTestCd = Trim(Mid(Rs.Fields("TestCd"), 1, 9))
'            tmpTestCd = tmpTestCd & Space(9 - Len(tmpTestCd)) & vbTab   ' �˻��ڵ�
'
'            If Trim(tmpTestCd) <> "" Then
'
'                tmpStatFg = medGetP(Rs.Fields("StatFlags"), 1, ";") ' �ǹ��� ���ް��� ����
'                tmpTestFg = medGetP(Rs.Fields("StatFlags"), 2, ";") ' �ǹ��� �˻簡�� ����
'
'    '            tmpStr = tmpStr & tmpRs.GetValue("SpcCd") & vbTab          ' ��ü�ڵ�
'                'tmpStr = tmpStr & tmpRs.GetValue("StatFg") & vbTab         ' ���޿���
'                tmpStatFg = Mid(tmpStatFg, gBuildingNo, 1)                   '** ���޿���(�ش�ǹ�)
'                Rs.Fields("statfg") = tmpStatFg
'    '            tmpStr = tmpStr & tmpRs.GetValue("WorkArea") & vbTab       ' WorkArea
'    '            tmpStr = tmpStr & tmpRs.GetValue("StoreCd") & vbTab        ' ��������
'    '            tmpStr = tmpStr & tmpRs.GetValue("RndFg") & vbTab          ' ��ħä�� ����
'    '            tmpStr = tmpStr & tmpRs.GetValue("TestDiv") & vbTab        ' �׽�Ʈ����
'    '            tmpStr = tmpStr & tmpRs.GetValue("MultiFg") & vbTab        ' ������ü����
'    '            tmpStr = tmpStr & tmpRs.GetValue("SpcGrp") & vbTab         ' ��ü��
'    '            tmpStr = tmpStr & tmpRs.GetValue("AbbrNm5") & vbTab        ' ����
'    '            tmpStr = tmpStr & tmpRs.GetValue("LabelCnt") & vbTab       ' ��������
'                tmpTestFg = Mid(tmpTestFg, gBuildingNo, 1)                   '** �˻簡�ɿ���(�ش�ǹ�)
'                Rs.Fields("testfg") = tmpTestFg
'
'                .AddItem tmpTestNm & tmpTestCd & "1"  '�˻������
'                .AddItem tmpTestCd & tmpTestNm & "2"  '�˻��ڵ����
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
            tmpTestNm = tmpTestNm & Space(40 - Len(tmpTestNm)) & vbTab  ' �˻��
            tmpTestCd = Trim(Mid(Rs.Fields("TestCd").Value & "", 1, 9))
            tmpTestCd = tmpTestCd & Space(9 - Len(tmpTestCd)) & vbTab   ' �˻��ڵ�
             
            If Trim(tmpTestCd) <> "" Then
            
                tmpStatFg = medGetP(Rs.Fields("StatFlags").Value & "", 1, ";") ' �ǹ��� ���ް��� ����
                tmpTestFg = medGetP(Rs.Fields("StatFlags").Value & "", 2, ";") ' �ǹ��� �˻簡�� ����
         
'                tmpStatFg = Mid(tmpStatFg, gBuildingNo, 1)                   '** ���޿���(�ش�ǹ�)
'                Rs.Fields("statfg") = tmpStatFg
'                tmpTestFg = Mid(tmpTestFg, gBuildingNo, 1)                   '** �˻簡�ɿ���(�ش�ǹ�)
'                Rs.Fields("testfg") = tmpTestFg
         
                .AddItem tmpTestNm & tmpTestCd & "1"  '�˻������
                .AddItem tmpTestCd & tmpTestNm & "2"  '�˻��ڵ����
                
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




'% �ٺ� ó�� ����Ʈ�� Click�ϸ� Ordersheet���� �� ������ �����Ѵ�.
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
        .Col = enORDSHEET.tcTESTNM:     Call tblList.GetText(tblList.Col + 1, Row, tmpOrdNm): .Value = tmpOrdNm  ' ó���
        .Col = enORDSHEET.tcTESTCD:     Call tblList.GetText(tblList.Col + 2, Row, tmpOrdCd): .Value = tmpOrdCd ' ó���ڵ�
        
'        Call rs.KeyChange(tmpOrdCd)
        
        Rs.Open objSQL.GetItemInfo(tmpOrdCd), DBConn
        
        .Col = enORDSHEET.tcINSURFG:    .Value = Rs.Fields("insurfg").Value & ""   ' �޿�����
        .Col = enORDSHEET.tcSPCCD:      .Value = Rs.Fields("spccd").Value & ""     ' ��ü�ڵ�
        .Col = enORDSHEET.tcREQDTTM:    .Value = ColDate
        .Col = enORDSHEET.tcSTATFG:     .Value = Rs.Fields("statfg").Value & ""    ' **���޿���(�ش�ǹ�)
    '***�ǹ����� ���
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
        .Col = enORDSHEET.tcSTORECD:    .Value = Rs.Fields("storecd").Value & ""   ' ��������
        .Col = enORDSHEET.tcRNDFG:      .Value = Rs.Fields("rndfg").Value & ""     ' ��ħä�� ����
        .Col = enORDSHEET.tcTESTDIV:    .Value = Rs.Fields("testdiv").Value & ""   ' �׽�Ʈ����(0:�Ϲ�,1:��Ÿ,2:�̻���)
        .Col = enORDSHEET.tcMULTIFG:    .Value = Rs.Fields("multifg").Value & ""   ' ������ü����
        .Col = enORDSHEET.tcSPCGRP:     .Value = Rs.Fields("spcgrp").Value & ""    ' ��ü��
        .Col = enORDSHEET.tcABBRNM:     .Value = Rs.Fields("abbrnm5").Value & ""   ' ����
        .Col = enORDSHEET.tcBARCNT:     .Value = Rs.Fields("labelcnt").Value & ""    ' ��������
        Dim tmpTestFg As String
'        tmpStatFg = medGetP("" & tmpRs.Fields("StatFlags").Value, 1, ";")   '�ǹ��� ���ް��� ����
        tmpTestFg = medGetP("" & Rs.Fields("StatFlags").Value, 2, ";")   '�ǹ��� �˻簡�� ����
        
        .Col = enORDSHEET.tcTESTFLAG:   .Value = tmpTestFg 'Rs.Fields("testfg").Value & ""    ' **�˻簡�ɿ���(�ش�ǹ�)
        
    '***�ǹ����� ���
        If P_ApplyBuildingInfo Then
            If .Value = "1" Then
                .Col = enORDSHEET.tcBUILDCD: .Value = gBuildingCd     ' ** �ش�ǹ����� �Ϲݰ˻� ������
                .Col = enORDSHEET.tcBUILDNM: .Value = gBuildingNm
            Else
                .Col = enORDSHEET.tcBUILDCD: .Value = CentralLab      ' ** �ش�ǹ����� �Ϲݰ˻� �Ұ����� --> �߾Ӱ˻�Ƿ�...
                .Col = enORDSHEET.tcBUILDNM: .Value = CentralLabNm
            End If
        Else
            .Col = enORDSHEET.tcBUILDCD: .Value = gBuildingCd     ' ** �ش�ǹ����� �Ϲݰ˻� ������
            .Col = enORDSHEET.tcBUILDNM: .Value = LABName
        End If
        .Col = enORDSHEET.tcSPCABBR:    .Value = Rs.Fields("spcnm").Value & ""     ' ��ü����
        .Col = enORDSHEET.tcLABDIV:     .Value = Rs.Fields("labdiv").Value & ""    ' ������ȣ �ο�����
        .Col = enORDSHEET.tcLABRANGE:   .Value = Rs.Fields("labrange").Value & ""  ' �̻��� ������ȣ ����
    End With
    
    Set Rs = Nothing
    Set objSQL = Nothing
End Sub
   

'% ó���ڵ带 �������� ������ü ����Ʈ��  ����.
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
            tmpStatFg = medGetP("" & .Fields("StatFlags").Value, 1, ";")    '�ǹ��� ���ް��� ����
            tmpTestFg = medGetP("" & .Fields("StatFlags").Value, 2, ";")    '�ǹ��� �˻簡�� ����
   
            tmpStr = ""
            tmpStr = tmpStr & Trim(.Fields("SpcCd").Value) & vbTab          '��ü�ڵ�
            tmpStr = tmpStr & Trim(.Fields("SpcNm").Value) & vbTab          '��ü��
            'tmpStr = tmpStr & Trim(.Fields("StatFg").Value) & vbTab        '���޿���
            tmpStr = tmpStr & Mid(tmpStatFg, gBuildingNo, 1) & vbTab         '**���޿���(�ش�ǹ�)
            tmpStr = tmpStr & Trim(.Fields("StoreCd").Value) & vbTab        '��������
            tmpStr = tmpStr & Trim(.Fields("MultiFg").Value) & vbTab        '������ü����
            tmpStr = tmpStr & Trim(.Fields("SpcGrp").Value) & vbTab         '��ü��
            tmpStr = tmpStr & Trim(.Fields("LabelCnt").Value) & vbTab       '��������
            tmpStr = tmpStr & Mid(tmpTestFg, gBuildingNo, 1) & vbTab         '**�˻簡�ɿ���(�ش�ǹ�)
            tmpStr = tmpStr & Trim(.Fields("SpcAbbr").Value) & vbTab        '��ü����
            tmpStr = tmpStr & Trim(.Fields("LabDiv").Value) & vbTab         '������ȣ �ο�����
            tmpStr = tmpStr & Trim(.Fields("LabRange").Value) & vbTab       '��ü���� ������ȣ �ο� ����
            tmpStr = tmpStr & Trim(.Fields("Seq").Value)                    '�켱����
            paraSpcList.AddItem tmpStr
            tmpRs.MoveNext
        Wend
    End With
    
NoData:
    Set tmpRs = Nothing
End Sub

'% ó��Header/Body ������ Ŭ������ Assign�Ѵ�.
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
            .Col = enORDSHEET.tcSPCCD:    KeyString = KeyString & .Value & vbTab        '��ü�ڵ�
            .Col = enORDSHEET.tcSTORECD:  KeyString = KeyString & .Value & vbTab        'StoreCd
            .Col = enORDSHEET.tcSTATFG:   KeyString = KeyString & 0 & vbTab 'CStr(Val(.Value)) & vbTab   '���޿���
            .Col = enORDSHEET.tcREQDTTM:  KeyString = KeyString & .Value                '���ä��ð�
         
            'HEADER ----------------------------------------------------------------
            If SaveKeyString <> KeyString Then
                    
                HSeq = OrdHeader.RecordCount + 1    'ó�� Header Seq
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
            BSeq = OrdHeader.Fields("BodyCnt") + 1   'ó��Body Seq
            
            KeyString = HSeq & COL_DIV & BSeq
            
            If Not OrdBody.Exists(KeyString) Then
                OrdBody.AddNew KeyString, ""
            Else
                OrdBody.KeyChange KeyString
            End If
            
            .Col = enORDSHEET.tcORDNO:  .Value = HSeq
            .Col = enORDSHEET.tcTESTCD:  OrdBody.Fields("OrdCd") = .Value    'ó���ڵ�
            .Col = enORDSHEET.tcSPCCD:   OrdBody.Fields("SpcCd") = .Value    '��ü�ڵ�
            .Col = enORDSHEET.tcSTORECD: OrdBody.Fields("StoreCd") = .Value  '��������
            .Col = enORDSHEET.tcINSURFG: OrdBody.Fields("InsDiv") = .Value              '�޿�����
            .Col = enORDSHEET.tcSTATCHK: OrdBody.Fields("StatFg") = CStr(Val(.Value))   '���޿���
            .Col = enORDSHEET.tcORDSEQ: .Value = BSeq
            
            OrdBody.Fields("DcFg") = ""         'DC ����
            OrdBody.Fields("DcDt") = ""         'DC ����
            OrdBody.Fields("DcNo") = "0"          'DC ��ȣ
            OrdBody.Fields("AttrCd") = ""       '�Ӽ�����
            OrdBody.Fields("ExamDt") = ""       '���� �ǽ�����
            OrdBody.Fields("ExamTm") = ""       '���� �ǽýð�
            OrdBody.Fields("ExamDoct") = ""     '���� �ǽ��ǻ�
            OrdBody.Fields("StsCd") = enStsCd.StsCd_LIS_Order       'Status
            OrdBody.Fields("DoneFg") = enStsCd.StsCd_LIS_Order      'Status
            
            OrdHeader.Fields("BodyCnt") = BSeq
        Next
    End With

    OrdHeader.Sort = True
    OrdBody.Sort = True
    
End Sub

'% ó��Header/Body ������ Array Bound�� �������Ѵ�.
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

'% Ŭ������ ó��Header/Body ������ DB�� �����Ѵ�.
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
   
' ���ҽ�
'    For i = 1 To BCnt 'OrdHeader.RecordCount
'        'BCnt = UBound(OrdHeader(I).OrdBody)
'
'        OrdHeader.KeyChange CStr(i)
'        OrdHeader.Fields("OrdNo") = LastOrdNo + i
'
'        sqlStmtH = objOrdSql.CreateSqlHeader(Me, OrdHeader)                'ó�� Header Insert Sql�� ����
'        Call DBConn.Execute(sqlStmtH)       'Sql ����
'
'        BCnt = Val(OrdHeader.Fields("BodyCnt"))
'        For j = 1 To BCnt
'            OrdBody.KeyChange CStr(i) & COL_DIV & CStr(j)
'            OrdBody.Fields("OrdNo") = OrdHeader.Fields("OrdNo")
'            sqlStmtB = objOrdSql.CreateSqlBody(Me, OrdBody)
'            Call DBConn.Execute(sqlStmtB)   'Sql ����
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
        
        sqlStmtH = objOrdSql.CreateSqlHeader(Me, OrdHeader)                'ó�� Header Insert Sql�� ����
        Call DBConn.Execute(sqlStmtH)       'Sql ����
        
'        BCnt = Val(OrdHeader.Fields("BodyCnt"))
'        For j = 1 To BCnt
'            If i = j Then
'                OrdBody.KeyChange CStr(i) & COL_DIV & CStr(i)
'                OrdBody.Fields("OrdNo") = OrdHeader.Fields("OrdNo")
                OrdBody.KeyChange 1 & COL_DIV & i
                
                OrdBody.Fields("OrdNo") = OrdHeader.Fields("OrdNo")
                sqlStmtB = objOrdSql.CreateSqlBody(Me, OrdBody)
                Call DBConn.Execute(sqlStmtB)   'Sql ����
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

'% ���� ������ ó���ȣ�� �����´�.
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