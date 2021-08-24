VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDonorBusiOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'+--------------------------------------------------------------------------------------+
'|  1.  Class   ��  : clsOrders
'|  2.  �� ��         : ó�泻��(HIS101, HIS102)�� �����Ѵ�.
'|  3. �ۼ���        : ��̰�
'|  4. �ۼ���        : 1999.06.01
'|  5. ������        : 2000.12
'|  6. ������        : Legends
'|  7. ���� ����     : BBS�� ������ �˻��Ƿڰ� �ӻ󺴸� �ý����� ó��� �ٸ��� BBS �˻��Ƿڿ�
'|                     �°� ������.
'|  CopyRight(C) 1999 ��ÿ�Ƽ����
'+--------------------------------------------------------------------------------------+


'%  Ŭ���� clsOrder�� Data Attributes

Public PtId As String           ' ȯ��ID
Public orddt As String         ' ó����
Public Bussdiv As String     ' ��������
Public BedIndt As String      ' �Կ���
Public DeptCd As String       ' �����
Public Orddoct As String     ' ó����
Public MajDoct  As String    ' ��ġ��
Public entid  As String        ' �Է���
Public EntDt As String          ' �Է���
Public EntTm  As String        ' �Է½ð�
Public DoneFg As String
Public OrdDiv As String         ' ó�汸��

Public OrgAccNo As String    ' ��������ȣ
Public SpOrdDiv As String     ' Ư��ó�汸��
Public RepeatFg As String
Public Receptno As String     '��������ȣ

'���߿� ����..
Public wardid As String
Public RoomId As String
Public HosilID As String

'= ó�� Body ='
Private Type OrderBody
   OrdSeq As Long     ' ó�� Seq
   OrdCd As String        ' ó���ڵ�
   SpcCd As String        ' ��ü�ڵ�
   StoreCd As String      ' ��������
   DcFg As String          ' DC ����
   DcDt As String           ' DC ����
   DcNo As Long        ' DC ó���ȣ
   Attrcd As String         ' �Ӽ�����
   ExamDt As String       ' ���� �ǽ�����
   ExamTm As String      ' ���� �ǽýð�
   ExamDoct As String   ' ���� �ǽ��ǻ�
   stscd As String        ' Status
   statfg As String        ' ���޿���
   InsDiv As String        ' �޿�����
   DoneFg As String
End Type

'= ó�� Header ='
Private Type OrderHeader
   OrdNo As Long      ' ó���ȣ
   Reqdt As String        ' �ǽÿ�û����
   ReqTm  As String      ' �ǽÿ�û�ð�
   OrdBody() As OrderBody
End Type


'���뺯�� - �����ϴ� ������Ʈ�κ��� ������
'Public MyOraSE As Object     'OraSession

Private OrdHeader() As OrderHeader
Private LastOrdNo As Long
Private sqlStmtH() As String
Private sqlStmtB() As String

Private objMySQL As New clsBBSSQLStatement
'Private DbConn As DrDatabase

Private mvarBuildingNo As String 'BuildingNo�� ���� ���� ���� ������Ƽ.
Private mvarDateTime As Date '���� ��¥�� �ִ´�.
Public Property Let DateTime(ByVal vData As Date)
    mvarDateTime = vData
End Property

Public Property Get DateTime() As Date
    DateTime = mvarDateTime
End Property




Public Property Let BuildingNo(ByVal vData As String)
    mvarBuildingNo = vData
End Property

Public Property Get BuildingNo() As String
    BuildingNo = mvarBuildingNo
End Property

'Public Sub setDbConn(ByRef pDbConn As DrDatabase)
''��� Ŀ�ؼ�
'
'    Set DbConn = pDbConn
'End Sub

'% ó���ڵ带 �������� ������ü ����Ʈ��  ����.
Public Sub SpcList(ByVal ParaTestCd As String, ByRef paraSpcList As Object)
   
   Dim SqlStmt As String
   Dim tmpRs As Recordset
   Dim tmpStr As String
   Dim tmpStatFg As String
   Dim tmpTestFg As String
   
   SqlStmt = objMySQL.GetSpecList(ParaTestCd)
   
   Set tmpRs = New Recordset
   tmpRs.Open SqlStmt, DBConn
   If tmpRs.EOF Then GoTo NoData:
   'tmpRs.MoveFirst
   paraSpcList.Clear
   
   With tmpRs
      While (Not tmpRs.EOF)
         tmpStatFg = medGetP("" & .Fields("StatFlags").Value, 1, ";")   '�ǹ��� ���ް��� ����
         tmpTestFg = medGetP("" & .Fields("StatFlags").Value, 2, ";")  '�ǹ��� �˻簡�� ����
   
         tmpStr = ""
         tmpStr = tmpStr & Trim(.Fields("SpcCd").Value & "") & Chr(9)   '��ü�ڵ�
         tmpStr = tmpStr & Trim(.Fields("SpcNm").Value & "") & Chr(9)   '��ü��
         'tmpStr = tmpStr & Trim(.Fields("StatFg").value & "") & Chr(9)    '���޿���
         tmpStr = tmpStr & Mid(tmpStatFg, BuildingNo, 1) & Chr(9)   '**���޿���(�ش�ǹ�)
         tmpStr = tmpStr & Trim(.Fields("StoreCd").Value & "") & Chr(9)  '��������
         tmpStr = tmpStr & Trim(.Fields("MultiFg").Value & "") & Chr(9)   '������ü����
         tmpStr = tmpStr & Trim(.Fields("SpcGrp").Value & "") & Chr(9)   '��ü��
         tmpStr = tmpStr & Trim(.Fields("LabelCnt").Value & "") & Chr(9) '��������
         tmpStr = tmpStr & Mid(tmpTestFg, BuildingNo, 1) & Chr(9)   '**�˻簡�ɿ���(�ش�ǹ�)
         tmpStr = tmpStr & Trim(.Fields("Seq").Value & "")                     '�켱����
         paraSpcList.AddItem tmpStr
         tmpRs.MoveNext
      Wend
   End With
NoData:
'   tmpRs.RsClose
   Set tmpRs = Nothing

End Sub

'% ó��Header/Body ������ Ŭ������ Assign�Ѵ�.
Public Sub MoveData(ByRef tblOrdSheet As Object)
   
   Dim i As Long
   Dim HSeq As Long
   Dim BSeq As Long
   Dim SaveKeyString As String
   Dim KeyString As String
   
   SaveKeyString = ""
   With tblOrdSheet
      For i = 1 To .DataRowCnt
         .Row = i
         .Col = 19: KeyString = .Value & Chr(9)                    '** Delivery Location
         '.Col = 12: KeyString = KeyString & .Value & Chr(9)   '�˻籸��  --> ���� 1999.10.08 by ��̰�
         .Col = 9:  KeyString = KeyString & .Value & Chr(9)   'WorkArea
         .Col = 4:  KeyString = KeyString & .Value & Chr(9)   '��ü�ڵ�
         .Col = 10: KeyString = KeyString & .Value & Chr(9)   'StoreCd
         .Col = 6:  KeyString = KeyString & CStr(Val(.Value)) & Chr(9)   '���޿���
         .Col = 7:  KeyString = KeyString & .Value                '���ä��ð�
         
         If SaveKeyString <> KeyString Then
            HSeq = NextSeq(1)
            SaveKeyString = KeyString
            .Col = 7: OrdHeader(HSeq).Reqdt = Format(.Value, CS_DateDbFormat)
                         OrdHeader(HSeq).ReqTm = Format(.Value, CS_TimeDbFormat)
         End If
         
         'OrdHeader(HSeq).OrdNo = HSeq    'ó��Header Seq
         .Col = 1: .Value = HSeq
         BSeq = NextSeq(2, HSeq)    'ó��Body Seq
         .Col = 3: OrdHeader(HSeq).OrdBody(BSeq).OrdCd = .Value   'ó���ڵ�
         .Col = 4: OrdHeader(HSeq).OrdBody(BSeq).SpcCd = .Value   '��ü�ڵ�
         .Col = 10: OrdHeader(HSeq).OrdBody(BSeq).StoreCd = .Value  '��������
                      OrdHeader(HSeq).OrdBody(BSeq).DcFg = ""           'DC ����
                      OrdHeader(HSeq).OrdBody(BSeq).DcDt = ""            'DC ����
                      OrdHeader(HSeq).OrdBody(BSeq).DcNo = 0            'DC ��ȣ
                      OrdHeader(HSeq).OrdBody(BSeq).Attrcd = ""          '�Ӽ�����
                      OrdHeader(HSeq).OrdBody(BSeq).ExamDt = ""        '���� �ǽ�����
                      OrdHeader(HSeq).OrdBody(BSeq).ExamTm = ""       '���� �ǽýð�
                      OrdHeader(HSeq).OrdBody(BSeq).ExamDoct = ""      '���� �ǽ��ǻ�
                      OrdHeader(HSeq).OrdBody(BSeq).stscd = "0"            'Status
                      OrdHeader(HSeq).OrdBody(BSeq).DoneFg = "0"          'Status
         .Col = 5: OrdHeader(HSeq).OrdBody(BSeq).InsDiv = .Value                   '�޿�����
         .Col = 6: OrdHeader(HSeq).OrdBody(BSeq).statfg = CStr(Val(.Value))    '���޿���
         .Col = 15: .Value = BSeq
      Next
   End With

End Sub

Public Function ExecuteSqlStmt() As Boolean

    Dim i As Long

On Error GoTo ErrExecute

    For i = 1 To UBound(sqlStmtH)
        'Debug.Print "Header i = " & i & " : " & sqlStmtH(i)
        DBConn.Execute sqlStmtH(i)
    Next
    
    For i = 1 To UBound(sqlStmtB)
       ' Debug.Print "Body i = " & i & " : " & sqlStmtB(i)
        DBConn.Execute sqlStmtB(i)
    Next
        
    ExecuteSqlStmt = True
    Exit Function
        
ErrExecute:
    ExecuteSqlStmt = False
End Function


'% ó��Header/Body ������ Array Bound�� �������Ѵ�.
Private Function NextSeq(ByVal TableFg As Long, Optional ByVal index As Variant)
   Dim lngBnd As Long
   Select Case TableFg
      Case 1:  'Header
         lngBnd = UBound(OrdHeader)
         ReDim Preserve OrdHeader(lngBnd + 1)
         ReDim Preserve OrdHeader(lngBnd + 1).OrdBody(0)
      Case 2:  'Body
         lngBnd = UBound(OrdHeader(index).OrdBody)
         ReDim Preserve OrdHeader(index).OrdBody(lngBnd + 1)
   End Select
   NextSeq = lngBnd + 1
End Function

'% Ŭ������ ó��Header/Body ������ DB�� �����Ѵ�.
Public Function CreateSqlStmt(ByRef StartOrdNo As Long, Optional ByRef ProgressBar As Variant) As Boolean
   
    Dim HCnt As Long, BCnt As Long
    Dim i As Long, j As Long
    
      
    LastOrdNo = GetLastNo(PtId, orddt)
    HCnt = UBound(OrdHeader)
    
    
     ReDim sqlStmtH(0)
     ReDim sqlStmtB(0)
    
On Error GoTo ErrCreateSQLStmt
    For i = 1 To HCnt
       BCnt = UBound(OrdHeader(i).OrdBody)
       
       ReDim Preserve sqlStmtH(i)
       
       OrdHeader(i).OrdNo = LastOrdNo + i
       sqlStmtH(i) = CreateSqlH(i)   'ó�� Header Insert Sql�� ����
       For j = 1 To BCnt
'            Debug.Print "Ubound(sqlstmtb) : " & UBound(sqlStmtB)
          ReDim Preserve sqlStmtB(UBound(sqlStmtB) + 1)
          
          sqlStmtB(UBound(sqlStmtB)) = CreateSqlB(i, j)  'ó�� Body Insert Sql�� ����
'          Debug.Print "UBound(sqlStmtB) = " & UBound(sqlStmtB) & " : " & "sqlStmtB(" & UBound(sqlStmtB) & ")" & " = " & sqlStmtB(UBound(sqlStmtB))
          
       Next
    Next
    StartOrdNo = LastOrdNo
    CreateSqlStmt = True
    Exit Function
   
ErrCreateSQLStmt:
    CreateSqlStmt = False
End Function

'% ���� ������ ó���ȣ�� �����´�.
Public Function GetLastNo(ByVal paraPtId As String, ByVal paraOrdDt As String)
   
   Dim tmpDs As Recordset
   Dim SqlStmt As String
   
   SqlStmt = objMySQL.GetLastOrdNo(paraPtId, paraOrdDt)
   Set tmpDs = New Recordset
   tmpDs.Open SqlStmt, DBConn
   If tmpDs.EOF Then
      GetLastNo = 0
   Else
      GetLastNo = tmpDs.Fields("ordno").Value & ""
   End If
   
   Set tmpDs = Nothing
End Function

Public Function CheckSameOrder(ByVal OrdTable As Object) As Long

   Dim i As Long, j As Long
   Dim SaveCode As String
   Dim SaveSpc As String
   Dim SaveDate As String
   
   CheckSameOrder = 0
   With OrdTable
      For i = 1 To .DataRowCnt
         .Row = i
         .Col = 3:  SaveCode = .Value
         .Col = 7:  SaveDate = .Value
         .Col = 4:  SaveSpc = .Value
         For j = i + 1 To .DataRowCnt
            .Row = j
            .Col = 3
            If .Value = SaveCode Then
               .Col = 7
               If .Value = SaveDate Then
                    .Col = 4
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

Private Sub Class_Initialize()
   Erase OrdHeader
   ReDim OrdHeader(0)
   ReDim OrdHeader(0).OrdBody(0)
End Sub

Private Function CreateSqlH(ByVal HSeq As Long) As String

   CreateSqlH = "Insert into " & _
                     "" & T_LAB101 & " (ptid, orddt, ordtm, ordno, bussdiv, bedindt, reqdt, " & _
                     "             reqtm, deptcd, orddoct, majdoct, entid, entdt, " & _
                     "             enttm, orddiv, ordfg, repeatfg, orgaccno, sporddiv, donefg, receptno, wardid, roomid, hosilid) " & _
                     "Values  (" & _
                        DBV("ptid", PtId, 1) & DBV("orddt", orddt, 1) & DBV("ordtm", Format(mvarDateTime, PRESENTTIME_FORMAT), 1) & DBV("ordno", OrdHeader(HSeq).OrdNo, 1) _
                        & DBV("bussdiv", Bussdiv, 1) & DBV("bedindt", BedIndt, 1) & DBV("reqdt", OrdHeader(HSeq).Reqdt, 1) _
                        & DBV("reqtm", OrdHeader(HSeq).ReqTm, 1) & DBV("deptcd", DeptCd, 1) & DBV("orddoct", Orddoct, 1) & DBV("majdoct", MajDoct, 1) & DBV("entid", entid, 1) _
                        & DBV("entdt", Format(mvarDateTime, PRESENTDATE_FORMAT), 1) & DBV("enttm", Format(mvarDateTime, PRESENTTIME_FORMAT), 1) & DBV("orddiv", OrdDiv, 1) & DBV("ordfg", "1", 1) & DBV("repeatfg", "", 1) & DBV("orgaccno", OrgAccNo, 1) _
                        & DBV("sporddiv", SpOrdDiv, 1) & DBV("donefg", DoneFg, 1) & DBV("receptno", Receptno, 1) & DBV("wardid", wardid, 1) & DBV("roomid", RoomId, 1) & DBV("hosilid", HosilID) & " ) "
End Function

Private Function CreateSqlB(ByVal HSeq As Long, ByVal BSeq As Long) As String
   With OrdHeader(HSeq).OrdBody(BSeq)
      CreateSqlB = "Insert into " & _
                         "" & T_LAB102 & " (ptid, orddt, ordno, ordseq, ordcd, spccd, storecd, dcfg, dcdt, dcno, " & _
                         "             attrcd, examdt, examtm, examdoct, stscd, statfg, insdiv,paydt, donefg) " & _
                         "Values  (" & _
                                DBV("ptid", PtId, 1) & DBV("orddt", orddt, 1) & DBV("ordno", OrdHeader(HSeq).OrdNo, 1) & DBV("ordseq", BSeq, 1) & DBV("ordcd", .OrdCd, 1) & DBV("spccd", .SpcCd, 1) _
                                & DBV("storecd", .StoreCd, 1) & DBV("dcfg", .DcFg, 1) & DBV("dcdt", .DcDt, 1) & DBV("dcno", .DcNo, 1) & DBV("attrcd", .Attrcd, 1) & DBV("examdt", .ExamDt, 1) _
                                & DBV("examtm", .ExamTm, 1) & DBV("examdoct", .ExamDoct, 1) & DBV("stscd", .stscd, 1) & DBV("statfg", .statfg, 1) & DBV("insdiv", .InsDiv, 1) & DBV("paydt", "2", 1) & DBV("donefg", .DoneFg) & " ) "
   End With
End Function

Private Sub Class_Terminate()
   
   Set objMySQL = Nothing
End Sub