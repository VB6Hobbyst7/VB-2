VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResultReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+--------------------------------------------------------------------------------------+
'|  1. Class ��  : clsResultReview
'|  2. ��    ��  : ��������� ��ȸ�Ѵ�.
'|  3. �� �� ��  : ��̰�
'|
'|  CopyRight(C) 1999 ��ÿ�Ƽ����
'+--------------------------------------------------------------------------------------+


'%  Ŭ���� clsResultView�� Data Attributes

Public PtId As String
Public Sex As String
Public AgeDay As Long
Public ColId As String
Public RcvId As String
Public VfyId As String
Public ColDtTm As String
Public RcvDtTm As String
Public VfyDtTm As String
Public MultiFg As String
Public FootNoteFg As String
Public TestDiv As String
Public RemarkCd As String
Public StatFg As String

Public DeptCd As String
Public DeptNm As String
Public WardId As String
Public WardNm As String
Public RoomId As String
Public BedId As String
Public HosilId As String
Public BedinDt As String

Private Type ResultTable
   TestCd As String  '/* �˻��׸� �ڵ�   */
   RstCd As String  '/* ����ڵ�(Alpha) */
   RstUnit As String  '/* Unit            */
   RstDiv As String
   HLDiv As String  '/* High/Low(H:Hig,L:Low) */
   DPDiv As String  '/* Delta/Panic(D:Delta,P:Panic) */
   SpcCd As String  '/* ��ü�ڵ�        */
   StatFg As String  '/* ���޿���('0':��,'1':��) */
   LastRst As String '/* �ֱٰ��        */
   LastVfyDtTm As String '/* �ֱٰ��Ȯ����  */
   'LastVfyTm As String '/* �ֱٰ��Ȯ�νð� */
   LastVfyId As String  '/* �ֱٰ��Ȯ����  */
   VfyDt As String  '/* ���Ȯ������    */
   VfyTm As String  '/* ���Ȯ�νð�    */
   VfyId As String  '/* ���Ȯ����      */
   AttrCd As String  '/* �Ӽ��ڵ�        */
   MfyFg As String  '/* ��������('0':��,'1':��) */
   GrpFg As String  '/* �׷��� �������('0':��,'1':��) */
   TxtFg As String  '/* TEXT �������('0':��,'1':��) */
   ValFg As String  '/* Valeu �������('0':��,'1':��) - ��Ÿ�˻翡���� ��� */
   RstType As String  '/* �������(N,Alpha,����,Free) */
   DetailFg As String  '/* ���׸� SEQ  */
   EqpCd As String  '/* ����ڵ� */
   PtId As String  '/* ȯ�� ID  */
   OrdDt As String  '/* ó���� */
   OrdNo As Integer  '/* ó���ȣ */
   OrdSeq As String  '/* ó��Seq */
   
   RstText As String  '/* Text ��� */
   SuppText As String  '/* Supplemental Report  */
   
   RefFromVal As Double  '/* ����ġ From */
   RefToVal As Double  '/* ����ġ To */
   RefCd As String  '/* ����ġ Code */
   
   TestShortNm As String  '/* �˻��(Short)  */
   TestLongNm As String  '/* �˻��(Long)  */
   RstCdNm As String '/* ����ڵ� �� */
   
   SenFg As String  '/* ��������� ���� */
End Type

'���Ӱ˻� ����
Private Type MultiSpc
   WorkArea As String
   AccDt As String
   AccSeq As String
   RemarkNm As String  '/* ��ü Remark */
   FootNote As String   '/* Foot Note */
End Type

'��������� ����
Private Type SenResult
   RstCd As String
   Row As String
   ForeColor As Long
End Type

Public Remark As String  '/* ��ü Remark */
Public FootNote As String   '/* Foot Note */
Public TextFg As Boolean
Public CommentFg As Boolean

'���뺯�� - �����ϴ� ������Ʈ�κ��� ������
'Public MyOraSE As Object     'OraSession
Public MyDb As Object     'OraDatabase

Public ResultCnt As Integer
Public SortStartRow As Long
Public SortEndRow As Long
Public SortFg As Boolean
Public OffSet As Integer

Public RstRow As Integer
Private MyResult() As ResultTable
Private MySenResult() As SenResult
Private RstForeColor() As Long

Public ResultClipText As String
Public SenClipText As String
Public RstTextBuffer As String
Public SamTextBuffer As String


Private MySql As New clsLISSqlReview
Private MyCollect As New clsLISCollectioin   '��������

Private Const MaxAntiCnt = 50
Private Const MaxColumns = 8


'% Method 1: Lab No�� �������� �������� �� ������� �˻�
Public Sub ResultQuery(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer)

   Call InitRtn  '�����ʱ�ȭ
   
   With MyCollect
      
      'Set .MyOraSE = MyOraSE
      Set .MyDb = MyDb
      
      'ä���������� Query
      
      'Debug.Print "Collect :", ",", Now
      
      Call .CollectQuery(pWorkArea, pAccDt, pAccSeq)

      PtId = .PtId
      Sex = .Sex
      AgeDay = .AgeDay
      ColId = .ColId
      RcvId = .RcvId
      VfyId = .VfyId
      ColDtTm = Format(.ColDt, CS_DateMask) & " " & Format(Mid(.ColTm, 1, 4), CS_TimeSMask)
      RcvDtTm = Format(.RcvDt, CS_DateMask) & " " & Format(Mid(.RcvTm, 1, 4), CS_TimeSMask)
      VfyDtTm = Format(.VfyDt, CS_DateMask) & " " & Format(Mid(.VfyTm, 1, 4), CS_TimeSMask)
      MultiFg = .MultiFg
      FootNoteFg = .FootNoteFg
      TestDiv = .TestDiv
      RemarkCd = .RmkCd
      StatFg = .StatFg
      
      DeptCd = .DeptCd
      DeptNm = GetDeptNm(DeptCd)
      WardId = .WardId
      WardNm = GetDeptNm(WardId)
      RoomId = .RoomId
      HosilId = .HosilId
      BedId = .BedId
      BedinDt = Format(.BedinDt, CS_DateMask)
      
      'ǲ��Ʈ & ��ü����ũ
      If FootNoteFg > "0" Then Call ReadFootNote(pWorkArea, pAccDt, pAccSeq)
      If Trim(RemarkCd) <> "" Then Call ReadRemark(.RmkCd)
      SamTextBuffer = Remark & FootNote
      
      '������� Query
      'Select Case TestDiv
      'Case "0", "1":  '�Ϲݰ˻���,'��Ÿ�˻���
         If Trim(.MultiFg) <> "" Then
            Call MultiResult(pWorkArea, pAccDt, pAccSeq, Sex, AgeDay)
         Else
            Call GeneralResult(pWorkArea, pAccDt, pAccSeq, Sex, AgeDay)
         End If
         Call SpecialResult(pWorkArea, pAccDt, pAccSeq)
      'Case "2":   '�̻����˻���
         Call MicrobeResult(pWorkArea, pAccDt, pAccSeq)
      'End Select
   
      'Debug.Print "Result End :", ",", Now

   End With
   
End Sub


'% Method 1-1 : Lab No�� �������� ����������� �Ұ߳����� �˻�
Public Sub ResultMore(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, _
                      ByVal pTestDiv As String, Optional ByVal pQuery As Boolean = True)

   Call InitRtn  '�����ʱ�ȭ
   
   With MyCollect
      
      'Set .MyOraSE = MyOraSE
      Set .MyDb = MyDb
      
      'ä���������� Query
      
      'Debug.Print "Collect :", ",", Now
      
      Call .CollectQuery(pWorkArea, pAccDt, pAccSeq)

      PtId = .PtId
      Sex = .Sex
      AgeDay = .AgeDay
      ColId = .ColId
      RcvId = .RcvId
      VfyId = .VfyId
      ColDtTm = Format(.ColDt, CS_DateMask) & " " & Format(Mid(.ColTm, 1, 4), CS_TimeSMask)
      RcvDtTm = Format(.RcvDt, CS_DateMask) & " " & Format(Mid(.RcvTm, 1, 4), CS_TimeSMask)
      VfyDtTm = Format(.VfyDt, CS_DateMask) & " " & Format(Mid(.VfyTm, 1, 4), CS_TimeSMask)
      MultiFg = .MultiFg
      FootNoteFg = .FootNoteFg
      TestDiv = .TestDiv
      RemarkCd = .RmkCd
      StatFg = .StatFg
      
      DeptCd = .DeptCd
      DeptNm = GetDeptNm(DeptCd)
      WardId = .WardId
      WardNm = GetDeptNm(WardId)
      RoomId = .RoomId
      HosilId = .HosilId
      BedId = .BedId
      BedinDt = Format(.BedinDt, CS_DateMask)
      
      If Not pQuery Then Exit Sub
      
      'ǲ��Ʈ & ��ü����ũ
      If FootNoteFg > "0" Then Call ReadFootNote(pWorkArea, pAccDt, pAccSeq)
      If Trim(RemarkCd) <> "" Then Call ReadRemark(.RmkCd)
      SamTextBuffer = Remark & FootNote
      
      '������� Query
      Select Case TestDiv
      Case "0", "1":  '�Ϲݰ˻���,'��Ÿ�˻���
         Call GeneralMore(pWorkArea, pAccDt, pAccSeq, Sex, AgeDay)
         Call SpecialMore(pWorkArea, pAccDt, pAccSeq)
      Case "2":   '�̻����˻���
         Call MicrobeMore(pWorkArea, pAccDt, pAccSeq)
      End Select
   
      'Debug.Print "Result End :", ",", Now

   End With
   
End Sub

'% ������ü�� ��� : ó���ȣ/Seq�� �������� ���� ��ü���� ����� ��� �����´�.

Public Sub MultiResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, ByVal pSex As String, ByVal pAgeDay As Long)

   Dim SqlStmt As String
   Dim tmpRs As Object
   Dim I As Integer
   Dim tmpWorkArea As String, tmpAccDt As String, tmpAccSeq As Integer
   
   SqlStmt = MySql.SqlMultiTest(pWorkArea, pAccDt, pAccSeq)
   Set tmpRs = OpenRecordSet(SqlStmt)
   
   With tmpRs
      If .EOF Then GoTo NoData
      While (Not .EOF)
         tmpWorkArea = "" & .Fields("WorkArea").Value
         tmpAccDt = "" & .Fields("AccDt").Value
         tmpAccSeq = Val("" & .Fields("AccSeq").Value)
         Call GeneralResult(tmpWorkArea, tmpAccDt, tmpAccSeq, pSex, pAgeDay)
         .MoveNext
      Wend
   End With
         
NoData:
   tmpRs.RsClose
   Set tmpRs = Nothing
   
End Sub


'% �Ϲݰ˻� ( LAB302 )  : Lab No�� �������� ������� �˻�
Public Sub GeneralResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, _
                                     ByVal pSex As String, ByVal pAgeDay As Long)
 
   Dim I As Integer
   Dim SqlStmt As String
   Dim ColCnt As Integer
   Dim tmpRs As New DrRecordSet
   Dim tmpRs1 As Object
   
   '������� �˻�
   SqlStmt = MySql.SqlQueryResults(pWorkArea, pAccDt, pAccSeq, "0", pSex, pAgeDay)
   'ColCnt = tmpRs.OpenCursor(DbConn, SqlStmt)
   Set tmpRs = OpenRecordSet(SqlStmt)
   
   If tmpRs.EOF Then GoTo NoData
   'If ColCnt = 0 Then GoTo NoData
   
   ResultCnt = tmpRs.RecordCount
   
   '������� Array
   ReDim Preserve MyResult(ResultCnt)
   ReDim Preserve RstForeColor(MaxColumns, ResultCnt)
   
   For I = 1 To ResultCnt
   'ResultCnt = 0
   'While (tmpRs.FetchCursor(ColCnt))
      
      'ResultCnt = ResultCnt + 1
      'ReDim Preserve MyResult(ResultCnt)
      'ReDim Preserve RstForeColor(MaxColumns, ResultCnt)
      
      Call AddRow(1)
      With MyResult(I)
         .TestCd = Trim("" & tmpRs.Fields("TestCd").Value)
         .RstCd = Trim("" & tmpRs.Fields("RstCd").Value)
         .RstUnit = Trim("" & tmpRs.Fields("RstUnit").Value)
         .HLDiv = Trim("" & tmpRs.Fields("HLDiv").Value)
         .DPDiv = Trim("" & tmpRs.Fields("DPDiv").Value)
         .SpcCd = Trim("" & tmpRs.Fields("SpcCd").Value)
         .StatFg = Trim("" & tmpRs.Fields("StatFg").Value)
         .RstDiv = Trim("" & tmpRs.Fields("RstDiv").Value)
         .LastRst = Trim("" & tmpRs.Fields("LastRst").Value)
         .LastVfyDtTm = Trim("" & tmpRs.Fields("LstVfyDtTm").Value)
         .LastVfyId = Trim("" & tmpRs.Fields("LastVfyNm").Value)
         .VfyDt = Trim("" & tmpRs.Fields("VfyDt").Value)
         .VfyTm = Trim("" & tmpRs.Fields("VfyTm").Value)
         .VfyId = Trim("" & tmpRs.Fields("VfyNm").Value)
         .AttrCd = Trim("" & tmpRs.Fields("AttrCd").Value)
         .MfyFg = Trim("" & tmpRs.Fields("MfyFg").Value)
         .GrpFg = Trim("" & tmpRs.Fields("GrpFg").Value)
         .TxtFg = Trim("" & tmpRs.Fields("TxtFg").Value)
         .RstType = Trim("" & tmpRs.Fields("RstType").Value)
         .DetailFg = Trim("" & tmpRs.Fields("DetailFg").Value)
         .EqpCd = Trim("" & tmpRs.Fields("EqpCd").Value)
         .PtId = Trim("" & tmpRs.Fields("PtId").Value)
         .OrdDt = Trim("" & tmpRs.Fields("OrdDt").Value)
         .OrdNo = Val("" & tmpRs.Fields("OrdNo").Value)
         .OrdSeq = Val("" & tmpRs.Fields("OrdSeq").Value)

         .RstText = Trim("" & tmpRs.Fields("TextResult").Value)
         'If .TxtFg = "1" Then .RstText = GetRstText(pWorkArea, pAccDt, pAccSeq, .TestCd)
         
         'Supplemental Report �� �ִ� ���...
         If .TxtFg = "2" Then .SuppText = GetSuppText(pWorkArea, pAccDt, pAccSeq, .TestCd)
         
         '�˻�� (�Ϲ��׸�/���׸�)
         If .DetailFg = "" Or .RstDiv = "*" Then
            .TestShortNm = Trim("" & tmpRs.Fields("TestShortNm").Value)
         Else
            .TestShortNm = "   " & Trim("" & tmpRs.Fields("TestShortNm").Value)
         End If
         .TestLongNm = Trim("" & tmpRs.Fields("TestLongNm").Value)  '�˻�� Full Name
         .RstCdNm = Trim("" & tmpRs.Fields("RstCdNm").Value)           '����ڵ��

         '�����(�ڵ��� ���..)
         If .VfyDt = "" Then
            If .RstDiv <> "*" Then .RstCd = "��Ȯ": RstForeColor(3, RstRow) = &HC0C0FF
            .HLDiv = ""
         Else
            RstForeColor(3, RstRow) = &H404080   '����
            If .RstCdNm <> "" Then .RstCd = .RstCdNm
            'High / Low
            If .RstCd <> "" Then
               If .HLDiv = "H" Then .HLDiv = "��": RstForeColor(5, RstRow) = &H7477EF '�ణ ������
               If .HLDiv = "L" Then .HLDiv = "��": RstForeColor(5, RstRow) = &HE48372 '�ణ �Ķ���
            End If
         End If
         
         
         '����ġ �˻�
         SqlStmt = MySql.SqlGetReference(.TestCd, .SpcCd, .VfyDt, "B", AgeDay)
         Set tmpRs1 = OpenRecordSet(SqlStmt)
         If tmpRs1.EOF Then  'ȯ�ڼ����� �ش��ϴ� ����ġ�� ���� ��� "B"(Both)�� �ش��ϴ� ����Ÿ �˻�
            tmpRs1.RsClose
            SqlStmt = MySql.SqlGetReference(.TestCd, .SpcCd, .VfyDt, Sex, AgeDay)
            Set tmpRs1 = OpenRecordSet(SqlStmt)
         End If
         If tmpRs1.EOF Then
            .RefFromVal = 0: .RefToVal = 0: .RefCd = ""
         Else
            .RefFromVal = Val("" & tmpRs1.Fields("RefValFrom").Value)
            .RefToVal = Val("" & tmpRs1.Fields("RefValTo").Value)
            .RefCd = Trim("" & tmpRs1.Fields("RefCd").Value)
            If .RefFromVal <> 0 Or .RefToVal <> 0 Then .RefCd = .RefFromVal & " - " & .RefToVal
         End If
         tmpRs1.RsClose
         '***************************************************************************
         '.RefFromVal = Val("" & tmpRs.Fields("RefValFrom").Value)
         '.RefToVal = Val("" & tmpRs.Fields("RefValTo").Value)
         '.RefCd = Trim("" & tmpRs.Fields("RefCd").Value)
         'If .RefFromVal <> 0 Or .RefToVal <> 0 Then .RefCd = .RefFromVal & " - " & .RefToVal
         '***************************************************************************
         
      
         RstForeColor(2, RstRow) = &H747474 '���� ȸ��
         'RstForeColor(3, RstRow) = &H404080 '����
         ResultClipText = ResultClipText & OneRow(I)       '������� Buffering
         RstTextBuffer = RstTextBuffer & OneTextResult(I)  '�ؽ�Ʈ��� Buffering
         
      End With
         
      tmpRs.MoveNext
      
   'Wend
   Next
   
NoData:
   tmpRs.RsClose
   Set tmpRs = Nothing
   Set tmpRs1 = Nothing
   
End Sub

'% �Ϲݰ˻� ( LAB302 )  : Text ��� �� Supplemental Report �� �˻�...
Public Sub GeneralMore(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, _
                                     ByVal pSex As String, ByVal pAgeDay As Long)
 
   Dim I As Integer
   Dim SqlStmt As String
   Dim ColCnt As Integer
   Dim tmpRs As New DrRecordSet
   Dim tmpRs1 As Object
   
   '������� �˻�
   SqlStmt = MySql.SqlQueryResults(pWorkArea, pAccDt, pAccSeq, "0", pSex, pAgeDay)
   Set tmpRs = OpenRecordSet(SqlStmt)
   
   If tmpRs.EOF Then GoTo NoData
   
   ResultCnt = tmpRs.RecordCount
   
   '������� Array
   ReDim Preserve MyResult(ResultCnt)
   ReDim Preserve RstForeColor(MaxColumns, ResultCnt)
   
   For I = 1 To ResultCnt
      
      Call AddRow(1)
      
      With MyResult(I)
         
         .TestCd = Trim("" & tmpRs.Fields("TestCd").Value)
         .TestShortNm = Trim("" & tmpRs.Fields("TestShortNm").Value)
         .RstText = Trim("" & tmpRs.Fields("TextResult").Value)
         
         'Supplemental Report �� �ִ� ���...
         If Trim("" & tmpRs.Fields("TxtFg").Value) = "2" Then .SuppText = GetSuppText(pWorkArea, pAccDt, pAccSeq, .TestCd)
         
         '�ؽ�Ʈ��� Buffering
         If .RstText <> "" Then
            ResultClipText = ResultClipText & "<< �˻� �Ұ� >> - " & .TestShortNm & vbCrLf
            ResultClipText = ResultClipText & .RstText & vbCrLf & vbCrLf
            TextFg = True
         End If
         If .SuppText <> "" Then
            ResultClipText = ResultClipText & "<< Supplemental Report >> " & vbCrLf
            ResultClipText = ResultClipText & .SuppText & vbCrLf
            TextFg = True
         End If
         
      End With
         
      tmpRs.MoveNext
      
   'Wend
   Next
   
NoData:
   tmpRs.RsClose
   Set tmpRs = Nothing
   Set tmpRs1 = Nothing
   
End Sub

      
'% ��Ÿ�˻� ( LAB351, LAB352, LAB353 ) : Lab No�� �������� ������� �˻�
Public Sub SpecialResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer)
 
   Dim I As Integer
   Dim j As Integer
   Dim SqlStmt As String
   Dim ColCnt As Integer
   Dim SvTestCd As String
   Dim tmpRs As New DrRecordSet
   
   '������� �˻�
   SqlStmt = MySql.SqlQueryResults(pWorkArea, pAccDt, pAccSeq, "1")
   Set tmpRs = OpenRecordSet(SqlStmt)
   
   If tmpRs.EOF Then GoTo NoData
   
   ResultCnt = tmpRs.RecordCount
   
   '������� Array
   ReDim Preserve MyResult(ResultCnt)
   ReDim Preserve RstForeColor(MaxColumns, ResultCnt)
   
   For I = 1 To ResultCnt
      
      Call AddRow(1)
      With MyResult(I)
         .TestCd = Trim("" & tmpRs.Fields("TestCd").Value)
         .ValFg = Trim("" & tmpRs.Fields("ValFg").Value)
         .TxtFg = Trim("" & tmpRs.Fields("TxtFg").Value)
         .MfyFg = Trim("" & tmpRs.Fields("MfySeq").Value)
         .RstType = Trim("" & tmpRs.Fields("RstType").Value)
         .VfyDt = Trim("" & tmpRs.Fields("VfyDt").Value)
         .VfyTm = Trim("" & tmpRs.Fields("VfyTm").Value)
         .VfyId = Trim("" & tmpRs.Fields("VfyNm").Value)
         .PtId = Trim("" & tmpRs.Fields("PtId").Value)
         .OrdDt = Trim("" & tmpRs.Fields("OrdDt").Value)
         .OrdNo = Trim("" & tmpRs.Fields("OrdNo").Value)
         .OrdSeq = Trim("" & tmpRs.Fields("OrdSeq").Value)

         '�ؽ�Ʈ ���
         If .TxtFg = ERT_TxtRst Then
            RstTextBuffer = RstTextBuffer & "<< �˻� �Ұ� >>  - " & tmpRs.Fields("TestShortNm").Value & vbCrLf
            RstTextBuffer = RstTextBuffer & tmpRs.Fields("TextResult").Value & vbCrLf & vbCrLf
            'Supplemental Report
            Dim SuppRs As Object
            SqlStmt = MySql.SqlETextRst(pWorkArea, pAccDt, pAccSeq, .TestCd)
            Set SuppRs = OpenRecordSet(SqlStmt)
            If Not SuppRs.EOF Then
               RstTextBuffer = RstTextBuffer & "<< Supplemental Report >> " & vbCrLf
               While (Not SuppRs.EOF)
                  RstTextBuffer = RstTextBuffer & SuppRs.Fields("TextResult").Value & vbCrLf
                  SuppRs.MoveNext
               Wend
               SuppRs.RsClose
               Set SuppRs = Nothing
            End If
            TextFg = True
         End If
                  
         '�˻�� (�Ϲ��׸�/���׸�)
         .TestShortNm = Trim("" & tmpRs.Fields("TestShortNm").Value)
         .TestLongNm = Trim("" & tmpRs.Fields("TestLongNm").Value)  '�˻�� Full Name
         
         ResultClipText = ResultClipText & .TestLongNm & vbCrLf
         
         '�� ��ġ��� �˻�...
         If .ValFg = ERT_ValRst Then Call GetRstValues(pWorkArea, pAccDt, pAccSeq, .TestCd, .MfyFg)
         
      End With
      tmpRs.MoveNext
      
   Next

NoData:
    tmpRs.RsClose
    Set tmpRs = Nothing
    
End Sub
         
'% ��Ÿ�˻� ( LAB351, LAB352, LAB353 ) : Text�����...
Public Sub SpecialMore(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer)
 
    Dim I As Integer
    Dim j As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim SvTestCd As String
    Dim tmpRs As New DrRecordSet
   
    '������� �˻�
    SqlStmt = MySql.SqlQueryResults(pWorkArea, pAccDt, pAccSeq, "1")
    Set tmpRs = OpenRecordSet(SqlStmt)
   
    If tmpRs.EOF Then GoTo NoData
   
    ResultCnt = tmpRs.RecordCount
   
    '������� Array
    ReDim Preserve MyResult(ResultCnt)
    ReDim Preserve RstForeColor(MaxColumns, ResultCnt)
   
    For I = 1 To ResultCnt
       
        Call AddRow(1)
        With MyResult(I)
            .TestCd = Trim("" & tmpRs.Fields("TestCd").Value)
            .TxtFg = Trim("" & tmpRs.Fields("TxtFg").Value)

            '�ؽ�Ʈ ���
            If .TxtFg = ERT_TxtRst Then
                ResultClipText = ResultClipText & "<< �˻� �Ұ� >>  - " & tmpRs.Fields("TestShortNm").Value & vbCrLf
                ResultClipText = ResultClipText & tmpRs.Fields("TextResult").Value & vbCrLf & vbCrLf
                'Supplemental Report
                Dim SuppRs As Object
                SqlStmt = MySql.SqlETextRst(pWorkArea, pAccDt, pAccSeq, .TestCd)
                Set SuppRs = OpenRecordSet(SqlStmt)
                If Not SuppRs.EOF Then
                    ResultClipText = ResultClipText & "<< Supplemental Report >> " & vbCrLf
                    While (Not SuppRs.EOF)
                        ResultClipText = ResultClipText & SuppRs.Fields("TextResult").Value & vbCrLf
                        SuppRs.MoveNext
                    Wend
                    SuppRs.RsClose
                    Set SuppRs = Nothing
                End If
                TextFg = True
            End If
                  
        End With
        tmpRs.MoveNext
      
    Next

NoData:
    tmpRs.RsClose
    Set tmpRs = Nothing
    
End Sub
         
'% �̻��� �˻� ( LAB404, LAB405 ) : Lab No�� �������� �̻��� ������� �˻�
Public Sub MicrobeResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer)
 
    Dim I As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpRs As New DrRecordSet
   
    '������� �˻�
    SqlStmt = MySql.SqlQueryResults(pWorkArea, pAccDt, pAccSeq, "2")
    Set tmpRs = OpenRecordSet(SqlStmt)
   
    If tmpRs.EOF Then GoTo NoData
   
    ResultCnt = tmpRs.RecordCount
   
    '������� Array
    ReDim Preserve MyResult(ResultCnt)
    ReDim Preserve RstForeColor(MaxColumns, ResultCnt)
   
        For I = 1 To ResultCnt
            
            With MyResult(I)
               .TestCd = Trim("" & tmpRs.Fields("TestCd").Value)
               .RstCd = Trim("" & tmpRs.Fields("RstCd").Value)
               .RstDiv = Trim("" & tmpRs.Fields("RstDiv").Value)
               .LastRst = Trim("" & tmpRs.Fields("LastRst").Value)
               .LastVfyDtTm = Trim("" & tmpRs.Fields("LstVfyDtTm").Value)
               .LastVfyId = Trim("" & tmpRs.Fields("LastVfyNm").Value)
               .VfyDt = Trim("" & tmpRs.Fields("VfyDt").Value)
               .VfyTm = Trim("" & tmpRs.Fields("VfyTm").Value)
               .VfyId = Trim("" & tmpRs.Fields("VfyNm").Value)
               .MfyFg = Trim("" & tmpRs.Fields("MfySeq").Value)
               .RstType = Trim("" & tmpRs.Fields("RstType").Value)
               .DetailFg = Trim("" & tmpRs.Fields("DetailFg").Value)
               .PtId = Trim("" & tmpRs.Fields("PtId").Value)
               .OrdDt = Trim("" & tmpRs.Fields("OrdDt").Value)
               .OrdNo = Trim("" & tmpRs.Fields("OrdNo").Value)
               .OrdSeq = Trim("" & tmpRs.Fields("OrdSeq").Value)
               .SenFg = Trim("" & tmpRs.Fields("SenFg").Value)
               
               '������ ��� ��ȸ
               If .SenFg = "Y" Then
                  
                  Call AddRow(1)
                  SenClipText = SenClipText & vbCrLf
   
                  Call AddRow(1)
                  RstForeColor(2, RstRow) = &H747474 '���� ȸ��
                  SenClipText = SenClipText & Trim(tmpRs.Fields("TestLongNm").Value) & vbCrLf
                  Call AddRow(1)
                  SenClipText = SenClipText & vbCrLf
                  Call SenResult(pWorkArea, pAccDt, pAccSeq, .TestCd)
               Else
                  '�˻�� (�Ϲ��׸�/���׸�)
                  Call AddRow(1)
                  If .DetailFg = "" Or .RstDiv = "*" Then
                     .TestShortNm = Trim("" & tmpRs.Fields("TestShortNm").Value)
                  Else
                     .TestShortNm = "   " & Trim(tmpRs.Fields("TestShortNm").Value)
                  End If
                  .TestLongNm = Trim("" & tmpRs.Fields("TestLongNm").Value)  '�˻�� Full Name
                  .RstCdNm = Trim("" & tmpRs.Fields("RstCdNm1").Value)           '����ڵ��
                  If .RstCdNm = "" Then .RstCdNm = Trim("" & tmpRs.Fields("RstCdNm2").Value)           '����ڵ��
   
                  '�����(�ڵ��� ���..)
                  If .VfyDt = "" Then
                     If .RstDiv <> "*" Then .RstCd = "��Ȯ": RstForeColor(3, RstRow) = &HC0C0FF
                  Else
                     RstForeColor(3, RstRow) = &H404080   '����
                     If .RstCdNm <> "" Then .RstCd = .RstCdNm
                  End If
                  RstForeColor(2, RstRow) = &H747474 '���� ȸ��
                  'RstForeColor(3, RstRow) = &H404080 '����
                  ResultClipText = ResultClipText & OneRow(I)       '������� Buffering
                  OffSet = OffSet + 1
               End If
   
            End With
            tmpRs.MoveNext
         
         'Wend
         Next

NoData:
   tmpRs.RsClose
   'tmpRs.CloseCursor
   Set tmpRs = Nothing
   
End Sub

'% �̻��� �˻� ( LAB404, LAB405 ) : �����������...
Public Sub MicrobeMore(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer)
 
   Dim I As Integer
   Dim SqlStmt As String
   Dim ColCnt As Integer
   Dim tmpRs As New DrRecordSet
   
   '������� �˻�
   SqlStmt = MySql.SqlQueryResults(pWorkArea, pAccDt, pAccSeq, "2")
   Set tmpRs = OpenRecordSet(SqlStmt)
   
   If tmpRs.EOF Then GoTo NoData
   
   ResultCnt = tmpRs.RecordCount
   
   '������� Array
   ReDim Preserve MyResult(ResultCnt)
   ReDim Preserve RstForeColor(MaxColumns, ResultCnt)
   
   For I = 1 To ResultCnt
            
      With MyResult(I)
         .TestCd = Trim("" & tmpRs.Fields("TestCd").Value)
         .SenFg = Trim("" & tmpRs.Fields("SenFg").Value)
         
         '������ ��� ��ȸ
         If .SenFg = "Y" Then
                  
            Call AddRow(1)
            SenClipText = SenClipText & vbCrLf
   
            Call AddRow(1)
            RstForeColor(2, RstRow) = &H747474 '���� ȸ��
            SenClipText = SenClipText & Trim(tmpRs.Fields("TestLongNm").Value) & vbCrLf
            Call AddRow(1)
            SenClipText = SenClipText & vbCrLf
            Call SenResult(pWorkArea, pAccDt, pAccSeq, .TestCd)
         End If

      End With
      tmpRs.MoveNext
   
   Next

NoData:
   tmpRs.RsClose
   Set tmpRs = Nothing
   
End Sub

'% ��������� ��ȸ
Public Sub SenResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, ByVal pTestCd As String)

   Dim SqlStmt As String
   Dim ColCnt As Integer
   Dim tmpRs As New DrRecordSet
   Dim tmpRs1 As New DrRecordSet
   Dim I As Integer, j As Integer
   Dim AntiList As ListBox, AntiCnt As Integer
   Dim AntiRst As String, AntiCd As String, AntiNm As String
   Dim AntiSeq As Integer, MicroCnt As Integer, GrowthQty As String
   Dim tmpTitle As String
   Dim tmpAntiCnt As Integer
   Dim objSql As New clsLISSqlStatement
   
   SqlStmt = MySql.SqlSenResult(pWorkArea, pAccDt, pAccSeq, pTestCd)
   Set tmpRs = OpenRecordSet(SqlStmt)
   
   If tmpRs.EOF Then GoTo NoData
   
   MicroCnt = tmpRs.RecordCount  '�հ���
   ReDim MySenResult(MaxAntiCnt, MicroCnt)
   
   Set AntiList = frmControls.lstUnsortedList   '�׻��� ����Ʈ(Unsorted)
   AntiList.Clear
   
   tmpAntiCnt = 0
   tmpTitle = "Antibiotics  " & vbTab
   
   '�̻��� ������ ������� Buffering (LAB405)
   For I = 1 To MicroCnt
      
      Call AddRow(1)
      RstForeColor(2, RstRow) = &H747474 '���� ȸ��
      RstForeColor(3, RstRow) = &H404080 '����
      
      AntiCnt = Val("" & tmpRs.Fields("SCnt").Value)  '�׻��� ����
      If tmpAntiCnt < AntiCnt Then tmpAntiCnt = AntiCnt
      If AntiCnt > 0 Then tmpTitle = tmpTitle & " " & CStr(I) & " "
      
      GrowthQty = Trim(tmpRs.Fields("GrowthQty").Value)   '�������
      SenClipText = SenClipText & Trim(tmpRs.Fields("Seq").Value) & "  "
      'SenClipText = SenClipText & GrowthQty & Space(6 - Len(GrowthQty)) & vbTab '�������
      SenClipText = SenClipText & GrowthQty & vbTab '�������
      SenClipText = SenClipText & Trim(tmpRs.Fields("MicroNm").Value) & vbCrLf   '�ո�
      
      For j = 1 To AntiCnt
         AntiRst = Trim("" & tmpRs.Fields("SRst" & CStr(j)).Value)
         AntiCd = medShift(AntiRst, ";")
         AntiSeq = medListFind(AntiList, AntiCd)
         '���ο� �׻����� ���� ��츸 ����Ʈ�� Add...
         If (AntiSeq < 0) Or (AntiCd <> AntiList.List(AntiSeq)) Then
            AntiList.AddItem AntiCd
            AntiSeq = AntiList.ListCount - 1
         End If
         
         MySenResult(AntiSeq, I).RstCd = AntiRst    '��� Keeping
      Next
      tmpRs.MoveNext
   Next
   
   Call AddRow(1): SenClipText = SenClipText & vbCrLf
   
   If tmpAntiCnt > 0 Then  '����������� ������...
      Call AddRow(1): SenClipText = SenClipText & "[ Susceptibility test ]" & vbCrLf
      RstForeColor(2, RstRow) = &HE48372  '�ణ �Ķ���
      Call AddRow(1): SenClipText = SenClipText & tmpTitle & vbCrLf
      RstForeColor(2, RstRow) = &H7477EF '�ణ ������
      RstForeColor(3, RstRow) = &H7477EF '�ణ ������
   End If

   '������ ��� Buffering
   For I = 1 To AntiList.ListCount
      Call AddRow(1)
      If I = 1 Then SortStartRow = RstRow
      RstForeColor(2, RstRow) = &H747474 '���� ȸ��
      RstForeColor(3, RstRow) = &H404080 '����
      AntiCd = AntiList.List(I - 1)
      Set tmpRs1 = OpenRecordSet(objSql.SqlCommonCode(T_LAB032, CD2_AntiBiotic, AntiCd))
      'AntiNm = Mid(Trim(tmpRs1.Fields("Text1").Value), 1, 15)   '�׻�����
      'SenClipText = SenClipText & AntiNm & Space(15 - Len(AntiNm)) & vbTab  '�׻�����
      AntiNm = Trim(tmpRs1.Fields("Text1").Value)   '�׻�����
      SenClipText = SenClipText & AntiNm & vbTab  '�׻�����
      tmpRs1.RsClose
      For j = 1 To MicroCnt
         If MySenResult(I - 1, j).RstCd = "" Then
            SenClipText = SenClipText & "   "
         Else
            SenClipText = SenClipText & " " & MySenResult(I - 1, j).RstCd & " " '�պ� ���������
         End If
      Next
      SenClipText = SenClipText & "                        " & vbTab
      SenClipText = SenClipText & vbTab & vbTab & vbTab & vbTab & AntiNm & vbCrLf
      'Call AddRow(1)
   Next
   SortEndRow = RstRow - 1
   SortFg = True
   
NoData:
   tmpRs.RsClose
   Set tmpRs = Nothing
   Set tmpRs1 = Nothing
   Set AntiList = Nothing
   Set objSql = Nothing
   
End Sub


'% ��ü����ũ Text
'% -. LAB034 : CdIndex-'C404'�� Text1
Public Function ReadRemark(ByVal RmkCd As String) As String
   
   Dim tmpSql As String
   Dim tmpRs As Object
   
   Remark = ""
   ReadRemark = ""
   tmpSql = MySql.SqlGetRemark(RmkCd)
   Set tmpRs = OpenRecordSet(tmpSql)
   
   If tmpRs.EOF Then GoTo NoData
   
   ReadRemark = Trim(tmpRs.Fields("Remark").Value)
   Remark = "<< Remark >>" & vbCrLf & Trim(tmpRs.Fields("Remark").Value) & vbCrLf & vbCrLf
   CommentFg = True

NoData:
   tmpRs.RsClose
   Set tmpRs = Nothing
   
End Function

'% FootNote Text
'% -. LAB304 : RstTxt
Public Function ReadFootNote(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer) As String
   
   Dim I As Integer
   Dim tmpSql As String
   Dim tmpRs As Object
   
   FootNote = ""
   ReadFootNote = ""
   tmpSql = MySql.SqlGetFootNote(pWorkArea, pAccDt, pAccSeq)
   Set tmpRs = OpenRecordSet(tmpSql)
   If tmpRs.EOF Then GoTo NoData
   
   FootNote = "<< Foot Note >>" & vbCrLf
   ReadFootNote = Trim(tmpRs.Fields("FootNote").Value)
   While (Not tmpRs.EOF)
      FootNote = FootNote & Trim(tmpRs.Fields("FootNote").Value) & vbCrLf
      tmpRs.MoveNext
   Wend
   CommentFg = True
   
NoData:
   tmpRs.RsClose
   Set tmpRs = Nothing
   
End Function


'% ������� Spread�� Display�� Clip Text�� �ۼ��Ѵ�.
'% -. �˻��(Short), ���, ����, High/Low, Delta/Panic, ����ġ, �˻��(Long), �ֱٰ��, �ֱٰ��Ȯ����, �Ӽ��ڵ�
Public Function OneRow(ByVal Row As Integer)

   OneRow = ""
   If Row > UBound(MyResult) Then Exit Function
   With MyResult(Row)
      OneRow = .TestShortNm & vbTab & .RstCd & vbTab & .RstUnit & vbTab & .HLDiv & vbTab & .DPDiv & vbTab & _
                     .RefCd & vbTab & .TestLongNm & vbTab & .LastRst & vbTab & .LastVfyDtTm & vbTab & .AttrCd & vbCrLf
   End With
   
End Function

'% �Ϲݰ�� - �ؽ�Ʈ��� ���� (������������)  Buffring
Public Function OneTextResult(ByVal Row As Integer)

   OneTextResult = ""
   If Row > UBound(MyResult) Then Exit Function
   If MyResult(Row).RstText <> "" Then
      OneTextResult = OneTextResult & "<< �˻� �Ұ� >> - " & MyResult(Row).TestShortNm & vbCrLf
      OneTextResult = OneTextResult & MyResult(Row).RstText & vbCrLf & vbCrLf
      TextFg = True
   End If
   If MyResult(Row).SuppText <> "" Then
      OneTextResult = OneTextResult & "<< Supplemental Report >> " & vbCrLf
      OneTextResult = OneTextResult & MyResult(Row).SuppText & vbCrLf
      TextFg = True
   End If
      
End Function


'% Work Area �� : LAB032 (C213) - Field1
Public Function GetWorkAreaNm(ByVal WorkAreaCd As String)

   Dim tmpRs As Object
   Dim objSql As New clsLISSqlStatement
   
   Set tmpRs = OpenRecordSet(objSql.SqlCommonCode(T_LAB032, CD2_WorkArea, WorkAreaCd))
   If tmpRs.EOF Then
      GetWorkAreaNm = ""
   Else
      GetWorkAreaNm = Trim(tmpRs.Fields("Field1").Value)
   End If
   tmpRs.RsClose
   Set tmpRs = Nothing
   Set objSql = Nothing
   
End Function

'% �ؽ�Ʈ��� ����
'% -. LAB303
Public Function GetRstText(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, ByVal pTestCd As String)

   Dim I As Integer
   Dim tmpRs As Object
   
   Set tmpRs = OpenRecordSet(MySql.SqlGetSuppText(pWorkArea, pAccDt, pAccSeq, pTestCd))
   If tmpRs.EOF Then
      GetRstText = ""
   Else
      GetRstText = Trim("" & tmpRs.Fields("RstTxt").Value)
   End If
   tmpRs.RsClose
   Set tmpRs = Nothing
   
End Function


'% �ؽ�Ʈ��� ��������
'% -. LAB305
Public Function GetSuppText(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, ByVal pTestCd As String)

   Dim I As Integer
   Dim tmpRs As Object
   
   Set tmpRs = OpenRecordSet(MySql.SqlGetSuppText(pWorkArea, pAccDt, pAccSeq, pTestCd))
   If tmpRs.EOF Then
      GetSuppText = ""
   Else
      GetSuppText = ""
      While (Not tmpRs.EOF)
         GetSuppText = GetSuppText & Trim(tmpRs.Fields("RstTxt").Value) & vbCrLf & _
                              "�� ������ : " & Trim(tmpRs.Fields("EmpNm").Value) & "   " & Trim(tmpRs.Fields("MfyDtTm").Value) & vbCrLf
         tmpRs.MoveNext
      Wend
   End If
   tmpRs.RsClose
   Set tmpRs = Nothing
   
End Function

'% ��Ÿ�˻� Value Result ( LAB352 )
Public Sub GetRstValues(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, _
                                    ByVal pTestCd As String, ByVal MfyFg As String)
   
   Dim tmpRs As Object
   Dim tmpSql As String
   Dim I As Integer
   
   tmpSql = MySql.SqlEValueRst(pWorkArea, pAccDt, pAccSeq, pTestCd, MfyFg)
   Set tmpRs = OpenRecordSet(tmpSql)
   If tmpRs.EOF Then GoTo NoData
   
   While (Not tmpRs.EOF)
      Call AddRow(1)
      ResultClipText = ResultClipText & "   " & Trim(tmpRs.Fields("Head").Value) & vbTab
      ResultClipText = ResultClipText & Trim(tmpRs.Fields("ValRst").Value) & vbTab
      ResultClipText = ResultClipText & Trim(tmpRs.Fields("Unit").Value) & vbCrLf
      RstForeColor(2, RstRow) = &H747474 '���� ȸ��
      RstForeColor(3, RstRow) = &H404080 '����
      tmpRs.MoveNext
   Wend
   
NoData:
   tmpRs.RsClose
   Set tnors = Nothing

End Sub

'% ���� ����
Public Function GetEmpNm(ByVal EmpId As String)

   GetEmpNm = ""
   If EmpId = "" Then Exit Function
   
   If ObjLISComCode.EmpId.Exists(EmpId) Then
      Call ObjLISComCode.EmpId.KeyChange(EmpId)
      GetEmpNm = ObjLISComCode.EmpId.Fields("EmpNm")
   Else
      GetEmpNm = ""
   End If
    
End Function

'% �ǻ��̸�
Public Function GetDoctNm(ByVal DoctId As String)

   GetDoctNm = ""
   If DoctId = "" Then Exit Function
   
   If ObjLISComCode.doct.Exists(DoctId) Then
      Call ObjLISComCode.doct.KeyChange(DoctId)
      GetDoctNm = ObjLISComCode.doct.Fields("DoctNm")
   Else
      GetDoctNm = ""
   End If
   
End Function

'% �μ��̸�
Public Function GetDeptNm(ByVal DeptCd As String)

   GetDeptNm = ""
   If DeptCd = "" Then Exit Function
   
   If ObjLISComCode.DeptCd.Exists(DeptCd) Then
      Call ObjLISComCode.DeptCd.KeyChange(DeptCd)
      GetDeptNm = ObjLISComCode.DeptCd.Fields("DeptNm")
   Else
      GetDeptNm = ""
   End If

End Function

Private Sub AddRow(ByVal Rows As Integer)
   RstRow = RstRow + Rows
   ReDim Preserve RstForeColor(MaxColumns, RstRow)
   
End Sub


'% High/Low �÷� ForeColor
Public Function Get_ForeColor(ByVal Col As Integer, ByVal Row As Integer) As Long

   Get_ForeColor = RstForeColor(Col, Row)
   'Get_HL_ForeColor = MyResult(Row).HLForeColor
   
End Function

'% Initialize Routine
Private Sub InitRtn()
   
   '���� �ʱ�ȭ
   RstRow = 0
   ResultCnt = 0
   Remark = ""
   FootNote = ""
   ResultClipText = ""
   SenClipText = ""
   SamTextBuffer = ""
   RstTextBuffer = ""
   SuppText = ""
   SortStartRow = 0
   SortEndRow = 0
   OffSet = 0
   
   CommentFg = False
   TextFg = False
   SortFg = False
   
   Erase MyResult
   Erase MySenResult
   ReDim MyResult(1)
   ReDim RstForeColor(MaxColumns, 1)

End Sub

Public Sub GetRelTest(ByRef objCombo As Object, ByVal pLabNo As String)

   Dim SqlStmt As String
   Dim pWorkArea As String
   Dim pAccDt As String
   Dim pAccSeq As String
   Dim tmpRs As New DrRecordSet
   
   pWorkArea = medGetP(pLabNo, 1, "-")
   pAccDt = medGetP(pLabNo, 2, "-")
   pAccSeq = medGetP(pLabNo, 3, "-")
   
   If Mid(pAccDt, 1, 1) = "9" Then
      pAccDt = "19" & pAccDt
   Else
      pAccDt = "20" & pAccDt
   End If
   SqlStmt = MySql.SqlGetRelTest(pWorkArea, pAccDt, pAccSeq)
   Set tmpRs = OpenRecordSet(SqlStmt)
   
   objCombo.Clear
   If tmpRs.EOF Then
      objCombo.AddItem "< ���� >"
      GoTo NoData
   End If
   
   With tmpRs
      For I = 1 To .RecordCount
         objCombo.AddItem .Fields("AbbrNm5").Value & "   " & .Fields("RstCd").Value & "   " & .Fields("RstUnit").Value & "   " & _
                                    .Fields("VfyDtTm").Value & "  by " & .Fields("VfyNm").Value
         .MoveNext
      Next
   End With
   
NoData:
   tmpRs.RsClose
   Set tmpRs = Nothing
   
End Sub









