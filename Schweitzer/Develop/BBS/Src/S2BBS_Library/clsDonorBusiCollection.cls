VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDonorBusiCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'+--------------------------------------------------------------------------------------+
'|  1. Class �� : clsCollection(DLL��)
'|  2. ��    �� : ä��/��������(LAB201)�� �����Ѵ�.
'|  3. �� �� �� : ��̰�
'|
'|  CopyRight(C) 1999 ��ÿ�Ƽ����
'|
'+--------------------------------------------------------------------------------------+


'%  Ŭ���� clsCollection�� Data Attributes
Public spcyy As String      '/* ��ü��ȣ-�⵵   */
Public spcno As Long        '/* ��ü��ȣ        */
Public PtId As String       '/* ȯ�� ID         */
Public PtNm As String       '/* ȯ�ڸ�         */
Public AgeDay As Long       '/* ȯ�� �Ϸ�       */
Public sex As String        '/* ȯ�� ����   */
Public BedIndt As String    '/* �Կ���          */
Public DeptCd As String     '/* �����          */
Public Orddoct As Integer   '/* ó����          */
Public MajDoct As Integer   '/* ��ġ��          */
Public WorkArea As String   '/* Work Area(������ȣ��) */
Public accdt As String      '/* ������(������ȣ��)    */
Public accseq As Integer    '/* ��������(������ȣ��)  */
Public stscd As String      '/* STATUS - '0 ':ó��,1:ä��,2:����,3:In-Process,4:���,5:Ȯ��,6:���� */
Public ReqTotCnt As Integer '/* Required Item Total Count */
Public ReqInputCnt As Integer  '/* Required Item Input Count */
Public VfyDt As String      '/* ���Ȯ������    */
Public VfyTm As String      '/* ���Ȯ�νð�    */
Public VfyId As Integer     '/* ���Ȯ����      */
Public coldt As String      '/* ä������        */
Public coltm As String      '/* ä���ð�        */
Public colid As Integer     '/* ä����          */
Public rcvdt As String      '/* ��������        */
Public rcvtm As String      '/* �����ð�        */
Public rcvid As Integer     '/* ������          */
Public EntDt As String      '/* �Է�����        */
Public EntTm As String      '/* �Է½ð�        */
Public entid As Integer     '/* �Է���          */
Public SpcCd As String      '/* ��ü�ڵ�        */
Public SpcNm As String      '/* ��ü��        */
Public MultiFg As String    '/* ������ü����  */
Public OrgAccNo As String   '/* �� ������ȣ     */
Public wardid As String     '/* �����ڵ�        */
Public RoomId As String     '/* �����ڵ�     */
Public BedID As String      '/* ħ���ڵ�    */
Public HosilID As String    '/* ȣ���ڵ�    */
Public FootNoteFg As String '/* FOOTNOTE����('0':��,'1':��) */
Public StoreCd As String    '/* ��ü���� �ڵ�   */
Public RptFg As String      '/* REPORT��¿���('0':��,'1':��) */
Public testdiv As String    '/* �˻籸�� - 1:��Ÿ�˻�,2:�̻����˻� */
Public QcFg As String       '/* QC����('0':��,'1':��) */
Public RmkCd As String      '/* ��ü Remark �ڵ� */
Public statfg As String     '/* ���޿���    */
Public BuildCd As String    '/* �ǹ��ڵ�    */
Public BuildNm As String    '/* �ǹ��ڵ��    */
Public orgbuildcd As String '/* ä���� �ǹ��ڵ�    */
Public TestNames As String  '/* Barcode�� ��µ� �˻���� */

Public orddt As String
Public OrdNo As Integer
Public OrdSeq As Integer
Public OrdCd As String

Public SpcGrp As String     '/* ��ü��  */
Public spcseq As String     '/* ������ü�� Seq */

'= Lab ä������ ���� Data Type ='
Private Type TpLabGroup
   BuildCd As String
   testdiv As String
   WorkArea As String
   SpcCd As String
   StoreCd As String
   statfg As String
   ColDate As String
   MultiFg As String
   SpcGrp As String
   orddt As String
   OrdNo As Integer
   OrdSeq As Integer
   OrdCd As String
   DeptCd As String     '/* �����          */
   Orddoct As Integer   '/* ó����          */
   MajDoct As Integer   '/* ��ġ��          */
   LabelCnt As Integer
   TestNm As String     '/* �˻� ����   */
End Type

Private Type LabNumbers
   WorkArea As String
   accdt As String
   accseq As Integer
   BuildCd As String
End Type
   
Private Type BarcodeLabel
   WorkArea As String
   accdt As String
   accseq As Integer
   coldt As String
   coltm As String
   orddt As String
   statfg As String
   spcno As String
   Location As String
   PtNm As String
   PtId As String
   SpcNm As String
   StoreCd As String
   wardid As String
   TestNames As String
   LabelCnt As String
End Type
   
   
Public ErrMsg As String
Public ColCount As Integer

'���뺯�� - �����ϴ� ������Ʈ�κ��� ������

Public CollectDone As Boolean

Private blnBatchCol As Boolean
Private ColDone As Boolean  '������ü ä���Ϸ� ����
Private ReadyToCollect As Boolean
Private KeyCount As Integer
Private MultiSpcFlag As String
Private BarCount As Integer
Private mvarSortList As Object
Private LabGroup() As TpLabGroup
Private CreateLabNo() As LabNumbers
Private BarcodeBuffer() As BarcodeLabel
Private SqlBuffer() As String
Private objMySQL As New clsBBSSQLStatement

Private sWorkDt As String
Private sWorkTm As String
Private iWorkSeq As Integer

Private ColSqlStmt() As String

'Private DbConn As DrDatabase

Private mvarDonorID As String
Private mvarDonorAccdt As String
Private mvarDateTime As Date    '���� ��¥�� �ִ´�.

Private objBarCode As clsBarcode

Public Property Let DateTime(ByVal vData As Date)
    mvarDateTime = vData
End Property

Public Property Get DateTime() As Date
    DateTime = mvarDateTime
End Property

Public Property Let donoraccdt(ByVal vData As String)
    mvarDonorAccdt = vData
End Property

Public Property Get donoraccdt() As String
    donoraccdt = mvarDonorAccdt
End Property

Public Property Let Donorid(ByVal vData As String)
    mvarDonorID = vData
End Property

Public Property Get Donorid() As String
    Donorid = mvarDonorID
End Property

'Public Sub setDbConn(ByRef pDbConn As DrDatabase)
''��� Ŀ�ؼ�
'
'    Set DbConn = pDbConn
'End Sub

Public Property Get SortList() As Object
    Set SortList = mvarSortList
End Property

Public Property Set SortList(ByVal vData As Object)
    Set mvarSortList = vData
End Property

'% ä�������� �����ϱ� ���� ���ʵ���Ÿ�� �����Ѵ�.
Public Sub AddLabCollect(ParamArray tmpData() As Variant)
   
   Dim tmpKeyString As String
   
   KeyCount = KeyCount + 1
   ReDim Preserve LabGroup(KeyCount)
   With LabGroup(KeyCount)
      .BuildCd = "" & tmpData(0)
      .testdiv = "" & tmpData(1)
      .WorkArea = "" & tmpData(2)
      .SpcCd = "" & tmpData(3)
      .StoreCd = "" & tmpData(4)
      .statfg = "" & tmpData(5)
      .ColDate = "" & tmpData(6)
      .MultiFg = "" & tmpData(7)
      .SpcGrp = "" & tmpData(8)
      .orddt = "" & tmpData(9)
      .OrdNo = Val("" & tmpData(10))
      .OrdSeq = Val("" & tmpData(11))
      .OrdCd = "" & tmpData(12)
      .DeptCd = "" & tmpData(13)
      .Orddoct = Val("" & tmpData(14))
      .MajDoct = Val("" & tmpData(15))
      .TestNm = "" & tmpData(16)
      .LabelCnt = Val("" & tmpData(17))
      tmpKeyString = .BuildCd & .WorkArea & .SpcCd & .StoreCd & .statfg & .ColDate
      mvarSortList.AddItem tmpKeyString & Chr(9) & KeyCount
   End With
   ReadyToCollect = True
   CollectDone = False
End Sub

'Public Sub SetWardCol(ByRef pWorkDt As String, ByRef pWorkTm As String)
'
'    Dim tmpRs As RECORDSET
'
'    blnBatchCol = True    '������ �ϰ�ä�� ����
'
'    sWorkDt = Format(Now, CS_DateDbFormat)
'    sWorkTm = Format(Now, CS_TimeDbFormat)
'
'    pWorkDt = sWorkDt
'    pWorkTm = sWorkTm
'
'    Set tmpRs = objMySQL.OpenRecordSet("SELECT max(seq) as Seq FROM " & T_LAB204 & " WHERE workdt = '" & sWorkDt & _
'                                            "' AND wardid = '" & wardid & "' AND worktm = '" & sWorkTm & "' ")
'    If tmpRs.EOF Then
'        iWorkSeq = 0
'    Else
'        iWorkSeq = Val(tmpRs.Fields("Seq").Value)
'    End If
'    tmpRs.RsClose
'    Set tmpRs = Nothing
'End Sub

'% ä��Procedure�� �����Ͽ� ä������������ �����Ѵ�.
'% ( ������ü�� ��� �Ѳ����� ��� ä��ó�� �Ѵ� --> �Ʒ� DoCollection_Backup�� �ѹ��� �ϳ���.. )
Public Function DoCollection(Optional ByRef ProgressBar As Variant) As Boolean
   
    Dim i As Integer, j As Integer
    Dim tmpStr As String
    Dim SubKey As String, SubIndex As Integer
    Dim NextKey As String
    Dim tmpSql As String
    Dim tmpSpcRs As Recordset
    Dim tmpMultiRs As Recordset
   
    If Not ReadyToCollect Then
       DoCollection = False
       Exit Function
    End If
       
    ReDim ColSqlStmt(0)

'    SaveKey = ""
    ColCount = 0
   
    stscd = "1" 'Status
    ReqTotCnt = 0       'Required Item Total Count
    ReqInputCnt = 0     'Required Item Input Count
    FootNoteFg = "0"    'FootNote����
    RptFg = "0"         '����Ʈ ��� ����
    QcFg = "0"          'QC����
    VfyDt = ""          '���Ȯ����
    VfyTm = ""          '���Ȯ�νð�
    VfyId = ""           '���Ȯ����
    rcvdt = ""          '������
    rcvtm = ""          '�����ð�
    rcvid = ""           '������
    RmkCd = ""          '��ü����ũ
    TestNames = ""      '�˻����String
       
    For i = 0 To mvarSortList.ListCount - 1
      
        tmpStr = mvarSortList.List(i)
        SubKey = medShift(tmpStr, vbTab)
        SubIndex = Val(tmpStr)
      
        If i < mvarSortList.ListCount - 1 Then
            tmpStr = mvarSortList.List(i + 1)
            NextKey = medShift(tmpStr, vbTab)
        Else
            NextKey = ""
        End If
        
        TestNames = TestNames & LabGroup(SubIndex).TestNm & ","      '** Barcode �󺧿� ��µ� �˻����...
        '>> ������ �޶��������� ä������ �Ѱ� ����...
        If NextKey <> SubKey Then
           
           BarCount = 0  '���ڵ������� �ʱ�ȭ
           
           With LabGroup(SubIndex)
            
                MultiSpcFlag = .MultiFg
                
                WorkArea = .WorkArea
                coldt = Format(mvarDateTime, CS_DateDbFormat)     'ä����
                coltm = Format(mvarDateTime, CS_TimeDbFormat)    'ä���ð�
                SpcCd = .SpcCd                          '��ü
                MultiFg = .MultiFg                      '������ü����
                StoreCd = .StoreCd                      '��ü��������
                testdiv = .testdiv                      '�˻籸��
                statfg = .statfg                        '���޿���
                SpcGrp = .SpcGrp                        '��ü��
                BuildCd = .BuildCd                      '** �ǹ��ڵ�
                
                orddt = .orddt
                OrdNo = .OrdNo
                OrdSeq = .OrdSeq
                
                Orddoct = .Orddoct
                MajDoct = .MajDoct
                DeptCd = .DeptCd
                
                Select Case BuildCd
                Case "10": BuildNm = "�߾�"
                Case "20": BuildNm = "����"
                Case "30": BuildNm = "����"
                Case "40": BuildNm = "����"
                Case "50": BuildNm = "����"
                End Select
                
                   spcno = GetSpcNo(spcyy)                  '��ü��ȣ
                   accseq = GetLabNo(WorkArea, accdt)       'AccSeq
                   If accseq = 0 Then GoTo Err_Trap      'AccSeq�� ����� �������� ���...
                   MultiFg = ""
                   ColDone = True
                   If BarCount < LabGroup(SubIndex).LabelCnt Then BarCount = LabGroup(SubIndex).LabelCnt   '������
                   Call SaveData(i, SubKey)  '����
                   '** Barcode ��³����� Buffering �Ѵ�.
                   Call PrintBarcode(ColCount, .ColDate)
                   TestNames = ""
                   If Not CollectDone Then GoTo Err_Trap
          End With
      End If
      'Progress Bar ����...
      If Not IsMissing(ProgressBar) Then
            If ProgressBar.Value = ProgressBar.Max Then ProgressBar.Max = ProgressBar.Max + 1
            ProgressBar.Value = ProgressBar.Value + 1
      End If
      DoEvents
   Next
   
   DoCollection = True
   Exit Function
   
Err_Trap:
   DoCollection = False
   Set tmpSpcRs = Nothing
   Set tmpMultiRs = Nothing
End Function

'% Ŭ������ ä�� ������ DB�� �����Ѵ�.
Public Sub SaveData(ByVal ii As Integer, ByVal KeyString As String)
   
   Dim SqlStmt As String
   Dim tmpStr As String
   Dim tmpRs As Recordset
   Dim i As Integer
   Dim lngRequestSeq As Long
   
   'ColCount ����
   ColCount = ColCount + 1
   ReDim Preserve CreateLabNo(ColCount)

   '--ä������ ���� ������ ���� ������ LabNo�� Keep�Ѵ�.
   CreateLabNo(ColCount).WorkArea = WorkArea
   CreateLabNo(ColCount).accdt = accdt
   CreateLabNo(ColCount).accseq = accseq
   CreateLabNo(ColCount).BuildCd = BuildCd
      
   '�˻��Ƿ� ������ �ְ����� �ҷ��´�.
    
   lngRequestSeq = GetMaxRequestSeq(mvarDonorID, mvarDonorAccdt, Format(mvarDateTime, PRESENTDATE_FORMAT)) + ColCount
    
   ReDim Preserve ColSqlStmt(UBound(ColSqlStmt) + 1)
   
   On Error GoTo Err_Trap

'--ä���������� ����
   ColSqlStmt(UBound(ColSqlStmt)) = CreateSql_Collection   '// Sql���� ����
'   Debug.Print "Colsqlstmt(" & UBound(ColSqlStmt) & ") : " & ColSqlStmt(UBound(ColSqlStmt))
   
'--�˻��Ƿڳ��� SQL
    
    ReDim Preserve ColSqlStmt(UBound(ColSqlStmt) + 1)
    
    ColSqlStmt(UBound(ColSqlStmt)) = objMySQL.SetTestRequest(mvarDonorID, mvarDonorAccdt, _
                            Format(mvarDateTime, PRESENTDATE_FORMAT), lngRequestSeq, CreateLabNo(ColCount).WorkArea, CreateLabNo(ColCount).accdt, CreateLabNo(ColCount).accseq)

'--������ ���� ���� ������Ʈ
    
    ReDim Preserve ColSqlStmt(UBound(ColSqlStmt) + 1)
    ColSqlStmt(UBound(ColSqlStmt)) = objMySQL.SetDonorAccHistoryUpdateByTmpID(mvarDonorID, mvarDonorAccdt, PtId)

'--ó�� Status Follow Up
   
   For i = 0 To mvarSortList.ListCount - 1
      tmpStr = mvarSortList.List(i)
      If KeyString <> medShift(tmpStr, Chr(9)) Then GoTo Skip
      With LabGroup(Val(tmpStr))
         If ColDone Then            '--> ���Ӱ˻��� ��� ��� ä�������� �Ǿ� ó��Status Follow Up.
            'LAB102�� DoneFg = '1' : ä���Ϸ�
            ReDim Preserve ColSqlStmt(UBound(ColSqlStmt) + 1)
            
            ColSqlStmt(UBound(ColSqlStmt)) = CreateSql_UpdateLabNo(PtId, .orddt, .OrdNo, .OrdSeq, "1", "1") '// Sql���� ����
         Else
            'LAB102�� DoneFg = '0' : ä���̿Ϸ�
            
            ReDim Preserve ColSqlStmt(UBound(ColSqlStmt) + 1)
            
            ColSqlStmt(UBound(ColSqlStmt)) = CreateSql_UpdateLabNo(PtId, .orddt, .OrdNo, .OrdSeq, "1")  '// Sql���� ����
         End If
            
         SqlStmt = CreateSql_HeaderUpdate(PtId, .orddt, .OrdNo, "1", 1)
        
         Set tmpRs = New Recordset
         tmpRs.Open SqlStmt, DBConn
         If tmpRs.EOF Then
            Set tmpRs = Nothing
            
            ReDim Preserve ColSqlStmt(UBound(ColSqlStmt) + 1)
            
            ColSqlStmt(UBound(ColSqlStmt)) = CreateSql_HeaderUpdate(PtId, .orddt, .OrdNo, "1", 2)
         Else
            Set tmpRs = Nothing
         End If
      End With
Skip:
   Next
     
   CollectDone = True
   Exit Sub

Err_Trap:
   CollectDone = False
   Set tmpRs = Nothing
End Sub

Private Function GetMaxRequestSeq(ByVal Donorid As String, ByVal donoraccdt As String, _
                                  ByVal orddt As String) As Long
'DonorID, DonorAccDt, OrdDt �� ���� ���ǿ���
'�˻��Ƿڹ�ȣ�� �ִ밪�� ���´�.
                                  
    Dim RsMaxRSeq As Recordset
    Dim objMaxRSeq As clsBBSSQLStatement
    
    Set objMaxRSeq = New clsBBSSQLStatement
    With objMaxRSeq
'        .setDbConn DbConn
        Set RsMaxRSeq = New Recordset
        RsMaxRSeq.Open .GetTestRequestMaxSeq(Donorid, donoraccdt, orddt), DBConn
    End With
    
    If RsMaxRSeq.EOF Then
        GetMaxRequestSeq = 0
    Else
        If IsNull(RsMaxRSeq.Fields("maxseq").Value) Then
            GetMaxRequestSeq = 0
        Else
            GetMaxRequestSeq = RsMaxRSeq.Fields("maxseq").Value & ""
        End If
    End If
    
    Set RsMaxRSeq = Nothing
    Set objMaxRSeq = Nothing
End Function

Public Function ExecuteSqlStmt() As Boolean
    
    Dim i As Long

On Error GoTo ErrExecute

    For i = 1 To UBound(ColSqlStmt)
'        Debug.Print " i  = " & i & " : " & ColSqlStmt(i)
        DBConn.Execute ColSqlStmt(i)
    Next
        
    ExecuteSqlStmt = True
        
    Exit Function
ErrExecute:
    ExecuteSqlStmt = False
End Function

'% ��ü��ȣ�� �ο��Ѵ�.
Private Function GetSpcNo(ByVal spcyy As String) As Long
   
   Dim SpcNoRs As Recordset
   
   Dim tmpSql As String
   
   GetSpcNo = 0
   
   '���� Lock�� �ɱ����� �ǹ̾��� update�� ����.
   'select�� Lock�� �Ȱɸ��� ������ ���� ª�� ��������
   '��ü��ȣ�� �ߺ��ɼ� ����.. �� ����.    2000.3.10  by ��̰�
   tmpSql = CreateSql_SpcNo(spcyy, 4)
   
   DBConn.BeginTrans
   
   On Error GoTo Err_Trap
   
   DBConn.Execute tmpSql, , adCmdText   'Lock �ɸ�
   
   tmpSql = CreateSql_SpcNo(spcyy, 1)
   Set SpcNoRs = New Recordset
   SpcNoRs.Open tmpSql, DBConn
   If SpcNoRs.EOF Then
      GetSpcNo = 1
      tmpSql = CreateSql_SpcNo(spcyy, 2, GetSpcNo)
   Else
      GetSpcNo = Val("" & SpcNoRs.Fields("Seq").Value) + 1
      tmpSql = CreateSql_SpcNo(spcyy, 3, GetSpcNo)
   End If
   Set SpcNoRs = Nothing
   
   DBConn.Execute tmpSql, , adCmdText
   DBConn.CommitTrans
   
   Exit Function
   
Err_Trap:
    DBConn.RollbackTrans
'    DBConn.DisplayErrors
   Set SpcNoRs = Nothing
   
End Function

'% ������ȣ�� �ο��Ѵ�.
Private Function GetLabNo(ByVal pWorkArea As String, ByRef pAccDt As String) As Integer
   
   Dim tmpRs As Recordset
   Dim tmpSql As String
   Dim LabDiv As String
   Dim tmpStr As String
   Dim tmpRng1 As Integer, tmpRng2 As Integer
   Dim tmpSpcGrp As String
   
   GetLabNo = 0
   '// Sql��'������ȣ �ο�..
   tmpSql = objMySQL.SqlCommonCode(T_COM003, CD2_WORKAREA, pWorkArea)
   '//
   Set tmpRs = New Recordset
   tmpRs.Open tmpSql, DBConn
   LabDiv = "" & tmpRs.Fields("Field2").Value
   
   tmpRng1 = 1
   tmpRng2 = 9999
   tmpSpcGrp = "0"
   Select Case LabDiv
   Case "1":       '�ϴ���
      pAccDt = Format(GetSystemDate, PRESENTDATE_FORMAT)
   Case "2":       '������
      pAccDt = Format(GetSystemDate, "yyyyMM")
   Case "3":       '�����
      pAccDt = Format(GetSystemDate, "yyyy")
   Case "4":       '��ü����
      pAccDt = Format(GetSystemDate, "yyyyMM")
      tmpSpcGrp = SpcGrp
      '// Sql��
      tmpSql = objMySQL.SqlCommonCode(T_COM003, CD2_SGROUP, tmpSpcGrp)
      '//
      Set tmpRs = New Recordset
      tmpRs.Open tmpSql, DBConn
      '��ü�����ε� ��ü�� �����Ͱ� ���� ���... (�̷����� �߻��ϸ� �ȵ���...)
      If Not tmpRs.EOF Then
         tmpStr = "" & tmpRs.Fields("Field2").Value '��ü���� Lab No ����
         tmpRng1 = Val(medGetP(tmpStr, 2, ";"))     '��ü���� Lab No ���۹�ȣ
         tmpRng2 = Val(medGetP(tmpStr, 3, ";"))     '��ü���� Lab No ����ȣ
      End If
   Case Else:
      pAccDt = Format(GetSystemDate, PRESENTDATE_FORMAT)
   End Select
   
   '���� Lock�� �ɱ����� �ǹ̾��� update�� ����.
   'select�� Lock�� �Ȱɸ��� ������ ���� ª�� ��������
   '������ȣ�� �ߺ��ɼ� ����.. �� ����.    2000.3.10  by ��̰�
   tmpSql = CreateSql_LabNo(pWorkArea, pAccDt, tmpSpcGrp, 4)
   
   On Error GoTo Err_Trap
   
   DBConn.BeginTrans
   DBConn.Execute tmpSql, , adCmdText   'Lock �ɸ�
   
   '// Sql���� ����
   tmpSql = CreateSql_LabNo(pWorkArea, pAccDt, tmpSpcGrp, 1)
   '//
   Set tmpRs = New Recordset
   tmpRs.Open tmpSql, DBConn
   If tmpRs.EOF Then
      GetLabNo = tmpRng1
      tmpSql = CreateSql_LabNo(pWorkArea, pAccDt, tmpSpcGrp, 2, GetLabNo)
   Else
      GetLabNo = Val(tmpRs.Fields("Seq").Value & "") + 1
      If GetLabNo > tmpRng2 Then
         MsgBox "������ȣ�� Range(" & tmpRng1 & "-" & tmpRng2 & ")�� ������ϴ�. : " & GetLabNo
         GoTo Err_Trap
      End If
      tmpSql = CreateSql_LabNo(pWorkArea, pAccDt, tmpSpcGrp, 3, GetLabNo)
   End If
   Set tmpRs = Nothing
      
   DBConn.Execute tmpSql, , adCmdText
   DBConn.CommitTrans
   
   Exit Function
   
Err_Trap:
   DBConn.RollbackTrans
'   DBConn.DisplayErrors
   Set tmpRs = Nothing
   Resume Next
   GetLabNo = 0
   Exit Function

End Function

Public Sub GetLabNumbers(ByVal AryIndex As Integer, ByRef pWorkArea As String, _
                                        ByRef pAccDt As String, ByRef pAccSeq As Integer, Optional ByRef pBuildCd As Variant)
   
   If CollectDone Then
      With CreateLabNo(AryIndex)
         pWorkArea = .WorkArea
         pAccDt = .accdt
         pAccSeq = .accseq
         pBuildCd = .BuildCd
      End With
   End If
End Sub


'% Method 1 : CollectQuery
'%                 Parameter�� ���� Sql�� �����ϰ�, �� �ʵ��� ����
'%                 Ŭ���� clsCollection�� Data Attribute�� �����Ѵ�.

Public Function CollectQuery(ByVal pWorkArea As String, _
                                         ByVal pAccDt As String, ByVal pAccSeq As Integer) As Boolean

   Dim SqlStmt As String
   Dim MyOraDS As Recordset
   
   SqlStmt = " SELECT a.* " & _
                 " FROM " & T_LAB201 & " a " & _
                 " WHERE " & DBW("a.workarea", pWorkArea, 2) & _
                 " AND   " & DBW("a.accdt", pAccDt, 2) & _
                 " AND   " & DBW("a.accseq", pAccSeq, 2)
   
   Set MyOraDS = New Recordset
   MyOraDS.Open SqlStmt, DBConn
   If MyOraDS.EOF Then
      CollectQuery = False
      Exit Function
   End If
   
   With MyOraDS
      
      .MoveFirst
         
      spcyy = Trim("" & .Fields("SpcYy").Value)
      spcno = Val("" & .Fields("SpcNo").Value)
      PtId = Trim("" & .Fields("PtId").Value)
      sex = Trim("" & .Fields("Sex").Value)
      AgeDay = Val("" & .Fields("AgeDay").Value)
      BedIndt = Trim("" & .Fields("BedInDt").Value)
      DeptCd = Trim("" & .Fields("DeptCd").Value)
      Orddoct = Val("" & .Fields("OrdDoct").Value)
      MajDoct = Val("" & .Fields("MajDoct").Value)
      WorkArea = Trim("" & .Fields("WorkArea").Value)
      accdt = Trim("" & .Fields("AccDt").Value)
      accseq = Val("" & .Fields("AccSeq").Value)
      stscd = Trim("" & .Fields("StsCd").Value)
      ReqTotCnt = Val("" & .Fields("ReqTotCnt").Value)
      ReqInputCnt = Val("" & .Fields("ReqInputCnt").Value)
      VfyDt = Trim("" & .Fields("VfyDt").Value)
      VfyTm = Trim("" & .Fields("VfyTm").Value)
      VfyId = Val("" & .Fields("VfyId").Value)
      coldt = Trim("" & .Fields("ColDt").Value)
      coltm = Trim("" & .Fields("ColTm").Value)
      colid = Val("" & .Fields("ColId").Value)
      rcvdt = Trim("" & .Fields("RcvDt").Value)
      rcvtm = Trim("" & .Fields("RcvTm").Value)
      rcvid = Val("" & .Fields("RcvId").Value)
      EntDt = Trim("" & .Fields("EntDt").Value)
      EntTm = Trim("" & .Fields("EntTm").Value)
      entid = Val("" & .Fields("EntId").Value)
      SpcCd = Trim("" & .Fields("SpcCd").Value)
      MultiFg = Trim("" & .Fields("MultiFg").Value)
      OrgAccNo = Trim("" & .Fields("OrgAccNo").Value)
      wardid = Trim("" & .Fields("WardId").Value)
      RoomId = Trim("" & .Fields("RoomId").Value)
      BedID = Trim("" & .Fields("BedId").Value)
      HosilID = Trim("" & .Fields("HosilId").Value)
      FootNoteFg = Trim("" & .Fields("FootNoteFg").Value)
      StoreCd = Trim("" & .Fields("StoreCd").Value)
      RptFg = Trim("" & .Fields("RptFg").Value)
      testdiv = Trim("" & .Fields("TestDiv").Value)
      QcFg = Trim("" & .Fields("QcFg").Value)
      RmkCd = Trim("" & .Fields("RmkCd").Value)
      statfg = Trim("" & .Fields("StatFg").Value)
      BuildCd = Trim("" & .Fields("BuildCd").Value)
      
   End With
   
End Function

Private Function CreateSql_Collection() As String

   CreateSql_Collection = "Insert into " & _
                        "" & T_LAB201 & " (spcyy, spcno, ptid, sex, ageday, bedindt, deptcd, orddoct, majdoct, " & _
                        "            workarea, accdt, accseq, stscd, reqtotcnt, reqinputcnt, " & _
                        "            vfydt, vfytm, vfyid, coldt, coltm, colid, rcvdt, rcvtm, " & _
                        "            rcvid, entdt, enttm, entid, spccd, multifg, orgaccno, wardid, roomid, " & _
                        "            bedid, footnotefg, storecd, rptfg, testdiv, qcfg, rmkcd, statfg, buildcd, orgbuildcd, hosilid) " & _
                        "Values  (" & DBV("spcyy", spcyy, 1) & DBV("spcno", spcno, 1) & DBV("ptid", PtId, 1) & DBV("sex", sex, 1) & DBV("ageday", AgeDay, 1) _
                                    & DBV("bedindt", BedIndt, 1) & DBV("deptcd", DeptCd, 1) & DBV("orddoct", Orddoct, 1) & DBV("majdoct", MajDoct, 1) _
                                    & DBV("workarea", WorkArea, 1) & DBV("accdt", accdt, 1) & DBV("accseq", accseq, 1) & DBV("stscd", stscd, 1) _
                                    & DBV("reqtotcnt", ReqTotCnt, 1) & DBV("reqinputcnt", ReqInputCnt, 1) & DBV("vfydt", VfyDt, 1) & DBV("vfytm", VfyTm, 1) & DBV("vfyid", VfyId, 1) _
                                    & DBV("coldt", coldt, 1) & DBV("coltm", coltm, 1) & DBV("colid", colid, 1) & DBV("rcvdt", rcvdt, 1) & DBV("rcvtm", rcvtm, 1) _
                                    & DBV("rcvid", rcvid, 1) & DBV("entdt", CS_DateDbFormat, 1) & DBV("enttm", CS_TimeDbFormat, 1) & DBV("entid", entid, 1) & DBV("spccd", SpcCd, 1) _
                                    & DBV("multifg", MultiFg, 1) & DBV("orgaccno", OrgAccNo, 1) & DBV("wardid", wardid, 1) & DBV("roomid", RoomId, 1) & DBV("bedid", BedID, 1) _
                                    & DBV("footnotefg", FootNoteFg, 1) & DBV("storecd", StoreCd, 1) & DBV("rptfg", RptFg, 1) & DBV("testdiv", testdiv, 1) & DBV("qcfg", QcFg, 1) _
                                    & DBV("rmkcd", RmkCd, 1) & DBV("statfg", statfg, 1) & DBV("buildcd", BuildCd, 1) & DBV("orgbuildcd", orgbuildcd, 1) & DBV("hosilid", HosilID) & ") "

End Function


Private Function CreateSql_MultiSpc() As String

   CreateSql_MultiSpc = "Insert into " & T_LAB203 & " ( ptid, orddt, ordno, ordseq, spcseq, spccd, workarea, accdt, accseq ) " & _
                        "Values ( " & DBV("ptid", PtId, 1) & DBV("orddt", orddt, 1) & DBV("ordno", OrdNo, 1) & DBV("ordseq", OrdSeq, 1) & DBV("spcseq", spcseq, 1) _
                                    & DBV("spccd", SpcCd, 1) & DBV("workarea", WorkArea, 1) & DBV("accdt", accdt, 1) & DBV("accseq", accseq) & " )"

End Function

Private Function CreateSql_WardCol() As String

   CreateSql_WardCol = "Insert into " & T_LAB204 & " ( workdt, wardid, worktm, seq, workarea, accdt, accseq, colid, buildcd, orgbuildcd ) " & _
                       "Values ( " & DBV("workdt", sWorkDt, 1) & DBV("wardid", wardid, 1) & DBV("worktm", sWorkTm, 1) & DBV("seq", iWorkSeq, 1) & DBV("workarea", WorkArea, 1) & _
                                     DBV("accdt", accdt, 1) & DBV("accseq", accseq, 1) & DBV("colid", colid, 1) & DBV("buildcd", BuildCd, 1) & DBV("orgbuildcd", orgbuildcd) & " )"

End Function

Private Function CreateSql_UpdateLabNo(ByVal PtId As String, ByVal orddt As String, ByVal OrdNo As Integer, _
                                                           ByVal OrdSeq As String, ByVal stscd As String, Optional ByVal DoneFg As Variant) As String
   Dim tmpStr As String
   
   If IsMissing(DoneFg) Then
      tmpStr = ""
   Else
      tmpStr = ", donefg " & DBV("donefg", DoneFg, 2) & " "
   End If
'   CreateSql_UpdateLabNo = "Update " & T_LAB102 & " " & _
'                            "Set    " & DBW("workarea=", WorkArea) & ", " & _
'                            "       " & DBW("accdt   =", accdt) & ", " & _
'                            "       " & DBW("accseq  =", accseq) & ", " & _
'                            "       " & DBW("stscd   =", stscd) & " " & tmpStr & _
'                            "WHERE  " & DBW("ptid    =", PtId) & " " & _
'                            "AND    " & DBW("orddt   =", orddt) & " " & _
'                            "AND    " & DBW("ordno   =", OrdNo) & " " & _
'                            "AND    " & DBW("ordseq  =", OrdSeq)
   '�������� ��
   CreateSql_UpdateLabNo = "Update mdbldort " & _
                            "Set    " & DBW("workarea=", WorkArea) & ", " & _
                            "       " & DBW("accdt   =", accdt) & ", " & _
                            "       " & DBW("accseq  =", accseq) & ", " & _
                            "       " & DBW("stscd   =", stscd) & " " & tmpStr & _
                            "WHERE  " & DBW("patno=", PtId) & " " & _
                            "AND     orddate=to_date(" & orddt & ",'yyyymmdd')" & _
                            "AND    " & DBW("ordseqno   =", OrdNo)

End Function

Private Function CreateSql_HeaderUpdate(ByVal PtId As String, ByVal orddt As String, ByVal OrdNo As Integer, _
                                                            ByVal DoneFg As String, ByVal flag As Integer)
   
'   MsgBox "ó�泻�� ������Ʈ", vbExclamation
   
   If flag = 1 Then
   CreateSql_HeaderUpdate = "SELECT * FROM " & T_LAB102 & " " & _
                            "WHERE  " & DBW("ptid=", PtId) & " " & _
                            "AND    " & DBW("orddt=", orddt) & " " & _
                            "AND    " & DBW("ordno=", OrdNo) & " " & _
                            "AND    " & DBW("dcfg=", "") & " " & _
                            "AND    " & DBW("donefg<", DoneFg)
    Else
'   CreateSql_HeaderUpdate = "Update " & T_LAB101 & " " & _
'                            "Set    " & DBW("donefg=", DoneFg) & " " & _
'                            "WHERE  " & DBW("ptid=", PtId) & " " & _
'                            "AND    " & DBW("orddt=", orddt) & " " & _
'                            "AND    " & DBW("ordno=", OrdNo)
   '�������� ��
   CreateSql_HeaderUpdate = "Update mdbldort " & _
                            "Set    " & DBW("donefg=", DoneFg) & " " & _
                            "WHERE  " & DBW("patno=", PtId) & " " & _
                            "AND  orddate=to_date(" & orddt & ",'yyyymmdd') " & _
                            "AND    " & DBW("ordseqno=", OrdNo)
   End If
End Function

Private Function CreateSql_SpcNo(ByVal spcyy As String, ByVal flag As Integer, Optional ByVal seq As Variant) As String
   Select Case flag
   Case 1:
      CreateSql_SpcNo = "SELECT seq " & _
                        "FROM   " & T_LAB099 & " " & _
                        "WHERE  " & DBW("cdindex=", "01") & " " & _
                        "AND    " & DBW("divcd1=", "SPCNO") & " " & _
                        "AND    " & DBW("divcd2=", spcyy) & " " & _
                        "AND    " & DBW("divcd3=", "0")
   Case 2:
      CreateSql_SpcNo = "Insert Into " & T_LAB099 & " (cdindex, divcd1, divcd2, divcd3, seq) " & _
                        "Values (" & DBV("cdindex", "01", 1) & DBV("divcd1", "SPCNO", 1) & DBV("divcd2", spcyy, 1) & DBV("divcd3", "0", 1) & DBV("seq", seq) & ")"
   Case 3:
      CreateSql_SpcNo = "Update " & T_LAB099 & " Set " & DBW("seq=", seq) & " " & _
                        "WHERE " & DBW("cdindex=", "01") & " " & _
                        "AND   " & DBW("divcd1=", "SPCNO") & " " & _
                        "AND   " & DBW("divcd2=", spcyy) & " " & _
                        "AND   " & DBW("divcd3=", "0")
   Case 4:
      CreateSql_SpcNo = "Update " & T_LAB099 & " Set seq=seq " & _
                        "WHERE " & DBW("cdindex=", "01") & " " & _
                        "AND   " & DBW("divcd1=", "SPCNO") & " " & _
                        "AND   " & DBW("divcd2=", spcyy) & " " & _
                        "AND   " & DBW("divcd3=", "0")
   End Select
End Function

Private Function CreateSql_LabNo(ByVal WorkArea As String, ByVal accdt As String, ByVal SpcGrp As String, _
                                                 ByVal flag As Integer, Optional ByVal seq As Variant) As String
   Select Case flag
   Case 1:
      CreateSql_LabNo = "SELECT seq " & _
                        "FROM   " & T_LAB099 & " " & _
                        "WHERE " & DBW("cdindex=", "02") & " " & _
                        "AND   " & DBW("divcd1 =", WorkArea) & " " & _
                        "AND   " & DBW("divcd2 =", accdt) & " " & _
                        "AND   " & DBW("divcd3 =", SpcGrp)
   Case 2:
      CreateSql_LabNo = "Insert Into " & T_LAB099 & " (cdindex, divcd1, divcd2, divcd3, seq) " & _
                        "Values (" & DBV("cdindex", "02", 1) & DBV("divcd1", WorkArea, 1) & DBV("divcd2", accdt, 1) & DBV("divcd3", SpcGrp, 1) & DBV("seq", seq) & ")"
   Case 3:
      CreateSql_LabNo = "Update " & T_LAB099 & " Set " & DBW("seq=", seq) & " " & _
                        "WHERE  " & DBW("cdindex =", "02") & " " & _
                        "AND    " & DBW("divcd1  = ", WorkArea) & " " & _
                        "AND    " & DBW("divcd2  = ", accdt) & " " & _
                        "AND    " & DBW("divcd3  = ", SpcGrp)
   Case 4:
      CreateSql_LabNo = "Update " & T_LAB099 & " Set seq=seq " & _
                        "WHERE " & DBW("cdindex = ", "02") & " " & _
                        "AND   " & DBW("divcd1  = ", WorkArea) & " " & _
                        "AND   " & DBW("divcd2  = ", accdt) & " " & _
                        "AND   " & DBW("divcd3  = ", SpcGrp)
   End Select
End Function

Private Sub Class_Initialize()
   blnBatchCol = False
   Set objMySQL = New clsBBSSQLStatement
   Set objBarCode = New clsBarcode
'   Call InitRtn
End Sub

Private Sub Class_Terminate()
   Set objMySQL = Nothing
   Set objBarCode = Nothing
End Sub

Public Sub InitRtn()
   KeyCount = 0
   ColCount = 0
   mvarSortList.Clear
   ReadyToCollect = False
   CollectDone = False
   ColDone = False
   Erase SqlBuffer
   ErrMsg = ""
End Sub

'---------------------------
'���ڵ� ��� �׸��� ��Ƴ���
'---------------------------
Public Function PrintBarcode(ByVal ColCnt As Integer, ByVal pColTm As String) As String

   Dim tmpSql As String
   Dim tmpRs As Recordset

   ReDim Preserve BarcodeBuffer(ColCnt)

   tmpSql = objMySQL.SqlCommonCode(T_COM003, CD2_SPECIMEN, SpcCd)
   Set tmpRs = New Recordset
   tmpRs.Open tmpSql, DBConn
   SpcNm = "" & tmpRs.Fields("Field5").Value
   Set tmpRs = Nothing

   With BarcodeBuffer(ColCnt)
      .Location = BuildNm
      .WorkArea = WorkArea
      .orddt = Mid(orddt, 5, 2) & "/" & Mid(orddt, 7, 2)
      .coldt = Mid(accdt, 3)
      .coltm = Mid(pColTm, 12, 5)
      .accseq = Format(accseq, "@@@@@")
      .statfg = statfg
      .spcno = spcyy & Format(spcno, CS_BarFormat)
      .PtNm = Mid(PtNm, 1, 5)
      .PtId = PtId
      .SpcNm = SpcNm
      .StoreCd = StoreCd
      If Trim(wardid) <> "" Then
         .wardid = HosilID  'WardId
      Else
         .wardid = DeptCd
      End If
      .TestNames = TestNames
      .LabelCnt = BarCount
   End With

End Function

'---------------------------
'���� ���ڵ尡 ��µǴ� �Լ�
'---------------------------
Public Function GetBarcodeLabel(ByVal AryIndex As Integer, Optional ByVal AccFg As Boolean = False) As String
MsgBox "���ڵ� ����ϴ� �ų�... �Ⱦ��� ����.. ", vbCritical
   If CollectDone Then
      With BarcodeBuffer(AryIndex)
'        Call ObjBBSComCode.Barinfo.Label_PrintOut(.Location, .WorkArea, .coldt, .accseq, .spcno, .PtId, .PtNm, .SpcNm, .StoreCd, .statfg, .wardid, .orddt, .coltm, .TestNames, .LabelCnt)
        Call objBarCode.Label_PrintOut(.Location, .WorkArea, .coldt, .accseq, .spcno, .PtId, .PtNm, .SpcNm, .StoreCd, .statfg, .wardid, .orddt, .coltm, .TestNames, .LabelCnt)
      End With
   End If

End Function
