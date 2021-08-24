VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpecimen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'+--------------------------------------------------------------------------------------+
'|  1.  Class   ��  : clsSpecimen
'|  2.  �� ��         : ������ü ������(LAB004) �κ��� �ش� �˻�������� ��ü������ �˻�
'|  3. �ۼ���        : ��̰�
'|  4. �ۼ���        : 1999.05.07
'|
'|  CopyRight(C) 1999 ��ÿ�Ƽ����
'+--------------------------------------------------------------------------------------+


'%  Ŭ���� clsSpecimen�� Data Attributes

Public TestCd       As String               ' �˻��׸� �ڵ�
Public SpcCd        As String               ' ��ü �ڵ�
Public Seq          As Integer              ' Seq
Public ApplyDt      As String               ' ������
Public SpcGrpCd     As String               ' ��ü�� �ڵ�
Public LabelCnt     As Integer              ' ���ڵ�� ������
Public RstUnit      As String               ' ��� ����
Public RndFg        As String               ' Round ä������('0':��,'1':��)
Public StatFlags    As String               ' �� �ǹ��� Stat ����
Public StatFg       As String               ' Stat ����('0':��,'1':��)
Public AvalVal      As Integer              ' ��ȿ ����
Public PanicFg      As String               ' Panic Check����('0':��,'1':��)
Public PanicFrVal   As Double               ' Panic Range(From)
Public PanicToVal   As Double               ' Panic Range(To)
Public DeltaFg      As String               ' Delta Check����('0':��,'1':��)
Public DeltaVal1    As Integer              ' Delta Value
Public DeltaVal2    As Integer              ' Delta Value
Public TestCost     As String               ' �����ڵ�
Public StoreCd      As String               ' ��������
Public TatAvg       As Integer              ' TAT�ҿ�ð�
Public SpcQty       As Double               ' ��ü ü�뷮(TLA)
Public SpcUnit      As String               ' ��ü ä�� ����
Public ExpDt        As String               ' �����
Public TATS         As String               ' �˻�ҿ�ð�
Public ArletFg      As String               ' Arlet Check����('0':��,'1':��)
Public ArletFrVal   As Double               ' Arlet Range(From)
Public ArletToVal   As Double               ' Arlet Range(To)


Public SpcName As String        ' ��ü��

'���뺯�� - �����ϴ� ������Ʈ�κ��� ������
'Public MyOraSE As Object     'OraSession
'Public MyOraDB As Object     'OraDatabase

'Public MySe As Object
'Public DbConn As Object



Public Sub GetStoreCd(ByRef lstList As Object)
   
   Dim objSqlMasters    As clsLISSqlMasters
   Dim tmpRs            As Recordset
   Dim SqlStmt          As String
   Dim tmpStr           As String
   Dim tmpTestCd        As String
   Dim tmpTestNm        As String
   Dim i                As Integer
   
   Set objSqlMasters = New clsLISSqlMasters
   SqlStmt = objSqlMasters.SqlStoreCd
   Set objSqlMasters = Nothing
   
   Set tmpRs = New Recordset
   tmpRs.Open SqlStmt, dbconn
   
   DoEvents
   With lstList
      medLockWindowUpdate (.hWnd)
      .Clear
      .AddItem ""
      While (Not tmpRs.EOF)
         .AddItem "" & tmpRs.Fields("cdval1").Value & "   " & tmpRs.Fields("field1").Value    '���������ڵ� ����Ʈ
         DoEvents
         tmpRs.MoveNext
      Wend
      '.Visible = False
      medLockWindowUpdate (0&)
   End With
   Set tmpRs = Nothing
End Sub

Public Sub GetBuildings(ByRef lstList As Object)
   
   Dim objSqlMasters    As clsLISSqlMasters
   Dim tmpRs            As Recordset
   Dim SqlStmt          As String
   Dim tmpStr           As String
   Dim tmpTestCd        As String
   Dim tmpTestNm        As String
   Dim i                As Integer
   Set objSqlMasters = New clsLISSqlMasters
   SqlStmt = objSqlMasters.SqlBuildings
   Set objSqlMasters = Nothing
   
   Set tmpRs = New Recordset
   tmpRs.Open SqlStmt, dbconn
   
   DoEvents
   With lstList
      medLockWindowUpdate (.hWnd)
      .Clear
      While (Not tmpRs.EOF)
         .AddItem "" & tmpRs.Fields("cdval1").Value & "   " & tmpRs.Fields("field1").Value    '���������ڵ� ����Ʈ
         DoEvents
         tmpRs.MoveNext
      Wend
      '.Visible = False
      medLockWindowUpdate (0&)
   End With
   Set tmpRs = Nothing

End Sub

Public Sub GetSpcGrp(ByRef lstList As Object)
   
   Dim objSqlMasters    As clsLISSqlMasters
   Dim tmpRs            As Recordset
   Dim SqlStmt          As String
   Dim tmpStr           As String
   Dim tmpTestCd        As String
   Dim tmpTestNm        As String
   Dim i                As Integer
   
   Set objSqlMasters = New clsLISSqlMasters
   SqlStmt = objSqlMasters.SqlSpcGrp
   Set objSqlMasters = Nothing
   
   Set tmpRs = New Recordset
   tmpRs.Open SqlStmt, dbconn
   
   DoEvents
   With lstList
      medLockWindowUpdate (.hWnd)
      .Clear
      While (Not tmpRs.EOF)
         .AddItem "" & tmpRs.Fields("cdval1").Value & "   " & tmpRs.Fields("field1").Value    '��ü�� ����Ʈ
         DoEvents
         tmpRs.MoveNext
      Wend
      '.Visible = False
      medLockWindowUpdate (0&)
   End With
   Set tmpRs = Nothing
End Sub

'% Method 1 : SpcQuery
'%                 Parameter�� ���� Sql�� �����ϰ�, �� �ʵ��� ����
'%                 Ŭ���� clsSpecimen�� Data Attribute�� �����Ѵ�.

Public Function SpcQuery(ByVal strTestCd As String, _
                        ByVal strSpcCd As String, _
                        Optional ByVal strApplyDt As Variant) As Boolean

    Dim objSqlMasters   As clsLISSqlMasters
    Dim MyRs            As Recordset
    Dim SqlStmt         As String
    
    Set objSqlMasters = New clsLISSqlMasters
    If IsMissing(strApplyDt) Then
        SqlStmt = objSqlMasters.SqlSpcQuery(strTestCd, strSpcCd)
    Else
        SqlStmt = objSqlMasters.SqlSpcQuery(strTestCd, strSpcCd, strApplyDt)
    End If
    Set objSqlMasters = Nothing
    
    MyRs.Open SqlStmt, dbconn
   
    With MyRs
      
        '.MoveFirst
           
        TestCd = "" & .Fields("TestCd").Value
        SpcCd = "" & .Fields("SpcCd").Value
        Seq = Val("" & .Fields("Seq").Value)
        ApplyDt = "" & .Fields("ApplyDt").Value
        SpcGrpCd = "" & .Fields("SpcGrpCd").Value
        LabelCnt = Val("" & .Fields("LabelCnt").Value)
        RstUnit = "" & .Fields("RstUnit").Value
        RndFg = "" & .Fields("RndFg").Value
        StatFg = "" & .Fields("StatFg").Value
        StatFlags = "" & .Fields("StatFlags").Value
        AvalVal = Val("" & .Fields("AvalVal").Value)
        PanicFg = "" & .Fields("PanicFg").Value
        PanicFrVal = Val("" & .Fields("PanicFrVal").Value)
        PanicToVal = Val("" & .Fields("PanicToVal").Value)
        DeltaFg = "" & .Fields("DeltaFg").Value
        DeltaVal1 = Val("" & .Fields("DeltaVal").Value)
        DeltaVal2 = Val("" & .Fields("DeltaVal2").Value)
        TestCost = "" & .Fields("TestCost").Value
        StoreCd = "" & .Fields("StoreCd").Value
        TatAvg = Val("" & .Fields("TatAvg").Value)
        SpcQty = Val("" & .Fields("SpcQty").Value)
        SpcUnit = "" & .Fields("SpcUnit").Value
        ExpDt = "" & .Fields("ExpDt").Value
        TATS = "" & .Fields("TATS").Value
        
    End With
   
    Set MyRs = Nothing
End Function
   
'% Method 2 : SpcInsert
'%                 Ŭ���� clsSpecimen�� Data Attribute�� ����� ����
'%                 Database�� �����Ѵ�.

Public Function SpcInsert() As Boolean

    Dim objSqlMasters   As clsLISSqlMasters
    Dim objTLab004      As clsTLab004
    Dim SqlStmt         As String
    
    Set objSqlMasters = New clsLISSqlMasters
    Set objTLab004 = New clsTLab004
    Call DataMove(objTLab004)
    SqlStmt = objSqlMasters.SqlSpcInsert(objTLab004)
    Set objTLab004 = Nothing
    Set objSqlMasters = Nothing
   
   
On Error GoTo Err_Trap
   
   dbconn.BeginTrans
   dbconn.Execute (SqlStmt)   'Sql ����
   dbconn.CommitTrans
   SpcInsert = True
   
   '25000������ ������ü���� ����..
'   Call Trans25000(SqlStmt)
   
   Exit Function
   
Err_Trap:
   SpcInsert = False
   dbconn.RollbackTrans
   MsgBox Err.Description, vbExclamation
End Function
   
'% Method 3 : SpcUpdate
'%                 Ŭ���� clsSpecimen�� Data Attribute�� ����� ����
'%                 Database�� Update�Ѵ�.

Public Function SpcUpdate() As Boolean

    Dim objSqlMasters   As clsLISSqlMasters
    Dim objTLab004      As clsTLab004
    Dim SqlStmt         As String
    
    Set objSqlMasters = New clsLISSqlMasters
    Set objTLab004 = New clsTLab004
    Call DataMove(objTLab004)
    SqlStmt = objSqlMasters.SqlSpcUpdate(objTLab004)
    Set objTLab004 = Nothing
    Set objSqlMasters = Nothing
  
   
On Error GoTo Err_Trap
   
   dbconn.BeginTrans
   dbconn.Execute (SqlStmt)   'Sql ����
   dbconn.CommitTrans
   SpcUpdate = True
   
   '25000������ ������ü���� ����..
'   Call Trans25000(SqlStmt)
   
   Exit Function
   
Err_Trap:
   SpcUpdate = False
   dbconn.RollbackTrans
   MsgBox Err.Description, vbExclamation
End Function
   
'% Method 4 : SpcDelete
'%                 Ŭ���� clsSpecimen�� TestCd, ApplyDt �� Key�� Data�� Delete�Ѵ�.

Public Function SpcDelete() As Boolean

    Dim objSqlMasters   As clsLISSqlMasters
    Dim objTLab004      As clsTLab004
    Dim SqlStmt         As String
    
    Set objSqlMasters = New clsLISSqlMasters
    Set objTLab004 = New clsTLab004
    Call DataMove(objTLab004)
    SqlStmt = objSqlMasters.SqlSpcDelete(objTLab004)
    Set objTLab004 = Nothing
    Set objSqlMasters = Nothing
   
   
On Error GoTo Err_Trap
   
   dbconn.BeginTrans
   dbconn.Execute (SqlStmt)   'Sql ����
   dbconn.CommitTrans
   SpcDelete = True
   
   '25000������ ������ü���� ����..
'   Call Trans25000(SqlStmt)
   
   Exit Function
   
Err_Trap:
   SpcDelete = False
   dbconn.RollbackTrans
   MsgBox Err.Description, vbExclamation
End Function
   
Public Function CheckCostCd(ByVal pCostCd As String) As Boolean
    Dim objSqlMasters   As clsLISSqlMasters
    Dim SqlStmt         As String
    Dim tmpRs           As Recordset
    
    Set objSqlMasters = New clsLISSqlMasters
    SqlStmt = objSqlMasters.SqlCheckCostCd(pCostCd)
    Set objSqlMasters = Nothing
    
    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, dbconn
    
    If tmpRs.EOF Then
        CheckCostCd = False
    Else
        CheckCostCd = True
    End If
    Set tmpRs = Nothing
    
End Function

Private Sub DataMove(ByRef objTLab004 As clsTLab004)
    With objTLab004
        .TestCd = TestCd
        .SpcCd = SpcCd
        .ApplyDt = ApplyDt
        .Seq = Seq
        .LabelCnt = LabelCnt
        .RstUnit = RstUnit
        .RndFg = RndFg
        .StatFg = StatFg
        .StatFlags = StatFlags
        .AvalVal = AvalVal
        .PanicFg = PanicFg
        .PanicFrVal = PanicFrVal
        .PanicToVal = PanicToVal
        .ArletFg = ArletFg
        .ArletFrVal = ArletFrVal
        .ArletToVal = ArletToVal
        .DeltaFg = DeltaFg
        .DeltaVal1 = DeltaVal1
        .TestCost = TestCost
        .StoreCd = StoreCd
        .TatAvg = TatAvg
        .SpcQty = SpcQty
        .SpcUnit = SpcUnit
        .ExpDt = ExpDt
        .DeltaVal2 = DeltaVal2
        .TATS = TATS
        
    End With
End Sub