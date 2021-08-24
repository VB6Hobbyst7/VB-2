VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'+--------------------------------------------------------------------------------------+
'|  1.  Class   ��  : clsReference
'|  2.  �� ��         : ����ġ ������(LAB005) �κ��� �ش� �˻�������� ����ġ������ �˻�
'|  3. �ۼ���        : ��̰�
'|  4. �ۼ���        : 1999.05.07
'|
'|  CopyRight(C) 1999 ��ÿ�Ƽ����
'+--------------------------------------------------------------------------------------+


'%  Ŭ���� clsReference�� Data Attributes

Public TestCd       As String           ' �˻��׸� �ڵ�
Public SpcCd        As String           ' ��ü �ڵ�
Public ApplyDt      As String           ' ������
Public ApplySex     As String           ' ���뼺��
Public AgeFrom      As Long             ' ���볪��(From)
Public AgeTo        As Long             ' ���볪��(To)
Public RefValFrom   As Double           ' ����ġ(From)
Public RefValTo     As Double           ' ����ġ(To)
Public ARefValFrom  As Double           ' AutoVerify ����ġ(From)
Public ARefValTo    As Double           ' AutoVerify ����ġ(To)
Public PanicFrVal   As Double           ' Panic ����ġ(From)
Public PanicToVal   As Double           ' Panic ����ġ(To)
Public ArletFrVal   As Double           ' Arlet ����ġ(From)
Public ArletToVal   As Double           ' Arlet ����ġ(To)
Public AMRFrVal     As Double           ' Panic ����ġ(From)
Public AMRToVal     As Double           ' Panic ����ġ(From)
Public RefCd        As String           ' ����ġ(Alpha)
Public RefText      As String           ' ��������ġ(Alpha)
Public ExpDt        As String           ' �����
Public RefDiv       As String           ' ����ġ ����(1: �Ϲ� ����ġ, 2: panic ����ġ)

'Private objSqlMasters

'% Method 1 : RefQuery
'%            Parameter�� ���� Sql�� �����ϰ�, �� �ʵ��� ����
'%            Ŭ���� clsReference�� Data Attribute�� �����Ѵ�.

Public Function RefQuery(ByVal strTestCd As String, _
                                      ByVal strSpcCd As String, _
                                      ByVal strApplyDt As String, _
                                      ByVal strApplySex As String, _
                                      ByVal intAgeFrom As Integer, _
                                      ByVal intAgeTo As Integer) As Boolean

    Dim objSqlMasters   As clsLISSqlMasters
    Dim MyRs            As Recordset
    Dim SqlStmt         As String
    
    Set objSqlMasters = New clsLISSqlMasters
    SqlStmt = objSqlMasters.SqlRefQuery(strTestCd, strSpcCd, strApplyDt, strApplySex, _
                                        intAgeFrom, intAgeTo)
    Set objSqlMasters = Nothing
   
    Set MyRs = New Recordset
    MyRs.Open SqlStmt, DBConn
    
    If MyRs.EOF Then GoTo NoData
   
    With MyRs
       
        .MoveFirst
           
        TestCd = "" & .Fields("TestCd").Value
        SpcCd = "" & .Fields("SpcCd").Value
        ApplyDt = "" & .Fields("ApplyDt").Value
        ApplySex = "" & .Fields("ApplySex").Value
        AgeFrom = "" & .Fields("AgeFrom").Value
        AgeTo = "" & .Fields("AgeTo").Value
'        AgeDiv = "" & .Fields("AgeDiv").Value
        RefValFrom = "" & .Fields("RefValFrom").Value
        RefValTo = "" & .Fields("RefValTo").Value
        RefCd = "" & .Fields("RefCd").Value
        ExpDt = "" & .Fields("ExpDt").Value
       
    End With

NoData:
    Set MyRs = Nothing
   
End Function
   
'% Method 2 : RefInsert
'%            Ŭ���� clsReference�� Data Attribute�� ����� ����
'%            Database�� �����Ѵ�.

Public Function RefInsert() As Boolean

    Dim objSqlMasters   As clsLISSqlMasters
    Dim objTLab005      As clsTLab005
    Dim SqlStmt         As String
    
    Set objSqlMasters = New clsLISSqlMasters
    Set objTLab005 = New clsTLab005
    
    Call DataMove(objTLab005)
    
    SqlStmt = objSqlMasters.SqlRefInsert(objTLab005)
    
    Set objTLab005 = Nothing
    Set objSqlMasters = Nothing
    
On Error GoTo Err_Trap
   
    DBConn.BeginTrans
    DBConn.Execute (SqlStmt)   'Sql ����
    DBConn.CommitTrans
    RefInsert = True
    Exit Function
   
Err_Trap:
    RefInsert = False
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation

End Function
   
'% Method 3 : RefUpdate
'%            Ŭ���� clsReference�� Data Attribute�� ����� ����
'%            Database�� Update�Ѵ�.

Public Function RefUpdate() As Boolean

    Dim objSqlMasters   As clsLISSqlMasters
    Dim objTLab005      As clsTLab005
    Dim SqlStmt         As String
    
    Set objSqlMasters = New clsLISSqlMasters
    Set objTLab005 = New clsTLab005
    Call DataMove(objTLab005)
    SqlStmt = objSqlMasters.SqlRefUpdate(objTLab005)
    Set objTLab005 = Nothing
    Set objSqlMasters = Nothing
   
   
On Error GoTo Err_Trap
   
    DBConn.BeginTrans
    DBConn.Execute (SqlStmt)   'Sql ����
    DBConn.CommitTrans
    RefUpdate = True
    Exit Function
   
Err_Trap:
    RefUpdate = False
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation

End Function
   
'% Method 4 : RefDelete
'%            Ŭ���� clsReference�� TestCd, ApplyDt �� Key�� Data�� Delete�Ѵ�.

Public Function RefDelete(Optional blnDelAll As Boolean = False) As Boolean

    Dim objTLab005      As clsTLab005
    Dim objSqlMasters   As clsLISSqlMasters
    Dim SqlStmt As String
    
    Set objSqlMasters = New clsLISSqlMasters
    Set objTLab005 = New clsTLab005
    Call DataMove(objTLab005)
    SqlStmt = objSqlMasters.SqlRefDelete(objTLab005, blnDelAll)
    Set objTLab005 = Nothing
    Set objSqlMasters = Nothing
   
On Error GoTo Err_Trap
   
    DBConn.BeginTrans
    DBConn.Execute (SqlStmt)   'Sql ����
    DBConn.CommitTrans
    RefDelete = True
    Exit Function
   
Err_Trap:
    RefDelete = False
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation

End Function
   
Private Sub DataMove(ByRef objTLab005 As clsTLab005)
    With objTLab005
        .TestCd = TestCd
        .SpcCd = SpcCd
        .ApplyDt = ApplyDt
        .ApplySex = ApplySex
        .AgeFrom = AgeFrom
        .AgeTo = AgeTo
        .RefValFrom = RefValFrom
        .RefValTo = RefValTo
        .ARefValFrom = ARefValFrom
        .ARefValTo = ARefValTo
        .RefCd = RefCd
        .RefText = RefText
        .ExpDt = ExpDt
        .PanicFrVal = PanicFrVal
        .PanicToVal = PanicToVal
        .ArletFrVal = ArletFrVal
        .ArletToVal = ArletToVal
        .AMRFrVal = AMRFrVal
        .AMRToVal = AMRToVal
        .RefDiv = RefDiv
    End With
End Sub