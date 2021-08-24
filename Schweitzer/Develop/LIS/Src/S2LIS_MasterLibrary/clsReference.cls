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
'|  1.  Class   명  : clsReference
'|  2.  기 능         : 기준치 마스터(LAB005) 로부터 해당 검사아이템의 기준치정보를 검색
'|  3. 작성자        : 김미경
'|  4. 작성일        : 1999.05.07
'|
'|  CopyRight(C) 1999 대련엠티에스
'+--------------------------------------------------------------------------------------+


'%  클래스 clsReference의 Data Attributes

Public TestCd       As String           ' 검사항목 코드
Public SpcCd        As String           ' 검체 코드
Public ApplyDt      As String           ' 적용일
Public ApplySex     As String           ' 적용성별
Public AgeFrom      As Long             ' 적용나이(From)
Public AgeTo        As Long             ' 적용나이(To)
Public RefValFrom   As Double           ' 참고치(From)
Public RefValTo     As Double           ' 참고치(To)
Public ARefValFrom  As Double           ' AutoVerify 참고치(From)
Public ARefValTo    As Double           ' AutoVerify 참고치(To)
Public PanicFrVal   As Double           ' Panic 참고치(From)
Public PanicToVal   As Double           ' Panic 참고치(To)
Public ArletFrVal   As Double           ' Arlet 참고치(From)
Public ArletToVal   As Double           ' Arlet 참고치(To)
Public AMRFrVal     As Double           ' Panic 참고치(From)
Public AMRToVal     As Double           ' Panic 참고치(From)
Public RefCd        As String           ' 참고치(Alpha)
Public RefText      As String           ' 다중참고치(Alpha)
Public ExpDt        As String           ' 폐기일
Public RefDiv       As String           ' 참고치 구분(1: 일반 참고치, 2: panic 참고치)

'Private objSqlMasters

'% Method 1 : RefQuery
'%            Parameter로 받은 Sql을 실행하고, 각 필드의 값을
'%            클래스 clsReference의 Data Attribute에 저장한다.

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
'%            클래스 clsReference의 Data Attribute에 저장된 값을
'%            Database로 저장한다.

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
    DBConn.Execute (SqlStmt)   'Sql 실행
    DBConn.CommitTrans
    RefInsert = True
    Exit Function
   
Err_Trap:
    RefInsert = False
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation

End Function
   
'% Method 3 : RefUpdate
'%            클래스 clsReference의 Data Attribute에 저장된 값을
'%            Database로 Update한다.

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
    DBConn.Execute (SqlStmt)   'Sql 실행
    DBConn.CommitTrans
    RefUpdate = True
    Exit Function
   
Err_Trap:
    RefUpdate = False
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation

End Function
   
'% Method 4 : RefDelete
'%            클래스 clsReference의 TestCd, ApplyDt 를 Key로 Data를 Delete한다.

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
    DBConn.Execute (SqlStmt)   'Sql 실행
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
