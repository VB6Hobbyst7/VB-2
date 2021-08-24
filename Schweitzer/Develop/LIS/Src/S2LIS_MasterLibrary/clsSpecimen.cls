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
'|  1.  Class   명  : clsSpecimen
'|  2.  기 능         : 지정검체 마스터(LAB004) 로부터 해당 검사아이템의 검체정보를 검색
'|  3. 작성자        : 김미경
'|  4. 작성일        : 1999.05.07
'|
'|  CopyRight(C) 1999 대련엠티에스
'+--------------------------------------------------------------------------------------+


'%  클래스 clsSpecimen의 Data Attributes

Public TestCd       As String               ' 검사항목 코드
Public SpcCd        As String               ' 검체 코드
Public Seq          As Integer              ' Seq
Public ApplyDt      As String               ' 적용일
Public SpcGrpCd     As String               ' 검체군 코드
Public LabelCnt     As Integer              ' 바코드라벨 출력장수
Public RstUnit      As String               ' 결과 단위
Public RndFg        As String               ' Round 채혈여부('0':부,'1':여)
Public StatFlags    As String               ' 각 건물별 Stat 여부
Public StatFg       As String               ' Stat 여부('0':부,'1':여)
Public AvalVal      As Integer              ' 유효 숫자
Public PanicFg      As String               ' Panic Check여부('0':부,'1':여)
Public PanicFrVal   As Double               ' Panic Range(From)
Public PanicToVal   As Double               ' Panic Range(To)
Public DeltaFg      As String               ' Delta Check여부('0':부,'1':여)
Public DeltaVal1    As Integer              ' Delta Value
Public DeltaVal2    As Integer              ' Delta Value
Public TestCost     As String               ' 수가코드
Public StoreCd      As String               ' 보관구분
Public TatAvg       As Integer              ' TAT소요시간
Public SpcQty       As Double               ' 검체 체취량(TLA)
Public SpcUnit      As String               ' 검체 채취 단위
Public ExpDt        As String               ' 폐기일
Public TATS         As String               ' 검사소요시간
Public ArletFg      As String               ' Arlet Check여부('0':부,'1':여)
Public ArletFrVal   As Double               ' Arlet Range(From)
Public ArletToVal   As Double               ' Arlet Range(To)


Public SpcName As String        ' 검체명

'공용변수 - 참조하는 프로젝트로부터 설정됨
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
         .AddItem "" & tmpRs.Fields("cdval1").Value & "   " & tmpRs.Fields("field1").Value    '보관구분코드 리스트
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
         .AddItem "" & tmpRs.Fields("cdval1").Value & "   " & tmpRs.Fields("field1").Value    '보관구분코드 리스트
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
         .AddItem "" & tmpRs.Fields("cdval1").Value & "   " & tmpRs.Fields("field1").Value    '검체군 리스트
         DoEvents
         tmpRs.MoveNext
      Wend
      '.Visible = False
      medLockWindowUpdate (0&)
   End With
   Set tmpRs = Nothing
End Sub

'% Method 1 : SpcQuery
'%                 Parameter로 받은 Sql을 실행하고, 각 필드의 값을
'%                 클래스 clsSpecimen의 Data Attribute에 저장한다.

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
'%                 클래스 clsSpecimen의 Data Attribute에 저장된 값을
'%                 Database로 저장한다.

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
   dbconn.Execute (SqlStmt)   'Sql 실행
   dbconn.CommitTrans
   SpcInsert = True
   
   '25000서버로 지정검체내역 전달..
'   Call Trans25000(SqlStmt)
   
   Exit Function
   
Err_Trap:
   SpcInsert = False
   dbconn.RollbackTrans
   MsgBox Err.Description, vbExclamation
End Function
   
'% Method 3 : SpcUpdate
'%                 클래스 clsSpecimen의 Data Attribute에 저장된 값을
'%                 Database로 Update한다.

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
   dbconn.Execute (SqlStmt)   'Sql 실행
   dbconn.CommitTrans
   SpcUpdate = True
   
   '25000서버로 지정검체내역 전달..
'   Call Trans25000(SqlStmt)
   
   Exit Function
   
Err_Trap:
   SpcUpdate = False
   dbconn.RollbackTrans
   MsgBox Err.Description, vbExclamation
End Function
   
'% Method 4 : SpcDelete
'%                 클래스 clsSpecimen의 TestCd, ApplyDt 를 Key로 Data를 Delete한다.

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
   dbconn.Execute (SqlStmt)   'Sql 실행
   dbconn.CommitTrans
   SpcDelete = True
   
   '25000서버로 지정검체내역 전달..
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
