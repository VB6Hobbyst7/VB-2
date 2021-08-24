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
'|  1.  Class   명  : clsOrders
'|  2.  기 능         : 처방내역(HIS101, HIS102)을 생성한다.
'|  3. 작성자        : 김미경
'|  4. 작성일        : 1999.06.01
'|  5. 수정일        : 2000.12
'|  6. 수정자        : Legends
'|  7. 수정 사유     : BBS의 헌혈자 검사의뢰가 임상병리 시스템의 처방과 다른바 BBS 검사의뢰에
'|                     맞게 수정함.
'|  CopyRight(C) 1999 대련엠티에스
'+--------------------------------------------------------------------------------------+


'%  클래스 clsOrder의 Data Attributes

Public PtId As String           ' 환자ID
Public orddt As String         ' 처방일
Public Bussdiv As String     ' 업무구분
Public BedIndt As String      ' 입원일
Public DeptCd As String       ' 진료과
Public Orddoct As String     ' 처방의
Public MajDoct  As String    ' 주치의
Public entid  As String        ' 입력자
Public EntDt As String          ' 입력일
Public EntTm  As String        ' 입력시간
Public DoneFg As String
Public OrdDiv As String         ' 처방구분

Public OrgAccNo As String    ' 원접수번호
Public SpOrdDiv As String     ' 특수처방구분
Public RepeatFg As String
Public Receptno As String     '영수증번호

'나중에 제거..
Public wardid As String
Public RoomId As String
Public HosilID As String

'= 처방 Body ='
Private Type OrderBody
   OrdSeq As Long     ' 처방 Seq
   OrdCd As String        ' 처방코드
   SpcCd As String        ' 검체코드
   StoreCd As String      ' 보관구분
   DcFg As String          ' DC 여부
   DcDt As String           ' DC 일자
   DcNo As Long        ' DC 처방번호
   Attrcd As String         ' 속성여부
   ExamDt As String       ' 실제 실시일자
   ExamTm As String      ' 실제 실시시간
   ExamDoct As String   ' 실제 실시의사
   stscd As String        ' Status
   statfg As String        ' 응급여부
   InsDiv As String        ' 급여구분
   DoneFg As String
End Type

'= 처방 Header ='
Private Type OrderHeader
   OrdNo As Long      ' 처방번호
   Reqdt As String        ' 실시요청일자
   ReqTm  As String      ' 실시요청시간
   OrdBody() As OrderBody
End Type


'공용변수 - 참조하는 프로젝트로부터 설정됨
'Public MyOraSE As Object     'OraSession

Private OrdHeader() As OrderHeader
Private LastOrdNo As Long
Private sqlStmtH() As String
Private sqlStmtB() As String

Private objMySQL As New clsBBSSQLStatement
'Private DbConn As DrDatabase

Private mvarBuildingNo As String 'BuildingNo를 가져 오기 위한 프라퍼티.
Private mvarDateTime As Date '오늘 날짜을 넣는다.
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
''디비 커넥션
'
'    Set DbConn = pDbConn
'End Sub

'% 처방코드를 기준으로 지정검체 리스트를  띄운다.
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
         tmpStatFg = medGetP("" & .Fields("StatFlags").Value, 1, ";")   '건물별 응급가능 여부
         tmpTestFg = medGetP("" & .Fields("StatFlags").Value, 2, ";")  '건물별 검사가능 여부
   
         tmpStr = ""
         tmpStr = tmpStr & Trim(.Fields("SpcCd").Value & "") & Chr(9)   '검체코드
         tmpStr = tmpStr & Trim(.Fields("SpcNm").Value & "") & Chr(9)   '검체명
         'tmpStr = tmpStr & Trim(.Fields("StatFg").value & "") & Chr(9)    '응급여부
         tmpStr = tmpStr & Mid(tmpStatFg, BuildingNo, 1) & Chr(9)   '**응급여부(해당건물)
         tmpStr = tmpStr & Trim(.Fields("StoreCd").Value & "") & Chr(9)  '보관구분
         tmpStr = tmpStr & Trim(.Fields("MultiFg").Value & "") & Chr(9)   '복수검체여부
         tmpStr = tmpStr & Trim(.Fields("SpcGrp").Value & "") & Chr(9)   '검체군
         tmpStr = tmpStr & Trim(.Fields("LabelCnt").Value & "") & Chr(9) '라벨출력장수
         tmpStr = tmpStr & Mid(tmpTestFg, BuildingNo, 1) & Chr(9)   '**검사가능여부(해당건물)
         tmpStr = tmpStr & Trim(.Fields("Seq").Value & "")                     '우선순위
         paraSpcList.AddItem tmpStr
         tmpRs.MoveNext
      Wend
   End With
NoData:
'   tmpRs.RsClose
   Set tmpRs = Nothing

End Sub

'% 처방Header/Body 내역을 클래스로 Assign한다.
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
         '.Col = 12: KeyString = KeyString & .Value & Chr(9)   '검사구분  --> 제외 1999.10.08 by 김미경
         .Col = 9:  KeyString = KeyString & .Value & Chr(9)   'WorkArea
         .Col = 4:  KeyString = KeyString & .Value & Chr(9)   '검체코드
         .Col = 10: KeyString = KeyString & .Value & Chr(9)   'StoreCd
         .Col = 6:  KeyString = KeyString & CStr(Val(.Value)) & Chr(9)   '응급여부
         .Col = 7:  KeyString = KeyString & .Value                '희망채취시간
         
         If SaveKeyString <> KeyString Then
            HSeq = NextSeq(1)
            SaveKeyString = KeyString
            .Col = 7: OrdHeader(HSeq).Reqdt = Format(.Value, CS_DateDbFormat)
                         OrdHeader(HSeq).ReqTm = Format(.Value, CS_TimeDbFormat)
         End If
         
         'OrdHeader(HSeq).OrdNo = HSeq    '처방Header Seq
         .Col = 1: .Value = HSeq
         BSeq = NextSeq(2, HSeq)    '처방Body Seq
         .Col = 3: OrdHeader(HSeq).OrdBody(BSeq).OrdCd = .Value   '처방코드
         .Col = 4: OrdHeader(HSeq).OrdBody(BSeq).SpcCd = .Value   '검체코드
         .Col = 10: OrdHeader(HSeq).OrdBody(BSeq).StoreCd = .Value  '보관구분
                      OrdHeader(HSeq).OrdBody(BSeq).DcFg = ""           'DC 여부
                      OrdHeader(HSeq).OrdBody(BSeq).DcDt = ""            'DC 일자
                      OrdHeader(HSeq).OrdBody(BSeq).DcNo = 0            'DC 번호
                      OrdHeader(HSeq).OrdBody(BSeq).Attrcd = ""          '속성여부
                      OrdHeader(HSeq).OrdBody(BSeq).ExamDt = ""        '실제 실시일자
                      OrdHeader(HSeq).OrdBody(BSeq).ExamTm = ""       '실제 실시시간
                      OrdHeader(HSeq).OrdBody(BSeq).ExamDoct = ""      '실제 실시의사
                      OrdHeader(HSeq).OrdBody(BSeq).stscd = "0"            'Status
                      OrdHeader(HSeq).OrdBody(BSeq).DoneFg = "0"          'Status
         .Col = 5: OrdHeader(HSeq).OrdBody(BSeq).InsDiv = .Value                   '급여구분
         .Col = 6: OrdHeader(HSeq).OrdBody(BSeq).statfg = CStr(Val(.Value))    '응급여부
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


'% 처방Header/Body 내역의 Array Bound를 재조정한다.
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

'% 클래스의 처방Header/Body 내역을 DB로 저장한다.
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
       sqlStmtH(i) = CreateSqlH(i)   '처방 Header Insert Sql문 생성
       For j = 1 To BCnt
'            Debug.Print "Ubound(sqlstmtb) : " & UBound(sqlStmtB)
          ReDim Preserve sqlStmtB(UBound(sqlStmtB) + 1)
          
          sqlStmtB(UBound(sqlStmtB)) = CreateSqlB(i, j)  '처방 Body Insert Sql문 생성
'          Debug.Print "UBound(sqlStmtB) = " & UBound(sqlStmtB) & " : " & "sqlStmtB(" & UBound(sqlStmtB) & ")" & " = " & sqlStmtB(UBound(sqlStmtB))
          
       Next
    Next
    StartOrdNo = LastOrdNo
    CreateSqlStmt = True
    Exit Function
   
ErrCreateSQLStmt:
    CreateSqlStmt = False
End Function

'% 가장 마지막 처방번호를 가져온다.
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
