VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDonorOkNot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function QueryJudgeList(ByVal Frdt As String, ByVal Todt As String, Optional reservedid As String = "") As Recordset
    Dim sSql As String
    Dim tmpRmk3 As String
    
'    If DBConn.Whatsthis = DBConn.ThisIsSybase Then
    If ObjSysInfo.dbtype = 1 Then
        tmpRmk3 = "convert(VarChar(255), d.rmk3) rmk3"
    Else
        tmpRmk3 = "d.rmk3 rmk3"
    End If

    sSql = " SELECT distinct  a.orddt,d.okdiv3,a.donorid,c.donornm,c.abo,c.rh,b.bldsrc,b.bldyy,b.bldno," & _
           " b.compocd,b.volumn,b.donoraccdt," & tmpRmk3 & _
           " FROM " & T_BBS603 & " d," & T_BBS601 & " c," & T_BBS602 & " b," & T_BBS605 & " a" & _
           " WHERE " & DBW("a.orddt >=", Frdt) & _
           " AND " & DBW("a.orddt<=", Todt) & _
           " AND a.donorid=b.donorid" & _
           " AND a.donoraccdt=b.donoraccdt" & _
           " AND a.donorid=d.donorid" & _
           " AND a.donoraccdt=d.donoraccdt" & _
           " AND a.donorid=c.donorid"
         
         
    If reservedid <> "" Then
        sSql = sSql & _
               " AND " & DBW("b.reservedid=", reservedid)
    End If
    sSql = sSql & " ORDER BY a.orddt,a.donorid,c.donornm"
    Set QueryJudgeList = New Recordset
    QueryJudgeList.Open sSql, DBConn
End Function

Public Function GetDonorBlood(ByVal Donorid As String, ByVal donoraccdt As String) As String
    Dim sSql As String
    
    sSql = " SELECT * FROM " & T_BBS401 & _
           " WHERE " & _
                     DBW("donorid=", Donorid) & _
           " AND " & DBW("donoraccdt=", donoraccdt)
    GetDonorBlood = sSql

End Function
Public Function Get_OkNotList(ByVal Frdt As String, ByVal Todt As String, Optional reservedid As String = "") As Recordset
'bbs506에서 헌혈자 적격 부적격 리스트를 조회한다.
'관련 테이블(헌혈자검사의뢰내역,헌혈자 접수내역,처방바디,결과내역,검사항목 마스터)
'조회데이터(헌혈자명,id,혈액번호,혈액형,제제 검사의뢰일,검사결과,검사명, 검사단위,)
    Dim sSql As String
    
    
    
    sSql = " SELECT a.orddt,g.donoraccdt,a.donorid,b.bldsrc,b.bldyy,b.bldno,b.compocd,b.volumn," & _
           " c.ordcd,d.rstcd,d.testcd,d.spccd,d.rstunit,d.rsttype,e.abbrnm5,f.abo,f.rh,f.donornm,f.dob,f.sex,g.okdiv3" & _
           " FROM " & T_BBS605 & " a," & T_BBS602 & " b," & T_LAB102 & " c," & _
                    T_LAB302 & " d," & T_LAB001 & " e," & T_BBS601 & " f," & T_BBS603 & " g" & _
           " WHERE" & _
                        DBW("a.orddt>=", Frdt) & " AND " & DBW("a.orddt<=", Todt) & _
                  " AND a.donorid=b.donorid" & _
                  " AND a.donoraccdt=b.donoraccdt" & _
                  " AND a.donorid=f.donorid" & _
                  " AND a.workarea=c.workarea" & _
                  " AND a.accdt=c.accdt" & _
                  " AND a.accseq=c.accseq" & _
                  " AND " & DBJ("c.workarea*=d.workarea") & _
                  " AND " & DBJ("c.accdt*=d.accdt") & _
                  " AND " & DBJ("c.accseq*=d.accseq") & _
                  " AND " & DBJ("c.ordcd*=d.testcd") & _
                  " AND " & DBJ("c.spccd*=d.spccd") & _
                  " AND c.ordcd=e.testcd" & _
                  " AND b.donorid=g.donorid" & _
                  " AND b.donoraccdt=g.donoraccdt" & _
                  " AND e.applydt=(SELECT max(z.applydt) " & _
                                  "FROM " & T_LAB001 & " z " & _
                                  "WHERE z.testcd=c.ordcd " & _
                                  "AND   z.applydt<=c.orddt) "
    
    If reservedid <> "" Then
        sSql = sSql & _
               " AND " & DBW("b.reservedid=", reservedid)
    End If
    sSql = sSql & " ORDER BY a.orddt"
    
    Set Get_OkNotList = New Recordset
    Get_OkNotList.Open sSql, DBConn

End Function
Public Function GetOkDiv3(ByVal Donorid As String, ByVal donoraccdt As String) As String
    Dim Rs As Recordset
    Dim sSql As String
    
    sSql = " SELECT okdiv3 FROM " & T_BBS603 & _
           " WHERE " & DBW("donorid=", Donorid) & " AND " & DBW("donoraccdt=", donoraccdt)
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    If Not Rs.EOF Then
        Select Case Rs.Fields("okdiv3").Value & ""
            Case "1": GetOkDiv3 = "적격"
            Case "0": GetOkDiv3 = "부적격"
            Case Else: GetOkDiv3 = "보류"
        End Select
    End If
    Set Rs = Nothing
End Function

Public Sub Reason_List(ByRef cbolist As Object)
   Dim Rs       As Recordset
   Dim i        As Integer
   Dim strReaCd As String
   Dim strReaNm As String

   Set Rs = New Recordset
   Set Rs = GetCom003RecordSet(BC2_DONOR_REASON)
   
   If Rs Is Nothing Then Exit Sub
    
   With cbolist
      .Clear
      .AddItem "판정사유선택"
      While Rs.EOF = False
         
         strReaNm = Rs.Fields("field1").Value & ""
         strReaCd = Rs.Fields("cdval1").Value & ""
         
         .AddItem strReaCd & vbTab & strReaNm
         
         Rs.MoveNext
      Wend
   End With
   
   
   Set Rs = Nothing
End Sub
Public Function Get_Judge_Reason_List() As Recordset
    Dim sSql As String
    
    sSql = "SELECT cdval1, field1 FROM " & T_COM003 & _
               " WHERE  " & DBW("cdindex", BC2_DONOR_REASON, 2)
    
    Set Get_Judge_Reason_List = New Recordset
    Get_Judge_Reason_List.Open sSql, DBConn
    
End Function
Public Function Set_DonorTestSave(ByVal Donorid As String, ByVal donoraccdt As String, _
                                  ByVal okfg As String, ByVal okdt As String, _
                                  ByVal rmk As String) As String
   
    Set_DonorTestSave = " update " & T_BBS603 & _
                        " set " & DBW("okdiv3", okfg, 3) & _
                                  DBW("okdt3", okdt, 3) & _
                                  DBW("rmk3", rmk, 3) & _
                                  DBW("stscd", DonorStatus.stsFinish, 2) & _
                        " WHERE " & _
                                       DBW("donorid", Donorid, 2) & _
                             " AND " & DBW("donoraccdt", donoraccdt, 2)
    

End Function
Public Function Delete_Rsncd(ByVal Donorid As String, ByVal donoraccdt As String) As String
'사유내역을 지운후 Update 해준다.

    Delete_Rsncd = " delete " & T_BBS607 & _
                   " WHERE " & DBW("donorid", Donorid, 2) & " AND  " & DBW("donoraccdt", donoraccdt, 2)
End Function
Public Function Set_MainRsncd(ByVal Donorid As String, ByVal donoraccdt As String, ByVal rsncd As String) As String
'주사유내역을 저장한다.

    Set_MainRsncd = " insert into " & T_BBS607 & "(donorid,donoraccdt,seq,rsncd)" & _
                " values(" & _
                DBV("donorid", Donorid, 1) & DBV("donoraccdt", donoraccdt, 1) & DBV("seq", 1, 1) & DBV("rsncd", rsncd) & ")"
End Function
Public Function Set_SubRsncd(ByVal Donorid As String, ByVal donoraccdt As String, _
                             ByVal rsncd As String, ByVal seq As Long) As String
'세부사유내역을 저장한다.

    Set_SubRsncd = " insert into " & T_BBS607 & "(donorid,donoraccdt,seq,rsncd)" & _
                " values(" & _
                DBV("donorid", Donorid, 1) & DBV("donoraccdt", donoraccdt, 1) & DBV("seq", seq, 1) & DBV("rsncd", rsncd) & ")"
End Function

Public Function GetWorkareaAccdtAccSeq(ByVal Donorid As String, ByVal donoraccdt As String) As Recordset
    Dim sSql As String
    
    sSql = " SELECT workarea,accdt,accseq FROM " & T_BBS605 & _
           " WHERE " & _
                     DBW("donorid=", Donorid) & " AND " & DBW("donoraccdt=", donoraccdt)
    
    Set GetWorkareaAccdtAccSeq = New Recordset
    GetWorkareaAccdtAccSeq.Open sSql, DBConn
End Function
Public Function GetTestResult(ByVal WorkArea As String, ByVal accdt As String, ByVal accseq As String) As Recordset
    Dim sSql As String
    
    sSql = " SELECT a.testcd,a.rstcd,a.rstunit,b.testnm,b.panelfg" & _
           " FROM " & T_LAB001 & " b," & T_LAB302 & " a" & _
           " WHERE " & _
                     DBW("a.workarea=", WorkArea) & _
           " AND " & DBW("a.accdt=", accdt) & _
           " AND " & DBW("a.accseq=", accseq) & _
           " AND a.testcd=b.testcd" & _
           " AND b.applydt=(SELECT max(z.applydt) FROM " & T_LAB001 & " z " & _
                           "WHERE z.testcd=a.testcd " & _
                           "AND " & DBW("z.applydt<=", Format(GetSystemDate, PRESENTDATE_FORMAT)) & ")" '& _
           " ORDER BY"
           
    
    '" AND " & DBW("a.rstdiv=", "R")
    Set GetTestResult = New Recordset
    GetTestResult.Open sSql, DBConn
End Function
Public Function GetResultCdNm(ByVal Wa As String, ByVal accdt As String, ByVal accseq As String) As String
    Dim sSql As String
    'Dim Rs   As String
    Dim Rs    As Recordset
    sSql = " SELECT a.ordcd ,b.testnm " & _
          " FROM " & T_LAB001 & " b," & T_LAB102 & " a" & _
          " WHERE " & _
                  DBW("a.workarea=", Wa) & _
          " AND " & DBW("a.accdt=", accdt) & _
          " AND " & DBW("a.accseq=", accseq) & " AND " & DBW("b.panelfg=", "G") & _
          " AND a.ordcd=b.testcd"
          
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    If Not Rs.EOF Then
        GetResultCdNm = Rs.Fields("ordcd").Value & "" & COL_DIV & Rs.Fields("testnm").Value & ""
    Else
        GetResultCdNm = ""
    End If
    Set Rs = Nothing
    
    
End Function

Public Function GetReason(ByVal Donorid As String, ByVal donoraccdt As String) As Recordset
    Dim sSql As String
    sSql = "SELECT seq,rsncd FROM " & T_BBS607 & " WHERE " & DBW("donorid=", Donorid) & " AND " & DBW("donoraccdt=", donoraccdt)
    Set GetReason = New Recordset
    GetReason.Open sSql, DBConn
    
End Function

