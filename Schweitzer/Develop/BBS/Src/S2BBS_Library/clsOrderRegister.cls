VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBBS501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'######## frmBBS501 에서 사용하였슴 #######

Public Function GetPopup(Optional ByVal busidiv As Integer = -1, Optional ByVal blnPtid As Boolean = True) As String
    'Popup String
    If busidiv = 0 Then      '환자 리스트
        If blnPtid = True Then
            GetPopup = "select a.ptid,b.ptnm from  " & T_HIS002 & " a," & T_HIS001 & " b" _
                      & " where a.ptid=b.ptid"
        Else
            GetPopup = "select ptid,ptnm from " & T_HIS001
        End If
    ElseIf busidiv = 1 Then  '처방의
        GetPopup = "select doctid,doctnm from " & T_HIS005
    ElseIf busidiv = 2 Then  '진료과
        GetPopup = "select deptcd,deptnm from " & T_HIS003
    ElseIf busidiv = 3 Then  '병실
        GetPopup = "select wardid,roomid+'-'+bedid from " & T_HIS004
    ElseIf busidiv = 4 Then
        GetPopup = "select Empid,EmpLngNm from " & T_COM006
    ElseIf busidiv = 5 Then  '수혈사유
        GetPopup = "select cdval1,text1 from " & T_COM003 & " where cdindex='B015'"
    ElseIf busidiv = 6 Then  '수술명
        GetPopup = "select ocd,onm from " & T_HIS007
    End If

End Function


Public Function GetOrdNO_Lab(ByVal Ptid As String, ByVal orddt As String) As String
    Dim strSql As String
    
    GetOrdNO_Lab = "select max(ordno) as ordno from " & T_LAB101 & _
                   " where ptid='" & Ptid & "' and orddt='" & orddt & "' "

End Function
Public Function Insert_Lab101(dbconn As DrDatabase, ByVal Ptid As String, orddt As String, ordno As Long, _
                             busidiv As String, BedIndt As String, Reqdt As String, deptcd As String, _
                             Orddoct As String, entid As String, dongFg As String, _
                             Receptno As String, WardId As String, hosilid As String, _
                             bedid As String) As Boolean
    '처방정보(헤더) insert
    Dim strSql As String
    Dim strOrdTm As String
    
    strOrdTm = Format(Now, "HHmm") & "00"
    If Receptno = "" Then
       Receptno = "Null"
    Else
        Receptno = CLng(Receptno)
    End If
    
    strSql = "insert into " & T_LAB101 & " (ptid,orddt,ordno,ordtm,Bussdiv,bedindt,reqdt, " & _
                          " reqtm,deptcd,orddoct,entid, " & _
                          " orddiv,Repeatfg,donefg,receptno,wardid,bedid,hosilid) " & _
                          " values (" & _
                          " '" & Ptid & "','" & orddt & "',   " & ordno & ", " & _
                          " '" & strOrdTm & "',   '" & busidiv & "','" & BedIndt & "','" & Reqdt & "', " & _
                          "    '000000',          '" & deptcd & "', '" & Orddoct & "', " & _
                          " '" & entid & "',  '" & c_workarea & "' ,'0','" & dongFg & "', " & _
                          "  " & Receptno & ",'" & WardId & "', " & _
                          " '" & bedid & "', '" & hosilid & "')"
              
On Error GoTo DBExecerror

    dbconn.BeginTrans
    dbconn.Execute (strSql)
    dbconn.CommitTrans
    Insert_Lab101 = True
    
    Exit Function

DBExecerror:
    Insert_Lab101 = False
    dbconn.RollbackTrans
    dbconn.DisplayErrors
End Function
Public Function Insert_Lab102(dbconn As DrDatabase, ByVal Ptid As String, orddt As String, ordno As Long, _
                             Ordseq As Long, OrdCd As String, status As String, _
                             statfg As String, dongFg As String, Receptno As String, _
                             UnitQty As Long) As Boolean
    '처방정보(내역) insert
    Dim strSql As String
    
    If Receptno = "" Then
       Receptno = "Null"
    Else
        Receptno = CLng(Receptno)
    End If
    strSql = "insert into " & T_LAB102 & " (ptid,orddt,ordno,Ordseq,OrdCd,stscd, " & _
                          " statfg,donefg,Receptno,UnitQty) " & _
                          " values (" & _
                          "'" & Ptid & "', '" & orddt & "', " & ordno & ", " & _
                          " " & Ordseq & ",'" & OrdCd & "','" & status & "', " & _
                          " '" & statfg & "', '" & dongFg & "', " & _
                          " " & Receptno & "," & UnitQty & ")"
             
On Error GoTo DBExecerror

    dbconn.BeginTrans
    dbconn.Execute (strSql)
    dbconn.CommitTrans
    Insert_Lab102 = True
    
    Exit Function

DBExecerror:
    Insert_Lab102 = False
    dbconn.RollbackTrans
    dbconn.DisplayErrors
End Function
Public Function Insert_Lab102_Collect(dbconn As DrDatabase, ByVal Ptid As String, orddt As String, ordno As Long, _
                             Ordseq As Long, OrdCd As String, stscd As String, _
                             statfg As String, dongFg As String, Receptno As String, _
                             UnitQty As Long, Accdt As String, accseq As Long, _
                             rcvdt As String, ByVal IrradFg As String, ByVal FilterFg As String) As Boolean
    '처방정보(내역) insert 처방 접수 버튼
    Dim strSql As String
    
    If Receptno = "" Then
       Receptno = "Null"
    Else
        Receptno = CLng(Receptno)
    End If
    strSql = "insert into " & T_LAB102 & " (ptid,orddt,ordno,Ordseq,OrdCd,stscd, " & _
                          " statfg,donefg,Receptno,UnitQty,workarea,accdt,accseq,rcvdt,rcvtm,irradfg,filterfg) " & _
                          " values (" & _
                          "'" & Ptid & "', '" & orddt & "', " & ordno & ", " & _
                          " " & Ordseq & ",'" & OrdCd & "','" & stscd & "', " & _
                          " '" & statfg & "', '" & dongFg & "', " & _
                          "  " & Receptno & "," & UnitQty & ",'" & c_workarea & "'," & _
                          " '" & Accdt & "',   " & accseq & ", '" & rcvdt & "','000000'," & _
                          " '" & IrradFg & "','" & FilterFg & "')"
             
On Error GoTo DBExecerror

    dbconn.BeginTrans
    dbconn.Execute (strSql)
    dbconn.CommitTrans
    Insert_Lab102_Collect = True
    
    Exit Function

DBExecerror:
    Insert_Lab102_Collect = False
    dbconn.RollbackTrans
    dbconn.DisplayErrors
End Function
Public Function Insert_Lab103(dbconn As DrDatabase, ByVal Ptid As String, orddt As String, ordno As Long, _
                             Attrcd As String, AttrVal As String) As Boolean
    '처방정보(속성) insert
    Dim strSql As String
    
    strSql = "insert into " & T_LAB103 & " (ptid,orddt,ordno,Attrcd,AttrVal) " & _
                          " values (" & _
                          "'" & Ptid & "',  '" & orddt & "'," & ordno & ", " & _
                          "'" & Attrcd & "','" & AttrVal & "')"
             
On Error GoTo DBExecerror

    dbconn.BeginTrans
    dbconn.Execute (strSql)
    dbconn.CommitTrans
    Insert_Lab103 = True
    
    Exit Function

DBExecerror:
    Insert_Lab103 = False
    dbconn.RollbackTrans
    dbconn.DisplayErrors
End Function
Public Function Delete_Order(dbconn As DrDatabase, ByVal Ptid As String, orddt As String, ordno As Long, ByVal TF As Boolean)
    'TF=true;lab101,TF=true;lab103
    Dim strSql As String
    If TF = True Then
        strSql = "delete * from " & T_LAB101 & " where ptid='" & Ptid & "', and orddt='" & orddt & "' and ordno='" & ordno & "'"
    Else
        strSql = "delete * from " & T_LAB103 & " where ptid='" & Ptid & "', and orddt='" & orddt & "' and ordno='" & ordno & "'"
    End If
    
On Error GoTo DBExecerror
    dbconn.BeginTrans
    dbconn.Execute (strSql)
    dbconn.CommitTrans
    
    Exit Function
DBExecerror:
    dbconn.RollbackTrans
    dbconn.DisplayErrors
End Function
Public Function GetBBS001() As DrRecordSet
    '수혈처방을 불러온다.
    Dim strSql As String
    
    strSql = "select testnm,testcd from " & T_BBS001
    Set GetBBS001 = New DrRecordSet
    GetBBS001.RsOpen , strSql
    
End Function
Public Function GetHIS007() As DrRecordSet
    Dim strSql As String
    
    strSql = "select ocd,onm from " & T_HIS007
    Set GetHIS007 = New DrRecordSet
    
    GetHIS007.RsOpen , strSql
End Function
Public Function Insert_BBS201(dbconn As DrDatabase, spcyy As String, spcno As String, coldt As String, _
                            coltm As String, colid As String, rcvdt As String, rcvtm As String, _
                            rcvid As String, storeleg As String, storerno As String, storecno As String, _
                            Ptid As String, busidiv As String, BedIndt As String, _
                            orddt As String, deptcd As String, buildcd As String, _
                            orgbuildcd As String) As Boolean
    Dim strSql As String
    
    strSql = "insert into " & T_BBS201 & "( " & _
           "spcyy,spcno,coldt,coltm,colid,rcvdt,rcvtm,rcvid,storeleg,storerno,storecno," & _
           "ptid,busidiv,bedindt,orddt,deptcd,buildcd,orgbuildcd) " & _
           "values (" & _
           "'" & spcyy & "', " & spcno & ",'" & coldt & "','" & coltm & "',   " & colid & "," & _
           "'" & rcvdt & "','" & rcvtm & "'," & rcvid & ", '" & storeleg & "'," & storerno & "," & _
           " " & storecno & ",'" & Ptid & "',  '" & busidiv & "','" & BedIndt & "'," & _
           "'" & orddt & "',  '" & deptcd & "','" & buildcd & "','" & orgbuildcd & "')"
           
On Error GoTo DBExecerror

    dbconn.BeginTrans
    dbconn.Execute (strSql)
    dbconn.CommitTrans
    Insert_BBS201 = True
    
    Exit Function

DBExecerror:
    Insert_BBS201 = False
    dbconn.RollbackTrans
    dbconn.DisplayErrors
End Function

