VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBldDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private stsORDER     As String
Private stsCOLLECT   As String
Private stsACCESS    As String
Private stsINPROCESS As String
Private mvarWorkSeq  As String

Public Property Let WorkSeq(ByVal vData As String)
    mvarWorkSeq = vData
End Property

Public Function GetBloodList(ByVal mode As Integer) As Recordset
    Dim sSql As String
    
    sSql = ""
    Select Case mode
        Case 0: '출고
            sSql = " SELECT distinct a.bldsrc,a.bldyy,a.bldno,e.ptid,a.compocd,b.componm " & _
                   " FROM " & T_BBS302 & " a, " & T_BBS006 & " b, " & T_BBS401 & " c, " & _
                             " " & T_BBS202 & " e " & _
                   " WHERE (" & DBW("a.rstv", "1", 2) & " or " & DBW("a.stat", "1", 2) & ") " & _
                   " AND (a.cancelfg='' or a.cancelfg is null or a.cancelfg='0')" & _
                   " AND b.compocd=a.compocd " & _
                   " AND a.bldsrc=c.bldsrc AND a.bldyy=c.bldyy AND a.bldno=c.bldno " & _
                   " AND a.workarea=e.workarea AND a.accdt=e.accdt AND a.accseq=e.accseq " & _
                   " " & _
                   " AND a.compocd=c.compocd " & _
                   " AND " & DBW("c.stscd", BBSBloodStatus.stsASSIGN, 2)
                   
        Case 1, 2, 3: '반환
            sSql = "SELECT distinct a.bldsrc,a.bldyy,a.bldno,e.ptid,d." & F_PTNM & ",a.compocd,b.componm " & _
                   "FROM " & T_BBS402 & " a, " & T_BBS006 & " b, " & T_BBS401 & " c, " & _
                             T_HIS001 & " d, " & T_BBS202 & " e " & _
                   "WHERE " & _
                   "    b.compocd=c.compocd " & _
                   "AND a.bldsrc=c.bldsrc AND a.bldyy=c.bldyy AND a.bldno=c.bldno " & _
                   "AND a.workarea=e.workarea AND a.accdt=e.accdt AND a.accseq=e.accseq " & _
                   "AND e.ptid=d." & F_PTID & " " & _
                   "AND a.compocd=c.compocd " & _
                   "AND " & DBW("a.deliverydt=", Format(GetSystemDate, "yyyymmdd")) & _
                   "AND " & DBW("c.stscd", BBSBloodStatus.stsDELIVERY, 2) & " AND (a.retfg='' or a.retfg is null)"
                   
    End Select
    
    Set GetBloodList = New Recordset
    Call GetBloodList.Open(sSql, DBConn)
'    If GetBloodList.DBerror = True Then
'        dbconn.DisplayErrors
'        Set GetBloodList = Nothing
'    End If
End Function

Public Function GetSQLBloodList(ByVal mode As Long) As String
    Dim sSql As String
    
    sSql = ""
    Select Case mode
        Case 0: '출고
            sSql = " SELECT DISTINCT a.bldsrc" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "a.bldyy" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "trim(" & FUNC_CONVERT("char", "a.bldno", "000000") & ")" & " AS bloodno, " & _
                    " e.ptid,d." & F_PTNM & ",a.compocd" & FUNC_CONCAT & "'" & COL_DIV & "'" & FUNC_CONCAT & "b.componm AS compo " & _
                    " FROM " & T_BBS302 & " a," & T_BBS202 & " e," & T_BBS006 & " b," & T_BBS401 & " c," & T_HIS001 & " d " & _
                    " WHERE a.workarea = e.workarea " & _
                    " AND a.accdt=e.accdt " & _
                    " AND a.accseq=e.accseq " & _
                    " AND a.bldsrc=c.bldsrc " & _
                    " AND a.bldyy=c.bldyy " & _
                    " AND a.bldno=c.bldno " & _
                    " AND a.compocd=c.compocd " & _
                    " AND ( a.rstv  ='1' OR  a.stat  ='1') " & _
                    " AND (a.cancelfg='' OR a.cancelfg IS NULL OR a.cancelfg='0') " & _
                    " AND a.compocd=b.compocd " & _
                    " AND e.ptid=d.patno " & _
                    " AND " & DBW("c.stscd", BBSBloodStatus.stsASSIGN, 2)
        Case 1, 2, 3: '반환,폐기
            sSql = " SELECT DISTINCT a.bldsrc" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "a.bldyy" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "trim(" & FUNC_CONVERT("char", "a.bldno", "000000") & ")" & " AS bloodno, " & _
                    " e.ptid,d." & F_PTNM & ",a.compocd" & FUNC_CONCAT & "'" & COL_DIV & "'" & FUNC_CONCAT & "b.componm AS compo " & _
                   " FROM " & T_BBS402 & " a, " & T_BBS006 & " b, " & T_BBS401 & " c, " & _
                             T_HIS001 & " d, " & T_BBS202 & " e " & _
                   " WHERE " & _
                   "    b.compocd=c.compocd " & _
                   " AND a.bldsrc=c.bldsrc AND a.bldyy=c.bldyy AND a.bldno=c.bldno " & _
                   " AND a.workarea=e.workarea AND a.accdt=e.accdt AND a.accseq=e.accseq " & _
                   " AND e.ptid=d." & F_PTID & " " & _
                   " AND a.compocd=c.compocd " & _
                   " and a.deliverydt " & FUNC_CONCAT & " a.deliverytm between '" & Format(DateAdd("h", -24, GetSystemDate), "yyyymmddhhmmss") & "' and '" & Format(GetSystemDate, "yyyymmddhhmmss") & "'" & _
                   "AND " & DBW("c.stscd", BBSBloodStatus.stsDELIVERY, 2) & " AND (a.retfg='' or a.retfg is null)"
'원본: 출고한지 24시간 이내의 것을 반환 받을 수있도록 하기 위해서  변경함
'                   " AND " & DBW("a.deliverydt=", Format(GetSystemDate, "yyyymmdd")) &
'        Case 2 '혈액 폐기(출고 후 반환 가능 시간 여부 체크 안함)
'            sSql = " SELECT DISTINCT a.bldsrc" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "a.bldyy" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "trim(" & FUNC_CONVERT("char", "a.bldno", "000000") & ")" & " AS bloodno, " & _
'                    " e.ptid,d." & F_PTNM & ",a.compocd" & FUNC_CONCAT & "'" & COL_DIV & "'" & FUNC_CONCAT & "b.componm AS compo " & _
'                   " FROM " & T_BBS402 & " a, " & T_BBS006 & " b, " & T_BBS401 & " c, " & _
'                             T_HIS001 & " d, " & T_BBS202 & " e " & _
'                   " WHERE " & _
'                   "    b.compocd=c.compocd " & _
'                   " AND a.bldsrc=c.bldsrc AND a.bldyy=c.bldyy AND a.bldno=c.bldno " & _
'                   " AND a.workarea=e.workarea AND a.accdt=e.accdt AND a.accseq=e.accseq " & _
'                   " AND e.ptid=d." & F_PTID & " " & _
'                   " AND a.compocd=c.compocd " & _
'                   " AND " & DBW("c.stscd", BBSBloodStatus.stsDELIVERY, 2) & " AND (a.retfg='' or a.retfg is null)"
        Case 4 '혈액은행 자체 폐기용 SQL
            sSql = " SELECT DISTINCT a.bldsrc" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "a.bldyy" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "trim(" & FUNC_CONVERT("char", "a.bldno", "000000") & ")" & " AS bloodno, " & _
                  " '' ptid, '' " & F_PTNM & ", a.compocd" & FUNC_CONCAT & "'" & COL_DIV & "'" & FUNC_CONCAT & "b.componm AS compo " & _
                  " FROM " & T_BBS401 & " a," & T_BBS006 & " b" & _
                  " WHERE " & _
                  "     a.stscd in(" & _
                                  DBV("stscd", BBSBloodStatus.stsENTER, 1) & _
                                  DBV("stscd", BBSBloodStatus.stsRETURN, 1) & _
                                  DBV("stscd", BBSBloodStatus.stsASSIGN) & ")" & _
                  " AND a.compocd=b.compocd"
    End Select
    'bloodno, ptid, ptnm, compo
    GetSQLBloodList = sSql
End Function

Public Function GetExpireHospital(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String) As Recordset
    Dim sSql As String
    
    sSql = " SELECT a.bldsrc,a.bldyy,a.bldno,a.compocd,b.componm,a.stscd " & _
           " FROM " & T_BBS401 & " a," & T_BBS006 & " b" & _
           " WHERE " & _
                             DBW("a.bldsrc", Bldsrc, 2) & _
                   " AND " & DBW("a.bldyy", Bldyy, 2) & _
                   " AND " & DBW("a.bldno", Bldno, 2) & _
                   " AND a.compocd=b.compocd " & _
                   " AND a.stscd in(" & _
                                  DBV("stscd", BBSBloodStatus.stsENTER, 1) & _
                                  DBV("stscd", BBSBloodStatus.stsRETURN, 1) & _
                                  DBV("stscd", BBSBloodStatus.stsASSIGN) & ")"
                          
    Set GetExpireHospital = New Recordset
    Call GetExpireHospital.Open(sSql, DBConn)
'    If GetExpireHospital.DBerror = True Then
'        dbconn.DisplayErrors
'        Set GetExpireHospital = Nothing
'    End If
    
End Function

Public Function GetCmdExpireHospital() As Recordset
    Dim sSql As String
    
    sSql = " SELECT distinct a.bldsrc,a.bldyy,a.bldno,a.compocd,b.componm,a.abo,a.rh,a.volumn,a.entdt,a.enttm " & _
          " FROM " & T_BBS401 & " a," & T_BBS006 & " b" & _
          " WHERE " & _
          "     a.stscd in(" & _
                          DBV("stscd", BBSBloodStatus.stsENTER, 1) & _
                          DBV("stscd", BBSBloodStatus.stsRETURN, 1) & _
                          DBV("stscd", BBSBloodStatus.stsASSIGN) & ")" & _
          " AND a.compocd=b.compocd"
    
    Set GetCmdExpireHospital = New Recordset
    Call GetCmdExpireHospital.Open(sSql, DBConn)
'    If GetCmdExpireHospital.DBerror = True Then
'        dbconn.DisplayErrors
'        Set GetCmdExpireHospital = Nothing
'    End If
End Function

Public Function GetBloodCompoList(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal mode As Integer) As Recordset
    '-----------------------------------------------------------------------------------
    'mode = 0 : 혈액출고
    '           XM되어있고(OK이거나 응급), 아직 출고되지 않은 혈액 또 취소되지 않은 혈액
    '       1 : 혈액반환
    '           출고되어있고, 반환이나 폐기, 회수되지 않은 혈액
    '       2 : 혈액폐기
    '           출고되어있고, 반환이나 폐기, 회수되지 않은 혈액
    '       3 : BAG회수
    '           출고되어있고, 반환이나 폐기, 회수되지 않은 혈액
    '-----------------------------------------------------------------------------------
    Dim sSql As String
    
    Select Case mode
        Case 0: '출고
            sSql = " SELECT distinct a.bldsrc,a.bldyy,a.bldno,a.compocd,b.componm " & _
                   " FROM " & T_BBS302 & " a, " & T_BBS006 & " b, " & T_BBS401 & " c " & _
                   " WHERE " & _
                                     DBW("a.bldsrc", Bldsrc, 2) & _
                           " AND " & DBW("a.bldyy", Bldyy, 2) & _
                           " AND " & DBW("a.bldno", Bldno, 2) & _
                           " AND (" & DBW("a.rstv=", "1") & " or" & DBW(" a.stat=", "1") & " or " & DBW("a.norstfg=", "1") & ") " & _
                           " AND (a.cancelfg='' or a.cancelfg is null or a.cancelfg='0')" & _
                           " AND b.compocd=a.compocd " & _
                           " AND a.bldsrc=c.bldsrc " & _
                           " AND a.bldyy=c.bldyy " & _
                           " AND a.bldno=c.bldno " & _
                           " AND a.compocd=c.compocd " & _
                           " AND " & DBW("c.stscd", BBSBloodStatus.stsASSIGN, 2)
                   
        Case 1, 2, 3: '반환
            sSql = " SELECT distinct a.bldsrc,a.bldyy,a.bldno,a.compocd,b.componm " & _
                   " FROM " & T_BBS302 & " a, " & T_BBS006 & " b, " & T_BBS401 & " c " & _
                   " WHERE " & _
                                     DBW("a.bldsrc", Bldsrc, 2) & _
                           " AND " & DBW("a.bldyy", Bldyy, 2) & _
                           " AND " & DBW("a.bldno", Bldno, 2) & _
                           " AND (" & DBW("a.rstv=", "1") & " or " & DBW("a.stat=", "1") & " or " & DBW("a.norstfg=", "1") & ") " & _
                           " AND (" & DBW("a.cancelfg<>", "1") & " or a.cancelfg is null) " & _
                           " AND b.compocd=a.compocd " & _
                           " AND a.bldsrc=c.bldsrc " & _
                           " AND a.bldyy=c.bldyy " & _
                           " AND a.bldno=c.bldno " & _
                           " AND a.compocd=c.compocd " & _
                           " AND " & DBW("c.stscd", BBSBloodStatus.stsDELIVERY, 2)
    End Select
    
    Set GetBloodCompoList = New Recordset
    Call GetBloodCompoList.Open(sSql, DBConn)
'    If GetBloodCompoList.DBerror = True Then
'        dbconn.DisplayErrors
'        Set GetBloodCompoList = Nothing
'    End If
End Function

Public Function GetXMExpireHospital(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal CompoCd As String) As Recordset
'입력된 혈액의 xm검사여부를 확인한다.
    Dim sSql   As String
    
    sSql = "SELECT  * FROM " & T_BBS302 & _
           " WHERE" & _
                  "     " & DBW("bldsrc", Bldsrc, 2) & _
                  " AND " & DBW("bldyy", Bldyy, 2) & _
                  " AND " & DBW("bldno", Bldno, 2) & _
                  " AND " & DBW("compocd", CompoCd, 2)
                  
    Set GetXMExpireHospital = New Recordset
    Call GetXMExpireHospital.Open(sSql, DBConn)
'    If GetXMExpireHospital.DBerror = True Then
'        dbconn.DisplayErrors
'        Set GetXMExpireHospital = Nothing
'    End If
End Function

Public Function GetXMExpireOrdInfo(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal CompoCd As String) As Recordset
'입력된 혈액의 처방정보를 가지고 온다.
    Dim sSql As String
    
    sSql = "SELECT j.ptid,j.orddt,j.ordno,j.ordseq,a.rstseq,a.bldsrc,a.bldyy,a.bldno,a.compocd,a.vfydt,a.vfytm,a.stat," & _
           "b.volumn,b.abo,b.rh,b.stscd," & _
           "c.reqdt,c.reqtm,c.deptcd,c.wardid,c.hosilid,c.bedid," & _
           "d.ordcd,d.accdt,d.accseq,d.mesg,d.unitqty,d.irradfg,d.filterfg," & _
           "e." & F_PTNM & " as ptnm," & F_SSN2("e") & " as ssn," & _
           "f.testnm," & _
           "g." & F_DEPTNM & " as deptnm" & _
    " FROM " & T_BBS302 & " a," & T_BBS401 & " b," & T_LAB101 & " c," & T_LAB102 & " d," & _
               T_HIS001 & " e," & T_BBS001 & " f," & T_HIS003 & " g," & _
               T_BBS202 & " j " & _
    " WHERE " & DBW("a.bldsrc", Bldsrc, 2) & _
    " AND " & DBW("a.bldyy", Bldyy, 2) & _
    " AND " & DBW("a.bldno", Bldno, 2) & _
    " AND " & DBW("a.compocd", CompoCd, 2) & _
    " AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd " & _
    " AND a.workarea=j.workarea AND a.accdt=j.accdt AND a.accseq=j.accseq " & _
    " AND j.ptid=c.ptid AND j.orddt=c.orddt AND j.ordno=c.ordno " & _
    " AND j.ptid=d.ptid AND j.orddt=d.orddt AND j.ordno=d.ordno AND j.ordseq=d.ordseq " & _
    " AND j.ptid=e." & F_PTID & _
    " AND d.ordcd=f.testcd " & _
    " AND " & DBJ("c.deptcd*=g." & F_DEPTCD) & _
    " AND b.stscd in (" & _
                    DBV("stscd", BBSBloodStatus.stsASSIGN, 1) & _
                    DBV("stscd", BBSBloodStatus.stsENTER, 1) & _
                    DBV("stscd", BBSBloodStatus.stsRETURN) & ")"
    
    Set GetXMExpireOrdInfo = New Recordset
    Call GetXMExpireOrdInfo.Open(sSql, DBConn)
'    If GetXMExpireOrdInfo.DBerror = True Then
'        dbconn.DisplayErrors
'        Set GetXMExpireOrdInfo = Nothing
'    End If
    
End Function

Public Function GetExpireBloodInfo(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal CompoCd As String) As Recordset
'자체폐기시에 혈액정보를 가지고 온다.
    Dim sSql   As String
    
    sSql = " SELECT a.abo,a.rh,a.volumn ,a.compocd,b.componm " & _
           " FROM " & T_BBS401 & " a," & T_BBS006 & " b" & _
           " WHERE" & _
                    "     " & DBW("a.bldsrc", Bldsrc, 2) & _
                    " AND " & DBW("a.bldyy", Bldyy, 2) & _
                    " AND " & DBW("a.bldno", Bldno, 2) & _
                    " AND " & DBW("a.compocd", CompoCd, 2) & _
                    " AND a.compocd=b.compocd"
          
    Set GetExpireBloodInfo = New Recordset
    Call GetExpireBloodInfo.Open(sSql, DBConn)
'    If GetExpireBloodInfo.DBerror = True Then
'        dbconn.DisplayErrors
'        Set GetExpireBloodInfo = Nothing
'    End If
End Function

Public Function GetBloodInfo(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal CompoCd As String, ByVal mode As Integer) As Recordset
    Dim sSql As String
    
    
    'mode = 0 : 출고
    '       1 : 반환
    '       2 : 폐기
    '       3 : 회수
    Select Case mode
        Case 0:
            sSql = "SELECT h.ptid,h.orddt,h.ordno,h.ordseq,a.rstseq,a.bldsrc,a.bldyy,a.bldno,a.compocd,a.vfydt,a.stat,a.rmk,a.filterfg," & _
                          "b.volumn,b.abo,b.rh,b.stscd,b.irrfg,b.entdt,b.available,b.coldt," & _
                          "c.reqdt,c.reqtm,c.deptcd,c.wardid,c.hosilid,c.bedid," & _
                          "d.ordcd,d.accdt,d.accseq,d.mesg,d.unitqty,d.irradfg," & _
                          "f.testnm,f.testdiv,d.newtestdiv,c.bussdiv,d.ocsordno" & _
                   " FROM " & T_BBS302 & " a," & T_BBS401 & " b," & T_LAB101 & " c," & T_LAB102 & " d," & _
                              T_BBS001 & " f," & T_BBS202 & " h " & _
                   " WHERE " & _
                                      DBW("a.bldsrc", Bldsrc, 2) & _
                            " AND " & DBW("a.bldyy", Bldyy, 2) & _
                            " AND " & DBW("a.bldno", Bldno, 2) & _
                            " AND " & DBW("a.compocd", CompoCd, 2) & _
                            " AND (a.cancelfg='' or a.cancelfg is null or a.cancelfg='0') AND (" & DBW("a.rstv", "1", 2) & " or " & DBW("a.stat=", "1") & ")" & _
                            " AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd " & _
                            " AND a.workarea=h.workarea AND a.accdt=h.accdt AND a.accseq=h.accseq " & _
                            " AND h.ptid=c.ptid AND h.orddt=c.orddt AND h.ordno=c.ordno " & _
                            " AND h.ptid=d.ptid AND h.orddt=d.orddt AND h.ordno=d.ordno AND h.ordseq=d.ordseq " & _
                            " AND d.ordcd=f.testcd " & _
                            " AND f.applydt=(SELECT max(z.applydt) FROM " & T_BBS001 & " z " & _
                                                                  "WHERE z.testcd=d.ordcd " & _
                                                                  "AND z.applydt<=d.orddt ) "
        
        Case 1, 2, 3:
            sSql = "SELECT j.ptid,j.orddt,j.ordno,j.ordseq,a.rstseq,a.bldsrc,a.bldyy,a.bldno,a.compocd,a.vfydt,a.stat,a.filterfg," & _
                          "b.volumn,b.abo,b.rh,b.stscd," & _
                          "c.reqdt,c.reqtm,c.deptcd,c.wardid,c.hosilid,c.bedid," & _
                          "d.ordcd,d.accdt,d.accseq,d.mesg,d.unitqty,d.irradfg,d.filterfg,d.newtestdiv,b.irrfg," & _
                          "f.testnm,h.workseq," & _
                          "h.deliverydt,h.deliverytm,h.deliveryseq,h.ordcd newordcd,i.testnm newordnm,d.ocsordno,c.bussdiv " & _
                   " FROM " & T_BBS302 & " a," & T_BBS401 & " b," & T_LAB101 & " c," & T_LAB102 & " d," & _
                              T_BBS001 & " f," & T_BBS402 & " h," & _
                              T_BBS001 & " i," & T_BBS202 & " j " & _
                   " WHERE " & DBW("a.bldsrc", Bldsrc, 2) & _
                   " AND " & DBW("a.bldyy", Bldyy, 2) & _
                   " AND " & DBW("a.bldno", Bldno, 2) & _
                   " AND " & DBW("a.compocd", CompoCd, 2) & _
                   " AND " & DBW("a.cancelfg<>", "1") & _
                   " AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd " & _
                   " AND a.workarea=h.workarea and a.accdt=h.accdt and a.accseq=h.accseq and a.rstseq=h.rstseq  " & _
                   " AND a.workarea=j.workarea AND a.accdt=j.accdt AND a.accseq=j.accseq " & _
                   " AND j.ptid=c.ptid AND j.orddt=c.orddt AND j.ordno=c.ordno " & _
                   " AND j.ptid=d.ptid AND j.orddt=d.orddt AND j.ordno=d.ordno AND j.ordseq=d.ordseq " & _
                   " AND d.ordcd=f.testcd " & _
                   " AND a.bldsrc=h.bldsrc AND a.bldyy=h.bldyy AND a.bldno=h.bldno AND a.compocd=h.compocd " & _
                   " AND " & DBJ("h.ordcd*=i.testcd") & _
                   " ORDER BY  a.vfydt desc"
    End Select
    
    Set GetBloodInfo = New Recordset
    Call GetBloodInfo.Open(sSql, DBConn)
'    If GetBloodInfo.DBerror = True Then
'        dbconn.DisplayErrors
'        Set GetBloodInfo = Nothing
'    End If
End Function

Public Function GetDeliveryHistory(ByVal PtId As String) As Recordset
    Dim sSql As String
    
    sSql = "SELECT a.deliverydt,a.deliverytm,a.bldsrc,a.bldyy,a.bldno,a.compocd,a.retfg,a.expfg,a.irrafg,a.filter," & _
           "       b.abo,b.rh,b.volumn,c.componm " & _
           " FROM " & T_BBS402 & " a," & T_BBS401 & " b," & T_BBS006 & " c," & T_BBS202 & " d " & _
           " WHERE " & _
                          DBW("d.ptid", PtId, 2) & _
                    " AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd " & _
                    " AND a.workarea=d.workarea AND a.accdt=d.accdt AND a.accseq=d.accseq " & _
                    " AND a.compocd=c.compocd " & _
                    " ORDER BY d.ptid,a.deliverydt desc,a.deliverytm desc"
           
    Set GetDeliveryHistory = New Recordset
    GetDeliveryHistory.Open sSql, DBConn
'    If GetDeliveryHistory.DBerror = True Then
'        dbconn.DisplayErrors
'        Set GetDeliveryHistory = Nothing
'    End If
End Function

Public Function BldBag(ByVal Bldsrc As String, _
                       ByVal Bldyy As String, _
                       ByVal Bldno As String, _
                       ByVal CompoCd As String, _
                       ByVal deliverydt As String, _
                       ByVal Deliveryseq As String, _
                       ByVal bagid As String, _
                       ByVal bagrcvid As String) As Boolean
    
    Dim today       As Date
    Dim sSql        As String
    
    Dim DrRS        As Recordset
    Dim strWorkArea As String
    Dim strAccDt    As String
    Dim strAccSeq   As String
    Dim strBagCnt   As String
    
    '--------------------------------------------------
    '처방상태의 bag 회수갯수를 수정하기 위해 갯수를 구한다.
    '--------------------------------------------------
    sSql = " SELECT a.workarea,a.accdt,a.accseq,a.bagcnt " & _
           " FROM  " & T_BBS203 & " a," & T_BBS402 & " b " & _
           " WHERE " & _
                   "a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq " & _
                   " AND " & DBW("b.bldsrc", Bldsrc, 2) & _
                   " AND " & DBW("b.bldyy", Bldyy, 2) & _
                   " AND " & DBW("b.bldno", Bldno, 2) & _
                   " AND " & DBW("b.compocd", CompoCd, 2) & _
                   " AND " & DBW("b.deliverydt", deliverydt, 2) & _
                   " AND " & DBW("b.deliveryseq", Deliveryseq, 2)
              
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
    If DrRS.EOF Then
'        dbconn.DisplayErrors
        Exit Function
    Else
        strWorkArea = DrRS.Fields("workarea").Value & ""
        strAccDt = DrRS.Fields("accdt").Value & ""
        strAccSeq = DrRS.Fields("accseq").Value & ""
        strBagCnt = DrRS.Fields("bagcnt").Value & "" + 1
    End If
    Set DrRS = Nothing
    
    
On Error GoTo BldBag_error
    DBConn.BeginTrans
    
    today = GetSystemDate
    
    sSql = " update " & T_BBS402 & " " & _
           " set " & DBW("bagfg", "1", 3) & _
                     DBW("bagdt", Format(today, PRESENTDATE_FORMAT), 3) & _
                     DBW("bagtm", Format(today, PRESENTTIME_FORMAT), 3) & _
                     DBW("bagid", bagid, 3) & _
                     DBW("bagrcvid", bagrcvid, 2) & _
           " WHERE" & _
                    "     " & DBW("bldsrc", Bldsrc, 2) & _
                    " AND " & DBW("bldyy", Bldyy, 2) & _
                    " AND " & DBW("bldno", Bldno, 2) & _
                    " AND " & DBW("compocd", CompoCd, 2) & _
                    " AND " & DBW("deliverydt", deliverydt, 2) & _
                    " AND " & DBW("deliveryseq", Deliveryseq, 2)
    DBConn.Execute sSql
    
    sSql = "update " & T_BBS401 & " " & _
           "set " & DBW("stscd", BBSBloodStatus.stsBAG, 2) & _
           "WHERE " & _
                           DBW("bldsrc", Bldsrc, 2) & _
                  "AND " & DBW("bldyy", Bldyy, 2) & _
                  "AND " & DBW("bldno", Bldno, 2) & _
                  "AND " & DBW("compocd", CompoCd, 2)
    DBConn.Execute sSql
    
    sSql = " update " & T_BBS203 & " " & _
           " set " & DBW("bagcnt", strBagCnt, 2) & _
           " WHERE " & DBW("workarea=", strWorkArea) & _
           " AND   " & DBW("accdt=", strAccDt) & _
           " AND   " & DBW("accseq=", strAccSeq)
    DBConn.Execute sSql
    
    DBConn.CommitTrans
    
    BldBag = True
    Exit Function
    
BldBag_error:
    DBConn.RollbackTrans
    
    BldBag = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function BldExpire(ByVal Bldsrc As String, _
                          ByVal Bldyy As String, _
                          ByVal Bldno As String, _
                          ByVal CompoCd As String, _
                          ByVal deliverydt As String, _
                          ByVal Deliveryseq As String, _
                          ByVal exprcvid As String, _
                          ByVal expid As String, _
                          ByVal Exprsncd As String, _
                          ByVal expbilldiv As String, _
                          ByVal exprsnrmk As String, _
                          ByVal PtId As String, _
                          ByVal OrdCd As String) As Boolean
                          

'HosExp=true 이면, 자체폐기

    Dim today       As Date
    Dim sSql        As String
    
    Dim strWorkArea As String
    Dim strAccDt    As String
    Dim strAccSeq   As String
    Dim strExpCnt   As String
    Dim DrRS        As Recordset
    Dim strDeliverytm As String


    today = GetSystemDate
    
On Error GoTo BldExpire_error
    DBConn.BeginTrans
    
    '--------------------------------------------------
    '처방상태의 폐기갯수를 수정하기 위해 갯수를 구한다.
    '--------------------------------------------------
    sSql = " SELECT a.workarea,a.accdt,a.accseq,a.expcnt,b.deliverytm " & _
           " FROM " & T_BBS203 & " a," & T_BBS402 & " b " & _
           " WHERE a.workarea=b.workarea" & _
           " AND   a.accdt=b.accdt" & _
           " AND   a.accseq=b.accseq " & _
           " AND   " & DBW("b.bldsrc", Bldsrc, 2) & _
           " AND   " & DBW("b.bldyy", Bldyy, 2) & _
           " AND   " & DBW("b.bldno", Bldno, 2) & _
           " AND   " & DBW("b.compocd", CompoCd, 2) & _
           " AND   " & DBW("b.deliverydt", deliverydt, 2) & _
           " AND   " & DBW("b.deliveryseq", Deliveryseq, 2)
    
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
    If DrRS.EOF Then
'        dbconn.DisplayErrors
        Exit Function
    Else
        strWorkArea = DrRS.Fields("workarea").Value & ""
        strAccDt = DrRS.Fields("accdt").Value & ""
        strAccSeq = DrRS.Fields("accseq").Value & ""
        strExpCnt = Val(DrRS.Fields("expcnt").Value & "") + 1
        strDeliverytm = DrRS.Fields("deliverytm").Value & ""
    End If
    Set DrRS = Nothing
    
    '
    ' 폐기수량 1 증가
    '
    sSql = " update " & T_BBS203 & _
           " set " & DBW("expcnt", strExpCnt, 2) & _
           " WHERE " & DBW("workarea=", strWorkArea) & _
           " AND   " & DBW("accdt=", strAccDt) & _
           " AND   " & DBW("accseq=", strAccSeq)
    DBConn.Execute sSql
    
    sSql = " update " & T_BBS401 & " " & _
           " set " & DBW("realexpdt", Format(today, PRESENTDATE_FORMAT), 3) & _
                     DBW("realexptm", Format(today, PRESENTTIME_FORMAT), 3) & _
                     DBW("exprcvid", exprcvid, 3) & _
                     DBW("expid", expid, 3) & _
                     DBW("exprsncd", Exprsncd, 3) & _
                     DBW("expbilldiv", expbilldiv, 3) & _
                     DBW("exprsnrmk", exprsnrmk, 3) & _
                     DBW("stscd", BBSBloodStatus.stsEXPIRE, 2) & _
           " WHERE " & _
                    "     " & DBW("bldsrc", Bldsrc, 2) & _
                    " AND " & DBW("bldyy", Bldyy, 2) & _
                    " AND " & DBW("bldno", Bldno, 2) & _
                    " AND " & DBW("compocd", CompoCd, 2)
    DBConn.Execute sSql
    
    sSql = " update " & T_BBS402 & " " & _
           " set " & DBW("expfg", BBSExpFg.stsExpfg, 3) & _
                     DBW("expdt", Format(today, PRESENTDATE_FORMAT), 3) & _
                     DBW("exptm", Format(today, PRESENTTIME_FORMAT), 2) & _
           " WHERE" & _
                    "     " & DBW("bldsrc", Bldsrc, 2) & _
                    " AND " & DBW("bldyy", Bldyy, 2) & _
                    " AND " & DBW("bldno", Bldno, 2) & _
                    " AND " & DBW("compocd", CompoCd, 2) & _
                    " AND " & DBW("deliverydt", deliverydt, 2) & _
                    " AND " & DBW("deliveryseq", Deliveryseq, 2)
                    
                    
    DBConn.Execute sSql
    
    
'########################
'전주 예수병원 폐기 처리
'IODATE, BLOODNO, BLDCODE, IOTYPE1, IOTYPE2, PATNO
'출고일자, 혈액번호, 처방코드, 입출고구분, 병동외래구분, 환자아이디
'            " where iodate=" & FUNC_CONVERT("date", deliverydt, "yyyymmdd")

    Dim Rs As Recordset
    
    Set Rs = New Recordset
    
    sSql = " select * from oras1.slxboutt " & _
            " where iodate=" & FUNC_CONVERT("date", deliverydt & strDeliverytm, "yyyymmdd hh24:mi:ss") & _
            " and bloodno='" & Bldsrc & Bldyy & Format(Bldno, String(6, "0")) & "'" & _
            " and bldcode='" & OrdCd & "'" & _
            " and iotype1='2'" & _
            " and patno ='" & PtId & "'"

    Rs.Open sSql, DBConn
    
    If Rs.EOF = False Then
'                " where iodate=" & FUNC_CONVERT("date", deliverydt, "yyyymmdd")
        sSql = " update oras1.slxboutt set iotype1='4' " & _
                " where iodate=" & FUNC_CONVERT("date", deliverydt & strDeliverytm, "yyyymmdd hh24:mi:ss") & _
                " and bloodno='" & Bldsrc & Bldyy & Format(Bldno, String(6, "0")) & "'" & _
                " and bldcode='" & OrdCd & "'" & _
                " and iotype1='2'" & _
                " and patno ='" & PtId & "'"
        
        DBConn.Execute sSql
    Else
        Err.Raise 9999, , "OCS에 폐기처리 내역생성 도중 오류가 발생하였습니다."
    End If
    
    Set Rs = Nothing
'##########################3
    
    
    '----------------------------------------------------------------------
    ' 병원별 수납처리 할때....반환여부 플래그 박아주는곳이다.....
    '----------------------------
    '----------------------------
    ' ★ 수가내역 수납처리
    '----------------------------
'    sSql = " UPDATE " & T_BBS903 & " SET " & DBW("orddiv", "4", 3) & "entdt=" & ConvToDate(Format(today, CS_DateDbFormat)) & "," & _
'                                             DBW("donefg", "0", 3) & DBW("mesg", exprsnrmk) & _
'           " WHERE " & DBW("workseq=", mvarWorkSeq)
'    DBConn.Execute sSql
    

    DBConn.CommitTrans
    
    BldExpire = True
    Exit Function
    
BldExpire_error:
    DBConn.RollbackTrans
    
    BldExpire = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function BldReturnHour() As Long
    Dim sSql As String
    Dim Rs   As Recordset
    
    sSql = "SELECT field1 FROM " & T_COM003 & _
           " WHERE " & _
                             DBW("cdindex", BC2_RETURN_HOUR, 2) & _
                   " AND  cdval1=(SELECT max(cdval1) FROM " & T_COM003 & " WHERE " & DBW("cdindex", BC2_RETURN_HOUR, 2) & " AND " & DBW("cdval1<=", "0") & ")"
    
  
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Rs.EOF Then
        BldReturnHour = 30
    Else
        BldReturnHour = CLng(Rs.Fields("field1").Value & "")
    End If
    
    Set Rs = Nothing
    
End Function

Public Function BldReturn(ByVal Bldsrc As String, _
                          ByVal Bldyy As String, _
                          ByVal Bldno As String, _
                          ByVal CompoCd As String, _
                          ByVal deliverydt As String, _
                          ByVal Deliveryseq As String, _
                          ByVal OrdCd As String, _
                          ByVal PtId As String, _
                          ByVal retid As String, _
                          ByVal retrcvid As String, _
                          ByVal retrmk As String) As Boolean
    
    Dim today       As Date
    Dim sSql        As String
    Dim RetCnt      As String
    Dim DrRS        As Recordset
    Dim strWorkArea As String
    Dim strAccDt    As String
    Dim strAccSeq   As String
    Dim strSQL      As String
    Dim i           As Long
    
    Dim strDeliverytm As String

    '--------------------------------------------------
    '처방상태의 반환갯수를 수정하기 위해 갯수를 구한다.
    '--------------------------------------------------
    sSql = " SELECT a.retcnt, b.deliverytm " & _
           " FROM " & T_BBS203 & " a," & T_BBS402 & " b " & _
           " WHERE" & _
                    " a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq " & _
                    " AND " & DBW("b.bldsrc", Bldsrc, 2) & _
                    " AND " & DBW("b.bldyy", Bldyy, 2) & _
                    " AND " & DBW("b.bldno", Bldno, 2) & _
                    " AND " & DBW("b.compocd", CompoCd, 2) & _
                    " AND " & DBW("b.deliverydt", deliverydt, 2) & _
                    " AND " & DBW("b.deliveryseq", Deliveryseq, 2)
    
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
    If DrRS.EOF Then
        MsgBox "반환처리할 혈액의 정보가 정확하지 않습니다.", vbExclamation
        Exit Function
    End If
    
    If DrRS.RecordCount < 1 Then
        RetCnt = 1
    Else
        RetCnt = Val(DrRS.Fields("retcnt").Value & "") + 1
        
        strDeliverytm = DrRS.Fields("deliverytm").Value & ""
    End If
    Set DrRS = Nothing
    
    
On Error GoTo BldReturn_error
    DBConn.BeginTrans
    
    today = GetSystemDate
    
    sSql = " update " & T_BBS402 & " " & _
           " set " & DBW("retfg", "1", 3) & _
           " " & DBW("retdt", Format(today, PRESENTDATE_FORMAT), 3) & _
           " " & DBW("rettm", Format(today, PRESENTTIME_FORMAT), 3) & _
           " " & DBW("retid", retid, 3) & _
           " " & DBW("retrcvid", retrcvid, 3) & _
           " " & DBW("retrmk", retrmk, 2) & _
           " WHERE" & _
                    "     " & DBW("bldsrc", Bldsrc, 2) & _
                    " AND " & DBW("bldyy", Bldyy, 2) & _
                    " AND " & DBW("bldno", Bldno, 2) & _
                    " AND " & DBW("compocd", CompoCd, 2) & _
                    " AND " & DBW("deliverydt", deliverydt, 2) & _
                    " AND " & DBW("deliveryseq", Deliveryseq, 2)
    
    DBConn.Execute sSql
    
    sSql = " update " & T_BBS401 & " " & _
           " set " & DBW("stscd", BBSBloodStatus.stsRETURN, 2) & _
           " WHERE " & _
                             DBW("bldsrc", Bldsrc, 2) & _
                   " AND " & DBW("bldyy", Bldyy, 2) & _
                   " AND " & DBW("bldno", Bldno, 2) & _
                   " AND " & DBW("compocd", CompoCd, 2)
    DBConn.Execute sSql


    sSql = " SELECT workarea,accdt,accseq,rstseq" & _
             " FROM " & T_BBS402 & _
             " WHERE " & DBW("bldsrc", Bldsrc, 2) & _
             " AND " & DBW("bldyy", Bldyy, 2) & _
             " AND " & DBW("bldno", Bldno, 2) & _
             " AND " & DBW("compocd", CompoCd, 2) & _
             " AND " & DBW("deliverydt", deliverydt, 2) & _
             " AND " & DBW("deliveryseq", Deliveryseq, 2)
    
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
    
    For i = 1 To DrRS.RecordCount
        sSql = " update " & T_BBS203 & " " & _
               " set " & DBW("retcnt", RetCnt, 2) & _
               " WHERE " & DBW("workarea=", DrRS.Fields("workarea").Value & "") & _
               " AND   " & DBW("accdt=", DrRS.Fields("accdt").Value & "") & _
               " AND   " & DBW("accseq=", DrRS.Fields("accseq").Value & "")
               
        DBConn.Execute sSql
        sSql = " update " & T_BBS302 & _
               " set " & DBW("canceldt", Format(today, PRESENTDATE_FORMAT), 3) & _
                         DBW("canceltm", Format(today, PRESENTTIME_FORMAT), 3) & _
                         DBW("cancelid", retid, 3) & _
                         DBW("cancelfg", "1", 2) & _
               " WHERE " & _
                         DBW("workarea=", DrRS.Fields("workarea").Value & "") & _
               " AND " & DBW("accdt=", DrRS.Fields("accdt").Value & "") & _
               " AND " & DBW("accseq=", DrRS.Fields("accseq").Value & "") & _
               " AND " & DBW("rstseq=", DrRS.Fields("rstseq").Value & "")
        strWorkArea = DrRS.Fields("workarea").Value & ""
        strAccDt = DrRS.Fields("accdt").Value & ""
        strAccSeq = DrRS.Fields("accseq").Value & ""
              
        DBConn.Execute sSql
        
        DrRS.MoveNext
    Next i
    
    
'########################
'전주 예수병원 반환 처리
'IODATE, BLOODNO, BLDCODE, IOTYPE1, IOTYPE2, PATNO
'출고일자, 혈액번호, 처방코드, 입출고구분, 병동외래구분, 환자아이디

'2005/06/15 modify by legends
'iodate에 시간까지 넣어 달라고 해서 넣어주는 걸로 수정함

'            " where iodate=" & FUNC_CONVERT("date", deliverydt , "yyyymmdd") & _

    Dim Rs As Recordset
    
    Set Rs = New Recordset
    
    sSql = " select * from oras1.slxboutt " & _
            " where iodate=" & FUNC_CONVERT("date", deliverydt & strDeliverytm, "yyyymmdd hh24:mi:ss") & _
            " and bloodno='" & Bldsrc & Bldyy & Format(Bldno, String(6, "0")) & "'" & _
            " and bldcode='" & OrdCd & "'" & _
            " and iotype1='2'" & _
            " and patno ='" & PtId & "'"

    Rs.Open sSql, DBConn
    
'                " where iodate=" & FUNC_CONVERT("date", deliverydt, "yyyymmdd") & _

    If Rs.EOF = False Then
        sSql = " update oras1.slxboutt set iotype1='3' " & _
                " where iodate=" & FUNC_CONVERT("date", deliverydt & strDeliverytm, "yyyymmdd hh24:mi:ss") & _
                " and bloodno='" & Bldsrc & Bldyy & Format(Bldno, String(6, "0")) & "'" & _
                " and bldcode='" & OrdCd & "'" & _
                " and iotype1='2'" & _
                " and patno ='" & PtId & "'"
        
        DBConn.Execute sSql
    Else
        Err.Raise 9999, , "OCS에 반환처리 내역생성 도중 오류가 발생하였습니다."
    End If
    
    Set Rs = Nothing
'##########################3
    
    '----------------------------------------------------------------------
    ' 병원별 수납처리 할때....반환여부 플래그 박아주는곳이다.....
    '----------------------------
    ' ★ 수가내역 수납처리
    '----------------------------
'    sSql = " UPDATE " & T_BBS903 & " SET " & DBW("orddiv", "1", 3) & "entdt=" & ConvToDate(Format(today, CS_DateDbFormat)) & "," & _
'                                             DBW("donefg", "0", 3) & DBW("mesg", retrmk) & _
'           " WHERE " & DBW("workseq=", mvarWorkSeq)
'    DBConn.Execute sSql
    '------------------------------------------------------------------------
    DBConn.CommitTrans
    BldReturn = True
    Exit Function
    
BldReturn_error:
    DBConn.RollbackTrans
    
    BldReturn = False
    MsgBox Err.Description, vbExclamation
End Function

'Private Function GetOrderKey() As String
'    Dim Rs      As Recordset
'    Dim strSql  As String
'
'    GetOrderKey = ""
'    strSql = " SELECT med_ocs.SEQ_ORDER_KEY.NEXTVAL MAXKEY FROM DUAL "
'    Set Rs = New Recordset
'    Rs.Open strSql, DBConn
'
'    If Not Rs.EOF Then GetOrderKey = Rs.Fields("MAXKEY").Value & ""
'
'    Set Rs = Nothing
'End Function

'Private Function ConvToDate(ByVal argDate As String) As String
'    ConvToDate = "To_Date('" & argDate & "', 'YYYYMMDD') "
'End Function

'Public Function Filter_SuNapSQL(ByVal OcsOrdNo As String, ByVal Bussdiv As String, ByVal BBSKey As String, ByVal CompoCd As String, _
'                                ByVal Filterfg As String, ByVal Mesg As String, ByVal OrdCd As String, ByRef Order_key As String) As String
'    Filter_SuNapSQL = OCS_SuNapSQL(OcsOrdNo, Bussdiv, BBSKey, CompoCd, Filterfg, Mesg, OrdCd, Order_key)
'End Function

'Private Function OCS_SuNapSQL(ByVal OcsOrdNo As String, ByVal Bussdiv As String, _
'                              ByVal BBSKey As String, ByVal CompoCd As String, _
'                              ByVal Filterfg As String, ByVal Mesg As String, _
'                              ByVal OrdCd As String, ByRef Order_key As String) As String
    
'    Dim Rs                  As Recordset
'    Dim sSql                As String
'
'    Dim Patient_No          As String
'    Dim Order_Date          As String
'    Dim Old_Order_Key       As String
'
'    Dim Item_Code           As String
'    Dim Item_Name           As String
'    Dim Qty_Of_Order        As String
'    Dim remark              As String
'    Dim Duration_Of_Order   As String
'    Dim Dosage_Code         As String
'    Dim Night_Day           As String
'    Dim Seq_No              As String
'    Dim Value_Of_Frequency  As String
'    Dim Self                As String
'    Dim Exam_Specimen       As String
'    Dim Stat_Normal         As String
'    Dim Pmpa_Code           As String
'    Dim Order_Dr_Code       As String
'    Dim Class_Of_Order      As String
'    Dim Group_Of_Order      As String
'
'    Dim Write_Date          As String
'    Dim Clinical_dept       As String
'    Dim Order_check         As String
'    Dim Unit_of_order       As String
'    Dim Request_Act_Date    As String
'    Dim Input_status        As String
'    Dim Write_status        As String
'    Dim Order_site          As String
'
'    Dim Contents_Of_Order   As String
'    Dim Dosage_Of_Order     As String
'    Dim Base_Contents       As String
'    Dim PRN                 As String
'    Dim Acting_Check        As String
'    Dim Dr_code             As String
'    Dim Frequency_Of_Order  As String
'    Dim Return_Value        As String
'
'    Dim strFields           As String
'
'    Dim CompoNm             As String
'    Dim OrdDiv              As String
'    Dim EntDt               As String
'    Dim DoneFg              As String
'    Dim strTable            As String
'
'    On Error GoTo Errors
'
'    Old_Order_Key = OcsOrdNo
'    Qty_Of_Order = "1"
'    OrdDiv = "3"
'    EntDt = Format(GetSystemDate, "YYYYMMDD")
'    DoneFg = "0"
'
'    strFields = " ,to_char(write_date,'YYYY-MM-DD HH24:MI:SS') as write_date,clinical_dept,order_check,unit_of_order,input_status,write_status,order_site,"
'
'    '처방정보 테이블 읽어오기
'    If Bussdiv = "1" Then
'        strTable = "med_ocs.opd_order_dmc  "
'        strFields = strFields & " '' as request_act_date"
'        sSql = " select patient_no,to_char(order_date,'yyyymmdd') as orderdate,order_key,item_code,item_name," & _
'               "        remark,duration_of_order,dosage_code,night_day,seq_no,value_of_frequency," & _
'               "        Self , Exam_Specimen, Stat_Normal, ordered_dr_code, Class_Of_Order, Group_Of_Order," & _
'               "        contents_of_order,dosage_of_order,base_contents,prn,acting_check,dr_code,frequency_of_order,return_value " & strFields & _
'               " From " & strTable & _
'               " where order_key=" & OcsOrdNo
'
'    Else
'        strTable = "med_ocs.ipd_order_dmc  "
'        strFields = strFields & " to_char(request_act_date,'yyyymmdd') as request_act_date"
'
'        sSql = " select patient_no,to_char(order_date,'yyyymmdd') as orderdate,order_key,item_code,item_name," & _
'               "        remark,duration_of_order,dosage_code,night_day,seq_no,value_of_frequency," & _
'               "        Self , Exam_Specimen, Stat_Normal, ordered_dr_code, Class_Of_Order, Group_Of_Order," & _
'               "        contents_of_order,dosage_of_order,base_contents,prn,acting_check,dr_code,frequency_of_order,return_value " & strFields & _
'               " From " & strTable & _
'               " where order_key=" & OcsOrdNo
'
'        Set Rs = Nothing
'        Set Rs = New Recordset
'        Rs.Open sSql, DBConn
'
'        If Rs.EOF Then
'            sSql = " select patient_no,to_char(order_date,'yyyymmdd') as orderdate,order_key,item_code,item_name," & _
'                   "        remark,duration_of_order,dosage_code,night_day,seq_no,value_of_frequency," & _
'                   "        Self , Exam_Specimen, Stat_Normal, ordered_dr_code, Class_Of_Order, Group_Of_Order," & _
'                   "        contents_of_order,dosage_of_order,base_contents,prn,acting_check,dr_code,frequency_of_order,return_value " & strFields & _
'                   " From med_ocs.ipd_order_update_dmc " & _
'                   " where order_key=" & OcsOrdNo
'        End If
'
'        Set Rs = Nothing
'
'    End If
'
'    Set Rs = Nothing
'    Set Rs = New Recordset
'
'    Rs.Open sSql, DBConn
'
'    If Not Rs.EOF Then
'        Patient_No = Rs.Fields("patient_no").Value & ""
'        Order_Date = Rs.Fields("orderdate").Value & ""
'        remark = Rs.Fields("remark").Value & ""
'        Duration_Of_Order = Rs.Fields("duration_of_order").Value & ""
'        Dosage_Code = Rs.Fields("dosage_code").Value & ""
'        Night_Day = Rs.Fields("night_day").Value & ""
'        Seq_No = Rs.Fields("seq_no").Value & ""
'        Value_Of_Frequency = Rs.Fields("value_of_frequency").Value & ""
'        Self = Rs.Fields("self").Value & ""
'        Exam_Specimen = Rs.Fields("exam_specimen").Value & ""
'        Stat_Normal = Rs.Fields("stat_normal").Value & ""
'        Order_Dr_Code = Rs.Fields("ordered_dr_code").Value & ""
'        Class_Of_Order = Rs.Fields("class_of_order").Value & ""
'        Group_Of_Order = Rs.Fields("group_of_order").Value & ""
'        Write_Date = Rs.Fields("write_date").Value & ""
'        Clinical_dept = Rs.Fields("clinical_dept").Value & ""
'        Order_check = Rs.Fields("order_check").Value & ""
'        Unit_of_order = Rs.Fields("unit_of_order").Value & ""
'        Request_Act_Date = Rs.Fields("request_act_date").Value & ""
'        Input_status = Rs.Fields("input_status").Value & ""
'        Write_status = Rs.Fields("write_status").Value & ""
'        Order_site = Rs.Fields("order_site").Value & ""
'
'        Contents_Of_Order = Rs.Fields("contents_of_order").Value & ""
'        Dosage_Of_Order = Rs.Fields("dosage_of_order").Value & ""
'        Base_Contents = Rs.Fields("base_contents").Value & ""
'        PRN = Rs.Fields("prn").Value & ""
'        Acting_Check = Rs.Fields("acting_check").Value & ""
'        Dr_code = Rs.Fields("dr_code").Value & ""
'        Frequency_Of_Order = Rs.Fields("frequency_of_order").Value & ""
'        Return_Value = Rs.Fields("return_value").Value & ""
'    Else
'        GoTo Errors
'    End If
'    Set Rs = Nothing
'
'    '출고항목에 대한 검사명 읽어오기
'    sSql = " SELECT item_code,item_namee,pmpa_code " & _
'           " FROM med_ocs.OPDIPD_OCS_SLIP_DMC " & _
'           " WHERE ITEM_CODE='" & Trim(OrdCd) & "'"
'    Set Rs = OpenRecordSet(sSql)
'
'    If Not Rs.EOF Then
'        Item_Code = Rs.Fields("item_code").Value & ""
'        Item_Name = Rs.Fields("item_namee").Value & ""
'        Pmpa_Code = Rs.Fields("pmpa_code").Value & ""
'    Else
'        GoTo Errors
'    End If
'    Set Rs = Nothing
'
'    '필터일경우 필터의 검사명 가져오기, 필터가 아닐경우 혈액제제명 가져오기
'    If Filterfg = "1" Then
'        CompoNm = Item_Name
'    Else
'        sSql = " SELECT COMPONM as CompoNm FROM " & T_BBS006 & _
'               " WHERE " & DBW("compocd=", Trim(CompoCd)) '& _
'               " AND (expdt='' or expdt is null)"
'        Set Rs = OpenRecordSet(sSql)
'
'
'        If Not Rs.EOF Then
'            CompoNm = Rs.Fields("CompoNm").Value & ""
'        Else
'            GoTo Errors
'        End If
'    End If
'
'    '새로운 order_key 발생
'    Order_key = GetOrderKey
'    'Insert 문장 생성
'    If Bussdiv = "2" Then
'        OCS_SuNapSQL = " INSERT INTO " & T_BBS903 & _
'                       " (Patient_No,order_date,old_order_key,order_key,item_code,item_name,qty_of_order,remark," & _
'                       "  duration_of_order,dosage_code,night_day,seq_no,value_of_frequency,self,exam_specimen," & _
'                       "  stat_normal,pmpa_code,ordered_dr_code,class_of_order,group_of_order,write_date,clinical_dept,order_check," & _
'                       "  unit_of_order,request_act_date,input_status,write_status,order_site, contents_of_order,dosage_of_order," & _
'                       "  base_contents,prn,acting_check,dr_code,frequency_of_order,return_value,bbskey,filterfg,componm,orddiv," & _
'                       "  entdt,donefg,bussdiv,mesg) values(" & _
'                       DBS(Patient_No, 1) & ConvToDate(Order_Date) & "," & DBN(Old_Order_Key, 1) & DBN(Order_key, 1) & _
'                       DBS(Item_Code, 1) & DBS(Item_Name, 1) & DBN(Qty_Of_Order, 1) & DBS(remark, 1) & _
'                       DBN(Duration_Of_Order, 1) & DBS(Dosage_Code, 1) & DBN(Night_Day, 1) & DBN(Seq_No, 1) & _
'                       DBN(Value_Of_Frequency, 1) & DBS(Self, 1) & DBS(Exam_Specimen, 1) & DBS(Stat_Normal, 1) & _
'                       DBS(Pmpa_Code, 1) & DBS(Order_Dr_Code, 1) & DBN(Class_Of_Order, 1) & DBN(Group_Of_Order, 1) & _
'                       " To_Date('" & Write_Date & "','YYYY-MM-DD HH24:MI:SS')" & "," & DBS(Clinical_dept, 1) & DBS(Order_check, 1) & DBS(Unit_of_order, 1) & _
'                       ConvToDate(Request_Act_Date) & "," & DBS(Input_status, 1) & DBS(Write_status, 1) & DBS(Order_site, 1) & _
'                       DBN(Contents_Of_Order, 1) & DBS(Dosage_Of_Order, 1) & DBN(Base_Contents, 1) & DBS(PRN, 1) & DBS(Acting_Check, 1) & DBS(Dr_code, 1) & _
'                       DBS(Frequency_Of_Order, 1) & DBS(Return_Value, 1) & DBS(BBSKey, 1) & DBS(Filterfg, 1) & DBS(CompoNm, 1) & DBS(OrdDiv, 1) & "sysdate," & _
'                       DBS(DoneFg, 1) & DBS(Bussdiv, 1) & DBS(Mesg) & ")"
'
'    ElseIf Bussdiv = "1" Then
'        OCS_SuNapSQL = " INSERT INTO " & T_BBS903 & _
'                       " (Patient_No,order_date,old_order_key,order_key,item_code,item_name,qty_of_order,remark," & _
'                       "  duration_of_order,dosage_code,night_day,seq_no,value_of_frequency,self,exam_specimen," & _
'                       "  stat_normal,pmpa_code,ordered_dr_code,class_of_order,group_of_order,write_date,clinical_dept,order_check," & _
'                       "  unit_of_order,input_status,write_status,order_site,contents_of_order,dosage_of_order,base_contents,prn," & _
'                       "  acting_check,dr_code,frequency_of_order,return_value, bbskey,filterfg,componm,orddiv," & _
'                       "  entdt,donefg,bussdiv,mesg) values(" & _
'                       DBS(Patient_No, 1) & ConvToDate(Order_Date) & "," & DBN(Old_Order_Key, 1) & DBN(Order_key, 1) & _
'                       DBS(Item_Code, 1) & DBS(Item_Name, 1) & DBN(Qty_Of_Order, 1) & DBS(remark, 1) & _
'                       DBN(Duration_Of_Order, 1) & DBS(Dosage_Code, 1) & DBN(Night_Day, 1) & DBN(Seq_No, 1) & _
'                       DBN(Value_Of_Frequency, 1) & DBS(Self, 1) & DBS(Exam_Specimen, 1) & DBS(Stat_Normal, 1) & _
'                       DBS(Pmpa_Code, 1) & DBS(Order_Dr_Code, 1) & DBN(Class_Of_Order, 1) & DBN(Group_Of_Order, 1) & _
'                       " To_Date('" & Write_Date & "','YYYY-MM-DD HH24:MI:SS')" & "," & DBS(Clinical_dept, 1) & DBS(Order_check, 1) & DBS(Unit_of_order, 1) & _
'                       DBS(Input_status, 1) & DBS(Write_status, 1) & DBS(Order_site, 1) & _
'                       DBN(Contents_Of_Order, 1) & DBS(Dosage_Of_Order, 1) & DBN(Base_Contents, 1) & DBS(PRN, 1) & DBS(Acting_Check, 1) & DBS(Dr_code, 1) & _
'                       DBS(Frequency_Of_Order, 1) & DBS(Return_Value, 1) & DBS(BBSKey, 1) & DBS(Filterfg, 1) & DBS(CompoNm, 1) & DBS(OrdDiv, 1) & "sysdate," & _
'                       DBS(DoneFg, 1) & DBS(Bussdiv, 1) & DBS(Mesg) & ")"
'    End If
'
'    Set Rs = Nothing
'    Exit Function
'
'Errors:
'    MsgBox Err.Description, vbInformation + vbOKOnly
'    Set Rs = Nothing
'End Function

Public Function BldDelivery(ByVal rec As String) As Boolean
    Dim today       As String
    Dim TodayTm     As String
    Dim Bldsrc      As String
    Dim Bldyy       As String
    Dim Bldno       As String
    Dim CompoCd     As String
    Dim deliverydt  As String
    Dim Deliveryseq As String
    Dim deliverytm  As String
    Dim Deliveryid  As String
    Dim rcvid       As String
    Dim WorkArea    As String
    Dim accdt       As String
    Dim accseq      As String
    Dim RstSeq      As String
    Dim OrdCd       As String
    Dim LocalCd     As String
    Dim irrafg      As String
    Dim filter      As String
    Dim retfg       As String
    Dim expfg       As String
    Dim WorkSeq As String
    Dim PtId        As String
    Dim orddt       As String
    Dim OrdNo       As String
    Dim OrdSeq      As String
    
    Dim Bussdiv     As String
    Dim OcsOrdNo    As String
    Dim OCS_OrderKey As String
    
    
    Dim DeliveryCnt As String
    Dim DrRS        As Recordset
    
    Dim sSql    As String
    
    Dim examdt  As String
    Dim examtm  As String
    Dim examid  As String
    
    
'수가계산내역변수
'pheresis 처방일경우는 추가재료 사용내역별로. 수가내역을 발생한다.
    
    Dim chargeno As Long          '처방번호
    
    today = Format(GetSystemDate, PRESENTDATE_FORMAT)
    TodayTm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    Bldsrc = medGetP(rec, 1, COL_DIV)
    Bldyy = medGetP(rec, 2, COL_DIV)
    Bldno = medGetP(rec, 3, COL_DIV)
    CompoCd = medGetP(rec, 4, COL_DIV)
    deliverydt = medGetP(rec, 5, COL_DIV)
    Deliveryseq = medGetP(rec, 6, COL_DIV)
    deliverytm = medGetP(rec, 7, COL_DIV)
    Deliveryid = medGetP(rec, 8, COL_DIV)
    rcvid = medGetP(rec, 9, COL_DIV)
    WorkArea = medGetP(rec, 10, COL_DIV)
    accdt = medGetP(rec, 11, COL_DIV)
    accseq = medGetP(rec, 12, COL_DIV)
    RstSeq = medGetP(rec, 13, COL_DIV)
    OrdCd = medGetP(rec, 14, COL_DIV)
    LocalCd = medGetP(rec, 15, COL_DIV)
    irrafg = medGetP(rec, 16, COL_DIV)
    If irrafg = "1" Then
        irrafg = "Y"
    Else
        irrafg = ""
    End If
    filter = medGetP(rec, 17, COL_DIV)
    retfg = medGetP(rec, 18, COL_DIV)
    expfg = medGetP(rec, 19, COL_DIV)
    
    PtId = medGetP(rec, 20, COL_DIV)
    orddt = medGetP(rec, 21, COL_DIV)
    OrdNo = medGetP(rec, 22, COL_DIV)
    OrdSeq = medGetP(rec, 23, COL_DIV)
    
    Bussdiv = medGetP(rec, 24, COL_DIV)
    OcsOrdNo = medGetP(rec, 25, COL_DIV)
    
    '----------------------------------------------
    'deliveryseq와 ordcd는 새로 구한다
    '----------------------------------------------
    Deliveryseq = GetMaxDeliverySeq(Bldsrc, Bldyy, Bldno, CompoCd, deliverydt) + 1
    
    '----------------------------------------------
    '처방 상태를 update하기 위해 출고갯수를 구한다.
    '----------------------------------------------
    sSql = " SELECT deliverycnt " & _
           " FROM " & T_BBS203 & " " & _
           " WHERE " & _
                               DBW("workarea=", WorkArea) & _
                     " AND " & DBW("accdt=", accdt) & _
                     " AND " & DBW("accseq=", accseq)
           
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
    If DrRS.EOF Then
        MsgBox "출고 정보가 정확하지 않습니다.", vbExclamation
        Exit Function
    End If
    If DrRS.RecordCount < 1 Then
        DeliveryCnt = 1
    Else
        DeliveryCnt = Val(DrRS.Fields("deliverycnt").Value & "") + 1
    End If
    Set DrRS = Nothing
 
     sSql = " SELECT vfyid, vfydt, vfytm " & _
           " FROM " & T_BBS302 & " " & _
           " WHERE " & _
                               DBW("workarea=", WorkArea) & _
                     " AND " & DBW("accdt=", accdt) & _
                     " AND " & DBW("accseq=", accseq)
           
    Set DrRS = New Recordset
    Call DrRS.Open(sSql, DBConn)
    If DrRS.EOF Then
        MsgBox "출고 정보가 정확하지 않습니다.", vbExclamation
        Exit Function
    End If
    If DrRS.RecordCount < 1 Then
        DeliveryCnt = 1
    Else
        examid = "" & DrRS.Fields("vfyid").Value
        examdt = "" & DrRS.Fields("vfydt").Value
        examtm = "" & DrRS.Fields("vfytm").Value
    End If
    Set DrRS = Nothing

 On Error GoTo BldDelivery_error
    '-------------------------------------------------------
    '출고처리를 한다.
    '-------------------------------------------------------
    '1: 병동수가내역 작성
    '2: 출고내역 생성 BBS402
    '3: 입고내역의 혈액상태 변경 BBS401 ( Assign -> 출고 )
    '4: 처방내역 상태 변경 LAB102 ( -> 출고 )
    '5: 출고갯수 Count 해준다.
    '-------------------------------------------------------

    '-------------------------------------------------------------------------------------------------
    ' 이하부분은 병원별로 수가내역을 작성하는 부분이므로 협의후 처리가 필요함
    ' -------------------------------------------------------------------------
    ' ★ 병원별 혈액은행 수납내역 작성
    ' ★ 입원 모두 발생
    ' ★ 테이블 : S2BBS_DELIVERY,출고시는 항상 INSERT
    ' ★ BBS 매칭 키 : BBSKEY(BLDSRC & BLDYY & BLDNO & COMPOCD),혈액처방의 경우
    ' ★               BBSKEY(FILTER CODE), FILTER 처방일경우
    ' -------------------------------------------------------------------------
    
'    '1.
'전주예수병원
'iodate(출고일자), bloodno(혈액번호), bldcode(수혈처방코드), iotype1(입고1,출고2,반납3,폐기4),
'iotype2(적십자1,병동2,외래3,환자4), patno, ordseqno, orddate만 입력
'IODATE, BLOODNO, BLDCODE, IOTYPE1, IOTYPE2, PATNO가 키
                    
'2005/06/15 modify by legends
'iodate에 시간까지 넣어 달라고 해서 넣어주는 걸로 수정함
'출고 시간까지 관리하므로 출고내역 유무를 판단하지 않음
                    
'    Dim Rs As Recordset
'
'    Set Rs = New Recordset
'
'    sSql = " select * from oras1.slxboutt " & _
'         " where iodate=" & FUNC_CONVERT("date", today, "yyyymmdd") & _
'         " and bloodno='" & Bldsrc & Bldyy & Format(Bldno, String(6, "0")) & "'" & _
'         " and bldcode='" & OrdCd & "'" & _
'         " and iotype1='2'" & _
'         " and iotype2='" & IIf(Bussdiv = "1", "3", "2") & "'" & _
'         " and patno ='" & PtId & "'"
'
'    Rs.Open sSql, DBConn
'
'    If Rs.EOF Then
'        sSql = " insert into oras1.slxboutt (iodate,bloodno,bldcode,irradyn,iotype1,iotype2,patno,ordseqno,orddate) values (" & _
'              FUNC_CONVERT("date", today, "yyyymmdd") & ",'" & Bldsrc & Bldyy & Format(Bldno, String(6, "0")) & "','" & OrdCd & "','" & irrafg & "'," & _
'              "'2','" & IIf(Bussdiv = "1", "3", "2") & "','" & PtId & "'," & OrdNo & "," & FUNC_CONVERT("date", orddt, "yyyymmdd") & ")"
'        DBConn.Execute sSql
'    End If
'
'    Set Rs = Nothing

    sSql = " insert into oras1.slxboutt (iodate,bloodno,bldcode,irradyn,iotype1,iotype2,patno,ordseqno,orddate) values (" & _
          FUNC_CONVERT("date", deliverydt & deliverytm, "yyyymmdd hh24:mi:ss") & ",'" & Bldsrc & Bldyy & Format(Bldno, String(6, "0")) & "','" & OrdCd & "','" & irrafg & "'," & _
          "'2','" & IIf(Bussdiv = "1", "3", "2") & "','" & PtId & "'," & OrdNo & "," & FUNC_CONVERT("date", orddt, "yyyymmdd") & ")"
    DBConn.Execute sSql

    
    '2
    sSql = " INSERT INTO " & T_BBS402 & "(bldsrc,bldyy,bldno,compocd,deliverydt,deliveryseq," & _
                                        "deliverytm,deliveryid,rcvid,workarea,accdt,accseq," & _
                                        "rstseq,ordcd,localcd,irrafg,filter,retfg,expfg,workseq) " & _
           " VALUES(" & _
                   DBV("bldsrc", Bldsrc, 1) & DBV("bldyy", Bldyy, 1) & DBV("bldno", Bldno, 1) & DBV("compocd", CompoCd, 1) & _
                   DBV("deliverydt", deliverydt, 1) & DBV("deliveryseq", Deliveryseq, 1) & DBV("deliverytm", deliverytm, 1) & _
                   DBV("deliveryid", Deliveryid, 1) & DBV("rcvid", rcvid, 1) & DBV("workarea", WorkArea, 1) & _
                   DBV("accdt", accdt, 1) & DBV("accseq", accseq, 1) & DBV("rstseq", RstSeq, 1) & DBV("ordcd", OrdCd, 1) & _
                   DBV("localcd", LocalCd, 1) & DBV("irrafg", irrafg, 1) & DBV("filter", filter, 1) & _
                   DBV("retfg", retfg, 1) & DBV("expfg", expfg, 1) & DBV("WorkSeq", OCS_OrderKey) & ")"
    DBConn.Execute sSql
    
    '3
    sSql = " update " & T_BBS401 & " " & _
           " set " & DBW("stscd", BBSBloodStatus.stsDELIVERY, 2) & _
           " WHERE " & _
                    "     " & DBW("bldsrc", Bldsrc, 2) & _
                    " AND " & DBW("bldyy", Bldyy, 2) & _
                    " AND " & DBW("bldno", Bldno, 2) & _
                    " AND " & DBW("compocd", CompoCd, 2)
    DBConn.Execute sSql
    
    '4
'    sSql = " update " & T_LAB102 & _
'           " set " & DBW("stscd", stsINPROCESS, 2) & _
'           " WHERE " & _
'                  "     " & DBW("ptid", PtId, 2) & _
'                  " AND " & DBW("orddt", orddt, 2) & _
'                  " AND " & DBW("ordno", OrdNo, 2) & _
'                  " AND " & DBW("ordseq", OrdSeq, 2) & _
'                  " AND " & DBW("stscd<", stsINPROCESS)
                  
'전주예수병원용
'상태전송 미포함
'    sSql = " update  mdbldort " & _
'           " set " & DBW("stscd", stsINPROCESS, 2) & _
'           " WHERE " & _
'                  "     " & DBW("patno", PtId, 2) & _
'                  " and orddate=to_date(" & orddt & " ,'yyyymmdd') " & _
'                  " AND " & DBW("ordseqno", OrdNo, 2) & _
'                  " AND " & DBW("stscd<", stsINPROCESS)
'상태전송 포함
'OCS에 상태를 R로 전송 (E로 전송?)
    sSql = " update  mdbldort " & _
           " set " & DBW("stscd", stsINPROCESS, 3) & DBW("procstat", "R", 2) & _
           " WHERE " & _
                  "     " & DBW("patno", PtId, 2) & _
                  " and orddate=to_date(" & orddt & " ,'yyyymmdd') " & _
                  " AND " & DBW("ordseqno", OrdNo, 2) & _
                  " AND " & DBW("stscd<", stsINPROCESS)
    
    DBConn.Execute sSql
    
'2005/02/04 혈액이 하나라도 출고되면 readydt를 박아준다.
    sSql = " update  mdbldort " & _
           " set readydt=to_date(" & deliverydt & deliverytm & ",'yyyymmdd hh24:mi:ss')" & _
           " WHERE " & DBW("patno", PtId, 2) & _
                  " and orddate=to_date(" & orddt & " ,'yyyymmdd') " & _
                  " AND " & DBW("ordseqno", OrdNo, 2)
    
    DBConn.Execute sSql
    '5
    sSql = " update " & T_BBS203 & _
           " set " & DBW("deliverycnt", DeliveryCnt, 2) & _
           " WHERE " & _
           "     " & DBW("workarea", WorkArea, 2) & _
           " AND " & DBW("accdt", accdt, 2) & _
           " AND " & DBW("accseq", accseq, 2)
    DBConn.Execute sSql
    
    BldDelivery = True
    Exit Function
    
BldDelivery_error:
    MsgBox Err.Description & sSql
    BldDelivery = False
End Function

Private Function GetTestDiv(ByVal testcd As String, ByVal applydt As String, testdiv As String) As Boolean
    Dim sSql As String
    Dim Rs As Recordset
    
    sSql = "SELECT testdiv " & _
           "FROM " & T_BBS001 & " " & _
           "WHERE testcd=" & DBV("testcd", testcd) & " " & _
           "AND applydt=(SELECT max(applydt) " & _
                        "FROM " & T_BBS001 & " " & _
                        "WHERE " & _
                                        DBW("testcd=", testcd) & " " & _
                               "AND " & DBW("applydt<=", applydt) & ") "
    Set Rs = New Recordset
    Call Rs.Open(sSql, DBConn)
    If Rs.EOF Then
'        dbconn.DisplayErrors
        GetTestDiv = False
    Else
        With Rs
            If .RecordCount > 0 Then
                GetTestDiv = True
                testdiv = Rs.Fields("testdiv").Value & ""
            Else
                GetTestDiv = False
                testdiv = ""
            End If
        End With
    End If
    Set Rs = Nothing
End Function

Private Function Pheresis_Enter(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal CompoCd As String) As Boolean
    Dim Rs   As New Recordset
    Dim sSql As String
    
    sSql = " SELECT * FROM " & T_BBS602 & _
           " WHERE " & _
                             DBW("bldsrc", Bldsrc, 2) & _
                   " AND " & DBW("bldyy", Bldyy, 2) & _
                   " AND " & DBW("bldno", Bldno, 2) & _
                   " AND " & DBW("compocd", CompoCd, 2) & _
                   " AND " & DBW("entfg", "1", 2) & _
                   " AND " & DBW("donorcd", "3", 2)
    
    Call Rs.Open(sSql, DBConn)
    
    If Rs.EOF Then
        Set Rs = Nothing
        Exit Function
    End If

On Error GoTo Pheresis_Enter_Error

    If Not Rs.EOF Then
        sSql = " update " & T_BBS603 & " set " & DBW("stscd", DonorStatus.stsFinish, 2) & _
               " WHERE " & DBW("donorid", Rs.Fields("donorid").Value & "", 2) & _
               " AND   " & DBW("donoraccdt", Rs.Fields("donoraccdt").Value & "", 2)
        DBConn.Execute sSql
    End If
    Pheresis_Enter = True
    Set Rs = Nothing
    Exit Function
Pheresis_Enter_Error:
    Pheresis_Enter = False
    Set Rs = Nothing
End Function

Private Function FindABORH(ByVal PtId As String, ByVal orddt As String, ByVal BedIndt As String, ByVal aborhcd As String) As Long
    Dim sSql As String
    Dim Rs As Recordset
    
    sSql = "SELECT count(*) cnt " & _
           "FROM " & T_INT002 & " " & _
           "WHERE " & DBW("ptid=", PtId) & " " & _
           "AND " & DBW("orddt>=", BedIndt) & " " & _
           "AND " & DBW("orddt<=", orddt) & " " & _
           "AND " & DBW("ordcd=", aborhcd) & " "
    Set Rs = New Recordset
    Call Rs.Open(sSql, DBConn)
    If Rs.EOF Then
        FindABORH = -1
    Else
        With Rs
            FindABORH = .Fields("cnt").Value & ""
        End With
    End If
    Set Rs = Nothing
End Function

Private Function GetABORh(astrABORH() As String) As Long

    Dim Rs As Recordset
    Dim i As Long
    
    Erase astrABORH
    

    Set Rs = GetCom003RecordSet(BC2_ABO_RH)
    
    If Rs.EOF Then GetABORh = -1: Exit Function
    
    With Rs
        GetABORh = .RecordCount
        For i = 1 To .RecordCount
            ReDim Preserve astrABORH(i - 1)
            astrABORH(i - 1) = .Fields("cdval1").Value & ""
            .MoveNext
        Next i
    End With
    Set Rs = Nothing
End Function

Private Function GetBedinDt(ByVal PtId As String, ByVal orddt As String, BedIndt As String) As Boolean
    Dim sSql As String
    Dim Rs As Recordset
    
    
    sSql = "SELECT max(" & F_BEDINDT & ") bedindt " & _
           "FROM " & T_HIS002 & " " & _
           "WHERE " & F_INPTID & "=" & DBV("ptid", PtId) & " " & _
           "AND   " & F_BEDINDT & "= " & FUNC_CONVERT("char", F_BEDINDT, "yyyymmdd") ' to_date(" & BedIndt & ",'YYYYMMDD') "
           
    Set Rs = New Recordset
    Call Rs.Open(sSql, DBConn)
    If Rs.EOF Then
        GetBedinDt = False
    Else
        GetBedinDt = True
        With Rs
            If .RecordCount < 1 Then
                BedIndt = ""
            Else
                If (.Fields("bedindt").Value & "") = "" Then
                    BedIndt = ""
                Else
                    BedIndt = .Fields("bedindt").Value & ""
                End If
            End If
        End With
    End If
    Set Rs = Nothing
    
End Function

Private Function Bussdiv_Chk(ByVal PtId As String, ByVal orddt As String, ByVal OrdNo As String, ByVal OrdSeq As String) As Boolean
'-----------------------------------
'입원환자여부를 체크한다(입원:True)
'-----------------------------------
    Dim sSql As String
    Dim Rs   As New Recordset
    
    sSql = " SELECT bussdiv FROM " & T_LAB101 & _
           " WHERE " & DBW("ptid", PtId, 2) & " AND " & DBW("orddt", orddt, 2) & " AND " & DBW("ordno", OrdNo, 2)
         
    Call Rs.Open(sSql, DBConn)
    
    If Rs.EOF Then
'        dbconn.DisplayErrors
        Set Rs = Nothing
        Exit Function
    End If
    
    Bussdiv_Chk = (Rs.Fields("bussdiv").Value & "" = "2")
    
    Set Rs = Nothing
End Function

Private Function NewTestdiv(ByVal PtId As String, ByVal orddt As String, ByVal OrdNo As String, ByVal OrdSeq As String) As Boolean
'--------------------------------------------------
'처방 바디의 newtestdiv를 체크해서 Y일때만 True이다.
'--------------------------------------------------
    Dim Rs   As New Recordset
    Dim sSql As String
     
    sSql = " SELECT newtestdiv FROM " & T_LAB102 & _
           " WHERE " & DBW("ptid", PtId, 2) & " AND " & DBW("orddt", orddt, 2) & _
           " AND " & DBW("ordno", OrdNo, 2) & " AND " & DBW("ordseq", OrdSeq, 2)
    Call Rs.Open(sSql, DBConn)
    If Rs.EOF Then
'        dbconn.DisplayErrors
        Set Rs = Nothing
        Exit Function
    End If
    If Not Rs.EOF Then
        If Rs.Fields("newtestdiv").Value & "" = "Y" Then
            NewTestdiv = True
        End If
    End If
    
    Set Rs = Nothing
End Function

Private Function Get_PheresisInfo(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, _
                                 ByVal CompoCd As String) As Recordset
    Dim sSql As String
    
    sSql = " SELECT a.pherefg,c.ordcd,c.qty" & _
           " FROM " & T_BBS401 & " a," & T_BBS602 & " b," & T_BBS606 & " c" & _
           " WHERE " & DBW("a.bldsrc", Bldsrc, 2) & _
           "  AND " & DBW("a.bldyy", Bldyy, 2) & _
           "  AND " & DBW("a.bldno", Bldno, 2) & _
           "  AND " & DBW("a.compocd", CompoCd, 2) & _
           "  AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd=b.compocd" & _
           "  AND b.donorid=c.donorid AND b.donoraccdt=c.donoraccdt"
           
    Set Get_PheresisInfo = New Recordset
    Call Get_PheresisInfo.Open(sSql, DBConn)
    
End Function

Public Function GetMaxChargeNo(PtId As String) As Long
''수가계산내역의 ordno를 구한다.
'    Dim Rs As RECORDSET
'    Dim sSql As String
'
'    sSql = " SELECT max(ordno) as maxordno FROM " & T_INT002 & _
'           " WHERE " & _
'           " " & DBW("ptid", PtId, 2) & _
'           " AND " & DBW("orddt", Format(GetSystemDate, PRESENTDATE_FORMAT), 2)
'
'    Set Rs = New RECORDSET
'    Call Rs.open(ssql,dbconn)
'    If Rs.DBerror = True Then
'        DBConn.DisplayErrors
'        GetMaxChargeNo = 1
'        Set Rs = Nothing
'        Exit Function
'    End If
'    If Rs.EOF Then
'        GetMaxChargeNo = 1
'    Else
'        If IsNull(Rs.Fields("maxordno").Value) = True Then
'            GetMaxChargeNo = 1
'        Else
'            GetMaxChargeNo = Val(Rs.Fields("maxordno").Value) + 1
'        End If
'    End If
'    Rs.RsClose
'    Set Rs = Nothing
End Function

Public Function GetMaxDeliverySeq(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal CompoCd As String, ByVal deliverydt As String) As Long
    Dim sSql As String
    Dim Rs As Recordset
    
    sSql = " SELECT max(deliveryseq) maxdeliveryseq " & _
           " FROM " & T_BBS402 & " " & _
           " WHERE " & DBW("bldsrc=", Bldsrc) & _
           " AND " & DBW("bldyy=", Bldyy) & _
           " AND " & DBW("bldno=", Bldno) & _
           " AND " & DBW("compocd=", CompoCd) & _
           " AND " & DBW("deliverydt=", deliverydt)
           
    Set Rs = New Recordset
    Call Rs.Open(sSql, DBConn)
    If Rs.EOF Then
'        dbconn.DisplayErrors
        GetMaxDeliverySeq = 0
        Set Rs = Nothing
        Exit Function
    End If
    If Rs.EOF = True Then
        GetMaxDeliverySeq = 0
    Else
        If IsNull(Rs.Fields("maxdeliveryseq").Value) = True Then
            GetMaxDeliverySeq = 0
        Else
            GetMaxDeliverySeq = Val(Rs.Fields("maxdeliveryseq").Value & "")
        End If
    End If
    Set Rs = Nothing
End Function

Private Function IsPheresis(ByVal CompoCd As String) As Boolean

    Dim Rs As Recordset
    

    Set Rs = GetComponent(CompoCd)
    
    If Rs Is Nothing Then IsPheresis = False: Exit Function
    If Rs.RecordCount < 1 Then
        IsPheresis = False
    Else
        IsPheresis = (Rs.Fields("pherefg").Value & "" = 1)
    End If
    Set Rs = Nothing
End Function

Public Function GetOrdCd(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, ByVal CompoCd As String, ByVal deliverydt As String, ByVal volume As String, ByVal testdiv As String, aryOrdCd() As String) As Long
    Dim sSql    As String
    Dim today   As String
    Dim Rs      As Recordset
    Dim i       As Long
    
    Erase aryOrdCd
    
    If testdiv = "0" Then
        sSql = " SELECT testcd,testnm " & _
               " FROM " & T_BBS001 & " a " & _
               " WHERE " & DBW("compocd", CompoCd, 2) & _
                       " AND " & DBW("volumn", volume, 2) & _
                       " AND " & DBW("testdiv", testdiv, 2) & _
                       " AND applydt = (SELECT max(applydt) FROM " & T_BBS001 & " b WHERE a.testcd=b.testcd AND " & DBW("b.applydt<=", deliverydt) & ") " & _
               " and (expdt='' or expdt is null) " & _
               " ORDER BY testcd "
    Else
        sSql = " SELECT testcd,testnm " & _
               " FROM " & T_BBS001 & " a " & _
               " WHERE " & DBW("compocd", CompoCd, 2) & _
               " AND " & DBW("testdiv", testdiv, 2) & _
               " AND applydt = (SELECT max(applydt) FROM " & T_BBS001 & " b WHERE a.testcd=b.testcd AND " & DBW("b.applydt<=", deliverydt) & ") " & _
               " and (expdt='' or expdt is null) " & _
               " ORDER BY testcd "
    End If
    
    Set Rs = New Recordset
    Call Rs.Open(sSql, DBConn)
    If Rs.EOF Then
'        dbconn.DisplayErrors
        GetOrdCd = 0
        Set Rs = Nothing
        Exit Function
    End If
    
    With Rs
        If .RecordCount < 1 Then
            GetOrdCd = 0
        Else
            GetOrdCd = .RecordCount
            ReDim Preserve aryOrdCd(.RecordCount - 1)
            For i = 1 To .RecordCount
                aryOrdCd(i - 1) = .Fields("testcd").Value & "" & " " & .Fields("testnm").Value & ""
                .MoveNext
            Next i
        End If
        
    End With
    Set Rs = Nothing
End Function
'혈액은행 자체폐기시......SQL 문장....

Public Function BldHosExpire(ByVal Bldsrc As String, _
                          ByVal Bldyy As String, _
                          ByVal Bldno As String, _
                          ByVal CompoCd As String, _
                          ByVal exprcvid As String, _
                          ByVal expid As String, _
                          ByVal Exprsncd As String, _
                          ByVal expbilldiv As String, _
                          ByVal exprsnrmk As String, _
                          Optional ByVal XmTest As Boolean = True) As Boolean
                          

'HosExp=true 이면, 자체폐기

    Dim today   As Date
    Dim sSql    As String
    Dim Rs    As Recordset
    
    Dim strWorkArea     As String
    Dim strAccDt        As String
    Dim strAccSeq       As String
    Dim strExpCnt       As String
    Dim strAssignCnt    As String

'수각계산내역변수
    
On Error GoTo BldExpire_error
    DBConn.BeginTrans
    
    today = GetSystemDate
   
    If XmTest = False Then
        '--------------------------------------------------
        '처방상태의 폐기갯수를 수정하기 위해 갯수를 구한다.
        '--------------------------------------------------
        sSql = " SELECT a.workarea,a.accdt,a.accseq,a.expcnt,a.assigncnt " & _
               " FROM " & T_BBS203 & " a," & T_BBS302 & " b " & _
               " WHERE a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq " & _
               " AND " & DBW("b.bldsrc", Bldsrc, 2) & _
               " AND " & DBW("b.bldyy", Bldyy, 2) & _
               " AND " & DBW("b.bldno", Bldno, 2) & _
               " AND " & DBW("b.compocd", CompoCd, 2)
                  
                  
        Set Rs = New Recordset
        Call Rs.Open(sSql, DBConn)
        If Rs.EOF Then GoTo BldExpire_error

        strWorkArea = Rs.Fields("workarea").Value & ""
        strAccDt = Rs.Fields("accdt").Value & ""
        strAccSeq = Rs.Fields("accseq").Value & ""
        strAssignCnt = CStr(Val(Rs.Fields("assigncnt").Value & "") + 1)
        strExpCnt = CStr(Val(Rs.Fields("expcnt").Value & "") + 1)

        Set Rs = Nothing
    
        sSql = " update " & T_BBS302 & " " & _
               " set " & DBW("cancelid", expid, 3) & _
                         DBW("cancelfg", BBSCancelStatus.stsCancel, 3) & _
                         DBW("canceltm", Format(today, PRESENTTIME_FORMAT), 3) & _
                         DBW("canceldt", Format(today, PRESENTDATE_FORMAT), 2) & _
               " WHERE " & _
                                   DBW("bldsrc", Bldsrc, 2) & _
                         " AND " & DBW("bldyy", Bldyy, 2) & _
                         " AND " & DBW("bldno", Bldno, 2) & _
                         " AND " & DBW("compocd", CompoCd, 2)
        DBConn.Execute sSql

        sSql = " update " & T_BBS203 & " " & _
               " set " & DBW("expcnt=", strExpCnt) & _
               " WHERE " & DBW("workarea=", strWorkArea) & _
               " AND   " & DBW("accdt=", strAccDt) & _
               " AND   " & DBW("accseq=", strAccSeq)
        DBConn.Execute sSql
    End If
    
    
    sSql = " update " & T_BBS401 & " " & _
           " set " & DBW("realexpdt", Format(today, PRESENTDATE_FORMAT), 3) & _
                     DBW("realexptm", Format(today, PRESENTTIME_FORMAT), 3) & _
                     DBW("exprcvid", exprcvid, 3) & _
                     DBW("expid", expid, 3) & _
                     DBW("exprsncd", Exprsncd, 3) & _
                     DBW("expbilldiv", expbilldiv, 3) & _
                     DBW("exprsnrmk", exprsnrmk, 3) & _
                     DBW("stscd", BBSBloodStatus.stsEXPIRE, 2) & _
           " WHERE " & _
                     "     " & DBW("bldsrc", Bldsrc, 2) & _
                     " AND " & DBW("bldyy", Bldyy, 2) & _
                     " AND " & DBW("bldno", Bldno, 2) & _
                     " AND " & DBW("compocd", CompoCd, 2)
    DBConn.Execute sSql
    
    DBConn.CommitTrans
    
    BldHosExpire = True
    Exit Function
    
BldExpire_error:
    DBConn.RollbackTrans
    
    BldHosExpire = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function GetIRROrder() As String
    Dim sSql As String
    Dim Rs   As Recordset
    sSql = " SELECT testnm FROM " & T_BBS001 & " WHERE " & DBW("orddiv=", "Z")
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Not Rs.EOF Then
        GetIRROrder = Rs.Fields("testnm").Value & ""
    Else
        GetIRROrder = ""
    End If
    Set Rs = Nothing
End Function

Private Sub Class_Initialize()
    If TRANS_REQUIRE_USED = True Then
        stsORDER = BBSOrdStatus.stsORDER
        stsCOLLECT = BBSOrdStatus.stsCOLLECT
        stsACCESS = BBSOrdStatus.stsACCESS
        stsINPROCESS = BBSOrdStatus.stsINPROCESS
    Else
        stsORDER = BBSOrderStatus.stsORDER
        stsCOLLECT = BBSOrderStatus.stsCOLLECT
        stsACCESS = BBSOrderStatus.stsACCESS
        stsINPROCESS = BBSOrderStatus.stsINPROCESS
    End If
End Sub

Private Function GetComponent(ByVal CompoCd As String, Optional QueryExp As Boolean = False) As Recordset
    Dim sSql As String

    sSql = "SELECT * " & _
           "FROM " & T_BBS006 & " " & _
           "WHERE " & DBW("compocd=", CompoCd) & " "
    If QueryExp = False Then
        sSql = sSql & _
               "AND (expdt is null or expdt='') "
    End If
    Set GetComponent = New Recordset
    GetComponent.Open sSql, DBConn
End Function
