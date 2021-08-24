VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHospital05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'성바오로에서 추가한 클래스 입니다.
'2002/07/04 :김정규

Public Function GetPtInformatio(ByVal qPtid As String) As String
'환자기본정보(환자명/주민번호)를 가지고 온다.
    Dim Rs   As Recordset
    Dim sSql As String
    
    sSql = " SELECT " & F_PTNM & " as ptnm," & F_SSN & " as ssn " & _
           " FROM " & T_HIS001 & _
           " WHERE " & DBW(F_PTID, qPtid, 2)
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    If Not Rs.EOF Then
        GetPtInformatio = Rs.Fields("ptnm").Value & "" & COL_DIV & Rs.Fields("ssn").Value & ""
    End If
    Set Rs = Nothing
    
End Function
Public Function GetDeilveryList(ByVal qPtid As String) As String
    Dim sSql As String
    
    
    sSql = " SELECT c.abo,c.rh,c.volumn,c.stscd,c.realexpdt," & _
                  " b.bldsrc,b.bldyy,b.bldno,b.compocd,b.deliverydt,b.deliveryseq,b.reactionfg" & _
           " FROM " & T_BBS401 & " c," & T_BBS506 & " b," & T_BBS202 & " a" & _
           " WHERE " & DBW("a.ptid=", qPtid) & " AND " & _
                   " a.workarea=b.workarea AND " & _
                   " a.accdt=b.accdt AND " & _
                   " a.accseq=b.accseq AND " & _
                   " b.bldsrc=c.bldsrc AND " & _
                   " b.bldyy=c.bldyy AND " & _
                   " b.bldno=c.bldno AND " & _
                   " b.compocd = c.compocd " & _
           " ORDER BY b.deliverydt,b.deliveryseq"
    
    GetDeilveryList = sSql
    
End Function
Public Function GetBloodInfo(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, _
                            ByVal CompoCd As String) As Recordset
    Dim sSql As String
    
    sSql = " SELECT  * " & _
           " FROM " & T_BBS401 & _
           " WHERE " & _
                   DBW("bldsrc", Bldsrc, 2) & " AND " & _
                   DBW("bldyy", Bldyy, 2) & " AND " & _
                   DBW("bldno", Bldno, 2) & " AND " & _
                   DBW("compocd", CompoCd, 2)
                   
    Set GetBloodInfo = New Recordset
    GetBloodInfo.Open sSql, DBConn
End Function
Public Function GetDeliveryNm(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, _
                            ByVal CompoCd As String, ByVal deliverydt As String, ByVal Deliveryseq As String) As String
    Dim sSql As String
    Dim Rs   As Recordset
    
    
    sSql = " SELECT deliveryid " & _
           " FROM " & T_BBS402 & _
           " WHERE " & _
                   DBW("bldsrc", Bldsrc, 2) & " AND " & _
                   DBW("bldyy", Bldyy, 2) & " AND " & _
                   DBW("bldno", Bldno, 2) & " AND " & _
                   DBW("compocd", CompoCd, 2) & " AND " & _
                   DBW("deliverydt", deliverydt, 2) & " AND " & _
                   DBW("deliveryseq", Deliveryseq, 2)
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    If Not Rs.EOF Then
        GetDeliveryNm = GetEmpNm(Rs.Fields("deliveryid").Value & "")
    End If
    Set Rs = Nothing
End Function

Public Function GetReactionSQL() As String
    '부작용리스트 조회
    Dim sSql As String
    
    sSql = " SELECT cdval1,field1" & _
           " FROM " & T_COM003 & _
           " WHERE " & DBW("cdindex", BC2_REACTION_REASON, 2) & _
           " ORDER BY cdval1"
    GetReactionSQL = sSql
End Function
Public Function UpdateSQL_BBS501(ByVal Bldsrc As String, ByVal Bldyy As String, ByVal Bldno As String, _
                                 ByVal CompoCd As String, ByVal deliverydt As String, _
                                 ByVal Deliveryseq As String, _
                                 ByVal ReactionCd As String, ByVal remark As String) As String
    Dim sSql As String
    
    sSql = " UPDATE " & T_BBS506 & _
           " SET " & _
                   DBW("reactiondt", Format(GetSystemDate, "YYYYMMDD"), 3) & _
                   DBW("reactioncd", ReactionCd, 3) & _
                   DBW("reactionid", ObjSysInfo.EmpID, 3) & _
                   DBW("remark", remark, 3) & DBW("reactionfg=", "1") & _
           " WHERE " & _
                   DBW("bldsrc=", Bldsrc) & " AND " & _
                   DBW("bldyy=", Bldyy) & " AND " & _
                   DBW("bldno=", Bldno) & " AND " & _
                   DBW("compocd=", CompoCd) & " AND " & _
                   DBW("deliverydt=", deliverydt) & " AND " & _
                   DBW("deliveryseq=", Deliveryseq)
    UpdateSQL_BBS501 = sSql
End Function
Public Function ReactionStatics(ByVal ReactionDt As String, Optional ByVal wardid As String) As clsDictionary
    Dim objdic As New clsDictionary
    Dim sSql   As String
    Dim RRS    As Recordset
    Dim Rs     As Recordset
    
    objdic.Clear
    objdic.FieldInialize "rsncd", "rsnnm,mon1,mon2,mon3,mon4,mon5,mon6,mon7,mon8,mon9,mon10,mon11,mon12"
    
    Set RRS = New Recordset
    RRS.Open GetReactionSQL, DBConn
    If Not RRS.EOF Then
        Do Until RRS.EOF
            objdic.AddNew RRS.Fields("cdval1").Value & "", Join(Array(RRS.Fields("field1").Value & "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), COL_DIV)
            RRS.MoveNext
        Loop
    End If
        
    Set RRS = Nothing
    
    If wardid <> "" Then
        sSql = " SELECT a.reactioncd,a.reactiondt" & _
               " FROM " & T_BBS202 & " b," & T_LAB101 & " c," & T_BBS506 & " a" & _
               " WHERE " & _
                         DBW("a.reactiondt>=", ReactionDt & "0101") & " AND " & _
                         DBW("a.reactiondt<=", ReactionDt & "1231") & " AND " & _
                         DBW("a.reactionfg=", "1") & " AND " & _
                       " a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq AND " & _
                       " b.ptid=c.ptid AND b.orddt=c.orddt AND b.ordno=c.ordno AND " & _
                         DBW("c.wardid=", wardid)
    Else
        sSql = " SELECT a.reactioncd,a.reactiondt" & _
               " FROM " & T_BBS506 & " a" & _
               " WHERE " & _
                         DBW("a.reactiondt>=", ReactionDt & "0101") & " AND " & _
                         DBW("a.reactiondt<=", ReactionDt & "1231") & " AND " & _
                         DBW("a.reactionfg=", "1")
                       
    End If
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    With objdic
        If Not Rs.EOF Then
            Do Until Rs.EOF
                If .Exists(Rs.Fields("reactioncd").Value & "") Then
                    .KeyChange Rs.Fields("reactioncd").Value & ""
                    Select Case Mid(Rs.Fields("reactiondt").Value, 5, 2)
                        Case "01": .Fields("mon1") = .Fields("mon1") + 1
                        Case "02": .Fields("mon2") = .Fields("mon2") + 1
                        Case "03": .Fields("mon3") = .Fields("mon3") + 1
                        Case "04": .Fields("mon4") = .Fields("mon4") + 1
                        Case "05": .Fields("mon5") = .Fields("mon5") + 1
                        Case "06": .Fields("mon6") = .Fields("mon6") + 1
                        Case "07": .Fields("mon7") = .Fields("mon7") + 1
                        Case "08": .Fields("mon8") = .Fields("mon8") + 1
                        Case "09": .Fields("mon9") = .Fields("mon9") + 1
                        Case "10": .Fields("mon10") = .Fields("mon10") + 1
                        Case "11": .Fields("mon11") = .Fields("mon11") + 1
                        Case "12": .Fields("mon12") = .Fields("mon12") + 1
                    End Select
                End If
                Rs.MoveNext
            Loop
        End If
    End With
    
    Set ReactionStatics = New clsDictionary
    Set ReactionStatics = objdic
    
    Set Rs = Nothing
End Function
Public Function GetExpReason(Optional ByVal ExpRsn As String = "") As String
    Dim sSql As String
    Dim Rs   As Recordset
    
    sSql = " SELECT * " & _
           " FROM " & T_COM003 & _
           " WHERE " & _
                  DBW("cdindex=", BC2_EXP_RESON)
    If ExpRsn <> "" Then
        sSql = sSql & " AND " & DBW("cdval1=", ExpRsn)
    End If
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    If Not Rs.EOF Then
        GetExpReason = Rs.Fields("field1").Value & ""
    End If
    Set Rs = Nothing
End Function
Public Function GetBloodDetailQuery(ByVal qEntDt As String) As String
    Dim sSql As String
    Dim sEntF As String
    Dim sEntT As String
    
    sEntF = medGetP(qEntDt, 1, COL_DIV)
    sEntT = medGetP(qEntDt, 2, COL_DIV)

    sSql = "SELECT a.entdt as dt,a.volumn ,a.abo,a.rh,b.abbrnm,count(*) as cnt " & _
         "FROM " & T_BBS401 & " a," & T_BBS006 & " b " & _
         "WHERE " & DBW("a.entdt >=", sEntF) & " " & _
         "AND " & DBW("a.entdt<=", sEntT) & " " & _
         "AND a.compocd =b.compocd " & _
         "GROUP BY entdt,volumn,abo,rh,abbrnm " & _
         "ORDER BY entdt, abo,volumn"
    
    GetBloodDetailQuery = sSql
    
End Function

Public Function GetBloodDeliveryQuery(ByVal qDeldt As String) As String
    Dim sSql As String
    Dim sDelF As String
    Dim sDelT As String
    
    '2005/05/30 modify by legends
    '실제 출고된 수량을 구하기 위해 반환된 혈액은 제외 "AND (c.retfg<>'1' or c.retfg is  null) " 조건 추가
    
    sDelF = medGetP(qDeldt, 1, COL_DIV)
    sDelT = medGetP(qDeldt, 2, COL_DIV)
    
    sSql = "SELECT c.deliverydt as dt,a.volumn,a.abo,a.rh,b.abbrnm,count(*) as cnt " & _
          "FROM " & T_BBS006 & " b," & T_BBS402 & " c," & T_BBS401 & " a " & _
          "WHERE " & _
                   DBW("a.stscd=", BBSBloodStatus.stsDELIVERY) & " " & _
          "AND " & DBW("c.deliverydt>=", sDelF) & " " & _
          "AND " & DBW("c.deliverydt<=", sDelT) & " " & _
          "AND a.bldsrc=c.bldsrc " & _
          "AND a.bldyy=c.bldyy " & _
          "AND a.bldno=c.bldno " & _
          "AND a.compocd=b.compocd " & _
          "AND (c.retfg<>'1' or c.retfg is  null) " & _
          "GROUP BY c.deliverydt,a.volumn,a.abo,a.rh,b.abbrnm " & _
          "ORDER BY deliverydt,abo,volumn"
    
    GetBloodDeliveryQuery = sSql
          
End Function
