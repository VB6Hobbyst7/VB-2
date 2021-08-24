VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWardColList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'#Const APS_FLAG = True
'#Const BBS_FLAG = True

Private iPWACoHeRow As Integer
Private fWhich As Object
Private iPageWidth As Integer
Private iPageHeight As Integer
Private iCurY As Integer
Private DataExist As Boolean
Private sLastDt As String
Private sLastTm As String
Private iRecordCount As Integer

Private pWardId As String
Private pWardNm As String
Private pWorkDt As String
Private pWorkTm As String
Private pBuildCd As String
Private pBuildNm As String
Private pTitleNm As String
    
Private SvBuildCd As String
Private SvBuildNm As String

Private mvarSetCrpt As Object
Private mvarTestDiv As String


Const iCm = 567
Const iLineHeight = 10

Public objDictionary As New clsDictionary

Private iposSEQ%, iposWorkNo%, iposPtName%, iposPtID%, iposSAge%, _
    iposIO%, iposRcv, iposSF%, iposTestCD%, iposSpccd%

Public Sub SetCrpt(ByVal Crystal As Object)
    Set mvarSetCrpt = Crystal
    
End Sub
Public Property Let TestDiv(ByVal vData As String)
    mvarTestDiv = vData
End Property


Public Property Get WardID() As String
    WardID = pWardId
End Property

Public Property Let WardID(ByVal vNewValue As String)
    pWardId = vNewValue
End Property

Public Property Get WardNm() As String
    WardNm = pWardNm
End Property

Public Property Let WardNm(ByVal vNewValue As String)
    pWardNm = vNewValue
End Property

Public Property Get WorkDt() As String
    WorkDt = pWorkDt
End Property

Public Property Let WorkDt(ByVal vNewValue As String)
    pWorkDt = vNewValue
End Property

Public Property Get WorkTm() As String
    WorkTm = pWorkTm
End Property

Public Property Let WorkTm(ByVal vNewValue As String)
    pWorkTm = vNewValue
End Property

Public Property Get BuildCd() As String
    BuildCd = pBuildCd
End Property

Public Property Let BuildCd(ByVal vNewValue As String)
    pBuildCd = vNewValue
End Property

Public Property Get BuildNm() As String
    BuildNm = pBuildNm
End Property

Public Property Let BuildNm(ByVal vNewValue As String)
    pBuildNm = vNewValue
End Property

Public Property Get TitleNm() As String
    TitleNm = pTitleNm
End Property

Public Property Let TitleNm(ByVal vNewValue As String)
    pTitleNm = vNewValue
End Property


Public Sub Print_ColList()
    Dim SqlStmt As String
    Dim tmpRs   As Recordset
    
    Call Crystal_CollectList
   
   
    SqlStmt = "select field1 as ColDate, field2 as ColTime from " & T_LAB031 & _
              " where " & DBW("cdindex=", LC2_ColListTm) & _
              " and " & DBW("cdval1=", pBuildCd) & _
              " and " & DBW("cdval2=", pWardId)
              
    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, DBConn
    
    If tmpRs.EOF Then
       DataExist = False
    Else
       DataExist = True
    End If
    
    Set tmpRs = Nothing
     
'MsgBox "현장검사 Report Print DB 정보 Setting"
'     If DataExist Then
'       SqlStmt = "update " & T_LAB031 & " set " & DBW("field1=", pWorkDt) & ", " & _
'                                                  DBW("field2=", sLastTm) & " " & _
'                 " where " & DBW("cdindex=", LC2_ColListTm) & _
'                 " and " & DBW("cdval1=", pBuildCd) & " " & _
'                 " and " & DBW("cdval2=", pWardId)
'     Else
'       SqlStmt = "insert into " & T_LAB031 & "(cdindex, cdval1, cdval2, field1, field2) " & _
'                     "values (" & _
'                                 DBV("cdindex", LC2_ColListTm, 1) & DBV("cdval1", pBuildCd, 1) & _
'                                 DBV("cdval2", pWardId, 1) & DBV("field1", pWorkDt, 1) & _
'                                 DBV("field2", sLastTm) & _
'                             ")"
'     End If
'
'     DBConn.BeginTrans
'     DBConn.Execute (SqlStmt)
'     DBConn.CommitTrans
    
End Sub



Private Function chkTestCD(sWorkarea As String, sAccDt As String, sAccSeq As String, stestdiv As String) As Boolean
    Dim sSQL2 As String
    Dim rsTestCode As Recordset
            
    sSQL2 = " select ordcd " & _
        " from " & T_LAB102 & _
        " where " & _
                          DBW("workarea=", sWorkarea) & _
                " and " & DBW("accdt=", Trim(sAccDt)) & _
                " and " & DBW("accseq=", Trim(sAccSeq))
        
    
    Set rsTestCode = New Recordset
    rsTestCode.Open sSQL2, DBConn
    
    If rsTestCode.EOF = True Then
        chkTestCD = False         ' not exitst
    Else
        chkTestCD = True              ' Exist
    End If
    
    Set rsTestCode = Nothing
End Function

Private Sub Crystal_CollectList()
    Dim sSQL        As String
    Dim rsWorksheet As Recordset
    Dim rsPtName    As Recordset
    Dim sSex        As String
    Dim strWorkInfo As String
    Dim strcoldt    As String
    Dim strColtm    As String
    
    Dim i As Integer
    
    Set rsWorksheet = CollectRecordSet
     
    If rsWorksheet.EOF = True Then ' record가 존재하지 않을경우
       MsgBox " 데이타가 존재하지 않습니다. "
       GoTo Nodata
    End If
    
    iRecordCount = rsWorksheet.RecordCount                      '검체수
    
    Dim objDic      As clsDictionary
    Dim sAge        As String
    Dim iSeqNum     As Integer
    Dim strStatFg   As String
    Dim WorkTestCd  As String
        
    Dim sICSString  As String
    Dim strEmpNm    As String
    Set objDic = New clsDictionary
    
    objDic.Clear
    objDic.FieldInialize "seq", "workno,ptnm,ptid,sexage,io,collectdt,testlist,spcnm,hosilid"
    
    
    sLastTm = ""
    With rsWorksheet
        .MoveFirst
        For i = 1 To .RecordCount
            sSQL = " select " & F_PTNM & " as ptnm," & F_DOB2 & " as dob ," & F_SSN2 & " as ssn from " & T_HIS001 & _
                    " where " & DBW(F_PTID & "=", "" & rsWorksheet.Fields("ptid").Value)
            objDic.Sort = False
            Set rsPtName = New Recordset
            rsPtName.Open sSQL, DBConn
            
            iSeqNum = iSeqNum + 1
            
            If ICSResultChk = True Then
                sICSString = ICSPatientString(.Fields("ptid").Value & "", enICSNum.LIS_ALL)
            Else
                strcoldt = Mid(.Fields("coldt").Value & "", 3)
                strColtm = " " & Mid(.Fields("coltm").Value & "", 1, 4)
                sICSString = strcoldt & strColtm
            End If
            
            Select Case .Fields("div").Value & ""
                Case LIS_ORDDIV:
                    If chkTestCD("" & .Fields("WorkArea").Value, "" & .Fields("accdt").Value, "" & .Fields("accseq").Value, _
                                 "" & .Fields("testdiv").Value) = True Then                   ' Exists
                        If .Fields("statfg").Value & "" = "" Or .Fields("statfg").Value & "" = "0" Then
                            strStatFg = ""
                        Else
                            strStatFg = "**"
                        End If
                        
                        sAge = ("" & .Fields("AgeDay").Value \ 365) + 1
                        
                        WorkTestCd = GetTestCD_LIS("" & .Fields("WorkArea").Value, "" & .Fields("AccDt").Value, _
                                                   "" & .Fields("AccSeq").Value, "" & .Fields("TestDiv").Value)
                        
                        strWorkInfo = .Fields("workarea").Value & "" & "-" & _
                                      Mid("" & .Fields("AccDt").Value, 3) & "-" & _
                                      .Fields("accseq").Value & ""
                        
                        objDic.AddNew iSeqNum & strStatFg, Join(Array(strWorkInfo, rsPtName.Fields("ptnm").Value & "", .Fields("ptid").Value & "", _
                                                         .Fields("sex").Value & "" & "/" & sAge & "", .Fields("deptcd").Value & "", _
                                                         sICSString, _
                                                         WorkTestCd, .Fields("spcnm").Value & "", .Fields("hosilid").Value & ""), COL_DIV)
                    End If
                Case BBS_ORDDIV:
                    sAge = medFindAge(rsPtName.Fields("dob").Value & "", "Y")
                    Select Case Mid(rsPtName.Fields("ssn").Value & "", 7, 1)
                        Case 1, 3: sSex = "M"
                        Case 2, 4: sSex = "F"
                        Case Else: sSex = "기타"
                    End Select
                    '응급여부체크
                    strStatFg = Get_statFg(.Fields("ptid").Value & "", .Fields("orddt").Value & "", .Fields("ordno").Value & "")
                    
                    
                    WorkTestCd = GetTestCd_BBS(.Fields("ptid").Value & "", .Fields("spcyy").Value & "", .Fields("spcno").Value)
                    
                    objDic.AddNew iSeqNum & strStatFg, Join(Array(.Fields("spcyy").Value & "-" & .Fields("spcno").Value, rsPtName.Fields("ptnm").Value & "", .Fields("ptid").Value & "", sSex & "/" & sAge, _
                                                      .Fields("deptcd").Value & "", sICSString, _
                                                      WorkTestCd, .Fields("spcnm").Value & "", .Fields("hosilid").Value & ""), COL_DIV)
                    
            End Select
            
            
            .MoveNext
        Next i
        Set rsPtName = Nothing
        
'        Dim objEmp As clsBasisData
        
'        Set objEmp = New clsBasisData
        strEmpNm = GetEmpNm(ObjSysInfo.EmpId)
'        Set objEmp = Nothing
        
'        strEmpNm = GetEmpName(objsysinfo.EmpId)
        
        If objDic.RecordCount > 0 Then
            strcoldt = Format(GetSystemDate, "yyyy/mm/dd")
            strColtm = " " & Format(GetSystemDate, "hh:mm")
            Call CollectList_Print(objDic, iRecordCount, pWardId, strcoldt & strColtm, strEmpNm, pBuildCd, pTitleNm)
        End If
    End With
    
    
    Set objDic = Nothing
Nodata:
    Set rsWorksheet = Nothing


End Sub
Private Function Get_statFg(ByVal ptid As String, ByVal OrdDt As String, ByVal ordno As Integer) As String
'-------------
'응급여부 판단
'-------------
    Dim sSQL As String
    Dim Rs   As Recordset
    
    Get_statFg = ""
    sSQL = " select b.statfg from " & T_LAB101 & " a," & T_LAB102 & " b" & _
            " where " & DBW("a.ptid=", ptid) & _
            " and " & DBW("a.orddt=", OrdDt) & _
            " and " & DBW("a.ordno=", ordno) & _
            " and a.ptid=b.ptid and a.orddt=b.orddt and a.ordno=b.ordno"
         
    Set Rs = New Recordset
    Rs.Open sSQL, DBConn
    
    If Rs.EOF Then
        Get_statFg = ""
        GoTo Nodata
    End If
    
    With Rs
        Do Until .EOF
            If .Fields("statfg").Value & "" = "1" Then
                Get_statFg = "**"
                GoTo Nodata
            End If
            .MoveNext
        Loop
    End With
    
Nodata:
    Set Rs = Nothing
End Function

Private Function GetTestCd_BBS(ByVal ptid As String, ByVal spcyy As String, ByVal spcno As String) As String
'-----------------------
'혈액은행 검사명 불러오기
'-----------------------
    Dim sSQL As String
    Dim Rs   As Recordset
    Dim ii   As Integer
    
    sSQL = " select a.ordcd,c.abbrnm5 " & _
         " from " & _
                   T_LAB102 & " a," & T_BBS102 & " b," & T_BBS001 & " c" & _
         " where " & _
                             DBW("b.spcyy=", spcyy) & _
                   " and " & DBW("b.spcno=", spcno) & _
                   " and b.ptid=a.ptid" & _
                   " and b.orddt=a.orddt" & _
                   " and b.ordno=a.ordno" & _
                   " and a.ordcd=c.testcd"
                   
    Set Rs = New Recordset
    Rs.Open sSQL, DBConn
    
    If Rs.EOF Then
        GetTestCd_BBS = ""
        GoTo Nodata
    End If
    With Rs
        .MoveFirst
        For ii = 1 To .RecordCount
            If mvarTestDiv = "1" Then
                GetTestCd_BBS = GetTestCd_BBS & .Fields("ordcd").Value & "" & ","
            Else
                GetTestCd_BBS = GetTestCd_BBS & .Fields("abbrnm5").Value & "" & ","
            End If
            .MoveNext
        Next
        GetTestCd_BBS = Mid(GetTestCd_BBS, 1, Len(GetTestCd_BBS) - 1)
    End With
    
Nodata:
    Set Rs = Nothing
End Function

Private Function GetTestCD_LIS(ByVal sWorkarea As String, ByVal sAccDt As String, _
                           ByVal sAccSeq As String, ByVal stestdiv As String) As String
'-----------------------
'임상병리검사명 불러오기
'-----------------------
    Dim sSQL2 As String
    Dim rsTestCode As Recordset
    Dim i%, tmpiposTestCD
    Dim sTable As String
    
    sSQL2 = " select a.ordcd,b.abbrnm5 " & _
        " from " & T_LAB102 & " a," & T_LAB001 & " b" & _
        " where " & _
                          DBW("a.workarea=", sWorkarea) & _
                " and " & DBW("a.accdt=", Trim(sAccDt)) & _
                " and " & DBW("a.accseq=", Trim(sAccSeq)) & _
                " and a.ordcd=b.testcd"
                
    Set rsTestCode = New Recordset
    rsTestCode.Open sSQL2, DBConn
    
    If rsTestCode.EOF = True Then
        GetTestCD_LIS = ""
        GoTo Nodata
    End If
    
    With rsTestCode
        rsTestCode.MoveFirst
        For i = 1 To rsTestCode.RecordCount
            If mvarTestDiv = "1" Then
                GetTestCD_LIS = GetTestCD_LIS & .Fields("ordcd").Value & "" & ","
            Else
                GetTestCD_LIS = GetTestCD_LIS & .Fields("abbrnm5").Value & "" & ","
            End If
            rsTestCode.MoveNext
        Next i
        GetTestCD_LIS = Mid(GetTestCD_LIS, 1, Len(GetTestCD_LIS) - 1)
    End With
    
Nodata:
    Set rsTestCode = Nothing
    
End Function

Private Sub CollectList_Print(ByVal objDic As clsDictionary, ByVal SpcCount As Integer, _
                              ByVal Ward As String, ByVal CollectDate As String, ByVal CollectNm As String, _
                              ByVal BuildInfo As String, ByVal TitleNm As String)
    Dim strTmp As String
    Dim intFNum As Integer
    Dim strRfile As String
    Dim strRptPath As String
    Dim ii      As Integer
    Dim jj      As Integer
    
    objDic.MoveFirst
    
    Do Until objDic.EOF
        strTmp = strTmp & objDic.Fields("seq") & vbTab & objDic.Fields("workno") & vbTab & objDic.Fields("ptnm") & vbTab & _
                          objDic.Fields("ptid") & vbTab & objDic.Fields("sexage") & vbTab & objDic.Fields("hosilid") & vbTab & _
                         objDic.Fields("collectdt") & vbTab & objDic.Fields("testlist") & vbTab & objDic.Fields("spcnm") & vbCr
        objDic.MoveNext
    Loop
    
    strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
    
    strRfile = InstallDir & "LIS\Rpt\CrystalReport.txt"
    strRptPath = InstallDir & "LIS\Rpt\CollectionList.rpt"
    
    If ICSResultChk = True Then
        strRptPath = InstallDir & "LIS\Rpt\CollListInfect.rpt"
    End If
  
  
    intFNum = FreeFile
    Open strRfile For Output As #intFNum
    Print #intFNum, strTmp
    Close #intFNum
    '
    With mvarSetCrpt
        If ObjSysInfo.UseBuildingInfo = "1" Then
            .ParameterFields(0) = "optLocation;" & "Deliver Location :" & BuildInfo & ";TRUE"
        Else
            .ParameterFields(0) = "optLocation;" & "" & ";TRUE"
        End If
        
        .ParameterFields(1) = "Ward;" & Ward & ";TRUE"
        .ParameterFields(2) = "ActionDate;" & CollectDate & ";TRUE"
        .ParameterFields(3) = "ActionNm;" & CollectNm & ";TRUE"
        .ParameterFields(4) = "SpcNm;" & "검체수 :  " & SpcCount & ";TRUE"
        .ParameterFields(5) = "title;" & TitleNm & ";TRUE"
        .ReportFileName = strRptPath
        .RetrieveDataFiles
        
        .WindowState = 0
        .WindowTitle = "채혈리스트"
        
        .Action = 1
        .Reset
    End With
   

End Sub

Private Function CollectRecordSet() As Recordset
    Dim sSQL(2) As String
    Dim strSQL  As String

    sSQL(0) = " select a.workarea, a.accdt, a.accseq, a.ptid, a.sex, a.ageday, a.deptcd, c.workdt as coldt, " & _
                " c.worktm as coltm, a.storecd, a.spccd, a.testdiv, a.buildcd,  d.field5 as SpcNm ,c.spcyy,c.spcno,a.statfg," & _
                " '' as orddt, 0 as ordno, i.roomno1 as hosilid," & _
                " 'L' as Div " & _
                " from " & _
                         T_LAB201 & " a, " & T_LAB032 & " d, " & T_LAB204 & " c, ORAA1.APIPDLST i " & _
                " where " & _
                                  DBW("c.workdt=", pWorkDt) & _
                        " and " & DBW("c.wardid=", pWardId) & _
                        " and " & DBW("c.worktm=", pWorkTm) & _
                        " and " & DBW("c.buildcd=", pBuildCd) & _
                        " and a.workarea = c.workarea" & _
                        " and a.accdt= c.accdt" & _
                        " and a.accseq=c.accseq" & _
                        " " & _
                        " " & _
                        " and " & DBW("d.cdindex=", LC3_Specimen) & _
                        " and d.cdval1 = a.spccd" & _
                        " AND a.ptid = i.patno " & _
                        " AND i.stayyn = 'Y' "
' 2012-04-24
' 현시점 병동/병실
'    sSQL(0) = " select a.workarea, a.accdt, a.accseq, a.ptid, a.sex, a.ageday, a.deptcd, c.workdt as coldt, " & _
'                " c.worktm as coltm, a.storecd, a.spccd, a.testdiv, a.buildcd,  d.field5 as SpcNm ,c.spcyy,c.spcno,a.statfg," & _
'                " '' as orddt, 0 as ordno,a.hosilid," & _
'                " 'L' as Div " & _
'                " from " & _
'                         T_LAB201 & " a, " & T_LAB032 & " d, " & T_LAB204 & " c " & _
'                " where " & _
'                                  DBW("c.workdt=", pWorkDt) & _
'                        " and " & DBW("c.wardid=", pWardId) & _
'                        " and " & DBW("c.worktm=", pWorkTm) & _
'                        " and " & DBW("c.buildcd=", pBuildCd) & _
'                        " and a.workarea = c.workarea" & _
'                        " and a.accdt= c.accdt" & _
'                        " and a.accseq=c.accseq" & _
'                        " " & _
'                        " " & _
'                        " and " & DBW("d.cdindex=", LC3_Specimen) & _
'                        " and d.cdval1 = a.spccd" '& _


    If P_IncludeBBSSystem Then
        '혈액은행
        sSQL(1) = " select c.workarea,c.accdt,c.accseq,a.ptid,'' as sex, 0 as ageday,b.deptcd," & _
                      " c.workdt as coldt,c.worktm as coltm,'' as storecd,'' as spccd,'' as testdiv,c.buildcd,'혈액' as spcnm,c.spcyy ,c.spcno,'' as statfg," & _
                      " b.orddt,b.ordno,b.hosilid," & _
                      " 'B' as Div " & _
                      " from " & T_BBS102 & " a," & T_LAB101 & " b," & T_LAB204 & " c" & _
                      " where " & _
                                      DBW("c.workdt=", pWorkDt) & _
                            " and " & DBW("c.wardid=", pWardId) & _
                            " and " & DBW("c.worktm=", pWorkTm) & _
                            " and " & DBW("c.buildcd=", pBuildCd) & _
                            " and c.spcyy=a.spcyy and c.spcno=a.spcno" & _
                            " and a.ptid=b.ptid and a.orddt=b.orddt and a.ordno=b.ordno"
    End If
    
    '해부병리
    sSQL(2) = " "
    
    strSQL = sSQL(0)
    If P_IncludeBBSSystem Then
        strSQL = strSQL & " Union all " & sSQL(1)
    End If
    
    strSQL = strSQL & " order by hosilid,ptid,buildcd,workarea,accdt,accseq"
    
    Set CollectRecordSet = New Recordset
    CollectRecordSet.Open strSQL, DBConn

End Function

'-----------------------
'이하 채혈리스트 재출력
'-----------------------
Public Function CollectQueryTF() As Boolean
    Dim Rs          As Recordset
    Dim RsName      As Recordset
    Dim i      As Integer
    Dim sSQL   As String
    Dim strStatFg As String
    Dim sAge    As String
    Dim strcoldt As String
    Dim strColtm As String
    Dim WorkTestCd As String
    Dim strWorkInfo As String
    Dim strSpcNm As String
    Dim sSex As String
    Dim strTmp As String
    
    Set Rs = CollectRecordSet
    
    If Rs.EOF Then
        CollectQueryTF = False
    Else
        With Rs
            .MoveFirst
            For i = 1 To .RecordCount
                strTmp = ""
                sSQL = " select " & F_PTNM & " as ptnm," & F_DOB2 & " as dob ," & F_SSN2 & " as ssn from " & T_HIS001 & _
                       " where " & DBW(F_PTID & "=", "" & .Fields("ptid").Value)
                objDictionary.Sort = False
                Set RsName = New Recordset
                RsName.Open sSQL, DBConn
                
                Select Case .Fields("div").Value & ""
                
                    Case LIS_ORDDIV:
                        If chkTestCD("" & .Fields("WorkArea").Value, "" & .Fields("accdt").Value, "" & .Fields("accseq").Value, _
                                     "" & .Fields("testdiv").Value) = True Then                   ' Exists
                            If .Fields("statfg").Value & "" = "" Or .Fields("statfg").Value & "" = "0" Then
                                strStatFg = ""
                            Else
                                strStatFg = "**"
                            End If
                            
                            sAge = ("" & .Fields("AgeDay").Value \ 365) + 1
                            strcoldt = Mid(.Fields("coldt").Value & "", 3)
                            strColtm = " " & Mid(.Fields("coltm").Value & "", 1, 4)
                            WorkTestCd = GetTestCD_LIS("" & .Fields("WorkArea").Value, "" & .Fields("AccDt").Value, _
                                                       "" & .Fields("AccSeq").Value, "" & .Fields("TestDiv").Value)
                            ' 검체코드: mvartestdiv=0 검체명 mvartestdiv=1
                            strSpcNm = IIf(mvarTestDiv = "0", .Fields("spcnm").Value & "", .Fields("spccd").Value & "")
                            
                            strWorkInfo = .Fields("workarea").Value & "" & "-" & _
                                          Mid("" & .Fields("AccDt").Value, 3) & "-" & _
                                          .Fields("accseq").Value & ""
                            
                            objDictionary.AddNew i & strStatFg, Join(Array(strWorkInfo, RsName.Fields("ptnm").Value & "", .Fields("ptid").Value & "", _
                                                             .Fields("sex").Value & "" & "/" & sAge & "", .Fields("deptcd").Value & "", _
                                                             strcoldt & strColtm, _
                                                             WorkTestCd, strSpcNm, .Fields("hosilid").Value & ""), COL_DIV)
                        End If
                    Case BBS_ORDDIV:
                        sAge = medFindAge(RsName.Fields("dob").Value & "", "Y")
                        Select Case Mid(RsName.Fields("ssn").Value & "", 7, 1)
                            Case 1, 3: sSex = "M"
                            Case 2, 4: sSex = "F"
                            Case Else: sSex = "기타"
                        End Select
                        '응급여부체크
                        strStatFg = Get_statFg(.Fields("ptid").Value & "", .Fields("orddt").Value & "", .Fields("ordno").Value & "")
                        strcoldt = Mid(.Fields("coldt").Value & "", 3)
                        strColtm = " " & Mid(.Fields("coltm").Value & "", 1, 4)
                        WorkTestCd = GetTestCd_BBS(.Fields("ptid").Value & "", .Fields("spcyy").Value & "", .Fields("spcno").Value)
                        ' 검체코드: mvartestdiv=0 검체명 mvartestdiv=1
                        strSpcNm = IIf(mvarTestDiv = "0", .Fields("spcnm").Value & "", "XM")
                        objDictionary.AddNew i & strStatFg, _
                                              Join(Array(.Fields("spcyy").Value & "-" & .Fields("spcno").Value, _
                                                         RsName.Fields("ptnm").Value & "", .Fields("ptid").Value & "", sSex & "/" & sAge, _
                                                         .Fields("deptcd").Value & "", strcoldt & strColtm, _
                                                         WorkTestCd, .Fields("spcnm").Value & "", .Fields("hosilid").Value & ""), COL_DIV)
                End Select
                .MoveNext
            Next
        End With
        CollectQueryTF = True
    End If
    
    Set Rs = Nothing
    Set RsName = Nothing
End Function
Public Function Get_Coltm()
    Dim Rs          As Recordset
    Dim strEmpNm    As String
    Dim sSQL        As String
    
    sSQL = " select distinct worktm,colid from " & T_LAB204 & _
           " where " & _
                     DBW("workdt=", pWorkDt) & _
           " and " & DBW("wardid=", pWardId) & _
           " and " & DBW("buildcd=", pBuildCd)
    
    Set Rs = New Recordset
    Rs.Open sSQL, DBConn
    
'    If rs.DBerror Then
'        Get_Coltm = ""
'        GoTo DataNull
'    End If
    
    If Not Rs.EOF Then
        Do Until Rs.EOF
'            Dim objEmp As clsBasisData
            
'            Set objEmp = Nothing
'            Set objEmp = New clsBasisData
            strEmpNm = GetEmpNm(Rs.Fields("colid").Value & "")
'            Set objEmp = Nothing
            
'            strEmpNm = GetEmpName(rs.Fields("colid").Value & "")
            Get_Coltm = Get_Coltm & Rs.Fields("worktm").Value & "" & Space(3) & _
                        Rs.Fields("colid").Value & "" & Space(2) & strEmpNm & COL_DIV
            Rs.MoveNext
        Loop
        Get_Coltm = Mid(Get_Coltm, 1, Len(Get_Coltm) - 1)
    Else
        Get_Coltm = ""
    End If
DataNull:
    Set Rs = Nothing
End Function

Public Sub RePrint_CollectList()
    Dim SpcCount As Integer
    Dim strcoldt As String
    Dim strColtm As String
    Dim strEmpNm As String
    
    SpcCount = objDictionary.RecordCount
    strcoldt = Format(GetSystemDate, "yyyy/mm/dd")
    strColtm = " " & Format(GetSystemDate, "hh:mm")
'    Dim objEmp As clsBasisData
'    Set objEmp = New clsBasisData
    strEmpNm = GetEmpNm(ObjSysInfo.EmpId)
'    Set objEmp = Nothing
'    strEmpNm = GetEmpName(objsysinfo.EmpId)
    
    
    Call CollectList_Print(objDictionary, SpcCount, pWardId, strcoldt & strColtm, strEmpNm, pBuildCd, pTitleNm)
    
End Sub

Private Sub Class_Initialize()
    objDictionary.Clear
    objDictionary.FieldInialize "seq", "workno,ptnm,ptid,sexage,io,collectdt,testlist,spcnm,hosilid"
End Sub


'---------------------------
'이하 RI채혈리스트 관련 모듈
'---------------------------
'2001-06-23 KJG
Public Function RI_CollectList(ByVal qWorkarea As String, ByVal qColdt As String, ByVal qColtm As String) As Recordset
    Dim sSQL As String
    
    sSQL = " select a.ptid,a.coldt,coltm,a.workarea,a.accdt,a.accseq,a.sex,d.field5 as spcnm," & _
           " c." & F_PTNM & " as ptnm,f.deptcd,f.wardid,f.hosilid,f.bussdiv" & _
           " from " & _
                      T_LAB101 & " f," & T_LAB102 & " e," & T_LAB032 & " d," & T_HIS001 & " c," & T_LAB201 & " a" & _
           " where " & _
                     DBW("a.workarea=", qWorkarea) & _
           " and " & DBW("a.coldt=", qColdt) & _
           " and a.coltm>=" & DBV("coltm", qColtm) & _
           " and " & DBW("d.cdindex=", LC3_Specimen) & _
           " and d.cdval1 = a.spccd" & _
           " and a.ptid=c." & F_PTID & _
           " and a.workarea=e.workarea and a.accdt=e.accdt and a.accseq=e.accseq" & _
           " and e.ptid=f.ptid and e.orddt=f.orddt and e.ordno=f.ordno" '& _
           " and f.bussdiv in('1','4')"
           
    sSQL = sSQL & _
                " union " & _
                           " select a.ptid,a.coldt,coltm,a.workarea,a.accdt,a.accseq,a.sex,d.field5 as spcnm," & _
                           " c." & F_PTNM & " as ptnm,f.deptcd,f.wardid,f.hosilid,f.bussdiv" & _
                           " from " & _
                                      T_LAB101 & " f," & T_LAB102 & " e," & T_LAB032 & " d," & T_HIS001 & " c," & T_LAB201 & " a" & _
                           " where " & _
                                     DBW("a.workarea=", qWorkarea) & _
                           " and " & DBW("a.coldt>", qColdt) & _
                           " and " & DBW("d.cdindex=", LC3_Specimen) & _
                           " and d.cdval1 = a.spccd" & _
                           " and a.ptid=c." & F_PTID & _
                           " and a.workarea=e.workarea and a.accdt=e.accdt and a.accseq=e.accseq" & _
                           " and e.ptid=f.ptid and e.orddt=f.orddt and e.ordno=f.ordno" '& _
                           " and f.bussdiv in('1','4')"
    sSQL = sSQL & " order by coldt,coltm          "
    Set RI_CollectList = New Recordset
    RI_CollectList.Open sSQL, DBConn
End Function

Public Function GetPtid_SSN(ByVal qPtid As String) As String
    Dim sSQL As String
    Dim Rs   As Recordset
    
    sSQL = " select " & F_SSN2 & " as ssn from " & T_HIS001 & " where " & DBW(F_PTID, qPtid, 2)
    
    Set Rs = New Recordset
    Rs.Open sSQL, DBConn
    
    If Not Rs.EOF Then
        GetPtid_SSN = Rs.Fields("ssn").Value & ""
    Else
        GetPtid_SSN = ""
    End If
    
    Set Rs = Nothing
End Function


Public Function RI_TESTLIST(ByVal sWorkarea As String, ByVal sAccDt As String, _
                            ByVal sAccSeq As String) As String
'-----------------------
'임상병리검사명 불러오기
'-----------------------
    Dim sSQL2 As String
    Dim rsTestCode As Recordset
    Dim i%, tmpiposTestCD
    Dim sTable As String
    
    sSQL2 = " select a.ordcd,b.abbrnm5 " & _
        " from " & T_LAB102 & " a," & T_LAB001 & " b" & _
        " where " & _
                          DBW("a.workarea=", sWorkarea) & _
                " and " & DBW("a.accdt=", Trim(sAccDt)) & _
                " and " & DBW("a.accseq=", Trim(sAccSeq)) & _
                " and a.ordcd=b.testcd"
                
    Set rsTestCode = New Recordset
    rsTestCode.Open sSQL2, DBConn
    
    If rsTestCode.EOF = True Then
        RI_TESTLIST = ""
        GoTo Nodata
    End If
    
    With rsTestCode
        rsTestCode.MoveFirst
        For i = 1 To rsTestCode.RecordCount
            If mvarTestDiv = "1" Then
                RI_TESTLIST = RI_TESTLIST & .Fields("ordcd").Value & "" & ","
            Else
                RI_TESTLIST = RI_TESTLIST & .Fields("abbrnm5").Value & "" & ","
            End If
            rsTestCode.MoveNext
        Next i
        RI_TESTLIST = Mid(RI_TESTLIST, 1, Len(RI_TESTLIST) - 1)
    End With
    
Nodata:
    Set rsTestCode = Nothing
    
End Function
Private Sub Class_Terminate()
    Set objDictionary = Nothing
End Sub
