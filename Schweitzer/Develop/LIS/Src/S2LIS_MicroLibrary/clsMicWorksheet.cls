VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISMicWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private objMicSql As New clsLISSqlMicRst
Private objMicLib As clsLISMicroLib

Public Sub GetSpcGroup(ByRef objSpcDic As clsDictionary)

    Dim objRs As Recordset
    Dim sTmp As String
    Dim sWA As String, sSEQ1 As String, sSEQ2 As String
    
'    objSpcDic.FieldInialize "grpcd", "grpnm,media,workarea,fseq,tseq,rptseq,wsgrp,excfg," & _
                                     "wsunit,fnshdt,fnshtm,count,worksheet,extable,excount"
    Set objRs = New Recordset
    objRs.Open objMicSql.SQLGetSpcGroup, DBConn
    
    With objRs
        objSpcDic.Sort = False
        While (Not objRs.EOF)
            sTmp = .Fields("field2").Value & ""
            sWA = medGetP(sTmp, 1, ";")
            sSEQ1 = medGetP(sTmp, 2, ";")
            sSEQ2 = medGetP(sTmp, 3, ";")
            objSpcDic.AddNew .Fields("cdval1").Value & "", medConcatString(COL_DIV, .Fields("field1").Value & "", "", _
            sWA, sSEQ1, sSEQ2, .Fields("field3").Value & "", .Fields("field4").Value & "", .Fields("field5").Value & "", _
            "", "", "", "", "", "", "")
            .MoveNext
        Wend
    End With
    Set objRs = Nothing
    
End Sub


Public Function DelSpcFromWorksheet(ByVal pDelString As String, _
                                    ByVal pSpcGrp As String, ByVal pWsUnit As String) As Boolean
                              
    Dim blnDel As Boolean
    Dim SqlStmt() As String
    Dim aryLabNo As Variant
    Dim i As Long
    Dim pWorkArea As String, pAccDt As String, pAccSeq As String
    
    aryLabNo = Split(pDelString, COL_DIV)
    
    ReDim SqlStmt(UBound(aryLabNo))
    
    For i = LBound(aryLabNo) To UBound(aryLabNo)
        
        If aryLabNo(i) = "" Then Exit For
        
        pWorkArea = medGetP(aryLabNo(i), 1, "-")
        pAccDt = medGetP(aryLabNo(i), 2, "-")
        pAccDt = IIf(Mid(pAccDt, 1, 1) = "9", "19" & pAccDt, "20" & pAccDt)
        pAccSeq = medGetP(aryLabNo(i), 3, "-")
    
        SqlStmt(i) = objMicSql.SQLDeleteSpc(pWorkArea, pAccDt, pAccSeq, pSpcGrp, pWsUnit)
    Next
    
On Error GoTo Err_Trap
    DBConn.BeginTrans
    For i = LBound(SqlStmt) To UBound(SqlStmt)
        If Trim(SqlStmt(i)) <> "" Then DBConn.Execute SqlStmt(i)
    Next
    DBConn.CommitTrans
        
    DelSpcFromWorksheet = True
    Exit Function

Err_Trap:
    DBConn.RollbackTrans
    DelSpcFromWorksheet = False
    
End Function

Public Function GetAddSpcInfo(ByVal pWorkArea, ByVal pAccDt As String, ByVal pAccSeq As String, _
                              ByVal pSpcGrp As String, ByVal pRTs As String) As String
                              
    Dim Rs As New Recordset
    Dim chkG As String, chkS As String, sTestFlag As String
    Dim strPtId As String, strPtNm As String, strSex As String, strAgeDay As String
    Dim strSpcNm As String, strTestCd As String, strTestNm As String, strRstType As String
    Dim strSpcCd As String
    Dim strRstCd As String
    
    GetAddSpcInfo = ""
        
    Set Rs = New Recordset
    Rs.Open objMicSql.SQLGetAddSpcInfo(pWorkArea, pAccDt, pAccSeq, pSpcGrp, pRTs), DBConn
    
    With Rs
        If Not .EOF Then
            strPtId = .Fields("ptid").Value & ""
            strPtNm = .Fields("ptnm").Value & ""
            strSex = .Fields("sex").Value & ""
            strAgeDay = .Fields("ageday").Value & ""
            strSpcNm = .Fields("spcnm").Value & ""
            strTestCd = .Fields("testcd").Value & ""
            strTestNm = .Fields("testnm").Value & ""
            strRstType = .Fields("rsttype").Value & ""
            strSpcCd = .Fields("spccd").Value & ""
            strRstCd = .Fields("rstcd").Value & ""
'            Select Case .Fields("rsttype").Value & ""
'                Case MRT_Stain:   chkG = MRT_Stain
'                Case MRT_GenSen:  chkS = MRT_GenSen
'                Case MRT_MicSen:  chkS = MRT_MicSen
'                Case MRT_AFC:     chkG = MRT_AFC
'                Case MRT_AFS:     chkG = MRT_AFS
'            End Select
'
'            Select Case True
'                Case chkG = MRT_Stain And (chkS = MRT_GenSen Or chkS = MRT_MicSen): sTestFlag = MRT_Both
'                Case chkG = "" And chkS = MRT_GenSen: sTestFlag = MRT_GenSen
'                Case chkG = "" And chkS = MRT_MicSen: sTestFlag = MRT_MicSen
'                Case chkG = MRT_Stain And chkS = "":  sTestFlag = MRT_Stain
'                Case chkG = MRT_AFC And chkS = "":    sTestFlag = MRT_AFC
'                Case chkG = MRT_AFS And chkS = "":    sTestFlag = MRT_AFS
'                Case chkG = "" And chkS = "":         sTestFlag = ""
'            End Select
    
            Select Case .Fields("rsttype").Value & ""
                Case MRT_Stain:   chkG = MRT_Stain
                Case MRT_GenSen:  chkS = MRT_GenSen
                Case MRT_MicSen:  chkS = MRT_MicSen
                Case MRT_AFC:     chkG = MRT_AFC
                Case MRT_AFS:     chkG = MRT_AFS
                Case MRT_AFN:     chkG = MRT_AFN
                Case MRT_Both:    chkG = MRT_Both
                Case MRT_OTHER:   chkG = MRT_OTHER
                Case MRT_FUC:     chkG = MRT_FUC
                
            End Select
            
    
             Select Case True
                 Case chkG = MRT_Stain And (chkS = MRT_GenSen Or chkS = MRT_MicSen)
                             sTestFlag = MRT_Both
                 Case chkG = "" And chkS = MRT_GenSen: sTestFlag = MRT_GenSen
                 Case chkG = "" And chkS = MRT_MicSen: sTestFlag = MRT_MicSen
                 Case chkG = MRT_Stain And chkS = "":  sTestFlag = MRT_Stain
                 Case chkG = MRT_AFC And chkS = "":    sTestFlag = MRT_AFC
                 Case chkG = MRT_AFS And chkS = "":    sTestFlag = MRT_AFS
                 Case chkG = MRT_AFN And chkS = "":    sTestFlag = MRT_AFN
                 Case chkG = MRT_FUC And chkS = "":    sTestFlag = MRT_FUC
                 Case chkG = "" And chkS = "":         sTestFlag = ""
             End Select
            
            .MoveNext
            While Not .EOF
                strTestCd = strTestCd & "','" & .Fields("testcd").Value & ""
                strTestNm = strTestNm & "," & .Fields("testnm").Value & ""
                .MoveNext
            Wend
    
            GetAddSpcInfo = strPtId & COL_DIV & strPtNm & COL_DIV & strSex & COL_DIV & strAgeDay & COL_DIV & strSpcNm & COL_DIV
            GetAddSpcInfo = GetAddSpcInfo & strTestCd & COL_DIV & strTestNm & COL_DIV & strRstType & COL_DIV
            GetAddSpcInfo = GetAddSpcInfo & sTestFlag & COL_DIV & strSpcCd & COL_DIV & strRstCd
                                 
        End If
    End With
    Set Rs = Nothing

End Function

Public Function SaveAddSpc(ByVal pWorkArea, ByVal pAccDt As String, ByVal pAccSeq As String, _
                           ByVal pSpcGrp As String, ByVal pWsUnit As String, ByVal pRTs As String) As Boolean

    Dim strSql As String
    
    strSql = objMicSql.SQLSaveAddSpc(pWorkArea, pAccDt, pAccSeq, pSpcGrp, pWsUnit, pRTs)
On Error GoTo Err_Trap
    DBConn.BeginTrans
    DBConn.Execute strSql
    DBConn.CommitTrans
    SaveAddSpc = True
    Exit Function
Err_Trap:
    DBConn.RollbackTrans
    SaveAddSpc = False
    
End Function

Public Sub GetLastWsUnit(ByVal sWsCd As String, ByRef sWsUnit As String, _
                         ByRef sWsDt As String, ByRef sWsTm As String)
                         
    Dim objRs As Recordset
    
    Set objRs = New Recordset
    objRs.Open objMicSql.SQLLastWsUnit(sWsCd), DBConn
    
    If objRs.EOF Then
        sWsUnit = ""
        sWsDt = ""
        sWsTm = ""
    Else
        sWsUnit = objRs.Fields("wsunit").Value
        sWsDt = objRs.Fields("fnshdt").Value
        sWsTm = objRs.Fields("fnshtm").Value
    End If
    Set objRs = Nothing
    
End Sub

Public Function GetMedias(ByVal pSGCD As String) As String
    
    Dim i As Integer, sSQL1 As String, dsMD As Recordset
    Dim sMDBuf As String

    ' 적용배지 Load
    Set dsMD = New Recordset
    dsMD.Open objMicSql.SQLGetMedias(pSGCD), DBConn
    
    If dsMD.EOF Then GetMedias = "":   Set dsMD = Nothing: Exit Function
    
    sMDBuf = ""
    dsMD.MoveFirst
    For i = 1 To dsMD.RecordCount
        If i = 1 Then
            sMDBuf = sMDBuf & "" & dsMD.Fields("media").Value & ""
        Else
            sMDBuf = sMDBuf & ", " & dsMD.Fields("media").Value & ""
        End If
        dsMD.MoveNext
    Next i
    
    Set dsMD = Nothing

    GetMedias = sMDBuf
    
End Function


Public Function GetRTypes(ByVal pSG As String) As String
    
    Dim dsRT As New Recordset
    Dim sRType() As String

    GetRTypes = ""
    dsRT.Open objMicSql.SQLGetRstType(pSG), DBConn
    
    If dsRT.EOF Then GoTo NoData
     
    Dim i As Integer
    
    sRType = Split(dsRT.Fields("field2").Value & "", ",")
    For i = LBound(sRType) To UBound(sRType)
        sRType(i) = "'" & sRType(i) & "'"
    Next
    GetRTypes = Join(sRType, ",")
     
NoData:
    Set dsRT = Nothing
   
End Function
    
Public Function GetWorkList(ByVal pWACD As String, ByVal pACCSEQ1 As Integer, ByVal pACCSEQ2 As String, _
                            ByVal pFR As String, ByVal pTO As String, ByVal pRTs As String, _
                            ByRef strFinDt As String, ByRef strFinTm As String, ByRef objWS As clsDictionary) As Long
    
    Dim strSql As String
    
    strSql = objMicSql.SQLGetWorkList(pWACD, pACCSEQ1, pACCSEQ2, pFR, pTO, pRTs)
    GetWorkList = MakeWorksheet(strSql, strFinDt, strFinTm, objWS)
    
End Function
    
Public Function GetWorkList_New(ByVal pWACD As String, ByVal pACCSEQ1 As String, ByVal pACCSEQ2 As String, _
                            ByVal pFR As String, ByVal pTO As String, ByVal pRTs As String, _
                            ByRef strFinDt As String, ByRef strFinTm As String, ByRef objWS As clsDictionary) As Long
    
    Dim strSql As String
    
    strSql = objMicSql.SQLGetWorkList(pWACD, pACCSEQ1, pACCSEQ2, pFR, pTO, pRTs)
    GetWorkList_New = MakeWorksheet(strSql, strFinDt, strFinTm, objWS)
    
End Function

Public Function GetExceptList(ByVal pWsCd As String, ByVal pWsUnit As String, _
                              ByRef strFinDt As String, ByRef strFinTm As String, ByRef objWS As clsDictionary) As Long
    
    Dim strSql As String
    
    strSql = objMicSql.SQLGetExcpList(pWsCd, pWsUnit)
    GetExceptList = MakeWorksheet(strSql, strFinDt, strFinTm, objWS)
    
End Function
        
Private Function MakeWorksheet(ByVal strSqlStmt As String, ByRef strFinDt As String, _
                               ByRef strFinTm As String, ByRef objWS As clsDictionary) As Long
    
    Dim dsWS        As New Recordset
    Dim sFinDTTM    As String
    Dim strFinDTTM  As String
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String
    Dim sAccNo      As String
    Dim sTestFlag   As String
    Dim sTest       As String
    Dim sTCd        As String
    Dim chkG        As String
    Dim chkS        As String
    Dim strData     As String
    Dim strAccNo    As String
    
    Dim blnFirst    As Boolean
        
    Dim sICSString  As String
    
        
    Set objMicLib = New clsLISMicroLib
    
    dsWS.Open strSqlStmt, DBConn
    
    MakeWorksheet = 0
    strFinDTTM = strFinDt & strFinTm
    sAccNo = "": sTest = "": sTCd = "": chkG = "": chkS = ""
    
    If Not dsWS.EOF Then
        While Not dsWS.EOF
            
            sICSString = ICSPatientString(dsWS.Fields("ptid").Value, enICSNum.LIS_ALL)
            
            If strAccNo <> dsWS.Fields("workarea").Value & "" & dsWS.Fields("accdt").Value & "" & dsWS.Fields("accseq").Value & "" Then
                blnFirst = False
            End If
            strAccNo = dsWS.Fields("workarea").Value & "" & dsWS.Fields("accdt").Value & "" & dsWS.Fields("accseq").Value & ""
            
            sWorkArea = dsWS.Fields("workarea").Value & ""
            sAccDt = dsWS.Fields("accdt").Value & ""
            sAccSeq = dsWS.Fields("accseq").Value & ""
            sAccNo = sWorkArea & "-" & Mid$(sAccDt, 3, Len(sAccDt) - 2) & "-" & sAccSeq
            
            If objWS.Exists(sAccNo) Then
                objWS.KeyChange sAccNo
                objWS.Fields("testcd") = objWS.Fields("testcd") & ";" & dsWS.Fields("testcd").Value & ""
                objWS.Fields("testnm") = objWS.Fields("testnm") & ";" & dsWS.Fields("abbrnm10").Value & ""
            Else
                With dsWS
                    chkG = "": chkS = ""
'최근결과를 보여주기위한 수정
'Modify By Legends 2003/08/11 졸라게 더운 여름날에

'                    strData = medConcatString(COL_DIV, _
'                                        .Fields("rcvdt").Value, .Fields("rcvtm").Value, .Fields("ptid").Value, .Fields("ptnm").Value, _
'                                        .Fields("sex").Value & "/" & (Val(.Fields("ageday").Value) \ 365), .Fields("wardid").Value, _
'                                        .Fields("spccd").Value, .Fields("spcnm").Value, .Fields("testcd").Value, .Fields("abbrnm10").Value, _
'                                        .Fields("rsttype").Value, chkG, chkS, "", "", sWorkArea, sAccDt, sAccSeq, .Fields("orddt").Value, _
'                                        .Fields("ordno").Value, .Fields("ordseq").Value)
                    
    ' 2008.10.23. 양성현 바코드 출력을 위한 조건을 삽입
                    
                    strData = medConcatString(COL_DIV, _
                                        .Fields("rcvdt").Value, .Fields("rcvtm").Value, .Fields("ptid").Value, .Fields("ptnm").Value & sICSString, _
                                        .Fields("sex").Value & "/" & (Val(.Fields("ageday").Value) \ 365), .Fields("wardid").Value, _
                                        .Fields("spccd").Value, .Fields("spcnm").Value, " ", " ", .Fields("testnm").Value, .Fields("testcd").Value, _
                                        .Fields("rsttype").Value, chkG, chkS, "", "", sWorkArea, sAccDt, .Fields("barcode").Value, sAccSeq, .Fields("orddt").Value, _
                                        .Fields("ordno").Value, .Fields("ordseq").Value)
                                        
                    objWS.AddNew sAccNo, strData
                    sFinDTTM = "" & .Fields("rcvdt").Value & "" & .Fields("rcvtm").Value & ""  '마지막 접수일시
                    If sFinDTTM > strFinDTTM Then strFinDTTM = sFinDTTM
                End With
                MakeWorksheet = MakeWorksheet + 1
            End If
            
            Select Case dsWS.Fields("rsttype").Value & ""
                Case MRT_Stain:   chkG = MRT_Stain
                Case MRT_GenSen:  chkS = MRT_GenSen
                Case MRT_MicSen:  chkS = MRT_MicSen
                Case MRT_AFC:     chkG = MRT_AFC
                Case MRT_AFS:     chkG = MRT_AFS
                Case MRT_AFN:     chkG = MRT_AFN
                Case MRT_FUC:     chkG = MRT_FUC
                Case MRT_FUS:     chkG = MRT_FUS
                Case MRT_Both:    chkG = MRT_Both
                Case MRT_OTHER:   chkG = MRT_OTHER
            End Select
            
            objWS.Fields("chkgram") = chkG
            objWS.Fields("chkculture") = chkS
            
            Select Case True
                Case chkG = MRT_Stain And (chkS = MRT_GenSen Or chkS = MRT_MicSen)
                    sTestFlag = MRT_Both
                Case chkG = "" And chkS = MRT_GenSen: sTestFlag = MRT_GenSen
                Case chkG = "" And chkS = MRT_MicSen: sTestFlag = MRT_MicSen
                Case chkG = MRT_Stain And chkS = "":  sTestFlag = MRT_Stain
                Case chkG = MRT_AFC And chkS = "":    sTestFlag = MRT_AFC
                Case chkG = MRT_AFS And chkS = "":    sTestFlag = MRT_AFS
                Case chkG = MRT_AFN And chkS = "":    sTestFlag = MRT_AFN
                Case chkG = MRT_FUC And chkS = "":    sTestFlag = MRT_FUC
                Case chkG = MRT_FUS And chkS = "":    sTestFlag = MRT_FUS
                Case chkG = "" And chkS = "":         sTestFlag = ""
            End Select
            
            objWS.Fields("testfg") = sTestFlag
            
            If blnFirst = False Then
                With objMicLib
                    .Workarea = sWorkArea
                    .Accdt = sAccDt
                    .Accseq = sAccSeq
                    .SpcCd = objWS.Fields("spccd")
                    .PtId = objWS.Fields("ptid")
                    .TestCd = objWS.Fields("testcd")
                End With
            
                objWS.Fields("lastrstcd") = objMicLib.GetNoGrowthLatestRstCd
                objWS.Fields("lastrstnm") = objMicLib.GetNoGrowthRst(objWS.Fields("lastrstcd"))
                
                blnFirst = True
            End If
            
            dsWS.MoveNext
        Wend
                   
        strFinDt = Mid(strFinDTTM, 1, 8)
        strFinTm = Mid(strFinDTTM, 9)
        
    End If
    
    Set dsWS = Nothing
    Set objMicLib = Nothing
End Function

Public Function GetWsUnitNo(ByVal sWA As String, ByVal sSG As String, ByVal sYY As String, _
                            ByRef sSeq As Long, ByRef sUNO As String) As Boolean
    
    Dim dsNo As Recordset
    Dim sql099 As String
    
    ' 번호부여정보 연도가 바뀌었는지 체크 (update or insert)
    Set dsNo = New Recordset
    dsNo.Open objMicSql.SQLGetWsUnitNo(sWA, sSG, sYY), DBConn

    ' 미리 만들어 둔 번호를 즉시 갱신
    If dsNo.EOF Then
        sSeq = 1
        sUNO = sYY & "0001"
        sql099 = " INSERT INTO " & T_COM099 & " (noindex,divcd1,divcd2,divcd3,seq) " & _
                 " VALUES (" & DBV("noindex", COM99_LIS_WSUnit, 1) & DBV("divcd1", sWA, 1) & _
                               DBV("divcd2", sSG, 1) & DBV("divcd3", sYY, 1) & DBV("seq", sSeq) & ")"
    Else
        sSeq = Val("" & dsNo.Fields("seq").Value) + 1
        sUNO = sYY & Format(sSeq, "0000")
        sql099 = " UPDATE " & T_COM099 & " SET " & DBW("seq = ", sSeq) & _
                 " WHERE  " & DBW("noindex = ", COM99_LIS_WSUnit) & _
                 " AND    " & DBW("divcd1  = ", sWA) & _
                 " AND    " & DBW("divcd2  = ", sSG) & _
                 " AND    " & DBW("divcd3  = ", sYY)
    End If
'
    Set dsNo = Nothing
    
    DBConn.BeginTrans
    DBConn.Execute sql099
    DBConn.CommitTrans
    GetWsUnitNo = True
    Exit Function
    
Err_Trap:
    DBConn.RollbackTrans
    GetWsUnitNo = False
    
End Function

Public Function SetWorksheetH(ByVal objGrp As clsDictionary, ByVal pSysDate As String, ByVal pEmpId As String) As Boolean

    Dim strSql As String
    
    On Error GoTo Err_Trap
    
    With objGrp
        strSql = " INSERT INTO " & T_LAB401 & " (wscd, wsunit, fnshdt, fnshtm, workdt, worktm, workid) " & _
                 " VALUES (" & DBV("wscd", .Fields("grpcd"), 1) & DBV("wsunit", .Fields("wsunit"), 1) & _
                               DBV("fnshdt", .Fields("fnshdt"), 1) & DBV("fnshtm", .Fields("fnshtm"), 1) & _
                               DBV("workdt", Mid(pSysDate, 1, 8), 1) & DBV("wscd", Mid(pSysDate, 10, 6), 1) & _
                               DBV("workid", pEmpId) & ")"
    End With
    DBConn.Execute strSql
    SetWorksheetH = True
    Exit Function
    
Err_Trap:
    SetWorksheetH = False

End Function

Public Function SetWorksheetB(ByVal sWsCd As String, ByVal sUNO As String, ByVal sWorkArea As String, _
                              ByVal sAccDt As String, sAccSeq As String, sSCFlag As String) As Boolean

            
    Dim strSql As String
    
    On Error GoTo Err_Trap
    
    strSql = " INSERT INTO " & T_LAB402 & " (wscd, wsunit, workarea, accdt, accseq, selfg, genfg, scfg, remark) " & _
             " VALUES (" & DBV("wscd", sWsCd, 1) & DBV("wsunit", sUNO, 1) & DBV("workarea", sWorkArea, 1) & _
                           DBV("accdt", sAccDt, 1) & DBV("accseq", sAccSeq, 1) & DBV("selfg", MWS_Ready, 1) & _
                           DBV("genfg", MWS_Ready, 1) & DBV("scfg", sSCFlag, 1) & DBV("remark", Format(Now, CS_DateDbFormat & " " & CS_TimeDbFormat)) & ")"
'
    DBConn.Execute strSql
    SetWorksheetB = True
    Exit Function
    
Err_Trap:
    SetWorksheetB = False

End Function
 
Public Function SetExceptList(ByVal sWsCd As String, ByVal sUNO As String, ByVal sWorkArea As String, _
                              ByVal sAccDt As String, sAccSeq As String) As Boolean

            
    Dim strSql As String
    
    On Error GoTo Err_Trap
    
    strSql = " INSERT INTO " & T_LAB403 & " (wscd, wsunit, workarea, accdt, accseq) " & _
             " VALUES (" & DBV("wscd", sWsCd, 1) & DBV("wsunit", sUNO, 1) & DBV("workarea", sWorkArea, 1) & _
                           DBV("accdt", sAccDt, 1) & DBV("accseq", sAccSeq) & ")"
            
    DBConn.Execute strSql
    SetExceptList = True
    Exit Function
    
Err_Trap:
    SetExceptList = False

End Function
 
Public Function SetStatus(ByVal sWorkArea As String, ByVal sAccDt As String, sAccSeq As String, sTCd As String) As Boolean

    Dim strSql As String

    On Error GoTo Err_Trap

    ' 접수 내역 Update
    strSql = " UPDATE " & T_LAB201 & _
             " SET    " & DBW("stscd    = ", enStsCd.StsCd_LIS_InProcess) & _
             " WHERE  " & DBW("workarea = ", sWorkArea) & _
             " AND    " & DBW("accdt    = ", sAccDt) & _
             " AND    " & DBW("accseq   = ", sAccSeq)
    DBConn.Execute strSql
    
    '** 원본 ===========================================================================
    ' 처방 Body Update
'    strSql = " UPDATE " & T_LAB102 & _
'             " SET    " & DBW("stscd    = ", enStsCd.StsCd_LIS_InProcess) & _
'             " WHERE  " & DBW("workarea = ", sWorkArea) & _
'             " AND    " & DBW("accdt    = ", sAccDt) & _
'             " AND    " & DBW("accseq   = ", sAccSeq) & _
'             " AND    ordcd  in (" & sTCd & ")"
    '===================================================================================
    
    '** 전주 예수 병원 변경 루틴========================================================
    Dim strWrkDiv   As String
    Dim Rs          As New ADODB.Recordset
    
    strSql = " select wrkdiv from " & T_LAB102 & _
             "  where workarea = " & DBS(sWorkArea) & _
             "    and accdt = " & DBS(sAccDt) & _
             "    and accseq = " & DBN(sAccSeq)
             
    Rs.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If Rs.EOF = False Then
        strWrkDiv = "" & Rs.Fields("wrkdiv").Value
    Else
        strWrkDiv = ""
    End If
    
    Rs.Close: Set Rs = Nothing
    
    ' 처방 Body Update
    '## 5.0.14: 이상대(2005-06-24)
    '   - 종건, 일건 컬럼이 틀린것을 수정
    If strWrkDiv = "3" Then         '종건
        strSql = " UPDATE su2examt " & _
                 " SET    " & DBW("stscd    = ", enStsCd.StsCd_LIS_InProcess) & _
                 " WHERE  " & DBW("workarea = ", sWorkArea) & _
                 " AND    " & DBW("accdt    = ", sAccDt) & _
                 " AND    " & DBW("accseq   = ", sAccSeq) & _
                 " AND    examcode in (" & sTCd & ")"
    ElseIf strWrkDiv = "4" Then     '일건
            strSql = " UPDATE sg2examt " & _
                     " SET    " & DBW("stscd    = ", enStsCd.StsCd_LIS_InProcess) & _
                     " WHERE  " & DBW("workarea = ", sWorkArea) & _
                     " AND    " & DBW("accdt    = ", sAccDt) & _
                     " AND    " & DBW("accseq   = ", sAccSeq) & _
                     " AND    examcode in (" & sTCd & ")"
    Else
        strSql = " UPDATE mdexmort " & _
                 " SET    " & DBW("stscd    = ", enStsCd.StsCd_LIS_InProcess) & _
                 " WHERE  " & DBW("workarea = ", sWorkArea) & _
                 " AND    " & DBW("accdt    = ", sAccDt) & _
                 " AND    " & DBW("accseq   = ", sAccSeq) & _
                 " AND    ordcd  in (" & sTCd & ")"
    End If
    '===================================================================================
    
    DBConn.Execute strSql

    ' 결과 내역 Update
    strSql = " UPDATE " & T_LAB404 & _
             " SET    " & DBW("stscd    = ", enStsCd.StsCd_LIS_InProcess) & _
             " WHERE  " & DBW("workarea = ", sWorkArea) & _
             " AND    " & DBW("accdt    = ", sAccDt) & _
             " AND    " & DBW("accseq   = ", sAccSeq) & _
             " AND    testcd  in (" & sTCd & ")"
    DBConn.Execute strSql

    SetStatus = True
    Exit Function

Err_Trap:
    SetStatus = False

End Function

Public Function SetBodyStatus(ByVal sPtId As String, ByVal sOrdDt As String, sOrdNo As String, sOrdSeq As String) As Boolean

    Dim strSql      As String
    Dim strExamDt   As String
    
    On Error GoTo Err_Trap
    
    '** 원본 ===========================================================================
    ' 처방 Body Update
'    strSql = " UPDATE " & T_LAB102 & _
'             " SET    " & DBW("stscd  = ", enStsCd.StsCd_LIS_InProcess) & _
'             " WHERE  " & DBW("ptid   = ", sPtId) & _
'             " AND    " & DBW("orddt  = ", sOrdDt) & _
'             " AND    " & DBW("ordno  = ", sOrdNo) & _
'             " AND    " & DBW("ordseq = ", sOrdSeq) & _
'             " AND    " & DBW("stscd  < ", enStsCd.StsCd_LIS_InProcess)
    '===================================================================================
    
    '** 전주 예수 병원 변경 루틴========================================================
    Dim strWrkDiv   As String
    Dim Rs          As New ADODB.Recordset
    
    strSql = " select wrkdiv from " & T_LAB102 & _
             "  where ptid = " & DBS(sPtId) & _
             "    and orddt = " & DBS(sOrdDt) & _
             "    and ordno = " & DBN(sOrdNo)
             
    Rs.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If Rs.EOF = False Then
        strWrkDiv = "" & Rs.Fields("wrkdiv").Value
    Else
        strWrkDiv = ""
    End If
    
    Rs.Close: Set Rs = Nothing
    
    '-- 실시일자
    strExamDt = "TO_DATE(" & DBS(Format$(GetSystemDate, "YYYYMMDD HH:MM:SS")) & ", 'yyyymmdd hh24:mi:ss')"
    
    '## 5.0.14: 이상대(2005-06-24)
    '   - 종건, 일건 컬럼이 틀린것을 수정
    ' 처방 Body Update
    If strWrkDiv = "3" Then         '종건
        strSql = " UPDATE su2examt " & _
                 " SET    " & DBW("stscd  = ", enStsCd.StsCd_LIS_InProcess) & _
                 ", examdate = " & strExamDt & _
                 " WHERE  patno = " & DBS(sPtId) & _
                 " AND    orddate = TO_DATE(" & DBS(sOrdDt) & ", 'yyyymmdd') " & _
                 " AND    ordseqno = " & DBN(sOrdNo) & _
                 " AND    " & DBW("stscd  < ", enStsCd.StsCd_LIS_InProcess)
    ElseIf strWrkDiv = "4" Then     '일건
            strSql = " UPDATE sg2examt " & _
                     " SET    " & DBW("stscd  = ", enStsCd.StsCd_LIS_InProcess) & _
                     ", examdate = " & strExamDt & _
                     " WHERE  patno = " & DBS(sPtId) & _
                     " AND    orddate = TO_DATE(" & DBS(sOrdDt) & ", 'yyyymmdd') " & _
                     " AND    ordseqno = " & DBN(sOrdNo) & _
                     " AND    " & DBW("stscd  < ", enStsCd.StsCd_LIS_InProcess)
    Else
        strSql = " UPDATE mdexmort " & _
                 " SET    " & DBW("stscd  = ", enStsCd.StsCd_LIS_InProcess) & _
                 ", execdate = " & strExamDt & _
                 " WHERE  patno = " & DBS(sPtId) & _
                 " AND    orddate = TO_DATE(" & DBS(sOrdDt) & ", 'yyyymmdd') " & _
                 " AND    ordseqno = " & DBN(sOrdNo) & _
                 " AND    " & DBW("stscd  < ", enStsCd.StsCd_LIS_InProcess)
    End If
    '===================================================================================
    DBConn.Execute strSql

    SetBodyStatus = True
    Exit Function

Err_Trap:
    SetBodyStatus = False

End Function

Private Sub Class_Initialize()
    Set objMicSql = Nothing
End Sub
