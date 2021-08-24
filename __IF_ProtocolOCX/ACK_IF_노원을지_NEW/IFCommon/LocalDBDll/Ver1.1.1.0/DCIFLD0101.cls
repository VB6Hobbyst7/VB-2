VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCIFLD0201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim adoTransCn As ADODB.Connection
Private mCurItemCnt As Integer
Private mPersonCnt As Integer

Public Function WMerge_MDB(ByVal sMachineCd As String, ByVal sWDate As String, ByVal sWSeq1 As String, ByVal sWSeq2 As String) As Long
On Error GoTo ErrorHandler

    Dim ADOCN As New ADODB.Connection
    Dim sSqldoc As String
    
    ADOCN.Open fGetCurDSN(sMachineCd)
    
    sSqldoc = " update IFRESULT set WSEQ ='" & sWSeq1 & "' " _
            & "  where WDATE = '" & sWDate & " '" _
            & "    and WSEQ  = '" & sWSeq2 & " '"
    ADOCN.Execute sSqldoc
    
    ADOCN.Close
    Set ADOCN = Nothing
    
    Exit Function
ErrorHandler:
    'WMerge_MDB = 0
    MsgBox "ERR(DCIFLD) : WMerge_MDB(" & Err.Description & ", " & Err.Number & ")"
End Function

Public Property Get CurItemCnt() As Integer
    CurItemCnt = mCurItemCnt
End Property

Public Property Get PersonCnt() As Integer
    PersonCnt = mPersonCnt
End Property

Public Function Del_IFTestItem(ByVal sMachineCd As String, ByVal iMode As Integer) As String
    On Error GoTo ErrHandler
    
    Dim sSql$
    Dim ADOCN As ADODB.Connection
    
    If iMode = 0 Then
        sSql = "DELETE * FROM IFTESTITEM "
    Else
    End If
    
    Set ADOCN = New ADODB.Connection
    
    ADOCN.Open fGetCurDSN(sMachineCd)
    
    ADOCN.Execute sSql
    
    Del_IFTestItem = "SUCCESS"
    
    ADOCN.Close
    Set ADOCN = Nothing
        
    Exit Function
ErrHandler:
    Del_IFTestItem = "0"
    MsgBox "ERR(DCIFLD) : Del_IFTestItem(" & Err.Description & ", " & Err.Number & ")"
End Function

Public Function Del_CalTestItem(ByVal sMachineCd As String, ByVal iMode As Integer) As String
    On Error GoTo ErrHandler
    
    Dim sSql$
    Dim ADOCN As ADODB.Connection
    
    If iMode = 0 Then
        sSql = "DELETE * FROM CALTESTITEM "
    Else
    End If
    
    Set ADOCN = New ADODB.Connection
    
    ADOCN.Open fGetCurDSN(sMachineCd)
    
    ADOCN.Execute sSql
    
    Del_CalTestItem = "SUCCESS"
    
    ADOCN.Close
    Set ADOCN = Nothing
        
    Exit Function
ErrHandler:
    Del_CalTestItem = "0"
    MsgBox "ERR(DCIFLD) : Del_CalTestItem(" & Err.Description & ", " & Err.Number & ")"
End Function

Public Function Get_IFTestItem(ByVal sMachineCd As String, ByVal iMode As Integer) As String
    On Error GoTo ErrHandler
    
    Dim ADORS As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
    
    sSql = ""
    
    If iMode = 0 Then
        sSql = sSql & " SELECT IFTESTSEQ, IFTESTNM, IFORDCD, IFRSTCD, IFSPCCD, IFSVRCD, DOTDIGIT, LHU, "
        sSql = sSql & "        JUDGEGBN, MREF1, MREF2, FREF1, FREF2, PANICL, PANICH, DELTAGBN, DELTAL, DELTAH, "
        sSql = sSql & "        LIMIT1GBN, LIMIT1, LIMIT2GBN, LIMIT2 "
        sSql = sSql & "   FROM IFTESTITEM "
        sSql = sSql & " ORDER BY IFTESTSEQ "
    Else
    End If
    
    ADORS.CursorType = adOpenStatic
    ADORS.Open sSql, fGetCurDSN(sMachineCd)
    
    If ADORS.EOF = False Then
        Get_IFTestItem = ADORS.RecordCount & Chr(3) & ADORS.GetString(adClipString, -1, Chr(124), Chr(3))
    Else
        Get_IFTestItem = "NONE"
    End If
    
    ADORS.Close
    
    Exit Function
ErrHandler:
    If ADORS.State = 1 Then
        ADORS.Close
    End If
    MsgBox "ERR(DCIFLD) : Get_IFTestItem(" & Err.Description & ", " & Err.Number & ")"
End Function

Public Function Add_IFFlagInfo(ByVal sMachineCd As String, _
                            ByVal sTSeq As String, ByVal sTFlag As String, ByVal sTFlagInfo, ByVal sTDispCd As String, _
                            ByVal sTUseYN As String, ByVal sTRmk As String) As String
    On Error GoTo ErrHandler
    
    Dim ADOCN   As New ADODB.Connection
    Dim SqlStr  As String
    Dim i%
    Dim aSeq()  As String
    Dim aFlag() As String
    Dim aFlagInfo() As String
    Dim aDispCd()   As String
    Dim aUseYN()    As String
    Dim aRmk()  As String
    
    ADOCN.Open fGetCurDSN(sMachineCd)
    
    SqlStr = " delete from ifflaginfo "
    
    ADOCN.Execute SqlStr
    
    aSeq() = Split(sTSeq, Chr(124))
    aFlag() = Split(sTFlag, Chr(124))
    aFlagInfo() = Split(sTFlagInfo, Chr(124))
    aDispCd() = Split(sTDispCd, Chr(124))
    aUseYN() = Split(sTUseYN, Chr(124))
    aRmk() = Split(sTRmk, Chr(124))
    
    For i = 0 To UBound(aSeq())
        If Trim(aSeq(i)) = "" Then Exit For
           
        SqlStr = " insert into ifflaginfo " _
                & "(flagseq, flagcd, flaginfo, dispcd, useyn, remark) values ( "
        SqlStr = SqlStr & "'" & Trim(aSeq(i)) & "', "
        SqlStr = SqlStr & "'" & Trim(aFlag(i)) & "', "
        SqlStr = SqlStr & "'" & Trim(Left(aFlagInfo(i), 100)) & "', "
        SqlStr = SqlStr & "'" & Trim(aDispCd(i)) & "', "
        SqlStr = SqlStr & "'" & Trim(aUseYN(i)) & "', "
        SqlStr = SqlStr & "'" & Trim(aRmk(i)) & "') "
        
        ADOCN.Execute SqlStr
    Next i
    
    ADOCN.Close: Set ADOCN = Nothing
    
    Add_IFFlagInfo = "SUCCESS"
    
ErrHandler:
    If Err <> 0 Then
        If ADOCN.State = 1 Then
            ADOCN.Close: Set ADOCN = Nothing
        End If
        Add_IFFlagInfo = "0"
        MsgBox "Add_IFFlagInfo Err - " & Err.Description, vbCritical
    End If
End Function
Public Function Get_IFFlagInfo(ByVal sMachineCd As String) As String
    On Error GoTo ErrHandler
    
    Dim ADORS   As New ADODB.Recordset
    Dim SqlStr  As String
    Dim aTmp()  As String
    
    SqlStr = " SELECT flagseq, flagcd, flaginfo, dispcd, useyn, remark " _
            & "  from ifflaginfo " _
            & " order by flagseq "
    
    ADORS.Open SqlStr, fGetCurDSN(sMachineCd), adOpenForwardOnly
    
    If ADORS.EOF = False Then
        Get_IFFlagInfo = ADORS.GetString(adClipString, -1, Chr(124), Chr(3))
        aTmp() = Split(Get_IFFlagInfo, Chr(3))
        
        Get_IFFlagInfo = UBound(aTmp()) & Chr(3) & Get_IFFlagInfo
    Else
        Get_IFFlagInfo = "NONE"
    End If
    
    ADORS.Close: Set ADORS = Nothing
    
ErrHandler:
    If Err <> 0 Then
        If ADORS.State = 1 Then
            ADORS.Close: Set ADORS = Nothing
        End If
        MsgBox "ERR(DCIFLD) : Get_IFFlagInfo(" & Err.Description & ", " & Err.Number & ")"
    End If
End Function
Public Function Get_CalTestItem(ByVal sMachineCd As String, ByVal iMode As Integer) As String
    On Error GoTo ErrHandler
    
    Dim ADORS As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
    
    sSql = ""
    
    If iMode = 0 Then
        sSql = sSql & " SELECT IFTESTSEQ, IFTESTNM, IFSVRCD, CALFORMULA, DOTDIGIT, LHU, "
        sSql = sSql & "        JUDGEGBN, MREF1, MREF2, FREF1, FREF2, PANICL, PANICH, DELTAGBN, DELTAL, DELTAH "
        sSql = sSql & "   FROM CALTESTITEM "
        sSql = sSql & " ORDER BY IFTESTSEQ "
    Else
    End If
    
    ADORS.CursorType = adOpenStatic
    ADORS.Open sSql, fGetCurDSN(sMachineCd)
    
    If ADORS.EOF = False Then
        Get_CalTestItem = ADORS.RecordCount & Chr(3) & ADORS.GetString(adClipString, -1, Chr(124), Chr(3))
    Else
        Get_CalTestItem = "NONE"
    End If
    
    ADORS.Close
    
    Exit Function
ErrHandler:
    If ADORS.State = 1 Then
        ADORS.Close
    End If
    MsgBox "ERR(DCIFLD) : Get_CalTestItem(" & Err.Description & ", " & Err.Number & ")"
End Function

Public Function Get_PreviousResult(ByVal sMachineCd$, ByVal sRegNo$, ByVal sWDate$, ByVal sWSeq$, ByVal sTotIFSeq$) As String
    On Error GoTo ErrHandler
    
    Dim ADORS As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
    
    sSql = ""
    sSql = sSql & " SELECT A.IFSEQ, RESULT1, WDATE+'-'+WSEQ AS WNO "
    sSql = sSql & "   FROM IFRESULT AS A, ( "
    sSql = sSql & "        SELECT IFSEQ, LEFT(MAX(WDATE+WSEQ), 8) AS WDT, RIGHT(MAX(WDATE+WSEQ), 4) AS WSQ "
    sSql = sSql & "          FROM IFRESULT "
    sSql = sSql & "         WHERE REGNO = '" & sRegNo & "' "
    sSql = sSql & "           AND IFSEQ IN (" & sTotIFSeq & ") "
    sSql = sSql & "           AND ((WDATE < '" & sWDate & "') OR (WDATE = '" & sWDate & "' AND WSEQ < '" & sWSeq & "')) "
    sSql = sSql & "         GROUP BY IFSEQ ) AS B "
    sSql = sSql & "  WHERE A.WDATE = B.WDT "
    sSql = sSql & "    AND A.WSEQ = B.WSQ "
    sSql = sSql & "    AND A.IFSEQ = B.IFSEQ "
    sSql = sSql & "    AND A.REGNO = '" & sRegNo & "' "
      
    ADORS.Open sSql, fGetCurDSN(sMachineCd), adOpenForwardOnly
    
    If ADORS.EOF = False Then
        Get_PreviousResult = ADORS.GetString(adClipString, -1, Chr(124), Chr(3))
    Else
        Get_PreviousResult = "NONE"
    End If
    
    ADORS.Close
    
    Exit Function
ErrHandler:
    If ADORS.State <> 0 Then
        ADORS.Close
    End If
    
    Get_PreviousResult = "NONE"
    
    MsgBox "Get_PreviousResult 오류 (AIFLD01.Dll) - (" & Err.Description & ")"
End Function

Public Function Get_LocalDsn(ByVal sMachineCd As String) As String
    Get_LocalDsn = "DSN=" & fGetCurDSN(sMachineCd) & ";UID=;PWD=;"
End Function

Public Function Get_WorkSheetList(ByVal sMachineCd$, ByVal sDate$) As String
    On Error GoTo ErrHandler
    
    Dim ADORS As New ADODB.Recordset
    Dim sSql As String
    
    sSql = " select A.WSEQ, A.IFSEQ,    A.JDATE,    A.JGBN, A.JNO, " _
         & "        A.RACK, A.POS,      A.REGNO,    A.NAME, A.SEX, " _
         & "        A.EMER, B.IFTESTNM, A.RESULT1 " _
         & "   from IFRESULT A, CALTESTITEM B " _
         & "  where A.IFSEQ = B.IFTESTSEQ  " _
         & "    and A.WDATE = '" & sDate$ & "'" _
         & " union " _
         & " select A.WSEQ, A.IFSEQ,    A.JDATE,    A.JGBN, A.JNO, " _
         & "        A.RACK, A.POS,      A.REGNO,    A.NAME, A.SEX, " _
         & "        A.EMER, B.IFTESTNM, A.RESULT1 " _
         & "   from IFRESULT A, IFTESTITEM B " _
         & "  where A.IFSEQ = B.IFTESTSEQ " _
         & "    and A.WDATE = '" & sDate$ & "'" _
         & "  order by A.WSEQ "
    ADORS.Open sSql, fGetCurDSN(sMachineCd), adOpenForwardOnly
    
    If ADORS.EOF = False Then
        Get_WorkSheetList = ADORS.GetString(adClipString, -1, Chr(124), Chr(3))
    Else
        Get_WorkSheetList = "NONE"
    End If
    
    ADORS.Close
    
    Exit Function
ErrHandler:
    If ADORS.State <> 0 Then
        ADORS.Close
    End If
    Get_WorkSheetList = "NONE"
    MsgBox "Get_WorkSheetList 오류 (AIFLD01.Dll) - (" & Err.Description & ")"
End Function

Public Function Add_IFTestItem(ByVal sMachineCd As String, ByVal iMode As Integer, ByVal sTSeq As String, _
                            ByVal sTNm As String, ByVal sTOcd As String, ByVal sTRcd As String, _
                            ByVal sTSPCcd As String, ByVal sTSVRcd As String, ByVal sTDot As String, ByVal sTLHU As String, _
                            ByVal sTJudge$, ByVal sTMRef1$, ByVal sTMRef2$, ByVal sTFRef1$, ByVal sTFRef2$, _
                            ByVal sTPanL$, ByVal sTPanH$, ByVal sTDelGbn$, ByVal sTDelL$, ByVal sTDelH$, _
                            ByVal sTLimit1Gbn$, ByVal sTLimit1$, ByVal sTLimit2Gbn$, ByVal sTLimit2$) As String
    
    On Error GoTo ErrHandler
    
    Dim ADOCN As ADODB.Connection
    Dim sSql$
    Dim sSeq$
    
    Set ADOCN = New ADODB.Connection
    
    ADOCN.Open fGetCurDSN(sMachineCd)
    
    sSql = ""
    
    If iMode = 0 Then
        sSql = sSql & " INSERT INTO IFTESTITEM ( "
        sSql = sSql & " IFTESTSEQ, "
        sSql = sSql & " IFTESTNM, "
        sSql = sSql & " IFORDCD, "
        sSql = sSql & " IFRSTCD, "
        sSql = sSql & " IFSPCCD, "
        sSql = sSql & " IFSVRCD, "
        sSql = sSql & " DOTDIGIT, "
        sSql = sSql & " LHU, "
        sSql = sSql & " JUDGEGBN, "
        sSql = sSql & " MREF1, "
        sSql = sSql & " MREF2, "
        sSql = sSql & " FREF1, "
        sSql = sSql & " FREF2, "
        sSql = sSql & " PANICL, "
        sSql = sSql & " PANICH, "
        sSql = sSql & " DELTAGBN, "
        sSql = sSql & " DELTAL, "
        sSql = sSql & " DELTAH, "
        sSql = sSql & " LIMIT1GBN, "
        sSql = sSql & " LIMIT1, "
        sSql = sSql & " LIMIT2GBN, "
        sSql = sSql & " LIMIT2 ) VALUES ( "
        sSql = sSql & "'" & sTSeq & "', "
        sSql = sSql & "'" & sTNm & "', "
        sSql = sSql & "'" & sTOcd & "', "
        sSql = sSql & "'" & sTRcd & "', "
        sSql = sSql & "'" & sTSPCcd & "', "
        sSql = sSql & "'" & sTSVRcd & "', "
        sSql = sSql & "'" & sTDot & "', "
        sSql = sSql & "'" & sTLHU & "', "
        sSql = sSql & "'" & sTJudge & "', "
        sSql = sSql & "'" & sTMRef1 & "', "
        sSql = sSql & "'" & sTMRef2 & "', "
        sSql = sSql & "'" & sTFRef1 & "', "
        sSql = sSql & "'" & sTFRef2 & "', "
        sSql = sSql & "'" & sTPanL & "', "
        sSql = sSql & "'" & sTPanH & "', "
        sSql = sSql & "'" & sTDelGbn & "', "
        sSql = sSql & "'" & sTDelL & "', "
        sSql = sSql & "'" & sTDelH & "', "
        sSql = sSql & "'" & sTLimit1Gbn & "', "
        sSql = sSql & "'" & sTLimit1 & "', "
        sSql = sSql & "'" & sTLimit2Gbn & "', "
        sSql = sSql & "'" & sTLimit2 & "' ) "
    Else
    End If
    
    ADOCN.Execute sSql
    
    Add_IFTestItem = "SUCCESS"
    
    ADOCN.Close
    Set ADOCN = Nothing
    
    Exit Function
ErrHandler:
    Add_IFTestItem = "0"
    MsgBox "ERR(DCIFLD) : Add_IFTestItem(" & Err.Description & ", " & Err.Number & ")"
End Function

Public Function Add_CalTestItem(ByVal sMachineCd As String, ByVal iMode As Integer, ByVal sTSeq As String, _
                ByVal sTNm As String, ByVal sTSVRcd As String, ByVal sTCalFormula As String, ByVal sTDot As String, ByVal sTLHU As String, _
                ByVal sTJudge$, ByVal sTMRef1$, ByVal sTMRef2$, ByVal sTFRef1$, ByVal sTFRef2$, _
                ByVal sTPanL$, ByVal sTPanH$, ByVal sTDelGbn$, ByVal sTDelL$, ByVal sTDelH$) As String
    
    On Error GoTo ErrHandler
    
    Dim ADOCN As ADODB.Connection
    Dim sSql$
    Dim sSeq$
    
    Set ADOCN = New ADODB.Connection
    
    ADOCN.Open fGetCurDSN(sMachineCd)
    
    sSql = ""
    
    If iMode = 0 Then
        sSql = sSql & " INSERT INTO CALTESTITEM ( "
        sSql = sSql & " IFTESTSEQ, "
        sSql = sSql & " IFTESTNM, "
        sSql = sSql & " IFSVRCD, "
        sSql = sSql & " CALFORMULA, "
        sSql = sSql & " DOTDIGIT, "
        sSql = sSql & " LHU, "
        sSql = sSql & " JUDGEGBN, "
        sSql = sSql & " MREF1, "
        sSql = sSql & " MREF2, "
        sSql = sSql & " FREF1, "
        sSql = sSql & " FREF2, "
        sSql = sSql & " PANICL, "
        sSql = sSql & " PANICH, "
        sSql = sSql & " DELTAGBN, "
        sSql = sSql & " DELTAL, "
        sSql = sSql & " DELTAH ) VALUES ( "
        sSql = sSql & "'" & sTSeq & "', "
        sSql = sSql & "'" & sTNm & "', "
        sSql = sSql & "'" & sTSVRcd & "', "
        sSql = sSql & "'" & sTCalFormula & "', "
        sSql = sSql & "'" & sTDot & "', "
        sSql = sSql & "'" & sTLHU & "', "
        sSql = sSql & "'" & sTJudge & "', "
        sSql = sSql & "'" & sTMRef1 & "', "
        sSql = sSql & "'" & sTMRef2 & "', "
        sSql = sSql & "'" & sTFRef1 & "', "
        sSql = sSql & "'" & sTFRef2 & "', "
        sSql = sSql & "'" & sTPanL & "', "
        sSql = sSql & "'" & sTPanH & "', "
        sSql = sSql & "'" & sTDelGbn & "', "
        sSql = sSql & "'" & sTDelL & "', "
        sSql = sSql & "'" & sTDelH & "' ) "
    Else
    End If
    
    ADOCN.Execute sSql
    
    Add_CalTestItem = "SUCCESS"
    
    ADOCN.Close
    Set ADOCN = Nothing
    
    Exit Function
ErrHandler:
    Add_CalTestItem = "0"
    MsgBox "ERR(DCIFLD) : Add_CalTestItem(" & Err.Description & ", " & Err.Number & ")"
End Function

Public Function Trans_SaveTestItem(ByVal sMachineCd$, ByVal iMode%, ByVal sTSeq$, _
                                ByVal sTNm$, ByVal sTOcd$, ByVal sTRcd$, ByVal sTSPCcd$, ByVal sTSVRcd$, _
                                ByVal sTDot$, ByVal sTLHU$, ByVal sTJudge$, ByVal sTMRef1$, ByVal sTMRef2$, ByVal sTFRef1$, ByVal sTFRef2$, _
                                ByVal sTPanL$, ByVal sTPanH$, sTDelGbn$, sTDelL$, sTDelH$, _
                                ByVal sTCalFormula$, ByVal sTLimit1Gbn$, ByVal sTLimit1$, ByVal sTLimit2Gbn$, ByVal sTLimit2$) As String
    On Error GoTo ErrHandler
        
    Dim sRtnVal$, sOneSeq$, sOneNm$, sOneOcd$, sOneRcd$, sOneSPCcd$, sOneSVRcd$
    Dim sOneDot$, sOneLHU$, sOneJudge$, sOneMRef1$, sOneMRef2$, sOneFRef1$, sOneFRef2$
    Dim sOnePanL$, sOnePanH$, sOneDelGbn$, sOneDelL$, sOneDelH$, sOneCalFormula$
    Dim sOneLimit1Gbn$, sOneLimit1$, sOneLimit2Gbn$, sOneLimit2$
    
    Set adoTransCn = New ADODB.Connection
    
    adoTransCn.CursorLocation = adUseServer
    adoTransCn.Open fGetCurDSN(sMachineCd)
    
    adoTransCn.BeginTrans
    
    If iMode = 0 Then
        sRtnVal = Del_IFTestItem(sMachineCd, 0)
        
        If IsNumeric(sRtnVal) = True Then
            adoTransCn.RollbackTrans
            Trans_SaveTestItem = "0"
            Exit Function
        End If
        
        sRtnVal = Del_CalTestItem(sMachineCd, 0)
        
        If IsNumeric(sRtnVal) = True Then
            adoTransCn.RollbackTrans
            Trans_SaveTestItem = "0"
            Exit Function
        End If
        
        Do
            sOneSeq = GetByOne(sTSeq, sTSeq)
            sOneNm = GetByOne(sTNm, sTNm)
            sOneOcd = GetByOne(sTOcd, sTOcd)
            sOneRcd = GetByOne(sTRcd, sTRcd)
            sOneSPCcd = GetByOne(sTSPCcd, sTSPCcd)
            sOneSVRcd = GetByOne(sTSVRcd, sTSVRcd)
            sOneDot = GetByOne(sTDot, sTDot)
            sOneLHU = GetByOne(sTLHU, sTLHU)
            sOneJudge = GetByOne(sTJudge, sTJudge)
            sOneMRef1 = GetByOne(sTMRef1, sTMRef1)
            sOneMRef2 = GetByOne(sTMRef2, sTMRef2)
            sOneFRef1 = GetByOne(sTFRef1, sTFRef1)
            sOneFRef2 = GetByOne(sTFRef2, sTFRef2)
            sOnePanL = GetByOne(sTPanL, sTPanL)
            sOnePanH = GetByOne(sTPanH, sTPanH)
            sOneDelGbn = GetByOne(sTDelGbn, sTDelGbn)
            sOneDelL = GetByOne(sTDelL, sTDelL)
            sOneDelH = GetByOne(sTDelH, sTDelH)
            sOneCalFormula = GetByOne(sTCalFormula, sTCalFormula)
            sOneLimit1Gbn = GetByOne(sTLimit1Gbn, sTLimit1Gbn)
            sOneLimit1 = GetByOne(sTLimit1, sTLimit1)
            sOneLimit2Gbn = GetByOne(sTLimit2Gbn, sTLimit2Gbn)
            sOneLimit2 = GetByOne(sTLimit2, sTLimit2)
            
            If sOneSeq = "" Then
                Exit Do
            End If
            
            If sOneCalFormula = "" And Len(sOneSeq) = 3 Then
                sRtnVal = Add_IFTestItem(sMachineCd, 0, sOneSeq, sOneNm, sOneOcd, sOneRcd, sOneSPCcd, _
                              sOneSVRcd, sOneDot, sOneLHU, sOneJudge, sOneMRef1, sOneMRef2, sOneFRef1, sOneFRef2, _
                              sOnePanL, sOnePanH, sOneDelGbn, sOneDelL, sOneDelH, sOneLimit1Gbn, sOneLimit1, sOneLimit2Gbn, sOneLimit2)
            ElseIf Len(sOneSeq) = 2 Then
                sRtnVal = Add_CalTestItem(sMachineCd, 0, sOneSeq, sOneNm, sOneSVRcd, sOneCalFormula, sOneDot, sOneLHU, _
                              sOneJudge, sOneMRef1, sOneMRef2, sOneFRef1, sOneFRef2, _
                              sOnePanL, sOnePanH, sOneDelGbn, sOneDelL, sOneDelH)
            End If
            
            If IsNumeric(sRtnVal) = True Then
                adoTransCn.RollbackTrans
                Trans_SaveTestItem = "0"
                Exit Function
            End If
        Loop
    Else
    End If
    
    adoTransCn.CommitTrans
    
    Trans_SaveTestItem = "SUCCESS"
    
    adoTransCn.Close
    Set adoTransCn = Nothing
    
    Exit Function
ErrHandler:
    adoTransCn.RollbackTrans
End Function

Public Function Add_IFResult(ByVal sMachineCd As String, ByVal iMode As Integer, _
                            ByVal sWDate As String, ByVal sWSeq As String, ByVal sIFSeq As String, _
                            ByVal sJDate As String, ByVal sJGbn As String, ByVal sJNo As String, _
                            ByVal sRack As String, ByVal sPos As String, ByVal sRegNo As String, _
                            ByVal sName As String, ByVal sSex As String, ByVal sEmer As String, _
                            ByVal sReRun As String, ByVal sOther As String, ByVal sRst1 As String, _
                            ByVal sRst2 As String, ByVal sFlag As String, ByVal sRegState As String, _
                            Optional ByVal iOrdCnt As Integer) As String
    On Error GoTo ErrHandler
    
    Dim ADOCN As ADODB.Connection
    Dim sSql$, sRtnVal$
    Dim sCIFSeq$, sCWSeq$, sCRst1$, sCRst2$, sOneRowIFSeq$, sOneRowRst1$, sOneRowRst2$
    Dim sCJDate$, sCJGbn$, sCJNo$, sCRack$, sCPos$, sCRegNo$, sCName$, sCSex$, sCEmer$, sCReRun$, sCOther$
    Dim sCFlag$, sOneRowFlag$
    Dim i%
    
    sCWSeq = " "
        
    Set ADOCN = New ADODB.Connection
    
    ADOCN.Open fGetCurDSN(sMachineCd)
    
    If iMode = 0 Then
        sSql = "INSERT INTO IFRESULT VALUES " _
                & "(WDATE, WSEQ, IFSEQ, JDATE, JGBN, JNO, RACK, POS, REGNO, NAME, SEX, EMER, RERUN, OTHER, RESULT1, RESULT2, REGSTATE, FLAG) " _
                & " values ( " & _
            "'" & sWDate & "', " & _
            "'" & sWSeq & "', " & _
            "'" & sIFSeq & "', " & _
            "'" & sJDate & "', " & _
            "'" & sJGbn & "', " & _
            "'" & sJNo & "', " & _
            "'" & sRack & "', " & _
            "'" & sPos & "', " & _
            "'" & sRegNo & "', " & _
            "'" & sName & "', " & _
            "'" & sSex & "', " & _
            "'" & sEmer & "', " & _
            "'" & sReRun & "', " & _
            "'" & sOther & "', " & _
            "'" & sRst1 & "', " & _
            "'" & sRst2 & "', " & _
            "'" & sRegState & "', " & _
            "'" & sFlag & "') "
               
            If sWDate = "" Or sWSeq = "" Or sIFSeq = "" Then
            Else
                ADOCN.Execute sSql
            End If
        
    ElseIf iMode = 1 Then
        For i = 1 To iOrdCnt
            sCIFSeq = GetByOne(sIFSeq, sIFSeq)
            sCRst1 = GetByOne(sRst1, sRst1)
            sCRst2 = GetByOne(sRst2, sRst2)
            sCFlag = GetByOne(sFlag, sFlag)
            
            sSql = "INSERT INTO IFRESULT " _
                    & " (WDATE, WSEQ, IFSEQ, JDATE, JGBN, JNO, RACK, POS, REGNO, NAME, SEX, EMER, RERUN, OTHER, RESULT1, RESULT2, REGSTATE, FLAG) " _
                    & " VALUES ( " & _
                "'" & sWDate & "', " & _
                "'" & sWSeq & "', " & _
                "'" & sCIFSeq & "', " & _
                "'" & sJDate & "', " & _
                "'" & sJGbn & "', " & _
                "'" & sJNo & "', " & _
                "'" & sRack & "', " & _
                "'" & sPos & "', " & _
                "'" & sRegNo & "', " & _
                "'" & sName & "', " & _
                "'" & sSex & "', " & _
                "'" & sEmer & "', " & _
                "'" & sReRun & "', " & _
                "'" & sCOther & "', " & _
                "'" & sCRst1 & "', " & _
                "'" & sCRst2 & "', " & _
                "'" & sRegState & "', " & _
                "'" & sCFlag & "') "
               
            If sWDate = "" Or sWSeq = "" Or sCIFSeq = "" Then
            Else
                ADOCN.Execute sSql
            End If
        Next
    
    ElseIf iMode = 2 Then
        Do While Not sCWSeq = ""
            sCWSeq = GetByOne(sWSeq, sWSeq)
            sCJDate = GetByOne(sJDate, sJDate)
            sCJGbn = GetByOne(sJGbn, sJGbn)
            sCJNo = GetByOne(sJNo, sJNo)
            sCRegNo = GetByOne(sRegNo, sRegNo)
            sCName = GetByOne(sName, sName)
            sCSex = GetByOne(sSex, sSex)
            sCEmer = GetByOne(sEmer, sEmer)
            sCReRun = GetByOne(sReRun, sReRun)
            sCOther = GetByOne(sOther, sOther)
            sCRack = GetByOne(sRack, sRack)
            sCPos = GetByOne(sPos, sPos)
            
            If sCWSeq = "" Then
                Exit Do
            End If
                    
            sCIFSeq = " "
            
            sOneRowIFSeq = GetByOneUserSymbol(sIFSeq, sIFSeq, Chr(3))
            sOneRowRst1 = GetByOneUserSymbol(sRst1, sRst1, Chr(3))
            sOneRowRst2 = GetByOneUserSymbol(sRst2, sRst2, Chr(3))
            sOneRowFlag = GetByOneUserSymbol(sFlag, sFlag, Chr(3))
            
            Do While Not sCIFSeq = ""
                sCIFSeq = GetByOne(sOneRowIFSeq, sOneRowIFSeq)
                sCRst1 = GetByOne(sOneRowRst1, sOneRowRst1)
                sCRst2 = GetByOne(sOneRowRst2, sOneRowRst2)
                sCFlag = GetByOne(sOneRowFlag, sOneRowFlag)
                
                If sCIFSeq = "" Then
                    Exit Do
                End If
                
                sSql = "INSERT INTO IFRESULT " _
                        & " (WDATE, WSEQ, IFSEQ, JDATE, JGBN, JNO, RACK, POS, REGNO, NAME, SEX, EMER, RERUN, OTHER, RESULT1, RESULT2, REGSTATE, FLAG) " _
                        & " VALUES ( " & _
                    "'" & sWDate & "', " & _
                    "'" & sCWSeq & "', " & _
                    "'" & sCIFSeq & "', " & _
                    "'" & sCJDate & "', " & _
                    "'" & sCJGbn & "', " & _
                    "'" & sCJNo & "', " & _
                    "'" & sCRack & "', " & _
                    "'" & sCPos & "', " & _
                    "'" & sCRegNo & "', " & _
                    "'" & sCName & "', " & _
                    "'" & sCSex & "', " & _
                    "'" & sCEmer & "', " & _
                    "'" & sCReRun & "', " & _
                    "'" & sCOther & "', " & _
                    "'" & sCRst1 & "', " & _
                    "'" & sCRst2 & "', " & _
                    "'" & sRegState & "', " & _
                    "'" & sFlag & "') "
                       
                ADOCN.Execute sSql
            Loop
        Loop
    End If
    
    Add_IFResult = "SUCCESS"
    
    ADOCN.Close
    Set ADOCN = Nothing
    
    Exit Function
ErrHandler:
    If Err.Number = -2147217900 Then
        If iMode = 0 Then
            sRtnVal = Update_IFResult(sMachineCd, 0, sWDate, sWSeq, sIFSeq, sJDate, sJGbn, sJNo, _
                                      sRack, sPos, sRegNo, sName, _
                                      sSex, sEmer, sReRun, sOther, _
                                      sRst1, sRst2, sFlag, sRegState, 1)
        ElseIf iMode = 1 Then
            sRtnVal = Update_IFResult(sMachineCd, 0, sWDate, sWSeq, sCIFSeq, sJDate, sJGbn, sJNo, _
                                      sRack, sPos, sRegNo, sName, _
                                      sSex, sEmer, sReRun, sOther, _
                                      sCRst1, sCRst2, sCFlag, sRegState, 1)
        ElseIf iMode = 2 Then
            sRtnVal = Update_IFResult(sMachineCd, 0, sWDate, sCWSeq, sCIFSeq, sCJDate, sCJGbn, sCJNo, _
                                      sCRack, sCPos, sCRegNo, sCName, _
                                      sCSex, sCEmer, sCReRun, sCOther, _
                                      sCRst1, sCRst2, sCFlag, sRegState, 1)
        Else
        
        End If
        
        If CBool(sRtnVal) = True Then
            Resume Next
        End If
    End If
    
    Add_IFResult = "0"
    
    MsgBox "ERR(DCIFLD) : Add_IFResult(" & Err.Description & ", " & Err.Number & ")"
End Function

Public Function Get_IFResult(ByVal sMachineCd As String, ByVal sWDate As String, ByVal sJDate As String, _
                            ByVal iGbn As Integer, Optional ByVal sCondition As String, Optional vParam As Variant) As String
    On Error GoTo ErrHandler
    
    Dim ADORS As New ADODB.Recordset
    Dim sSql As String
    
'    Call cmbGbn.AddItem("ALL", 0)
'    Call cmbGbn.AddItem("등록", 1)
'    Call cmbGbn.AddItem("미등록", 2)
'    Call cmbGbn.AddItem("JGbn(A)", 3)
'    Call cmbGbn.AddItem("JGbn(R)", 4)
'    Call cmbGbn.AddItem("JGbn(N)", 5)
'    Call cmbGbn.AddItem("JNo(A)", 6)
'    Call cmbGbn.AddItem("JNo(L)", 7)
        
    Call GetOrdRstCfg(sMachineCd)
        
    sSql = "SELECT WSEQ, JDATE, JGBN, JNO, RACK, POS, REGNO, NAME, SEX," _
        & " EMER, RERUN, OTHER, IFSEQ, RESULT1, RESULT2, FLAG " _
        & " FROM IFRESULT" _
        & " WHERE WDATE = '" & sWDate & "'"

    Select Case iGbn
        Case 0
        
        Case 1
            sSql = sSql & " AND REGSTATE = '1'"
        
        Case 2
            sSql = sSql & " AND REGSTATE = '0'"
        
        Case 3
            sSql = sSql & " AND JNO = '" & sCondition & "' "
                
        Case 4
            sSql = ""
                    
            sSql = "  SELECT WSEQ, JDATE, JGBN, JNO, RACK, POS, REGNO, NAME, SEX, EMER, RERUN, OTHER, IFSEQ, RESULT1, RESULT2, FLAG " _
                   & "  FROM IFRESULT " _
                   & " WHERE WDATE = '" & sWDate & "' " _
                   & "   AND JNO = '" & sCondition & "' " _
                   & "   AND WSEQ = ("
            sSql = sSql & "  SELECT TOP 1 WSEQ " _
                        & " FROM IFRESULT" _
                        & " WHERE WDATE = '" & sWDate & "'" _
                        & "   AND JNO = '" & sCondition & "' " _
                        & " ORDER BY WSEQ DESC, IFSEQ ASC " _
                        & " ) "
                  
        Case Else
        
    End Select
    
    If iGbn = 4 Then
    Else
        sSql = sSql & " ORDER BY WSEQ, IFSEQ "
    End If
    
    ADORS.CursorType = adOpenStatic
    ADORS.Open sSql, fGetCurDSN(sMachineCd)
    
    If ADORS.EOF = False Then
        vParam = ADORS.GetRows(-1)
        Get_IFResult = "OK"
        'Get_IFResult = adoRs.GetString(adClipString, -1, Chr(124), Chr(3))
    Else
        Get_IFResult = "NONE"
    End If

    ADORS.Close

    Exit Function
ErrHandler:
    ADORS.Close
    MsgBox "ERR(DCIFLD) : Get_IFResult(" & Err.Description & ", " & Err.Number & ")"
End Function

Public Function Get_LastIFResult(ByVal sMachineCd As String, ByVal sWDate As String) As String
    On Error GoTo ErrHandler
    
    Dim ADORS As New ADODB.Recordset
    Dim sSql As String

    sSql = "SELECT TOP 1 WSEQ " _
        & " FROM IFRESULT" _
        & " WHERE WDATE = '" & sWDate & "'" _
        & " ORDER BY WSEQ DESC"

    ADORS.CursorType = adOpenStatic
    ADORS.Open sSql, fGetCurDSN(sMachineCd)
    
    If ADORS.EOF = False Then
        Get_LastIFResult = ADORS.GetString(adClipString, -1, "|", Chr$(3))
    Else
        Get_LastIFResult = "NONE"
    End If

    ADORS.Close

    Exit Function
ErrHandler:
    If ADORS.State = 1 Then
        ADORS.Close
    End If
    MsgBox "ERR(DCIFLD) : Get_LastIFResult(" & Err.Description & ", " & Err.Number & ")"
End Function

Public Function Update_IFResult(ByVal sMachineCd As String, ByVal iMode As Integer, _
                              ByVal sWDate As String, ByVal sWSeq As String, ByVal sIFSeq As String, _
                              ByVal sJDate As String, ByVal sJGbn As String, ByVal sJNo As String, _
                              ByVal sRack As String, ByVal sPos As String, ByVal sRegNo As String, _
                              ByVal sName As String, ByVal sSex As String, ByVal sEmer As String, _
                              ByVal sReRun As String, ByVal sOther As String, ByVal sResult1 As String, _
                              ByVal sResult2 As String, ByVal sFlag As String, ByVal sRegState As String, _
                              Optional ByVal iItemCnt As Integer, Optional iExist% = 0) As Boolean
    On Error GoTo ErrHandler

    Dim ADOCN As New ADODB.Connection
    Dim ADORS As New ADODB.Recordset
    Dim sSql$, sSql2$
    Dim sSeq$
    Dim iCnt%
    Dim sTmpIFSeq$
    
    ADOCN.Open "" & fGetCurDSN(sMachineCd) & "", "", ""
    
    Select Case iMode
        '한 항목 등록시
        Case 0
            iExist = 0
            
            sSql = ""
            sSql = sSql & " SELECT COUNT(A.UNIQUE_JNO) "
            sSql = sSql & "   FROM ( SELECT DISTINCT JNO AS UNIQUE_JNO FROM IFRESULT "
            sSql = sSql & "           WHERE WDATE = '" & sWDate & "' "
            sSql = sSql & "             AND WSEQ = '" & sWSeq & "' ) A "
                        
            ADORS.Open sSql, ADOCN, adOpenForwardOnly
            
            If ADORS.EOF = False Then
                If Trim("" & ADORS(0)) >= "2" Then
                    iExist = 1
                Else
                    iExist = 0
                End If
            Else
                iExist = 0
            End If
            
            ADORS.Close
            
            sSql = ""
            sSql = sSql & " UPDATE IFRESULT SET "
                If iExist = 0 Then
            sSql = sSql & "        JNO = '" & sJNo & "', "
                End If
            sSql = sSql & "        JDATE = '" & sJDate & "', "
            sSql = sSql & "        JGBN = '" & sJGbn & "', "
            sSql = sSql & "        RACK = '" & sRack & "', "
            sSql = sSql & "        POS = '" & sPos & "', "
            sSql = sSql & "        REGNO = '" & sRegNo & "', "
            sSql = sSql & "        NAME = '" & sName & "', "
            sSql = sSql & "        SEX = '" & sSex & "', "
            sSql = sSql & "        EMER = '" & sEmer & "', "
            sSql = sSql & "        RERUN = '" & sReRun & "', "
            sSql = sSql & "        OTHER = '" & sOther & "', "
            sSql = sSql & "        RESULT1 = '" & sResult1 & "', "
            sSql = sSql & "        RESULT2 = '" & sResult2 & "', "
            sSql = sSql & "        FLAG = '" & sFlag & "', "
            sSql = sSql & "        REGSTATE = '" & sRegState & "' "
            sSql = sSql & "  WHERE WDATE = '" & sWDate & "' "
            sSql = sSql & "    AND WSEQ = '" & sWSeq & "' "
            sSql = sSql & "    AND IFSEQ = '" & sIFSeq & "' "

            ADOCN.Execute sSql
            
        '여러항목 등록시
        Case 1
            iExist = 0
            
            sSql = ""
            sSql = sSql & " SELECT COUNT(A.UNIQUE_JNO) "
            sSql = sSql & "   FROM ( SELECT DISTINCT JNO AS UNIQUE_JNO FROM IFRESULT "
            sSql = sSql & "           WHERE WDATE = '" & sWDate & "' "
            sSql = sSql & "             AND WSEQ = '" & sWSeq & "' ) A "
            
            ADORS.Open sSql, ADOCN, adOpenForwardOnly
            
            If ADORS.EOF = False Then
                If Trim("" & ADORS(0)) >= "2" Then
                    iExist = 1
                Else
                    iExist = 0
                End If
            Else
                iExist = 0
            End If
            
            ADORS.Close
                
            For iCnt = 1 To iItemCnt
                sSql = ""
                sSql = sSql & " UPDATE IFRESULT SET "
                    If iExist = 0 Then
                sSql = sSql & "        JNO = '" & sJNo & "', "
                    End If
                sSql = sSql & "        JDATE = '" & sJDate & "', "
                sSql = sSql & "        JGBN = '" & sJGbn & "', "
                sSql = sSql & "        RACK = '" & sRack & "', "
                sSql = sSql & "        POS = '" & sPos & "', "
                sSql = sSql & "        REGNO = '" & sRegNo & "', "
                sSql = sSql & "        NAME = '" & sName & "', "
                sSql = sSql & "        SEX = '" & sSex & "', "
                sSql = sSql & "        EMER = '" & sEmer & "', "
                sSql = sSql & "        RERUN = '" & sReRun & "', "
                sSql = sSql & "        OTHER = '" & sOther & "', "
                sSql = sSql & "        RESULT1 = '" & sResult1 & "', "
                sSql = sSql & "        RESULT2 = '" & sResult2 & "', "
                sSql = sSql & "        FLAG = '" & sFlag & "', "
                sSql = sSql & "        REGSTATE = '" & sRegState & "' "
                sSql = sSql & "  WHERE WDATE = '" & sWDate & "' "
                sSql = sSql & "    AND WSEQ = '" & sWSeq & "' "
                sSql = sSql & "    AND IFSEQ = '" & GetByOne(sIFSeq, sIFSeq) & "' "
    
                ADOCN.Execute sSql
            Next
        
        '서버등록후 등록상태 표기
        Case 2
            For iCnt = 1 To iItemCnt
                sSql = "UPDATE IFRESULT " & _
                       "   SET REGSTATE = '1' " & _
                       " WHERE WDATE = '" & sWDate & "' " & _
                       "   AND WSEQ = '" & GetByOne(sWSeq, sWSeq) & "' "
                
                ADOCN.Execute sSql
                
            Next
            
        Case Else
        
    End Select

    Update_IFResult = True

    ADOCN.Close
    Set ADOCN = Nothing

    Exit Function
ErrHandler:
    Update_IFResult = False
    MsgBox "ERR(DCIFLD) : Update_IFResult(" & Err.Description & ", " & Err.Number & ")"
End Function

Public Sub Update_PatInfo(ByVal sMachineCd As String, _
                        ByVal sWDate As String, ByVal sWSeq As String, _
                        ByVal sJDate As String, ByVal sJGbn As String, ByVal sJNo As String, _
                        ByVal sRegNo As String, _
                        ByVal sName As String, ByVal sSex As String, ByVal sEmer As String, _
                        ByVal sReRun As String, ByVal sOther As String)
    On Error GoTo ErrRtn
    
    Dim SqlStr  As String
    Dim ADOCN   As New ADODB.Connection
    
    ADOCN.Open fGetCurDSN(sMachineCd)
        
    SqlStr = " UPDATE IFRESULT " _
            & "   SET JNO = '" & sJNo & "', " _
            & "       JDATE = '" & sJDate & "', " _
            & "       JGBN = '" & sJGbn & "', " _
            & "       REGNO = '" & sRegNo & "', " _
            & "       NAME = '" & sName & "', " _
            & "       SEX = '" & sSex & "', " _
            & "       EMER = '" & sEmer & "', " _
            & "       RERUN = '" & sReRun & "', " _
            & "       OTHER = '" & sOther & "', " _
            & "       REGSTATE = '0' " _
            & " WHERE WDATE = '" & sWDate & "'" _
            & "   AND WSEQ = '" & sWSeq & "'"
    
    ADOCN.Execute SqlStr
    
    ADOCN.Close: Set ADOCN = Nothing
    
    Exit Sub
ErrRtn:
    If ADOCN.State = 1 Then
        ADOCN.Close: Set ADOCN = Nothing
    End If
    MsgBox "ERR(DCIFLD) : Update_PatInfo(" & Err.Description
End Sub

Public Function Edit_IFResult(ByVal sMachineCd As String, ByVal iMode As Integer, _
                              ByVal sWDate As String, ByVal sWSeq As String, ByVal sIFSeq As String, _
                              ByVal sJDate As String, ByVal sJGbn As String, ByVal sJNo As String, _
                              ByVal sRegNo As String, _
                              ByVal sName As String, ByVal sSex As String, ByVal sEmer As String, _
                              ByVal sReRun As String, ByVal sOther As String, ByVal sResult1 As String, _
                              ByVal sResult2 As String, ByVal sFlag As String, ByVal sRegState As String, _
                              Optional ByVal iItemCnt As Integer, Optional iExist% = 0) As Boolean
    On Error GoTo ErrHandler

    Dim ADOCN As New ADODB.Connection
    Dim ADORS As New ADODB.Recordset
    Dim sSql$, sSql2$
    Dim sSeq$
    Dim iCnt%
    Dim sTmpIFSeq$
    
    ADOCN.Open fGetCurDSN(sMachineCd)
    
    Select Case iMode
        '한 항목 등록시
        Case 0
            iExist = 0
            
            sSql = ""
            sSql = sSql & " SELECT COUNT(A.UNIQUE_JNO) "
            sSql = sSql & "   FROM ( SELECT DISTINCT JNO AS UNIQUE_JNO FROM IFRESULT "
            sSql = sSql & "           WHERE WDATE = '" & sWDate & "' "
            sSql = sSql & "             AND WSEQ = '" & sWSeq & "' ) A "
                        
            ADORS.Open sSql, ADOCN, adOpenForwardOnly
            
            If ADORS.EOF = False Then
                If Trim("" & ADORS(0)) >= "2" Then
                    iExist = 1
                Else
                    iExist = 0
                End If
            Else
                iExist = 0
            End If
            
            ADORS.Close
            
            sSql = ""
            sSql = sSql & " UPDATE IFRESULT SET "
            
            If iExist = 0 Then
                sSql = sSql & "   JDATE = '" & sJDate & "', "
                sSql = sSql & "   JGBN = '" & sJGbn & "', "
                sSql = sSql & "   JNO = '" & sJNo & "', "
                sSql = sSql & "   REGNO = '" & sRegNo & "', "
                sSql = sSql & "   NAME = '" & sName & "', "
                sSql = sSql & "   SEX = '" & sSex & "', "
                sSql = sSql & "   EMER = '" & sEmer & "', "
                sSql = sSql & "   RERUN = '" & sReRun & "', "
                sSql = sSql & "   OTHER = '" & sOther & "', "
            End If
            
            sSql = sSql & "       RESULT1 = '" & sResult1 & "', "
            sSql = sSql & "       RESULT2 = '" & sResult2 & "', "
            sSql = sSql & "       REGSTATE = '" & sRegState & "', "
            sSql = sSql & "       FLAG = '" & sFlag & "'"
            sSql = sSql & "      WHERE WDATE = '" & sWDate & "' "
            sSql = sSql & "        AND WSEQ = '" & sWSeq & "' "
            sSql = sSql & "        AND IFSEQ = '" & sIFSeq & "' "

            ADOCN.Execute sSql
            
        '여러항목 등록시
        Case 1
            iExist = 0
            
            sSql = ""
            sSql = sSql & " SELECT COUNT(A.UNIQUE_JNO) "
            sSql = sSql & "   FROM ( SELECT DISTINCT JNO AS UNIQUE_JNO FROM IFRESULT "
            sSql = sSql & "           WHERE WDATE = '" & sWDate & "' "
            sSql = sSql & "             AND WSEQ = '" & sWSeq & "' ) A "
            
            ADORS.Open sSql, ADOCN, adOpenForwardOnly
            
            If ADORS.EOF = False Then
                If Trim("" & ADORS(0)) >= "2" Then
                    iExist = 1
                Else
                    iExist = 0
                End If
            Else
                iExist = 0
            End If
            
            ADORS.Close
                
            For iCnt = 1 To iItemCnt
                sSql = ""
                sSql = sSql & " UPDATE IFRESULT SET "
                
                If iExist = 0 Then
                    sSql = sSql & "   JDATE = '" & sJDate & "', "
                    sSql = sSql & "   JGBN = '" & sJGbn & "', "
                    sSql = sSql & "   JNO = '" & sJNo & "', "
                    sSql = sSql & "   REGNO = '" & sRegNo & "', "
                    sSql = sSql & "   NAME = '" & sName & "', "
                    sSql = sSql & "   SEX = '" & sSex & "', "
                    sSql = sSql & "   EMER = '" & sEmer & "', "
                    sSql = sSql & "   RERUN = '" & sReRun & "', "
                    sSql = sSql & "   OTHER = '" & sOther & "', "
                End If
                
                sSql = sSql & "       RESULT1 = '" & GetByOne(sResult1, sResult1) & "', "
                sSql = sSql & "       RESULT2 = '" & GetByOne(sResult2, sResult2) & "', "
                sSql = sSql & "       FLAG = '" & GetByOne(sFlag, sFlag) & "', "
                sSql = sSql & "       REGSTATE = '" & sRegState & "' "
                sSql = sSql & "      WHERE WDATE = '" & sWDate & "' "
                sSql = sSql & "        AND WSEQ = '" & sWSeq & "' "
                sSql = sSql & "        AND IFSEQ = '" & GetByOne(sIFSeq, sIFSeq) & "' "
    
                ADOCN.Execute sSql
            Next
        
        '서버등록후 등록상태 표기
        Case 2
            For iCnt = 1 To iItemCnt
                sSql = "UPDATE IFRESULT " & _
                       "   SET REGSTATE = '1' " & _
                       " WHERE WDATE = '" & sWDate & "' " & _
                       "   AND WSEQ = '" & GetByOne(sWSeq, sWSeq) & "' "
                
                ADOCN.Execute sSql
                
            Next
            
        Case Else
        
    End Select

    Edit_IFResult = True

    ADOCN.Close
    Set ADOCN = Nothing

    Exit Function
ErrHandler:
    Edit_IFResult = False
    MsgBox "ERR(DCIFLD) : Edit_IFResult(" & Err.Description & ", " & Err.Number & ")"
End Function

Public Function Del_IFResult(ByVal sMachineCd As String, ByVal iMode As Integer, ByVal sWDate As String, ByVal sWSeq As String, Optional ByVal sIFSeq As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim sSql$
    Dim ADOCN As ADODB.Connection
    Dim lngDeleteOK&
    
    sSql = "DELETE * FROM IFRESULT "
    
    If iMode = 0 Then
        sSql = sSql & "WHERE WDATE = '" & sWDate & "'" _
            & " AND WSEQ = '" & sWSeq & "'"

    ElseIf iMode = 1 Then
        sSql = sSql & "WHERE WDATE = '" & sWDate & "'" _
            & " AND WSEQ = '" & sWSeq & "'" _
            & " AND IFSEQ = '" & sIFSeq & "'"
    ElseIf iMode = 2 Then
        sSql = sSql & "WHERE WDATE <= '" & sWDate & "' "
    End If

    Set ADOCN = New ADODB.Connection

    ADOCN.Open fGetCurDSN(sMachineCd)

    ADOCN.Execute sSql, lngDeleteOK

    If lngDeleteOK = 0 Then
        Del_IFResult = False
    Else
        Del_IFResult = True
    End If

    ADOCN.Close
    Set ADOCN = Nothing

    Exit Function
ErrHandler:
    Del_IFResult = False
    MsgBox "ERR(DCIFLD) : Del_IFResult(" & Err.Description & ", " & Err.Number & ")"
End Function

Public Function FetchOrder(ByVal sMachineCd As String, ByVal sGbn As String, _
            ByVal sCWDate As String, ByVal sCJDate As String, Optional ByVal sConditon As String) As String

    On Error GoTo ErrHandler
    
    Dim ADORS As New ADODB.Recordset
    Dim sSql As String
    
    If sGbn = "OLD" Then
    Else
        mCurItemCnt = 0
        mPersonCnt = 0
        Exit Function
    End If
    
    sSql = "SELECT COUNT(*) " & _
               "FROM IFRESULT " & _
               "WHERE WDATE = '" & sCWDate & "' " & _
               "AND (RESULT1 = '' OR RESULT1 IS NULL) " & _
               "GROUP BY WSEQ "

    ADORS.CursorType = adOpenStatic
    ADORS.Open sSql, "DSN=" & fGetCurDSN(sMachineCd) & ";UID=;PWD=;"
          
    mPersonCnt = ADORS.RecordCount
    
    If mPersonCnt = 0 Then
        ADORS.Close
        Exit Function
    End If
    
    ADORS.Close
    
    sSql = "SELECT WSEQ, '1', JDATE, JGBN, JNO, RACK, POS, REGNO, NAME, SEX, EMER, 'Y', OTHER, 'N', 'N', IFSEQ " & _
               "FROM IFRESULT " & _
               "WHERE WDATE = '" & sCWDate & "' " & _
               "AND (RESULT1 = '' OR RESULT1 IS NULL) " & _
               "ORDER BY WSEQ, IFSEQ "
    
    ADORS.CursorType = adOpenStatic
    ADORS.Open sSql, fGetCurDSN(sMachineCd)
    
    If ADORS.RecordCount = 0 Then
        mCurItemCnt = 0
        Exit Function
    End If
    
    If ADORS.EOF = False Then
        mCurItemCnt = ADORS.RecordCount
        FetchOrder = ADORS.GetString(adClipString, -1, "|", Chr$(3))
        FetchOrder = Replace(FetchOrder, Chr$(3), "|" & Chr$(3))
    Else
        mCurItemCnt = 0
        FetchOrder = "NONE"
    End If

    ADORS.Close

    Exit Function
ErrHandler:
    ADORS.Close
    MsgBox "ERR(DCIFLD) : FetchOrder(" & Err.Description & ", " & Err.Number & ")"
End Function

Public Function Set_WorkSheetList(ByVal sMachineCd$, ByVal aDate As String, _
                                  ByRef aData() As String, _
                                  ByRef aItemNm() As String, _
                                  ByRef aResult() As String) As Long
    On Error GoTo ErrHandler
    
    Dim ADOCN As New ADODB.Connection
    
    Dim iRow As Integer
    Dim iCnt As Integer
    
    Dim sSql As String
    
    Dim sFields As String
    Dim sValues As String
    
    ADOCN.Open fGetCurDSN(sMachineCd)
    
    sSql = " delete * FROM IFRESULT_PRT where WDATE = '" & aDate & "' "
    ADOCN.Execute sSql
    
    sSql = ""
    For iRow = 1 To UBound(aData, 2)
        sFields = " WDATE, WSEQ, IFSEQ, SEQNO, JDATE, JGBN, JNO, RACK, POS,  REGNO, NAME, SEX, EMER "
        
        For iCnt = 1 To UBound(aData, 1)
            Select Case iCnt
                Case 1:    sValues = "'" & aDate & "', '" & aData(iCnt, iRow) & "'"
                Case 3:    sValues = sValues & ", " & aData(iCnt, iRow)
                Case Else: sValues = sValues & ", '" & aData(iCnt, iRow) & "'"
            End Select
        Next
        
        For iCnt = 1 To UBound(aItemNm)
            sFields = sFields & ", ORDNM" & CStr(iCnt) & ", RESULT" & CStr(iCnt)
            sValues = sValues & ", '" & aItemNm(iCnt, iRow) & "'" & ", '" & aResult(iCnt, iRow) & "'"
        Next
        
        sSql = " insert into IFRESULT_PRT (" & sFields & ") values (" & sValues & ")"
        ADOCN.Execute sSql
    Next
    
    ADOCN.Close
    Set ADOCN = Nothing
    
    Exit Function
ErrHandler:
    If ADOCN.State <> 0 Then
        ADOCN.Close
    End If
    
    MsgBox "Set_WorkSheetList 오류 (AIFLD01.Dll) - (" & Err.Description & ")"
End Function

