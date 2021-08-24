VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCJ0101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim AdoConn     As New ADODB.Connection
Dim RegNo       As String
Dim LabNo      As String

Function Add_Person(sPerson As String) As String
' 접수신상 insert
    Dim sSql        As String
    Dim sRegno      As String
    Dim sName       As String
    Dim sBornYear   As String
    Dim sIdLeft     As String
    Dim sSex        As String
    Dim sIdRight    As String
    
    sRegno = GetByOne(sPerson, sPerson)     '1
    sName = GetByOne(sPerson, sPerson)      '2
    sBornYear = GetByOne(sPerson, sPerson)  '3
    sIdLeft = GetByOne(sPerson, sPerson)    '4
    sSex = GetByOne(sPerson, sPerson)       '5
    sIdRight = GetByOne(sPerson, sPerson)   '6
    
    On Error GoTo ErrHandler
    
    sSql = " INSERT INTO PERSON (REGNO, NAME, BORNYEAR, IDLEFT, SEX, IDRIGHT) " & _
           " VALUES ('" & sRegno & "','" & sName & "','" & sBornYear & "','" & sIdLeft & "','" & sSex & "','" & sIdRight & "')"

    AdoConn.Execute sSql
    
    Add_Person = "SUCCESS"
Exit Function
ErrHandler:
    Add_Person = Str(Err.Number)
    If Str(Err.Number) <> "-2147217900" Then
        MsgBox "ADD (DLL) - " & Err.Description
    End If
    
End Function

Function Add_JubSu(sJubSu As String) As String
' 접수 MASTE Insert
    
    Dim RCnt        As Integer
    Dim iCnt        As Integer
    Dim sSql        As String
    Dim sPartCd     As String
    Dim sLabDate    As String
    Dim sPartGbn    As String
    Dim sLabSeq     As String
    Dim sLabTime    As String
    Dim sSpcCd      As String
    Dim sRegno      As String
    Dim sName       As String
    Dim sAge        As String
    Dim sIdLeft     As String
    Dim sSex        As String
    Dim sIdRight    As String
    Dim sDeptCd     As String
    Dim sWard       As String
    Dim sJubSuGbn   As String
    Dim sEmerGbn    As String
    Dim sDoctor     As String
    Dim sMedicalcmt As String
    Dim sSpecialOrd As String
    
    sPartCd = GetByOne(sJubSu, sJubSu)        '1
    sLabDate = GetByOne(sJubSu, sJubSu)       '2
    sPartGbn = GetByOne(sJubSu, sJubSu)       '3
    sSpcCd = GetByOne(sJubSu, sJubSu)         '4
    sRegno = GetByOne(sJubSu, sJubSu)         '5
    RegNo = sRegno
    sLabSeq = Format(Get_LastSeq(sLabDate, sPartCd, sPartGbn), "00000")
    sLabTime = Get_Date("TS")
'    sSpcCd = GetByOne(sJubSu, sJubSu)         '4
'    sRegno = GetByOne(sJubSu, sJubSu)         '5
    sName = GetByOne(sJubSu, sJubSu)          '6
    sAge = GetByOne(sJubSu, sJubSu)           '7
    sIdLeft = GetByOne(sJubSu, sJubSu)        '8
    sSex = GetByOne(sJubSu, sJubSu)           '9
    sIdRight = GetByOne(sJubSu, sJubSu)       '10
    sDeptCd = GetByOne(sJubSu, sJubSu)        '11
    sWard = GetByOne(sJubSu, sJubSu)          '12
    sJubSuGbn = GetByOne(sJubSu, sJubSu)      '13
    sEmerGbn = GetByOne(sJubSu, sJubSu)       '14
    sDoctor = GetByOne(sJubSu, sJubSu)        '15
    sMedicalcmt = GetByOne(sJubSu, sJubSu)    '16
    sSpecialOrd = GetByOne(sJubSu, sJubSu)    '17
    
    On Error GoTo ErrHandler
    
    sSql = " INSERT INTO H_JUBSU (LABDATE, PARTGBN, LABSEQ, LABTIME, SPECIMENCD, " & _
           "                      REGNO, NAME, AGE, IDLEFT, SEX, IDRIGHT, DEPTCD, " & _
           "                      WARD, JUBSUGBN, EMERGBN, DOCTOR, MEDICALREMARK, " & _
           "                      SPECIALORD, SLIPYN) " & _
           " VALUES ('" & sLabDate & "','" & sPartGbn & "','" & sLabSeq & "','" & sLabTime & "','" & sSpcCd & _
           "','" & sRegno & "','" & sName & "','" & sAge & "','" & sIdLeft & "','" & sSex & "','" & sIdRight & "','" & sDeptCd & _
           "','" & sWard & "','" & sJubSuGbn & "','" & sEmerGbn & "','" & sDoctor & "','" & sMedicalcmt & _
           "','" & sSpecialOrd & "','0') "
           
    AdoConn.Execute sSql
    
    Add_JubSu = sLabSeq
Exit Function
ErrHandler:
    Add_JubSu = Str(Err.Number)
    MsgBox "ADD (DLL) - " & Err.Description
    
End Function

Function Add_Result(sResult As String) As String
' 접수 내역 Detail insert
        
    Dim iCnt        As Integer
    Dim iOcnt       As Integer
    Dim sSql        As String
    Dim sPartCd     As String
    Dim sLabDate    As String
    Dim sPartGbn    As String
    Dim sLabSeq     As String
    Dim sSpcCd      As String
    Dim sOrdSeq     As String
    Dim sSubMcd     As String
    Dim sRegno      As String
    Dim sName       As String
    Dim sAge        As String
    Dim sIdLeft     As String
    Dim sSex        As String
    Dim sIdRight    As String
    Dim sRCd        As String
    
    sLabSeq = GetByOne(sResult, sResult)    '1
    sPartCd = GetByOne(sResult, sResult)    '2
    sLabDate = GetByOne(sResult, sResult)   '3
    sPartGbn = GetByOne(sResult, sResult)   '4
    sSpcCd = GetByOne(sResult, sResult)     '5
    
    sRegno = GetByOne(sResult, sResult)     '6
    sName = GetByOne(sResult, sResult)      '7
    sAge = GetByOne(sResult, sResult)       '8
    sIdLeft = GetByOne(sResult, sResult)    '9
    sSex = GetByOne(sResult, sResult)       '10
    sIdRight = GetByOne(sResult, sResult)   '11
    iOcnt = Val(GetByOne(sResult, sResult)) '12
    
    
    On Error GoTo ErrHandler
    
    For iCnt = 1 To iOcnt
    
        sOrdSeq = GetByOne(sResult, sResult)    '13 * 3n
        sSubMcd = GetByOne(sResult, sResult)    '14 * 3n
        sRCd = GetByOne(sResult, sResult)       '15 * 3n
            
        sSql = " INSERT INTO H_RESULT " & _
               " SELECT '" & sLabDate & "' AS LABDATE,'" & sPartGbn & "' AS PARTGBN," & _
               "'" & sLabSeq & "' AS LABSEQ,'" & sSpcCd & "' AS SPECIMENCD," & _
               "TESTITEMSEQ, SUBMCD,'" & sRegno & "' AS REGNO,'" & sName & "' AS NAME," & _
               "'" & sAge & "' AS AGE,'" & sIdLeft & "' AS IDLEFT,'" & sSex & "' AS SEX," & _
               "'" & sIdRight & "' AS IDRIGHT,'' AS RESULTDATE, '' AS RESULTTIME,'" & sRCd & "' AS ROUTINECD, '' AS MACHINECD,'' AS RESULT,"
        If sRegno <> "" Then
            sSql = sSql & " DELTAGBN, "
        Else
            sSql = sSql & "'N' AS DELTAGBN, "
        End If
        sSql = sSql & " REFGBN, PANICGBN, " & _
               "'' AS DELTAMARK, '' AS REFMARK, '' AS PANICMARK, '' AS BEFORELABNO, '' AS BEFORERESULT, 0 AS INTERVAL" & _
               "  FROM TESTITEM " & _
               " WHERE PARTCD  = '" & sPartCd & "'" & _
               "   AND PARTGBN = '" & sPartGbn & "'" & _
               "   AND SPECIMENCD = '" & sSpcCd & "'" & _
               "   AND TESTITEMSEQ = '" & sOrdSeq & "'"
        If Left(Trim(sSubMcd), 2) = "NN" Then   ' Sub 검사항목일 경우 여러 Row를 동시에 INSERT
               sSql = sSql & "   AND SUBMCD = '" & sSubMcd & "'"
        End If
    
        AdoConn.Execute sSql
    
    Next iCnt
    
    Add_Result = "SUCCESS"
Exit Function
ErrHandler:
    If Err.Number = -2147217900 Then
        Resume Next
    End If
    Add_Result = Str(Err.Number)
    MsgBox "ADD (DLL) - " & Err.Description
    
End Function

Function Add_LabNo(LabDate As String, PartCd As String, PartGbn As String) As String

    Dim sSql    As String
    Dim iCnt    As Integer
    
    On Error GoTo ErrHandler
    
    If LabNo = "" Then
        sSql = " INSERT INTO H_LABNO (LABDATE, PARTGBN, LABSEQ, REGNO) " & _
               " VALUES ('" & LabDate & "','" & PartGbn & "','" & "00001','" & RegNo & "') "
    Else
        sSql = " INSERT INTO H_LABNO (LABDATE, PARTGBN, LABSEQ, REGNO) " & _
           " VALUES ('" & LabDate & "','" & PartGbn & "','" & LabNo & "','" & RegNo & "' ) "
    End If
    
    AdoConn.Execute sSql
    
    Add_LabNo = "SUCCESS"
Exit Function
ErrHandler:
    Add_LabNo = "ERROR"
    MsgBox "EDIT (DLL) - " & Err.Description
    
End Function

'U1) 환자 신상 Master정보를 수정한다.
Function Edit_Person(sPerson As String) As String
' 접수신상 update
    
    Dim sSql        As String
    Dim sRegno      As String
    Dim sName       As String
    Dim sBornYear   As String
    Dim sIdLeft     As String
    Dim sSex        As String
    Dim sIdRight    As String
        
    sRegno = GetByOne(sPerson, sPerson)     '1
    sName = GetByOne(sPerson, sPerson)      '2
    sBornYear = GetByOne(sPerson, sPerson)  '3
    sIdLeft = GetByOne(sPerson, sPerson)    '4
    sSex = GetByOne(sPerson, sPerson)       '5
    sIdRight = GetByOne(sPerson, sPerson)   '6
    
    On Error GoTo ErrHandler
    
    sSql = " UPDATE PERSON " & _
           "    SET NAME = '" & sName & "'," & _
           "        BORNYEAR = '" & sBornYear & "'," & _
           "        IDLEFT = '" & sIdLeft & "'," & _
           "        SEX    = '" & sSex & "'," & _
           "        IDRIGHT = '" & sIdRight & "'" & _
           "  WHERE REGNO = '" & sRegno & "'"

    AdoConn.Execute sSql
    
    Edit_Person = "SUCCESS"
Exit Function
ErrHandler:
    Edit_Person = Str(Err.Number)
    MsgBox "EDIT (DLL) - " & Err.Description
    
End Function

Function Edit_JubSu(sJubSu As String) As String
' 접수 Maste Update를 가장한 Insert
    
    Dim RCnt        As Integer
    Dim iCnt        As Integer
    Dim sSql        As String
    Dim sPartCd     As String
    Dim sLabDate    As String
    Dim sPartGbn    As String
    Dim sLabSeq     As String
    Dim sSpcCd      As String
    Dim sRegno      As String
    Dim sName       As String
    Dim sAge        As String
    Dim sIdLeft     As String
    Dim sSex        As String
    Dim sIdRight    As String
    Dim sDeptCd     As String
    Dim sWard       As String
    Dim sJubSuGbn   As String
    Dim sEmerGbn    As String
    Dim sDoctor     As String
    Dim sMedicalcmt As String
    Dim sSpecialOrd As String
    
    sPartCd = GetByOne(sJubSu, sJubSu)        '1
    sLabSeq = GetByOne(sJubSu, sJubSu)        '0
    sLabDate = GetByOne(sJubSu, sJubSu)       '2
    sPartGbn = GetByOne(sJubSu, sJubSu)       '3
    sSpcCd = GetByOne(sJubSu, sJubSu)         '4
    sRegno = GetByOne(sJubSu, sJubSu)         '5
    sName = GetByOne(sJubSu, sJubSu)          '6
    sAge = GetByOne(sJubSu, sJubSu)           '7
    sIdLeft = GetByOne(sJubSu, sJubSu)        '8
    sSex = GetByOne(sJubSu, sJubSu)           '9
    sIdRight = GetByOne(sJubSu, sJubSu)       '10
    sDeptCd = GetByOne(sJubSu, sJubSu)        '11
    sWard = GetByOne(sJubSu, sJubSu)          '12
    sJubSuGbn = GetByOne(sJubSu, sJubSu)      '13
    sEmerGbn = GetByOne(sJubSu, sJubSu)       '14
    sDoctor = GetByOne(sJubSu, sJubSu)        '15
    sMedicalcmt = GetByOne(sJubSu, sJubSu)    '16
    sSpecialOrd = GetByOne(sJubSu, sJubSu)    '17
    
    On Error GoTo ErrHandler
    
    sSql = " UPDATE H_JUBSU " & _
           "    SET SPECIMENCD = '" & sSpcCd & "'," & _
           "        REGNO = '" & sRegno & "'," & _
           "        NAME  = '" & sName & "'," & _
           "        AGE   = '" & sAge & "'," & _
           "        IDLEFT = '" & sIdLeft & "'," & _
           "        SEX = '" & sSex & "'," & _
           "        IDRIGHT = '" & sIdRight & "'," & _
           "        DEPTCD = '" & sDeptCd & "'," & _
           "        WARD = '" & sWard & "'," & _
           "        JUBSUGBN = '" & sJubSuGbn & "'," & _
           "        EMERGBN = '" & sEmerGbn & "'," & _
           "        DOCTOR = '" & sDoctor & "'," & _
           "        MEDICALREMARK = '" & sMedicalcmt & "'," & _
           "        SPECIALORD = '" & sSpecialOrd & "'" & _
           "  WHERE LABDATE = '" & sLabDate & "'" & _
           "    AND PARTGBN = '" & sPartGbn & "'" & _
           "    AND LABSEQ = '" & sLabSeq & "'"

    AdoConn.Execute sSql
    
    Edit_JubSu = "SUCCESS"
Exit Function
ErrHandler:
    Edit_JubSu = Str(Err.Number)
    MsgBox "EDIT (DLL) - " & Err.Description
    
End Function

'U3) 접수 Detail한 정보를 수정한다.
Function Edit_Result(sPerson As String, sAge As String, sLabDate As String, sSlipCd As String, sLabSeq As String) As String
' 접수 내역 Detail Update

    Dim RCnt        As Integer
    Dim iCnt        As Integer
    Dim sSql        As String
    Dim sPartCd     As String
    Dim sPartGbn    As String
    Dim sRegno      As String
    Dim sName       As String
    Dim sBoneYear   As String
    Dim sIdLeft     As String
    Dim sSex        As String
    Dim sIdRight    As String
    
    sPartCd = Left(sSlipCd, 1)
    sPartGbn = Mid(sSlipCd, 2, 2)
'sPerson = txtRegNo.Text & "|" & txtname.Text & "|" & sBornYear & "|" & _
'              txtidleft.Text & "|" & sSex & "|" & txtidright.Text & "|"
              
    sRegno = GetByOne(sPerson, sPerson)
    sName = GetByOne(sPerson, sPerson)
    sBoneYear = GetByOne(sPerson, sPerson)  ' 생년월일
    
    sIdLeft = GetByOne(sPerson, sPerson)
    sSex = GetByOne(sPerson, sPerson)
    sIdRight = GetByOne(sPerson, sPerson)
    
    On Error GoTo ErrHandler

    sSql = " UPDATE H_RESULT " & _
           "    SET REGNO = '" & sRegno & "'," & _
           "        NAME = '" & sName & "'," & _
           "        AGE = '" & sAge & "'," & _
           "        IDLEFT = '" & sIdLeft & "'," & _
           "        SEX = '" & sSex & "'," & _
           "        IDRIGHT = '" & sIdRight & "'" & _
           "  WHERE LABDATE     = '" & sLabDate & "'" & _
           "    AND PARTGBN     = '" & sPartGbn & "'" & _
           "    AND LABSEQ      = '" & sLabSeq & "'"
           
    AdoConn.Execute sSql

    Edit_Result = "SUCCESS"
Exit Function
ErrHandler:
    Edit_Result = Str(Err.Number)
    MsgBox "EDIT (DLL) - " & Err.Description

End Function

Function Edit_LabNo(LabDate As String, PartCd As String, PartGbn As String, LabSeq As String) As String
' 검체번호를 증가시킨후 재등록 시킴

    Dim sSql        As String
    
    On Error GoTo ErrHandler
    
    sSql = " UPDATE H_LABNO " & _
           "    SET LABSEQ = '" & LabSeq & "'" & _
           "  WHERE LABDATE = '" & LabDate & "'" & _
           "    AND PARTGBN = '" & PartGbn & "'"

    AdoConn.Execute sSql
    
    Edit_LabNo = "SUCCESS"
Exit Function
ErrHandler:
    Edit_LabNo = Str(Err.Number)
    MsgBox "EDIT (DLL) - " & Err.Description

End Function

Function Del_JubSu(sJubSu As String) As String
' 접수 내역 Maste Delete
    Dim sSql        As String
    Dim sPartCd     As String
    Dim sSlipCd     As String
    Dim sLabDate    As String
    Dim sPartGbn    As String
    Dim sLabSeq     As String
    
    sLabDate = GetByOne(sJubSu, sJubSu)
    sSlipCd = GetByOne(sJubSu, sJubSu)
    sLabSeq = GetByOne(sJubSu, sJubSu)
    sPartCd = Left(sSlipCd, 1)
    sPartGbn = Mid(sSlipCd, 2, 2)
    
    Del_JubSu = "ERROR"
    
    On Error GoTo ErrHandler
    
    sSql = " DELETE FROM H_JUBSU " & _
           "  WHERE LABDATE = '" & sLabDate & "'" & _
           "    AND PARTGBN = '" & sPartGbn & "'" & _
           "    AND LABSEQ = '" & sLabSeq & "'"
           
    AdoConn.Execute sSql
    
    Del_JubSu = "SUCCESS"
    
Exit Function
ErrHandler:
    Del_JubSu = Str(Err.Number)
    MsgBox "DEL (JubSu DLL) - " & Err.Description
    
End Function

'D2) 접수 Detail내용을 삭제한다.
Function Del_Result(sResult As String, sGbn As String) As String
' 접수 상세 내역 Delete 종류에 따라

    Dim sSql        As String
    Dim sPartCd     As String
    Dim sLabDate    As String
    Dim sSlipCd     As String
    Dim sPartGbn    As String
    Dim sLabSeq     As String
    
    sLabDate = GetByOne(sResult, sResult)
    sSlipCd = GetByOne(sResult, sResult)
    sLabSeq = GetByOne(sResult, sResult)
    sPartCd = Left(sSlipCd, 1)
    sPartGbn = Mid(sSlipCd, 2, 2)
    
    Del_Result = "ERROR"
    
    On Error GoTo ErrHandler
      
    sSql = " DELETE FROM H_RESULT " & _
           "  WHERE LABDATE = '" & sLabDate & "'" & _
           "    AND PARTGBN = '" & sPartGbn & "'" & _
           "    AND LABSEQ = '" & sLabSeq & "'"
    If sGbn = "PART" Then
       sSql = sSql & "  AND (TRIM(RESULT) = '' OR RESULT IS NULL) "
    End If
'    If Kind = "P" Then
'        sSql = sSql & "    AND SPECIMENCD = '" & sSpcCd & "'" & _
'                      "    AND TESTITEMSEQ = '" & sTestItemSeq & "'" & _
'                      "    AND SUBMCD = '" & sSubMcd & "'"
'    End If
    
    AdoConn.Execute sSql
        
    Del_Result = "SUCCESS"

Exit Function
ErrHandler:
    Del_Result = Str(Err.Number)
    MsgBox "DEL (JubSu DLL) - " & Err.Description
        
End Function

'S1) Get The Now Date & Time
'sopt "D" date형 "T" type형 "A" D와 T형을 모두 포함
Function Get_Date(sopt As String) As String
    
    Get_Date = sNow_Date(sopt)
    
End Function

'S2) Get Default SlipCd & SpcCd
Function Get_User_Env(UserCd As String) As String

    Dim AdoRecord   As New ADODB.Recordset
    Dim sSql        As String
    Dim sSlipCd     As String * 3
    Dim sSpcCd      As String * 3
    Dim sSlipNm     As String * 40
    Dim sSpcNm      As String * 30
    
    sSlipCd = "": sSpcCd = "": sSlipNm = "": sSpcNm = ""
    
    On Error GoTo ErrHandler
    sSql = "SELECT USER.SLIPCD, SLIPNM, USER.SPECIMENCD, SPECIMENBRIEFNM " & _
            " FROM USER, PART, SPECIMEN " & _
            "WHERE USERCD          = '" & UserCd & "'" & _
            "  AND USER.SLIPCD     = PART.PARTCD + PART.PARTGBN " & _
            "  AND USER.SPECIMENCD = SPECIMEN.SPECIMENCD "
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
' Packet 사용자 기본 슬립코드|슬립명|검체코드|검체명
        AdoRecord.MoveFirst
        
        sSlipCd = AdoRecord("SLIPCD") & ""
        sSlipNm = AdoRecord("SLIPNM") & ""
        sSpcCd = AdoRecord("SPECIMENCD") & ""
        sSpcNm = AdoRecord("SPECIMENBRIEFNM") & ""
    End If

    AdoRecord.Close
    
    Get_User_Env = sSlipCd & sSpcCd & sSlipNm & sSpcNm

Exit Function
ErrHandler:
    MsgBox "GetItem(접수 DLL) - " & Err.Description
    
End Function

'S3) Get SlipName to SlipCode
Function Get_SlipNm(SlipCd As String) As String

    Dim AdoRecord   As New ADODB.Recordset
    Dim sSql        As String
    Dim sSlipCd     As String * 3
    Dim sSlipNm     As String * 40
    
    sSlipCd = "": sSlipNm = ""
    
    On Error GoTo ErrHandler
    sSql = "SELECT SLIPNM " & _
            " FROM PART " & _
            "WHERE PARTCD + PARTGBN = '" & SlipCd & "'"
            
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
' Packet 슬립명
        AdoRecord.MoveFirst
        sSlipNm = AdoRecord("SLIPNM") & ""
    End If

    AdoRecord.Close
    
    Get_SlipNm = sSlipNm

Exit Function

ErrHandler:
    MsgBox "GetItem(접수 DLL) - " & Err.Description
    
End Function

'S4) Get Specimen Name to Specimen Code
Function Get_SpcNm(SpcCd As String) As String

    Dim AdoRecord   As New ADODB.Recordset
    Dim sSql        As String
    Dim sSpcCd      As String * 3
    Dim sSpcNm      As String * 30

    sSpcCd = "": sSpcNm = ""
    
    On Error GoTo ErrHandler
    
    sSql = "SELECT SPECIMENBRIEFNM " & _
            " FROM SPECIMEN " & _
            "WHERE SPECIMENCD = '" & SpcCd & "'"
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
' Packet 검체명
        AdoRecord.MoveFirst
        sSpcNm = AdoRecord("SPECIMENBRIEFNM") & ""
    End If

    AdoRecord.Close
    
    Get_SpcNm = sSpcNm

Exit Function
ErrHandler:
    MsgBox "GetItem(접수 DLL) - " & Err.Description
    
End Function

'S5) 환자 정보 취득(MASTER)
Function Get_RegInfo(RegNo As String) As String
    
    Dim AdoRecord   As New ADODB.Recordset
    Dim sSql        As String
    Dim sName       As String
    Dim sBornYear   As String
    Dim sAge        As String
    Dim sIdLeft     As String
    Dim sSex        As String
    Dim sIdRight    As String
    
    sName = "": sAge = "": sIdLeft = "": sSex = "": sIdRight = ""
    
    On Error GoTo ErrHandler
    
    sSql = "SELECT NAME,BORNYEAR,IDLEFT,SEX,IDRIGHT " & _
            " FROM PERSON " & _
            "WHERE REGNO = '" & RegNo & "'"
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
' Packet 주민번호왼쪽|주민번호오른쪽|나이|성별|이름
        AdoRecord.MoveFirst
        sName = AdoRecord("NAME") & "|"
        sBornYear = AdoRecord("BORNYEAR") & "|"
        sIdLeft = AdoRecord("IDLEFT") & "|"
        sSex = AdoRecord("SEX") & ""
        sIdRight = AdoRecord("IDRIGHT") & "|"
        If Trim(sSex) = "1" Or Trim(sSex) = "3" Then
            sSex = "남" & "|"
        ElseIf Trim(sSex) = "2" Or Trim(sSex) = "4" Then
            sSex = "여" & "|"
        Else
            sSex = "|"
        End If
        sAge = Str(Val(Left(Get_Date("D"), 4)) - Val(sBornYear)) & "|"
    End If

    AdoRecord.Close
    
    Get_RegInfo = sIdLeft & sIdRight & sAge & sSex & sName
    
Exit Function
ErrHandler:
    MsgBox "GetItem(접수 DLL) - " & Err.Description
    
End Function

Function Get_NameInfo(RegNo As String) As String
    
    Dim AdoRecord   As New ADODB.Recordset
    Dim sSql        As String
    Dim sName       As String
    Dim sBornYear   As String
    Dim sAge        As String
    Dim sIdLeft     As String
    Dim sSex        As String
    Dim sIdRight    As String
    
    '-- 2001.08.08(osw)
    Dim sRegno      As String
    
    sName = "": sAge = "": sIdLeft = "": sSex = "": sIdRight = ""
    sRegno = ""
    
    On Error GoTo ErrHandler
    
    sSql = "SELECT REGNO,BORNYEAR,IDLEFT,SEX,IDRIGHT " & _
            " FROM PERSON " & _
            "WHERE NAME = '" & RegNo & "'"
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
' Packet 주민번호왼쪽|주민번호오른쪽|나이|성별|이름
        AdoRecord.MoveFirst
        sRegno = AdoRecord("REGNO") & "|"
        sBornYear = AdoRecord("BORNYEAR") & "|"
        sIdLeft = AdoRecord("IDLEFT") & "|"
        sSex = AdoRecord("SEX") & ""
        sIdRight = AdoRecord("IDRIGHT") & "|"
        If Trim(sSex) = "1" Or Trim(sSex) = "3" Then
            sSex = "남" & "|"
        ElseIf Trim(sSex) = "2" Or Trim(sSex) = "4" Then
            sSex = "여" & "|"
        Else
            sSex = "|"
        End If
        sAge = Str(Val(Left(Get_Date("D"), 4)) - Val(sBornYear)) & "|"
    End If

    AdoRecord.Close
    
    Get_NameInfo = sIdLeft & sIdRight & sAge & sSex & sRegno
    
Exit Function
ErrHandler:
    MsgBox "GetItem(접수 DLL) - " & Err.Description
    
End Function

'S6) 진료과명 취득
Function Get_DeptNm(DeptCd As String) As String

    Dim AdoRecord   As New ADODB.Recordset
    Dim sSql        As String
    Dim sDeptNm     As String * 30

    sDeptNm = ""
    
    On Error GoTo ErrHandler
    
    sSql = "SELECT DEPTNM " & _
            " FROM DEPT " & _
            "WHERE DEPTCD = '" & DeptCd & "'"
            
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
' Packet 진료과명
        AdoRecord.MoveFirst
        sDeptNm = AdoRecord("DEPTNM") & ""
    End If

    AdoRecord.Close
    
    Get_DeptNm = sDeptNm

Exit Function
ErrHandler:
    MsgBox "GetItem(접수 DLL) - " & Err.Description

End Function

Function Get_JubSu(LabDate As String, SlipCd As String, LabSeq As String) As String
    
    Dim AdoRecord   As New ADODB.Recordset
    Dim sSql        As String
    Dim sJubSu      As String
    Dim RCnt        As Integer
    Dim iCnt        As Integer
        
    sJubSu = ""
    
    On Error GoTo ErrHandler

' 신상자료 취득
    sSql = " SELECT SPECIMENCD, JUBSU.REGNO, JUBSU.NAME, AGE, JUBSU.IDLEFT, JUBSU.SEX, " & _
           "        JUBSU.IDRIGHT, DEPT.DEPTCD, DEPT.DEPTNM, WARD, JUBSUGBN, EMERGBN, " & _
           "        DOCTOR, MEDICALREMARK, SPECIALORD " & _
           "   FROM PERSON RIGHT JOIN ( H_JUBSU JUBSU LEFT JOIN DEPT DEPT " & _
           "     ON JUBSU.DEPTCD = DEPT.DEPTCD) ON PERSON.REGNO = JUBSU.REGNO " & _
           "  WHERE LABDATE = '" & LabDate & "'" & _
           "    AND PARTGBN = '" & Mid(SlipCd, 2, 2) & "'" & _
           "    AND LABSEQ  = '" & LabSeq & "'"

    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
'Packet 1.검체코드|2.등록번호|3.이름|4.나이|5.주민번호왼쪽|6.성별|7.주민번호오른쪽|
'        8.진료과번호|9.진료과명|10.병동|11.접수구분|12.응급구분|13.Doctor|
'        14.병상명|15.특진구분
        AdoRecord.MoveFirst
        sJubSu = AdoRecord("SPECIMENCD") & "|"
        sJubSu = sJubSu & AdoRecord("REGNO") & "|"
        sJubSu = sJubSu & AdoRecord("NAME") & "|"
        sJubSu = sJubSu & AdoRecord("AGE") & "|"
        sJubSu = sJubSu & AdoRecord("IDLEFT") & "|"
        sJubSu = sJubSu & AdoRecord("SEX") & "|"
        sJubSu = sJubSu & AdoRecord("IDRIGHT") & "|"
        sJubSu = sJubSu & AdoRecord("DEPTCD") & "|"
        sJubSu = sJubSu & AdoRecord("DEPTNM") & "|"
        sJubSu = sJubSu & AdoRecord("WARD") & "|"
        sJubSu = sJubSu & AdoRecord("JUBSUGBN") & "|"
        sJubSu = sJubSu & AdoRecord("EMERGBN") & "|"
        sJubSu = sJubSu & AdoRecord("DOCTOR") & "|"
        sJubSu = sJubSu & AdoRecord("MEDICALREMARK") & "|"
        sJubSu = sJubSu & AdoRecord("SPECIALORD") & "|"
    End If
    
    AdoRecord.Close
    
' 접수자료 취득(결과입력 여부 확인정보 필요)
    sSql = " SELECT '" & Left(SlipCd, 1) & "' & RESULT.PARTGBN & RESULT.SPECIMENCD & RESULT.TESTITEMSEQ & RESULT.SUBMCD AS ORDCD, ITEM.PRINTNM, RESULT.ROUTINECD " & _
           "   FROM H_RESULT RESULT, TESTITEM ITEM" & _
           "  WHERE RESULT.LABDATE   = '" & LabDate & "'" & _
           "    AND RESULT.PARTGBN   = '" & Mid(SlipCd, 2, 2) & "'" & _
           "    AND RESULT.LABSEQ    = '" & LabSeq & "'" & _
           "    AND RESULT.SUBMCD      IN ('NNNN','00NN') " & _
           "    AND ITEM.PARTCD      = '" & Left(SlipCd, 1) & "'" & _
           "    AND ITEM.PARTGBN     = RESULT.PARTGBN " & _
           "    AND ITEM.SPECIMENCD  = RESULT.SPECIMENCD " & _
           "    AND ITEM.TESTITEMSEQ = RESULT.TESTITEMSEQ " & _
           "    AND ITEM.SUBMCD      = RESULT.SUBMCD " & _
           "  ORDER BY ITEM.PARTGBN & ITEM.SPECIMENCD & ITEM.TESTITEMSEQ & ITEM.SUBMCD  "

    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
        AdoRecord.MoveFirst
        RCnt = AdoRecord.RecordCount
        
        sJubSu = sJubSu & Str(RCnt) & "|"                   ' 검사건수
        For iCnt = 1 To RCnt
            sJubSu = sJubSu & AdoRecord("PRINTNM") & "|"    ' 검사명
            sJubSu = sJubSu & AdoRecord("ORDCD") & "|"      ' 검사코드
            sJubSu = sJubSu & AdoRecord("ROUTINECD") & "|"  ' 집단검사명
            AdoRecord.MoveNext
        Next iCnt
    Else
        sJubSu = sJubSu & Str(0) & "|"                      ' 검사건수
    End If

    AdoRecord.Close

' 묶음 검사항목 존재시 처리
' 접수자료 취득(결과입력 여부 확인정보 필요)
    sSql = " SELECT DISTINCT RT.PARTCD + RT.ROUTINECD AS RTCD, RT.ROUTINENM " & _
           "   FROM H_RESULT RESULT, ROUTINE RT" & _
           "  WHERE RESULT.LABDATE           = '" & LabDate & "'" & _
           "    AND RESULT.PARTGBN           = '" & Mid(SlipCd, 2, 2) & "'" & _
           "    AND RESULT.LABSEQ            = '" & LabSeq & "'" & _
           "    AND RESULT.SUBMCD              IN ('NNNN','00NN') " & _
           "    AND RT.PARTCD + RT.ROUTINECD = RESULT.ROUTINECD "

    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
        AdoRecord.MoveFirst
        RCnt = AdoRecord.RecordCount
        
        sJubSu = sJubSu & Str(RCnt) & "|"                   ' 검사건수
        For iCnt = 1 To RCnt
            sJubSu = sJubSu & AdoRecord("ROUTINENM") & "|"  ' 집단검사코드
            sJubSu = sJubSu & AdoRecord("RTCD") & "|"       ' 집단검사명
            AdoRecord.MoveNext
        Next iCnt
    Else
        sJubSu = sJubSu & Str(0) & "|"                      ' 검사건수
    End If

    AdoRecord.Close
    
    Get_JubSu = sJubSu

Exit Function
ErrHandler:
    MsgBox "GetItem(접수 DLL) - " & Err.Description

End Function

'S8) 접수 List 취득
Function Get_JubSuList(LabDate As String, PartCd As String, sSearch As String) As String

    Dim AdoRecord   As New ADODB.Recordset
    Dim sSql        As String
    Dim sJubSuList  As String
    Dim sSlipCd     As String
    Dim sLabSeq     As String
    Dim RCnt        As Integer
    Dim iCnt        As Integer

    sJubSuList = ""
    
    On Error GoTo ErrHandler
        
    If sSearch = "YES" Then
' 검사완료된 검체 LIST
        sSql = " SELECT LABDATE, PARTGBN, LABSEQ " & _
               "  FROM H_JUBSU " & _
               " WHERE LABDATE = '" & LabDate & "'" & _
               "   AND REGNO IS NOT NULL " & _
               "   AND REGNO > '' " & _
               " ORDER BY LABDATE, PARTGBN, LABSEQ "
    ElseIf sSearch = "NO" Then
' 검사미완료된 검체 LIST
        sSql = " SELECT PARTGBN, LABSEQ " & _
               "  FROM H_JUBSU " & _
               " WHERE LABDATE = '" & LabDate & "'" & _
               "   AND (REGNO IS NULL " & _
               "     OR REGNO = '') " & _
               " ORDER BY LABDATE, PARTGBN, LABSEQ "
    ElseIf Left(sSearch, 1) = "R" Then
' 등록번호로 검체 LIST조회
        sSql = " SELECT PARTGBN, LABSEQ " & _
               "  FROM H_JUBSU " & _
               " WHERE LABDATE = '" & LabDate & "'" & _
               "   AND REGNO   = '" & Mid(sSearch, 2, 15) & "'" & _
               " ORDER BY LABDATE, PARTGBN, LABSEQ "
    Else
' 검체번호로 검체번호 조회
        sSql = " SELECT PARTGBN, LABSEQ " & _
               "  FROM H_JUBSU " & _
               " WHERE LABDATE = '" & LabDate & "'" & _
               "   AND PARTGBN = '" & Left(sSearch, 2) & "'" & _
               "   AND LABSEQ  = '" & Mid(sSearch, 3, 5) & "'"
    End If
            
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
' 검체건수 |{ Slip코드 | 작업일련번호}n
        AdoRecord.MoveFirst
        RCnt = AdoRecord.RecordCount
        sJubSuList = Str(RCnt) & "|"
        For iCnt = 1 To RCnt
            If AdoRecord("PARTGBN") = "01" Then
                sSlipCd = "C01" & "|"
            ElseIf AdoRecord("PARTGBN") = "02" Then
                sSlipCd = "H02" & "|"
            Else
                sSlipCd = "U04" & "|"
            End If
'            sSlipCd = PartCd & AdoRecord("PARTGBN") & "|"
            sLabSeq = AdoRecord("LABSEQ") & "|"
            sJubSuList = sJubSuList & sSlipCd & sLabSeq
            AdoRecord.MoveNext
        Next iCnt
    End If
    AdoRecord.Close
    
    Get_JubSuList = sJubSuList

Exit Function
ErrHandler:
    MsgBox "GetItem(접수 DLL) - " & Err.Description
    
End Function

'S9) Routine코드의 검사명과 풀어진 검사항목 취득
Function Get_Routine(PartCd As String, RCd As String, SpcCd As String, UserCd As String) As String

    Dim AdoRecord       As New ADODB.Recordset
        
    Dim RCnt            As Integer
    Dim iCnt            As Integer
    Dim sSql            As String
    Dim sRountineNm     As String
    Dim sItemCd         As String
    Dim sPrintNm        As String
    
    Dim sTESTITEMCD     As String
    Dim sPartCd         As String
    Dim sPartGbn        As String
    Dim sSpcCd          As String
    Dim sItemSeq        As String
    Dim sTestNmCfg      As String
    
    sRountineNm = ""
    sTestNmCfg = fGetCurTestItemNmCfg
    
    On Error GoTo ErrHandler
    
    If sTestNmCfg = "T" Then
        sSql = "SELECT ROUTINE.ROUTINENM, ROUTINE.PARTCD & ROUTINE.PARTGBN & ROUTINE.SPECIMENCD & ROUTINE.TESTITEMSEQ & TESTITEM.SUBMCD AS ITEMCD,TESTITEM.TESTITEMNM " & _
                " FROM ROUTINE, TESTITEM " & _
                "WHERE ROUTINE.PARTCD = '" & PartCd & "'" & _
                "  AND ROUTINE.ROUTINECD = '" & RCd & "'" & _
                "  AND ROUTINE.PARTCD = TESTITEM.PARTCD " & _
                "  AND ROUTINE.PARTGBN = TESTITEM.PARTGBN " & _
                "  AND ROUTINE.SPECIMENCD = '" & SpcCd & "'" & _
                "  AND ROUTINE.SPECIMENCD = TESTITEM.SPECIMENCD " & _
                "  AND ROUTINE.TESTITEMSEQ = TESTITEM.TESTITEMSEQ " & _
                "  AND ROUTINE.SUBMCD = TESTITEM.SUBMCD " & _
                "  AND TESTITEM.SUBMCD IN ('NNNN', '00NN') "
    ElseIf sTestNmCfg = "P" Then
        sSql = "SELECT ROUTINE.ROUTINENM, ROUTINE.PARTCD & ROUTINE.PARTGBN & ROUTINE.SPECIMENCD & ROUTINE.TESTITEMSEQ & TESTITEM.SUBMCD AS ITEMCD,TESTITEM.PRINTNM " & _
                " FROM ROUTINE, TESTITEM " & _
                "WHERE ROUTINE.PARTCD = '" & PartCd & "'" & _
                "  AND ROUTINE.ROUTINECD = '" & RCd & "'" & _
                "  AND ROUTINE.PARTCD = TESTITEM.PARTCD " & _
                "  AND ROUTINE.PARTGBN = TESTITEM.PARTGBN " & _
                "  AND ROUTINE.SPECIMENCD = '" & SpcCd & "'" & _
                "  AND ROUTINE.SPECIMENCD = TESTITEM.SPECIMENCD " & _
                "  AND ROUTINE.TESTITEMSEQ = TESTITEM.TESTITEMSEQ " & _
                "  AND ROUTINE.SUBMCD = TESTITEM.SUBMCD " & _
                "  AND TESTITEM.SUBMCD IN ('NNNN', '00NN') "
    End If
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
' Packet 묶음처방명 | 풀어진 검사건수 | {검사코드| 검사명|}n
        RCnt = AdoRecord.RecordCount
        AdoRecord.MoveFirst
        
        sRountineNm = AdoRecord("ROUTINENM") & "|"
        sRountineNm = sRountineNm & Str(RCnt) & "|"
        
        For iCnt = 1 To RCnt
            sItemCd = AdoRecord("ITEMCD") & "|"
            sPrintNm = AdoRecord(2) & "|"
            sRountineNm = sRountineNm & sPrintNm & sItemCd
            AdoRecord.MoveNext
        Next iCnt
        
    End If
    AdoRecord.Close
    
    Get_Routine = sRountineNm
   
Exit Function
ErrHandler:
    MsgBox "GetItem(접수 DLL) - " & Err.Description

End Function

'S10) 검사항목명 취득
Function Get_Order(SlipCd As String, SpcCd As String, OrdCd As String) As String

    Dim AdoRecord       As New ADODB.Recordset
    Dim sSql            As String
    Dim sPrintNm        As String
    Dim sDeltaGbn       As String
    Dim sPanicGbn       As String
    Dim sTestNmCfg      As String
    
    
    sPrintNm = ""
    sTestNmCfg = fGetCurTestItemNmCfg
    
    On Error GoTo ErrHandler
    
    If sTestNmCfg = "T" Then
        sSql = "SELECT TESTITEMNM, SUBMCD " & _
                " FROM TESTITEM " & _
                "WHERE PARTCD = '" & Left(SlipCd, 1) & "'" & _
                "  AND PARTGBN = '" & Mid(SlipCd, 2, 2) & "'" & _
                "  AND SPECIMENCD = '" & SpcCd & "'" & _
                "  AND TESTITEMSEQ = '" & OrdCd & "'"
    ElseIf sTestNmCfg = "P" Then
        sSql = "SELECT PRINTNM, SUBMCD " & _
                " FROM TESTITEM " & _
                "WHERE PARTCD = '" & Left(SlipCd, 1) & "'" & _
                "  AND PARTGBN = '" & Mid(SlipCd, 2, 2) & "'" & _
                "  AND SPECIMENCD = '" & SpcCd & "'" & _
                "  AND TESTITEMSEQ = '" & OrdCd & "'"
    End If
            
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
        AdoRecord.MoveFirst
' Packet 검사명
        sPrintNm = AdoRecord("SUBMCD") & "|"
        sPrintNm = sPrintNm & AdoRecord(0) & "|"
    End If
    AdoRecord.Close
    
    Get_Order = sPrintNm

Exit Function
ErrHandler:
    MsgBox "GetItem(접수 DLL) - " & Err.Description

End Function
'S11) 등록환자와 검사항목을 키로 하여 가장 최근의 검사결과를 얻어온다.
Function Get_LastResult(LabDate As String, LabSeq As String, RegNo As String, PartCd As String, PartGbn As String, SpcCd As String, TestItemSeq As String, SubMcd As String) As String
        
    Dim AdoRecord   As New ADODB.Recordset
    Dim sSql        As String
    Dim sLabNo      As String
    Dim sResult     As String
    Dim sInterval   As String
    
    sLabNo = "": sResult = "": sInterval = ""
    
    On Error GoTo ErrHandler
    
    sSql = "SELECT TOP 1 LABDATE & '" & PartCd & "' & PARTGBN & LABSEQ AS LABNO, RESULT" & _
            " FROM H_RESULT " & _
            "WHERE REGNO       = '" & RegNo & "'" & _
            "  AND PARTGBN     = '" & PartGbn & "'" & _
            "  AND SPECIMENCD  = '" & SpcCd & "'" & _
            "  AND TESTITEMSEQ = '" & TestItemSeq & "'" & _
            "  AND SUBMCD      = '" & SubMcd & "'" & _
            "  AND ((LABDATE    < '" & LabDate & "')" & _
            "  OR  (LABDATE    = '" & LabDate & "'" & _
            "  AND LABSEQ      < '" & LabSeq & "'))" & _
            "  AND RESULTDATE IS NOT NULL " & _
            "  AND RESULTTIME IS NOT NULL " & _
            "  ORDER BY LABDATE DESC, LABSEQ DESC"
            
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
' Packet 검체번호, 결과
        AdoRecord.MoveFirst
        
        sLabNo = AdoRecord("LABNO") & "|"
        sResult = AdoRecord("RESULT") & "|"
        sInterval = Trim(Str(CDate(Left(LabDate, 4) & "-" & Mid(LabDate, 5, 2) & "-" & Mid(LabDate, 7, 2)) - CDate(Left(sLabNo, 4) & "-" & Mid(sLabNo, 5, 2) & "-" & Mid(sLabNo, 7, 2)))) & "|"
    End If

    AdoRecord.Close
    
    Get_LastResult = sLabNo & sResult & sInterval

Exit Function
ErrHandler:
    AdoRecord.Close
    MsgBox "GetItem(접수 DLL) - " & Err.Description
    
End Function
'Function Over_Chk(LabDate As String, PartCd As String, PartGbn As String, PartSeq As String) As Boolean
'    Dim AdoRecord As New ADODB.Recordset
'    Dim sSql        As String
'    Dim sLabNo      As String
'
'    Over_Chk = False
'
'    sSql = "SELECT LABSEQ " & _
'           "  FROM H_LABNO " & _
'           " WHERE LABDATE = '" & LabDate & "'" & _
'           "   AND LABSEQ > '" & PartSeq & "'"
'
'    AdoRecord.CursorType = adOpenStatic
'    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
'
'    If AdoRecord.RecordCount <> 0 Then
'        AdoRecord.MoveFirst
'        Over_Chk = True
'    End If
'
'
'End Function
'S12) Get 해당 파트의 마지막 LabNo
Function Get_LastSeq(LabDate As String, PartCd As String, PartGbn As String) As String
    
    Dim AdoRecord As New ADODB.Recordset
    Dim sSql        As String
    Dim sLabNo      As String
    
    sLabNo = ""
    LabNo = ""
    
    On Error GoTo ErrHandler
    
    sSql = "SELECT LABSEQ " & _
           "  FROM H_LABNO " & _
           " WHERE LABDATE = '" & LabDate & "'" & _
           "   AND PARTGBN = '" & PartGbn & "'" & _
           "   AND REGNO = '" & RegNo & "'"
           
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
' Packet 검체번호, 결과
        AdoRecord.MoveFirst
                
        sLabNo = AdoRecord("LABSEQ") & ""
        sLabNo = Format(Trim(Val(sLabNo) + 1), "00000")
        
        AdoRecord.Close

        sSql = "SELECT LABSEQ " & _
               "  FROM H_JUBSU " & _
               " WHERE LABDATE = '" & LabDate & "'" & _
               "   AND PARTGBN = '" & PartGbn & "'"

        AdoRecord.CursorType = adOpenStatic
        AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"

        If AdoRecord.RecordCount <> 0 Then
            AdoRecord.MoveFirst
            sLabNo = Format(Trim(Val(sLabNo) + 1), "00000")
        End If
        
        If Edit_LabNo(LabDate, PartCd, PartGbn, sLabNo) <> "SUCCESS" Then
            sLabNo = "ERROR"
        End If
    Else
        AdoRecord.Close
        
        sSql = "SELECT LABSEQ " & _
               "  FROM H_LABNO " & _
               " WHERE LABDATE = '" & LabDate & "'" & _
               "   AND REGNO = '" & RegNo & "'"
    
        AdoRecord.CursorType = adOpenStatic
        AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"

        If AdoRecord.RecordCount <> 0 Then
            '-- 기존등록된 LABSEQ
            LabNo = AdoRecord("LABSEQ")
            If Add_LabNo(LabDate, PartCd, PartGbn) <> "SUCCESS" Then
                sLabNo = "ERROR"
            Else
                If LabNo <> "" Then
                    sLabNo = LabNo
                Else
                    sLabNo = "1"
                End If
            End If
        Else
            '-- 신규 & 추가 LABSEQ
            AdoRecord.Close
            
            sSql = "SELECT LABSEQ " & _
                   "  FROM H_LABNO " & _
                   " WHERE LABDATE = '" & LabDate & "'" & _
                   " Order By LABSEQ Desc "
                   
            AdoRecord.CursorType = adOpenStatic
            AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
            If AdoRecord.RecordCount <> 0 Then
                LabNo = AdoRecord("LABSEQ")
                LabNo = Format(Trim(Val(LabNo) + 1), "00000")
                
                If Add_LabNo(LabDate, PartCd, PartGbn) <> "SUCCESS" Then
                    sLabNo = "ERROR"
                Else
                    sLabNo = LabNo
                End If
            Else
                If Add_LabNo(LabDate, PartCd, PartGbn) <> "SUCCESS" Then
                    sLabNo = "ERROR"
                Else
                    sLabNo = "1"
                End If
            End If
            
        
        End If
    End If

    AdoRecord.Close
    
    Get_LastSeq = sLabNo
    
Exit Function
ErrHandler:
    MsgBox "GetItem(접수 DLL) - " & Err.Description
End Function
''S12) Get 해당 파트의 마지막 LabNo
'Function Get_LastSeq_back(LabDate As String, PartCd As String, PartGbn As String) As String
'
'    Dim AdoRecord As New ADODB.Recordset
'    Dim sSql        As String
'    Dim sLabNo      As String
'
'    sLabNo = ""
'    LabNo = ""
'
'    On Error GoTo ErrHandler
'
'    sSql = "SELECT LABSEQ " & _
'           "  FROM H_LABNO " & _
'           " WHERE LABDATE = '" & LabDate & "'" & _
'           "   AND PARTGBN = '" & PartGbn & "'"
'
'    '-- 접수내역이 있는지 확인(등록일 & 등록번호)
''    sSql = "SELECT LABSEQ " & _
'           "  FROM H_JUBSU " & _
'           " WHERE LABDATE = '" & LabDate & "'" & _
'           "   AND REGNO = '" & RegNo & "'"
'
'    AdoRecord.CursorType = adOpenStatic
'    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
'
'    If AdoRecord.RecordCount <> 0 Then
'        AdoRecord.MoveFirst
'        sLabNo = AdoRecord("LABSEQ") & ""
'        LabNo = sLabNo
'    Else
'
'    End If
'
'    AdoRecord.Close
'
'    sSql = "SELECT LABSEQ " & _
'           "  FROM H_LABNO " & _
'           " WHERE LABDATE = '" & LabDate & "'" & _
'           "   AND PARTGBN = '" & PartGbn & "'"
'
'    AdoRecord.CursorType = adOpenStatic
'    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
'
'    If AdoRecord.RecordCount <> 0 Then
'        AdoRecord.MoveFirst
'
'        If Trim(sLabNo) = "" Then
'            sLabNo = AdoRecord("LABSEQ") & ""
'            sLabNo = Format(Trim(Val(sLabNo) + 1), "00000")
'        End If
'
'        If Edit_LabNo(LabDate, PartCd, PartGbn, sLabNo) <> "SUCCESS" Then
'            sLabNo = "ERROR"
'        End If
'    Else
'        AdoRecord.Close
'
'        sSql = "SELECT LABSEQ " & _
'               "  FROM H_LABNO " & _
'               " WHERE LABDATE = '" & LabDate & "' " & _
'               " Order By LABSEQ Desc "
'
'        AdoRecord.CursorType = adOpenStatic
'        AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
'
'        If AdoRecord.RecordCount <> 0 Then
'            AdoRecord.MoveFirst
'
'            If sLabNo = "" Then
'                sLabNo = AdoRecord("LABSEQ") & ""
'                sLabNo = Format(Trim(Val(sLabNo) + 1), "00000")
'                LabNo = sLabNo
'            End If
'            If Add_LabNo(LabDate, PartCd, PartGbn) <> "SUCCESS" Then
'                sLabNo = "ERROR"
'            Else
'                If LabNo <> "" Then
'                    sLabNo = LabNo
'                Else
'                    sLabNo = "1"
'                End If
'            End If
'        Else
'            ' INSERT LABSEQ
'            If Add_LabNo(LabDate, PartCd, PartGbn) <> "SUCCESS" Then
'                sLabNo = "ERROR"
'            Else
'                sLabNo = "1"
'            End If
'        End If
'    End If
'
'    AdoRecord.Close
'
'    If sLabNo <> "" Then
'        Get_LastSeq = sLabNo
'    Else
'        Get_LastSeq = "1"
'    End If
'Exit Function
'ErrHandler:
'    MsgBox "GetItem(접수 DLL) - " & Err.Description
'End Function

'S13) 결과의 갯수를 파악하는 함수
Function Get_ResultCnt(LabDate As String, SlipCd As String, LabSeq As String) As Integer

    Dim AdoRecord       As New ADODB.Recordset
    Dim sSql            As String
    Dim iResCnt         As Integer
    
    On Error GoTo ErrHandler
    
    sSql = "SELECT LABDATE " & _
            " FROM H_RESULT " & _
            "WHERE LABDATE = '" & LabDate & "'" & _
            "  AND PARTGBN = '" & Mid(SlipCd, 2, 2) & "'" & _
            "  AND LABSEQ = '" & LabSeq & "'"
            
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
        iResCnt = AdoRecord.RecordCount
    Else
        iResCnt = 0
    End If
    AdoRecord.Close
    
    Get_ResultCnt = iResCnt

Exit Function
ErrHandler:
    MsgBox "GetItem(접수 DLL) - " & Err.Description

End Function
Function Get_LastLabNo(LabDate As String, PartCd As String, PartGbn As String) As String
    
    Dim AdoRecord As New ADODB.Recordset
    Dim sSql        As String
    Dim sLabNo      As String
    
    sLabNo = ""
    
    On Error GoTo ErrHandler
    
    sSql = "SELECT LABSEQ " & _
           "  FROM H_LABNO " & _
           " WHERE LABDATE = '" & LabDate & "'" & _
           "   AND PARTGBN = '" & PartGbn & "'"

    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;", adOpenForwardOnly
    
    If AdoRecord.EOF = False Then
' Packet 검체번호, 결과
        sLabNo = AdoRecord("LABSEQ") & ""
    Else
        sLabNo = ""
    End If

    AdoRecord.Close
    
    Get_LastLabNo = sLabNo
    
Exit Function
ErrHandler:
    MsgBox "GetItem(접수 DLL) - " & Err.Description


End Function

Function Tran_JubSu(sTrans1 As String, sTrans1_F As Integer, sTrans2 As String, sTrans2_F As Integer, sTrans3 As String, sTrans3_F As Integer, sTrans4 As String) As String
    
    Dim sTrans5     As String
    Dim sPartCd     As String
    Dim sPartGbn    As String
    Dim sLabDate    As String
    Dim sSlipCd     As String
    Dim sLabSeq     As String
    Dim sTrans      As String
    Dim sAge        As String
    Dim stmp        As String ' 사용하지 않는 스트링
    
    Dim sRtnVal     As String

    AdoConn.CursorLocation = adUseServer
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    AdoConn.BeginTrans
    
' 환자 신상정보
        If Trim(sTrans1) <> "" Then
            sTrans5 = sTrans1
            If sTrans1_F = True Then    ' 기존 존재 환자
                If Edit_Person(sTrans5) <> "SUCCESS" Then
                    AdoConn.RollbackTrans
                    AdoConn.Close
                    Exit Function
                End If
            
            Else                        ' 완전 신규 환자
                sRtnVal = Add_Person(sTrans5)

                If sRtnVal = "-2147217900" Then     '중복 Insert인 경우 Update
                    If Edit_Person(sTrans5) <> "SUCCESS" Then
                        AdoConn.RollbackTrans
                        AdoConn.Close
                        Exit Function
                    End If
                ElseIf sRtnVal <> "SUCCESS" Then
                    AdoConn.RollbackTrans
                    AdoConn.Close
                    Exit Function
                End If
            End If
        End If
        
' 환자 접수정보
        If sTrans2_F = True Then   ' 환자 신상자료를 수정했을 경우
        ' 삭제할 Key값을 조사
            sTrans = sTrans2
            sPartCd = GetByOne(sTrans, sTrans)
            sLabSeq = GetByOne(sTrans, sTrans)
            sLabDate = GetByOne(sTrans, sTrans)
            sPartGbn = GetByOne(sTrans, sTrans)
            stmp = GetByOne(sTrans, sTrans)
            stmp = GetByOne(sTrans, sTrans)
            stmp = GetByOne(sTrans, sTrans)
            sAge = GetByOne(sTrans, sTrans)
            
            sSlipCd = sPartCd & sPartGbn
            sTrans = ""
            
            sTrans5 = sLabDate & "|" & sSlipCd & "|" & sLabSeq & "|"
            sTrans4 = sLabDate & "|" & sSlipCd & "|" & sLabSeq & "|"
            
            sTrans = GetByOne(sTrans3, sTrans3)
            sTrans = GetByOne(sTrans3, sTrans3)
            sTrans = GetByOne(sTrans3, sTrans3)
            sTrans3 = sPartCd & "|" & sLabDate & "|" & sTrans3
' 결과 자료 삭제
            If Del_Result(sTrans5, "PART") <> "SUCCESS" Then
                AdoConn.RollbackTrans
                AdoConn.Close
                Exit Function
            End If

' 결과 자료 부분 UPDATE
            If Edit_Result(sTrans1, sAge, sLabDate, sSlipCd, sLabSeq) <> "SUCCESS" Then
                AdoConn.RollbackTrans
                AdoConn.Close
            End If
            
' 접수 Master수정
            If Edit_JubSu(sTrans2) <> "SUCCESS" Then
                AdoConn.RollbackTrans
                AdoConn.Close
                Exit Function
            End If
' 결과 자료 신규등록
            If sTrans3_F = True Then    ' 검사항목까지 수정했을 경우
                If Add_Result(sLabSeq & "|" & sTrans3) <> "SUCCESS" Then
                    AdoConn.RollbackTrans
                    AdoConn.Close
                    Exit Function
                End If
            End If
        Else
            sLabSeq = Add_JubSu(sTrans2)  ' 접수 신규 환자
            If IsNumeric(sLabSeq) = False Then
                AdoConn.RollbackTrans
                AdoConn.Close
                Exit Function
            Else
' 검체의 검사정보
                If Add_Result(sLabSeq & "|" & sTrans3) <> "SUCCESS" Then
                    AdoConn.RollbackTrans
                    AdoConn.Close
                    Exit Function
                End If
            End If
        End If
    AdoConn.CommitTrans
    
    AdoConn.BeginTrans
' Delta정보 처리
    If Tran_Delta(sLabSeq & "|" & sTrans4) <> "SUCCESS" Then
        AdoConn.RollbackTrans
        AdoConn.Close
        Exit Function
    End If
    AdoConn.CommitTrans
    
    AdoConn.Close
        
    Tran_JubSu = sLabSeq
    
End Function

Function Tran_JubSu_Del(sTrans1 As String) As String
    
    Dim LabNo       As String
    Dim sDelStr     As String
    
    Tran_JubSu_Del = "ERROR"
    
    AdoConn.CursorLocation = adUseServer
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    AdoConn.BeginTrans
        
' 환자 접수정보
    sDelStr = sTrans1
    If Del_JubSu(sDelStr) <> "SUCCESS" Then
        AdoConn.RollbackTrans
        AdoConn.Close
        Exit Function
    End If
            
' 검체의 검사정보
    sDelStr = sTrans1
    If Del_Result(sDelStr, "ALL") <> "SUCCESS" Then
        AdoConn.RollbackTrans
        AdoConn.Close
        Exit Function
    End If
    
    AdoConn.CommitTrans
    AdoConn.Close
    
    Tran_JubSu_Del = "SUCCESS"
        
End Function

'T3) Delta처리를 한다.
Function Tran_Delta(sTrans1 As String) As String

    Dim AdoRecord       As New ADODB.Recordset
    
    Dim RCnt        As Integer
    Dim iCnt        As Integer
    Dim sSql        As String
    Dim sLabDate    As String
    Dim sPartCd     As String
    Dim sSlipCd     As String
    Dim sPartGbn    As String
    Dim sLabSeq     As String
    Dim sRegno      As String
    Dim sSpcCd      As String
    Dim sTestSeq    As String
    Dim sSubMcd     As String
    Dim sResult     As String
    
    Dim sLastLabNo  As String
    Dim sLastResult As String
    Dim sInterval   As String
    
    sLabSeq = GetByOne(sTrans1, sTrans1)    '1
    sLabDate = GetByOne(sTrans1, sTrans1)   '2
    sSlipCd = GetByOne(sTrans1, sTrans1)    '3
    sPartCd = Left(sSlipCd, 1)
    sPartGbn = Mid(sSlipCd, 2, 2)
        
' Delta를 해야 하는 ITEM만을 신상정보와 함께 취득(물론 Regno가 존재해야함)
' SELECT
    On Error GoTo ErrHandler
' 등록번호가 존재하고, Delta항목에 걸리는 것만 처리
    sSql = "SELECT REGNO, SPECIMENCD, TESTITEMSEQ, SUBMCD " & _
            " FROM H_RESULT " & _
            "WHERE LABDATE = '" & sLabDate & "'" & _
            "  AND PARTGBN = '" & sPartGbn & "'" & _
            "  AND LABSEQ  = '" & sLabSeq & "'" & _
            "  AND DELTAGBN IN ('1','2') "
            
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount <> 0 Then
' Packet 묶음처방명 | 풀어진 검사건수 | {검사코드| 검사명|}n
        RCnt = AdoRecord.RecordCount
        AdoRecord.MoveFirst
        
        For iCnt = 1 To RCnt
            sRegno = AdoRecord("REGNO") & ""
            sSpcCd = AdoRecord("SPECIMENCD") & ""
            sTestSeq = AdoRecord("TESTITEMSEQ") & ""
            sSubMcd = AdoRecord("SUBMCD") & ""
    ' 가져온 각 ITEM에 대해 가장 최근의 결과를 취득하여 다시 UPDATE처리
    ' SELECT
            sResult = Get_LastResult(sLabDate, sLabSeq, sRegno, sPartCd, sPartGbn, sSpcCd, sTestSeq, sSubMcd)
            sLastLabNo = GetByOne(sResult, sResult)
            sLastResult = GetByOne(sResult, sResult)
            sInterval = GetByOne(sResult, sResult)
            If Trim(sInterval) = "" Then
                sInterval = 0
            End If
    ' UPDATE
            sSql = " UPDATE H_RESULT " & _
                   "    SET BEFORELABNO  = '" & sLastLabNo & "', " & _
                   "        BEFORERESULT = '" & sLastResult & "', " & _
                   "        INTERVAL = " & sInterval & " " & _
                   " WHERE LABDATE = '" & sLabDate & "' " & _
                   "   AND PARTGBN = '" & sPartGbn & "' " & _
                   "   AND LABSEQ  = '" & sLabSeq & "' " & _
                   "   AND SPECIMENCD = '" & sSpcCd & "' " & _
                   "   AND TESTITEMSEQ = '" & sTestSeq & "' " & _
                   "   AND SUBMCD = '" & sSubMcd & "' "
                   
            AdoConn.Execute sSql
    
            AdoRecord.MoveNext
        Next iCnt
        
    End If
    AdoRecord.Close
   
    Tran_Delta = "SUCCESS"
Exit Function
ErrHandler:
    MsgBox "GetItem(접수 DLL) - " & Err.Description
    Tran_Delta = "ERROR"
    
End Function
