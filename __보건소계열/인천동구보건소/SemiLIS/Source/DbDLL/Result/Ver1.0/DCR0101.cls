VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCR0101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarTotLABDATE As String '지역 복사
Private mvarTotPARTGBN As String '지역 복사
Private mvarTotLABSEQ As String '지역 복사
Private mvarCurItemCnt As String '지역 복사
Private mvarCurItemCnt2 As String '지역 복사
Private mvarTot이름 As String '지역 복사
Private mvarTot나이 As String '지역 복사
Private mvarTot성별 As String '지역 복사
Private mvarTot진료과 As String '지역 복사
Private mvarTot병실 As String '지역 복사
Private mvarTot접수구분 As String '지역 복사
Private mvarTot응급 As String '지역 복사
Private mvarTot특진 As String '지역 복사
Private mvarTot의사 As String '지역 복사
Private mvarTot소견 As String '지역 복사
Private mvarTot결과일 As String '지역 복사
Private mvarTot결과시간 As String '지역 복사
Private mvarTot검체코드 As String '지역 복사
Private mvarTot등록번호 As String '지역 복사
Private mvarTot진료과명 As String '지역 복사
Private mvarTot검체약명 As String '지역 복사
Private mvarTot검체원명 As String '지역 복사
Private mvarTot검사코드 As String '지역 복사
Private mvarTot검사약명 As String '지역 복사
Private mvarTot검사결과 As String '지역 복사
Private mvarTot참고치 As String '지역 복사
Private mvarTotPANIC As String '지역 복사
Private mvarTotDELTA As String '지역 복사
Private mvarTot소견코드 As String '지역 복사
Private mvarTot소견내용 As String '지역 복사
Private mvarTotREFGBN As String '지역 복사
Private mvarTotPANICGBN As String '지역 복사
Private mvarTotDELTAGBN As String '지역 복사
Private mvarTotDELTAVAL As String '지역 복사
Private mvarTotPANJUNGGBN As String '지역 복사
Private mvarTotPANICLOW As String '지역 복사
Private mvarTotPANICHIGH As String '지역 복사
Private mvarTotREFLOW As String '지역 복사
Private mvarTotREFHIGH As String '지역 복사
Private mvarTotUPPERLIMIT As String '지역 복사
Private mvarTotLOWERLIMIT As String '지역 복사
Private mvarTotGRAYUPPER As String '지역 복사
Private mvarTotGRAYLOWER As String '지역 복사
Private mvarTotOTHERFLAG As String '지역 복사
Private mvarTotREFLETTER As String '지역 복사
Private mvarTot서브코드 As String '지역 복사
Private mvarTotUNIT As String '지역 복사
Private mvarTot배경색 As String '지역 복사
Private mvarTot전경색 As String '지역 복사
Private mvarTot연산기준 As String '지역 복사
Private mvarTot이상기준 As String '지역 복사
Private mvarTotPARTCD As String '지역 복사
'속성값을 가질 지역 변수
Private mvarTotSTATUSCD As String '지역 복사
'속성값을 가질 지역 변수
Private mvarTot검사코드F As String '지역 복사
Private mvarTot서브코드F As String '지역 복사
'속성값을 가질 지역 변수
Private mvarCurItemCnt3 As String '지역 복사
'속성값을 가질 지역 변수
Private mvarTotFLAGYN As String '지역 복사
'속성값을 가질 지역 변수
Private mvarTot이전결과 As String '지역 복사
'DB연결(ADO)
Dim AdoConn     As New ADODB.Connection
'속성값을 가질 지역 변수
Private mvarTot결과완료 As String '지역 복사

Function Add_Flag(ByVal sLabInfo As String, ByVal sSpc As String, ByVal sFlag As String, ByVal iFlagCnt As Integer) As String
    Dim sSql        As String
    Dim i%, sExamCd$, sSubCd$, sFlags$, sOneFlag$
    Dim sLabDate$, sPartCd$, sPartGbn$, sLabseq$
    
    On Error GoTo ErrHandler
    
    sLabDate = Left(sLabInfo, 8)
    sPartCd = Mid(sLabInfo, 9, 1)
    sPartGbn = Mid(sLabInfo, 10, 2)
    sLabseq = Right(sLabInfo, 5)
    
    For i = 1 To iFlagCnt
        sExamCd = GetByOne(sFlag, sFlag)
        sSubCd = GetByOne(sFlag, sFlag)
        sFlags = GetByOne(sFlag, sFlag) & " "
        While (sFlags <> "")
            sOneFlag = GetBYSpc(sFlags, sFlags)
            
            If Trim$(sOneFlag) = "" Then
            Else
                sSql = " INSERT INTO FLAG (LABDATE,    PARTCD,      PARTGBN, LABSEQ, " & _
                       "                   SPECIMENCD, TESTITEMSEQ, SUBMCD,  STATUSCD ) " & _
                       " VALUES ('" & sLabDate & "','" & sPartCd & "','" & sPartGbn & "','" & sLabseq & "'," & _
                       "         '" & sSpc & "','" & sExamCd & "','" & sSubCd & "','" & sOneFlag & "') "
            
                AdoConn.Execute sSql
            End If
        Wend
    Next
    
    Add_Flag = "SUCCESS"
Exit Function
ErrHandler:
    Add_Flag = Str(Err.Number)
    MsgBox "Add_Flag(AR01.DLL) - " & Err.Description
End Function

Function Del_Flag(ByVal sLabInfo As String, ByVal sSpc As String, ByVal sFlag As String, ByVal iFlagCnt As Integer) As String
    Dim sSql        As String
    Dim i%, sExamCd$, sSubCd$, sFlags$
    Dim sLabDate$, sPartCd$, sPartGbn$, sLabseq$
    
    On Error GoTo ErrHandler
    
    sLabDate = Left(sLabInfo, 8)
    sPartCd = Mid(sLabInfo, 9, 1)
    sPartGbn = Mid(sLabInfo, 10, 2)
    sLabseq = Right(sLabInfo, 5)
    
    For i = 1 To iFlagCnt
        sExamCd = GetByOne(sFlag, sFlag)
        sSubCd = GetByOne(sFlag, sFlag)
        sFlags = GetByOne(sFlag, sFlag) & " "
        
        sSql = " DELETE FROM FLAG " & _
           "  WHERE LABDATE = '" & sLabDate & "' " & _
           "    AND PARTCD  = '" & sPartCd & "' " & _
           "    AND PARTGBN = '" & sPartGbn & "' " & _
           "    AND LABSEQ  = '" & sLabseq & "' " & _
           "    AND SPECIMENCD = '" & sSpc & "' " & _
           "    AND TESTITEMSEQ = '" & sExamCd & "' " & _
           "    AND SUBMCD = '" & sSubCd & "' "
    
        AdoConn.Execute sSql
    Next
    
    Del_Flag = "SUCCESS"
    
Exit Function
ErrHandler:
    Del_Flag = Str(Err.Number)
    MsgBox "Del_Flag(AR01.DLL) - " & Err.Description
End Function

Function Add_Abnormal(ByVal sLabInfo As String, ByVal sSpc As String, ByVal sAbnormal As String) As String
    Dim sSql        As String
    Dim i%, sExamCd$, sSubCd$, sRstVal$, sRef$, sPanic$, sDelta$, sRPD$
    Dim sLabDate$, sPartCd$, sPartGbn$, sLabseq$
    Dim iAbnormalCnt%, iRPD%
    
    On Error GoTo ErrHandler
    
    sLabDate = Left(sLabInfo, 8)
    sPartCd = Mid(sLabInfo, 9, 1)
    sPartGbn = Mid(sLabInfo, 10, 2)
    sLabseq = Right(sLabInfo, 5)
    iAbnormalCnt = CInt(GetByOne(sAbnormal, sAbnormal))
    
    For i = 1 To iAbnormalCnt
        sExamCd = GetByOne(sAbnormal, sAbnormal)
        sSubCd = GetByOne(sAbnormal, sAbnormal)
        sRstVal = GetByOne(sAbnormal, sAbnormal)
        sRef = GetByOne(sAbnormal, sAbnormal)
        sPanic = GetByOne(sAbnormal, sAbnormal)
        sDelta = GetByOne(sAbnormal, sAbnormal)
        
        iRPD = 0
        
        If sRef = "" Then
        ElseIf sRef = "L" Then
            iRPD = iRPD + 1
        ElseIf sRef = "H" Then
            iRPD = iRPD + 2
        ElseIf sRef = "Pos" Then
            iRPD = iRPD + 16
        Else
            iRPD = iRPD + 32
        End If
        
        If sPanic = "P" Then
            iRPD = iRPD + 4
        End If
        
        If sDelta = "D" Then
            iRPD = iRPD + 8
        End If
        
        sRPD = CStr(iRPD)
        
        If iRPD = 0 Then
        Else
            sSql = "INSERT INTO H_ABNORMAL ( " & _
                   "LABDATE, PARTGBN, LABSEQ, SPECIMENCD, TESTITEMSEQ, SUBMCD, " & _
                   "ABNORMALCD, RESULT ) " & _
                   "VALUES ( " & _
                   "'" & sLabDate & "', '" & sPartGbn & "', '" & sLabseq & "', " & _
                   "'" & sSpc & "', '" & sExamCd & "', '" & sSubCd & "', " & _
                   "'" & sRPD & "', '" & sRstVal & _
                   "' ) "
            
            AdoConn.Execute sSql
        End If
    Next
    
    Add_Abnormal = "SUCCESS"
Exit Function
ErrHandler:
    Add_Abnormal = Str(Err.Number)
    MsgBox "Add_Abnormal(AR01.DLL) - " & Err.Description
End Function

Function Del_Abnormal(ByVal sLabInfo As String, ByVal sSpc As String, ByVal sAbnormal As String) As String
    Dim sSql        As String
    Dim sLabDate$, sPartCd$, sPartGbn$, sLabseq$
    Dim iAbnormalCnt%, i%
    Dim sExamCd     As String
    Dim sSubCd      As String
    Dim sRstVal     As String
    Dim sRef        As String
    Dim sPanic      As String
    Dim sDelta      As String
    
    On Error GoTo ErrHandler
    
    sLabDate = Left(sLabInfo, 8)
    sPartCd = Mid(sLabInfo, 9, 1)
    sPartGbn = Mid(sLabInfo, 10, 2)
    sLabseq = Right(sLabInfo, 5)
    iAbnormalCnt = CInt(GetByOne(sAbnormal, sAbnormal))
    
    For i = 1 To iAbnormalCnt
        sExamCd = GetByOne(sAbnormal, sAbnormal)
        sSubCd = GetByOne(sAbnormal, sAbnormal)
        sRstVal = GetByOne(sAbnormal, sAbnormal)
        sRef = GetByOne(sAbnormal, sAbnormal)
        sPanic = GetByOne(sAbnormal, sAbnormal)
        sDelta = GetByOne(sAbnormal, sAbnormal)
        
        sSql = " DELETE FROM H_ABNORMAL " & _
               "  WHERE LABDATE = '" & sLabDate & "'" & _
               "    AND PARTGBN = '" & sPartGbn & "'" & _
               "    AND LABSEQ  = '" & sLabseq & "' " & _
               "    AND SPECIMENCD = '" & sSpc & "' " & _
               "    AND TESTITEMSEQ = '" & sExamCd & "' " & _
               "    AND SUBMCD = '" & sSubCd & "' "
           
        AdoConn.Execute sSql
    Next
    
    Del_Abnormal = "SUCCESS"
    
Exit Function
ErrHandler:
    Del_Abnormal = Str(Err.Number)
    MsgBox "Del_Abnormal(AR01.DLL) - " & Err.Description
End Function

Public Property Let Tot결과완료(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot결과완료 = 5
    mvarTot결과완료 = vData
End Property


Private Function GetBYSpc(ByVal tStr As String, sOriginal As String) As String
    Dim Pos%
    
    Pos = InStr(tStr, " ")
    
    If Pos = 0 Then
    Else
        GetBYSpc = Mid$(tStr, 1, Pos - 1)
        sOriginal = Mid$(sOriginal, Pos + 1, Len(sOriginal) - Pos)
    End If
End Function


Public Property Get Tot결과완료() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot결과완료
    Tot결과완료 = mvarTot결과완료
End Property



Public Property Let Tot이전결과(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot이전결과 = 5
    mvarTot이전결과 = vData
End Property


Public Property Get Tot이전결과() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot이전결과
    Tot이전결과 = mvarTot이전결과
End Property



Public Property Let TotFLAGYN(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotFLAGYN = 5
    mvarTotFLAGYN = vData
End Property


Public Property Get TotFLAGYN() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotFLAGYN
    TotFLAGYN = mvarTotFLAGYN
End Property




Public Property Get CurItemCnt3() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.CurItemCnt3
    CurItemCnt3 = mvarCurItemCnt3
End Property

Public Property Let Tot서브코드F(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot서브코드F = 5
    mvarTot서브코드F = vData
End Property


Public Property Get Tot서브코드F() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot서브코드F
    Tot서브코드F = mvarTot서브코드F
End Property



Public Property Let Tot검사코드F(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot검사코드F = 5
    mvarTot검사코드F = vData
End Property


Public Property Get Tot검사코드F() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot검사코드F
    Tot검사코드F = mvarTot검사코드F
End Property



Public Property Let TotSTATUSCD(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotSTATUSCD = 5
    mvarTotSTATUSCD = vData
End Property


Public Property Get TotSTATUSCD() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotSTATUSCD
    TotSTATUSCD = mvarTotSTATUSCD
End Property



Public Property Let TotPARTCD(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotPARTCD = 5
    mvarTotPARTCD = vData
End Property


Public Property Get TotPARTCD() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotPARTCD
    TotPARTCD = mvarTotPARTCD
End Property



Public Property Let Tot이상기준(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot이상기준 = 5
    mvarTot이상기준 = vData
End Property


Public Property Get Tot이상기준() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot이상기준
    Tot이상기준 = mvarTot이상기준
End Property



Public Property Let Tot연산기준(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot연산기준 = 5
    mvarTot연산기준 = vData
End Property


Public Property Get Tot연산기준() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot연산기준
    Tot연산기준 = mvarTot연산기준
End Property



Public Property Let Tot전경색(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot전경색 = 5
    mvarTot전경색 = vData
End Property


Public Property Get Tot전경색() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot전경색
    Tot전경색 = mvarTot전경색
End Property



Public Property Let Tot배경색(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot배경색 = 5
    mvarTot배경색 = vData
End Property


Public Property Get Tot배경색() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot배경색
    Tot배경색 = mvarTot배경색
End Property



Public Property Let TotUNIT(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotUNIT = 5
    mvarTotUNIT = vData
End Property


Public Property Get TotUNIT() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotUNIT
    TotUNIT = mvarTotUNIT
End Property



Public Property Let Tot서브코드(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot서브코드 = 5
    mvarTot서브코드 = vData
End Property


Public Property Get Tot서브코드() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot서브코드
    Tot서브코드 = mvarTot서브코드
End Property



Public Property Let TotREFLETTER(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotREFLETTER = 5
    mvarTotREFLETTER = vData
End Property


Public Property Get TotREFLETTER() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotREFLETTER
    TotREFLETTER = mvarTotREFLETTER
End Property



Public Property Let TotOTHERFLAG(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotOTHERFLAG = 5
    mvarTotOTHERFLAG = vData
End Property


Public Property Get TotOTHERFLAG() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotOTHERFLAG
    TotOTHERFLAG = mvarTotOTHERFLAG
End Property



Public Property Let TotGRAYLOWER(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotGRAYLOWER = 5
    mvarTotGRAYLOWER = vData
End Property


Public Property Get TotGRAYLOWER() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotGRAYLOWER
    TotGRAYLOWER = mvarTotGRAYLOWER
End Property



Public Property Let TotGRAYUPPER(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotGRAYUPPER = 5
    mvarTotGRAYUPPER = vData
End Property


Public Property Get TotGRAYUPPER() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotGRAYUPPER
    TotGRAYUPPER = mvarTotGRAYUPPER
End Property



Public Property Let TotLOWERLIMIT(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotLOWERLIMIT = 5
    mvarTotLOWERLIMIT = vData
End Property


Public Property Get TotLOWERLIMIT() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotLOWERLIMIT
    TotLOWERLIMIT = mvarTotLOWERLIMIT
End Property



Public Property Let TotUPPERLIMIT(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotUPPERLIMIT = 5
    mvarTotUPPERLIMIT = vData
End Property


Public Property Get TotUPPERLIMIT() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotUPPERLIMIT
    TotUPPERLIMIT = mvarTotUPPERLIMIT
End Property



Public Property Let TotREFHIGH(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotREFHIGH = 5
    mvarTotREFHIGH = vData
End Property


Public Property Get TotREFHIGH() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotREFHIGH
    TotREFHIGH = mvarTotREFHIGH
End Property



Public Property Let TotREFLOW(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotREFLOW = 5
    mvarTotREFLOW = vData
End Property


Public Property Get TotREFLOW() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotREFLOW
    TotREFLOW = mvarTotREFLOW
End Property



Public Property Let TotPANICHIGH(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotPANICHIGH = 5
    mvarTotPANICHIGH = vData
End Property


Public Property Get TotPANICHIGH() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotPANICHIGH
    TotPANICHIGH = mvarTotPANICHIGH
End Property



Public Property Let TotPANICLOW(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotPANICLOW = 5
    mvarTotPANICLOW = vData
End Property


Public Property Get TotPANICLOW() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotPANICLOW
    TotPANICLOW = mvarTotPANICLOW
End Property



Public Property Let TotPANJUNGGBN(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotPANJUNGGBN = 5
    mvarTotPANJUNGGBN = vData
End Property


Public Property Get TotPANJUNGGBN() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotPANJUNGGBN
    TotPANJUNGGBN = mvarTotPANJUNGGBN
End Property



Public Property Let TotDELTAVAL(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotDELTAVAL = 5
    mvarTotDELTAVAL = vData
End Property


Public Property Get TotDELTAVAL() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotDELTAVAL
    TotDELTAVAL = mvarTotDELTAVAL
End Property



Public Property Let TotDELTAGBN(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotDELTAGBN = 5
    mvarTotDELTAGBN = vData
End Property


Public Property Get TotDELTAGBN() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotDELTAGBN
    TotDELTAGBN = mvarTotDELTAGBN
End Property



Public Property Let TotPANICGBN(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotPANICGBN = 5
    mvarTotPANICGBN = vData
End Property


Public Property Get TotPANICGBN() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotPANICGBN
    TotPANICGBN = mvarTotPANICGBN
End Property



Public Property Let TotREFGBN(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotREFGBN = 5
    mvarTotREFGBN = vData
End Property


Public Property Get TotREFGBN() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotREFGBN
    TotREFGBN = mvarTotREFGBN
End Property



Public Property Get CurItemCnt2() As String
    CurItemCnt2 = mvarCurItemCnt2
End Property

Public Property Let Tot소견내용(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot소견내용 = 5
    mvarTot소견내용 = vData
End Property


Public Property Get Tot소견내용() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot소견내용
    Tot소견내용 = mvarTot소견내용
End Property



Public Property Let Tot소견코드(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot소견코드 = 5
    mvarTot소견코드 = vData
End Property


Public Property Get Tot소견코드() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot소견코드
    Tot소견코드 = mvarTot소견코드
End Property



Public Property Let TotDELTA(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotDELTA = 5
    mvarTotDELTA = vData
End Property


Public Property Get TotDELTA() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotDELTA
    TotDELTA = mvarTotDELTA
End Property



Public Property Let TotPANIC(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotPANIC = 5
    mvarTotPANIC = vData
End Property


Public Property Get TotPANIC() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotPANIC
    TotPANIC = mvarTotPANIC
End Property



Public Property Let Tot참고치(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot참고치 = 5
    mvarTot참고치 = vData
End Property


Public Property Get Tot참고치() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot참고치
    Tot참고치 = mvarTot참고치
End Property



Public Property Let Tot검사결과(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot검사결과 = 5
    mvarTot검사결과 = vData
End Property


Public Property Get Tot검사결과() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot검사결과
    Tot검사결과 = mvarTot검사결과
End Property



Public Property Let Tot검사약명(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot검사약명 = 5
    mvarTot검사약명 = vData
End Property


Public Property Get Tot검사약명() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot검사약명
    Tot검사약명 = mvarTot검사약명
End Property



Public Property Let Tot검사코드(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot검사코드 = 5
    mvarTot검사코드 = vData
End Property


Public Property Get Tot검사코드() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot검사코드
    Tot검사코드 = mvarTot검사코드
End Property



Public Property Let Tot검체원명(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot검체원명 = 5
    mvarTot검체원명 = vData
End Property


Public Property Get Tot검체원명() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot검체원명
    Tot검체원명 = mvarTot검체원명
End Property



Public Property Let Tot검체약명(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot검체약명 = 5
    mvarTot검체약명 = vData
End Property


Public Property Get Tot검체약명() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot검체약명
    Tot검체약명 = mvarTot검체약명
End Property



Public Property Let Tot진료과명(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot진료과명 = 5
    mvarTot진료과명 = vData
End Property


Public Property Get Tot진료과명() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot진료과명
    Tot진료과명 = mvarTot진료과명
End Property




Public Property Let Tot등록번호(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot등록번호 = 5
    mvarTot등록번호 = vData
End Property

Public Property Get Tot등록번호() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot등록번호
    Tot등록번호 = mvarTot등록번호
End Property

Public Property Let Tot검체코드(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot검체코드 = 5
    mvarTot검체코드 = vData
End Property

Public Property Get Tot검체코드() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot검체코드
    Tot검체코드 = mvarTot검체코드
End Property

Public Property Let Tot결과시간(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot결과시간 = 5
    mvarTot결과시간 = vData
End Property

Public Property Get Tot결과시간() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot결과시간
    Tot결과시간 = mvarTot결과시간
End Property

Public Property Let Tot결과일(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot결과일 = 5
    mvarTot결과일 = vData
End Property

Public Property Get Tot결과일() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot결과일
    Tot결과일 = mvarTot결과일
End Property

Public Property Let Tot소견(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot소견 = 5
    mvarTot소견 = vData
End Property

Public Property Get Tot소견() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot소견
    Tot소견 = mvarTot소견
End Property

Public Property Let Tot의사(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot의사 = 5
    mvarTot의사 = vData
End Property

Public Property Get Tot의사() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot의사
    Tot의사 = mvarTot의사
End Property

Public Property Let Tot특진(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot특진 = 5
    mvarTot특진 = vData
End Property

Public Property Get Tot특진() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot특진
    Tot특진 = mvarTot특진
End Property

Public Property Let Tot응급(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot응급 = 5
    mvarTot응급 = vData
End Property

Public Property Get Tot응급() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot응급
    Tot응급 = mvarTot응급
End Property

Public Property Let Tot접수구분(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot접수구분 = 5
    mvarTot접수구분 = vData
End Property

Public Property Get Tot접수구분() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot접수구분
    Tot접수구분 = mvarTot접수구분
End Property

Public Property Let Tot병실(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot병실 = 5
    mvarTot병실 = vData
End Property

Public Property Get Tot병실() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot병실
    Tot병실 = mvarTot병실
End Property

Public Property Let Tot진료과(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot진료과 = 5
    mvarTot진료과 = vData
End Property

Public Property Get Tot진료과() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot진료과
    Tot진료과 = mvarTot진료과
End Property

Public Property Let Tot성별(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot성별 = 5
    mvarTot성별 = vData
End Property

Public Property Get Tot성별() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot성별
    Tot성별 = mvarTot성별
End Property

Public Property Let Tot나이(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot나이 = 5
    mvarTot나이 = vData
End Property

Public Property Get Tot나이() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot나이
    Tot나이 = mvarTot나이
End Property

Public Property Let Tot이름(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.Tot이름 = 5
    mvarTot이름 = vData
End Property

Public Property Get Tot이름() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Tot이름
    Tot이름 = mvarTot이름
End Property

'속성값을 가질 지역 변수
Public Property Let TotLABSEQ(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotLABSEQ = 5
    mvarTotLABSEQ = vData
End Property

Public Property Get TotLABSEQ() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotLABSEQ
    TotLABSEQ = mvarTotLABSEQ
End Property

Public Property Let TotPARTGBN(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.TotPARTGBN = 5
    mvarTotPARTGBN = vData
End Property

Public Property Get TotPARTGBN() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.TotPARTGBN
    TotPARTGBN = mvarTotPARTGBN
End Property

Public Property Get CurItemCnt() As String
    CurItemCnt = mvarCurItemCnt
End Property

Public Property Let TotLABDATE(ByVal vData As String)
    mvarTotLABDATE = vData
End Property

Public Property Get TotLABDATE() As String
    TotLABDATE = mvarTotLABDATE
End Property

Public Sub Get_Pat(ByVal iMode As Integer, ByVal iMode2 As Integer, ByVal sField01 As String, ByVal sField02 As String, ByVal sField03 As String, ByVal sField04 As String, ByVal sPart As String)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
    
    sSql = "SELECT LABDATE,       PARTGBN,       LABSEQ,     SPECIMENCD, REGNO, " & _
           "       NAME,          AGE,           SEX,        DEPTCD,     WARD, " & _
           "       JUBSUGBN,      EMERGBN,       SPECIALORD, DOCTOR,     MEDICALREMARK, " & _
           "       RESULTENDDATE, RESULTENDTIME, RESULTENDYN " & _
           "  FROM H_JUBSU "

    If iMode2 = 1 Then
        sSql = sSql & "WHERE LABDATE = " & "'" & sField01 & "'"
    Else
        If iMode <> 1 Then
            sSql = sSql & "WHERE RESULTENDDATE = " & "'" & sField01 & "'"
        Else
            sSql = sSql & "WHERE LABDATE = " & "'" & sField01 & "'"
        End If
    End If
    
    If iMode = 1 Then '접수번호
        sSql = sSql & "  AND LABDATE = " & "'" & sField02 & "'" & _
                      "  AND PARTGBN = " & "'" & sField03 & "'" & _
                      "  AND LABSEQ  = " & "'" & sField04 & "'"
    ElseIf iMode = 2 Then '등록번호
        sSql = sSql & "  AND REGNO   = " & "'" & sField02 & "'"
    ElseIf iMode = 3 Then '완
'        sSql = sSql & "  AND RESULTENDYN = '1'"
        sSql = sSql & "  AND Not IsNull(RESULTENDDATE) AND Not IsNull(RESULTENDTIME) "
    ElseIf iMode = 4 Then '미완
'        sSql = sSql & "  AND RESULTENDYN = '0'"
        sSql = sSql & "  AND IsNull(RESULTENDDATE) AND IsNull(RESULTENDTIME) "
    End If
    
    sSql = sSql & " ORDER BY LABDATE, PARTGBN, LABSEQ "
    
    AdoConn.CursorType = adOpenStatic
    AdoConn.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoConn.RecordCount = 0 Then
        mvarCurItemCnt = 0
        AdoConn.Close
        Exit Sub
    Else
        i = AdoConn.RecordCount
        mvarCurItemCnt = i
        
        AdoConn.MoveFirst
        
        For j = 1 To i
            mvarTotLABDATE = mvarTotLABDATE & AdoConn("LABDATE") & "|"
            mvarTotPARTGBN = mvarTotPARTGBN & AdoConn("PARTGBN") & "|"
            mvarTotLABSEQ = mvarTotLABSEQ & AdoConn("LABSEQ") & "|"
            mvarTot검체코드 = mvarTot검체코드 & AdoConn("SPECIMENCD") & "|"
            mvarTot등록번호 = mvarTot등록번호 & AdoConn("REGNO") & "|"
            mvarTot이름 = mvarTot이름 & AdoConn("NAME") & "|"
            mvarTot나이 = mvarTot나이 & AdoConn("AGE") & "|"
            mvarTot성별 = mvarTot성별 & AdoConn("SEX") & "|"
            mvarTot진료과 = mvarTot진료과 & AdoConn("DEPTCD") & "|"
            mvarTot병실 = mvarTot병실 & AdoConn("WARD") & "|"
            mvarTot접수구분 = mvarTot접수구분 & AdoConn("JUBSUGBN") & "|"
            mvarTot응급 = mvarTot응급 & AdoConn("EMERGBN") & "|"
            mvarTot특진 = mvarTot특진 & AdoConn("SPECIALORD") & "|"
            mvarTot의사 = mvarTot의사 & AdoConn("DOCTOR") & "|"
            mvarTot소견 = mvarTot소견 & AdoConn("MEDICALREMARK") & "|"
            mvarTot결과일 = mvarTot결과일 & AdoConn("RESULTENDDATE") & "|"
            mvarTot결과시간 = mvarTot결과시간 & AdoConn("RESULTENDTIME") & "|"
            mvarTot결과완료 = mvarTot결과완료 & AdoConn("RESULTENDYN") & "|"
            
            AdoConn.MoveNext
        Next
    End If
    
    AdoConn.Close
    
    Exit Sub
ErrHandler:
    MsgBox "Get_Pat(AR01.DLL) - " & Err.Description
    AdoConn.Close
End Sub

Public Sub Get_진료과()
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
    
    sSql = "SELECT DEPTCD, DEPTNM FROM DEPT"
    
    AdoConn.CursorType = adOpenStatic
    AdoConn.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoConn.RecordCount = 0 Then
        mvarCurItemCnt = 0
        AdoConn.Close
        Exit Sub
    Else
        i = AdoConn.RecordCount
        mvarCurItemCnt = i
        
        AdoConn.MoveFirst
        
        For j = 1 To i
            mvarTot진료과 = mvarTot진료과 & AdoConn("DEPTCD") & "|"
            mvarTot진료과명 = mvarTot진료과명 & AdoConn("DEPTNM") & "|"
            AdoConn.MoveNext
        Next
    End If
    
    AdoConn.Close
    
    Exit Sub
ErrHandler:
    MsgBox "Get_진료과(AR01.DLL) - " & Err.Description
End Sub

Public Sub Get_검체()
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
        
    sSql = "SELECT SPECIMENCD, SPECIMENBRIEFNM, SPECIMENFULLNM FROM SPECIMEN "
    
    AdoConn.CursorType = adOpenStatic
    AdoConn.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoConn.RecordCount = 0 Then
        mvarCurItemCnt = 0
        AdoConn.Close
        Exit Sub
    Else
        i = AdoConn.RecordCount
        mvarCurItemCnt = i
        
        AdoConn.MoveFirst
        
        For j = 1 To i
            mvarTot검체코드 = mvarTot검체코드 & AdoConn("SPECIMENCD") & "|"
            mvarTot검체약명 = mvarTot검체약명 & AdoConn("SPECIMENBRIEFNM") & "|"
            mvarTot검체원명 = mvarTot검체원명 & AdoConn("SPECIMENFULLNM") & "|"
            AdoConn.MoveNext
        Next
    End If
    
    AdoConn.Close
    
    Exit Sub
ErrHandler:
    MsgBox "Get_검체(AR01.DLL) - " & Err.Description
End Sub

Public Sub Get_결과(ByVal sField01 As String, ByVal sField02 As String, ByVal sField03 As String, ByVal sField04 As String, ByVal sField05 As String, ByVal sField06 As String)
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
    Dim sBuf As String
    
    sBuf = fGetCurTestItemNmCfg
    
    '검사결과 및 참고치 가져오기
    If sField05 = "1" Then
        If sBuf = "T" Then
            sSql = "SELECT A.TESTITEMSEQ, A.SUBMCD,      B.TESTITEMNM,     A.RESULT,     A.REFMARK, " & _
                   "       A.PANICMARK,   A.DELTAMARK,   B.REFGBN,      B.PANICGBN,   B.DELTAGBN, " & _
                   "       B.DELTAVAL,    B.PANJUNGGBN,  B.PANICLOW,    B.PANICHIGH,  B.REFLOWM, " & _
                   "       B.REFHIGHM,    B.UPPERLIMITM, B.LOWERLIMITM, B.GRAYUPPERM, B.GRAYLOWERM, " & _
                   "       B.OTHERFLAG,   B.REFLETTER,   B.UNIT,        B.FLAGYN,     A.BEFORERESULT " & _
                   "  FROM H_RESULT A, TESTITEM B " & _
                   " WHERE A.LABDATE     = " & "'" & sField01 & "'" & _
                   "   AND A.PARTGBN     = " & "'" & sField02 & "'" & _
                   "   AND A.LABSEQ      = " & "'" & sField03 & "'" & _
                   "   AND A.SPECIMENCD  = " & "'" & sField04 & "'" & _
                   "   AND B.PARTCD      = " & "'" & sField06 & "'" & _
                   "   AND B.PARTGBN     = A.PARTGBN " & _
                   "   AND B.SPECIMENCD  = A.SPECIMENCD " & _
                   "   AND B.TESTITEMSEQ = A.TESTITEMSEQ " & _
                   "   AND B.SUBMCD      = A.SUBMCD " & _
                   " ORDER BY B.PARTGBN, B.SPECIMENCD, B.DISPLAYORDER "
        ElseIf sBuf = "P" Then
            sSql = "SELECT A.TESTITEMSEQ, A.SUBMCD,      B.PRINTNM,     A.RESULT,     A.REFMARK, " & _
                   "       A.PANICMARK,   A.DELTAMARK,   B.REFGBN,      B.PANICGBN,   B.DELTAGBN, " & _
                   "       B.DELTAVAL,    B.PANJUNGGBN,  B.PANICLOW,    B.PANICHIGH,  B.REFLOWM, " & _
                   "       B.REFHIGHM,    B.UPPERLIMITM, B.LOWERLIMITM, B.GRAYUPPERM, B.GRAYLOWERM, " & _
                   "       B.OTHERFLAG,   B.REFLETTER,   B.UNIT,        B.FLAGYN,     A.BEFORERESULT " & _
                   "  FROM H_RESULT A, TESTITEM B " & _
                   " WHERE A.LABDATE     = " & "'" & sField01 & "'" & _
                   "   AND A.PARTGBN     = " & "'" & sField02 & "'" & _
                   "   AND A.LABSEQ      = " & "'" & sField03 & "'" & _
                   "   AND A.SPECIMENCD  = " & "'" & sField04 & "'" & _
                   "   AND B.PARTCD      = " & "'" & sField06 & "'" & _
                   "   AND B.PARTGBN     = A.PARTGBN " & _
                   "   AND B.SPECIMENCD  = A.SPECIMENCD " & _
                   "   AND B.TESTITEMSEQ = A.TESTITEMSEQ " & _
                   "   AND B.SUBMCD      = A.SUBMCD " & _
                   " ORDER BY B.PARTGBN, B.SPECIMENCD, B.DISPLAYORDER "
        End If
    Else
        If sBuf = "T" Then
            sSql = "SELECT A.TESTITEMSEQ, A.SUBMCD,      B.TESTITEMNM,     A.RESULT,     A.REFMARK, " & _
                   "       A.PANICMARK,   A.DELTAMARK,   B.REFGBN,      B.PANICGBN,   B.DELTAGBN, " & _
                   "       B.DELTAVAL,    B.PANJUNGGBN,  B.PANICLOW,    B.PANICHIGH,  B.REFLOWF, " & _
                   "       B.REFHIGHF,    B.UPPERLIMITF, B.LOWERLIMITF, B.GRAYUPPERF, B.GRAYLOWERF, " & _
                   "       B.OTHERFLAG,   B.REFLETTER,   B.UNIT,        B.FLAGYN,     A.BEFORERESULT " & _
                   "  FROM H_RESULT A, TESTITEM B " & _
                   " WHERE A.LABDATE     = " & "'" & sField01 & "'" & _
                   "   AND A.PARTGBN     = " & "'" & sField02 & "'" & _
                   "   AND A.LABSEQ      = " & "'" & sField03 & "'" & _
                   "   AND A.SPECIMENCD  = " & "'" & sField04 & "'" & _
                   "   AND B.PARTCD      = " & "'" & sField06 & "'" & _
                   "   AND B.PARTGBN     = A.PARTGBN " & _
                   "   AND B.SPECIMENCD  = A.SPECIMENCD " & _
                   "   AND B.TESTITEMSEQ = A.TESTITEMSEQ " & _
                   "   AND B.SUBMCD      = A.SUBMCD " & _
                   " ORDER BY B.PARTGBN, B.SPECIMENCD, B.DISPLAYORDER "
        ElseIf sBuf = "P" Then
            sSql = "SELECT A.TESTITEMSEQ, A.SUBMCD,      B.PRINTNM,     A.RESULT,     A.REFMARK, " & _
                   "       A.PANICMARK,   A.DELTAMARK,   B.REFGBN,      B.PANICGBN,   B.DELTAGBN, " & _
                   "       B.DELTAVAL,    B.PANJUNGGBN,  B.PANICLOW,    B.PANICHIGH,  B.REFLOWF, " & _
                   "       B.REFHIGHF,    B.UPPERLIMITF, B.LOWERLIMITF, B.GRAYUPPERF, B.GRAYLOWERF, " & _
                   "       B.OTHERFLAG,   B.REFLETTER,   B.UNIT,        B.FLAGYN,     A.BEFORERESULT " & _
                   "  FROM H_RESULT A, TESTITEM B " & _
                   " WHERE A.LABDATE     = " & "'" & sField01 & "'" & _
                   "   AND A.PARTGBN     = " & "'" & sField02 & "'" & _
                   "   AND A.LABSEQ      = " & "'" & sField03 & "'" & _
                   "   AND A.SPECIMENCD  = " & "'" & sField04 & "'" & _
                   "   AND B.PARTCD      = " & "'" & sField06 & "'" & _
                   "   AND B.PARTGBN     = A.PARTGBN " & _
                   "   AND B.SPECIMENCD  = A.SPECIMENCD " & _
                   "   AND B.TESTITEMSEQ = A.TESTITEMSEQ " & _
                   "   AND B.SUBMCD      = A.SUBMCD " & _
                   " ORDER BY B.PARTGBN, B.SPECIMENCD, B.DISPLAYORDER "
        End If
    End If
    
    AdoConn.CursorType = adOpenStatic
    AdoConn.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoConn.RecordCount = 0 Then
        mvarCurItemCnt = 0
        mvarCurItemCnt2 = 0
        mvarCurItemCnt3 = 0
        AdoConn.Close
        Exit Sub
    Else
        i = AdoConn.RecordCount
        mvarCurItemCnt = i
        
        AdoConn.MoveFirst
        
        For j = 1 To i
            mvarTot검사코드 = mvarTot검사코드 & AdoConn("TESTITEMSEQ") & "|"
            mvarTot서브코드 = mvarTot서브코드 & AdoConn("SUBMCD") & "|"
            mvarTot검사약명 = mvarTot검사약명 & AdoConn(2) & "|"
            mvarTot검사결과 = mvarTot검사결과 & AdoConn("RESULT") & "|"
            mvarTot참고치 = mvarTot참고치 & AdoConn("REFMARK") & "|"
            mvarTotPANIC = mvarTotPANIC & AdoConn("PANICMARK") & "|"
            mvarTotDELTA = mvarTotDELTA & AdoConn("DELTAMARK") & "|"
            mvarTotREFGBN = mvarTotREFGBN & AdoConn("REFGBN") & "|"
            mvarTotPANICGBN = mvarTotPANICGBN & AdoConn("PANICGBN") & "|"
            mvarTotDELTAGBN = mvarTotDELTAGBN & AdoConn("DELTAGBN") & "|"
            mvarTotDELTAVAL = mvarTotDELTAVAL & AdoConn("DELTAVAL") & "|"
            mvarTotPANJUNGGBN = mvarTotPANJUNGGBN & AdoConn("PANJUNGGBN") & "|"
            mvarTotPANICLOW = mvarTotPANICLOW & AdoConn("PANICLOW") & "|"
            mvarTotPANICHIGH = mvarTotPANICHIGH & AdoConn("PANICHIGH") & "|"
            If sField05 = "1" Then
                mvarTotREFLOW = mvarTotREFLOW & AdoConn("REFLOWM") & "|"
                mvarTotREFHIGH = mvarTotREFHIGH & AdoConn("REFHIGHM") & "|"
                mvarTotUPPERLIMIT = mvarTotUPPERLIMIT & AdoConn("UPPERLIMITM") & "|"
                mvarTotLOWERLIMIT = mvarTotLOWERLIMIT & AdoConn("LOWERLIMITM") & "|"
                mvarTotGRAYUPPER = mvarTotGRAYUPPER & AdoConn("GRAYUPPERM") & "|"
                mvarTotGRAYLOWER = mvarTotGRAYLOWER & AdoConn("GRAYLOWERM") & "|"
            Else
                mvarTotREFLOW = mvarTotREFLOW & AdoConn("REFLOWF") & "|"
                mvarTotREFHIGH = mvarTotREFHIGH & AdoConn("REFHIGHF") & "|"
                mvarTotUPPERLIMIT = mvarTotUPPERLIMIT & AdoConn("UPPERLIMITF") & "|"
                mvarTotLOWERLIMIT = mvarTotLOWERLIMIT & AdoConn("LOWERLIMITF") & "|"
                mvarTotGRAYUPPER = mvarTotGRAYUPPER & AdoConn("GRAYUPPERF") & "|"
                mvarTotGRAYLOWER = mvarTotGRAYLOWER & AdoConn("GRAYLOWERF") & "|"
            End If
            mvarTotOTHERFLAG = mvarTotOTHERFLAG & AdoConn("OTHERFLAG") & "|"
            mvarTotREFLETTER = mvarTotREFLETTER & AdoConn("REFLETTER") & "|"
            mvarTotUNIT = mvarTotUNIT & AdoConn("UNIT") & "|"
            mvarTotFLAGYN = mvarTotFLAGYN & AdoConn("FLAGYN") & "|"
            mvarTot이전결과 = mvarTot이전결과 & AdoConn("BEFORERESULT") & "|"
            
            AdoConn.MoveNext
        Next
    End If
    
    AdoConn.Close
    
    '검사항목별 FLAG 가져오기
    sSql = "SELECT B.TESTITEMSEQ, B.SUBMCD, B.STATUSCD, B.BACKCOLOR, B.FORECOLOR" & _
           "  FROM FLAG A, STATUS B " & _
           " WHERE A.LABDATE     = " & "'" & sField01 & "'" & _
           "   AND A.PARTCD      = " & "'" & sField06 & "'" & _
           "   AND A.PARTGBN     = " & "'" & sField02 & "'" & _
           "   AND A.LABSEQ      = " & "'" & sField03 & "'" & _
           "   AND A.SPECIMENCD  = " & "'" & sField04 & "'" & _
           "   AND B.STATUSCD    = A.STATUSCD " & _
           "   AND B.PARTCD      = A.PARTCD " & _
           "   AND B.PARTGBN     = A.PARTGBN " & _
           "   AND B.SPECIMENCD  = A.SPECIMENCD " & _
           "   AND B.TESTITEMSEQ = A.TESTITEMSEQ " & _
           "   AND B.SUBMCD      = A.SUBMCD " & _
           " ORDER BY B.PARTCD, B.PARTGBN, B.SPECIMENCD, B.TESTITEMSEQ, B.SUBMCD "
    
    AdoConn.CursorType = adOpenStatic
    AdoConn.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoConn.RecordCount = 0 Then
        mvarCurItemCnt2 = 0
        AdoConn.Close
        GoTo NEXTSQL
    Else
        i = AdoConn.RecordCount
        mvarCurItemCnt2 = i
        
        AdoConn.MoveFirst

        For j = 1 To i
            mvarTot검사코드F = mvarTot검사코드F & AdoConn("TESTITEMSEQ") & "|"
            mvarTot서브코드F = mvarTot서브코드F & AdoConn("SUBMCD") & "|"
            mvarTotSTATUSCD = mvarTotSTATUSCD & AdoConn("STATUSCD") & "|"
            mvarTot배경색 = mvarTot배경색 & AdoConn("BACKCOLOR") & "|"
            mvarTot전경색 = mvarTot전경색 & AdoConn("FORECOLOR") & "|"
            AdoConn.MoveNext
        Next
    End If
    
    AdoConn.Close
    
NEXTSQL:

    '소견 가져오기
    sSql = "SELECT COMMENTCD, COMMENT" & _
           "  FROM COMMENT " & _
           " WHERE LABDATE = " & "'" & sField01 & "'" & _
           "   AND PARTCD  = " & "'" & sField06 & "'" & _
           "   AND PARTGBN = " & "'" & sField02 & "'" & _
           "   AND LABSEQ  = " & "'" & sField03 & "'" & _
           " ORDER BY COMMENTSEQ "
    
    AdoConn.CursorType = adOpenStatic
    AdoConn.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoConn.RecordCount = 0 Then
        mvarCurItemCnt3 = 0
        AdoConn.Close
        Exit Sub
    Else
        i = AdoConn.RecordCount
        mvarCurItemCnt3 = i
        
        AdoConn.MoveFirst
        
        For j = 1 To i
            mvarTot소견코드 = mvarTot소견코드 & AdoConn("COMMENTCD") & "|"
            mvarTot소견내용 = mvarTot소견내용 & AdoConn("COMMENT") & "|"
            AdoConn.MoveNext
        Next
    End If
    
    AdoConn.Close
    
    Exit Sub
ErrHandler:
    MsgBox "Get_결과(AR01.DLL) - " & Err.Description
End Sub
Public Sub Get_소견()
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
        
    sSql = "SELECT PARTCD, COMMENTCD, COMMENTNM " & _
           "  FROM COMMENTCD "
    
    AdoConn.CursorType = adOpenStatic
    AdoConn.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoConn.RecordCount = 0 Then
        mvarCurItemCnt = 0
        AdoConn.Close
        Exit Sub
    Else
        i = AdoConn.RecordCount
        mvarCurItemCnt = i
        
        AdoConn.MoveFirst
        
        For j = 1 To i
            mvarTotPARTCD = mvarTotPARTCD & AdoConn("PARTCD") & "|"
            mvarTot소견코드 = mvarTot소견코드 & AdoConn("COMMENTCD") & "|"
            mvarTot소견내용 = mvarTot소견내용 & AdoConn("COMMENTNM") & "|"
            AdoConn.MoveNext
        Next
    End If
    
    AdoConn.Close
    
    Exit Sub
ErrHandler:
    MsgBox "Get_소견(AR01.DLL) - " & Err.Description
End Sub

Public Sub Get_FLAG()
    On Error GoTo ErrHandler
    
    Dim AdoConn As New ADODB.Recordset
    Dim sSql As String
    Dim i As Integer
    Dim j As Integer
        
    sSql = "SELECT STATUSCD, PARTCD,      PARTGBN,      SPECIMENCD, TESTITEMSEQ, " & _
           "       SUBMCD,   CRITICALVAL, COMPOPERATOR, BACKCOLOR,  FORECOLOR " & _
           "  FROM STATUS " & _
           " ORDER BY PARTCD, PARTGBN, SPECIMENCD, TESTITEMSEQ, SUBMCD "
    
    AdoConn.CursorType = adOpenStatic
    AdoConn.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoConn.RecordCount = 0 Then
        mvarCurItemCnt = 0
        AdoConn.Close
        Exit Sub
    Else
        i = AdoConn.RecordCount
        mvarCurItemCnt = i
        
        AdoConn.MoveFirst
        
        For j = 1 To i
            mvarTotSTATUSCD = mvarTotSTATUSCD & AdoConn("STATUSCD") & "|"
            mvarTotPARTCD = mvarTotPARTCD & AdoConn("PARTCD") & "|"
            mvarTotPARTGBN = mvarTotPARTGBN & AdoConn("PARTGBN") & "|"
            mvarTot검체코드 = mvarTot검체코드 & AdoConn("SPECIMENCD") & "|"
            mvarTot검사코드 = mvarTot검사코드 & AdoConn("TESTITEMSEQ") & "|"
            mvarTot서브코드 = mvarTot서브코드 & AdoConn("SUBMCD") & "|"
            mvarTot이상기준 = mvarTot이상기준 & AdoConn("CRITICALVAL") & "|"
            mvarTot연산기준 = mvarTot연산기준 & AdoConn("COMPOPERATOR") & "|"
            mvarTot배경색 = mvarTot배경색 & AdoConn("BACKCOLOR") & "|"
            mvarTot전경색 = mvarTot전경색 & AdoConn("FORECOLOR") & "|"
            AdoConn.MoveNext
        Next
    End If
    
    AdoConn.Close
    
    Exit Sub
ErrHandler:
    MsgBox "Get_FLAG(AR01.DLL) - " & Err.Description
End Sub

Function Tran_Result(ByVal sLabInfo As String, ByVal sJubSu As String, ByVal sCmt As String, _
                      ByVal lCmtCnt As Long, ByVal sRst As String, ByVal sRstDate As String, _
                      ByVal sRstTime As String, ByVal sSpc As String, ByVal sFlag As String, _
                      ByVal iFlagCnt As Integer, ByVal sAbnormal As String, ByVal sReRunYN As String) As String
                      
    Dim RstEndDate$, RstDateTime$, MedRem$
    Dim RtnFlag$
    
    AdoConn.CursorLocation = adUseServer
    AdoConn.Open "" & fGetCurDSN & "", "", ""
    
    AdoConn.BeginTrans
' 접수정보
    If Len(sJubSu) <> 3 Then
        RtnFlag = Edit_Jubsu(sLabInfo, sJubSu, sRstDate, sRstTime)
    End If
    If IsNumeric(RtnFlag) = True Then
        AdoConn.RollbackTrans
        Exit Function
    End If

' 결과정보
    If sRst <> "" Then
        RtnFlag = Edit_Rst(sLabInfo, sRst, sRstDate, sRstTime, sSpc)
    End If
    If IsNumeric(RtnFlag) = True Then
        AdoConn.RollbackTrans
        Exit Function
    End If

'Abnormal정보
    If Len(sAbnormal) <> 0 Then
        RtnFlag = Del_Abnormal(sLabInfo, sSpc, sAbnormal)
        If IsNumeric(RtnFlag) = True Then
            AdoConn.RollbackTrans
            Exit Function
        End If
        
        RtnFlag = Add_Abnormal(sLabInfo, sSpc, sAbnormal)
        If IsNumeric(RtnFlag) = True Then
            AdoConn.RollbackTrans
            Exit Function
        End If
'''    ElseIf sAbnormal = "" And sReRunYN = "1" Then
'''        RtnFlag = Del_Abnormal(sLabInfo)
'''        If IsNumeric(RtnFlag) = True Then
'''            AdoConn.RollbackTrans
'''            Exit Function
'''        End If
    End If
    
' Comment정보
    If sCmt = "NULL" Then
        RtnFlag = Del_Cmt(sLabInfo)
        If IsNumeric(RtnFlag) = True Then
            AdoConn.RollbackTrans
            Exit Function
        End If
    Else
        If Len(sCmt) <> 0 Then
            RtnFlag = Add_Cmt(sLabInfo, sCmt, lCmtCnt)
            
            If IsNumeric(RtnFlag) = True Then
                AdoConn.RollbackTrans
                Exit Function
            End If
        End If
    End If
    
' FLAG정보
    If Len(sFlag) <> 0 Then
        RtnFlag = Del_Flag(sLabInfo, sSpc, sFlag, iFlagCnt)
        If IsNumeric(RtnFlag) = True Then
            AdoConn.RollbackTrans
            Exit Function
        End If
        
        RtnFlag = Add_Flag(sLabInfo, sSpc, sFlag, iFlagCnt)
        If IsNumeric(RtnFlag) = True Then
            AdoConn.RollbackTrans
            Exit Function
        End If
    End If
    
    AdoConn.CommitTrans
    
    AdoConn.Close
    
    Tran_Result = "SUCCESS"
    
End Function

Function Add_Cmt(ByVal sLabInfo As String, ByVal sCmt As String, ByVal lCmtCnt As Long) As String
    Dim sSql        As String
    Dim i%, iMaxCnt%, sCmtCd$, sComment$, sCnt$
    Dim sLabDate$, sPartCd$, sPartGbn$, sLabseq$
    Dim AdoRecord As New ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    sLabDate = Left(sLabInfo, 8)
    sPartCd = Mid(sLabInfo, 9, 1)
    sPartGbn = Mid(sLabInfo, 10, 2)
    sLabseq = Right(sLabInfo, 5)
    
    sSql = "SELECT TOP 1 COMMENTSEQ " & _
           "  FROM COMMENT " & _
           "  WHERE LABDATE = '" & sLabDate & "' " & _
           "    AND PARTCD = '" & sPartCd & "' " & _
           "    AND PARTGBN = '" & sPartGbn & "' " & _
           "    AND LABSEQ = '" & sLabseq & "' " & _
           "  ORDER BY COMMENTSEQ DESC "
    
    AdoRecord.CursorType = adOpenStatic
    AdoRecord.Open sSql, "DSN=" & fGetCurDSN & ";UID=;PWD=;"
    
    If AdoRecord.RecordCount = 0 Then
        iMaxCnt = 0
    Else
        iMaxCnt = CInt(AdoRecord(0))
    End If
    
    AdoRecord.Close
    
    For i = 1 To lCmtCnt
        sCmtCd = GetByOne(sCmt, sCmt)
        sComment = GetByOne(sCmt, sCmt)
        sCnt = Format$(i + iMaxCnt, "000")
        
        sSql = " INSERT INTO COMMENT (LABDATE,    PARTCD,    PARTGBN, LABSEQ, " & _
               "                      COMMENTSEQ, COMMENTCD, COMMENT        ) " & _
               " VALUES ('" & sLabDate & "', '" & sPartCd & "', '" & sPartGbn & "', '" & sLabseq & "', " & _
               "         '" & sCnt & "', '" & sCmtCd & "', '" & sComment & "') "
    
        AdoConn.Execute sSql
    Next
    
    Add_Cmt = "SUCCESS"
Exit Function
ErrHandler:
    Add_Cmt = Str(Err.Number)
    MsgBox "Add_Cmt(AR01.DLL) - " & Err.Description
End Function

Function Del_Cmt(ByVal sLabInfo As String, Optional ByVal sCmt As String, Optional ByVal iCmtCnt As Integer) As String
    Dim sSql        As String
    Dim sLabDate$, sPartCd$, sPartGbn$, sLabseq$
    Dim sCmtCd$, sComment$
    Dim i%
    
    On Error GoTo ErrHandler
    
    sLabDate = Left(sLabInfo, 8)
    sPartCd = Mid(sLabInfo, 9, 1)
    sPartGbn = Mid(sLabInfo, 10, 2)
    sLabseq = Right(sLabInfo, 5)
    
    If sCmt = "" Then
        sSql = " DELETE FROM COMMENT " & _
               "  WHERE LABDATE = '" & sLabDate & "'" & _
               "    AND PARTCD  = '" & sPartCd & "'" & _
               "    AND PARTGBN = '" & sPartGbn & "'" & _
               "    AND LABSEQ  = '" & sLabseq & "'"
               
        AdoConn.Execute sSql
    Else
        For i = 1 To iCmtCnt
            sCmtCd = GetByOne(sCmt, sCmt)
            sComment = GetByOne(sCmt, sCmt)
            
            sSql = " DELETE FROM COMMENT " & _
                   "  WHERE LABDATE = '" & sLabDate & "'" & _
                   "    AND PARTCD  = '" & sPartCd & "'" & _
                   "    AND PARTGBN = '" & sPartGbn & "'" & _
                   "    AND LABSEQ  = '" & sLabseq & "'" & _
                   "    AND COMMENTCD  = '" & sCmtCd & "'" & _
                   "    AND COMMENT  = '" & sComment & "' "
                   
            AdoConn.Execute sSql
        Next
    End If
    
    Del_Cmt = "SUCCESS"
    
Exit Function
ErrHandler:
    Del_Cmt = Str(Err.Number)
    MsgBox "Del_Cmt(AR01.DLL) - " & Err.Description
End Function

Function Del_CmtByOne(ByVal sLabInfo As String, ByVal sCmt As String, ByVal iCmtCnt As Integer) As String
    Dim sSql        As String
    Dim sLabDate$, sPartCd$, sPartGbn$, sLabseq$
    Dim sCmtCd$, sComment$
    Dim i%
    Dim iAdoConn As New ADODB.Connection
    
    On Error GoTo ErrHandler
    
    sLabDate = Left(sLabInfo, 8)
    sPartCd = Mid(sLabInfo, 9, 1)
    sPartGbn = Mid(sLabInfo, 10, 2)
    sLabseq = Right(sLabInfo, 5)
    
    iAdoConn.Open "" & fGetCurDSN & "", "", ""
    
    For i = 1 To iCmtCnt
        sCmtCd = GetByOne(sCmt, sCmt)
        sComment = GetByOne(sCmt, sCmt)
        
        sSql = " DELETE FROM COMMENT " & _
               "  WHERE LABDATE = '" & sLabDate & "'" & _
               "    AND PARTCD  = '" & sPartCd & "'" & _
               "    AND PARTGBN = '" & sPartGbn & "'" & _
               "    AND LABSEQ  = '" & sLabseq & "'" & _
               "    AND COMMENTCD  = '" & sCmtCd & "'" & _
               "    AND COMMENT  = '" & sComment & "' "
               
        iAdoConn.Execute sSql
    Next
    
    Del_CmtByOne = "SUCCESS"
    
    iAdoConn.Close
    
    Exit Function
    
ErrHandler:
    iAdoConn.Close
    Del_CmtByOne = Str(Err.Number)
    MsgBox "Del_CmtByOne(AR01.DLL) - " & Err.Description
End Function

Function Edit_Jubsu(ByVal sLabInfo As String, ByVal sJubSu As String, ByVal sDate As String, ByVal sTime As String) As String

    Dim sMedRem$, sLabDate$, sPartCd$, sPartGbn$, sLabseq$, sRstEdit$, sReRun$, sSlip$
    Dim sSql        As String
    Dim iMode%
    
    On Error GoTo ErrHandler
    
    sLabDate = Left(sLabInfo, 8)
    sPartCd = Mid(sLabInfo, 9, 1)
    sPartGbn = Mid(sLabInfo, 10, 2)
    sLabseq = Right(sLabInfo, 5)
    sMedRem = GetByOne(sJubSu, sJubSu)
    sRstEdit = GetByOne(sJubSu, sJubSu)
    sReRun = GetByOne(sJubSu, sJubSu)
    
    iMode = 0
    
    sSql = " UPDATE H_JUBSU SET"
    If sRstEdit = "Y" Then
        sSql = sSql & " RESULTENDDATE = '" & sDate & "', " & _
                      " RESULTENDTIME = '" & sTime & "', " & _
                      " RESULTENDYN   = '1' "
        iMode = 1
    ElseIf sRstEdit = "N" Then
        sSql = sSql & " RESULTENDDATE = '', " & _
                      " RESULTENDTIME = '', " & _
                      " RESULTENDYN   = '0' "
        iMode = 1
    End If
    
    If sMedRem = "" Then
    ElseIf sMedRem = Chr(13) Then
        If iMode = 0 Then
            sSql = sSql & " MEDICALREMARK = '' "
        Else
            sSql = sSql & ", MEDICALREMARK = '' "
        End If
        
        iMode = 1
    ElseIf sMedRem <> "" Then
        If iMode = 0 Then
            sSql = sSql & " MEDICALREMARK = '" & sMedRem & "' "
        Else
            sSql = sSql & ", MEDICALREMARK = '" & sMedRem & "' "
        End If
        
        iMode = 1
    End If
    
    If sReRun = "" Then
    Else
        If iMode = 0 Then
            sSql = sSql & " RERUNYN = '1'"
        Else
            sSql = sSql & ", RERUNYN = '1'"
        End If
    End If
    
    sSql = sSql & "  WHERE LABDATE       = '" & sLabDate & "'" & _
                  "    AND PARTGBN       = '" & sPartGbn & "'" & _
                  "    AND LABSEQ        = '" & sLabseq & "'"

    AdoConn.Execute sSql
    
    Edit_Jubsu = "SUCCESS"
Exit Function
ErrHandler:
    Edit_Jubsu = Str(Err.Number)
    MsgBox "Edit_Jubsu(AR01.DLL) - " & Err.Description
End Function

Function Edit_Rst(ByVal sLabInfo As String, ByVal sRst As String, ByVal sDate As String, _
                    ByVal sTime As String, ByVal sSpc As String) As String
    Dim i%
    Dim sLabDate$, sPartCd$, sPartGbn$, sLabseq$, sRstCnt$
    Dim sRstVal$, sDelta$, sRef$, sPanic$, sExamCd$, sSubCd$
    Dim sSql$
    
    On Error GoTo ErrHandler
    
    sLabDate = Left(sLabInfo, 8)
    sPartCd = Mid(sLabInfo, 9, 1)
    sPartGbn = Mid(sLabInfo, 10, 2)
    sLabseq = Right(sLabInfo, 5)
    sRstCnt = GetByOne(sRst, sRst)
    
    For i = 1 To Val(sRstCnt)
        sExamCd = GetByOne(sRst, sRst)
        sSubCd = GetByOne(sRst, sRst)
        sRstVal = GetByOne(sRst, sRst)
        sRef = GetByOne(sRst, sRst)
        sPanic = GetByOne(sRst, sRst)
        sDelta = GetByOne(sRst, sRst)
        
        sSql = " UPDATE H_RESULT SET" & _
               "        RESULTDATE  = '" & sDate & "'," & _
               "        RESULTTIME  = '" & sTime & "'," & _
               "        RESULT      = '" & sRstVal & "'," & _
               "        DELTAMARK   = '" & sDelta & "'," & _
               "        REFMARK     = '" & sRef & "'," & _
               "        PANICMARK   = '" & sPanic & "'" & _
               "  WHERE LABDATE     = '" & sLabDate & "'" & _
               "    AND PARTGBN     = '" & sPartGbn & "'" & _
               "    AND LABSEQ      = '" & sLabseq & "'" & _
               "    AND SPECIMENCD  = '" & sSpc & "'" & _
               "    AND TESTITEMSEQ = '" & sExamCd & "'" & _
               "    AND SUBMCD      = '" & sSubCd & "'"
    
        AdoConn.Execute sSql
    Next
    
    Edit_Rst = "SUCCESS"
Exit Function
ErrHandler:
    Edit_Rst = Str(Err.Number)
    MsgBox "Edit_Rst(AR01.DLL) - " & Err.Description
End Function

Function Edit_Delta(sLabInfo As String, sRst As String, sDate As String, sTime As String, sSpc As String) As String
    Dim i%
    Dim sLabDate$, sPartCd$, sPartGbn$, sLabseq$, sRstCnt$
    Dim sRstVal$, sDelta$, sRef$, sPanic$, sExamCd$, sSubCd$
    Dim sSql$
    
    On Error GoTo ErrHandler
    
    sLabDate = Left(sLabInfo, 8)
    sPartCd = Mid(sLabInfo, 9, 1)
    sPartGbn = Mid(sLabInfo, 10, 2)
    sLabseq = Right(sLabInfo, 5)
    sRstCnt = GetByOne(sRst, sRst)
    
    For i = 1 To Val(sRstCnt)
        sExamCd = GetByOne(sRst, sRst)
        sSubCd = GetByOne(sRst, sRst)
        sRstVal = GetByOne(sRst, sRst)
        sRef = GetByOne(sRst, sRst)
        sPanic = GetByOne(sRst, sRst)
        sDelta = GetByOne(sRst, sRst)
        
        sSql = " UPDATE H_RESULT SET" & _
               "        RESULTDATE  = '" & sDate & "'," & _
               "        RESULTTIME  = '" & sTime & "'," & _
               "        RESULT      = '" & sRstVal & "'," & _
               "        DELTAMARK   = '" & sDelta & "'," & _
               "        REFMARK     = '" & sRef & "'," & _
               "        PANICMARK   = '" & sPanic & "'" & _
               "  WHERE LABDATE     = '" & sLabDate & "'" & _
               "    AND PARTGBN     = '" & sPartGbn & "'" & _
               "    AND LABSEQ      = '" & sLabseq & "'" & _
               "    AND SPECIMENCD  = '" & sSpc & "'" & _
               "    AND TESTITEMSEQ = '" & sExamCd & "'" & _
               "    AND SUBMCD      = '" & sSubCd & "'"
    
        AdoConn.Execute sSql
    Next
    
    Edit_Delta = "SUCCESS"
Exit Function
ErrHandler:
    Edit_Delta = Str(Err.Number)
    MsgBox "Edit_Delta(AR01.DLL) - " & Err.Description
End Function

