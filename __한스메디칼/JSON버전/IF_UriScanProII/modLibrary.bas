Attribute VB_Name = "modLibrary"
Option Explicit

Private Declare Function GetIpAddrTable_API Lib "IpHlpApi" Alias "GetIpAddrTable" (pIPAddrTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long
 
Public Function GetIpAddrTable()
   Dim Buf(0 To 511) As Byte
   Dim BufSize As Long: BufSize = UBound(Buf) + 1
   Dim rc As Long
   rc = GetIpAddrTable_API(Buf(0), BufSize, 1)
   If rc <> 0 Then Err.Raise vbObjectError, , "GetIpAddrTable failed with return value " & rc
   Dim NrOfEntries As Integer: NrOfEntries = Buf(1) * 256 + Buf(0)
   If NrOfEntries = 0 Then GetIpAddrTable = Array(): Exit Function
   ReDim IpAddrs(0 To NrOfEntries - 1) As String
   Dim i As Integer
   For i = 0 To NrOfEntries - 1
      Dim j As Integer, s As String: s = ""
      For j = 0 To 3: s = s & IIf(j > 0, ".", "") & Buf(4 + i * 24 + j): Next
      IpAddrs(i) = s
      Next
   GetIpAddrTable = IpAddrs
End Function
   
'-- 지금날짜와 검사일자 비교한다
Public Function DateCompare(ByVal FDate As String) As String
    
    DateCompare = FDate
    If FDate <> Format(Now, "yyyymmdd") Then
        DateCompare = Format(Now, "yyyymmdd")
    End If
    
End Function

'-1 은 해당 값이 없는 항목임
' C3815 코드 확인할것
'-- CRR 결과판정
Public Function getCRRValue(ByVal pTestCd As String, ByVal pResult As String) As String
    Dim strCRR      As String
    Dim dblLow      As Double
    Dim dblHigh     As Double
    
    dblLow = -1
    dblHigh = -1
    strCRR = ""
    
    getCRRValue = pResult
    
    If Not IsNumeric(pResult) Then
        Exit Function
    End If
    
    Select Case pTestCd
        '생화학
        Case "B2570":       dblLow = 2.6:       dblHigh = 6000      'AST
        Case "B2580":       dblLow = 2.2:       dblHigh = 6600      'ALT
        Case "C3711":       dblLow = 1.3:       dblHigh = 2100      'GLU
        Case "C2411":       dblLow = 0.7:       dblHigh = 800       'TCHO
        Case "C3720":       dblLow = 0.03:      dblHigh = 40        'TBIL
        Case "C3721":       dblLow = 0.04:      dblHigh = 28.6      'DBIL
        Case "C2200":       dblLow = 0.1:       dblHigh = 15.8      'TP
        Case "C2210":       dblLow = 0.2:       dblHigh = 9.9       'ALB
        Case "C2602":       dblLow = 0:         dblHigh = 6600      'ALP
        Case "C3730":       dblLow = 2.25:      dblHigh = 200       'BUN
        Case "C3750":       dblLow = 0.1:       dblHigh = 60        'CREA
        Case "C2443":       dblLow = 1.1:       dblHigh = 2000      'TG
        Case "B2710":       dblLow = 5.4:       dblHigh = 4200      'LDH
        Case "C3038":       dblLow = 3.3:       dblHigh = 7200      'rGTP
        Case "C3780":       dblLow = 0.14:      dblHigh = 80        'UA
        Case "C2243":       dblLow = 0.01:      dblHigh = 62.4      'CRP
        Case "C4903":       dblLow = 3:         dblHigh = 900       'RF(RA)
        Case "C2420":       dblLow = 2:         dblHigh = 230       'HDL-C
        Case "C2430":       dblLow = 1:         dblHigh = 1000      'LDL-C
        Case "C3795":       dblLow = 0.35:      dblHigh = 26.8      'CA
        Case "C3794":       dblLow = 0.1:       dblHigh = 35        'P
        Case "B2630":       dblLow = 6:         dblHigh = 7800      'CK(CPK)
        Case "C2490":       dblLow = 4:         dblHigh = 1000      'FE
        Case "B2611":       dblLow = 1.8:       dblHigh = 4500      'AMY
        Case "C3870":       dblLow = 5.87:      dblHigh = 587       'NH3(AMM)
        Case "C3812N1":     dblLow = 0:         dblHigh = 50        'TCO2
        Case "C2200N2":     dblLow = 0:         dblHigh = 400       'Micro TP
        Case "C2302N6":     dblLow = 0.03:      dblHigh = -1        'Micro ALB
        Case "C3791":       dblLow = 100:       dblHigh = 207       'NA
        Case "C3792":       dblLow = 1:         dblHigh = 109.6     'K
        Case "C3793":       dblLow = 15:        dblHigh = 200       'Cl
        Case "C3825":       dblLow = 3.5:       dblHigh = 18.5      'HBA1C
        Case "C3815N1":     dblLow = 6.001:     dblHigh = 8         'PH
        Case "C3815N2":     dblLow = 5:         dblHigh = 250       'PCO2
        Case "C3815N3":     dblLow = 0:         dblHigh = 749       'PO2
        Case "C3720N1":     dblLow = 0:         dblHigh = 30        'BIL
        Case "C3800":       dblLow = 0:         dblHigh = 2000      'OSMO
        Case "C3800-1":     dblLow = 0:         dblHigh = 2000      'OSMO
        Case "C3797N2":     dblLow = 0.1:       dblHigh = -1        'MG
        Case "C2621N1":     dblLow = 3:         dblHigh = -1        'LIPASE
        Case "C3796N1":     dblLow = 0.2:       dblHigh = -1        'IoN-O2
    
        '면역
        Case "C3290":      dblLow = 0.1:        dblHigh = 8         'T3
        Case "C3340":      dblLow = 0.1:        dblHigh = 12        'FT4
        Case "C3360":      dblLow = 0.01:       dblHigh = 150       'TSH
        Case "C2520":      dblLow = 0.5:        dblHigh = 1650      'FERR
        Case "C4212":      dblLow = 1.3:        dblHigh = 200000    'AFP
        Case "C4220":      dblLow = 0.5:        dblHigh = 10000     'CEA
        Case "C4280":      dblLow = 0.01:       dblHigh = 100       'PSA
        Case "C3520":      dblLow = 2:          dblHigh = 200000    'ThCG
        Case "C4230":      dblLow = 1.2:        dblHigh = 50000     'CA199
        Case "C4240":      dblLow = 2:          dblHigh = 600       'CA125
        Case "C4802":      dblLow = 0.1:        dblHigh = 1000      'HBSAG
        Case "C4812":      dblLow = 1:          dblHigh = 1000      'HBSAB
        Case "C4861-1":    dblLow = 0:          dblHigh = 100       'HAV T
        Case "C4862":      dblLow = 0.02:       dblHigh = 7         'HAV M
        Case "C4872":      dblLow = 0:          dblHigh = 11        'HCV
        Case "C4872-1":    dblLow = 0:          dblHigh = 11        'HCV
        Case "C4872-2":    dblLow = 0:          dblHigh = 11        'HCV
        Case "C4712":      dblLow = 0.05:       dblHigh = 50        'HIV
        Case "C3942-1":    dblLow = 0.006:      dblHigh = 50        'TNI
        Case "B2640":      dblLow = 0.18:       dblHigh = 300       'CKMB
    
    End Select
    
'    If dblLow <> -1 Then
'        If dblLow > CDbl(pResult) Then
'            strCRR = "<" & Space(1) & dblLow
'        ElseIf dblHigh < CDbl(pResult) Then
'            strCRR = ">" & Space(1) & dblHigh
'        Else
'            strCRR = pResult
'        End If
'    Else
'        strCRR = pResult
'    End If
    
    If dblLow > CDbl(pResult) Then
        strCRR = "<" & Space(1) & dblLow
    Else
        If dblHigh = -1 Then
            strCRR = pResult
        Else
            If dblHigh < CDbl(pResult) Then
                strCRR = ">" & Space(1) & dblHigh
            Else
                strCRR = pResult
            End If
        End If
    End If
    
    getCRRValue = strCRR

End Function


Public Function SetText(ByRef vasTable As Object, ByVal SetStr As String, ByVal vasRow As Long, ByVal vasCol As Long) As Boolean
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    vasTable.Text = SetStr
End Function

Public Function GetText(ByRef vasTable As Object, ByVal vasRow As Long, ByVal vasCol As Long) As String
    If vasRow < 0 Or vasCol < 0 Then
        Exit Function
    End If
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    GetText = vasTable.Text
End Function

Public Function spdActiveCell(ByRef vasTable As Object, ByVal vasRow As Long, ByVal vasCol As Long) As Boolean
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    vasTable.Action = 0
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 문자열을 구분자를 이용해 구분해 지정한 위치의 문자열을 구함
'   인수 :
'       1.pText      : 구분자로 구성된 문자열
'       2.pPosiion   : 위치
'       3.pDelimiter : 구분자
'-----------------------------------------------------------------------------'
Public Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
                      ByVal pDelimiter As String) As String
    
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    'pPosition 인수가 1인 경우 For문 Skip
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    '해당 컬럼
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
    If intPos2 = 0 Then intPos2 = Len(pText) + 1
    
    mGetP = Mid$(pText, intPos1, intPos2 - intPos1)
    Exit Function
    
ReturnNull:
    mGetP = ""
End Function

'문장 양쪽에 Single quote 를 붙인다.
Public Function STS(ByVal strStmt As String) As String
    Dim strTmp As String
    
    strTmp = Replace(strStmt, "'", "''")
    
    STS = "'" & strTmp & "'"
End Function

Public Function PedLeftStr(ByVal pData As String, ByVal pLen As Integer, ByVal pVal As Integer)
    Dim intLen  As Integer
    
    PedLeftStr = ""
    intLen = pLen - Len(pData)
    
    PedLeftStr = Space(intLen)
    PedLeftStr = Replace(PedLeftStr, " ", pVal)
    PedLeftStr = PedLeftStr & pData
    
End Function


Public Function PedRighttStr(ByVal pData As String, ByVal pLen As Integer, ByVal pVal As Integer)
    Dim intLen  As Integer
    
    PedRighttStr = ""
    intLen = pLen - Len(pData)
    
    PedRighttStr = Space(intLen)
    PedRighttStr = Replace(PedRighttStr, " ", pVal)
    PedRighttStr = pData & PedRighttStr
    
End Function


Public Sub SetRawData(argSQL As String)
    Dim FilNum
    Dim sFileName As String
    
    If gHOSP.LOQWRITE = "0" Then
        Exit Sub
    End If
    
    FilNum = FreeFile
    
    If Dir(App.PATH & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.PATH & "\Log")
    End If
    
    sFileName = gHOSP.MACHNM & "_" & Format(CDate(frmMain.dtpToday), "yyyy-mm-dd")
    
    Open App.PATH & "\Log\" & sFileName & ".txt" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
    
End Sub

Public Sub SetSQLData(ByVal strName As String, ByVal argSQL As String, Optional ByVal argMode As String)
    Dim FilNum
    Dim sFileName As String
    
    If gHOSP.LOQWRITE = "0" Then
        Exit Sub
    End If
    
    FilNum = FreeFile
        
    If Dir(App.PATH & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.PATH & "\Log")
    End If
    
    sFileName = gHOSP.MACHNM & "_" & Format(CDate(frmMain.dtpToday), "yyyy-mm-dd") & "_" & strName
    
    If argMode = "A" Then
        Open App.PATH & "\Log\" & sFileName & ".txt" For Append As FilNum
    Else
        Open App.PATH & "\Log\" & sFileName & ".txt" For Output As FilNum
    End If
    Print #FilNum, argSQL
    Close FilNum
    
End Sub

Public Sub DeleteRow(ByVal vasTable As Object, ByVal argRow1 As Integer, ByVal argRow2 As Integer)
    vasTable.Row = argRow1
    vasTable.Row2 = argRow2
    vasTable.Col = 1
    vasTable.Col2 = vasTable.MaxCols
    vasTable.BlockMode = True
    vasTable.Action = 5
    vasTable.BlockMode = False
End Sub

Public Sub Deletecol(ByVal vasTable As Object, ByVal argCol1 As Integer, ByVal argCol2 As Integer)
    vasTable.Row = 1
    vasTable.Row2 = vasTable.MaxRows
    vasTable.Col = argCol1
    vasTable.Col2 = argCol2
    vasTable.BlockMode = True
    vasTable.Action = 6
    vasTable.BlockMode = False
End Sub

Public Sub SetBackColor(asTable As vaSpread, ByVal asRow1 As Long, ByVal asRow2 As Long, ByVal asCol1 As Long, ByVal asCol2 As Long, asR As Variant, asG As Variant, asB As Variant)
    asTable.Row = asRow1
    asTable.Row2 = asRow2
    asTable.Col = asCol1
    asTable.Col2 = asCol2
    asTable.BlockMode = True
    asTable.BackColor = RGB(asR, asG, asB)
    asTable.BlockMode = False
End Sub

Public Sub SetForeColor(asTable As vaSpread, ByVal asRow1 As Long, ByVal asRow2 As Long, ByVal asCol1 As Long, ByVal asCol2 As Long, asR As Variant, asG As Variant, asB As Variant)
    asTable.Row = asRow1
    asTable.Row2 = asRow2
    asTable.Col = asCol1
    asTable.Col2 = asCol2
    asTable.BlockMode = True
    asTable.ForeColor = RGB(asR, asG, asB)
    asTable.BlockMode = False
End Sub

Public Function getJsonVar(ByRef v_strData As String) As clsJson

    Dim objResult As New clsJson
    Dim objCurrent As clsJson
    Dim i As Integer
    Dim IngStartPos As Long
    Dim IngEndPos As Long
    Dim IngLength As Long
    Dim strkey As String
    Dim strValue As String
    Dim strType As String
    Dim bMode As Boolean
    '-- 숫자값 처리용
    Dim bISNMode As Boolean
    Dim IngPrevStartPos As Long
    Dim IngPrevEndPos As Long
    
    mJsonData = ""
    bMode = False
    bISNMode = False
    IngStartPos = 0
    IngLength = Len(v_strData)
    
    Set objCurrent = objResult

    Do
        IngStartPos = InStr(IngStartPos + 1, v_strData, Chr$(34))
        IngEndPos = InStr(IngStartPos + 1, v_strData, Chr$(34))
        If IngEndPos = 0 Or IngStartPos = 0 Then Exit Do
    
        If bMode Then
            strValue = Mid$(v_strData, IngStartPos + 1, IngEndPos - IngStartPos - 1)
        Else
            If bISNMode = True Then
                strValue = mGetP(Mid$(v_strData, IngPrevStartPos + 2), 1, ",")
                bISNMode = False
            End If
            If strValue <> "" Then
                'Debug.Print Trim(strkey) & "@" & Trim(strValue)
                SetRawData Trim(strkey) & "@" & Trim(strValue)
                
                mJsonData = mJsonData & strkey & "@" & strValue & vbCr
                'SetRawData "[getJsonVar]" & mJsonData
    '            frmJson.Text2 = frmJson.Text2 & Trim(strkey) & ":" & Trim(strValue) & vbNewLine
            End If
            
            strkey = Mid$(v_strData, IngStartPos + 1, IngEndPos - IngStartPos - 1)
            
        End If
        
        'If strkey = "programs" Or strValue = "programs" Then Stop
        If strValue <> "" Then
            'Debug.Print Trim(strkey) & "@" & Trim(strValue)
            SetRawData Trim(strkey) & "@" & Trim(strValue)
            
            mJsonData = mJsonData & strkey & "@" & strValue & vbCr
            
'            frmJson.Text2 = frmJson.Text2 & Trim(strkey) & ":" & Trim(strValue) & vbNewLine
        End If
        'Debug.Print Mid$(v_strData, IngEndPos + 1, 1)
        
        Select Case Mid$(v_strData, IngEndPos + 1, 1)
            Case ":"
                'Debug.Print Mid$(v_strData, IngEndPos + 2, 1)
                Select Case Mid$(v_strData, IngEndPos + 2, 1)
                    Case "{"
                        Set objCurrent = objCurrent.addChild(strkey)
                    Case Chr$(34)
                        bMode = True
                    Case "["
                        Set objCurrent = objCurrent.addChild(strkey).addChild()
                    Case ")"
                        'Stop
                    Case Else
                        If IsNumeric(Mid$(v_strData, IngEndPos + 2, 1)) Then
                            bISNMode = True
                        End If
                End Select
            Case ","
                Call objCurrent.addChild(strkey, strValue)
                strkey = ""
                strValue = ""
                bMode = False
    
            Case "}"
                Call objCurrent.addChild(strkey, strValue)
                Set objCurrent = objCurrent.getParent
                strkey = ""
                strValue = ""
                bMode = False
    
                If Mid$(v_strData, IngEndPos + 3, 1) = "{" Then
                    Set objCurrent = objCurrent.addChild()
                End If
        End Select
        IngStartPos = IngEndPos
        IngPrevStartPos = IngStartPos
        IngPrevEndPos = IngEndPos
    Loop
    
    Set getJsonVar = objResult

End Function

Public Function getJsonVarPT(ByRef v_strData As String) As clsJson

    Dim objResult As New clsJsonPT
    Dim objCurrent As clsJsonPT
    Dim i As Integer
    Dim IngStartPos As Long
    Dim IngEndPos As Long
    Dim IngLength As Long
    Dim strkey As String
    Dim strValue As String
    Dim strType As String
    Dim bMode As Boolean
    '-- 숫자값 처리용
    Dim bISNMode As Boolean
    Dim IngPrevStartPos As Long
    Dim IngPrevEndPos As Long
    
    mJsonData = ""
    bMode = False
    bISNMode = False
    IngStartPos = 0
    IngLength = Len(v_strData)
    
    Set objCurrent = objResult

    Do
        IngStartPos = InStr(IngStartPos + 1, v_strData, Chr$(34))
        IngEndPos = InStr(IngStartPos + 1, v_strData, Chr$(34))
        If IngEndPos = 0 Or IngStartPos = 0 Then Exit Do
    
        If bMode Then
            strValue = Mid$(v_strData, IngStartPos + 1, IngEndPos - IngStartPos - 1)
        Else
            If bISNMode = True Then
                strValue = mGetP(Mid$(v_strData, IngPrevStartPos + 2), 1, ",")
                bISNMode = False
            End If
            If strValue <> "" Then
                'Debug.Print Trim(strkey) & "@" & Trim(strValue)
                SetRawData Trim(strkey) & "@" & Trim(strValue)
                
                mJsonData = mJsonData & strkey & "@" & strValue & vbCr
                'SetRawData "[getJsonVar]" & mJsonData
    '            frmJson.Text2 = frmJson.Text2 & Trim(strkey) & ":" & Trim(strValue) & vbNewLine
            End If
            
            strkey = Mid$(v_strData, IngStartPos + 1, IngEndPos - IngStartPos - 1)
            
        End If
        
        'If strkey = "programs" Or strValue = "programs" Then Stop
        If strValue <> "" Then
            'Debug.Print Trim(strkey) & "@" & Trim(strValue)
            SetRawData Trim(strkey) & "@" & Trim(strValue)
            
            mJsonData = mJsonData & strkey & "@" & strValue & vbCr
            
'            frmJson.Text2 = frmJson.Text2 & Trim(strkey) & ":" & Trim(strValue) & vbNewLine
        End If
        'Debug.Print Mid$(v_strData, IngEndPos + 1, 1)
        
        Select Case Mid$(v_strData, IngEndPos + 1, 1)
            Case ":"
                'Debug.Print Mid$(v_strData, IngEndPos + 2, 1)
                Select Case Mid$(v_strData, IngEndPos + 2, 1)
                    Case "{"
                        Set objCurrent = objCurrent.addChild(strkey)
                    Case Chr$(34)
                        bMode = True
                    Case "["
                        Set objCurrent = objCurrent.addChild(strkey).addChild()
                    Case ")"
                        'Stop
                    Case Else
                        If IsNumeric(Mid$(v_strData, IngEndPos + 2, 1)) Then
                            bISNMode = True
                        End If
                End Select
            Case ","
                Call objCurrent.addChild(strkey, strValue)
                strkey = ""
                strValue = ""
                bMode = False
    
            Case "}"
                Call objCurrent.addChild(strkey, strValue)
                Set objCurrent = objCurrent.getParent
                strkey = ""
                strValue = ""
                bMode = False
    
                If Mid$(v_strData, IngEndPos + 3, 1) = "{" Then
                    Set objCurrent = objCurrent.addChild()
                End If
        End Select
        IngStartPos = IngEndPos
        IngPrevStartPos = IngStartPos
        IngPrevEndPos = IngEndPos
    Loop
    
    Set getJsonVarPT = objResult

End Function



Public Function JsonSend(intAct As Integer, P() As Variant) As Variant
    Dim strURL      As String
    Dim strHeader   As String
    Dim varPara()   As Variant
    Dim varVal()    As Variant
    Dim strVHDV     As String
    
    '로그인
    If intAct = 1 Then
        strURL = gJSON.LOGIN
        strHeader = "srchMap"

        ReDim Preserve varPara(0) As Variant
        varPara(0) = "USER_ID"

        ReDim Preserve varVal(0) As Variant
        varVal(0) = P(0)
    
    '워크조회
    ElseIf intAct = 2 Then
        strURL = gJSON.WORK
        strHeader = "srchMap"

        ReDim Preserve varPara(4) As Variant
        varPara(0) = "SES_HSPT_CD"
        varPara(1) = "FROM_DATE"
        varPara(2) = "TO_DATE"
        varPara(3) = "LLRG_CD"
        varPara(4) = "VHDV_CD"

        ReDim Preserve varVal(4) As Variant
        varVal(0) = P(0)
        varVal(1) = P(1)
        varVal(2) = P(2)
        varVal(3) = P(3)
        varVal(4) = P(4)
        
    '바코드조회
    ElseIf intAct = 3 Then
        strURL = gJSON.BARCD
        strHeader = "srchMap"

        ReDim Preserve varPara(1) As Variant
        varPara(0) = "SES_HSPT_CD"
        varPara(1) = "BARCDNO"

        ReDim Preserve varVal(1) As Variant
        varVal(0) = P(0)
        varVal(1) = P(1)
    
    '결과저장
    ElseIf intAct = 9 Then
        strURL = gJSON.SAVE
        strHeader = "saveList"
        
        ReDim Preserve varPara(11) As Variant
        varPara(0) = "SES_HSPT_CD"
        varPara(1) = "BARCDNO"
        varPara(2) = "PRSCRT_CODENO"
        varPara(3) = "INSP_CLSFCT_CODENO"
        varPara(4) = "SMPORE_CD"
        varPara(5) = "INSP_EQP_YN"
        varPara(6) = "INSP_EQP_CODENO"
        varPara(7) = "RLTY_RSLT_CTNT"
        varPara(8) = "APLY_RSLT_CTNT"
        varPara(9) = "RSLT_STATE_CD"
        varPara(10) = "SES_USER_ID"
        varPara(11) = "SES_USER_IP"

        ReDim Preserve varVal(11) As Variant
        varVal(0) = P(0)
        varVal(1) = P(1)
        varVal(2) = P(2)     '처방코드
        varVal(3) = P(3)      '검사코드
        varVal(4) = P(4)      '검체코드(소변)
        varVal(5) = P(5)
        varVal(6) = P(6)      '장비코드
        varVal(7) = P(7)    '결과
        varVal(8) = P(8)    '결과
        varVal(9) = P(9)
        varVal(10) = P(10)
        varVal(11) = P(11)
    End If
    
    If frmMain.chkTest.Value = "1" Then
        JsonSend = JSONRPC(strURL, strHeader, varPara, varVal, intAct)
    Else
        JsonSend = JSONRPC(strURL, strHeader, varPara, varVal, -1)
    End If
    
End Function


Public Function JSONRPC(URL$, JSONPostHeader$, P() As Variant, V() As Variant, Optional intAct As Integer) As String
    Dim http    As Object
    Dim i       As Integer
    Dim JSONPostBody$()
  
On Error GoTo RST
    
    If intAct = -1 Then
       Set http = CreateObject("Winhttp.WinHttpRequest.5.1")
         
       http.Open "POST", URL, False
       http.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
       http.setRequestHeader "Accept", "application/json"
    
       ReDim Preserve JSONPostBody$(UBound(P))
       
       For i = 0 To UBound(P)
           JSONPostBody(i) = MakeJSONFromParams(P(i)) & ":" & MakeJSONFromParams(V(i))
       Next
       
       If JSONPostHeader = "saveList" Then
           If JSONPostHeader <> "" Then
               http.send "{" & MakeJSONFromParams(JSONPostHeader) & ":" & "[" & "{" & Join(JSONPostBody, ",") & "}" & "]" & "}"
               SetRawData "[SAVE]" & "{" & MakeJSONFromParams(JSONPostHeader) & ":" & "[" & "{" & Join(JSONPostBody, ",") & "}" & "]" & "}"
           
           Else
               http.send "{" & Join(JSONPostBody, ",") & "}"
           End If
       Else
           If JSONPostHeader <> "" Then
               http.send "{" & MakeJSONFromParams(JSONPostHeader) & ":" & "{" & Join(JSONPostBody, ",") & "}" & "}"
           Else
               http.send "{" & Join(JSONPostBody, ",") & "}"
           End If
       End If
       
       JSONRPC = http.responseText
       SetRawData JSONRPC
       
       Set http = Nothing
    Else
        '=============== 테스트 용 ===============
        '-- 오더파일명과 경로를 지정한다.
        Dim strPath     As String
        Dim strBuffer   As String
        Dim TextLine
        
        strBuffer = ""
        If intAct = 1 Then
            strPath = App.PATH & "\JSON_LOG\login.txt"
        ElseIf intAct = 2 Then
            strPath = App.PATH & "\JSON_LOG\work.txt"
        'ElseIf intAct = 3 Then
        '    strPath = App.PATH & "\JSON_LOG\barcode1.txt"
        ElseIf intAct = 3 Then
            strPath = App.PATH & "\JSON_LOG\barcode10.txt"
        ElseIf intAct = 5 Then
            strPath = App.PATH & "\JSON_LOG\save.txt"
        End If
        
        Open strPath For Input As #11 ' 파일을 엽니다.
    
        Do While Not EOF(11) ' 파일의 끝을 만날 때까지 반복합니다.
            Line Input #11, TextLine ' 변수로 데이터 행을 읽어들입니다.
            strBuffer = strBuffer & TextLine
        Loop
    
        Close #11 ' 파일을 닫습니다
    
        JSONRPC = strBuffer
        '=============== 테스트 용 ===============
    End If
    

Exit Function
RST:
    
    SetRawData Err.Number & vbNewLine & Err.Description
    
    JSONRPC = ""
    Set http = Nothing

End Function

'Public Function JSONRPC_Direct(URL$) As String
'    Dim httpd    As Object
'
'On Error GoTo RST
'
'    Set httpd = CreateObject("Winhttp.WinHttpRequest.5.1")
'
'    If optOpen(0).Value = True Then
'        httpd.Open "POST", URL, False
'    Else
'        httpd.Open "GET", URL, False
'    End If
'
'    httpd.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
'    httpd.setRequestHeader "Accept", "application/json"
'    httpd.send txtSend.Text
'
'    JSONRPC_Direct = httpd.responseText
'
'Exit Function
'RST:
'    MsgBox Err.Number & vbCr & Err.Description
'
'End Function



Public Function MakeJSONFromParams(ByVal P) As String 'Helper-function for the above main-request-function
    Dim Tmp$
    
    Select Case VarType(P)
        Case vbString:        Tmp = """" & P & """"
        Case vbBoolean:       Tmp = IIf(P, "true", "false")
        Case vbEmpty, vbNull: Tmp = "null"
        Case Else:            Tmp = Str$(P)
    End Select
    
    MakeJSONFromParams = Tmp
    
End Function

Public Function MakeJSONArrayFromParams(ByVal PArr) As String 'Helper-function for the above main-request-function
    Dim Tmp$(), P
    
    Tmp = Split(vbNullString)
    
    For Each P In PArr
        ReDim Preserve Tmp(0 To UBound(Tmp) + 1)
        Select Case VarType(P)
            Case vbString:        Tmp(UBound(Tmp)) = """" & P & """"
            Case vbBoolean:       Tmp(UBound(Tmp)) = IIf(P, "true", "false")
            Case vbEmpty, vbNull: Tmp(UBound(Tmp)) = "null"
            Case Else:            Tmp(UBound(Tmp)) = Str$(P)
        End Select
    Next
    
    MakeJSONArrayFromParams = "[" & Join(Tmp, ",") & "]"

End Function
