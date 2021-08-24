Attribute VB_Name = "modLibrary"
Option Explicit

'-- 지금날짜와 검사일자 비교한다
Public Function DateCompare(ByVal FDate As String) As String
    
    DateCompare = FDate
    If FDate <> Format(Now, "yyyymmdd") Then
        DateCompare = Format(Now, "yyyymmdd")
    End If
    
End Function

'-- 숫자외의 구분자를 모두 없앤다
Public Function SeperatorCls(ByVal asStr As String) As String
    Dim i       As Integer
    Dim StrLen  As Integer
    Dim RtStr   As String
    
    RtStr = ""

    For i = 1 To Len(asStr)
        If IsNumeric(Mid(asStr, i, 1)) Then
            RtStr = RtStr & Mid(asStr, i, 1)
        End If
    Next i
    
    SeperatorCls = RtStr
End Function

'-- 주민번호로 나이와 성별을 찾아온다.
'-- 인수 :  주민번호,현재날자
Public Sub CalAgeSex(ByRef asPNRN As String, ByVal asCurDate As String)
    Dim sBirth As String
    Dim sStart As String
    
    mPatient.SEX = ""
    mPatient.AGE = ""
    
    If Mid(asPNRN, 1, 1) = "_" Or Mid(asPNRN, 1, 1) = "" Then
        Exit Sub
    End If
    
    '수치만 가져온다 -,= 제거
    asPNRN = SeperatorCls(asPNRN)
    
    sStart = Trim(Mid(Trim(asPNRN), 7, 1))
    sBirth = Trim(Mid(Trim(asPNRN), 1, 6))
    
    If Mid(sBirth, 3, 4) = "0000" Then
        sBirth = Mid(sBirth, 1, 2) & "0101"
    End If
    
    Select Case sStart
        Case "1", "3", "5", "7"
            mPatient.SEX = "M"
        Case "2", "4", "6", "8"
            mPatient.SEX = "F"
    End Select

    Select Case sStart
        Case "1", "2"
            sBirth = "19"
        Case "3", "4"
            sBirth = "20"
        Case "7", "8"
            sBirth = "18"
        Case Else
            sBirth = "19"
    End Select
    
    sBirth = sBirth & Mid(asPNRN, 1, 2) & "/" & Mid(asPNRN, 3, 2) & "/" & Mid(asPNRN, 5, 2)
    'sBirth = sBirth & "-01"
    'Else
    '   sBirth = sBirth & "/" & Mid(asPNRN, 3, 2)
    'End If
    'If Mid(asPNRN, 5, 2) = "00" Then
 '      sBirth = sBirth & "-01"
    'Else
    '    sBirth = sBirth & "/" & Mid(asPNRN, 5, 2)
    'End If
    
    mPatient.AGE = DateDiff("yyyy", sBirth, asCurDate) + 1
    mPatient.AGE = mPatient.AGE - 1
    
End Sub

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

Public Function SetTag(ByRef vasTable As Object, ByVal SetStr As String, ByVal vasRow As Long, ByVal vasCol As Long) As Boolean
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    vasTable.CellTag = SetStr
End Function

Public Function SetToolTip(ByRef vasTable As Object, ByVal SetStr As String, ByVal vasRow As Long, ByVal vasCol As Long) As Boolean
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    vasTable.ToolTipText = SetStr
End Function

Public Function GetText(ByRef vasTable As Object, ByVal vasRow As Long, ByVal vasCol As Long) As String
    If vasRow < 0 Or vasCol < 0 Then
        Exit Function
    End If
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    GetText = vasTable.Text
End Function

Public Function GetTag(ByRef vasTable As Object, ByVal vasRow As Long, ByVal vasCol As Long) As String
    If vasRow < 0 Or vasCol < 0 Then
        Exit Function
    End If
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    GetTag = vasTable.CellTag
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
    Dim sFileName   As String
    Dim FindFile    As String
    
    If gHOSP.LOQWRITE = "0" Then
        Exit Sub
    End If
    
    FilNum = FreeFile
    
    If Dir(App.PATH & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.PATH & "\Log")
    End If
    
    sFileName = gHOSP.MACHNM & "_" & Format(CDate(Now), "yyyy-mm-dd")
    
    Open App.PATH & "\Log\" & sFileName & ".txt" For Append As FilNum
    
    Print #FilNum, argSQL;
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
    
    sFileName = gHOSP.MACHNM & "_" & Format(CDate(Now), "yyyy-mm-dd") & "_" & strName
    
    If argMode = "A" Then
        Open App.PATH & "\Log\" & sFileName & ".txt" For Append As FilNum
    Else
        Open App.PATH & "\Log\" & sFileName & ".txt" For Output As FilNum
    End If
    Print #FilNum, argSQL
    Close FilNum
    
End Sub

Public Sub SetErrData(ByVal strName As String, ByVal argSQL As String, Optional ByVal argMode As String)
    Dim FilNum
    Dim sFileName As String
    
    If gHOSP.LOQWRITE = "0" Then
        Exit Sub
    End If
    
    FilNum = FreeFile
        
    If Dir(App.PATH & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.PATH & "\Log")
    End If
    
    sFileName = gHOSP.MACHNM & "_" & Format(CDate(Now), "yyyy-mm-dd") & "_" & strName
    
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

Public Sub SetBackColor(asTable As Object, ByVal asRow1 As Long, ByVal asRow2 As Long, ByVal asCol1 As Long, ByVal asCol2 As Long, asR As Variant, asG As Variant, asB As Variant)
    asTable.Row = asRow1
    asTable.Row2 = asRow2
    asTable.Col = asCol1
    asTable.Col2 = asCol2
    asTable.BlockMode = True
    asTable.BackColor = RGB(asR, asG, asB)
    asTable.BlockMode = False
End Sub

Public Sub SetForeColor(asTable As Object, ByVal asRow1 As Long, ByVal asRow2 As Long, ByVal asCol1 As Long, ByVal asCol2 As Long, asR As Variant, asG As Variant, asB As Variant)
    asTable.Row = asRow1
    asTable.Row2 = asRow2
    asTable.Col = asCol1
    asTable.Col2 = asCol2
    asTable.BlockMode = True
    asTable.ForeColor = RGB(asR, asG, asB)
    asTable.BlockMode = False
End Sub


