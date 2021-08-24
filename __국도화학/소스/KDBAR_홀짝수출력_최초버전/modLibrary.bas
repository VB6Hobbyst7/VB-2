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
    
    If gKUKDO.LOQWRITE = "0" Then
        Exit Sub
    End If
    
    FilNum = FreeFile
    
    If Dir(App.PATH & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.PATH & "\Log")
    End If
    
    sFileName = Format(CDate(Now), "yyyy-mm-dd")
    
    Open App.PATH & "\Log\" & sFileName & ".txt" For Append As FilNum
    
    Print #FilNum, argSQL;
    Close FilNum

End Sub

Public Sub SetSQLData(ByVal strName As String, ByVal argSQL As String, Optional ByVal argMode As String)
    Dim FilNum
    Dim sFileName As String
    
    If gKUKDO.LOQWRITE = "0" Then
        Exit Sub
    End If
    
    FilNum = FreeFile
        
    If Dir(App.PATH & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.PATH & "\Log")
    End If
    
    sFileName = Format(CDate(Now), "yyyy-mm-dd") & "_" & strName
    
    If argMode = "A" Then
        Open App.PATH & "\Log\" & sFileName & ".txt" For Append As FilNum
    Else
        Open App.PATH & "\Log\" & sFileName & ".txt" For Output As FilNum
    End If
    Print #FilNum, argSQL
    Close FilNum
    
End Sub

Public Sub SetPrtData(ByVal strName As String, ByVal argSQL As String, Optional ByVal argMode As String)
    Dim FilNum
    Dim sFileName As String
    
    If gKUKDO.LOQWRITE = "0" Then
        Exit Sub
    End If
    
    FilNum = FreeFile
        
    If Dir(App.PATH & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.PATH & "\Log")
    End If
    
    sFileName = strName & "_" & Format(CDate(Now), "yyyy-mm-dd")
    
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
    
    If gKUKDO.LOQWRITE = "0" Then
        Exit Sub
    End If
    
    FilNum = FreeFile
        
    If Dir(App.PATH & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.PATH & "\Log")
    End If
    
    sFileName = Format(CDate(Now), "yyyy-mm-dd") & "_" & strName
    
    If argMode = "A" Then
        Open App.PATH & "\Log\" & sFileName & ".txt" For Append As FilNum
    Else
        Open App.PATH & "\Log\" & sFileName & ".txt" For Output As FilNum
    End If
    Print #FilNum, argSQL
    Close FilNum
    
End Sub

Public Sub SetXMLData(ByVal strXmlName As String, ByVal sRcvData As Variant)  ', Optional ByVal argMode As String
    Dim STM         As ADODB.Stream

On Error GoTo ErrHandle
        
    If Dir(App.PATH & "\Xml", vbDirectory) <> "Xml" Then
        MkDir (App.PATH & "\Xml")
    End If
    
    '## 파일오픈
    Set STM = New ADODB.Stream
    
    STM.Open
    STM.Type = adTypeText
    STM.Charset = "utf-8"
    STM.WriteText sRcvData
    STM.SaveToFile App.PATH & "\Xml\" & strXmlName, adSaveCreateOverWrite ' adSaveCreateNotExist
    STM.Close
    Set STM = Nothing
    
Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gKUKDO.MACHNM & "_SetXMLData" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

    
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

'Y:4자리숫자,M:2자리숫자,D:2자리숫자
Public Function Get_YMD(ByVal pCode As String, ByVal pData As String) As String
    Dim strY    As String
    Dim strD    As String
    
    Get_YMD = ""
    strY = ""
    strD = ""
    
    If Mid(pCode, 1, 1) = "Y" And Len(pData) = 4 And IsNumeric(pData) Then
        Select Case UCase(pCode)
            Case "Y1"
                Get_YMD = pData
            
            Case "Y2"   '2099 년까지만 유효함
                Get_YMD = Mid(pData, 3, 2)
                
            Case "Y3"
'                Get_YMD = CCur(pData) - 2020
'
'                If CCur(Get_YMD) < 0 Then
'                    Get_YMD = ""
'                End If
                
                Get_YMD = Mid(pData, 4, 1)
                
            Case "Y4"   '2036(Z) 년까지 유효
                        '2010=A
                
                strY = CCur(pData) - 1945
                Get_YMD = Chr(strY)
            
            Case "Y5"   '2032(Z) 년까지 유효
                        '2011=A
                        'I(73),O(79),U(85),V(86) 제외
                
                strY = CCur(pData) - 1946
                If CCur(strY) >= 73 Then strY = strY + 1    'I
                If CCur(strY) >= 79 Then strY = strY + 1    'O
                If CCur(strY) >= 85 Then strY = strY + 1    'U
                If CCur(strY) >= 86 Then strY = strY + 1    'V
                Get_YMD = Chr(strY)
            
            Case "Y6"   '2034(Z) 년까지 유효
                        '2010=A
                        'N(78),O(79) 제외
                
                strY = CCur(pData) - 1945
                Get_YMD = Chr(strY)
                If CCur(strY) >= 78 Then strY = strY + 1    'N
                If CCur(strY) >= 79 Then strY = strY + 1    'O
                Get_YMD = Chr(strY)
            
            Case "Y7"   '2032(Z) 년까지 유효
                        '2011=A

                strY = CCur(pData) - 1946
                Get_YMD = Chr(strY)
                
                
        End Select
    
    ElseIf Mid(pCode, 1, 1) = "M" And IsNumeric(pData) Then
        Select Case UCase(pCode)
            Case "M1"
                Get_YMD = pData
            
            Case "M2"   '1,2,3,4,5,6,7,8,9,A(10),B(11),C(12)
                If CCur(pData) < 10 Then
                    Get_YMD = Val(pData)
                ElseIf CCur(pData) = 10 Then
                    Get_YMD = "A"
                ElseIf CCur(pData) = 11 Then
                    Get_YMD = "B"
                ElseIf CCur(pData) = 12 Then
                    Get_YMD = "C"
                End If
                
            Case "M3"   '1,2,3,4,5,6,7,8,9,O(10),D(11),N(12) :Octovet,November,December
                If CCur(pData) < 10 Then
                    Get_YMD = Val(pData)
                ElseIf CCur(pData) = 10 Then
                    Get_YMD = "O"
                ElseIf CCur(pData) = 11 Then
                    Get_YMD = "N"
                ElseIf CCur(pData) = 12 Then
                    Get_YMD = "D"
                End If
                
        End Select
        
    ElseIf Mid(pCode, 1, 1) = "D" And IsNumeric(pData) Then
        Select Case UCase(pCode)
            Case "D1"
                Get_YMD = Format(pData, "00")
            
            Case "D2"   '1,2,3,4,5,6,7,8,9,A(10),B(11),C(12).......Z(35)
                If CCur(pData) < 10 Then
                    Get_YMD = Val(pData)
                ElseIf CCur(pData) >= 10 Then
                    strD = 55 + CCur(pData)
                    Get_YMD = Chr(strD)
                End If
                
            Case "D3"   '1,2,3,4,5,6,7,8,9,A(10),B(11),C(12).......Z(33)   I,O 제외
                If CCur(pData) < 10 Then
                    Get_YMD = Val(pData)
                ElseIf CCur(pData) >= 10 Then
                    strD = 55 + CCur(pData)
                    If CCur(strD) >= 73 Then strD = strD + 1    'I
                    If CCur(strD) >= 79 Then strD = strD + 1    'O
                    
                    Get_YMD = Chr(strD)
                End If
                
            Case "D4"   '1,2,3,4,5,6,7,8,9,A(10),B(11),C(12).......Z(31)   I,O,U,V 제외
                If CCur(pData) < 10 Then
                    Get_YMD = Val(pData)
                ElseIf CCur(pData) >= 10 Then
                    strD = 55 + CCur(pData)
                    If CCur(strD) >= 73 Then strD = strD + 1    'I
                    If CCur(strD) >= 79 Then strD = strD + 1    'O
                    If CCur(strD) >= 85 Then strD = strD + 1    'U
                    If CCur(strD) >= 86 Then strD = strD + 1    'V
                    
                    Get_YMD = Chr(strD)
                End If
                
        End Select
    End If
    
End Function

'Y:4자리숫자,M:2자리숫자,D:2자리숫자
Public Function Get_Len(ByVal pCode As String, ByVal pData As String) As String
    Get_Len = ""
    
    If pCode = "L1" And IsNumeric(pData) Then
        Select Case pData
            Case "10":        Get_Len = "A"
            Case "100":       Get_Len = "B"
            Case "1000":      Get_Len = "C"
            Case "10000":     Get_Len = "D"
            Case "100000":    Get_Len = "E"
            Case "1000000":   Get_Len = "F"
        End Select
    End If
    
End Function

Public Function GetLotNo(ByVal pMakeDate As String, ByVal pSeq As String, ByVal pPackCd As String, ByVal pCompUserCd) As String
    Dim strLotNo  As String
    
    GetLotNo = ""
    
    If Len(pMakeDate) <> 10 Then
        pMakeDate = Format(pMakeDate, "####-##-##")
    End If
    
    strLotNo = ""
    strLotNo = strLotNo & Get_YMD("Y6", Year(pMakeDate))
    strLotNo = strLotNo & Get_YMD("M3", MONTH(pMakeDate))
    strLotNo = strLotNo & "F"   'ITEM구분 ACF = F
    strLotNo = strLotNo & Format(Day(pMakeDate), "00")
    strLotNo = strLotNo & pSeq
    strLotNo = strLotNo & pPackCd
    strLotNo = strLotNo & pCompUserCd
    
    GetLotNo = strLotNo
    
End Function















