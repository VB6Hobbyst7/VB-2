Attribute VB_Name = "modJupsu"
Option Explicit

Public i                    As Integer
Public j                    As Integer
Public sMsg                 As String
Public sTitle               As String
Public nRow(1)              As Integer
Public GStrPtno             As String
Public GStrJdate            As String
Public GstrJupsuGubun       As String   '대기환자= D, 접수환자=J 구분 Flag
Public sQueryITem           As String
Public sJdate               As String
Public hWndReturn           As Long

Public iBlockRow            As Integer
Public iBlockRow2           As Integer

Public sLabelALLPrintIPD    As String
Public iSLip(11 To 50)      As Integer

Public GstrIOGubun          As String * 3 'OPD IPD

'Query
Public gSio                 As String
Public gSver                As String


Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Function convResultFormat(ByVal sRet As String) As String
    Dim nLength     As Integer
    
    Dim sLeft       As String * 6
    Dim sRight      As String * 4
    Dim nLeft       As Integer
    Dim nRight      As Integer
        
   '결과 Data 중 소숫점 을 기준으로 정렬시키는 함수 = 654321.123
    
   'Data Error Check
    nLength = Len(sRet)
    If nLength = 0 Then Exit Function                    'NULL Data 는 Exit
    
    If nLength > 11 Then                                 '자릿수가소숫점포함 11자리가 넘으면
        convResultFormat = sRet: Exit Function: End If   '    Data 를 그대로 Return
        
    If False = IsNumeric(sRet) Then                      'Character Data가 포함되어 있으면
        convResultFormat = sRet: Exit Function: End If   '    Data 를 그대로 Return

    
    nLeft = InStr(1, sRet, ".", vbTextCompare)
    If nLeft = 0 Then      '소숫점이 없는 Data
        RSet sLeft = sRet
    ElseIf nLeft > 0 Then
        RSet sLeft = Left(sRet, nLeft - 1)
        LSet sRight = Mid(sRet, nLeft, (Len(sRet) - nLeft) + 1) '소숫점때문에 +1....
    End If
    
    convResultFormat = sLeft & sRight
    
End Function

Public Function SetComboBox(ByVal sCombo As Object, ByVal sCompString As String, Optional nLtCnt As Integer = 0) As Integer
    
    
    If Trim(sCompString) = "" Then
        sCombo.ListIndex = -1
        Exit Function
    End If
    
    SetComboBox = False
    
    If Val(nLtCnt) > 0 Then
        GoSub String_LeftCut_Sub
    Else
        GoSub String_Normal_Sub
    End If
    Exit Function
    
String_Normal_Sub:
    For i = 0 To sCombo.ListCount - 1
        If Trim(sCombo.List(i)) = Trim(sCompString) Then
            sCombo.ListIndex = i
            SetComboBox = True
            Exit For
        End If
    Next
    Return
    
String_LeftCut_Sub:
    nLtCnt = Len(Trim(sCompString))
    For i = 0 To sCombo.ListCount - 1
        If Left(Trim(sCombo.List(i)), nLtCnt) = Trim(sCompString) Then
            sCombo.ListIndex = i
            SetComboBox = True
            Exit For
        
        End If
    Next
    Return
    
End Function


Public Function SpreadRowTopLine(ByVal sSpread As Object, ByVal iRow As Integer) As Integer
        
        sSpread.Row = iRow
        sSpread.Row2 = iRow
        sSpread.Col = 1
        sSpread.Col2 = sSpread.MaxCols
        sSpread.BlockMode = True
        sSpread.CellBorderType = SS_BORDER_TYPE_TOP
        sSpread.CellBorderStyle = CellBorderStyleSolid
        sSpread.Action = SS_ACTION_SET_CELL_BORDER
        sSpread.BlockMode = False

End Function


Public Function IsExamMaster(ByVal sPtno As String) As Integer
    Dim adoExamID       As ADODB.Recordset
    
    strSql = " SELECT Ptno FROM TWEXAM_IDnoMst WHERE Ptno = '" & sPtno & "'"
    
    If adoSetOpen(strSql, adoExamID) Then
        IsExamMaster = True
        Call adoSetClose(adoExamID)
    Else
        IsExamMaster = False
    End If

End Function

Public Function Bi_Check(ByVal sBi As String) As String

    Select Case Trim$(sBi)
        Case "11": Bi_Check = "공단"
        Case "12": Bi_Check = "직장"
        Case "13": Bi_Check = "지역"
        Case "14": Bi_Check = "지장1"
        Case "15": Bi_Check = "지역2"
        Case "16": Bi_Check = "직장1"
        Case "17": Bi_Check = "지역2"
        Case "21": Bi_Check = "보호1종"
        Case "22": Bi_Check = "보호2종"
        Case "23": Bi_Check = "의료시혜"
        Case "24": Bi_Check = "행려"
        Case "31": Bi_Check = "산재"
        Case "32": Bi_Check = "공상"
        Case "41": Bi_Check = "공단100%"
        Case "42": Bi_Check = "직장100%"
        Case "43": Bi_Check = "지역100%"
        Case "44": Bi_Check = "가족계획"
        Case "51": Bi_Check = "일반"
        Case "52": Bi_Check = "자보"
        Case "53": Bi_Check = "자보100%"
        Case "54": Bi_Check = "계약"
        Case "61": Bi_Check = "국내선박"
        Case "65": Bi_Check = "외국인"
        Case Else: Bi_Check = sBi
    End Select
    
End Function
Public Function Quot_Conv(ByVal sString As String) As Variant
    Dim sRecvStr
    Dim nStart      As Integer
    Dim sTemp       As String
    
    If Trim(Len(sString)) = "" Then Exit Function
    
    For nStart = 1 To Len(Trim(sString))
        sTemp = Mid(sString, nStart, 1)
        If Mid(sString, nStart, 1) = "'" Then
            sTemp = "''"
        ElseIf Mid(sString, nStart, 1) = """" Then
            sTemp = """"
        End If
        sRecvStr = sRecvStr & sTemp
    Next
    
    Quot_Conv = sRecvStr
    
    
End Function

Public Function ClearForm(ByVal sForm As Object) As Integer
    
    For i = 0 To sForm.Count - 1
        If TypeOf sForm.Controls(i) Is TextBox Then
            sForm.Controls(i).Text = ""
        ElseIf TypeOf sForm.Controls(i) Is ComboBox Then
            If sForm.Controls(i).Style = vbComboDropdownList Then
                sForm.Controls(i).ListIndex = -1
            Else
                sForm.Controls(i).Text = ""
            End If
        ElseIf TypeOf sForm.Controls(i) Is fpSpread Then
            sForm.Controls(i).Row = 1:
            sForm.Controls(i).Row2 = sForm.Controls(i).DataRowCnt
            sForm.Controls(i).Col = 1:
            sForm.Controls(i).Col2 = sForm.Controls(i).DataColCnt
            sForm.Controls(i).BlockMode = True
            sForm.Controls(i).Text = ""
            sForm.Controls(i).BlockMode = False
        End If
    Next
    
End Function
Public Function Spread_Set_Clear(ByVal sSpread As Object) As Integer
    
    sSpread.Row = 1
    
    sSpread.Row2 = sSpread.DataRowCnt
    sSpread.Col = 1
    sSpread.Col2 = sSpread.DataColCnt
    sSpread.BlockMode = True
    sSpread.Action = ActionClear
    sSpread.BlockMode = False
    
    
End Function
Public Function Dual_Date_Get(ByVal sFormat As String) As String
    Dim adoDual     As ADODB.Recordset
    
    If Trim(sFormat) = "" Then sFormat = "yyyy-MM-dd"
    
'o  strSql = " SELECT TO_CHAR(SysDate, '" & sFormat & "') ToDate FROM sys.Dual"
    strSql = " SELECT TO_CHAR(SysDate, '" & sFormat & "') ToDate FROM Dual"
    
    If False = adoSetOpen(strSql, adoDual) Then
        Dual_Date_Get = Format(Now, "yyyy-MM-dd")
        Exit Function
    End If
    
    Dual_Date_Get = adoDual.Fields("ToDate").Value & ""
    
    adoDual.Close
    If Not adoDual Is Nothing Then
        Set adoDual = Nothing
    End If
        
    Exit Function

End Function
Public Function Dual_Date_Cal_Get(ByVal sFormat As String, Optional sCnt = 0) As String
                                '(ex. Dual_Date_Cal_Get("yyyy-MM-dd", -7))
    Dim nReturn As Integer
    Dim adoDual As ADODB.Recordset
    
    On Error GoTo Error_Get
    
    If Trim(sFormat) = "" Then sFormat = "yyyy-MM-dd"
    
    
'o  strSql = " SELECT TO_CHAR(SysDate + " & sCnt & ", '" & sFormat & "') ToDate " & " FROM Sys.Dual"
    strSql = " SELECT TO_CHAR(SysDate + " & sCnt & ", '" & sFormat & "') ToDate " & " FROM Dual"
    If adoSetOpen(strSql, adoDual) Then
        Dual_Date_Cal_Get = Trim(adoDual.Fields("ToDate").Value)
        Call adoSetClose(adoDual)
    End If
    Exit Function
    
Error_Get:
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description
    Dual_Date_Cal_Get = Format(Now, "yyyy-MM-dd")
    Exit Function


End Function

Public Function IsAdmission(sPano As String) As Integer
    Dim adoIPD      As ADODB.Recordset
    
    'amSet1 : 0 = 재원중, 1=수납, 2=계산, 3=가퇴원, 9=심사완료
    'amSet6 : * = ghkswk rnqnsqusrud
    
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TWIPD_MASTER  INDEX_IPDMST2)  */ "
    
    strSql = ""
    strSql = strSql & " SELECT Ptno"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWIPD_MASTER"
    strSql = strSql & " WHERE  PTNO = '" & Trim$(sPano) & "'"
    
    If False = adoSetOpen(strSql, adoIPD) Then
        IsAdmission = False
        Exit Function
    End If
    
    If adoIPD.RecordCount = 0 Then
        IsAdmission = False
    Else
        IsAdmission = True
        adoIPD.Close
        If Not adoIPD Is Nothing Then Set adoIPD = Nothing
    End If

    
End Function

Public Function SetAge_Check(ByVal sJumin1 As String, sJumin2 As String) As String
    Dim nBirth  As Long
    Dim nTodate As Long
    
    If Trim$(sJumin1) = "" Then Exit Function
    If Trim$(sJumin2) = "" Then Exit Function
    If Len(Trim$(sJumin1)) <> 6 Then Exit Function
    If Len(Trim$(sJumin2)) <> 7 Then Exit Function
    
    nTodate = Format(CLng(Dual_Date_Get("yyyyMMdd")))
    
    Select Case Left(sJumin2, 1)
        Case "0", "9": nBirth = CLng(Trim("18" + sJumin1))  '1800년대 생년월일
        Case "1", "2": nBirth = CLng(Trim("19" + sJumin1))  '1900년대 생년월일
        Case "3", "4": nBirth = CLng(Trim("20" + sJumin1))  '2000년대 생년월일
        Case "7", "8": nBirth = CLng(Trim("19" + sJumin1))  '외국인 1900년대 Setting
        Case Else:     nBirth = CLng(Trim("19" + sJumin1))  'Default = 1900년대
    End Select
    
    Select Case nTodate - nBirth
        Case Is < 10000:    SetAge_Check = "1"                                      '1세미만
        Case Is < 100000:   SetAge_Check = Left(Trim(Str(nTodate - nBirth)), 1)     '10세이하
        Case Is < 1000000:  SetAge_Check = Left(Trim(Str(nTodate - nBirth)), 2)     '100세이하
        Case Is < 10000000: SetAge_Check = Left(Trim(Str(nTodate - nBirth)), 3)     '100세이상
        Case Else:          SetAge_Check = ""
    End Select
    
    
    
End Function

Public Function IsRoutineCode(ByVal sCode As String) As Integer
    Dim adoRt       As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & " SELECT RoutinCD"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Routine"
    strSql = strSql & " WHERE  RoutinCd = '" & sCode & "'"
    
    If False = adoSetOpen(strSql, adoRt) Then
        IsRoutineCode = False
        Exit Function
    Else
        IsRoutineCode = True
        Call adoSetClose(adoRt)
    End If
    
End Function

Public Function Get_RoutineName(ByVal sRoutineCD As String) As String
    Dim adoRoutin       As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & " SELECT DISTINCT RoutinNm"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Routine"
    strSql = strSql & " WHERE  RoutinCd = '" & Trim(sRoutineCD) & "'"
    If False = adoSetOpen(strSql, adoRoutin) Then
        Get_RoutineName = ""
    Else
        Get_RoutineName = adoRoutin.Fields("RoutinNm").Value & ""
        Call adoSetClose(adoRoutin)
    End If
    
End Function

Public Function Get_ItemName(ByVal sItemCode As String) As String
    Dim adoiTemNM       As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & " SELECT ITEMNM"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_itemML"
    strSql = strSql & " WHERE  CODEKY  = '" & Trim(sItemCode) & "'"
    
    If False = adoSetOpen(strSql, adoiTemNM) Then
        Get_ItemName = ""
    Else
        Get_ItemName = adoiTemNM.Fields("ITEMNM").Value & ""
        Call adoSetClose(adoiTemNM)
    End If
    
    
End Function

Public Function GET_SLipname(ByVal sSLipcode As String) As String
    Dim adoSpecode      As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & " SELECT CODENM"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Specode"
    strSql = strSql & " WHERE  CODEGU = '12'"
    strSql = strSql & " AND    CODEKY = '" & sSLipcode & "'"
    
    If False = adoSetOpen(strSql, adoSpecode) Then
        GET_SLipname = ""
    Else
        GET_SLipname = adoSpecode.Fields("Codenm").Value & ""
        Call adoSetClose(adoSpecode)
    End If
    
End Function
Public Function Get_Status(ByVal JeobsuDt As String, ByVal iSLno1 As Integer, ByVal iSLno2 As Integer) As String
    Dim adoStat     As ADODB.Recordset
    Dim sSqlSt      As String
    
    Get_Status = ""
    
    sSqlSt = ""
    sSqlSt = sSqlSt & " SELECT Status"
    sSqlSt = sSqlSt & " FROM   TWEXAM_General"
    sSqlSt = sSqlSt & " WHERE  JeobsuDt = TO_DATE('" & JeobsuDt & "','yyyy-MM-dd')"
    sSqlSt = sSqlSt & " AND    SLipno1  = " & iSLno1
    sSqlSt = sSqlSt & " AND    SLipno2  = " & iSLno2
    
    If False = adoSetOpen(sSqlSt, adoStat) Then Exit Function
    Get_Status = Trim(adoStat.Fields("Status").Value & "")
    
    
    
End Function
Public Function Get_NextLabno() As Integer
    Dim adoLabno        As ADODB.Recordset
    
    strSql = " SELECT SEQ_LABNO.NEXTVAL SLno2 FROM DUAL"
    
    If False = adoSetOpen(strSql, adoLabno) Then
        Get_NextLabno = 0
        Exit Function
    Else
        Get_NextLabno = Val(adoLabno.Fields("SLno2").Value & "")
        Call adoSetClose(adoLabno)
    End If

End Function
Public Function Get_MatchLabno() As Integer
    Dim adoMatch        As ADODB.Recordset
    
    strSql = " SELECT SEQ_MATCHNO.NEXTVAL MatchNO FROM DUAL"
    
    If False = adoSetOpen(strSql, adoMatch) Then
        Get_MatchLabno = 0
        Exit Function
    Else
        Get_MatchLabno = Val(adoMatch.Fields("MatchNO").Value & "")
        Call adoSetClose(adoMatch)
    End If

End Function






Public Function Get_Data_Labno(ByVal sJeobsuDt As String, ByVal iSLipno1 As Integer, ByVal sIO As String) As Integer
    Dim sRowID      As String
    Dim sSLipno2    As String
    Dim iRet        As Integer
    
    'IF sio = "O"외래  ,  "I" = 입원
    '검사종목별 Labno 를 분리하여 Select
    '입원은 10001 부터 , 외래는 00001 부터 시작함


    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID RwID"
    strSql = strSql & " FROM   TWEXAM_LABNO a"
    strSql = strSql & " WHERE  a.iDGubun  = '" & sIO & "'"
    strSql = strSql & " AND    a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.SLipno1  = " & iSLipno1
    strSql = strSql & " For UPdate"
    
    adoConnect.BeginTrans
    If adoSetOpen(strSql, adoSet) Then
        sRowID = adoSet.Fields("RwID").Value & ""
        Get_Data_Labno = Val(adoSet.Fields("SLipno2").Value & "")
        
        strSql = ""
        strSql = strSql & " UPDATE TWEXAM_LABNO"
        strSql = strSql & " SET    SLipno2  =  " & Get_Data_Labno + 1
        strSql = strSql & " WHERE  ROWID   = '" & sRowID & "'"
        
        If adoExec(strSql) Then
            adoConnect.CommitTrans
        Else
            adoConnect.RollbackTrans
        End If
        Call adoSetClose(adoSet)
    Else
        'Data 가 없어 새로운 Row를 시작할때
        '입원은 10001부터 시작한다.(외래는 00001 으로 시작함)
        strSql = ""
        strSql = strSql & " INSERT INTO TWEXAM_LABNO"
        strSql = strSql & "       ( JeobsuDt, iDGubun, SLipno1, SLipno2)"
        strSql = strSql & " VALUES(      TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD'),"
        strSql = strSql & "         '" & sIO & "',"
        strSql = strSql & "          " & iSLipno1 & ","
        
        Select Case Trim(sIO)
            Case "O": strSql = strSql & "              2)"
            Case "I": strSql = strSql & "          10002)"
        End Select

        If adoExec(strSql) Then
            adoConnect.CommitTrans
        Else
            adoConnect.RollbackTrans
        End If
        
        Select Case Trim(sIO)
            Case "O": Get_Data_Labno = 1
            Case "I": Get_Data_Labno = 10001
        End Select
    End If
    iRet = adoExec("COMMIT")
  

End Function

Public Function GetWardCode_FromRoom(ByVal sRoomCode As String) As String
    Dim adoWd       As ADODB.Recordset
    
    sRoomCode = Trim(sRoomCode)
    
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TWBas_ROOM INDEX_ROOM0) */"
    
    strSql = ""
    strSql = strSql & " SELECT a.WARDCode"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBas_ROOM a"
    strSql = strSql & " WHERE  a.RoomCode = '" & sRoomCode & "'"
    
    If False = adoSetOpen(strSql, adoWd) Then
        GetWardCode_FromRoom = ""
        Exit Function
    End If
    
    GetWardCode_FromRoom = adoWd.Fields("WARDCode").Value & ""
    Call adoSetClose(adoWd)

End Function
