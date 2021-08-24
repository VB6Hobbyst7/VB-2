Attribute VB_Name = "modResult"
Option Explicit

Public sRefDataMin      As String
Public sRefDataMax      As String
Public gOiLLQryPtno     As String
Public gSRmkSLipno      As String * 2
Public gResultPtno      As String

Public gSMicroCheck     As String


Public hWndReturn       As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long

Public Type SensResultType
    ItemCd(99)      As String
    Rcode(99)       As String
    AntiName(99)    As String * 20
    Sens(99)        As String
    Value(99)       As String
    Result(99)      As String
End Type

Public SensResult    As SensResultType

Public Sub SensResultClear()
    Dim nClsCnt     As Integer
    
    For nClsCnt = 0 To 99
        SensResult.ItemCd(nClsCnt) = ""
        SensResult.Rcode(nClsCnt) = ""
        SensResult.AntiName(nClsCnt) = ""
        SensResult.Sens(nClsCnt) = ""
        SensResult.Value(nClsCnt) = ""
        SensResult.Result(nClsCnt) = ""
    Next
    
End Sub

Public Function apiSetFocus(ByVal hWndFocus As Long) As Long
    Call SetFocus(hWndFocus)
    
End Function

Public Function SpreadSetClear(ByVal sObject As Object) As Integer
    
    sObject.Row = 1
    sObject.Row2 = sObject.DataRowCnt
    sObject.Col = 1
    sObject.Col2 = sObject.DataColCnt
    sObject.BlockMode = True
    sObject.Action = SS_ACTION_CLEAR_TEXT
    sObject.BlockMode = False
    
End Function
Public Function convResultFormat(ByVal sRet As String) As String
    Dim nLength     As Integer
    
    Dim sLeft       As String * 6
    Dim sRight      As String * 4
    Dim nLeft       As Integer
    Dim nRight      As Integer
        
   '��� Data �� �Ҽ��� �� �������� ���Ľ�Ű�� �Լ� = 654321.123
    
   'Data Error Check
    nLength = Len(sRet)
    If nLength = 0 Then Exit Function                    'NULL Data �� Exit
    
    If nLength > 11 Then                                 '�ڸ������Ҽ������� 11�ڸ��� ������
        convResultFormat = sRet: Exit Function: End If   '    Data �� �״�� Return
        
    If False = IsNumeric(sRet) Then                      'Character Data�� ���ԵǾ� ������
        convResultFormat = sRet: Exit Function: End If   '    Data �� �״�� Return

    
    nLeft = InStr(1, sRet, ".", vbTextCompare)
    If nLeft = 0 Then      '�Ҽ����� ���� Data
        RSet sLeft = sRet
    ElseIf nLeft > 0 Then
        RSet sLeft = Left(sRet, nLeft - 1)
        LSet sRight = Mid(sRet, nLeft, (Len(sRet) - nLeft) + 1) '�Ҽ��������� +1....
    End If
    
    convResultFormat = sLeft & sRight
    
End Function


Public Function SELECT_RefData(ByVal sSex As String, ByVal sAge As String, ByVal sItemCd As String) As Integer
    Dim adoRef      As ADODB.Recordset
    
    sRefDataMin = ""
    sRefDataMax = ""
    
    strSql = ""
    strSql = strSql & " SELECT * "
    strSql = strSql & " FROM   TWEXAM_REFDATA"
    strSql = strSql & " WHERE  ITEMCODE  = '" & sItemCd & "'"
    strSql = strSql & " AND    AGEMIN   <=  " & Val(sAge)
    strSql = strSql & " AND    AGEMAX   >=  " & Val(sAge)
    strSql = strSql & " AND    APPDATE   =     (SELECT MAX(APPDATE)"
    strSql = strSql & "                         FROM   TWEXAM_REFDATA"
    strSql = strSql & "                         WHERE  ITEMCODE = '" & sItemCd & "'"
    strSql = strSql & "                         AND    AGEMIN  <=  " & Val(sAge)
    strSql = strSql & "                         AND    AGEMAX  >=  " & Val(sAge) & ")"
    
    If adoSetOpen(strSql, adoRef) Then
        If sSex = "M" Then
            sRefDataMin = Trim(adoRef.Fields("M_MIN").Value & "")
            sRefDataMax = Trim(adoRef.Fields("M_MAX").Value & "")
        End If
        If sSex = "F" Then
            sRefDataMin = Val(Trim(adoRef.Fields("F_MIN").Value & ""))
            sRefDataMax = Val(Trim(adoRef.Fields("F_MAX").Value & ""))
        End If
        Call adoSetClose(adoRef)
    End If
    

End Function
Public Function GET_General_Status(ByVal JeobsuDt As String, ByVal SLno1 As Integer, ByVal SLno2 As Integer) As String
    Dim adoRet      As ADODB.Recordset
    
    GET_General_Status = ""
    
    strSql = ""
    strSql = strSql & " SELECT Status "
    strSql = strSql & " FROM   TWEXAM_GENERAL"
    strSql = strSql & " WHERE  JeobsuDt = TO_DATE('" & JeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    SLipno1  = " & SLno1
    strSql = strSql & " AND    SLipno2  = " & SLno2
    
    If False = adoSetOpen(strSql, adoRet) Then
        Exit Function
    End If
    
    GET_General_Status = adoRet.Fields("Status").Value & ""
    
    
    
End Function

Public Function Get_Result_Text(ByVal sItemCd As String) As String
    Dim nReturn     As Integer
    Dim adoResult   As ADODB.Recordset
    
    
    Get_Result_Text = ""
    
    strSql = ""
    strSql = strSql & " SELECT RET"
    strSql = strSql & " FROM   TWEXAM_RET"
    strSql = strSql & " WHERE  ITEMCD = '" & Trim$(sItemCd) & "'"
    strSql = strSql & " AND    RetGB  = 'A'"
    strSql = strSql & " ORDER  BY  Seqno"
    If False = adoSetOpen(strSql, adoResult) Then Exit Function
        
    Do Until adoResult.EOF
        Get_Result_Text = Get_Result_Text & _
                          RTrim(adoResult.Fields("RET").Value & "") & Chr$(9)
        adoResult.MoveNext
    Loop
    Call adoSetClose(adoResult)
    
    
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

Public Function Set_CheckBox_SqlSum(ByVal sObjectName As Object, ByVal SqlCompWard As String) As String
    Dim iWhereCnt   As Integer
    
    'CheckBox �� Tag �� ���� Data �� Setting ��Ų��.
    'ex) Set_CheckBox_SqlSum(chkWhere, " a.Status")
    
    
    
    '��� Check �Ǿ����� Count �Ѵ�.
    iWhereCnt = 0
    For i = sObjectName.LBound To sObjectName.UBound
        If sObjectName(i).Value = "1" Then iWhereCnt = iWhereCnt + 1
    Next
    'Check �Ȱ��� ������ ������ "" Setting��Ű�� Return
    If iWhereCnt = 0 Then Set_CheckBox_SqlSum = "": Exit Function
    
    'Sql ������ ��ģ��.
    Set_CheckBox_SqlSum = " AND ("
    For i = sObjectName.LBound To sObjectName.UBound
        If sObjectName(i).Value = "1" Then
            Set_CheckBox_SqlSum = Set_CheckBox_SqlSum & " " & SqlCompWard & " = '" & sObjectName(i).Tag & "' OR"
        End If
    Next
    
    '�� ������ ������ OR �� ")" �� �ٲ۴�. �׷��� Sql������ �ϼ�������!.
    If Right(Set_CheckBox_SqlSum, 2) = "OR" Then
        Set_CheckBox_SqlSum = Left(Set_CheckBox_SqlSum, Len(Set_CheckBox_SqlSum) - 2) & ")"
    End If


End Function
Public Function Get_OrgName(ByVal sRcode As String) As String
    Dim adoOrg      As ADODB.Recordset
    
    strSql = " SELECT ORG_NAME FROM TWEXAM_ORGLIST WHERE ORG_CODE = '" & Trim(sRcode) & "'"
    If False = adoSetOpen(strSql, adoOrg) Then
        Get_OrgName = ""
        Exit Function
    Else
        Get_OrgName = adoOrg.Fields("ORG_Name").Value & ""
        Call adoSetClose(adoOrg)
    End If
    
End Function

Public Function Get_iTemName(ByVal sItemCd As String) As String
    Dim adoItem      As ADODB.Recordset
    
    strSql = " SELECT itemnm FROM TWEXAM_ItemML WHERE Codeky = '" & Trim(sItemCd) & "'"
    If False = adoSetOpen(strSql, adoItem) Then
        Get_iTemName = ""
        Exit Function
    Else
        Get_iTemName = adoItem.Fields("ItemNM").Value & ""
        Call adoSetClose(adoItem)
    End If
    
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

Public Function GET_SLipname(ByVal sSLipno1 As String) As String
    Dim adoSLip     As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & " SELECT CODENM"
    strSql = strSql & " FROM   TWEXAM_Specode"
    strSql = strSql & " WHERE  CODEGU = '12'"
    strSql = strSql & " AND    CODEKY = '" & sSLipno1 & "'"
    If False = adoSetOpen(strSql, adoSLip) Then
        GET_SLipname = ""
        Exit Function
    End If
    GET_SLipname = Trim(adoSLip.Fields("Codenm").Value & "")
    Call adoSetClose(adoSLip)
    
End Function
Public Function GET_WardName(ByVal sRoomCode As String) As String
    Dim adoWD       As ADODB.Recordset
    Dim sSqlStr     As String
    
    
    sSqlStr = ""
    sSqlStr = sSqlStr & " SELECT WardCode"
    sSqlStr = sSqlStr & " FROM   TW_MIS_PMPA.TWBAS_Room"
    sSqlStr = sSqlStr & " WHERE  ROOMCode = '" & sRoomCode & "'"
    
    If False = adoSetOpen(sSqlStr, adoWD) Then
        GET_WardName = ""
        Exit Function
    End If
    GET_WardName = adoWD.Fields("WardCode").Value & ""
    
    Call adoSetClose(adoWD)
    
    
End Function
Public Function convLabnoToExpand(ByVal sComp5 As String) As String
    
    convLabnoToExpand = Format(DateAdd("d", Val(sComp5), "2000-10-01"), "YYYYMMDD")
        
    
End Function

Public Function convLabnoToComp(ByVal sYear8 As String) As String
    Dim sconvYear      As String
    
    sconvYear = Left(sYear8, 4) & "-" & Mid(sYear8, 5, 2) & "-" & Mid(sYear8, 7)
    
    convLabnoToComp = Format(DateDiff("d", "2000-10-01", sconvYear), "00000")
    
End Function




Public Function Get_MicroSeqno(ByVal iSLipno1 As Integer, ByVal strSampleCode As String, ByVal strDate As String) As Integer
    Dim strGubun        As String
    Dim nGubun          As Long
    Dim sSqlM           As String
    Dim nSeqno          As Long
    Dim bExists         As Boolean
    Dim sSampleGubun    As String
    
   
   
    '/ 0001~1999 = ����,�⵵,ȣ����ü
    '/ 2001~3999 = �񴢻��ı� ��ü
    '/ 4001~4999 = ��ȭ���ü
    '/ 5001~6999 = ü�׹ױ�Ÿ
    '/ 7001~8999 = ���׹�� ��ü
    '/ 9001~9999 = ����(AFB)
    
    
    GoSub Get_SampleGubun
    
    GoSub Exists_Mdate         '��ü�� Seqno �� �����´�. ������ ������, ����No Set
    
    If bExists = True Then
        GoSub MicroSeqno_Update
    Else
        GoSub MicroSeqno_Insert
    End If
    
    Get_MicroSeqno = nSeqno
    Exit Function
    
MicroSeqno_Insert:
    
    strSql = ""
    strSql = strSql & " INSERT INTO TWEXAM_MSeq"
    strSql = strSql & "       (Jdate, MCode, Seqno)"
    strSql = strSql & " VALUES( TO_DATE('" & strDate & "','yyyy-MM-dd'),"
    strSql = strSql & "         '" & sSampleGubun & "',"
    strSql = strSql & "          " & nSeqno & ")"
    adoConnect.BeginTrans
    
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    Return
    

MicroSeqno_Update:
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_MSeq"
    strSql = strSql & " SET    Seqno = " & nSeqno
    strSql = strSql & " WHERE  TO_CHAR(JDATE, 'yyyy-MM') = '" & Left(strDate, 7) & "'"
    strSql = strSql & " AND    MCODE = '" & sSampleGubun & "'"
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    Return
    
    

Exists_Mdate:
    Dim adoM        As ADODB.Recordset
    
    sSqlM = ""
    sSqlM = sSqlM & " SELECT *"
    sSqlM = sSqlM & " FROM   TWEXAM_MSeq"
    sSqlM = sSqlM & " WHERE  TO_CHAR(JDATE, 'yyyy-MM') = '" & Left(strDate, 7) & "'"
    sSqlM = sSqlM & " AND    MCODE = '" & sSampleGubun & "'"
    
    If False = adoSetOpen(sSqlM, adoM) Then
        nSeqno = nGubun + 1
        bExists = False      '���� Flag
        Return
    End If
    
    nSeqno = Val(adoM.Fields("Seqno").Value & "") + 1
    bExists = True     '���� Flag
    Call adoSetClose(adoM)
    
    Return



Get_SampleGubun:
    
    If iSLipno1 = 44 Then        'AFB ������ ����........
        nGubun = 9000
        sSampleGubun = "9"
    Else
        Select Case Trim(strSampleCode)
            Case "M2101", "M2102", "M2201", "M2202":          nGubun = 0
                                                              sSampleGubun = "0"
            Case "M2401", "M2402", "M2403", "M2405", "M2601": nGubun = 2000
                                                              sSampleGubun = "2"
            Case "M2701", "M2702", "M2703":                   nGubun = 4000
                                                              sSampleGubun = "4"
            Case "M2301", "M2302", "M2304", "M2305", "M2308": nGubun = 5000
                                                              sSampleGubun = "5"
            Case "M2309", "M2310", "M2311", "M2312", "M2399": nGubun = 5000
                                                              sSampleGubun = "5"
            Case "M2501", "M2503", "M2506", "M2507", "M2508": nGubun = 5000
                                                              sSampleGubun = "5"
            Case "M2509", "M2804":                            nGubun = 5000
                                                              sSampleGubun = "5"
            Case "M2001", "M2002":                            nGubun = 7000
                                                              sSampleGubun = "7"
            Case Else:                                        nGubun = 20000
                                                              sSampleGubun = "20"
        End Select
    End If
    
    Return
    
    
End Function


Public Function Get_CutOFFData(ByVal sArgItemCd As String, ByVal sResult1 As String) As String


    Get_CutOFFData = sResult1
    
    If IsNumeric(sResult1) = False Then Exit Function
    
    'HBs Ag ----------------------
    If Trim(sArgItemCd) = "310131" Then
        Select Case Val(sResult1)
            Case Is > 1:   Get_CutOFFData = "POSITIVE"
            Case 0.9 To 1: Get_CutOFFData = "Borderline"
            Case Is < 0.9: Get_CutOFFData = "NEGATIVE"
        End Select
    End If
    'HBs Ab ----------------------
    If Trim(sArgItemCd) = "310132" Then
        Select Case Val(sResult1)
            Case Is < 8:  Get_CutOFFData = "NEGATIVE"
            Case 8 To 12: Get_CutOFFData = "Borderline"
            Case Is > 12: Get_CutOFFData = "POSITIVE"
        End Select
    End If
    'anti-HIV ----------------------
    If Trim(sArgItemCd) = "310137" Or Trim(sArgItemCd) = "310138" Then
        Select Case Val(sResult1)
            Case Is >= 1:     Get_CutOFFData = "POSITIVE"
            Case 0.9 To 0.99: Get_CutOFFData = "Borderline"
            Case Is < 0.9:    Get_CutOFFData = "NEGATIVE"
        End Select
    End If
    'anti-HCV ----------------------
    If Trim(sArgItemCd) = "310135" Or Trim(sArgItemCd) = "310136" Then
        Select Case Val(sResult1)
            Case Is >= 1:     Get_CutOFFData = "POSITIVE"
            Case 0.9 To 0.99: Get_CutOFFData = "Borderline"
            Case Is < 0.9:    Get_CutOFFData = "NEGATIVE"
        End Select
    End If
    'HBc ----------------------
    If Trim(sArgItemCd) = "310141" Then
        Select Case Val(sResult1)
            Case Is < 1:      Get_CutOFFData = "POSITIVE"
            Case 1 To 1.19:   Get_CutOFFData = "Borderline"
            Case Is > 1.2:    Get_CutOFFData = "NEGATIVE"
        End Select
    End If
    'HBc ----------------------
    If Trim(sArgItemCd) = "310142" Then
        Select Case Val(sResult1)
            Case Is >= 1.2:   Get_CutOFFData = "POSITIVE"
            Case 0.8 To 1.19: Get_CutOFFData = "Borderline"
            Case Is < 0.8:    Get_CutOFFData = "NEGATIVE"
        End Select
    End If
    
    
End Function

