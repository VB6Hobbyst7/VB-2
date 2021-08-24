Attribute VB_Name = "modBasic"
Option Explicit

'Public strSql      As String
Public i           As Integer
Public j           As Integer
Public sMsg        As String
Public sTitle      As String



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

Public Function Dual_Date_Get(ByVal sFormat As String) As String
    Dim adoDual     As ADODB.Recordset
    
    If Trim(sFormat) = "" Then sFormat = "yyyy-MM-dd"
    
'O  strSql = " SELECT TO_CHAR(SysDate, '" & sFormat & "') ToDate FROM sys.Dual"
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
Public Function IsAdmission(sPano As String) As Integer
    Dim adoIPD      As ADODB.Recordset
    
    'amSet1
    '  0 = 재원중, 1=수납, 2=계산, 3=가퇴원, 9=심사완료
'o  strSql = ""
'o  strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWIPD_MASTER  INDEX_IPDMST2)  */ "

    strSql = ""
    strSql = strSql & " SELECT Ptno "
    strSql = strSql & " FROM   TW_MIS_PMPA.TWIPD_MASTER"
    strSql = strSql & " WHERE  PTNO    = '" & Trim$(sPano) & "'"
    strSql = strSql & " AND    amSet6  = ' '"
    strSql = strSql & " AND    amSet1  = '0'"
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

