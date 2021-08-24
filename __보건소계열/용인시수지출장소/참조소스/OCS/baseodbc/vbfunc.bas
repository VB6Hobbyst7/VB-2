Attribute VB_Name = "VbFunction"
Option Explicit

Global GstrSysTime      As String

Global rs1              As ADODB.Recordset
Global rs2              As ADODB.Recordset
Global RdoRes           As ADODB.Recordset

Global GstrDate         As String

'/========================================================================================================
Public Const IME_CMODE_NATIVE = &H1
Public Const IME_CMODE_HANGEUL = IME_CMODE_NATIVE
Public Const IME_CMODE_ALPHANUMERIC = &H0
Public Const IME_SMODE_NONE = &H0
Declare Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As Long) As Long
Declare Function ImmSetConversionStatus Lib "imm32.dll" (ByVal hIMC As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long

Public Sub cvtToHan(ByRef ArgObject As Object)
   Dim hIMC                 As Long
   
   hIMC = ImmGetContext(ArgObject.hwnd)
   ImmSetConversionStatus hIMC, IME_CMODE_HANGEUL, IME_SMODE_NONE
   
End Sub

Public Sub cvtToEng(ByRef ArgObject As Object)
   Dim hIMC                 As Long
   
   hIMC = ImmGetContext(ArgObject.hwnd)
   ImmSetConversionStatus hIMC, IME_CMODE_ALPHANUMERIC, IME_SMODE_NONE
   
End Sub

Function Date_Format(ArgForDate As String) As String

    Dim StrYY       As String
    Dim StrMM       As String
    Dim strDD       As String
    
    Select Case Len(ArgForDate)
        Case 10:    GoSub MID_10    'yyyy-mm-dd
        Case 8:     GoSub MID_8     'yyyymmdd or yy-mm-dd
        Case 6:     GoSub MID_6     'yymmdd
        Case Else:  Date_Format = "": Exit Function
    End Select

    Date_Format = ""
    If StrYY < "0000" Or StrYY > "9999" Then Exit Function
    If StrMM < "01" Or StrMM > "12" Then Exit Function
    If strDD < "01" Or strDD > "31" Then Exit Function
    If StrMM = "02" And strDD > "29" Then Exit Function
    
    If StrMM = "04" Or StrMM = "06" Or StrMM = "09" Or StrMM = "11" Then
        If strDD > "30" Then Exit Function
    End If
    
    If StrMM = "02" And strDD = "29" Then
        If StrYY <> "1996" And StrYY <> "2004" And _
           StrYY <> "2008" And StrYY <> "2012" And _
           StrYY <> "2016" And StrYY <> "2020" Then
           Exit Function
        End If
    End If
    
    Date_Format = StrYY & "-" & StrMM & "-" & strDD
    
    Exit Function


'/------------------------------------------------------------------------------------/

MID_10:
    
    StrYY = Mid$(ArgForDate, 1, 4)
    StrMM = Mid$(ArgForDate, 6, 2)
    strDD = Mid$(ArgForDate, 9, 2)

    Return


'/------------------------------------------------------------------------------------/

MID_8:
    If IsNumeric(ArgForDate) Then
        StrYY = Mid$(ArgForDate, 1, 4)
        StrMM = Mid$(ArgForDate, 5, 2)
        strDD = Mid$(ArgForDate, 7, 2)
    Else
        Select Case Mid$(ArgForDate, 1, 2)
            Case "00" To "10":
                StrYY = "20" & Mid$(ArgForDate, 1, 2)
            Case Else:
                StrYY = "19" & Mid$(ArgForDate, 1, 2)
        End Select
        StrMM = Mid$(ArgForDate, 4, 2)
        strDD = Mid$(ArgForDate, 7, 2)
    End If

    Return


'/------------------------------------------------------------------------------------/

MID_6:
    
    Select Case Mid$(ArgForDate, 1, 2)
        Case "00" To "10":
            StrYY = "20" & Mid$(ArgForDate, 1, 2)
        Case Else:
            StrYY = "19" & Mid$(ArgForDate, 1, 2)
    End Select
    StrMM = Mid$(ArgForDate, 3, 2)
    strDD = Mid$(ArgForDate, 5, 2)

    Return


End Function
Sub FORM_CENTER(FormName As Form) 'center it on the screen
    
    FormName.Top = (Screen.Height - FormName.Height - 400) \ 2
    FormName.Left = (Screen.Width - FormName.Width) \ 2

End Sub

Function AGE_DAY_GESAN(ArgJumin$, ArgGdate$) As Integer
    ' ArgJumin$ : 생년월일(6) + 주민번호(7)
    ' ArgGDate$ : 나이를 계산할 기준일자 (yyyy-mm-dd)
    ' *** 주민번호가 오류인 경우 999일로 처리함 ***

    Dim ArgMonth                As Double
    Dim ArgJuminLen             As Integer
    Dim ArgBirth                As String
    Dim ArgSex                  As String
    Dim ArgAGE                  As Integer

    '주민번호가 7보다 적으면 오류
    '기준일자는 반드시 'YYYY-MM-DD' Type이여야 함
    ArgJuminLen = Len(Trim(ArgJumin$))
    If ArgJuminLen < 7 Then AGE_DAY_GESAN = 999: Exit Function
    If Len(ArgGdate$) <> 10 Then AGE_DAY_GESAN = 999: Exit Function

    '성별을 Setting
    ArgSex = "1"
    If ArgJuminLen > 6 Then ArgSex = Mid(ArgJumin$, 7, 1)
    If ArgSex = "-" Then
        If ArgJuminLen > 7 Then
            ArgSex = Mid(ArgJumin$, 8, 1)
        Else
            ArgSex = "1"
        End If
    End If

    '생년월일을 YYYY-MM-DD Type으로 변경
    If ArgSex = "1" Or ArgSex = "2" Then
        ArgBirth = "19" & Left(ArgJumin$, 2) & "-" & Mid(ArgJumin$, 3, 2) & "-" & Mid(ArgJumin$, 5, 2)
    ElseIf ArgSex = "3" Or ArgSex = "4" Then
        ArgBirth = "18" & Left(ArgJumin$, 2) & "-" & Mid(ArgJumin$, 3, 2) & "-" & Mid(ArgJumin$, 5, 2)
    ElseIf ArgSex = "5" Or ArgSex = "6" Then
        ArgBirth = "20" & Left(ArgJumin$, 2) & "-" & Mid(ArgJumin$, 3, 2) & "-" & Mid(ArgJumin$, 5, 2)
    Else
        AGE_DAY_GESAN = 999: Exit Function
    End If

    '기준일자가 생년월일 보다 적으면 12개월 처리
    If ArgBirth >= ArgGdate$ Then AGE_DAY_GESAN = 999: Exit Function

    '주민번호가 오류이면 999일 처리
    If Not IsDate(Right(ArgBirth, 8)) Then AGE_DAY_GESAN = 999: Exit Function

    ArgAGE = DATE_ILSU(ArgGdate, ArgBirth) + 1
    If ArgAGE > 999 Then ArgAGE = 999
    
    AGE_DAY_GESAN = ArgAGE

End Function

Function AGE_MONTH_GESAN(ArgJumin$, ArgGdate$) As Integer
    ' ArgJumin$ : 생년월일(6) + 주민번호(7)
    ' ArgGDate$ : 나이를 계산할 기준일자 (yyyy-mm-dd)
    ' *** 주민번호가 오류인 경우 12개월로 처리함 ***

    Dim ArgMonth                As Double
    Dim ArgJuminLen             As Integer
    Dim ArgBirth                As String
    Dim ArgSex                  As String
    Dim ArgAGE                  As Integer

    '주민번호가 7보다 적으면 오류
    '기준일자는 반드시 'YYYY-MM-DD' Type이여야 함
    ArgJuminLen = Len(Trim(ArgJumin$))
    If ArgJuminLen < 7 Then AGE_MONTH_GESAN = 12: Exit Function
    If Len(ArgGdate$) <> 10 Then AGE_MONTH_GESAN = 12: Exit Function

    '성별을 Setting
    ArgSex = "1"
    If ArgJuminLen > 6 Then ArgSex = Mid(ArgJumin$, 7, 1)
    If ArgSex = "-" Then
        If ArgJuminLen > 7 Then
            ArgSex = Mid(ArgJumin$, 8, 1)
        Else
            ArgSex = "1"
        End If
    End If

    '생년월일을 YYYY-MM-DD Type으로 변경
    If ArgSex = "1" Or ArgSex = "2" Then
        ArgBirth = "19" & Left(ArgJumin$, 2) & "-" & Mid(ArgJumin$, 3, 2) & "-" & Mid(ArgJumin$, 5, 2)
    ElseIf ArgSex = "3" Or ArgSex = "4" Then
        ArgBirth = "18" & Left(ArgJumin$, 2) & "-" & Mid(ArgJumin$, 3, 2) & "-" & Mid(ArgJumin$, 5, 2)
    ElseIf ArgSex = "5" Or ArgSex = "6" Then
        ArgBirth = "20" & Left(ArgJumin$, 2) & "-" & Mid(ArgJumin$, 3, 2) & "-" & Mid(ArgJumin$, 5, 2)
    Else
        AGE_MONTH_GESAN = 12: Exit Function
    End If

    '기준일자가 생년월일 보다 적으면 12개월 처리
    If ArgBirth >= ArgGdate$ Then AGE_MONTH_GESAN = 12: Exit Function

    '주민번호가 오류이면 12개월 처리
    If Not IsDate(Right(ArgBirth, 8)) Then AGE_MONTH_GESAN = 12: Exit Function

    ArgAGE = 0
    strSQL = " SELECT MONTHS_BETWEEN(TO_DATE('" & ArgGdate & "','YYYY-MM-DD'),"
    strSQL = strSQL & "TO_DATE('" & ArgBirth & "','YYYY-MM-DD')) cAge FROM DUAL"
    Result = AdoOpenSet(rs2, strSQL)
    If rowindicator = 1 Then ArgAGE = Fix(AdoGetNumber(rs2, "cAge", 0))
    
    rs2.Close: Set rs2 = Nothing
    
    If ArgAGE >= 12 Then ArgAGE = 12
    AGE_MONTH_GESAN = ArgAGE

End Function

Function AGE_YEAR_GESAN(ArgJumin$, ArgGdate$) As Integer
    ' ArgJumin$ : 생년월일(6) + 주민번호(7)
    ' ArgGDate$ : 나이를 계산할 기준일자 (yyyy-mm-dd)
    ' *** 주민번호가 오류인 경우 10살로 처리함 ***

    Dim ArgMonth                As Double
    Dim ArgJuminLen             As Integer
    Dim ArgBirth                As String
    Dim ArgSex                  As String
    Dim ArgAGE                  As Integer

    '주민번호가 7보다 적으면 오류
    '기준일자는 반드시 'YYYY-MM-DD' Type이여야 함
    ArgJuminLen = Len(Trim(ArgJumin$))
    If ArgJuminLen < 7 Then AGE_YEAR_GESAN = 10: Exit Function
    If Len(ArgGdate$) <> 10 Then AGE_YEAR_GESAN = 10: Exit Function

    '성별을 Setting
    ArgSex = "1"
    If ArgJuminLen > 6 Then ArgSex = Mid(ArgJumin$, 7, 1)
    If ArgSex = "-" Then
        If ArgJuminLen > 7 Then
            ArgSex = Mid(ArgJumin$, 8, 1)
        Else
            ArgSex = "1"
        End If
    End If

    '생년월일을 YYYY-MM-DD Type으로 변경
    If ArgSex = "1" Or ArgSex = "2" Then
        ArgBirth = "19" & Left(ArgJumin$, 2) & "-" & Mid(ArgJumin$, 3, 2) & "-" & Mid(ArgJumin$, 5, 2)
    ElseIf ArgSex = "3" Or ArgSex = "4" Then
        ArgBirth = "18" & Left(ArgJumin$, 2) & "-" & Mid(ArgJumin$, 3, 2) & "-" & Mid(ArgJumin$, 5, 2)
    ElseIf ArgSex = "5" Or ArgSex = "6" Then
        ArgBirth = "20" & Left(ArgJumin$, 2) & "-" & Mid(ArgJumin$, 3, 2) & "-" & Mid(ArgJumin$, 5, 2)
    Else
        AGE_YEAR_GESAN = 10: Exit Function
    End If

    '기준일자가 생년월일 보다 적으면 0살 처리
    If ArgBirth >= ArgGdate$ Then AGE_YEAR_GESAN = 10: Exit Function

    '주민번호가 오류이면 10살 처리
    If Not IsDate(Right(ArgBirth, 8)) Then AGE_YEAR_GESAN = 10: Exit Function
    
    ArgAGE = 0
    strSQL = " SELECT MONTHS_BETWEEN(TO_DATE('" & ArgGdate & "','YYYY-MM-DD'),"
    strSQL = strSQL & "TO_DATE('" & ArgBirth & "','YYYY-MM-DD')) cAge FROM DUAL"
    Result = AdoOpenSet(rs2, strSQL)
    If rowindicator = 1 Then ArgAGE = Fix(AdoGetNumber(rs2, "cAge", 0) / 12)
    
    rs2.Close: Set rs2 = Nothing
    
    AGE_YEAR_GESAN = ArgAGE

End Function

Function ComboYYMM_TO_YYMM(ArgCombo) As String  'ComboType (yyyy년 mm월분) ==> yyyymm으로 변환
    Dim Inx             As Integer
    Dim ArgReturn       As String
    Dim ArgData         As String
    
    ArgReturn = ""
    For Inx = 1 To Len(ArgCombo)
        ArgData = Mid(ArgCombo, Inx, 1)
        If ArgData >= "0" And ArgData <= "9" Then
            ArgReturn = ArgReturn & ArgData
        End If
    Next Inx
    If Len(ArgReturn) <> 6 Then ArgReturn = ""
    ComboYYMM_TO_YYMM = ArgReturn

End Function

Function DATE_ADD(ArgDate$, ArgIlsu%) As String

    If Len(ArgDate$) <> 10 Then DATE_ADD = "": Exit Function

    strSQL = " SELECT TO_CHAR(TO_DATE('" & ArgDate & "','YYYY-MM-DD')"
    If ArgIlsu% < 0 Then
        strSQL = strSQL & "-" & ArgIlsu% * -1
    Else
        strSQL = strSQL & "+" & ArgIlsu%
    End If
    strSQL = strSQL & ",'YYYY-MM-DD') AddDate FROM DUAL"
    Result = AdoOpenSet(rs2, strSQL)
    If rowindicator = 1 Then
        DATE_ADD = AdoGetString(rs2, "AddDate", 0)
    Else
        DATE_ADD = ""
    End If
    
    rs2.Close: Set rs2 = Nothing
    
End Function

Function DATE_ILSU(ArgTdate$, ArgFdate$) As Integer

    If Len(Trim(ArgFdate$)) <> 10 Or Len(Trim(ArgTdate$)) <> 10 Then DATE_ILSU = 0: Exit Function

    If ArgFdate$ > ArgTdate$ Then DATE_ILSU = 0: Exit Function

    strSQL = " SELECT TO_DATE('" & ArgTdate$ & "','YYYY-MM-DD') - "
    strSQL = strSQL & "TO_DATE('" & ArgFdate$ & "','YYYY-MM-DD') Gigan FROM DUAL"
    Result = AdoOpenSet(rs2, strSQL)
    If rowindicator = 1 Then
        DATE_ILSU = AdoGetNumber(rs2, "Gigan", 0)
    Else
        DATE_ILSU = 0
    End If
    
    rs2.Close: Set rs2 = Nothing
    
End Function

Function DATE_TO_YYMM(ArgDate$) As String

    ' YYYY-MM-DD => YYYYMM, YYYYMMDD => YYYYMM으로 변환

    If Len(ArgDate$) = 10 Then
        DATE_TO_YYMM = Left(ArgDate$, 4) & Mid$(ArgDate$, 6, 2)
    ElseIf Len(ArgDate$) = 8 Then
        DATE_TO_YYMM = Left(ArgDate$, 6)
    Else
        DATE_TO_YYMM = ""
    End If

End Function

Function DATE_YYMM_ADD(ArgYYMM As String, ArgAdd%) As String
    Dim ArgI, ArgJ          As Integer
    Dim ArgYY, ArgMM        As Integer

    If Len(ArgYYMM) <> 6 Or ArgAdd% = 0 Then DATE_YYMM_ADD = ArgYYMM: Exit Function

    ArgYY = Val(Left$(ArgYYMM, 4))
    ArgMM = Val(Right$(ArgYYMM, 2))

    ArgJ = ArgAdd%
    If ArgJ < 0 Then ArgJ = ArgJ * -1

    For ArgI = 1 To ArgJ
        If ArgAdd% < 0 Then
            ArgMM = ArgMM - 1
            If ArgMM = 0 Then ArgMM = 12: ArgYY = ArgYY - 1
        Else
            ArgMM = ArgMM + 1
            If ArgMM = 13 Then ArgMM = 1: ArgYY = ArgYY + 1
        End If
    Next ArgI

    DATE_YYMM_ADD = Format(ArgYY, "0000") & Format(ArgMM, "00")

End Function

Function READ_LASTDAY(ArgDate$) As String
    
    strSQL = " SELECT TO_CHAR(LAST_DAY(TO_DATE('" & ArgDate$ & "','YYYY-MM-DD')),'YYYY-MM-DD') Lday FROM DUAL "
    Result = AdoOpenSet(rs2, strSQL)
    
    If rowindicator = 1 Then
        READ_LASTDAY = AdoGetString(rs2, "LDay", 0)
    Else
        READ_LASTDAY = ""
    End If
    
    rs2.Close: Set rs2 = Nothing
    
End Function

Sub READ_SYSDATE()
    strSQL = " SELECT TO_CHAR(SYSDATE,'YYYY-MM-DD HH24:MI') Sdate FROM DUAL "
    Result = AdoOpenSet(rs2, strSQL)
    
    If rowindicator = 1 Then
        GstrSysDate = Left(AdoGetString(rs2, "Sdate", 0), 10)
        GstrSysTime = Right(AdoGetString(rs2, "Sdate", 0), 5)
    Else
        GstrSysDate = "":  GstrSysTime = ""
    End If
    
    rs2.Close: Set rs2 = Nothing
    
End Sub

Function READ_YOIL(ArgDate$) As String
    
    strSQL = " SELECT TO_CHAR(TO_DATE('" & ArgDate$ & "','YYYY-MM-DD'),'DY') Yoil FROM DUAL "
    Result = AdoOpenSet(rs2, strSQL)
    
    If rowindicator = 1 Then
        Select Case UCase(AdoGetString(rs2, "Yoil", 0))
            Case "SUN": READ_YOIL = "일요일"
            Case "MON": READ_YOIL = "월요일"
            Case "TUE": READ_YOIL = "화요일"
            Case "WED": READ_YOIL = "수요일"
            Case "THU": READ_YOIL = "목요일"
            Case "FRI": READ_YOIL = "금요일"
            Case "SAT": READ_YOIL = "토요일"
            Case Else:  READ_YOIL = Trim(AdoGetString(rs2, "Yoil", 0))
        End Select
    Else
        READ_YOIL = ""
    End If
    
    rs2.Close: Set rs2 = Nothing
    
End Function
