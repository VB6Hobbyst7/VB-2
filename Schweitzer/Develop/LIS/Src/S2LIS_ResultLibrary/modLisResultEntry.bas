Attribute VB_Name = "modResultEntry"
Option Explicit
' Template (LAB099) NoIndex 상수
'
Public OraDS As clsLisSqlResult
Public OraErr As clsLisError
Public gblnDBConnection As Boolean
Public glngErrorNo As Long
Public gstrErrorMsg As String
'
'Using API
Public Const WM_RBUTTONDOWN = &H204
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub Main()
'Create the Data Class
    Set OraDS = New clsLisSqlResult
    Set OraErr = New clsLisError
    OraErr.Reset
    gblnDBConnection = False
    'OraDS.InitConnect
End Sub

Public Function StrMergy(ParamArray aryStr()) As String
    Dim ii As Long
    StrMergy = ""
    For ii = LBound(aryStr) To UBound(aryStr)
       If ii = 0 Then
          StrMergy = CStr(aryStr(ii))
       Else
          StrMergy = StrMergy & vbTab & CStr(aryStr(ii))
       End If
    Next ii
End Function

'Function medFindAge(ByVal strBirthDate As String, ByVal strAgeType As String, _
'                    Optional ByVal strSysDate) As String
'    Dim strFormatBirth As String
'    Dim strFormatSys As String
'
'    strFormatBirth = Format(strBirthDate, "####/##/##")
'    If Not IsDate(strFormatBirth) Then strFormatBirth = Mid(strBirthDate, 1, 4) & "/01/01"
'    'strFormatBirth = Format(strFormatBirth, "yy-mm-dd")
'
'    'If IsMissing(strSysDate) Then
'        strFormatSys = Format(Now, "yyyy-mm-dd")
'    'Else
'    '    strFormatSys = Format(strSysDate, "####/##/##")
'    'End If
'
'    Select Case UCase(strAgeType)
'    Case "Y":        '년령
'        medFindAge = DateDiff("yyyy", strFormatBirth, strFormatSys)
'    Case "M":        '월령
'        medFindAge = DateDiff("m", strFormatBirth, strFormatSys)
'    Case "D":        '일령
'        medFindAge = DateDiff("d", strFormatBirth, strFormatSys)
'    End Select
'
'End Function

Public Function FormatNum(MyNumber As Double, FormatStr As String)
    FormatNum = Format(MyNumber, FormatStr)
    If Len(FormatNum) < Len(FormatStr) Then FormatNum = Space(Len(FormatStr) - Len(FormatNum)) & FormatNum
End Function

Function PadToString(intValue, intDigits)
    PadToString = String(intDigits - Len(intValue), "0") & intValue
End Function

'Public Sub CallBeep(ByVal pBeep As Long)
'    Dim ii As Long
'    For ii = 1 To pBeep
'        Beep
'    Next ii
'End Sub

Public Function FormatAccDt(ByVal strval As String) As String
    If Mid(strval, 1, 1) = "9" Then
        FormatAccDt = "19" & Trim(strval)
    Else
        FormatAccDt = "20" & Trim(strval)
    End If
End Function

Public Function DBAccNo(ByVal strval As String) As String
    
    Dim aryTmp() As String
    Dim ii As Integer
    Dim intLen As Integer
    
    aryTmp = Split(strval, "-")
    For ii = 1 To 2
       aryTmp(ii) = Trim(aryTmp(ii))
    Next
    If Mid(aryTmp(1), 1, 1) = "9" Then
       aryTmp(1) = "19" & aryTmp(1)
    Else
       aryTmp(1) = "20" & aryTmp(1)
    End If
    '
    DBAccNo = Join(aryTmp, "-")
    
End Function

'Public Function medShift(ByRef strText As String, ByVal strDeli As Variant) As String
'    Dim CNTA, CNTB As Integer
'    Dim Delimiter As String
'
'    medShift = "": CNTA = 0: CNTB = 0
'
'    CNTA = InStr(1, strText, strDeli)
'    If CNTA = 0 Then
'        medShift = strText
'        strText = ""
'        Exit Function
'    End If
'
'    medShift = Mid$(strText, 1, CNTA - Len(strDeli))
'    strText = Mid$(strText, CNTA + Len(strDeli))
'
'End Function


