Attribute VB_Name = "modHStrCode"
Option Explicit

Public Function LenH(ByVal anystr As String) As Integer
    LenH = LenB(StrConv(anystr, vbFromUnicode))
End Function

Public Function LeftH(ByVal anystr As String, ByVal nPos As Integer) As String
    LeftH = StrConv(LeftB(StrConv(anystr, vbFromUnicode), nPos), vbUnicode)
End Function

Public Function RightH(ByVal anystr As String, ByVal nPos As Integer) As String
    RightH = StrConv(RightB(StrConv(anystr, vbFromUnicode), nPos), vbUnicode)
End Function

Public Function MidH(ByVal anystr As String, ByVal nStartPos As Integer, nSize As Integer) As String
    MidH = StrConv(MidB(StrConv(anystr, vbFromUnicode), nStartPos, nSize), vbUnicode)
End Function

Public Function fGetHexaCode(ByVal InputString As String) As String

    Dim strUniCode As String

    Dim i As Integer

    strUniCode = StrConv(InputString, vbFromUnicode)

    For i = 1 To LenB(strUniCode)
'        fGetHexaCode = fGetHexaCode & Hex(AscB(MidB(strUniCode, i, 1)))
        fGetHexaCode = fGetHexaCode & Right("00" & Hex(AscB(MidB(strUniCode, i, 1))), 2)
    Next i

End Function

Public Sub fGetByteData(ByVal vData As Variant, ByRef bSend() As Byte)

'    Dim a(100) As Byte '입력예상되는 최대 길이의 두배로 잡으세요
'    Dim nVal As Integer 'ascii값저장
'    Dim aVal As String '한글 ascii값을 문자열로 자장
'    Dim nLen As Integer '문자열의 길이저장
'    Dim k As Integer
        
    Dim aByte() As Byte
    Dim iLen%, iVal%
    Dim ii%, kk%
    Dim sVal$       '한글 ascii값을 문자열로 자장
    
    iLen = Len(vData)
    
'    ReDim aByte(LenH(vData))
    ReDim aByte(Len(vData) * 2)
    
    kk = 0
    
    For ii = 1 To iLen
        iVal = Asc(Mid(vData, ii, 1))
    
        '한글인 경우는 아스키값이 음수입니다.
        If iVal < 0 Then
            sVal = Hex(iVal)
            
            aByte(kk) = Val("&h" & Left(sVal, 2))
            kk = kk + 1
            aByte(kk) = Val("&h" & Right(sVal, 2))
        Else
            aByte(kk) = iVal
        End If
        
        kk = kk + 1
    Next
    
    '입력된 문자를 Byte Array에 저장
    bSend = aByte
    
End Sub
Public Sub fGetByteData_temp(ByVal vData As Variant, ByRef bSend() As Byte)
        
    Dim aByte() As Byte
    Dim iLen%, iVal%
    Dim ii%, kk%
    Dim sVal$       '한글 ascii값을 문자열로 자장
    
    iLen = Len(vData)
    
    ReDim aByte(LenH(vData))
    
    kk = 0
    
    For ii = 1 To iLen
        iVal = Asc(Mid(vData, ii, 1))
    
        '한글인 경우는 아스키값이 음수입니다.
        If iVal < 0 Then
            sVal = Hex(iVal)
            
            aByte(kk) = "&h" & Left(sVal, 2)
            kk = kk + 1
            aByte(kk) = "&h" & Right(sVal, 2)
        Else
            aByte(kk) = "&h" & Hex(iVal)
        End If
        
        kk = kk + 1
    Next
    
'    '입력된 문자를 Byte Array에 저장
'    fGetByteData = aByte
    bSend = aByte
    
End Sub
'바이트 어래이에 저장된 아스키 값을 문자(한글 영문 숫자)로 바꾼다.
Public Function Fu_Read_Name(ByRef szTemp1() As Byte) As String

    Dim i As Integer
    Dim strTemp As String
    Dim rName As String
    Dim szTemp As String
    
    For i = 0 To 100
        If szTemp1(i) >= &H80 Then
            strTemp = "&H" & Hex(szTemp1(i))
            i = i + 1
            strTemp = strTemp & Hex(szTemp1(i))
            rName = rName & Chr(Val(strTemp))
        ElseIf szTemp1(i) = &H0 Then
            Exit For
        Else
            rName = rName & Chr(szTemp1(i))
        End If
    Next i
    
    Fu_Read_Name = rName

End Function



