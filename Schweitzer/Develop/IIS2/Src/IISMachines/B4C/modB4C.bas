Attribute VB_Name = "modB4C"
Option Explicit

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Private Declare Function WideCharToMultiByte Lib "kernel32" _
                         (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Any, _
                         ByVal cchWideChar As Long, ByVal lpMultiByteStr As Any, ByVal cchMultiByte As Long, _
                         ByVal lpDefaultChar As Any, ByVal lpUsedDefaultChar As Long) As Long
                        
Private Declare Function MultiByteToWideChar Lib "kernel32.dll" _
                         (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef MultiByteStr As Any, _
                         ByVal cbMultiByte As Long, ByRef WideCharStr As Any, ByVal cchWideChar As Long) As Long

Const CP_ACP As Long = 0 ' default to ANSI code page
Const CP_UTF8 As Long = 65001


Function UrlEncode(ByVal urlText As String) As String
     Dim i As Long
     Dim ansi() As Byte
     Dim ascii As Integer
     Dim encText As String
    
     ansi = StrConv(urlText, vbFromUnicode)
     encText = ""
    
     For i = 0 To UBound(ansi)
          ascii = ansi(i)
         
          Select Case ascii
               Case 48 To 57, 65 To 90, 97 To 122
                    encText = encText & Chr(ascii)
               Case 32
                    encText = encText & "+"
               Case Else
                    If ascii < 16 Then
                         encText = encText & "%0" & Hex(ascii)
                    Else
                         encText = encText & "%" & Hex(ascii)
                    End If
          End Select
     Next i
    
    
    
     UrlEncode = encText
End Function

Public Function URLdecode(ByRef Text As String) As String
    Const Hex = "0123456789ABCDEF"
    Dim lngA As Long, lngB As Long, lngChar As Long, lngChar2 As Long
    URLdecode = Text
    lngB = 1
    For lngA = 1 To LenB(Text) - 1 Step 2
        lngChar = Asc(MidB$(URLdecode, lngA, 2))
        Select Case lngChar
            Case 37
                lngChar = InStr(Hex, MidB$(Text, lngA + 2, 2)) - 1
                If lngChar >= 0 Then
                    lngChar2 = InStr(Hex, MidB$(Text, lngA + 4, 2)) - 1
                    If lngChar2 >= 0 Then
                        MidB$(URLdecode, lngB, 2) = Chr$((lngChar * &H10&) Or lngChar2)
                        lngA = lngA + 4
                    Else
                        If lngB < lngA Then MidB$(URLdecode, lngB, 2) = MidB$(Text, lngA, 2)
                    End If
                Else
                    If lngB < lngA Then MidB$(URLdecode, lngB, 2) = MidB$(Text, lngA, 2)
                End If
            Case 43
                MidB$(URLdecode, lngB, 2) = " "
            Case Else
                If lngB < lngA Then MidB$(URLdecode, lngB, 2) = MidB$(Text, lngA, 2)
        End Select
        lngB = lngB + 2
    Next lngA
    URLdecode = LeftB$(URLdecode, lngB - 1)
End Function

'Public Function UrlEncode(ByRef Text As String) As String
'    Const Hex = "0123456789ABCDEF"
'    Dim lngA As Long, lngChar As Long
'    UrlEncode = Text
'    For lngA = LenB(UrlEncode) - 1 To 1 Step -2
'        lngChar = Asc(MidB$(UrlEncode, lngA, 2))
'        Select Case lngChar
'            Case 48 To 57, 65 To 90, 97 To 122
'            Case 32
'                MidB$(UrlEncode, lngA, 2) = "+"
'            Case Else
'                UrlEncode = LeftB$(UrlEncode, lngA - 1) & "%" & Mid$(Hex, (lngChar And &HF0) \ &H10 + 1, 1) & Mid$(Hex, (lngChar And &HF&) + 1, 1) & MidB$(UrlEncode, lngA + 2)
'        End Select
'    Next lngA
'End Function

      
Public Function URLdecshort(ByRef Text As String) As String
    Dim strArray() As String, lngA As Long
    strArray = Split(Replace(Text, "+", " "), "%")
    For lngA = 1 To UBound(strArray)
        strArray(lngA) = Chr$("&H" & Left$(strArray(lngA), 2)) & Mid$(strArray(lngA), 3)
    Next lngA
    URLdecshort = Join(strArray, vbNullString)
End Function

Public Function URLencshort(ByRef Text As String) As String
    Dim lngA As Long, strChar As String
    For lngA = 1 To Len(Text)
        strChar = Mid$(Text, lngA, 1)
        If strChar Like "[A-Za-z0-9]" Then
        ElseIf strChar = " " Then
            strChar = "+"
        Else
            strChar = "%" & Right$("0" & Hex$(Asc(strChar)), 2)
        End If
        URLencshort = URLencshort & strChar
    Next lngA
End Function


'### 2차 검사

Public Function g_xFile_Chk_UTF8(ByRef xBuf() As Byte) As Boolean
     Dim Tmp() As Byte
     Dim i As Long
     Dim x As Long
     Dim r As Long
    
     On Error GoTo Err

     x = UBound(xBuf)
    
     i = x + 1
    
     '### 일단 UTF-8이라 생각하고 ANSI로 변환 시작...
     r = MultiByteToWideChar(CP_UTF8, 0&, xBuf(0), i, 0&, 0&)
    
     If r Then
          ReDim Tmp(r * 2 - 1)
          r = MultiByteToWideChar(CP_UTF8, 0&, xBuf(0), i, Tmp(0), r)
     End If
    
     '### UTF-8이라면 한글등이 정상 변환되고 ANSI라면 변환이 않되고 Chr(32)로 표현됨...
     For i = 0 To x
          If xBuf(i) > 128 Then '### 숫자와 영문은 어차피 비교해봐야 같으니까... 한글등 2Byte 문자만 확인
               If Tmp(i * 2) = 32 And Tmp(i * 2 + 1) = 0 Then '### 한글등 2Byte 문자를 표현하지 못하면 ANSI라고 믿자...
                    g_xFile_Chk_UTF8 = False
                    Exit Function '### 전체를 계속 비교를 하다보면 조건을 만족하지 못하는 경우도 있으니 그냥 빠져나가자...
               Else
                    g_xFile_Chk_UTF8 = True     '### 한글등 2Byte 문자 표현이 가능하다면 UTF-8이라고 믿자...
                    Exit Function
               End If
          End If
     Next

     On Error GoTo 0
     GoTo End_Exit

Err:
'     Call g_xMsg_Err("xFile", "g_xFile_Chk_UTF8", g_Log_Path)
     Err.Clear
End_Exit:
     Erase Tmp
End Function

'---------------------------------------------------------------------------------------
' 함 수 명 : g_xFile_Get_Text_Format
' 작성일자 : 2009-07-10 11:03
' 작 성 자 : 어성화 pally4u@paran.com
' 함수설명 : 텍스트 파일(txt,dat,log등등 텍스트 형태로 저장된 파일)의 인코딩(포맷)정보를 얻어옴
' 인자설명 : xBuf:텍스트 파일의 바이트배열
' 리 턴 값 : 0:Unicode(Little Endian), 1:Unicode(Big Endian), 2:UTF-8, -1:ANSI 또는 Text가 아닌 파일, 에러발생시:99
'---------------------------------------------------------------------------------------
Public Function g_xFile_Get_Text_Format(ByRef xBuf() As Byte) As Long
'### 텍스트 파일의 BOM(Byte Order Mark)를 1차적으로 검사하고 ANSI일 경우 2차(g_xFile_Chk_UTF8) 검사를 수행하도록 변경함
     On Error GoTo Err

     If xBuf(0) = &HFF And xBuf(1) = &HFE Then
          '### Unicode (Little Endian: x86 기반의 Windows인 경우)
          g_xFile_Get_Text_Format = 0
     ElseIf xBuf(0) = &HFE And xBuf(1) = &HFF Then
          '### Unicode (Big Endian)
          g_xFile_Get_Text_Format = 1
     ElseIf xBuf(0) = &HEF And xBuf(1) = &HBB And xBuf(2) = &HBF Then
          '### UTF-8
          g_xFile_Get_Text_Format = 2
     Else
          If g_xFile_Chk_UTF8(xBuf) Then '### 2차 검사 수행
               '### BOM(Byte Order Mark)가 없는 UTF-8
               g_xFile_Get_Text_Format = 2
          Else
               '### ANSI 또는 Text가 아닌 파일 포맷
               g_xFile_Get_Text_Format = -1
          End If
     End If

     On Error GoTo 0
     GoTo End_Exit

Err:
     g_xFile_Get_Text_Format = 99
'     Call g_xMsg_Err("xFile", "g_xFile_Get_Text_Format", g_Log_Path)
     Err.Clear
End_Exit:
End Function

'---------------------------------------------------------------------------------------
' 함 수 명 : g_xFile_Get_File_To_Byte
' 작성일자 : 2009-04-22 08:11
' 작 성 자 : 어성화 pally4u@paran.com
' 함수설명 : 해당 파일의 내용을 읽어서 Byte배열로 반환
' 인자설명 : fName: 파일명, xBuf: 반환받을 byte배열
' 리 턴 값 : 성공시 True
'---------------------------------------------------------------------------------------
Public Function g_xFile_Get_File_To_Byte(ByRef fName As String, ByRef xBuf() As Byte) As Boolean
     Dim fNo As Integer     '파일번호
    
     On Error GoTo Err

     fNo = FreeFile
    
     Open fName For Binary Access Read As #fNo
    
     ReDim xBuf(LOF(fNo) - 1) As Byte     '### 넘겨받은 배열을 읽을 파일의 크기에 맞게 재설정
'     ReDim xBuf(FileLen(fName)) As Byte
    
     Get #fNo, , xBuf
    
     g_xFile_Get_File_To_Byte = True

     On Error GoTo 0
     GoTo End_Exit

Err:
'     Call g_xMsg_Err("xFile", "g_xFile_Get_File_To_Byte", g_Log_Path)
     Err.Clear
End_Exit:
     Close #fNo
End Function

 
 
 


