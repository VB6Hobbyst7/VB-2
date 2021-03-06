Attribute VB_Name = "modCFX96_RV16"
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


'### 2?? ????

Public Function g_xFile_Chk_UTF8(ByRef xBuf() As Byte) As Boolean
     Dim Tmp() As Byte
     Dim i As Long
     Dim x As Long
     Dim r As Long
    
     On Error GoTo Err

     x = UBound(xBuf)
    
     i = x + 1
    
     '### ???? UTF-8???? ???????? ANSI?? ???? ????...
     r = MultiByteToWideChar(CP_UTF8, 0&, xBuf(0), i, 0&, 0&)
    
     If r Then
          ReDim Tmp(r * 2 - 1)
          r = MultiByteToWideChar(CP_UTF8, 0&, xBuf(0), i, Tmp(0), r)
     End If
    
     '### UTF-8?????? ???????? ???? ???????? ANSI???? ?????? ?????? Chr(32)?? ??????...
     For i = 0 To x
          If xBuf(i) > 128 Then '### ?????? ?????? ?????? ?????????? ????????... ?????? 2Byte ?????? ????
               If Tmp(i * 2) = 32 And Tmp(i * 2 + 1) = 0 Then '### ?????? 2Byte ?????? ???????? ?????? ANSI???? ????...
                    g_xFile_Chk_UTF8 = False
                    Exit Function '### ?????? ???? ?????? ???????? ?????? ???????? ?????? ?????? ?????? ???? ??????????...
               Else
                    g_xFile_Chk_UTF8 = True     '### ?????? 2Byte ???? ?????? ?????????? UTF-8?????? ????...
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
' ?? ?? ?? : g_xFile_Get_Text_Format
' ???????? : 2009-07-10 11:03
' ?? ?? ?? : ?????? pally4u@paran.com
' ???????? : ?????? ????(txt,dat,log???? ?????? ?????? ?????? ????)?? ??????(????)?????? ??????
' ???????? : xBuf:?????? ?????? ??????????
' ?? ?? ?? : 0:Unicode(Little Endian), 1:Unicode(Big Endian), 2:UTF-8, -1:ANSI ???? Text?? ???? ????, ??????????:99
'---------------------------------------------------------------------------------------
Public Function g_xFile_Get_Text_Format(ByRef xBuf() As Byte) As Long
'### ?????? ?????? BOM(Byte Order Mark)?? 1???????? ???????? ANSI?? ???? 2??(g_xFile_Chk_UTF8) ?????? ?????????? ??????
     On Error GoTo Err

     If xBuf(0) = &HFF And xBuf(1) = &HFE Then
          '### Unicode (Little Endian: x86 ?????? Windows?? ????)
          g_xFile_Get_Text_Format = 0
     ElseIf xBuf(0) = &HFE And xBuf(1) = &HFF Then
          '### Unicode (Big Endian)
          g_xFile_Get_Text_Format = 1
     ElseIf xBuf(0) = &HEF And xBuf(1) = &HBB And xBuf(2) = &HBF Then
          '### UTF-8
          g_xFile_Get_Text_Format = 2
     Else
          If g_xFile_Chk_UTF8(xBuf) Then '### 2?? ???? ????
               '### BOM(Byte Order Mark)?? ???? UTF-8
               g_xFile_Get_Text_Format = 2
          Else
               '### ANSI ???? Text?? ???? ???? ????
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
' ?? ?? ?? : g_xFile_Get_File_To_Byte
' ???????? : 2009-04-22 08:11
' ?? ?? ?? : ?????? pally4u@paran.com
' ???????? : ???? ?????? ?????? ?????? Byte?????? ????
' ???????? : fName: ??????, xBuf: ???????? byte????
' ?? ?? ?? : ?????? True
'---------------------------------------------------------------------------------------
Public Function g_xFile_Get_File_To_Byte(ByRef fName As String, ByRef xBuf() As Byte) As Boolean
     Dim fNo As Integer     '????????
    
     On Error GoTo Err

     fNo = FreeFile
    
     Open fName For Binary Access Read As #fNo
    
     ReDim xBuf(LOF(fNo) - 1) As Byte     '### ???????? ?????? ???? ?????? ?????? ???? ??????
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

 
 
 


