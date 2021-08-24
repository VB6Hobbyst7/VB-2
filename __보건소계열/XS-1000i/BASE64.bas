Attribute VB_Name = "Base64"
Option Explicit


Public Const base64_alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="

'  Text1 = makeB64("메디메이트medimate가a나b")
'  Text2 = makeUB64(Text1)

Public Function makeB64(ByVal vbstrValue As String) As String
   Dim arrByte(2) As Variant
   Dim arrStringByte() As Byte
   Dim baseValue As Long
   Dim arrBase64(3) As Byte
   Dim i, j
   Dim tmpByte As Long
   Dim lastFlag As Boolean
   Dim LastString As String
   
   j = 0
   For i = 1 To Len(vbstrValue)
      ReDim Preserve arrStringByte(j)
      tmpByte = Asc(Mid(vbstrValue, i, 1))
      If (tmpByte And &HFF00) <> 0 Then
         arrStringByte(j) = ((tmpByte And &HFF00) \ 2 ^ 8) And &HFF
         j = j + 1
         ReDim Preserve arrStringByte(j)
         arrStringByte(j) = tmpByte And &HFF
      Else
         arrStringByte(j) = tmpByte And &HFF
      End If
      j = j + 1
   Next
   
   For i = 0 To UBound(arrStringByte) Step 3
      arrByte(0) = arrStringByte(i)
      If i + 1 > UBound(arrStringByte) Then
         arrByte(1) = 0
         lastFlag = True
      Else
         arrByte(1) = arrStringByte(i + 1)
         lastFlag = False
      End If
      If i + 2 > UBound(arrStringByte) Then
         arrByte(2) = 0
         lastFlag = True
      Else
         arrByte(2) = arrStringByte(i + 2)
         lastFlag = False
      End If
   
      baseValue = "&H" & Hex(arrByte(0)) & "0000"
      baseValue = baseValue Or ("&H" & Hex(arrByte(1)) & "00")
      baseValue = baseValue Or ("&H" & Hex(arrByte(2)))
   
      arrBase64(0) = (baseValue \ 2 ^ 18) And &H3F
      LastString = LastString & returnBase64(arrBase64(0), False)
      arrBase64(1) = baseValue \ 2 ^ 12 And &H3F
      LastString = LastString & returnBase64(arrBase64(1), lastFlag)
      arrBase64(2) = baseValue \ 2 ^ 6 And &H3F
      LastString = LastString & returnBase64(arrBase64(2), lastFlag)
      arrBase64(3) = baseValue And &H3F
      LastString = LastString & returnBase64(arrBase64(3), lastFlag)
   Next
   
   makeB64 = LastString
End Function

Public Function makeUB64(ByVal vbstrString As String) As String
   Dim i, Ls
   Dim midValue As Long
   Dim arrValue(2) As Byte
   Dim arrTmpValue(3) As Byte
   Dim tmpHanCode As Byte
   Dim tmpString As String
   Dim temp As Long
   
   Ls = Len(vbstrString)
   
   For i = 1 To Ls Step 4
      arrTmpValue(0) = returnUBase64(Mid(vbstrString, i, 1))
'      MsgBox Hex(arrTmpValue(0))
      arrTmpValue(1) = returnUBase64(Mid(vbstrString, i + 1, 1))
'      MsgBox Hex(arrTmpValue(1))
      arrTmpValue(2) = returnUBase64(Mid(vbstrString, i + 2, 1))
'      MsgBox Hex(arrTmpValue(2))
      arrTmpValue(3) = returnUBase64(Mid(vbstrString, i + 3, 1))
'      MsgBox Hex(arrTmpValue(3))
      
      arrValue(0) = (arrTmpValue(0) * 2 ^ 2)
      arrValue(0) = arrValue(0) Or ((arrTmpValue(1) \ 2 ^ 4) And &H3)
      If tmpHanCode <> 0 Then
         temp = "&H" & Hex(tmpHanCode) & "00"
         temp = temp Or arrValue(0)
         
         tmpString = tmpString & Chr(temp)
         tmpHanCode = 0
      Else
         If (arrValue(0) And &H80) = 0 Then
            tmpString = tmpString & Chr(arrValue(0))
         Else
            tmpHanCode = arrValue(0)
         End If
      End If
            
      arrValue(1) = (arrTmpValue(1) And &HF) * 2 ^ 4
      arrValue(1) = arrValue(1) Or ((arrTmpValue(2) \ 2 ^ 2) And &HF)
      If tmpHanCode <> 0 Then
         temp = "&H" & Hex(tmpHanCode) & "00"
         temp = temp Or arrValue(1)
         
         tmpString = tmpString & Chr(temp)
         tmpHanCode = 0
      Else
         If (arrValue(1) And &H80) = 0 Then
            tmpString = tmpString & Chr(arrValue(1))
         Else
            tmpHanCode = arrValue(1)
         End If
      End If
      
      arrValue(2) = (arrTmpValue(2) And &H3) * 2 ^ 6
      arrValue(2) = arrValue(2) Or arrTmpValue(3)
      If tmpHanCode <> 0 Then
         temp = "&H" & Hex(tmpHanCode) & "00"
         temp = temp Or arrValue(2)
         
         tmpString = tmpString & Chr(temp)
         tmpHanCode = 0
      Else
         If (arrValue(2) And &H80) = 0 Then
            tmpString = tmpString & Chr(arrValue(2))
         Else
            tmpHanCode = arrValue(2)
         End If
      End If
      
      makeUB64 = tmpString
      
'      midValue = "&H" & Hex(arrTmpValue(0)) & "
   Next
End Function

Public Function returnBase64(reNum As Byte, lFlag As Boolean) As String
   If lFlag And reNum = 0 Then
      returnBase64 = "="
   Else
      returnBase64 = Mid(base64_alphabet, reNum + 1, 1)
   End If
End Function

Public Function returnUBase64(vbstrValue As String) As Byte
   If vbstrValue = "=" Then
      returnUBase64 = 0
   Else
      returnUBase64 = InStr(base64_alphabet, vbstrValue) - 1
   End If
End Function


