Attribute VB_Name = "Module1"
' Author: Le Duc Hong         http://www.vovisoft.com
Option Explicit
Function HexToDec(ByVal HexString) As Long
   ' Convert a Hex String (Max 4 characters) to decimal, assuming valid Hex digits
   '-------------------------------------------------------------------------------
   Dim i, TLen, AHex, NHex
   Dim Num As Long
   Num = 0
   TLen = Len(HexString)
   If (TLen = 0) Or (TLen > 4) Then
      HexToDec = 0
      Exit Function
   End If
   For i = 1 To TLen
      AHex = Mid(HexString, i, 1)
      If (AHex <= "9") Then
         NHex = Asc(AHex) - Asc("0")
      Else
         ' It is in range of "A".."F"
         NHex = 10 + (Asc(AHex) - Asc("A"))
      End If
      Num = Num * 16 + NHex
   Next
   HexToDec = Num
End Function

Public Function ReadTextFile(FileName) As String
' Write a Unicode String to UTF-16LE Text file
' Remember to Project | References "Microsoft Scripting Runtime" to support
'    FileSystemObject  & TextStream
   Dim Fs As FileSystemObject
   Dim TS As TextStream
   '  Create a FileSystem Object
   Set Fs = CreateObject("Scripting.FileSystemObject")
   ' Open TextStream for Input.
   ' TriStateTrue means Read Unicode UTF-16LE
   Set TS = Fs.OpenTextFile(FileName, ForReading, False, TristateTrue)
   ReadTextFile = TS.ReadAll  ' Read the whole content of the text file in one stroke
   TS.Close ' Close the Text Stream
   Set Fs = Nothing  ' Dispose FileSystem Object
End Function
Public Sub WriteTextFile(FileName, StrOutText)
' Read a Unicode String from UTF-16LE Text file
' Remember to Project | References "Microsoft Scripting Runtime" to support
'    FileSystemObject  & TextStream
   Dim Fs As FileSystemObject
   Dim TS As TextStream
   '  Create a FileSystem Object
   Set Fs = CreateObject("Scripting.FileSystemObject")
   ' Open TextStream for Output, create file if necesssary
   ' TriStateTrue means Write Unicode  UTF-16LE
   Set TS = Fs.OpenTextFile(FileName, ForWriting, True, TristateTrue)
   TS.Write StrOutText  ' Write the whole StrOutText string in one stroke
   TS.Close ' Close the Text Stream
   Set Fs = Nothing  ' Dispose FileSystem Object
End Sub
Function GetLocalDirectory() As String
' Return the folder where this program EXE resides
   Dim TStr
   TStr = App.Path  ' Get folder where this program EXE resides
   ' Append a back slash if it does not end with one
   If Right(TStr, 1) <> "\" Then TStr = TStr & "\"
   GetLocalDirectory = TStr ' Return it
End Function

