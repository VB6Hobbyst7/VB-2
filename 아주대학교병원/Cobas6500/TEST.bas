Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
     ByVal codepage As Long, _
     ByVal dwFlags As Long, _
     ByVal lpWideCharStr As Long, _
     ByVal cchWideChar As Long, _
     ByVal lpMultiByteStr As Long, _
     ByVal cchMultiByte As Long, _
     ByVal lpDefaultChar As Long, _
     ByVal lpUsedDefaultChar As Long _
) As Long
     
Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
     ByVal codepage As Long, _
     ByVal dwFlags As Long, _
     ByVal lpMultiByteStr As Long, _
     ByVal cchMultiByte As Long, _
     ByVal lpWideCharStr As Long, _
     ByVal cchWideChar As Long _
) As Long
Private Const CP_UTF8 As Long = 65001



'### EUC-KR -> 한글(디코딩)
Public Function URLDecodeAnsi(Str As String) As String
On Error GoTo ErrLbl

     Dim AnsiArr() As Byte, i As Long, Buf() As String
     Buf = Split(Str, "%")
     ReDim AnsiArr(UBound(Buf) - 1)
    
     For i = 1 To UBound(Buf)
          AnsiArr(i - 1) = Val("&H" & Buf(i))
     Next i
    
     URLDecodeAnsi = StrConv(AnsiArr, vbUnicode)

ErrLbl:
End Function

'### UTF-8 -> 한글(디코딩)
Public Function URLDecodeUTF8(Str As String) As String
On Error GoTo ErrLbl

     Dim MultiArr() As Byte, strSplit() As String, i As Long, Converted() As Byte
     strSplit = Split(Str, "%")
     ReDim MultiArr(UBound(strSplit) - 1)
    
     For i = 1 To UBound(strSplit)
          MultiArr(i - 1) = Val("&H" & strSplit(i))
     Next i
    
     Dim BufSize As Long
     BufSize = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(MultiArr(0)), UBound(MultiArr) + 1&, 0&, 0&)
    
     If BufSize > 0 Then
          ReDim Converted(BufSize * 2 + 1)
          MultiByteToWideChar CP_UTF8, 0&, VarPtr(MultiArr(0)), UBound(MultiArr) + 1&, VarPtr(Converted(0)), BufSize
     End If
    
     URLDecodeUTF8 = Converted

ErrLbl:
End Function


'//////////////////////////////////////////////////////////////



'### 한글 -> UTF-8(인코딩)
Function URLEncodeUTF8(Str As String) As String
On Error GoTo ErrLbl

     Dim BufSize As Long, MultiArr() As Byte, Buf As String, i As Long
     Dim UniArr() As Byte
     UniArr = Str
    
     BufSize = WideCharToMultiByte(CP_UTF8, 0&, VarPtr(UniArr(0)), (UBound(UniArr) + 1) / 2, 0&, 0&, 0&, 0&)
    
     If BufSize > 0 Then
          ReDim MultiArr(BufSize - 1&)
          WideCharToMultiByte CP_UTF8, 0&, VarPtr(UniArr(0)), (UBound(UniArr) + 1) / 2, VarPtr(MultiArr(0)), BufSize, 0&, 0&
     End If
    
     For i = 0 To UBound(MultiArr)
          Buf = Buf & "%" & Hex$(MultiArr(i))
     Next i
    
     URLEncodeUTF8 = Buf

ErrLbl:
End Function

Function EncodeUTF8_ADOStream(strText As String) As Byte()

On Error GoTo ErrEncode


    Dim oStream As New ADODB.Stream
    
    Dim data() As Byte
    
    
    With oStream
        .Charset = "UTF-8"
        .mode = adModeReadWrite
        .Type = adTypeText
        .Open
        .WriteText strText
        .Flush
        .Position = 0
        .Type = adTypeBinary
        .Read 3
        
        data = .Read()
        .Close: Set oStream = Nothing
    End With
    
    
    EncodeUTF8_ADOStream = data

Exit Function

ErrEncode:


End Function




Public Function ConvertStringToUtf8Bytes(ByRef strText As String) As Byte()

    Dim objStream As ADODB.Stream
    Dim data() As Byte
    
    ' init stream
    Set objStream = New ADODB.Stream
    objStream.Charset = "utf-8"
    objStream.mode = adModeReadWrite
    objStream.Type = adTypeText
    objStream.Open
    
    ' write bytes into stream
    objStream.WriteText strText
    objStream.Flush
    
    ' rewind stream and read text
    objStream.Position = 0
    objStream.Type = adTypeBinary
    objStream.Read 3 ' skip first 3 bytes as this is the utf-8 marker
    data = objStream.Read()
    
    ' close up and return
    objStream.Close
    ConvertStringToUtf8Bytes = data

End Function

Public Function ConvertUtf8BytesToString(ByRef data() As Byte) As String

    Dim objStream As ADODB.Stream
    Dim strTmp As String
    
    ' init stream
    Set objStream = New ADODB.Stream
    objStream.Charset = "utf-8"
    objStream.mode = adModeReadWrite
    objStream.Type = adTypeBinary
    objStream.Open
    
    ' write bytes into stream
    objStream.Write data
    objStream.Flush
    
    ' rewind stream and read text
    objStream.Position = 0
    objStream.Type = adTypeText
    strTmp = objStream.ReadText
    
    ' close up and return
    objStream.Close
    ConvertUtf8BytesToString = strTmp

End Function


