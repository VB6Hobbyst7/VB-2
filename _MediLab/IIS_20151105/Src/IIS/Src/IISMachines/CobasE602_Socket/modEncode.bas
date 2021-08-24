Attribute VB_Name = "modEncode"
Option Explicit

Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
     ByVal CodePage As Long, _
     ByVal dwFlags As Long, _
     ByVal lpWideCharStr As Long, _
     ByVal cchWideChar As Long, _
     ByVal lpMultiByteStr As Long, _
     ByVal cchMultiByte As Long, _
     ByVal lpDefaultChar As Long, _
     ByVal lpUsedDefaultChar As Long _
) As Long

 

Private Const CP_UTF8 As Long = 65001


Public Function URLEncodeUTF8(Str As String) As String
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
