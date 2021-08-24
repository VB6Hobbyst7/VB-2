Attribute VB_Name = "modPNV"
Option Explicit



Private Const CHUNK_SIZE& = 4096&
Private Const CP_UTF8 As Long = 65001

Private Declare Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cbMultiByte As Long, ByRef lpWideCharStr As Any, ByVal cchWideChar As Long) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" ( _
     ByRef Destination As Any, _
     ByRef Source As Any, _
     ByVal Length As Long _
)

Type Pnv_API
    APIURL     As String
    APIOrdPath As String
    APIRstPath As String
End Type

Public PnVAPI As Pnv_API


Public Function OpenURLWithIE2(ByVal sURL As String, ByVal sHeader As String, ByVal sBody As String, ByRef Inet As Inet) As String
     
    Dim TotBuf()        As Byte
    Dim ChunkedBuf()    As Byte
    Dim Converted()     As Byte
    Dim ni              As Long

     With Inet
          .Cancel
          .Execute sURL, "POST", sBody, sHeader
          
          Do While .StillExecuting
               DoEvents
          Loop
          
          ChunkedBuf() = .GetChunk(CHUNK_SIZE, icByteArray)
          
          Do While UBound(ChunkedBuf) >= 0
               ni = ni + UBound(ChunkedBuf) + 1
               ReDim Preserve TotBuf(ni - 1)
               RtlMoveMemory TotBuf(ni - UBound(ChunkedBuf) - 1), ChunkedBuf(0), UBound(ChunkedBuf) + 1&
               ChunkedBuf() = .GetChunk(CHUNK_SIZE, icByteArray)
          Loop
     End With
    
     Dim lSize As Long
     lSize = MultiByteToWideChar(CP_UTF8, 0&, TotBuf(0), UBound(TotBuf) + 1&, ByVal 0&, 0&)
    
     ReDim Converted(lSize * 2 - 1)
     MultiByteToWideChar CP_UTF8, 0&, TotBuf(0), UBound(TotBuf) + 1&, Converted(0), lSize
    
     OpenURLWithIE2 = Converted
     
End Function

