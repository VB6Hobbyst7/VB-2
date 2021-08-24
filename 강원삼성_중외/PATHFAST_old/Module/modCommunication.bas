Attribute VB_Name = "modCommunication"
Option Explicit

Public Const COM_NUL    As Long = &H0
Public Const COM_SOH    As Long = &H1
Public Const COM_STX    As Long = &H2
Public Const COM_ETX    As Long = &H3
Public Const COM_EOT    As Long = &H4
Public Const COM_ENQ    As Long = &H5
Public Const COM_ACK    As Long = &H6
Public Const COM_BEL    As Long = &H7
Public Const COM_BS     As Long = &H8
Public Const COM_HT     As Long = &H9
Public Const COM_LF     As Long = &HA
Public Const COM_VT     As Long = &HB
Public Const COM_FF     As Long = &HC
Public Const COM_CR     As Long = &HD
Public Const COM_SO     As Long = &HE
Public Const COM_SI     As Long = &HF
Public Const COM_DEL    As Long = &H10
Public Const COM_DC1    As Long = &H11
Public Const COM_DC2    As Long = &H12
Public Const COM_DC3    As Long = &H13
Public Const COM_DC4    As Long = &H14
Public Const COM_NACK   As Long = &H15
Public Const COM_SYN    As Long = &H16
Public Const COM_ETB    As Long = &H17
Public Const COM_CAN    As Long = &H18
Public Const COM_EM     As Long = &H19
Public Const COM_SUB    As Long = &H1A
Public Const COM_ESC    As Long = &H1B
Public Const COM_FS     As Long = &H1C
Public Const COM_GS     As Long = &H1D
Public Const COM_RS     As Long = &H1E
Public Const COM_US     As Long = &H1F
Public Const COM_SP     As Long = &H20

' 통신데이타 저장 여부
Public mlngRecLen           As Long             ' 이전에 장비에서 들어온 내용 길이

Public Function Get_CRC(ByVal StrCRC As String) As String
    Dim bytTemp() As Byte
    Dim i As Integer
    Dim XOR_Rst As Byte
    
    bytTemp = StrConv(StrCRC, vbFromUnicode)
    XOR_Rst = bytTemp(LBound(bytTemp))
    
    For i = (LBound(bytTemp) + 1) To UBound(bytTemp)
        XOR_Rst = XOR_Rst Xor bytTemp(i)
    Next i
    
    If XOR_Rst = COM_ETX Then
        XOR_Rst = &H7F
    End If
    
    Get_CRC = Chr(XOR_Rst)
End Function

'Check Sum
Public Function Get_SUM(ByVal strTmp As String, ByVal MaxLen As Integer) As String
    Dim SumVal          As Variant
    Dim SumStr          As String
    Dim strByte()       As Byte
    Dim i               As Long
    
'    strByte = StrConv(strTmp, vbFromUnicode)
    strByte = strTmp
    For i = LBound(strByte) To UBound(strByte)
        SumVal = SumVal + strByte(i)
    Next i
    
    SumVal = SumVal Mod (256 ^ MaxLen)
        
    For i = 1 To MaxLen
        SumStr = Format(Hex(SumVal Mod 256), "0#") & SumStr
        SumVal = SumVal / (256 ^ i)
    Next i
    
    Get_SUM = SumStr
End Function

'AXSYM
Public Function CHK_SUM(ByVal strTmp As String, ByVal MaxLen As Integer) As String
    Dim SumVal          As Variant
    Dim strByte()       As Byte
    Dim i               As Long

    strByte = StrConv(strTmp, vbFromUnicode)

    For i = LBound(strByte) To UBound(strByte)
        SumVal = SumVal + strByte(i)
    Next i

    CHK_SUM = Right(Hex(SumVal), MaxLen)
End Function

Public Sub Delay(fTime As Single)
  Dim fStartTime As Single
  fStartTime = Timer
  Do While Timer < fStartTime + fTime
      DoEvents
  Loop
End Sub

