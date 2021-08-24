Attribute VB_Name = "modLISStatisticLibrary"
Option Explicit

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'-------------------------------------------------------'
'   근태관리에서 사용하는 API (by 이상대, 2002-11-07)
'-------------------------------------------------------'
'INI File에Data를 쓰는 API Function
'Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
'    ByVal lpApplicationName As String, _
'    ByVal lpKeyName As Any, _
'    ByVal lpString As Any, _
'    ByVal lpFileName As String _
') As Long

' INI File에서 Data를 읽는 API Function
'Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
'    ByVal lpApplicationName As String, _
'    ByVal lpKeyName As Any, _
'    ByVal lpDefault As String, _
'    ByVal lpReturnedString As String, _
'    ByVal nSize As Long, _
'    ByVal lpFileName As String _
') As Long


Global gIsDeveloper As Boolean
Global gBuildingCd  As String
Global gEmpId       As String

Public Const PrtLeft = 5      '시작위치(x좌표)
Public Const LineSpace = 6    '행사이의 간격(높이)
Public lngCurYPos As Long

Public Sub P_PrtSet()
    Dim ii As Integer
    
    Printer.Font = "굴림체"
    Printer.FontSize = 9
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = vbPRORPortrait '/* 좁게
    Printer.ScaleMode = vbMillimeters
    
End Sub
Public Function Print_Setting(ByVal sStr As String, _
                              ByVal aBaseX As Single, _
                              ByVal aBaseY As Single, _
                              Optional ByVal SpcWidth As Single, _
                              Optional ByVal WAlign As String, _
                              Optional ByVal HAlign As String, _
                              Optional ByVal blnLineAdd As Boolean = True) As Integer
                          
    '/* 가로 정렬 */
    Select Case WAlign
        Case "C", "c"  '/* 가운데 정렬*/
            Printer.CurrentX = aBaseX + (SpcWidth - Printer.TextWidth(sStr)) / 2
        Case "R", "r"  '/* 오른쪽 정렬 */
            Printer.CurrentX = aBaseX + (SpcWidth - Printer.TextWidth(sStr)) - 0.5
        Case "L", "l"  '/* 왼쪽 정렬 */
            Printer.CurrentX = aBaseX + 0.5
        Case Else      '/* 가운데 정렬*/
            Printer.CurrentX = aBaseX + (SpcWidth - Printer.TextWidth(sStr)) / 2
            'Printer.CurrentX = aBaseX + 0.5
    End Select
    
    '/* 세로 정렬 */
    Select Case HAlign
        Case "C", "c", "M", "m" '/* 중앙정렬 */
            Printer.CurrentY = lngCurYPos + (aBaseY - Printer.TextHeight(sStr)) / 2
'                    Printer.CurrentY = aBaseY + (SpcHeight - Printer.TextHeight(sStr)) / 2
        Case "B", "b" '/* 아래정렬 */
            Printer.CurrentY = lngCurYPos + (aBaseY - Printer.TextHeight(sStr)) - 1
        Case Else     '/* 위쪽정렬 */
            'Printer.CurrentY = lngCurYPos + 1
            Printer.CurrentY = lngCurYPos + (aBaseY - Printer.TextHeight(sStr)) / 2
    End Select
    If blnLineAdd Then lngCurYPos = lngCurYPos + aBaseY
    
    Printer.Print sStr
            
End Function

Public Function IsLastForm() As Boolean

    Dim i As Long
    Dim tmpFrm As Form
    
    i = 0
    IsLastForm = False
    
    For Each tmpFrm In Forms
        i = i + 1
    Next
    If i = 0 Then IsLastForm = True

End Function

Public Sub BlobToFile(fld As ADODB.Field, ByVal FName As String, _
                     Optional FieldSize As Long = -1, _
                     Optional Threshold As Long = 1048576)
'
' Assumes file does not exist '1048576
' Data cannot exceed approx. 2Gb in size
'
Dim F As Long, bData() As Byte, sData As String
  F = FreeFile
  Open FName For Binary As #F
  Select Case fld.Type
    Case adLongVarBinary
      If FieldSize = -1 Then   ' blob field is of unknown size
        WriteFromUnsizedBinary F, fld
      Else                     ' blob field is of known size
        If FieldSize > Threshold Then   ' very large actual data
          WriteFromBinary F, fld, FieldSize
        Else                            ' smallish actual data
          bData = fld.Value
          Put #F, , bData  ' PUT tacks on overhead if use fld.Value
        End If
      End If
    Case adLongVarChar, adLongVarWChar
      If FieldSize = -1 Then
        WriteFromUnsizedText F, fld
      Else
        If FieldSize > Threshold Then
          WriteFromText F, fld, FieldSize
        Else
          sData = fld.Value
          Put #F, , sData  ' PUT tacks on overhead if use fld.Value
        End If
      End If
  End Select
  Close #F
End Sub

Public Sub WriteFromBinary(ByVal F As Long, fld As ADODB.Field, _
                    ByVal FieldSize As Long)
Dim Data() As Byte, BytesRead As Long
  Do While FieldSize <> BytesRead
    If FieldSize - BytesRead < P_BLOCK_SIZE Then
      Data = fld.GetChunk(FieldSize - P_BLOCK_SIZE)
      BytesRead = FieldSize
    Else
      Data = fld.GetChunk(P_BLOCK_SIZE)
      BytesRead = BytesRead + P_BLOCK_SIZE
    End If
    Put #F, , Data
  Loop
End Sub

Public Sub WriteFromUnsizedBinary(ByVal F As Long, fld As ADODB.Field)
Dim Data() As Byte, Temp As Variant
  Do
    Temp = fld.GetChunk(P_BLOCK_SIZE)
    If IsNull(Temp) Then Exit Do
    Data = Temp
    Put #F, , Data
  Loop While LenB(Temp) = P_BLOCK_SIZE
End Sub

Public Sub WriteFromText(ByVal F As Long, fld As ADODB.Field, _
                  ByVal FieldSize As Long)
Dim Data As String, CharsRead As Long
  Do While FieldSize <> CharsRead
    If FieldSize - CharsRead < P_BLOCK_SIZE Then
      Data = fld.GetChunk(FieldSize - P_BLOCK_SIZE)
      CharsRead = FieldSize
    Else
      Data = fld.GetChunk(P_BLOCK_SIZE)
      CharsRead = CharsRead + P_BLOCK_SIZE
    End If
    Put #F, , Data
  Loop
End Sub

Public Sub WriteFromUnsizedText(ByVal F As Long, fld As ADODB.Field)
Dim Data As String, Temp As Variant
  Do
    Temp = fld.GetChunk(P_BLOCK_SIZE)
    If IsNull(Temp) Then Exit Do
    Data = Temp
    Put #F, , Data
  Loop While Len(Temp) = P_BLOCK_SIZE
End Sub
