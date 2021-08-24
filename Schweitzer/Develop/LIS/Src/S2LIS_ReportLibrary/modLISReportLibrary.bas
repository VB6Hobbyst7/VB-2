Attribute VB_Name = "modLISReportLibrary"

Option Explicit

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Global gIsDeveloper As Boolean
Global gBuildingCd As String
Global gEmpId As String
Global gParentWhnd As Long

Global gWardid     As String
Global gWardNm     As String

Public PrtLeft         As Long
Public LineSpace       As Long
Public LastLineYpos    As Long
Public Twidth          As Long
Public lngCurYPos      As Long

Global gUsingInWardMenu As Boolean

'Global Const P_BLOCK_SIZE = 16384
'Global Const P_SLIDE_DB_PATH = "C:\Schweitzer\Image\"

Public frmCount     As Long

Public Function P_FIX(ByVal sStr As String, ByVal aBaseX As Single, ByVal aBaseY As Single, _
                          Optional ByVal SpcWidth As Single, _
                          Optional ByVal WAlign As String, _
                          Optional ByVal SpcHeight As Single, _
                          Optional ByVal HAlign As String, _
                          Optional ByVal StrFix As String, _
                          Optional ByVal StrFixRow As Integer) As Integer

    'strFixRow 줄간격
    'strFix
    Dim sglTmp!, sglLnHeight!
    Dim sData$(), sTmp$
    Dim iCnt%, iLineCnt%, iWidthLen%
    Dim iChk%

    '-[구분자 유무 체크 ]-
    iChk = InStr(1, sStr, "|")

    If iChk > 0 Then
        '-[ "|" (ascii:124) 구분으로 나누기 ]-
        iLineCnt = 1
        ReDim sData$(iLineCnt)
        For iCnt = 1 To Len(sStr)
            sTmp = Mid$(sStr, iCnt, 1)

            If sTmp = "|" Then
                iLineCnt = iLineCnt + 1
                ReDim Preserve sData$(iLineCnt)
            Else
                sData(iLineCnt) = sData(iLineCnt) & sTmp
            End If
        Next

        '-[ "|"구분으로 나눈것 출력 ]-
        If SpcHeight = 0 Then Exit Function
        sglLnHeight = SpcHeight / iLineCnt

        For iCnt = 1 To iLineCnt
            sglTmp = aBaseY + ((iCnt - 1) * sglLnHeight)
            Call P_FIX(sData(iCnt), aBaseX, sglTmp, SpcWidth, WAlign, sglLnHeight, HAlign)
        Next

    Else
        If SpcWidth >= Printer.TextWidth(sStr) Or _
           StrFix = "" Or SpcWidth = 0 Then
            '/* 가로 정렬 */
            Select Case WAlign
                Case "C", "c"  '/* 가운데 정렬*/
                    Printer.CurrentX = aBaseX + (SpcWidth - Printer.TextWidth(sStr)) / 2
                Case "R", "r"  '/* 오른쪽 정렬 */
                    Printer.CurrentX = aBaseX + (SpcWidth - Printer.TextWidth(sStr)) - 0.5
                Case Else      '/* 왼쪽 정렬 */
                    Printer.CurrentX = aBaseX + 0.5
            End Select

            '/* 세로 정렬 */
            Select Case HAlign
                Case "C", "c", "M", "m" '/* 중앙정렬 */
'                    Printer.CurrentY = lngCurYPos + (aBaseY - Printer.TextHeight(sStr)) / 2
                    Printer.CurrentY = aBaseY + (SpcHeight - Printer.TextHeight(sStr)) / 2
                Case "B", "b" '/* 아래정렬 */
                    Printer.CurrentY = aBaseY + (SpcHeight - Printer.TextHeight(sStr)) - 1
                Case Else     '/* 위쪽정렬 */
                    Printer.CurrentY = aBaseY + 1
            End Select
'            lngCurYPos = lngCurYPos + aBaseY

            Printer.Print sStr

        Else
            iWidthLen = (SpcWidth - 1) / Printer.TextWidth("A")

            Call Print_Fix(sStr, sData(), iLineCnt, iWidthLen)
            Select Case StrFix
                Case "W", "w" '/* Wordwrap */
                    '-[ 줄수가 지정되면 지정된 줄수만큼만 표시 ]-
                    If StrFixRow < iLineCnt And StrFixRow > 0 Then iLineCnt = StrFixRow - 1

                    For iCnt = 0 To iLineCnt
                        Printer.CurrentX = aBaseX + 0.5
                        Printer.CurrentY = aBaseY + Printer.TextHeight("A") * iCnt + 1
                        Printer.Print sData(iCnt)
                    Next
                    P_FIX = iLineCnt + 1
                Case Else '/* Prefix */
                    Call P_FIX(sData(0), aBaseX, aBaseY, SpcWidth, WAlign, SpcHeight, HAlign)
                    P_FIX = 1
            End Select

        End If

    End If
End Function

Public Sub Print_Fix(ByVal 문자 As String, _
                 ByRef 문자열() As String, _
                 ByRef LineCnt As Integer, _
                 ByVal aStrLenth As Integer)
    
    Dim iTextLenth As Integer
    Dim iCnt As Integer
    Dim sTmp$
    
    Dim iTextLine As Integer
    Dim iStringLenth As Integer
    Dim sStringBuffer As String
    Dim sTextBuffer() As String
    
    ReDim sTextBuffer(1) As String
    
    iTextLine = 0
    iStringLenth = 0
    iTextLenth = Len(문자)
    
    For iCnt = 1 To iTextLenth
        
        If Mid(문자, iCnt, 1) = "'" Then
            sTmp = """"
        Else
            sTmp = Mid(문자, iCnt, 1)
        End If
        
        Select Case Asc(sTmp)
            Case 13 ', 20, 10
                iTextLine = iTextLine + 1
                ReDim Preserve sTextBuffer(iTextLine) As String
            Case Is > 31
                iStringLenth = iStringLenth + 1
                
                If iStringLenth > aStrLenth Then
                    iTextLine = iTextLine + 1
                    ReDim Preserve sTextBuffer(iTextLine) As String
                    iStringLenth = 1
                End If
                
                sTextBuffer(iTextLine) = sTextBuffer(iTextLine) & sTmp
                
            Case Is < 0
                iStringLenth = iStringLenth + 2
                
                If iStringLenth > aStrLenth Then
                    iTextLine = iTextLine + 1
                    ReDim Preserve sTextBuffer(iTextLine) As String
                    iStringLenth = 2
                End If
                sTextBuffer(iTextLine) = sTextBuffer(iTextLine) & sTmp
            
        End Select
    Next iCnt

    ReDim 문자열(iTextLine) As String
    
    For iCnt = 0 To iTextLine
        문자열(iCnt) = sTextBuffer(iCnt)
    Next iCnt
    
    LineCnt = iTextLine
    
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
        Case Else      '/* 왼쪽 정렬 */
            Printer.CurrentX = aBaseX + 0.5
    End Select
    
    '/* 세로 정렬 */
    Select Case HAlign
        Case "C", "c", "M", "m" '/* 중앙정렬 */
            Printer.CurrentY = lngCurYPos + (aBaseY - Printer.TextHeight(sStr)) / 2
'                    Printer.CurrentY = aBaseY + (SpcHeight - Printer.TextHeight(sStr)) / 2
        Case "B", "b" '/* 아래정렬 */
            Printer.CurrentY = lngCurYPos + (aBaseY - Printer.TextHeight(sStr)) - 1
        Case Else     '/* 위쪽정렬 */
            Printer.CurrentY = lngCurYPos + 1
    End Select
    If blnLineAdd Then lngCurYPos = lngCurYPos + aBaseY
    
    Printer.Print sStr
    
End Function


Public Function Select_Print(sPrinterName As String)
    Dim X As Printer
'    If sPrinterName = "" Then Set Printer = Nothing: Exit Function
    For Each X In Printers
        If X.DeviceName = sPrinterName Then
        Set Printer = X
        End If
    Next
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

Public Sub FileToBlob(ByVal FName As String, fld As ADODB.Field, _
               Optional Threshold As Long = 1048576)
'
' Assumes file exists
' Assumes calling routine does the UPDATE
' File cannot exceed approx. 2Gb in size
'
Dim F As Long, Data() As Byte, FileSize As Long
  F = FreeFile
  Open FName For Binary As #F
  FileSize = LOF(F)
  Select Case fld.Type
    Case adLongVarBinary
      If FileSize > Threshold Then
        ReadToBinary F, fld, FileSize
      Else
        Data = InputB(FileSize, F)
        fld.Value = Data
      End If
    Case adLongVarChar, adLongVarWChar
      If FileSize > Threshold Then
        ReadToText F, fld, FileSize
      Else
        fld.Value = Input(FileSize, F)
      End If
  End Select
  Close #F
End Sub

Public Sub ReadToBinary(ByVal F As Long, fld As ADODB.Field, _
                 ByVal FileSize As Long)
Dim Data() As Byte, BytesRead As Long
  Do While FileSize <> BytesRead
    If FileSize - BytesRead < P_BLOCK_SIZE Then
      Data = InputB(FileSize - BytesRead, F)

      BytesRead = FileSize
    Else
      Data = InputB(P_BLOCK_SIZE, F)
      BytesRead = BytesRead + P_BLOCK_SIZE
    End If
    fld.AppendChunk Data
  Loop
End Sub

Public Sub ReadToText(ByVal F As Long, fld As ADODB.Field, _
               ByVal FileSize As Long)
Dim Data As String, CharsRead As Long
  Do While FileSize <> CharsRead
    If FileSize - CharsRead < P_BLOCK_SIZE Then
      Data = Input(FileSize - CharsRead, F)
      CharsRead = FileSize
    Else
      Data = Input(P_BLOCK_SIZE, F)
      CharsRead = CharsRead + P_BLOCK_SIZE
    End If
    fld.AppendChunk Data
  Loop
End Sub

Public Sub Print_WaterMark()
    Dim strAddress As String
    Dim PageNumber As Integer
    Dim dwLen As Long
    Dim strDate As String
    Dim strIP As String
    Dim strString As String
    
    SocketsInitialize
    strIP = GetTheIP
    SocketsCleanup
 
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    GetComputerName strString, dwLen
    strString = Left(strString, dwLen)
    strDate = Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    PageNumber = PageNumber + 1
    LastLineYpos = Printer.ScaleHeight - 90             '마지막라인Y위치
    
    Printer.FontSize = 9 'Printer.FontBold = True
    Printer.ForeColor = &HC0C0C0
    Call P_FIX("출력일시 : " & strDate & Space(5) & "출력IP : " & strIP & Space(5) & "출력ID : " & ObjMyUser.EmpLngNm & "(" & ObjMyUser.EmpId & ")", PrtLeft, LastLineYpos, 7500 - PrtLeft, "C", , "C")
    Printer.ForeColor = vbBlack
End Sub

