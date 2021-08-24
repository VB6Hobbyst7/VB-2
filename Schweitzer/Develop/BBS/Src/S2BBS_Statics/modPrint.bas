Attribute VB_Name = "modPrint"
Option Explicit

Public Const PrtLeft = 10      '시작위치(x좌표)
Public Const PrtTop = 20      '시작위치(y좌표)
Public Const LineSpace = 6    '행사이의 간격(높이)
Public lngCurYPos As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Sub P_PrtSet()
    Dim ii As Integer
    
    Printer.Font = "굴림체"
    Printer.FontSize = 10
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = vbPRORPortrait '/* 좁게
    Printer.ScaleMode = vbMillimeters
    
End Sub

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
        Case "L", "l"
            Printer.CurrentX = aBaseX + 0.5
        Case Else      '/* 왼쪽 정렬 */
            Printer.CurrentX = aBaseX + (SpcWidth - Printer.TextWidth(sStr)) / 2
    End Select
    
    '/* 세로 정렬 */
    Select Case HAlign
        Case "C", "c", "M", "m" '/* 중앙정렬 */
            Printer.CurrentY = lngCurYPos + (aBaseY - Printer.TextHeight(sStr)) / 2
'                    Printer.CurrentY = aBaseY + (SpcHeight - Printer.TextHeight(sStr)) / 2
        Case "B", "b" '/* 아래정렬 */
            Printer.CurrentY = lngCurYPos + (aBaseY - Printer.TextHeight(sStr)) - 1
        Case Else     '/* 위쪽정렬 */
            Printer.CurrentY = lngCurYPos + (aBaseY - Printer.TextHeight(sStr)) / 2
    End Select
    If blnLineAdd Then lngCurYPos = lngCurYPos + aBaseY
    
    Printer.Print sStr
            
End Function
