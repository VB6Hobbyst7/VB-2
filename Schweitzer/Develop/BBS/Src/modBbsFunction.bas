Attribute VB_Name = "modBbsFunction"
Option Explicit

Global gBloodRequestMusic As String
Global gblnEndSystem As Boolean
Global Const CD2_FileServer = "C232"    ' File Server Location

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                              (ByVal hwnd As Long, ByVal wMsg As Long, _
                              ByVal wParam As Long, ByVal lParam As Any) As Long

Public PrtLeft         As Long
Public LineSpace       As Long
Public LastLineYpos    As Long
Public Twidth          As Long
Public lngCurYPos      As Long

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



'Public Sub InitLvwHead(ByRef objLvw As ListView, ByVal strHead As String, _
'                       ByVal strSize As String)
'    Dim ii As Integer
'    Dim aryTitle() As String
'    Dim aryWidth() As String
'
'    aryTitle = Split(strHead, ",")
'    aryWidth = Split(strSize, ",")
'    If UBound(aryWidth) < UBound(aryTitle) Then
'        ReDim Preserve aryWidth(UBound(aryTitle))
'    End If
'    With objLvw
'        .ColumnHeaders.Clear
'        For ii = 0 To UBound(aryTitle)
'            If aryWidth(ii) = "" Then
'               aryWidth(ii) = "0"
'            End If
'            .ColumnHeaders.Add ii + 1, aryTitle(ii), aryTitle(ii), _
'                              (.Width \ (UBound(aryTitle) + 1)) + Val(aryWidth(ii)), vbLeftJustify
'        Next ii
'        .View = lvwReport            ' Report Style
'    End With
'End Sub

Public Sub Crystal_Print(ByVal CrystalNm As CrystalReport, ByVal strTmp As String, _
                            ByVal strFilePath As String, ByVal strRptPath As String)
    
    'CrystalNm:Crystal컨트롤 Name
    'strTmp: Record String(출력값)
    'strRptPath: Rpt파일 경로
    'strFilePath: text Fil 경로
    
    Dim intFNum As Integer
    
    intFNum = FreeFile
    Open strFilePath For Output As #intFNum
    Print #intFNum, strTmp
    Close #intFNum
    With CrystalNm
        .ReportFileName = strRptPath
        .RetrieveDataFiles
        .WindowState = crptMaximized
        .Action = 1
        .Reset
    End With
End Sub
Public Function SDA_String(ByVal SSN As String) As String
    Dim strTmp As String
    Dim strSEX As String
    Dim strAge As String
    Dim strDOB As String
    
    Dim strYY  As String
    Dim strMM  As String
    Dim strDD  As String
    
    strYY = Trim(Mid(SSN, 1, 2))
    strMM = Trim(Mid(SSN, 3, 2))
    strDD = Trim(Mid(SSN, 5, 2))
    
    If Val(strMM) < 1 Then strMM = "01"
    If Val(strMM) > 12 Then strMM = "12"
    If Val(strDD) < 1 Then strDD = "01"
    If Val(strDD) > 31 Then strDD = "31"
    
    
    On Error Resume Next
    
    If IsDate(strYY & "-" & strMM & "-" & strDD) = False Then
        strDD = "01"
    End If
    
    strSEX = "기타": strAge = "": strDOB = ""
    
    If SSN <> "" Then
        strTmp = Mid(SSN, 7, 1)
        Select Case strTmp
            Case "0": strSEX = "여": strDOB = "18" & strYY & "-" & strMM & "-" & strDD
            Case "1": strSEX = "남": strDOB = "19" & strYY & "-" & strMM & "-" & strDD
            Case "2": strSEX = "여": strDOB = "19" & strYY & "-" & strMM & "-" & strDD
            Case "3": strSEX = "남": strDOB = "20" & strYY & "-" & strMM & "-" & strDD
            Case "4": strSEX = "여": strDOB = "20" & strYY & "-" & strMM & "-" & strDD
            Case Else: strSEX = "남": strDOB = "19" & strYY & "-" & strMM & "-" & strDD
        End Select
        
        If Len(SSN) = 13 Then
            strAge = medFindAge(Replace(strDOB, "-", ""), "Y")
        Else
            strAge = ""
        End If
        SDA_String = strSEX & COL_DIV & strDOB & COL_DIV & strAge
    Else
        SDA_String = "" & COL_DIV & "" & COL_DIV & ""
    End If
End Function

Public Function Trim0(ByVal vData As String) As String
    Dim l As Long
    Dim i As Long
    
    l = Len(vData)
    For i = l To 1 Step -1
        If Asc(Mid(vData, i, 1)) = 0 Then
            vData = Mid(vData, 1, i - 1)
        End If
    Next i
    Trim0 = vData
End Function

Public Function GetBBS_Ptinfo(ByVal qPtid As String, _
                              Optional ByRef pSSN As String, Optional ByRef qPtnm As String, _
                              Optional ByRef qSex As String, Optional ByRef qDob As String)

    Dim Rs  As Recordset
    Dim objPt As clsPatient
    
    Set objPt = New clsPatient
    Set Rs = New Recordset
    Rs.Open objPt.GetSQLPt(qPtid), DBConn
    
    If Not Rs.EOF Then
        qPtnm = Rs.Fields("ptnm").value & ""
        qSex = Rs.Fields("sex").value & ""
        qDob = Rs.Fields("dob").value & ""
        pSSN = Rs.Fields("ssn").value & ""
    End If
    Set Rs = Nothing
    Set objPt = Nothing
End Function

