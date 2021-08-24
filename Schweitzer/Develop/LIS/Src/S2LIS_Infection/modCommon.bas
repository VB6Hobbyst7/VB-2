Attribute VB_Name = "modCommon"
Option Explicit


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                              (ByVal hwnd As Long, ByVal wMsg As Long, _
                              ByVal wParam As Long, ByVal lParam As Any) As Long
                              
'# medLockWindowUpdate
Private Declare Function LockWindowUpdate Lib "user32" _
                (ByVal hwndLock As Long) As Long
                              

Public ObjMyUser As New clsDSMLogOn

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

Public frmCount     As Long


                              
'*-----------------------------------------------------------------
'*  1. 기능 : SQL문장의 필드값을 셋팅시 "'"문자열을 조정한다
'*            필드값외에 문자열의 형태를 변형한다.
'*  2. Parameter : strValue - 필드(문자:varchar,char)값
'*               : optNum - 리턴문자열의 모양정의치
'*-----------------------------------------------------------------

Public Function DBS(ByVal strValue As String, Optional ByVal optNum As Long) As String
'String Conversion For Database INSERT,UPDATE

    strValue = Replace(strValue, "'", "''")
    If UCase(strValue) = "NULL" Then
        DBS = strValue
    Else
        DBS = "'" & CStr(strValue) & "'"
    End If
    If IsMissing(optNum) = False Then
        Select Case optNum
            Case 0
                DBS = DBS
            Case 1
                DBS = DBS & ","
            Case 2
                DBS = "=" & DBS
            Case 3
                DBS = "=" & DBS & ","
            Case Else
        End Select
    End If
   '
End Function

'*-----------------------------------------------------------------
'*  1. 기능 : SQL문장의 필드값을 셋팅시 수치형자료를 조정한다
'*            필드값외에 문자열의 형태를 변형한다.
'*  2. Parameter : strValue - 필드값(수치형)
'*               : optNum - 리턴수치형 자료의 모양정의치
'*-----------------------------------------------------------------

Public Function DBN(ByVal NumValue As String, Optional ByVal optNum As Long) As String
'Number Conversion For Database INSERT,UPDATE
    
    If UCase(NumValue) = "NULL" Then
        DBN = NumValue
    Else
        DBN = CStr(Val(NumValue))
    End If
    If DBN = "" Then DBN = "0"
    If IsMissing(optNum) = False Then
        Select Case optNum
            Case 0
                DBN = DBN
            Case 1
                DBN = DBN & ","
            Case 2
                DBN = " = " & DBN
            Case 3
                DBN = " = " & DBN & ","
            Case Else
        End Select
    End If
   '
End Function

'*-----------------------------------------------------------------
'*  1. 기능 : Delimiter로 구분하여 지정 위치의 String을 읽어온다.
'*            (Mumps의 $P()함수 이용하여 Data Read 하는 경우)
'*  2. 관련변수 :
'*  3. Parameter : strtext - Delimiter로 묶여있는 대상 문자열
'*                 intDepth - Delimiter Level (1 - 5)
'*                 intPosition - 선택 대상 문자열 위치
'*                 strDeli - Optional, 사용자정의의 구분자
'*  4. ReturnValue : 선택된 문자열
'*-----------------------------------------------------------------
Public Function medGetP(ByVal strText As String, _
                  ByVal intPosition As Integer, ByVal Delimiter As String) As String
Dim intPos1 As Integer, intPos2 As Integer, i As Integer

    intPos1 = 0: intPos2 = 0
    
    ' intPosition 인수가 1인 경우 For문 Skip
    For i = 1 To intPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, strText, Delimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    ' 해당 컬럼
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, strText, Delimiter)
    If intPos2 = 0 Then intPos2 = Len(strText) + 1
    
    medGetP = Mid$(strText, intPos1, intPos2 - intPos1)
    
    Exit Function
    
ReturnNull:
    medGetP = ""
    
End Function

'*-----------------------------------------------------------------
'*  1. 기능 : Delimiter로 구분된 첫번째 String을 읽어오고
'*            나머지 문자열을 그대로 남긴다.
'*  2. ReturnValue : 선택된 문자열
'*                   strText 자신은 Shift가 이루어진다.
'*-----------------------------------------------------------------
Public Function medShift(ByRef strText As String, ByVal strDeli As Variant) As String
Dim CNTA, CNTB As Integer
Dim Delimiter As String

    medShift = "": CNTA = 0: CNTB = 0
    
    CNTA = InStr(1, strText, strDeli)
    If CNTA = 0 Then
        medShift = strText
        strText = ""
        Exit Function
    End If
    
    medShift = Mid$(strText, 1, CNTA - Len(strDeli))
    strText = Mid$(strText, CNTA + Len(strDeli))
    
End Function

'*-----------------------------------------------------------------
'*  1. 기능 : 데이타가 많은 리스트 상에서 특정 내용(String)을 찾는 경우
'*  2. Parameter : lstList - 대상 리스트
'*                 strTmp  - Search할 String
'*  3. ReturnValue :
'*          원하는 문자열을 찾았을 경우에는 해당 Listindex를 리턴
'*          찾지 못했을 경우에는 근접 단어의 Listindex를 리턴
'*          근접 단어 조차도 찾지 못한 경우에는 -1을 리턴
'*-----------------------------------------------------------------
Public Function medListFind(ByVal lstList As Object, ByVal strTmp As String)
    
    medListFind = SendMessage(lstList.hwnd, &H18F, -1, strTmp)

End Function

'*-----------------------------------------------------------------
'*  1. 기능 : 리스트의 Redraw를 막아서 Add 또는 Scroll 속도를 향상한다.
'*  2. Parameter : hwndLock :lstMicCd.hwnd - Lock, &0 - Unlock
'*-----------------------------------------------------------------
Public Sub medLockWindowUpdate(ByVal hwndLock As Long)
    
    Call LockWindowUpdate(hwndLock)
  
End Sub

Public Sub SelFocus(ByRef Obj As Object)
    Obj.SelStart = 0
    Obj.SelLength = Len(Obj)
End Sub


Public Function Null2Space(Code As Variant) As String
    '
    ' Null String을 확인한다.
    ' Parameter    : Variant
    ' Return Value : Null이면 Space Value Else Parameter
    '
    Null2Space = Trim(IIf(IsNull(Code), "", Trim(Code)))
End Function

Public Function Null2Zero(Code As Variant) As Double
    '
    ' Null 을 확인한다.
    ' Parameter    : Variant
    ' Return Value : Null이면 0 Else Parameter
    '
    Null2Zero = Trim(IIf(IsNull(Code) Or Code = "", 0, Trim(Code)))
End Function


'** Report 모듈
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


Public Function InstallDir() As String
    Dim tmpDir As String
    
    tmpDir = GetSetting("Schweitzer2000", "InstallDir", "InstallDir", "")
    If tmpDir <> "" Then
        If Mid(tmpDir, Len(tmpDir), 1) = "\" Then
            tmpDir = tmpDir
        Else
            tmpDir = tmpDir & "\"
        End If
    End If

    InstallDir = tmpDir
End Function


'Public Sub BlobToFile(fld As ADODB.Field, ByVal FName As String, _
'                     Optional FieldSize As Long = -1, _
'                     Optional Threshold As Long = 1048576)
''
'' Assumes file does not exist '1048576
'' Data cannot exceed approx. 2Gb in size
''
'Dim F As Long, bData() As Byte, sData As String
'  F = FreeFile
'  Open FName For Binary As #F
'  Select Case fld.Type
'    Case adLongVarBinary
'      If FieldSize = -1 Then   ' blob field is of unknown size
'        WriteFromUnsizedBinary F, fld
'      Else                     ' blob field is of known size
'        If FieldSize > Threshold Then   ' very large actual data
'          WriteFromBinary F, fld, FieldSize
'        Else                            ' smallish actual data
'          bData = fld.Value
'          Put #F, , bData  ' PUT tacks on overhead if use fld.Value
'        End If
'      End If
'    Case adLongVarChar, adLongVarWChar
'      If FieldSize = -1 Then
'        WriteFromUnsizedText F, fld
'      Else
'        If FieldSize > Threshold Then
'          WriteFromText F, fld, FieldSize
'        Else
'          sData = fld.Value
'          Put #F, , sData  ' PUT tacks on overhead if use fld.Value
'        End If
'      End If
'  End Select
'  Close #F
'End Sub

'Public Sub WriteFromBinary(ByVal F As Long, fld As ADODB.Field, _
'                    ByVal FieldSize As Long)
'Dim Data() As Byte, BytesRead As Long
'  Do While FieldSize <> BytesRead
'    If FieldSize - BytesRead < P_BLOCK_SIZE Then
'      Data = fld.GetChunk(FieldSize - P_BLOCK_SIZE)
'      BytesRead = FieldSize
'    Else
'      Data = fld.GetChunk(P_BLOCK_SIZE)
'      BytesRead = BytesRead + P_BLOCK_SIZE
'    End If
'    Put #F, , Data
'  Loop
'End Sub
'
'Public Sub WriteFromUnsizedBinary(ByVal F As Long, fld As ADODB.Field)
'Dim Data() As Byte, Temp As Variant
'  Do
'    Temp = fld.GetChunk(P_BLOCK_SIZE)
'    If IsNull(Temp) Then Exit Do
'    Data = Temp
'    Put #F, , Data
'  Loop While LenB(Temp) = P_BLOCK_SIZE
'End Sub
'
'Public Sub WriteFromText(ByVal F As Long, fld As ADODB.Field, _
'                  ByVal FieldSize As Long)
'Dim Data As String, CharsRead As Long
'  Do While FieldSize <> CharsRead
'    If FieldSize - CharsRead < P_BLOCK_SIZE Then
'      Data = fld.GetChunk(FieldSize - P_BLOCK_SIZE)
'      CharsRead = FieldSize
'    Else
'      Data = fld.GetChunk(P_BLOCK_SIZE)
'      CharsRead = CharsRead + P_BLOCK_SIZE
'    End If
'    Put #F, , Data
'  Loop
'End Sub
'
'Public Sub WriteFromUnsizedText(ByVal F As Long, fld As ADODB.Field)
'Dim Data As String, Temp As Variant
'  Do
'    Temp = fld.GetChunk(P_BLOCK_SIZE)
'    If IsNull(Temp) Then Exit Do
'    Data = Temp
'    Put #F, , Data
'  Loop While Len(Temp) = P_BLOCK_SIZE
'End Sub
'
'Public Sub FileToBlob(ByVal FName As String, fld As ADODB.Field, _
'               Optional Threshold As Long = 1048576)
''
'' Assumes file exists
'' Assumes calling routine does the UPDATE
'' File cannot exceed approx. 2Gb in size
''
'Dim F As Long, Data() As Byte, FileSize As Long
'  F = FreeFile
'  Open FName For Binary As #F
'  FileSize = LOF(F)
'  Select Case fld.Type
'    Case adLongVarBinary
'      If FileSize > Threshold Then
'        ReadToBinary F, fld, FileSize
'      Else
'        Data = InputB(FileSize, F)
'        fld.Value = Data
'      End If
'    Case adLongVarChar, adLongVarWChar
'      If FileSize > Threshold Then
'        ReadToText F, fld, FileSize
'      Else
'        fld.Value = Input(FileSize, F)
'      End If
'  End Select
'  Close #F
'End Sub
'
'Public Sub ReadToBinary(ByVal F As Long, fld As ADODB.Field, _
'                 ByVal FileSize As Long)
'Dim Data() As Byte, BytesRead As Long
'  Do While FileSize <> BytesRead
'    If FileSize - BytesRead < P_BLOCK_SIZE Then
'      Data = InputB(FileSize - BytesRead, F)
'
'      BytesRead = FileSize
'    Else
'      Data = InputB(P_BLOCK_SIZE, F)
'      BytesRead = BytesRead + P_BLOCK_SIZE
'    End If
'    fld.AppendChunk Data
'  Loop
'End Sub
'
'Public Sub ReadToText(ByVal F As Long, fld As ADODB.Field, _
'               ByVal FileSize As Long)
'Dim Data As String, CharsRead As Long
'  Do While FileSize <> CharsRead
'    If FileSize - CharsRead < P_BLOCK_SIZE Then
'      Data = Input(FileSize - CharsRead, F)
'      CharsRead = FileSize
'    Else
'      Data = Input(P_BLOCK_SIZE, F)
'      CharsRead = CharsRead + P_BLOCK_SIZE
'    End If
'    fld.AppendChunk Data
'  Loop
'End Sub
'



