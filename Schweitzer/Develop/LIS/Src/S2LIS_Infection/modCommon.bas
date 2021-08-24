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
'*  1. ��� : SQL������ �ʵ尪�� ���ý� "'"���ڿ��� �����Ѵ�
'*            �ʵ尪�ܿ� ���ڿ��� ���¸� �����Ѵ�.
'*  2. Parameter : strValue - �ʵ�(����:varchar,char)��
'*               : optNum - ���Ϲ��ڿ��� �������ġ
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
'*  1. ��� : SQL������ �ʵ尪�� ���ý� ��ġ���ڷḦ �����Ѵ�
'*            �ʵ尪�ܿ� ���ڿ��� ���¸� �����Ѵ�.
'*  2. Parameter : strValue - �ʵ尪(��ġ��)
'*               : optNum - ���ϼ�ġ�� �ڷ��� �������ġ
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
'*  1. ��� : Delimiter�� �����Ͽ� ���� ��ġ�� String�� �о�´�.
'*            (Mumps�� $P()�Լ� �̿��Ͽ� Data Read �ϴ� ���)
'*  2. ���ú��� :
'*  3. Parameter : strtext - Delimiter�� �����ִ� ��� ���ڿ�
'*                 intDepth - Delimiter Level (1 - 5)
'*                 intPosition - ���� ��� ���ڿ� ��ġ
'*                 strDeli - Optional, ����������� ������
'*  4. ReturnValue : ���õ� ���ڿ�
'*-----------------------------------------------------------------
Public Function medGetP(ByVal strText As String, _
                  ByVal intPosition As Integer, ByVal Delimiter As String) As String
Dim intPos1 As Integer, intPos2 As Integer, i As Integer

    intPos1 = 0: intPos2 = 0
    
    ' intPosition �μ��� 1�� ��� For�� Skip
    For i = 1 To intPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, strText, Delimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    ' �ش� �÷�
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, strText, Delimiter)
    If intPos2 = 0 Then intPos2 = Len(strText) + 1
    
    medGetP = Mid$(strText, intPos1, intPos2 - intPos1)
    
    Exit Function
    
ReturnNull:
    medGetP = ""
    
End Function

'*-----------------------------------------------------------------
'*  1. ��� : Delimiter�� ���е� ù��° String�� �о����
'*            ������ ���ڿ��� �״�� �����.
'*  2. ReturnValue : ���õ� ���ڿ�
'*                   strText �ڽ��� Shift�� �̷������.
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
'*  1. ��� : ����Ÿ�� ���� ����Ʈ �󿡼� Ư�� ����(String)�� ã�� ���
'*  2. Parameter : lstList - ��� ����Ʈ
'*                 strTmp  - Search�� String
'*  3. ReturnValue :
'*          ���ϴ� ���ڿ��� ã���� ��쿡�� �ش� Listindex�� ����
'*          ã�� ������ ��쿡�� ���� �ܾ��� Listindex�� ����
'*          ���� �ܾ� ������ ã�� ���� ��쿡�� -1�� ����
'*-----------------------------------------------------------------
Public Function medListFind(ByVal lstList As Object, ByVal strTmp As String)
    
    medListFind = SendMessage(lstList.hwnd, &H18F, -1, strTmp)

End Function

'*-----------------------------------------------------------------
'*  1. ��� : ����Ʈ�� Redraw�� ���Ƽ� Add �Ǵ� Scroll �ӵ��� ����Ѵ�.
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
    ' Null String�� Ȯ���Ѵ�.
    ' Parameter    : Variant
    ' Return Value : Null�̸� Space Value Else Parameter
    '
    Null2Space = Trim(IIf(IsNull(Code), "", Trim(Code)))
End Function

Public Function Null2Zero(Code As Variant) As Double
    '
    ' Null �� Ȯ���Ѵ�.
    ' Parameter    : Variant
    ' Return Value : Null�̸� 0 Else Parameter
    '
    Null2Zero = Trim(IIf(IsNull(Code) Or Code = "", 0, Trim(Code)))
End Function


'** Report ���
Public Function P_FIX(ByVal sStr As String, ByVal aBaseX As Single, ByVal aBaseY As Single, _
                          Optional ByVal SpcWidth As Single, _
                          Optional ByVal WAlign As String, _
                          Optional ByVal SpcHeight As Single, _
                          Optional ByVal HAlign As String, _
                          Optional ByVal StrFix As String, _
                          Optional ByVal StrFixRow As Integer) As Integer

    'strFixRow �ٰ���
    'strFix
    Dim sglTmp!, sglLnHeight!
    Dim sData$(), sTmp$
    Dim iCnt%, iLineCnt%, iWidthLen%
    Dim iChk%

    '-[������ ���� üũ ]-
    iChk = InStr(1, sStr, "|")

    If iChk > 0 Then
        '-[ "|" (ascii:124) �������� ������ ]-
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

        '-[ "|"�������� ������ ��� ]-
        If SpcHeight = 0 Then Exit Function
        sglLnHeight = SpcHeight / iLineCnt

        For iCnt = 1 To iLineCnt
            sglTmp = aBaseY + ((iCnt - 1) * sglLnHeight)
            Call P_FIX(sData(iCnt), aBaseX, sglTmp, SpcWidth, WAlign, sglLnHeight, HAlign)
        Next

    Else
        If SpcWidth >= Printer.TextWidth(sStr) Or _
           StrFix = "" Or SpcWidth = 0 Then
            '/* ���� ���� */
            Select Case WAlign
                Case "C", "c"  '/* ��� ����*/
                    Printer.CurrentX = aBaseX + (SpcWidth - Printer.TextWidth(sStr)) / 2
                Case "R", "r"  '/* ������ ���� */
                    Printer.CurrentX = aBaseX + (SpcWidth - Printer.TextWidth(sStr)) - 0.5
                Case Else      '/* ���� ���� */
                    Printer.CurrentX = aBaseX + 0.5
            End Select

            '/* ���� ���� */
            Select Case HAlign
                Case "C", "c", "M", "m" '/* �߾����� */
'                    Printer.CurrentY = lngCurYPos + (aBaseY - Printer.TextHeight(sStr)) / 2
                    Printer.CurrentY = aBaseY + (SpcHeight - Printer.TextHeight(sStr)) / 2
                Case "B", "b" '/* �Ʒ����� */
                    Printer.CurrentY = aBaseY + (SpcHeight - Printer.TextHeight(sStr)) - 1
                Case Else     '/* �������� */
                    Printer.CurrentY = aBaseY + 1
            End Select
'            lngCurYPos = lngCurYPos + aBaseY

            Printer.Print sStr

        Else
            iWidthLen = (SpcWidth - 1) / Printer.TextWidth("A")

            Call Print_Fix(sStr, sData(), iLineCnt, iWidthLen)
            Select Case StrFix
                Case "W", "w" '/* Wordwrap */
                    '-[ �ټ��� �����Ǹ� ������ �ټ���ŭ�� ǥ�� ]-
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

Public Sub Print_Fix(ByVal ���� As String, _
                 ByRef ���ڿ�() As String, _
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
    iTextLenth = Len(����)
    
    For iCnt = 1 To iTextLenth
        
        If Mid(����, iCnt, 1) = "'" Then
            sTmp = """"
        Else
            sTmp = Mid(����, iCnt, 1)
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

    ReDim ���ڿ�(iTextLine) As String
    
    For iCnt = 0 To iTextLine
        ���ڿ�(iCnt) = sTextBuffer(iCnt)
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
                          
    '/* ���� ���� */
    Select Case WAlign
        Case "C", "c"  '/* ��� ����*/
            Printer.CurrentX = aBaseX + (SpcWidth - Printer.TextWidth(sStr)) / 2
        Case "R", "r"  '/* ������ ���� */
            Printer.CurrentX = aBaseX + (SpcWidth - Printer.TextWidth(sStr)) - 0.5
        Case Else      '/* ���� ���� */
            Printer.CurrentX = aBaseX + 0.5
    End Select
    
    '/* ���� ���� */
    Select Case HAlign
        Case "C", "c", "M", "m" '/* �߾����� */
            Printer.CurrentY = lngCurYPos + (aBaseY - Printer.TextHeight(sStr)) / 2
'                    Printer.CurrentY = aBaseY + (SpcHeight - Printer.TextHeight(sStr)) / 2
        Case "B", "b" '/* �Ʒ����� */
            Printer.CurrentY = lngCurYPos + (aBaseY - Printer.TextHeight(sStr)) - 1
        Case Else     '/* �������� */
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



