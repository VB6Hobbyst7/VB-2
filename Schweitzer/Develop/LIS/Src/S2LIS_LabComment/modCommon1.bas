Attribute VB_Name = "modCommon1"

'*-----------------------------------------------------------------
'*  1. ���� : �ݺ����� ����Ÿ ó���� ������ ��ƾ
'*  2. �ֿ� Object :
'*  3. �ֿ� Procedure :
'*  4. �ֿ� Algorithm :
'*  5. Calling Form :
'*  6. Called By :
'*  7. Reference Form :
'*  8. ������/������ : 1998.2.12 by Jeong,Kwangseok
'*  9. ������/������ : 1998.6.12 by �� �� ��
'* 10. Ư����� :
'*-----------------------------------------------------------------

Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'#
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'#
Public Const LB_SETTABSTOPS = &H192
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                              (ByVal hwnd As Long, ByVal wMsg As Long, _
                              ByVal wParam As Long, ByVal lParam As Any) As Long
Declare Function SendMessage1 Lib "user32" Alias "SendMessageA" _
                              (ByVal hwnd As Long, ByVal wMsg As Long, _
                              ByVal wParam As Long, lParam As Any) As Long


Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
                    (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, _
                    ByVal lpsz2 As String) As Long
'#
Declare Function GetListIndex Lib "user32" Alias "SendMessageA" _
                              (ByVal hwnd As Long, ByVal wMsg As Long, _
                              ByVal wParam As Long, ByVal lParam As String) As Long
'#
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (lpString1 As Any, _
                                                                                      lpString2 As Any) As Long
Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" _
                    (ByVal lpDevMode As Long, ByVal dwFlags As Long) As Long
Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32
Public Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
'#
Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Public Const HWND_NOTOPMOST = -2   'Not Always top
Public Const HWND_TOPMOST = -1  'Always top
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
'#
Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
'#
Public Const IME_HANGUL = &H1
Public Const IME_ENGLISH = &H0
Public Const IME_NONE = &H0
Declare Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As Long) As Long
Declare Function ImmSetConversionStatus Lib "imm32.dll" _
                (ByVal hIMC As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
'#
Declare Function LockWindowUpdate Lib "user32" _
                (ByVal hwndLock As Long) As Long
'#
Declare Function FlashWindow Lib "user32" _
                (ByVal hwnd As Long, ByVal bInvert As Long) As Long
'#
Declare Function ShowScrollBar Lib "user32" _
                              (ByVal hwnd As Long, ByVal sFlag As Integer, _
                              ByVal sBool As Boolean) As Long
'#
Declare Function ExtFloodFill Lib "gdi32" _
                (ByVal hDC As Long, ByVal X As Long, _
                  ByVal Y As Long, ByVal crColor As Long, _
                  ByVal wFillType As Long) As Long
'#
Declare Function FillRgn Lib "gdi32" _
                (ByVal hDC As Long, ByVal hRgn As Long, _
                  ByVal hBrush As Long) As Long
'#
Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Global rectTmp As RECT
Declare Function FillRect Lib "user32" _
                (ByVal hDC As Long, lpRect As RECT, _
                  ByVal hBrush As Long) As Long
'#
Declare Function SetBkColor Lib "gdi32" _
                (ByVal hDC As Long, ByVal crColor As Long) As Long
'#
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
'#
Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
'#
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWNORMAL = 1

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
            (ByVal hwnd As Long, ByVal lpOperation As String, _
             ByVal lpFile As String, ByVal lpParameters As String, _
            ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'#
Declare Function SetCursorPos Lib "user32.dll" (ByVal X As Long, ByVal Y As Long) As Long
Declare Function ClipCursor Lib "user32.dll" (lpRect As RECT) As Long
Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetDesktopWindow Lib "user32.dll" () As Long
'#
Declare Function DrawIconEx& Lib "user32" (ByVal hDC As Long, ByVal xleft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long)
Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, _
ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long

'# Screen Lock ����
Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long

'# Homepage ����
Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" _
    (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As _
    String) As Long


Global hDC As Long
Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Declare Function timeSetEvent Lib "winmm.dll" (ByVal uDelay As Long, ByVal uResolution As Long, ByVal lpFunction As Long, ByVal dwUser As Long, ByVal uFlags As Long) As Long
Declare Function timeKillEvent Lib "winmm.dll" (ByVal uID As Long) As Long


Enum Language
   lngKorean = 0
   lngEnglish = 1
End Enum
Global medLan As Language
Global clsLan As Object


Public Function Get_SysDate()

    Dim tmpRs As DrRecordSet
    
    Set tmpRs = OpenRecordSet("select " & CS_SybaseDate & " as Today")
    If tmpRs.EOF Then
        Get_SysDate = Format(Now, CS_DateDbFormat)
    Else
        Get_SysDate = tmpRs.Fields("Today").Value
    End If
    
    tmpRs.RsClose
    Set tmpRs = Nothing

End Function

Public Function Get_SysTime()

    Dim tmpRs As DrRecordSet
    
    Set tmpRs = OpenRecordSet("select " & CS_SybaseTime & " as Time")
    If tmpRs.EOF Then
        Get_SysDate = Format(Now, CS_TimeDbFormat)
    Else
        Get_SysTime = tmpRs.Fields("Time").Value
    End If
    
    tmpRs.RsClose
    Set tmpRs = Nothing

End Function

'% ��ǻ�� �̸� ��������..
Public Function medGetComNm()

   Dim sBuffer$, nSize As Long, rtn As Long
   sBuffer = String(256, Chr(0))
   rtn = GetComputerName(sBuffer$, Len(sBuffer))
   medGetComNm = sBuffer
   
End Function

'*-----------------------------------------------------------------
'*  1. ��� : ������� �μ�����ŭ �ݺ��ؼ� �Ҹ�����.
'*  2. ���ú��� :
'*  3. Parameter : intCnt (Beep Count)
'*-----------------------------------------------------------------
Public Sub medBeep(ByVal intCNT As Integer)
Dim I As Integer
    
    For I = 1 To intCNT
        Call Beep
    Next I
    
End Sub

'*-----------------------------------------------------------------
'*  1. ��� : M���� Delimiter������ ���� �̸� ���ǵ� �� ��ȯ
'*  2. ���ú��� :
'*  3. Parameter : intDepth - Delimiter Level (1 - 5)
'*  4. ReturnValue : �� ���ǵ� Character
'*-----------------------------------------------------------------
Public Function medDelimiter(ByVal intDepth As Integer) As String
Const DelimiterCode As Integer = 20

    If intDepth < 1 Or intDepth > 5 Then intDepth = 1
    
    medDelimiter = Chr$(DelimiterCode + intDepth)
    
End Function

'*-----------------------------------------------------------------
'*  1. ��� : �� string�� �� Level���� Delimiter�� �̿��Ͽ� Join
'*  2. ���ú��� :
'*  3. Parameter : intDepth - Delimiter Level (1 - 9)
'*                 strValue() - ������ ���Ѿ��� �����µ��� ���
'*                              null�� ""�� �ڵ� ġȯ�Ͽ� ó��
'*  4. ReturnValue : Delimiter �����Ͽ� Join�� �ϳ��� string
'*-----------------------------------------------------------------
Public Function medJoin(ByVal intDepth As Integer, ParamArray strValue() As Variant)
Dim I As Integer, intArrayCount As Integer
Dim Delimiter As String * 1

    intArrayCount = UBound(strValue)
    If intArrayCount < 0 Then medJoin = "": Exit Function
    
    Delimiter = medDelimiter(intDepth)
    
    If IsMissing(strValue(0)) Then
        medJoin = ""
    Else
        medJoin = strValue(0)
    End If
    
    For I = 1 To intArrayCount
        If IsMissing(strValue(I)) Then strValue(I) = ""
        medJoin = medJoin & Delimiter & strValue(I)
    Next I
    
End Function


'*-----------------------------------------------------------------
'*  1. ��� : �� string�� BLBX�� Key���Ŀ� �°� ����
'*  2. ���ú��� :
'*  3. Parameter :    strValue() - ������ ���Ѿ��� �����µ��� ���
'*                            null�� ""�� �ڵ� ġȯ�Ͽ� ó��
'*  4. ReturnValue : Delimiter �����Ͽ� Join�� �ϳ��� string
'*-----------------------------------------------------------------
Public Function medSetKey(ParamArray strValue() As Variant)
Dim I As Integer, intArrayCount As Integer
Dim Delimiter As String * 1

    intArrayCount = UBound(strValue)
    If intArrayCount < 0 Then medSetKey = "": Exit Function
    
    Delimiter = ";"
    
    If IsMissing(strValue(0)) Then
        medSetKey = ""
    Else
        medSetKey = strValue(0)
    End If
    
    For I = 1 To intArrayCount
        If IsMissing(strValue(I)) Then strValue(I) = ""
        medSetKey = medSetKey & Delimiter & strValue(I)
    Next I
    
End Function

'*-----------------------------------------------------------------
'*  1. ��� : �ش� SpreadSheet�� ���� ��ġ�� ����Ÿ�� Write �Ѵ�.
'*  2. ���ú��� :
'*  3. Parameter : objTable - ���� ������ ���̺� ����
'*                 Col,Row  - Col,Row ��ġ ����
'*                 StoreData, FontSize - �����Ϳ� ���� ũ��(Optional)
'*  4. ReturnValue :
'*-----------------------------------------------------------------
Public Sub medWriteTableXY(ByRef objTable As Object, _
                        ByVal Col As Integer, ByVal Row As Integer, _
                        ByVal StoreData As String, Optional ByVal FontSize As Variant)

    objTable.Col = Col
    objTable.Row = Row
        
    If IsMissing(FontSize) = False Then
        objTable.FontSize = FontSize
    End If
    objTable.Value = StoreData
        
End Sub

'*-----------------------------------------------------------------
'*  1. ��� : �ش� SpreadSheet�� ���� ��ġ�� ����Ÿ�� Read �Ѵ�.
'*  2. ���ú��� :
'*  3. Parameter : objTable - ���� ������ ���̺� ����
'*                 Col1,Col2,Row1,Row2 - ���� ����
'*  4. ReturnValue :
'*-----------------------------------------------------------------
Public Function medReadTableXY(ByRef objTable As Object, _
                    ByVal Col As Integer, ByVal Row As Integer) As String

    objTable.Col = Col
    objTable.Row = Row
    
    medReadTableXY = objTable.Value
    
End Function


'*-----------------------------------------------------------------
'*  1. ��� : �ش� SpreadSheet�� ���� �������� ����Ÿ�� Write �Ѵ�.
'*  2. ���ú��� :
'*  3. Parameter : objTable - ���� ������ ���̺� ����
'*                 Col1,Col2,Row1,Row2 - ���� ����
'*                 StoreData, FontSize - �����Ϳ� ���� ũ��
'*                 Flag - (Missing : ���̺���ü, 1 : �κ�)
'*  4. ReturnValue :
'*-----------------------------------------------------------------
Public Sub medWriteTable(ByRef objTable As Object, _
                        ByVal Col1 As Integer, ByVal COL2 As Integer, _
                        ByVal Row1 As Integer, ByVal Row2 As Integer, _
                        ByVal StoreData As String, Optional ByVal FontSize As Variant, _
                        Optional ByVal flag As Variant)

    If IsMissing(flag) Then
        objTable.MaxRows = Row2
        objTable.MaxCols = COL2
    Else
        If objTable.MaxRows < Row2 Then objTable.MaxRows = Row2
        If objTable.MaxCols < COL2 Then objTable.MaxCols = COL2
    End If
    
    objTable.Col = Col1: objTable.COL2 = COL2
    objTable.Row = Row1: objTable.Row2 = Row2
    
    objTable.BlockMode = True
    If IsMissing(FontSize) = False Then
        objTable.FontSize = FontSize
    End If
    objTable.ClipValue = StoreData
    objTable.BlockMode = False
    
End Sub

'*-----------------------------------------------------------------
'*  1. ��� : �ش� SpreadSheet�� ���� ������ ����Ÿ�� Read �Ѵ�.
'*  2. ���ú��� :
'*  3. Parameter : objTable - ���� ������ ���̺� ����
'*                 Col1,Col2,Row1,Row2 - ���� ����
'*  4. ReturnValue : Colume�� ASC(9), Row�� ASC(13)+ASC(10)����
'*                   ���е� �ϳ��� String���� ����
'*-----------------------------------------------------------------
Public Function medReadTable(ByRef objTable As Object, _
                    ByVal Col1 As Integer, ByVal COL2 As Integer, _
                    ByVal Row1 As Integer, ByVal Row2 As Integer) As String

    objTable.Col = Col1: objTable.COL2 = COL2
    objTable.Row = Row1: objTable.Row2 = Row2
    
    medReadTable = ""
    
    objTable.BlockMode = True
    medReadTable = objTable.ClipValue
    objTable.BlockMode = False

End Function

'*-----------------------------------------------------------------
'*  1. ��� : Listbox �Ǵ� Combobox�� �ش� �ڷḦ Setting �Ѵ�.
'*  2. ���ú��� :
'*  3. Parameter : objList - ���� ������ ����Ʈ�ڽ� �Ǵ� �޺��ڽ� ����
'*                 intDepth - Delimiter Level (1 - 9)
'*                 strText - ������ ITEM�� ������ �ǵ���Ÿ
'*                 intFlag - 0 (�ʱ�ȭ�� ���ۼ�)
'*                           1 (���� ����Ÿ�� �����ϸ鼭 �ڿ� ADD)
'*  4. ReturnValue :
'*-----------------------------------------------------------------
Public Sub medWriteList(ByRef objList As Object, ByVal intDepth As Integer, _
                        ByVal strText As String, ByVal intFlag As Integer)
Dim intPos1 As Integer, intPos2 As Integer
Dim intLength As Integer, Delimiter As String * 1

    If intFlag = 0 Then objList.Clear       ' �ʱ�ȭ ����
    
    intLength = Len(strText)
    If intLength <= 0 Then Exit Sub         ' ����Ÿ ����
    
    Delimiter = medDelimiter(intDepth)
    intPos1 = 0: intPos2 = 0
    
    Do While intPos2 < intLength
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, strText, Delimiter)
       If intPos2 = 0 Then Exit Do
       objList.AddItem Mid$(strText, intPos1, intPos2 - intPos1)
    Loop
    objList.AddItem Mid(strText, intPos1, Len(strText) - intPos1 + 1)
   
    objList.Visible = True
    objList.ZOrder 0
        
End Sub

Public Sub medPrintList(ByRef objList As Object, ByVal intDepth As Integer, _
                        ByVal strText As String, ByVal intFlag As Integer)
Dim strAddWord As String, strTempChar As String * 1
Dim I As Integer, intLength As Integer, Delimiter As String * 1

    If intFlag = 0 Then objList.Clear       ' �ʱ�ȭ ����
    
    intLength = Len(strText)
    If intLength <= 0 Then Exit Sub         ' ����Ÿ ����
    
    Delimiter = medDelimiter(intDepth)
    
    strAddWord = ""
    For I = 1 To intLength
        strTempChar = Mid$(strText, I, 1)
        If strTempChar = Delimiter Then
            objList.AddItem strAddWord
            strAddWord = ""
        Else
            strAddWord = strAddWord & strTempChar
        End If
    Next I
    
    objList.AddItem strAddWord
   
    objList.Visible = True
    objList.ZOrder 0
        
End Sub

'*-----------------------------------------------------------------
'*  1. ��� : Printer Queue�� ��ġ�� �����Ͽ� Data�� ������
'*  2. ���ú��� :
'*  3. Parameter : XPos,YPos - ���� ������ X,Y ��ǥ�� ����
'*                 PrintText, FontSize - ��� Text�� ũ��
'*  4. ReturnValue :
'*-----------------------------------------------------------------
Public Sub PTextXY(ByVal XPos As String, ByVal YPos As String, _
           ByVal PrintText As String, ByVal FontSize As Integer)

    Printer.CurrentX = XPos
    Printer.CurrentY = YPos
    
    Printer.FontSize = FontSize
    Printer.Print PrintText
    
End Sub
'*-----------------------------------------------------------------
'*  1. ��� : medPrinterOpen() - Printer Port�� Open�Ѵ�.
'*               medPrinterClose() - Printer Port�� Close�Ѵ�.
'*               medPrint() - String�� Open�Ǿ��ִ� Port�� Print�Ѵ�.
'*  2. ���ú��� :
'*  3. Parameter :   intFileNo - �����ִ� �����Ʈ�� ��ȣ
'*                           strData - ����� ���ڿ�
'*-----------------------------------------------------------------
Function medPrinterOpen() As Integer
Dim FileNo As Integer

    On Error GoTo ErrorHandler
    
    medPrinterOpen = 0
    FileNo = FreeFile
    medPrinterOpen = FileNo
    Open "LPT1" For Output As #FileNo
    Exit Function
    
ErrorHandler:
    MsgBox "There was a problem openning your printer port."
    medPrinterOpen = -1

End Function

Function medPrinterClose(ByVal FileNo As Integer) As Integer
    
    On Error GoTo ErrorHandler
    Close #FileNo
    medPrinterClose = 0
    Exit Function
    
ErrorHandler:
    MsgBox "There was a problem closing your printer port."
    medPrinterClose = -1

End Function

Function medPrint(ByVal intFileNo As Integer, ByVal strData As String) As Integer
    
    On Error GoTo ErrorHandler
    Print #intFileNo, strData
    medPrint = 0
    Exit Function
    
ErrorHandler:
    MsgBox "There was a problem printing to your printer."
    medPrint = -1

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
Dim intPos1 As Integer, intPos2 As Integer, I As Integer

    intPos1 = 0: intPos2 = 0
    
    ' intPosition �μ��� 1�� ��� For�� Skip
    For I = 1 To intPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, strText, Delimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next I
    
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
'*  1. ��� : Delimiter�� �����Ͽ� ���� ��ġ�� String�� ġȯ�Ѵ�.
'*            (Mumps�� $P()�Լ� �̿��Ͽ� Data Store �ϴ� ���)
'*  2. ���ú��� :
'*  3. Parameter : strtext - Delimiter�� �����ִ� ��� ���ڿ�
'*                 intDepth - Delimiter Level (1 - 5)
'*                 intPosition - ���� ��� ���ڿ� ��ġ
'*                 strWord - ġȯ �� ���ڿ�
'*  4. ReturnValue : ġȯ�� ���ڿ�
'*-----------------------------------------------------------------
Public Function medSetP(ByVal strText As String, ByVal intDepth As Integer, _
                  ByVal intPosition As Integer, ByVal strWord As String, _
                  Optional ByVal strDeli As Variant) As String
Dim intPos1 As Integer, intPos2 As Integer, I As Integer
Dim strHead As String, strTail As String
Dim Delimiter As String

    If intPosition <= 0 Then GoTo ReturnMe ' ����Ÿ ����
    
    If IsMissing(strDeli) Then
        Delimiter = medDelimiter(intDepth)
    Else
        Delimiter = strDeli
    End If
    intPos1 = 0: intPos2 = 0
    
    ' intPosition �μ��� 1�� ��� For�� Skip
    For I = 1 To intPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, strText, Delimiter)
       If intPos2 = 0 Then GoTo AddDelimiter
    Next I
    
    ' �ش� �÷�
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, strText, Delimiter)
    If intPos2 = 0 Then intPos2 = Len(strText) + 1
    
    'strHead = Mid(strText.Text, 1, intPos1 - 1)
    'strTail = Right$(strText.Text, Len(strText) - intPos2 + 1)
    'medSetP = strHead & strWord & strTail
    
    Exit Function
    
AddDelimiter:
    medSetP = strText & String(intPosition - I, Delimiter) & strWord
    Exit Function
    
ReturnMe:
    medSetP = strText
    
End Function

'*-----------------------------------------------------------------
'*  1. ��� : BLBX�� Delimiter�� ���е� ù��° String�� �о����
'*               ������ ���ڿ��� �״�� �����.
'*  2. ���ú��� :
'*  3. Parameter : strText - Delimiter�� �����ִ� ��� ���ڿ�
'*  4. ReturnValue : ���õ� ���ڿ�
'*                   strText �ڽ��� Shift�� �̷������.
'*-----------------------------------------------------------------
Public Function medGetKey(ByRef strText As String, ByVal strDeli As String) As String
Dim CNTA, CNTB As Integer
    
    medGetKey = "": CNTA = 0: CNTB = 0
    
    CNTA = InStr(1, strText, strDeli)
    If CNTA = 0 Then
        medGetKey = strText
        strText = ""
        Exit Function
    End If
    
    medGetKey = Mid$(strText, 1, CNTA - 1)
    strText = Mid$(strText, CNTA + 1)

End Function

'*-----------------------------------------------------------------
'*  1. ��� : Delimiter�� ���е� ù��° String�� �о����
'*            ������ ���ڿ��� �״�� �����.
'*  2. ���ú��� :
'*  3. ReturnValue : ���õ� ���ڿ�
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
'*  1. ��� : message code�� �Ѱ��־� �ش� message�� �޴´�.
'*  2. ���ú��� :
'*  3. Parameter : strCode - Message Code
'*  4. ReturnValue : �޼���
'*-----------------------------------------------------------------
Public Function medGetMsg(ByVal strCode As String) As String
Dim strFromM As String, strCheck As String

    If Len(strCode) < 6 Then
        medGetMsg = ""
        Exit Function
    End If
    
    'medGetMsg = medMVB.MExe("medUtil", "GETMSG", strCode, 1)

End Function

'*-----------------------------------------------------------------
'*  1. ��� : ������Ϸ� ���� ���
'*  2. ���ú��� :
'*  3. Parameter : strBirthDate: �������(yyyymmdd)
'*                 strType:���̸� ��,��,�� �� ��� �������� ������ ������
'*                     ( Y,M,D )
'*                 strSysDate : ����� ������ �Ǵ� ����(yyyymmdd)
'*                     strSysDate�� Optional ������ ������ ���ڷ� ���̸� ���
'*  4. ReturnValue : ���� ����(Year����)
'*-----------------------------------------------------------------
Function medFindAge(ByVal strBirthDate As String, ByVal strAgeType As String, _
             Optional ByVal strSysDate) As String
Dim strFormatBirth As String
Dim strFormatSys As String

    strFormatBirth = Format(Format(strBirthDate, "####/##/##"), "yy-mm-dd")
    
    If IsMissing(strSysDate) Then
        strFormatSys = Format(Now, "yy-mm-dd")
    Else
        strFormatSys = Format(Format(strSysDate, "####/##/##"), "yy-mm-dd")
    End If
    
    Select Case UCase(strAgeType)
    Case "Y":        '���
        medFindAge = DateDiff("yyyy", strFormatBirth, strFormatSys)
    Case "M":        '����
        medFindAge = DateDiff("m", strFormatBirth, strFormatSys)
    Case "D":        '�Ϸ�
        medFindAge = DateDiff("d", strFormatBirth, strFormatSys)
    End Select
    
End Function

'*-----------------------------------------------------------------
'*  1. ��� : Server�� ��¥��  Return
'*  2. ���ú��� :
'*  3. Parameter : intDateLenth:Return�� ���ϴ� Date Format��
'*                 Length
'*  4. ReturnValue : Server�� ��¥
'*-----------------------------------------------------------------
Function medSysDate(Optional ByVal intDateLength) As String
    
    'medSysDate = medMVB.MExe("%AJOU", "HDATE", "", 1)
    If IsMissing(intDateLength) Then Exit Function
    If (IsMissing(intDateLength) = False) And (intDateLength <= 8) Then
        'medSysDate = Right(medSysDate, intDateLength)
    End If
End Function

'*-----------------------------------------------------------------
'*  1. ��� : Server�� �ð���  Return
'*  2. ���ú��� :
'*  3. Parameter : intTimeLength:Return�� ���ϴ� Time Format��
'*                 Length
'*  4. ReturnValue : Server�� �ð�
'*-----------------------------------------------------------------
Function medSysTime(Optional ByVal intTimeLength) As String
    
    'medSysTime = medMVB.MExe("%AJOU", "HTIME", "", 1)
    If IsMissing(intTimeLength) Then Exit Function
    If (IsMissing(intTimeLength) = False) And (intTimeLength <= 6) Then
        'medSysTime = Left(medSysTime, intTimeLength)
    End If
End Function

'*-----------------------------------------------------------------
'*  1. ��� : Client�� ���糯¥�� �ð���  Return
'*  2. ���ú��� :
'*  3. Parameter :
'*  4. ReturnValue : Client�� ���糯¥�� �ð�
'*-----------------------------------------------------------------
Function medThisTime() As String
Dim yy As String, mm As String, DD As String, TM As String
    
    DD = Format(Date, "YY/MM/DD")
    TM = Format(Time, "HH:MM:SS")
    medThisTime = DD & "   " & TM
    
End Function


'*-----------------------------------------------------------------
'*  1. ��� : �ش� Data�� ���ڷμ� �� Data�� ��ȿ���� Check
'*  2. ���ú��� :
'*  3. Parameter : strDate : Check�ϰ��� �ϴ� Data
'*                 yyyymmdd(8�ڸ�) ���ĸ� ����
'*  4. ReturnValue : True or False
'*-----------------------------------------------------------------
Public Function medDateChk(ByVal strDate As String) As Boolean
Dim strFormatData As String
    
    If Len(strDate) <> 8 Then
        medDateChk = False
        Exit Function
    End If
    
    medDateChk = IsDate(strFormatData)

End Function

'*-----------------------------------------------------------------
'*  1. ��� : �ش� Date�� ������ Return
'*  2. ���ú��� :
'*  3. Parameter : strDate   - Check�ϰ��� �ϴ� Date
'*                                          Date���Ŀ� �´� Data�� ������
'*                         intOption - Return�� ���ϴ� ������ ����(1,2,3,4)
'*                                          ex)1:Sunday, 2:Sun, 3:�Ͽ���, 4:��
'*  4. ReturnValue : ����(����, �ѱ�)
'*-----------------------------------------------------------------
Public Function medWeekday(ByVal strDate As Date, _
               ByVal intOption As Integer) As String
Dim aryPattern As Variant
Dim aryWeekday As Variant

    aryWeekday = Array("��", "��", "ȭ", "��", "��", "��", "��")
    aryPattern = Array("ddd", "dddd")

    If IsDate(strDate) = False Then   '�μ����� date���Ŀ���
        medWeekday = ""
        Exit Function
    End If

    If intOption < 3 Then            '����
        medWeekday = Format(strDate, aryPattern(intOption - 1))
    Else
        medWeekday = aryWeekday(Weekday(strDate) - 1)
        If intOption = 4 Then        '�ѱ���ü
            medWeekday = medWeekday + "����"
        End If
    End If
    
End Function

'*-----------------------------------------------------------------
'*  1. ��� : Keyascii ���� Upper Case������ ��ȯ��
'*  2. ���ú��� :
'*  3. Parameter : intKeyAscii :Keypress���� �߻��ϴ� Keyascii��
'*  4. ReturnValue : Alphabet�빮�ڿ� �ش��ϴ� Ascii��
'*  5. ��뿹�� :
'*               Private Sub Text2_KeyPress(KeyAscii As Integer)
'*                   KeyAscii = medToUCase(KeyAscii)
'*               End Sub
'*-----------------------------------------------------------------
Public Function medUCase(ByVal intKeyAscii As Integer) As Integer

    medUCase = Asc(UCase(Chr(intKeyAscii)))
    
End Function

'*-----------------------------------------------------------------
'*  1. ��� : �ش� Spreat Sheet�� Data�� Clear
'*  2. ���ú��� :
'*  3. Parameter : objTable : Clear�� Table(���� Form����)
'*                 blnCol : ColHeader�� Clear�� ������
'*                 blnRow : RowHeader�� Clear�� ������
'*                 ex) Form1.Table1
'*  4. ReturnValue :
'*-----------------------------------------------------------------
Public Sub medClearTable(ByVal objTable As Object, _
                        Optional ByVal blnCol As Boolean, _
                        Optional ByVal blnRow As Boolean)
Dim ii As Integer
    
    objTable.Col = 1
    objTable.COL2 = objTable.MaxCols
    objTable.Row = 1
    objTable.Row2 = objTable.MaxRows
    objTable.BlockMode = True
    objTable.Action = ActionClearText   ' Clear the data from the cells
    objTable.BlockMode = False
    
    ' Turn block mode off
    '��)Header�� ��� Value�� Null�� �ָ� Default���� ��Ÿ���Ƿ�
    '   ������ Space�� Insert�մϴ�.
    'Column Header Clear
    If blnCol = True Then
        For ii = 1 To objTable.MaxCols
            objTable.Row = 0
            objTable.Col = ii
            objTable.Value = " "
        Next ii
    End If
    'Row Header Clear
    If blnRow = True Then
        For ii = 1 To objTable.MaxRows
            objTable.Col = 0
            objTable.Row = ii
            objTable.Value = " "
        Next ii
    End If
End Sub



'*-----------------------------------------------------------------
'*  1. ��� : �����ð����� �ð��� ������ų ���
'*  2. Parameter : Interval : ������ų �ð�
'*-----------------------------------------------------------------
Sub medSleep(ByVal Interval As Long)

    Sleep (Interval)
    
End Sub


'*-----------------------------------------------------------------
'*  1. ��� : ListBox�� Horizental Scroll Bar�� �������� �ش�.
'*              (Default ListBox�� �ش�)
'*  2. ���ú��� :
'*  3. Parameter : lstList : �ش� ListBox Control
'*-----------------------------------------------------------------
Sub medHorScrol(ByVal lstList As Object)
   
    SendMessage lstList.hwnd, &H194, 3 * (lstList.WIDTH / Screen.TwipsPerPixelX)

End Sub


'*-----------------------------------------------------------------
'*  1. ��� : medplay Mode�� �����Ѵ�.
'*  2. ���ú��� :
'*  3. Parameter : ������ Mode (1:640*480, 2:800*600, 3:1024*768)
'*-----------------------------------------------------------------
Sub medSetMode(ByVal intMode As Integer)
Dim XX As Integer, yy As Integer
    
    On Error GoTo ErrorHandler

    XX = Choose(intMode, 640, 800, 1024)
    yy = Choose(intMode, 480, 600, 768)
    
    If SetDisplayMode(XX, yy, -1) = 0 Then
        MsgBox CStr(XX) & "*" & CStr(yy) & " ���� ����Ǿ����ϴ�.", vbInformation
    Else
        MsgBox "���÷��� ��� ������ �����߽��ϴ�.", vbInformation
    End If
    Exit Sub
    
ErrorHandler:
        MsgBox "���÷��� ��� ������ �����߽��ϴ�.", vbInformation
        
End Sub

'*-----------------------------------------------------------------
'*  1. ��� : medplay Mode�� �����Ѵ�.(Called by medSetMode())
'*-----------------------------------------------------------------
Public Function SetDisplayMode(WIDTH As Integer, HEIGHT As Integer, COLOR As Integer) As Long
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Const DM_BITSPERPEL = &H40000
Dim NewDevMode As DEVMODE
Dim pDevMode As Long
    With NewDevMode
        .dmSize = 122
        If COLOR = -1 Then
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
        Else
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
        End If
        .dmPelsWidth = WIDTH
        .dmPelsHeight = HEIGHT
        If COLOR <> -1 Then
            .dmBitsPerPel = COLOR
        End If
    End With
    pDevMode = lstrcpy(NewDevMode, NewDevMode)
    SetDisplayMode = ChangeDisplaySettings(pDevMode, 0)
End Function

'*-----------------------------------------------------------------
'*  1. ��� : Win32 ȯ�� �Ͽ��� �޸��� ���¸� �˾Ƴ���.
'*  3. Parameter : TotMem : ��ü ������ �޸�
'*                         AvailMem : ��� ������ ������ �޸�
'*-----------------------------------------------------------------
Public Sub medSysMem(ByRef TotMem As Long, ByRef AvailMem As Long)
Dim ms As MEMORYSTATUS

    ms.dwLength = Len(ms)
    GlobalMemoryStatus ms
    'MsgBox "��ü ������ �޸� : " & ms.dwTotalPhys & vbCRLF & _
                  "��� ������ ������ �޸� : " & ms.dwAvailPhys
    TotMem = ms.dwTotalPhys
    AvailMem = ms.dwAvailPhys

End Sub


'*-----------------------------------------------------------------
'*  1. ��� : �ѱ��Է��� Set�Ѵ�.
'*  3. Parameter : Scr :�Է��� ������ Control
'*-----------------------------------------------------------------
Public Sub medHanOn(Src As Object)
Dim hIME As Long

  hIME = ImmGetContext(Src.hwnd)
  ImmSetConversionStatus hIME, IME_HANGUL, IME_NONE
  Src.SetFocus

End Sub

'*-----------------------------------------------------------------
'*  1. ��� : �����Է��� Set�Ѵ�.
'*  3. Parameter : Scr :�Է��� ������ Control
'*-----------------------------------------------------------------
Public Sub medEngOn(Src As Object)
Dim hIME As Long

  hIME = ImmGetContext(Src.hwnd)
  ImmSetConversionStatus hIME, IME_ENGLISH, IME_NONE
  Src.SetFocus
  
End Sub


'*-----------------------------------------------------------------
'*  1. ��� : �ش����� �׻� ���� ���ְ� �Ѵ�.
'*  3. Parameter : frmForm - �ش� ��
'*                         OnOff - 0 : ����, 1 : ����
'*-----------------------------------------------------------------
Sub medAlwaysOn(ByVal frmForm As Form, ByVal OnOff As Integer)
Dim hWndMode As Integer

    hWndMode = Choose(OnOff + 1, -2, -1)
    SetWindowPos frmForm.hwnd, hWndMode, 0, 0, 10, 10, _
                                        SWP_NOMOVE Or SWP_NOSIZE
'    SetWindowPos frmForm.hwnd, HWND_TOPMOST, 0, 0, 10, 10, _
                                        SWP_NOMOVE Or SWP_NOSIZE

End Sub
 

'*-----------------------------------------------------------------
'*  1. ��� : ����Ÿ�� ���� ����Ʈ �󿡼� Ư�� ����(String)�� ã�� ���
'*  2. ���ú��� :
'*  3. Parameter : lstList : ��� ����Ʈ
'*                         strTmp : Search�� String
'*  4. ReturnValue :
'*          ���ϴ� ���ڿ��� ã���� ��쿡�� �ش� Listindex�� ����
'*          ã�� ������ ��쿡�� ���� �ܾ��� Listindex�� ����
'*          ���� �ܾ� ������ ã�� ���� ��쿡�� -1�� ����
'*-----------------------------------------------------------------
Function medListFind(ByVal lstList As Object, ByVal strTmp As String)
    
    medListFind = SendMessage(lstList.hwnd, &H18F, -1, strTmp)

End Function
    
Function medComboFind(ByVal cboCombo As Object, ByVal strTmp As String)
    
    Dim I As Integer
    
    With cboCombo
      For I = 0 To .ListCount - 1
         If .List(I) Like (strTmp & "*") Then
            medComboFind = I
            Exit Function
         End If
      Next
   End With
   medComboFind = -1

End Function
    
    
'*-----------------------------------------------------------------
'*  1. ��� : ���ٸ� ������·� �����ش�.
'*  2. ���ú��� :
'*  3. Parameter : objToolBar : Toolbar Control
'*  4. ReturnValue :
'*-----------------------------------------------------------------
Sub medCoolbar(ByVal objToolBar As Object)
Dim lRet As Long, lHandle As Long

    lHandle = FindWindowEx(objToolBar.hwnd, "ToolbarWindow32", vbNullString)
    lRet = SendMessage(lHandle, 1081&)
    lRet = lRet Or 2048&
    SendMessage lHandle, 1080&, 0, lRet
    objToolBar.Refresh

End Sub
    
'*-----------------------------------------------------------------
'*  1. ��� : ���� ĸ�ǹٸ� �����̰� �Ѵ�.
'*  2. ���ú��� :
'*  3. Parameter : frmForm - �ش� ��
'*                         blnOpt - True�϶� Flash, False�϶� ��������
'*  4. ReturnValue : �ش� �����찡 Ȱ��ȭ�Ǿ��ٸ� True�� ��ȯ,
'*                            �� ���� ���� False�� ��ȯ.
'*-----------------------------------------------------------------
Function medFlashWnd(ByVal frmForm As Form, ByVal blnOpt As Boolean)

    medFlashWnd = FlashWindow(frmForm.hwnd, blnOpt)

End Function
    
'*-----------------------------------------------------------------
'*  1. ��� : �������� ���̺� Lock�� �����ϰų� �����Ѵ�.
'*  2. ���ú��� :
'*  3. Parameter : objTable - ���� ������ ���̺� ����
'*                         blnOpt - True�϶� Lock, False�϶� Unlock
'*-----------------------------------------------------------------
Function medTableLock(ByVal objTable As Object, ByVal blnOpt As Boolean)

    objTable.Row = -1
    objTable.Col = -1
    objTable.Lock = blnOpt

End Function

'*-----------------------------------------------------------------
'*  1. ��� : �������� ���̺��� �� �Ǵ� ���� �������� �����Ѵ�.
'*  2. ���ú��� :
'*  3. Parameter : objTable - ���� ������ ���̺� ����
'*                         Col1,Col2,Row1,Row2 - ���� ����
'*                         SortBy - ���ı���(1:��Column, 2:��Row)
'*                         KeyOrder - ����Key��  Option�� 2���� �迭
'*                  �� :        KeyOrder(1,1) = N -  ù��° Key(N ��° �� �Ǵ� ��)
'*                               KeyOrder(1,2) = 1  : ��������, 2 : ��������
'*                               KeyOrder(2,1) = N  - �ι�° Key(N ��° �� �Ǵ� ��)
'*                               KeyOrder(2,2) = 1  : ��������, 2 : ��������
'*-----------------------------------------------------------------
Function medTableSort(ByVal objTable As Object, _
                        ByVal Col1 As Integer, ByVal COL2 As Integer, _
                        ByVal Row1 As Integer, ByVal Row2 As Integer, _
                        ByVal SortBy As String, ByRef KeyOrder() As Variant)
Dim I As Integer

    objTable.Col = Col1: objTable.COL2 = COL2
    objTable.Row = Row1: objTable.Row2 = Row2
    objTable.BlockMode = True
    
    objTable.SortBy = SortBy
    
    For I = 1 To UBound(KeyOrder)
        objTable.SortKey(I) = KeyOrder(I, 1)
        objTable.SortKeyOrder(I) = KeyOrder(I, 2)
    Next I
    
    objTable.Action = 25
    objTable.BlockMode = False

End Function


'*-----------------------------------------------------------------
'*  1. ��� : ����Ʈ�ڽ��� 2�� �̻��� �÷��� ����Ÿ�� ���÷��� �Ҷ�,
'*               �� �ణ�� ������ ����ϰ� ���ش�.
'*  2. ���ú��� :
'*  3. Parameter : objList - ���� ������ ����Ʈ ����
'*                         iColCount - �÷��� ����
'*                         iColLen - �÷��� ����
'*                                (��)   iColLen(0) = 80    ' 80/4 = 20 Character
'*                                         iColLen(1) = 160  ' 160/4 = 40 Character
'*                                         iColLen(2) = 240  ' 240/4 = 60 Character
'*-----------------------------------------------------------------
Sub medListAlign(objList As Object, iColCount As Long, ParamArray iColLen() As Variant)
Dim Ret As Long, I As Integer
Dim iTab() As Long, iCnt As Integer
        
    On Error GoTo ErrorHandler
                
    ReDim iTab(0 To iColCount - 1)
    For I = 0 To iColCount - 1
        If IsMissing(iColLen(I)) Then
            iTab(I) = 20   'Default Length
        Else
            iTab(I) = iColLen(I)
        End If
    Next I
    
    Ret = SendMessage(objList.hwnd, LB_SETTABSTOPS, iColCount, iTab(0))
    Exit Sub
    
ErrorHandler:
    MsgBox "Error!!!", vbOKOnly, "Message..."

End Sub


'*-----------------------------------------------------------------
'*  1. ��� : �ش� ���Ͽ� ����Ǿ� �ִ� ���α׷��� ȣ���Ͽ� �����Ѵ�.
'*  2. ���ú��� :
'*  3. Parameter : frmForm  - ���� ��
'*                         FileName - �ش� ���ϸ�
'*-----------------------------------------------------------------
Sub medShell(ByVal frmForm As Form, ByVal FileName As String)
    
    ShellExecute frmForm.hwnd, "Open", FileName, vbNullString, vbNullString, SW_SHOWDEFAULT

End Sub


Sub Dither(vObj As Object)
    Dim intLoop As Integer
      vObj.DrawStyle = vbInsideSolid
      vObj.DrawMode = vbCopyPen
      vObj.ScaleMode = vbPixels
      vObj.DrawWidth = 4
      vObj.ScaleWidth = 100
      vObj.ScaleHeight = 255
      '--------------------------------------------------
      ' �Ķ���(0, 0, 255)���� ����������(0, 0, 0)����
      ' ���������� ĥ�� ������. ���� �����θ� ĥ�Ѵٴ�
      ' ������ �ִ�. �� ����� �ٲ��...
      '--------------------------------------------------
      For intLoop = 0 To 255
         vObj.Line (0, intLoop)-(100, intLoop - 1), RGB(intLoop, intLoop, intLoop), B
      Next intLoop
      
End Sub

'*-----------------------------------------------------------------
'*  1. ��� : ���ڿ� ���� Ư�� string�� �ٸ� string���� ��ġ�Ѵ�.
'*  2. ���ú��� :
'*  3. Parameter : strOrigin - ��� ���ڿ�
'*                       transFrom - �ٲ� ���ڿ�
'*                       transTo - �� ���ڿ�
'*-----------------------------------------------------------------
Function medTR(ByVal strOrigin As String, ByVal transFrom As String, transTo As String)
Dim I As Integer
Dim intLen As Integer

    I = 1
    intLen = Len(transFrom)
    Do While I <= Len(strOrigin)
        If Mid(strOrigin, I, intLen) = transFrom Then
            strOrigin = Mid(strOrigin, 1, I - 1) & transTo & Mid(strOrigin, I + intLen)
        End If
        I = I + 1
    Loop
    medTR = strOrigin
            
End Function

Sub Dither1(vObj As Object)

    Dim intLoop As Integer
      
    vObj.DrawStyle = vbInsideSolid
    vObj.DrawMode = vbCopyPen
    vObj.ScaleMode = vbPixels
    vObj.DrawWidth = 4
    vObj.ScaleWidth = 255
    vObj.ScaleHeight = 100
    
    '--------------------------------------------------
    ' �Ķ���(0, 0, 255)���� ����������(0, 0, 0)����
    ' ���������� ĥ�� ������. ���� �����θ� ĥ�Ѵٴ�
    ' ������ �ִ�. �� ����� �ٲ��...
    '--------------------------------------------------
    For intLoop = 0 To 255
        If intLoop <= 127 Then
            vObj.Line (intLoop, 0)-(intLoop - 1, 100), RGB(intLoop * 0.76, intLoop * 0.76, intLoop * 2), B
        Else
            vObj.Line (intLoop, 0)-(intLoop - 1, 100), RGB(intLoop * 0.76, intLoop * 0.76, 255 - (0.5 * (intLoop - 127))), B
            'vObj.Line (230, 0)-(230 - 1, 100), RGB(230 * 0.76, 230 * 0.76, 255 - (0.5 * (230 - 127))), B
        End If
    Next intLoop
    'For intLoop = 256 To 510
    '    vObj.Line (intLoop, 0)-(intLoop - 1, 100), RGB(193, 193, 510 - intLoop + 193), B
   ' Next intLoop
    
End Sub


Function medGetMessage(ByVal MsgKey As Integer)
   Select Case MsgKey
   Case CONNECT_ERROR:
      medGetMessage = "Database ������ �ȉ���ϴ�. ����Ƿ� ���� �ٶ��ϴ�."
   End Select
End Function


Public Function DBStr(ByVal strValue As String, _
   Optional ByVal optNum As Variant) As String
'String Conversion For Database INSERT,UPDATE
   DBStr = "'" & CStr(strValue) & "'"
   If IsMissing(optNum) = False Then
      Select Case optNum
         Case 1
            DBStr = DBStr & ","
         Case 2
            DBStr = "=" & DBStr
         Case 3
            DBStr = "=" & DBStr & ","
         Case Else
      End Select
   End If
   '
End Function


Public Function DBNum(ByVal NumValue As String, _
   Optional ByVal optNum As Variant) As String
'Number Conversion For Database INSERT,UPDATE
   DBNum = CStr(NumValue)
   If DBNum = "" Then DBNum = "null"
   If IsMissing(optNum) = False Then
      Select Case optNum
         Case 1
            DBNum = DBNum & ","
         Case 2
            DBNum = " = " & DBNum
         Case 3
            DBNum = " = " & DBNum & ","
         Case Else
      End Select
   End If
   '
End Function


Public Sub UnloadForms(ByVal MyForm As Object)

   Dim tmpForm As Form
   
   For Each tmpForm In Forms
      With tmpForm
         If .Name <> "medMain" And .Name <> "medSplash" And .Name <> MyForm.Name Then
            Unload tmpForm
         End If
      End With
   Next

End Sub

