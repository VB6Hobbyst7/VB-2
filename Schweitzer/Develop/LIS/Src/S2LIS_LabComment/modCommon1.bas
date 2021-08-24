Attribute VB_Name = "modCommon1"

'*-----------------------------------------------------------------
'*  1. 목적 : 반복적인 데이타 처리의 공통사용 루틴
'*  2. 주요 Object :
'*  3. 주요 Procedure :
'*  4. 주요 Algorithm :
'*  5. Calling Form :
'*  6. Called By :
'*  7. Reference Form :
'*  8. 개발일/개발자 : 1998.2.12 by Jeong,Kwangseok
'*  9. 수정일/수정자 : 1998.6.12 by 김 미 경
'* 10. 특기사항 :
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

'# Screen Lock 관련
Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long

'# Homepage 띄우기
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

'% 컴퓨터 이름 가져오기..
Public Function medGetComNm()

   Dim sBuffer$, nSize As Long, rtn As Long
   sBuffer = String(256, Chr(0))
   rtn = GetComputerName(sBuffer$, Len(sBuffer))
   medGetComNm = sBuffer
   
End Function

'*-----------------------------------------------------------------
'*  1. 기능 : 경고음을 인수값만큼 반복해서 소리낸다.
'*  2. 관련변수 :
'*  3. Parameter : intCnt (Beep Count)
'*-----------------------------------------------------------------
Public Sub medBeep(ByVal intCNT As Integer)
Dim I As Integer
    
    For I = 1 To intCNT
        Call Beep
    Next I
    
End Sub

'*-----------------------------------------------------------------
'*  1. 기능 : M과의 Delimiter공유를 위해 미리 정의된 값 반환
'*  2. 관련변수 :
'*  3. Parameter : intDepth - Delimiter Level (1 - 5)
'*  4. ReturnValue : 기 정의된 Character
'*-----------------------------------------------------------------
Public Function medDelimiter(ByVal intDepth As Integer) As String
Const DelimiterCode As Integer = 20

    If intDepth < 1 Or intDepth > 5 Then intDepth = 1
    
    medDelimiter = Chr$(DelimiterCode + intDepth)
    
End Function

'*-----------------------------------------------------------------
'*  1. 기능 : 각 string을 각 Level단위 Delimiter를 이용하여 Join
'*  2. 관련변수 :
'*  3. Parameter : intDepth - Delimiter Level (1 - 9)
'*                 strValue() - 갯수에 제한없이 들어오는데로 사용
'*                              null은 ""로 자동 치환하여 처리
'*  4. ReturnValue : Delimiter 구분하여 Join된 하나의 string
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
'*  1. 기능 : 각 string을 BLBX의 Key형식에 맞게 연결
'*  2. 관련변수 :
'*  3. Parameter :    strValue() - 갯수에 제한없이 들어오는데로 사용
'*                            null은 ""로 자동 치환하여 처리
'*  4. ReturnValue : Delimiter 구분하여 Join된 하나의 string
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
'*  1. 기능 : 해당 SpreadSheet의 지정 위치에 데이타를 Write 한다.
'*  2. 관련변수 :
'*  3. Parameter : objTable - 현재 폼내의 테이블 설정
'*                 Col,Row  - Col,Row 위치 설정
'*                 StoreData, FontSize - 데이터와 글자 크기(Optional)
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
'*  1. 기능 : 해당 SpreadSheet의 지정 위치의 데이타를 Read 한다.
'*  2. 관련변수 :
'*  3. Parameter : objTable - 현재 폼내의 테이블 설정
'*                 Col1,Col2,Row1,Row2 - 범위 설정
'*  4. ReturnValue :
'*-----------------------------------------------------------------
Public Function medReadTableXY(ByRef objTable As Object, _
                    ByVal Col As Integer, ByVal Row As Integer) As String

    objTable.Col = Col
    objTable.Row = Row
    
    medReadTableXY = objTable.Value
    
End Function


'*-----------------------------------------------------------------
'*  1. 기능 : 해당 SpreadSheet의 지정 범위내에 데이타를 Write 한다.
'*  2. 관련변수 :
'*  3. Parameter : objTable - 현재 폼내의 테이블 설정
'*                 Col1,Col2,Row1,Row2 - 범위 설정
'*                 StoreData, FontSize - 데이터와 글자 크기
'*                 Flag - (Missing : 테이블전체, 1 : 부분)
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
'*  1. 기능 : 해당 SpreadSheet의 지정 범위의 데이타를 Read 한다.
'*  2. 관련변수 :
'*  3. Parameter : objTable - 현재 폼내의 테이블 설정
'*                 Col1,Col2,Row1,Row2 - 범위 설정
'*  4. ReturnValue : Colume은 ASC(9), Row는 ASC(13)+ASC(10)으로
'*                   구분된 하나의 String으로 리턴
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
'*  1. 기능 : Listbox 또는 Combobox에 해당 자료를 Setting 한다.
'*  2. 관련변수 :
'*  3. Parameter : objList - 현재 폼내의 리스트박스 또는 콤보박스 지정
'*                 intDepth - Delimiter Level (1 - 9)
'*                 strText - 삽입할 ITEM의 갯수와 실데이타
'*                 intFlag - 0 (초기화후 재작성)
'*                           1 (기존 데이타를 유지하면서 뒤에 ADD)
'*  4. ReturnValue :
'*-----------------------------------------------------------------
Public Sub medWriteList(ByRef objList As Object, ByVal intDepth As Integer, _
                        ByVal strText As String, ByVal intFlag As Integer)
Dim intPos1 As Integer, intPos2 As Integer
Dim intLength As Integer, Delimiter As String * 1

    If intFlag = 0 Then objList.Clear       ' 초기화 여부
    
    intLength = Len(strText)
    If intLength <= 0 Then Exit Sub         ' 데이타 없슴
    
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

    If intFlag = 0 Then objList.Clear       ' 초기화 여부
    
    intLength = Len(strText)
    If intLength <= 0 Then Exit Sub         ' 데이타 없슴
    
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
'*  1. 기능 : Printer Queue로 위치를 지정하여 Data를 보낸다
'*  2. 관련변수 :
'*  3. Parameter : XPos,YPos - 현재 폼내의 X,Y 좌표값 설정
'*                 PrintText, FontSize - 대상 Text와 크기
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
'*  1. 기능 : medPrinterOpen() - Printer Port를 Open한다.
'*               medPrinterClose() - Printer Port를 Close한다.
'*               medPrint() - String을 Open되어있는 Port로 Print한다.
'*  2. 관련변수 :
'*  3. Parameter :   intFileNo - 열려있는 출력포트의 번호
'*                           strData - 출력할 문자열
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
Dim intPos1 As Integer, intPos2 As Integer, I As Integer

    intPos1 = 0: intPos2 = 0
    
    ' intPosition 인수가 1인 경우 For문 Skip
    For I = 1 To intPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, strText, Delimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next I
    
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
'*  1. 기능 : Delimiter로 구분하여 지정 위치의 String을 치환한다.
'*            (Mumps의 $P()함수 이용하여 Data Store 하는 경우)
'*  2. 관련변수 :
'*  3. Parameter : strtext - Delimiter로 묶여있는 대상 문자열
'*                 intDepth - Delimiter Level (1 - 5)
'*                 intPosition - 선택 대상 문자열 위치
'*                 strWord - 치환 할 문자열
'*  4. ReturnValue : 치환된 문자열
'*-----------------------------------------------------------------
Public Function medSetP(ByVal strText As String, ByVal intDepth As Integer, _
                  ByVal intPosition As Integer, ByVal strWord As String, _
                  Optional ByVal strDeli As Variant) As String
Dim intPos1 As Integer, intPos2 As Integer, I As Integer
Dim strHead As String, strTail As String
Dim Delimiter As String

    If intPosition <= 0 Then GoTo ReturnMe ' 데이타 없슴
    
    If IsMissing(strDeli) Then
        Delimiter = medDelimiter(intDepth)
    Else
        Delimiter = strDeli
    End If
    intPos1 = 0: intPos2 = 0
    
    ' intPosition 인수가 1인 경우 For문 Skip
    For I = 1 To intPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, strText, Delimiter)
       If intPos2 = 0 Then GoTo AddDelimiter
    Next I
    
    ' 해당 컬럼
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
'*  1. 기능 : BLBX의 Delimiter로 구분된 첫번째 String을 읽어오고
'*               나머지 문자열은 그대로 남긴다.
'*  2. 관련변수 :
'*  3. Parameter : strText - Delimiter로 묶여있는 대상 문자열
'*  4. ReturnValue : 선택된 문자열
'*                   strText 자신은 Shift가 이루어진다.
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
'*  1. 기능 : Delimiter로 구분된 첫번째 String을 읽어오고
'*            나머지 문자열을 그대로 남긴다.
'*  2. 관련변수 :
'*  3. ReturnValue : 선택된 문자열
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
'*  1. 기능 : message code를 넘겨주어 해당 message를 받는다.
'*  2. 관련변수 :
'*  3. Parameter : strCode - Message Code
'*  4. ReturnValue : 메세지
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
'*  1. 기능 : 생년월일로 나이 계산
'*  2. 관련변수 :
'*  3. Parameter : strBirthDate: 생년월일(yyyymmdd)
'*                 strType:나이를 년,월,일 중 어느 기준으로 계한할 것인지
'*                     ( Y,M,D )
'*                 strSysDate : 계산의 기준이 되는 날자(yyyymmdd)
'*                     strSysDate는 Optional 없으면 현재의 날자로 나이를 계산
'*  4. ReturnValue : 계산된 나이(Year기준)
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
    Case "Y":        '년령
        medFindAge = DateDiff("yyyy", strFormatBirth, strFormatSys)
    Case "M":        '월령
        medFindAge = DateDiff("m", strFormatBirth, strFormatSys)
    Case "D":        '일령
        medFindAge = DateDiff("d", strFormatBirth, strFormatSys)
    End Select
    
End Function

'*-----------------------------------------------------------------
'*  1. 기능 : Server의 날짜를  Return
'*  2. 관련변수 :
'*  3. Parameter : intDateLenth:Return을 원하는 Date Format의
'*                 Length
'*  4. ReturnValue : Server의 날짜
'*-----------------------------------------------------------------
Function medSysDate(Optional ByVal intDateLength) As String
    
    'medSysDate = medMVB.MExe("%AJOU", "HDATE", "", 1)
    If IsMissing(intDateLength) Then Exit Function
    If (IsMissing(intDateLength) = False) And (intDateLength <= 8) Then
        'medSysDate = Right(medSysDate, intDateLength)
    End If
End Function

'*-----------------------------------------------------------------
'*  1. 기능 : Server의 시간을  Return
'*  2. 관련변수 :
'*  3. Parameter : intTimeLength:Return을 원하는 Time Format의
'*                 Length
'*  4. ReturnValue : Server의 시간
'*-----------------------------------------------------------------
Function medSysTime(Optional ByVal intTimeLength) As String
    
    'medSysTime = medMVB.MExe("%AJOU", "HTIME", "", 1)
    If IsMissing(intTimeLength) Then Exit Function
    If (IsMissing(intTimeLength) = False) And (intTimeLength <= 6) Then
        'medSysTime = Left(medSysTime, intTimeLength)
    End If
End Function

'*-----------------------------------------------------------------
'*  1. 기능 : Client의 현재날짜와 시간을  Return
'*  2. 관련변수 :
'*  3. Parameter :
'*  4. ReturnValue : Client의 현재날짜와 시간
'*-----------------------------------------------------------------
Function medThisTime() As String
Dim yy As String, mm As String, DD As String, TM As String
    
    DD = Format(Date, "YY/MM/DD")
    TM = Format(Time, "HH:MM:SS")
    medThisTime = DD & "   " & TM
    
End Function


'*-----------------------------------------------------------------
'*  1. 기능 : 해당 Data가 날자로서 그 Data가 유효한지 Check
'*  2. 관련변수 :
'*  3. Parameter : strDate : Check하고자 하는 Data
'*                 yyyymmdd(8자리) 형식만 가능
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
'*  1. 기능 : 해당 Date의 요일을 Return
'*  2. 관련변수 :
'*  3. Parameter : strDate   - Check하고자 하는 Date
'*                                          Date형식에 맞는 Data만 가능함
'*                         intOption - Return을 원하는 요일의 형식(1,2,3,4)
'*                                          ex)1:Sunday, 2:Sun, 3:일요일, 4:일
'*  4. ReturnValue : 요일(영문, 한글)
'*-----------------------------------------------------------------
Public Function medWeekday(ByVal strDate As Date, _
               ByVal intOption As Integer) As String
Dim aryPattern As Variant
Dim aryWeekday As Variant

    aryWeekday = Array("일", "월", "화", "수", "목", "금", "토")
    aryPattern = Array("ddd", "dddd")

    If IsDate(strDate) = False Then   '인수값의 date형식여부
        medWeekday = ""
        Exit Function
    End If

    If intOption < 3 Then            '영문
        medWeekday = Format(strDate, aryPattern(intOption - 1))
    Else
        medWeekday = aryWeekday(Weekday(strDate) - 1)
        If intOption = 4 Then        '한글전체
            medWeekday = medWeekday + "요일"
        End If
    End If
    
End Function

'*-----------------------------------------------------------------
'*  1. 기능 : Keyascii 값을 Upper Case값으로 변환함
'*  2. 관련변수 :
'*  3. Parameter : intKeyAscii :Keypress에서 발생하는 Keyascii값
'*  4. ReturnValue : Alphabet대문자에 해당하는 Ascii값
'*  5. 사용예제 :
'*               Private Sub Text2_KeyPress(KeyAscii As Integer)
'*                   KeyAscii = medToUCase(KeyAscii)
'*               End Sub
'*-----------------------------------------------------------------
Public Function medUCase(ByVal intKeyAscii As Integer) As Integer

    medUCase = Asc(UCase(Chr(intKeyAscii)))
    
End Function

'*-----------------------------------------------------------------
'*  1. 기능 : 해당 Spreat Sheet의 Data를 Clear
'*  2. 관련변수 :
'*  3. Parameter : objTable : Clear할 Table(속한 Form포함)
'*                 blnCol : ColHeader를 Clear할 것인지
'*                 blnRow : RowHeader를 Clear할 것인지
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
    '주)Header의 경우 Value를 Null로 주면 Default값이 나타나므로
    '   강제로 Space를 Insert합니다.
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
'*  1. 기능 : 일정시간동안 시간을 지연시킬 경우
'*  2. Parameter : Interval : 지연시킬 시간
'*-----------------------------------------------------------------
Sub medSleep(ByVal Interval As Long)

    Sleep (Interval)
    
End Sub


'*-----------------------------------------------------------------
'*  1. 기능 : ListBox에 Horizental Scroll Bar를 생성시켜 준다.
'*              (Default ListBox만 해당)
'*  2. 관련변수 :
'*  3. Parameter : lstList : 해당 ListBox Control
'*-----------------------------------------------------------------
Sub medHorScrol(ByVal lstList As Object)
   
    SendMessage lstList.hwnd, &H194, 3 * (lstList.WIDTH / Screen.TwipsPerPixelX)

End Sub


'*-----------------------------------------------------------------
'*  1. 기능 : medplay Mode를 변경한다.
'*  2. 관련변수 :
'*  3. Parameter : 변경할 Mode (1:640*480, 2:800*600, 3:1024*768)
'*-----------------------------------------------------------------
Sub medSetMode(ByVal intMode As Integer)
Dim XX As Integer, yy As Integer
    
    On Error GoTo ErrorHandler

    XX = Choose(intMode, 640, 800, 1024)
    yy = Choose(intMode, 480, 600, 768)
    
    If SetDisplayMode(XX, yy, -1) = 0 Then
        MsgBox CStr(XX) & "*" & CStr(yy) & " 모드로 변경되었습니다.", vbInformation
    Else
        MsgBox "디스플레이 모드 변경이 실패했습니다.", vbInformation
    End If
    Exit Sub
    
ErrorHandler:
        MsgBox "디스플레이 모드 변경이 실패했습니다.", vbInformation
        
End Sub

'*-----------------------------------------------------------------
'*  1. 기능 : medplay Mode를 변경한다.(Called by medSetMode())
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
'*  1. 기능 : Win32 환경 하에서 메모리의 상태를 알아낸다.
'*  3. Parameter : TotMem : 전체 물리적 메모리
'*                         AvailMem : 사용 가능한 물리적 메모리
'*-----------------------------------------------------------------
Public Sub medSysMem(ByRef TotMem As Long, ByRef AvailMem As Long)
Dim ms As MEMORYSTATUS

    ms.dwLength = Len(ms)
    GlobalMemoryStatus ms
    'MsgBox "전체 물리적 메모리 : " & ms.dwTotalPhys & vbCRLF & _
                  "사용 가능한 물리적 메모리 : " & ms.dwAvailPhys
    TotMem = ms.dwTotalPhys
    AvailMem = ms.dwAvailPhys

End Sub


'*-----------------------------------------------------------------
'*  1. 기능 : 한글입력을 Set한다.
'*  3. Parameter : Scr :입력이 행해질 Control
'*-----------------------------------------------------------------
Public Sub medHanOn(Src As Object)
Dim hIME As Long

  hIME = ImmGetContext(Src.hwnd)
  ImmSetConversionStatus hIME, IME_HANGUL, IME_NONE
  Src.SetFocus

End Sub

'*-----------------------------------------------------------------
'*  1. 기능 : 영문입력을 Set한다.
'*  3. Parameter : Scr :입력이 행해질 Control
'*-----------------------------------------------------------------
Public Sub medEngOn(Src As Object)
Dim hIME As Long

  hIME = ImmGetContext(Src.hwnd)
  ImmSetConversionStatus hIME, IME_ENGLISH, IME_NONE
  Src.SetFocus
  
End Sub


'*-----------------------------------------------------------------
'*  1. 기능 : 해당폼을 항상 위에 떠있게 한다.
'*  3. Parameter : frmForm - 해당 폼
'*                         OnOff - 0 : 해제, 1 : 설정
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
'*  1. 기능 : 데이타가 많은 리스트 상에서 특정 내용(String)을 찾는 경우
'*  2. 관련변수 :
'*  3. Parameter : lstList : 대상 리스트
'*                         strTmp : Search할 String
'*  4. ReturnValue :
'*          원하는 문자열을 찾았을 경우에는 해당 Listindex를 리턴
'*          찾지 못했을 경우에는 근접 단어의 Listindex를 리턴
'*          근접 단어 조차도 찾지 못한 경우에는 -1을 리턴
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
'*  1. 기능 : 툴바를 쿨바형태로 보여준다.
'*  2. 관련변수 :
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
'*  1. 기능 : 폼의 캡션바를 깜박이게 한다.
'*  2. 관련변수 :
'*  3. Parameter : frmForm - 해당 폼
'*                         blnOpt - True일땐 Flash, False일땐 원래상태
'*  4. ReturnValue : 해당 윈도우가 활성화되었다면 True를 반환,
'*                            그 외의 경우는 False를 반환.
'*-----------------------------------------------------------------
Function medFlashWnd(ByVal frmForm As Form, ByVal blnOpt As Boolean)

    medFlashWnd = FlashWindow(frmForm.hwnd, blnOpt)

End Function
    
'*-----------------------------------------------------------------
'*  1. 기능 : 스프레드 테이블에 Lock을 설정하거나 해제한다.
'*  2. 관련변수 :
'*  3. Parameter : objTable - 현재 폼내의 테이블 설정
'*                         blnOpt - True일땐 Lock, False일땐 Unlock
'*-----------------------------------------------------------------
Function medTableLock(ByVal objTable As Object, ByVal blnOpt As Boolean)

    objTable.Row = -1
    objTable.Col = -1
    objTable.Lock = blnOpt

End Function

'*-----------------------------------------------------------------
'*  1. 기능 : 스프레드 테이블을 행 또는 열을 기준으로 정렬한다.
'*  2. 관련변수 :
'*  3. Parameter : objTable - 현재 폼내의 테이블 설정
'*                         Col1,Col2,Row1,Row2 - 범위 설정
'*                         SortBy - 정렬기준(1:열Column, 2:행Row)
'*                         KeyOrder - 정렬Key와  Option의 2차원 배열
'*                  예 :        KeyOrder(1,1) = N -  첫번째 Key(N 번째 행 또는 열)
'*                               KeyOrder(1,2) = 1  : 오름차순, 2 : 내림차순
'*                               KeyOrder(2,1) = N  - 두번째 Key(N 번째 행 또는 열)
'*                               KeyOrder(2,2) = 1  : 오름차순, 2 : 내림차순
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
'*  1. 기능 : 리스트박스에 2개 이상의 컬럼의 데이타를 디스플레이 할때,
'*               각 행간의 정렬을 깔끔하게 해준다.
'*  2. 관련변수 :
'*  3. Parameter : objList - 현재 폼내의 리스트 설정
'*                         iColCount - 컬럼의 갯수
'*                         iColLen - 컬럼의 길이
'*                                (예)   iColLen(0) = 80    ' 80/4 = 20 Character
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
'*  1. 기능 : 해당 파일에 연결되어 있는 프로그램을 호출하여 실행한다.
'*  2. 관련변수 :
'*  3. Parameter : frmForm  - 현재 폼
'*                         FileName - 해당 파일명
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
      ' 파란색(0, 0, 255)에서 검정색으로(0, 0, 0)으로
      ' 점차적으로 칠해 나간다. 폼의 폭으로만 칠한다는
      ' 단점이 있다. 즉 사이즈가 바뀌면...
      '--------------------------------------------------
      For intLoop = 0 To 255
         vObj.Line (0, intLoop)-(100, intLoop - 1), RGB(intLoop, intLoop, intLoop), B
      Next intLoop
      
End Sub

'*-----------------------------------------------------------------
'*  1. 기능 : 문자열 내의 특정 string을 다른 string으로 대치한다.
'*  2. 관련변수 :
'*  3. Parameter : strOrigin - 대상 문자열
'*                       transFrom - 바꿀 문자열
'*                       transTo - 새 문자열
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
    ' 파란색(0, 0, 255)에서 검정색으로(0, 0, 0)으로
    ' 점차적으로 칠해 나간다. 폼의 폭으로만 칠한다는
    ' 단점이 있다. 즉 사이즈가 바뀌면...
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
      medGetMessage = "Database 연결이 안됬습니다. 전산실로 문의 바랍니다."
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

