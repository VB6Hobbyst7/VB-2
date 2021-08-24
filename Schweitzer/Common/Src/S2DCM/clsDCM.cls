VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDCM"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



'*==========================================================================*'
'*                                                                          *'
'*           S2DCM(Schweitzer2000 Daeryun Common Module) Class              *'
'*                                                                          *'
'*               (c)Copyright 2000 DAERYUN MTS Co.,Ltd.                     *'
'*                                                                          *'
'*==========================================================================*'


'# medBeep
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'# medGetComNm
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'# medHorScrol, medListFind
Private Const LB_SETTABSTOPS = &H192
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                              (ByVal hwnd As Long, ByVal wMsg As Long, _
                              ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function SendMessage1 Lib "user32" Alias "SendMessageA" _
                              (ByVal hwnd As Long, ByVal wMsg As Long, _
                              ByVal wParam As Long, lParam As Any) As Long

'# SetDisplayMode (Called by medSetMode())
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (lpString1 As Any, _
                                                          lpString2 As Any) As Long
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" _
                    (ByVal lpDevMode As Long, ByVal dwFlags As Long) As Long
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Type DEVMODE
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

'# medAlwaysOn
Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Private Const HWND_NOTOPMOST = -2   'Not Always top
Private Const HWND_TOPMOST = -1  'Always top
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

'# medSysMem
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

'# medHanOn, medEngOn
Private Const IME_HANGUL = &H1
Private Const IME_ENGLISH = &H0
Private Const IME_NONE = &H0
Private Declare Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ImmSetConversionStatus Lib "imm32.dll" _
                (ByVal hIMC As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long

'# medLockWindowUpdate
Private Declare Function LockWindowUpdate Lib "user32" _
                (ByVal hwndLock As Long) As Long

'# medFlashWnd
Private Declare Function FlashWindow Lib "user32" _
                (ByVal hwnd As Long, ByVal bInvert As Long) As Long

'#
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private rectTmp As RECT
Private Declare Function FillRect Lib "user32" _
                (ByVal hDC As Long, lpRect As RECT, _
                  ByVal hBrush As Long) As Long

'# medShell
Private Const SW_SHOWDEFAULT = 10
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWNORMAL = 1

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
            (ByVal hwnd As Long, ByVal lpOperation As String, _
             ByVal lpFile As String, ByVal lpParameters As String, _
            ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'# Screen Lock 관련
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

'# Homepage 띄우기
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" _
    (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As _
    String) As Long



'*-----------------------------------------------------------------
'*  1. 기능 : 컴퓨터 이름 가져오기..
'*-----------------------------------------------------------------
Public Function medGetComNm()

   Dim sBuffer$, nSize As Long, rtn As Long
   sBuffer = String(256, Chr(0))
   rtn = GetComputerName(sBuffer$, Len(sBuffer))
   medGetComNm = sBuffer
   
End Function

'*-----------------------------------------------------------------
'*  1. 기능 : 경고음을 인수값만큼 반복해서 소리낸다.
'*  2. Parameter : intCnt (Beep Count)
'*-----------------------------------------------------------------
Public Sub medBeep(ByVal intCNT As Integer)
Dim i As Integer
    
    For i = 1 To intCNT
        Call Beep
    Next i
    
End Sub


'*-----------------------------------------------------------------
'*  1. 기능 : 해당 SpreadSheet의 지정 위치에 데이타를 Write 한다.
'*  2. Parameter : objTable - 현재 폼내의 테이블 설정
'*                 Col,Row  - Col,Row 위치 설정
'*                 StoreData- Write 할 String
'*                 FontSize - 글자 크기(Optional)
'*  3. ReturnValue :
'*-----------------------------------------------------------------
Public Sub medSetText(ByRef objTable As Object, _
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
'*  2. Parameter : objTable - 현재 폼내의 테이블 설정
'*                 Col1,Col2,Row1,Row2 - 범위 설정
'*  3. ReturnValue :
'*-----------------------------------------------------------------
Public Function medGetText(ByRef objTable As Object, _
                    ByVal Col As Integer, ByVal Row As Integer) As String

    objTable.Col = Col
    objTable.Row = Row
    
    medGetText = objTable.Value
    
End Function


'*-----------------------------------------------------------------
'*  1. 기능 : 해당 SpreadSheet의 지정 범위내에 데이타를 Write 한다.
'*  2. Parameter : objTable - 현재 폼내의 테이블 설정
'*                 Col1,Col2,Row1,Row2 - 범위 설정
'*                 StoreData, FontSize - 데이터와 글자 크기
'*  3. ReturnValue :
'*-----------------------------------------------------------------
Public Sub medSetBlockText(ByRef objTable As Object, _
                           ByVal Col1 As Integer, ByVal COL2 As Integer, _
                           ByVal Row1 As Integer, ByVal Row2 As Integer, _
                           ByVal StoreData As String, Optional ByVal FontSize As Variant)

    If objTable.MaxRows < Row2 Then objTable.MaxRows = Row2
    If objTable.MaxCols < COL2 Then objTable.MaxCols = COL2
    
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
'*  2. Parameter : objTable - 현재 폼내의 테이블 설정
'*                 Col1,Col2,Row1,Row2 - 범위 설정
'*  3. ReturnValue : Colume은 ASC(9), Row는 ASC(13)+ASC(10)으로
'*                   구분된 하나의 String으로 리턴
'*-----------------------------------------------------------------
Public Function medGetBlockText(ByRef objTable As Object, _
                    ByVal Col1 As Integer, ByVal COL2 As Integer, _
                    ByVal Row1 As Integer, ByVal Row2 As Integer) As String

    objTable.Col = Col1: objTable.COL2 = COL2
    objTable.Row = Row1: objTable.Row2 = Row2
    
    medGetBlockText = ""
    
    objTable.BlockMode = True
    medGetBlockText = objTable.ClipValue
    objTable.BlockMode = False

End Function


'*-----------------------------------------------------------------
'*  1. 기능 : Delimiter로 구분하여 지정 위치의 String을 읽어온다.
'*            (Mumps의 $P()함수 이용하여 Data Read 하는 경우)
'*  2. Parameter : strtext      - Delimiter로 묶여있는 대상 문자열
'*                 intPosition  - 선택 대상 문자열 위치
'*                 Delimiter    - 사용자정의의 구분자
'*  3. ReturnValue : 선택된 문자열
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
'*  1. 기능 : Delimiter로 구분하여 지정 위치의 String을 치환한다.
'*            (Mumps의 $P()함수 이용하여 Data Store 하는 경우)
'*  2. Parameter : strtext - Delimiter로 묶여있는 대상 문자열
'*                 intPosition - 선택 대상 문자열 위치
'*                 strWord - 치환 할 문자열
'*                 strDeli - Delimiter
'*  3. ReturnValue : 치환된 문자열
'*-----------------------------------------------------------------
Public Function medSetP(ByVal strText As String, _
                        ByVal intPosition As Integer, ByVal strWord As String, _
                        ByVal strDeli As Variant) As String
Dim intPos1 As Integer, intPos2 As Integer, i As Integer
Dim strHead As String, strTail As String
Dim Delimiter As String

    If intPosition <= 0 Then GoTo ReturnMe ' 데이타 없슴
    
'    If IsMissing(strDeli) Then
'        Delimiter = medDelimiter(intDepth)
'    Else
        Delimiter = strDeli
'    End If
    intPos1 = 0: intPos2 = 0
    
    ' intPosition 인수가 1인 경우 For문 Skip
    For i = 1 To intPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, strText, Delimiter)
       If intPos2 = 0 Then GoTo AddDelimiter
    Next i
    
    ' 해당 컬럼
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, strText, Delimiter)
    If intPos2 = 0 Then intPos2 = Len(strText) + 1
    
    'strHead = Mid(strText.Text, 1, intPos1 - 1)
    'strTail = Right$(strText.Text, Len(strText) - intPos2 + 1)
    'medSetP = strHead & strWord & strTail
    
    Exit Function
    
AddDelimiter:
    medSetP = strText & String(intPosition - i, Delimiter) & strWord
    Exit Function
    
ReturnMe:
    medSetP = strText
    
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
'*  1. 기능 : 생년월일로 나이 계산
'*  2. Parameter : strBirthDate: 생년월일(yyyymmdd)
'*                 strType:나이를 년,월,일 중 어느 기준으로 계한할 것인지
'*                     ( Y,M,D )
'*                 strSysDate : 계산의 기준이 되는 날자(yyyymmdd)
'*                     strSysDate는 Optional 없으면 현재의 날자로 나이를 계산
'*  3. ReturnValue : 계산된 나이(Year기준)
'*-----------------------------------------------------------------
Public Function medFindAge(ByVal strBirthDate As String, ByVal strAgeType As String, _
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
'*  1. 기능 : Client의 현재날짜와 시간을  Return
'*  2. Parameter :
'*  3. ReturnValue : Client의 현재날짜와 시간 (YY/MM/DD  HH:MM:SS)
'*-----------------------------------------------------------------
Public Function medThisTime() As String
Dim yy As String, mm As String, DD As String, TM As String
    
    DD = Format(Date, "YY/MM/DD")
    TM = Format(Time, "HH:MM:SS")
    medThisTime = DD & "   " & TM
    
End Function


'*-----------------------------------------------------------------
'*  1. 기능 : 해당 Date의 요일을 Return
'*  2. Parameter : strDate  - Check하고자 하는 Date
'*                            Date형식에 맞는 Data만 가능함
'*                 intOption- Return을 원하는 요일의 형식(1,2,3,4)
'*                            ex)1:Sunday, 2:Sun, 3:일요일, 4:일
'*  3. ReturnValue : 요일(영문, 한글)
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
'*  2. Parameter : intKeyAscii :Keypress에서 발생하는 Keyascii값
'*  3. ReturnValue : Alphabet대문자에 해당하는 Ascii값
'*  4. 사용예제 :
'*               Private Sub Text2_KeyPress(KeyAscii As Integer)
'*                   KeyAscii = medToUCase(KeyAscii)
'*               End Sub
'*-----------------------------------------------------------------
Public Function medUCase(ByVal intKeyAscii As Integer) As Integer

    medUCase = Asc(UCase(Chr(intKeyAscii)))
    
End Function

'*-----------------------------------------------------------------
'*  1. 기능 : 해당 Spread Sheet의 Data를 Clear
'*  2. Parameter : objTable : Clear할 Table(속한 Form포함)
'*                 blnCol   : ColHeader를 Clear할 것인지
'*                 blnRow   : RowHeader를 Clear할 것인지
'*  3. ReturnValue :
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
' 1. 기능 : Main Form의 StatusBar에 Message를 Display한다.
'*-----------------------------------------------------------------
Public Sub medSetMsg(ByVal frmMain As Object, ByVal MSG As String, Optional BeepOk)
   
    frmMain.stsBar.Panels(2) = MSG
    '
    If IsMissing(BeepOk) = False Then
      medBeep (3)
    End If

End Sub


'*-----------------------------------------------------------------
'*  1. 기능 : 일정시간동안 시간을 지연시킬 경우
'*  2. Parameter : Interval - 지연시킬 시간
'*-----------------------------------------------------------------
Public Sub medSleep(ByVal Interval As Long)

    Sleep (Interval)
    
End Sub


'*-----------------------------------------------------------------
'*  1. 기능 : ListBox에 Horizental Scroll Bar를 생성시켜 준다.
'*              (Default ListBox만 해당)
'*  2. Parameter : lstList - 해당 ListBox Control
'*-----------------------------------------------------------------
Public Sub medHorScrol(ByVal lstList As Object)
   
    SendMessage lstList.hwnd, &H194, 3 * (lstList.WIDTH / Screen.TwipsPerPixelX), 0&

End Sub


'*-----------------------------------------------------------------
'*  1. 기능 : medplay Mode를 변경한다.
'*  2. Parameter : intMode - 변경할 Mode (1:640*480, 2:800*600, 3:1024*768)
'*-----------------------------------------------------------------
Public Sub medSetMode(ByVal intMode As Integer)
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
'*  2. Parameter : TotMem   - 전체 물리적 메모리
'*                 AvailMem - 사용 가능한 물리적 메모리
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
'*  2. Parameter : Scr :입력이 행해질 Control
'*-----------------------------------------------------------------
Public Sub medHanOn(Src As Object)
Dim hIME As Long

  hIME = ImmGetContext(Src.hwnd)
  ImmSetConversionStatus hIME, IME_HANGUL, IME_NONE
  Src.SetFocus

End Sub

'*-----------------------------------------------------------------
'*  1. 기능 : 영문입력을 Set한다.
'*  2. Parameter : Scr :입력이 행해질 Control
'*-----------------------------------------------------------------
Public Sub medEngOn(Src As Object)
Dim hIME As Long

  hIME = ImmGetContext(Src.hwnd)
  ImmSetConversionStatus hIME, IME_ENGLISH, IME_NONE
  Src.SetFocus
  
End Sub

'*-----------------------------------------------------------------
'*  1. 기능 : 리스트의 Redraw를 막아서 Add 또는 Scroll 속도를 향상한다.
'*  2. Parameter : hwndLock :lstMicCd.hwnd - Lock, &0 - Unlock
'*-----------------------------------------------------------------
Public Sub medLockWindowUpdate(ByVal hwndLock As Long)
    
    Call LockWindowUpdate(hwndLock)
  
End Sub


'*-----------------------------------------------------------------
'*  1. 기능 : 해당폼을 항상 위에 떠있게 한다.
'*  2. Parameter : frmForm - 해당 폼
'*                 OnOff - 0 : 해제, 1 : 설정
'*-----------------------------------------------------------------
Public Sub medAlwaysOn(ByVal frmForm As Object, ByVal OnOff As Integer)
Dim hWndMode As Integer

    hWndMode = Choose(OnOff + 1, -2, -1)
    SetWindowPos frmForm.hwnd, hWndMode, 0, 0, 10, 10, _
                                        SWP_NOMOVE Or SWP_NOSIZE
'    SetWindowPos frmForm.hwnd, HWND_TOPMOST, 0, 0, 10, 10, _
                                        SWP_NOMOVE Or SWP_NOSIZE

End Sub
 

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
    
Public Function medComboFind(ByVal cboCombo As Object, ByVal strTmp As String)
    
    Dim i As Integer
    
    With cboCombo
      For i = 0 To .ListCount - 1
         If .List(i) Like (strTmp & "*") Then
            medComboFind = i
            Exit Function
         End If
      Next
   End With
   medComboFind = -1

End Function
    
    
'*-----------------------------------------------------------------
'*  1. 기능 : 폼의 캡션바를 깜박이게 한다.
'*  2. Parameter   : frmForm - 해당 폼
'*                   blnOpt  - True일땐 Flash, False일땐 원래상태
'*  3. ReturnValue : 해당 윈도우가 활성화되었다면 True를 반환,
'*                   그 외의 경우는 False를 반환.
'*-----------------------------------------------------------------
Public Function medFlashWnd(ByVal frmForm As Object, ByVal blnOpt As Boolean)

    medFlashWnd = FlashWindow(frmForm.hwnd, blnOpt)

End Function
    
'*-----------------------------------------------------------------
'*  1. 기능 : 스프레드 테이블에 Lock을 설정하거나 해제한다.
'*  2. Parameter : objTable - 현재 폼내의 테이블 설정
'*                 blnOpt   - True일땐 Lock, False일땐 Unlock
'*-----------------------------------------------------------------
Public Function medTableLock(ByVal objTable As Object, ByVal blnOpt As Boolean)

    objTable.Row = -1
    objTable.Col = -1
    objTable.Protect = True
    objTable.Lock = blnOpt

End Function

'*-----------------------------------------------------------------
'*  1. 기능 : 스프레드 테이블을 행 또는 열을 기준으로 정렬한다.
'*  2. Parameter : objTable - 현재 폼내의 테이블 설정
'*                 Col1,Col2,Row1,Row2 - 범위 설정
'*                 SortBy - 정렬기준(1:열Column, 2:행Row)
'*                 KeyOrder - 정렬Key와  Option의 2차원 배열
'*                 예 : KeyOrder(1,1) = N -  첫번째 Key(N 번째 행 또는 열)
'*                      KeyOrder(1,2) = 1  : 오름차순, 2 : 내림차순
'*                      KeyOrder(2,1) = N  - 두번째 Key(N 번째 행 또는 열)
'*                      KeyOrder(2,2) = 1  : 오름차순, 2 : 내림차순
'*-----------------------------------------------------------------
Public Function medTableSort(ByVal objTable As Object, _
                             ByVal Col1 As Integer, ByVal COL2 As Integer, _
                             ByVal Row1 As Integer, ByVal Row2 As Integer, _
                             ByVal SortBy As String, ByRef KeyOrder() As Variant)
Dim i As Integer

    objTable.Col = Col1: objTable.COL2 = COL2
    objTable.Row = Row1: objTable.Row2 = Row2
    objTable.BlockMode = True
    
    objTable.SortBy = SortBy
    
    For i = 1 To UBound(KeyOrder)
        objTable.SortKey(i) = KeyOrder(i, 1)
        objTable.SortKeyOrder(i) = KeyOrder(i, 2)
    Next i
    
    objTable.Action = 25
    objTable.BlockMode = False

End Function


'*-----------------------------------------------------------------
'*  1. 기능 : 해당 파일에 연결되어 있는 프로그램을 호출하여 실행한다.
'*  2. Parameter : frmForm  - 현재 폼
'*                 FileName - 해당 파일명
'*-----------------------------------------------------------------
Public Sub medShell(ByVal frmForm As Object, ByVal FileName As String)
    
    ShellExecute frmForm.hwnd, "Open", FileName, vbNullString, vbNullString, SW_SHOWDEFAULT

End Sub


'*-----------------------------------------------------------------
'*  1. 기능 : 문자열 내의 특정 string을 다른 string으로 대치한다.
'*  2. Parameter : strOrigin - 대상 문자열
'*                 transFrom - 바꿀 문자열
'*                 transTo - 새 문자열
'*-----------------------------------------------------------------
Public Function medReplace(ByVal strOrigin As String, ByVal transFrom As String, transTo As String)
Dim i As Integer
Dim intLen As Integer

    i = 1
    intLen = Len(transFrom)
    Do While i <= Len(strOrigin)
        If Mid(strOrigin, i, intLen) = transFrom Then
            strOrigin = Mid(strOrigin, 1, i - 1) & transTo & Mid(strOrigin, i + intLen)
            i = i + intLen
        Else
            i = i + 1
        End If
    Loop
    medReplace = strOrigin
            
End Function

Public Sub Dither(vObj As Object)
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


Public Sub Dither1(vObj As Object)

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


'*-----------------------------------------------------------------
'*  1. 기능 : SQL문 작성 시 Character필드의 처리문을 반환.
'*  2. Parameter : strValue - Data
'*                 optNum - 1:insert, 2:where, 3:update
'*-----------------------------------------------------------------
Public Function DBStr(ByVal strValue As String, _
   Optional ByVal optNum As Variant) As String
'String Conversion For Database INSERT,UPDATE
   Dim iPos As Integer
   
   iPos = InStr(1, strValue, "'")
   While iPos > 0
        strValue = Mid(strValue, 1, iPos) & "'" & Mid(strValue, iPos + 1)
        iPos = InStr(iPos + 2, strValue, "'")
   Wend
   
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


'*-----------------------------------------------------------------
'*  1. 기능 : SQL문 작성 시 Numeric필드의 처리문을 반환.
'*  2. Parameter : NumValue - Data
'*                 optNum - 1:insert, 2:where, 3:update
'*-----------------------------------------------------------------
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


'*-----------------------------------------------------------------
'*  1. 기능 : FormNames에 폼명이 지정되지 않은 폼들을 모두 Unload한다.
'*  2. Parameter : FormNames - 언로드시키지 말아야 할 폼명들
'*                             예) "medMain,medSplash,medLogOn"
'*-----------------------------------------------------------------
Public Sub medUnloadForms(ByVal FormNames As String)

   Dim tmpForm As Form
   
   For Each tmpForm In Forms
      With tmpForm
         If InStr(1, UCase(FormNames), UCase(.Name)) = 0 Then
            Unload tmpForm
         End If
      End With
   Next

End Sub


' List View 에 데이터를 Display한다.
Public Sub DataLoadLvw(ByRef objLvw As Object, _
   ByVal RowDel As String, ByVal ColDel As String, _
   ByVal strData As String, Optional strTag As String)
Dim itmX As ListItem
Dim strTmp As String
Dim aryTmp() As String
Dim aryTag() As String
Dim ii As Integer
Dim jj As Integer
Dim intCol As Integer
   aryTmp = Split(medGetP(strData, 1, RowDel), ColDel)
   If IsMissing(strTag) Then
      strTag = ""
   End If
   aryTag = Split(strTag, RowDel)
   intCol = UBound(aryTmp) + 1
   '
   aryTmp = Split(strData, RowDel)
   If UBound(aryTmp) > UBound(aryTag) Then
      ReDim Preserve aryTag(UBound(aryTmp))
   End If
   If (UBound(aryTmp) + 1) < 1 Then Exit Sub
   For ii = 0 To UBound(aryTmp)
      For jj = 1 To intCol
         If jj = 1 Then
            Set itmX = objLvw.ListItems.Add(, , medGetP(aryTmp(ii), jj, ColDel))
         Else
            If medGetP(aryTmp(ii), jj, ColDel) <> "" Then
               itmX.SubItems(jj - 1) = medGetP(aryTmp(ii), jj, ColDel)
            Else
               itmX.SubItems(jj - 1) = " "
            End If
         End If
         itmX.Tag = aryTag(ii)
      Next jj

   Next ii
   Set itmX = Nothing
   '
End Sub

'ListView의 Header를 Set한다.
Public Sub InitLvwHead(ByRef objLvw As Object, ByVal strHead As String, _
   ByVal strSize As String)
Dim ii As Integer
Dim aryTitle() As String
Dim aryWidth() As String
   aryTitle = Split(strHead, ",")
   aryWidth = Split(strSize, ",")
   If UBound(aryWidth) < UBound(aryTitle) Then
      ReDim Preserve aryWidth(UBound(aryTitle))
   End If
   With objLvw
      .ColumnHeaders.Clear
      For ii = 0 To UBound(aryTitle)
         If aryWidth(ii) = "" Then
            aryWidth(ii) = "0"
         End If
        .ColumnHeaders.Add ii + 1, aryTitle(ii), aryTitle(ii), _
            (.WIDTH \ (UBound(aryTitle) + 1)) + Val(aryWidth(ii)), vbLeftJustify
      Next ii
      .View = lvwReport            ' Report Style
   End With
End Sub

'Public Sub Crystal_Print(ByVal CrystalNm As Object, ByVal strTmp As String, _
'                            ByVal strFilePath As String, ByVal strRptPath As String)
'
'    'CrystalNm:Crystal컨트롤 Name
'    'strTmp: Record String(출력값)
'    'strRptPath: Rpt파일 경로
'    'strFilePath: text Fil 경로
'
'    Dim intFNum As Integer
'
'    intFNum = FreeFile
'    Open strFilePath For Output As #intFNum
'    Print #intFNum, strTmp
'    Close #intFNum
'    With CrystalNm
'        .ReportFileName = strRptPath
'        .RetrieveDataFiles
'        .WindowState = crptMaximized
'        .Action = 1
'        .Reset
'    End With
'End Sub


