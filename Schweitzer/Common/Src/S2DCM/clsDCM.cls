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

'# Screen Lock ����
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

'# Homepage ����
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" _
    (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As _
    String) As Long



'*-----------------------------------------------------------------
'*  1. ��� : ��ǻ�� �̸� ��������..
'*-----------------------------------------------------------------
Public Function medGetComNm()

   Dim sBuffer$, nSize As Long, rtn As Long
   sBuffer = String(256, Chr(0))
   rtn = GetComputerName(sBuffer$, Len(sBuffer))
   medGetComNm = sBuffer
   
End Function

'*-----------------------------------------------------------------
'*  1. ��� : ������� �μ�����ŭ �ݺ��ؼ� �Ҹ�����.
'*  2. Parameter : intCnt (Beep Count)
'*-----------------------------------------------------------------
Public Sub medBeep(ByVal intCNT As Integer)
Dim i As Integer
    
    For i = 1 To intCNT
        Call Beep
    Next i
    
End Sub


'*-----------------------------------------------------------------
'*  1. ��� : �ش� SpreadSheet�� ���� ��ġ�� ����Ÿ�� Write �Ѵ�.
'*  2. Parameter : objTable - ���� ������ ���̺� ����
'*                 Col,Row  - Col,Row ��ġ ����
'*                 StoreData- Write �� String
'*                 FontSize - ���� ũ��(Optional)
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
'*  1. ��� : �ش� SpreadSheet�� ���� ��ġ�� ����Ÿ�� Read �Ѵ�.
'*  2. Parameter : objTable - ���� ������ ���̺� ����
'*                 Col1,Col2,Row1,Row2 - ���� ����
'*  3. ReturnValue :
'*-----------------------------------------------------------------
Public Function medGetText(ByRef objTable As Object, _
                    ByVal Col As Integer, ByVal Row As Integer) As String

    objTable.Col = Col
    objTable.Row = Row
    
    medGetText = objTable.Value
    
End Function


'*-----------------------------------------------------------------
'*  1. ��� : �ش� SpreadSheet�� ���� �������� ����Ÿ�� Write �Ѵ�.
'*  2. Parameter : objTable - ���� ������ ���̺� ����
'*                 Col1,Col2,Row1,Row2 - ���� ����
'*                 StoreData, FontSize - �����Ϳ� ���� ũ��
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
'*  1. ��� : �ش� SpreadSheet�� ���� ������ ����Ÿ�� Read �Ѵ�.
'*  2. Parameter : objTable - ���� ������ ���̺� ����
'*                 Col1,Col2,Row1,Row2 - ���� ����
'*  3. ReturnValue : Colume�� ASC(9), Row�� ASC(13)+ASC(10)����
'*                   ���е� �ϳ��� String���� ����
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
'*  1. ��� : Delimiter�� �����Ͽ� ���� ��ġ�� String�� �о�´�.
'*            (Mumps�� $P()�Լ� �̿��Ͽ� Data Read �ϴ� ���)
'*  2. Parameter : strtext      - Delimiter�� �����ִ� ��� ���ڿ�
'*                 intPosition  - ���� ��� ���ڿ� ��ġ
'*                 Delimiter    - ����������� ������
'*  3. ReturnValue : ���õ� ���ڿ�
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
'*  1. ��� : Delimiter�� �����Ͽ� ���� ��ġ�� String�� ġȯ�Ѵ�.
'*            (Mumps�� $P()�Լ� �̿��Ͽ� Data Store �ϴ� ���)
'*  2. Parameter : strtext - Delimiter�� �����ִ� ��� ���ڿ�
'*                 intPosition - ���� ��� ���ڿ� ��ġ
'*                 strWord - ġȯ �� ���ڿ�
'*                 strDeli - Delimiter
'*  3. ReturnValue : ġȯ�� ���ڿ�
'*-----------------------------------------------------------------
Public Function medSetP(ByVal strText As String, _
                        ByVal intPosition As Integer, ByVal strWord As String, _
                        ByVal strDeli As Variant) As String
Dim intPos1 As Integer, intPos2 As Integer, i As Integer
Dim strHead As String, strTail As String
Dim Delimiter As String

    If intPosition <= 0 Then GoTo ReturnMe ' ����Ÿ ����
    
'    If IsMissing(strDeli) Then
'        Delimiter = medDelimiter(intDepth)
'    Else
        Delimiter = strDeli
'    End If
    intPos1 = 0: intPos2 = 0
    
    ' intPosition �μ��� 1�� ��� For�� Skip
    For i = 1 To intPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, strText, Delimiter)
       If intPos2 = 0 Then GoTo AddDelimiter
    Next i
    
    ' �ش� �÷�
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
'*  1. ��� : ������Ϸ� ���� ���
'*  2. Parameter : strBirthDate: �������(yyyymmdd)
'*                 strType:���̸� ��,��,�� �� ��� �������� ������ ������
'*                     ( Y,M,D )
'*                 strSysDate : ����� ������ �Ǵ� ����(yyyymmdd)
'*                     strSysDate�� Optional ������ ������ ���ڷ� ���̸� ���
'*  3. ReturnValue : ���� ����(Year����)
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
    Case "Y":        '���
        medFindAge = DateDiff("yyyy", strFormatBirth, strFormatSys)
    Case "M":        '����
        medFindAge = DateDiff("m", strFormatBirth, strFormatSys)
    Case "D":        '�Ϸ�
        medFindAge = DateDiff("d", strFormatBirth, strFormatSys)
    End Select
    
End Function


'*-----------------------------------------------------------------
'*  1. ��� : Client�� ���糯¥�� �ð���  Return
'*  2. Parameter :
'*  3. ReturnValue : Client�� ���糯¥�� �ð� (YY/MM/DD  HH:MM:SS)
'*-----------------------------------------------------------------
Public Function medThisTime() As String
Dim yy As String, mm As String, DD As String, TM As String
    
    DD = Format(Date, "YY/MM/DD")
    TM = Format(Time, "HH:MM:SS")
    medThisTime = DD & "   " & TM
    
End Function


'*-----------------------------------------------------------------
'*  1. ��� : �ش� Date�� ������ Return
'*  2. Parameter : strDate  - Check�ϰ��� �ϴ� Date
'*                            Date���Ŀ� �´� Data�� ������
'*                 intOption- Return�� ���ϴ� ������ ����(1,2,3,4)
'*                            ex)1:Sunday, 2:Sun, 3:�Ͽ���, 4:��
'*  3. ReturnValue : ����(����, �ѱ�)
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
'*  2. Parameter : intKeyAscii :Keypress���� �߻��ϴ� Keyascii��
'*  3. ReturnValue : Alphabet�빮�ڿ� �ش��ϴ� Ascii��
'*  4. ��뿹�� :
'*               Private Sub Text2_KeyPress(KeyAscii As Integer)
'*                   KeyAscii = medToUCase(KeyAscii)
'*               End Sub
'*-----------------------------------------------------------------
Public Function medUCase(ByVal intKeyAscii As Integer) As Integer

    medUCase = Asc(UCase(Chr(intKeyAscii)))
    
End Function

'*-----------------------------------------------------------------
'*  1. ��� : �ش� Spread Sheet�� Data�� Clear
'*  2. Parameter : objTable : Clear�� Table(���� Form����)
'*                 blnCol   : ColHeader�� Clear�� ������
'*                 blnRow   : RowHeader�� Clear�� ������
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
' 1. ��� : Main Form�� StatusBar�� Message�� Display�Ѵ�.
'*-----------------------------------------------------------------
Public Sub medSetMsg(ByVal frmMain As Object, ByVal MSG As String, Optional BeepOk)
   
    frmMain.stsBar.Panels(2) = MSG
    '
    If IsMissing(BeepOk) = False Then
      medBeep (3)
    End If

End Sub


'*-----------------------------------------------------------------
'*  1. ��� : �����ð����� �ð��� ������ų ���
'*  2. Parameter : Interval - ������ų �ð�
'*-----------------------------------------------------------------
Public Sub medSleep(ByVal Interval As Long)

    Sleep (Interval)
    
End Sub


'*-----------------------------------------------------------------
'*  1. ��� : ListBox�� Horizental Scroll Bar�� �������� �ش�.
'*              (Default ListBox�� �ش�)
'*  2. Parameter : lstList - �ش� ListBox Control
'*-----------------------------------------------------------------
Public Sub medHorScrol(ByVal lstList As Object)
   
    SendMessage lstList.hwnd, &H194, 3 * (lstList.WIDTH / Screen.TwipsPerPixelX), 0&

End Sub


'*-----------------------------------------------------------------
'*  1. ��� : medplay Mode�� �����Ѵ�.
'*  2. Parameter : intMode - ������ Mode (1:640*480, 2:800*600, 3:1024*768)
'*-----------------------------------------------------------------
Public Sub medSetMode(ByVal intMode As Integer)
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
'*  2. Parameter : TotMem   - ��ü ������ �޸�
'*                 AvailMem - ��� ������ ������ �޸�
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
'*  2. Parameter : Scr :�Է��� ������ Control
'*-----------------------------------------------------------------
Public Sub medHanOn(Src As Object)
Dim hIME As Long

  hIME = ImmGetContext(Src.hwnd)
  ImmSetConversionStatus hIME, IME_HANGUL, IME_NONE
  Src.SetFocus

End Sub

'*-----------------------------------------------------------------
'*  1. ��� : �����Է��� Set�Ѵ�.
'*  2. Parameter : Scr :�Է��� ������ Control
'*-----------------------------------------------------------------
Public Sub medEngOn(Src As Object)
Dim hIME As Long

  hIME = ImmGetContext(Src.hwnd)
  ImmSetConversionStatus hIME, IME_ENGLISH, IME_NONE
  Src.SetFocus
  
End Sub

'*-----------------------------------------------------------------
'*  1. ��� : ����Ʈ�� Redraw�� ���Ƽ� Add �Ǵ� Scroll �ӵ��� ����Ѵ�.
'*  2. Parameter : hwndLock :lstMicCd.hwnd - Lock, &0 - Unlock
'*-----------------------------------------------------------------
Public Sub medLockWindowUpdate(ByVal hwndLock As Long)
    
    Call LockWindowUpdate(hwndLock)
  
End Sub


'*-----------------------------------------------------------------
'*  1. ��� : �ش����� �׻� ���� ���ְ� �Ѵ�.
'*  2. Parameter : frmForm - �ش� ��
'*                 OnOff - 0 : ����, 1 : ����
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
'*  1. ��� : ���� ĸ�ǹٸ� �����̰� �Ѵ�.
'*  2. Parameter   : frmForm - �ش� ��
'*                   blnOpt  - True�϶� Flash, False�϶� ��������
'*  3. ReturnValue : �ش� �����찡 Ȱ��ȭ�Ǿ��ٸ� True�� ��ȯ,
'*                   �� ���� ���� False�� ��ȯ.
'*-----------------------------------------------------------------
Public Function medFlashWnd(ByVal frmForm As Object, ByVal blnOpt As Boolean)

    medFlashWnd = FlashWindow(frmForm.hwnd, blnOpt)

End Function
    
'*-----------------------------------------------------------------
'*  1. ��� : �������� ���̺��� Lock�� �����ϰų� �����Ѵ�.
'*  2. Parameter : objTable - ���� ������ ���̺� ����
'*                 blnOpt   - True�϶� Lock, False�϶� Unlock
'*-----------------------------------------------------------------
Public Function medTableLock(ByVal objTable As Object, ByVal blnOpt As Boolean)

    objTable.Row = -1
    objTable.Col = -1
    objTable.Protect = True
    objTable.Lock = blnOpt

End Function

'*-----------------------------------------------------------------
'*  1. ��� : �������� ���̺��� �� �Ǵ� ���� �������� �����Ѵ�.
'*  2. Parameter : objTable - ���� ������ ���̺� ����
'*                 Col1,Col2,Row1,Row2 - ���� ����
'*                 SortBy - ���ı���(1:��Column, 2:��Row)
'*                 KeyOrder - ����Key��  Option�� 2���� �迭
'*                 �� : KeyOrder(1,1) = N -  ù��° Key(N ��° �� �Ǵ� ��)
'*                      KeyOrder(1,2) = 1  : ��������, 2 : ��������
'*                      KeyOrder(2,1) = N  - �ι�° Key(N ��° �� �Ǵ� ��)
'*                      KeyOrder(2,2) = 1  : ��������, 2 : ��������
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
'*  1. ��� : �ش� ���Ͽ� ����Ǿ� �ִ� ���α׷��� ȣ���Ͽ� �����Ѵ�.
'*  2. Parameter : frmForm  - ���� ��
'*                 FileName - �ش� ���ϸ�
'*-----------------------------------------------------------------
Public Sub medShell(ByVal frmForm As Object, ByVal FileName As String)
    
    ShellExecute frmForm.hwnd, "Open", FileName, vbNullString, vbNullString, SW_SHOWDEFAULT

End Sub


'*-----------------------------------------------------------------
'*  1. ��� : ���ڿ� ���� Ư�� string�� �ٸ� string���� ��ġ�Ѵ�.
'*  2. Parameter : strOrigin - ��� ���ڿ�
'*                 transFrom - �ٲ� ���ڿ�
'*                 transTo - �� ���ڿ�
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
      ' �Ķ���(0, 0, 255)���� ����������(0, 0, 0)����
      ' ���������� ĥ�� ������. ���� �����θ� ĥ�Ѵٴ�
      ' ������ �ִ�. �� ����� �ٲ��...
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


'*-----------------------------------------------------------------
'*  1. ��� : SQL�� �ۼ� �� Character�ʵ��� ó������ ��ȯ.
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
'*  1. ��� : SQL�� �ۼ� �� Numeric�ʵ��� ó������ ��ȯ.
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
'*  1. ��� : FormNames�� ������ �������� ���� ������ ��� Unload�Ѵ�.
'*  2. Parameter : FormNames - ��ε��Ű�� ���ƾ� �� ������
'*                             ��) "medMain,medSplash,medLogOn"
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


' List View �� �����͸� Display�Ѵ�.
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

'ListView�� Header�� Set�Ѵ�.
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
'    'CrystalNm:Crystal��Ʈ�� Name
'    'strTmp: Record String(��°�)
'    'strRptPath: Rpt���� ���
'    'strFilePath: text Fil ���
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

