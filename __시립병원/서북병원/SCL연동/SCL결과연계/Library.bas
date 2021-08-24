Attribute VB_Name = "Library"
Option Explicit

Public Type PatGen
    Age As String
    Sex As String
End Type
Public gPatGen As PatGen

Declare Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As Long) As Long
Declare Function ImmGetConversionStatus Lib "imm32.dll" (ByVal hIMC As Long, lpdw As Long, lpdw2 As Long) As Long
Declare Function ImmSetConversionStatus Lib "imm32.dll" (ByVal hIMC As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long

' -------------------------------------------------------------
' 윈도우 API 함수 선언
' -------------------------------------------------------------
' //작업 Window 지정
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

' //윈도우 활성화
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

' //지정한 핸들값을 가진 윈도우의 스타일을 반환
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

' //Parenet Window를 반환
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

' //지정한 Handle값을 가진 윈도우에 Child Window로 지정
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

' //위도위의 위치, 크기, 순서 변경 API
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


' //새로운 Registry Key를 생성한다
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

' //열려있는 Registry Key를 닫는다
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

' //Registry에 등록된 Sub Key를 원하는 Access형태로 연다.
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

' //열려있는 Registry Key의 정보중에서 특정한 정보의 Data를 읽어드린다.
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
        
'Binary Registry
Private Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
        
' //지정된 Registry Key의 값을 설정한다.
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

' //지정된 Registry Key을 삭제한다
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

'실행
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' ------------------------------------------------------------
' * SetWindowPos API 함수 Enum 상수
' ------------------------------------------------------------
Public Enum udtMOST
    E_HWND_TOPMOST = -1
    E_HWND_NOTOPMOST = -2
End Enum

' ------------------------------------------------------------
' * Registry API 함수 Enum 상수
' ------------------------------------------------------------
Public Enum udtHKEY
    HKEY_CURRENT_USER = &H80000001
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum

Public Enum enumREGTYPE
    REG_NONE = 0
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
    REG_DWORD = 4
    REG_DWORD_LITTLE_ENDIAN = 4
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7
End Enum

Public Function fnActiveFormIsAppoint(ByVal hwnd As Long, Optional pMaximized As Boolean = True)
    Dim frmObject As Form   ' Form Object
'    Dim hWndMDI As Long     ' Main Form 핸들값
    
    '* 윈도우를 최대화 시킨다.
    Call fnForeGroundWindow(hwnd)
                
    '* 최상위 윈도우로 지정한다.
    If pMaximized = True Then
        Call fnMostWindowPosition(hwnd, E_HWND_TOPMOST)
    End If
                
End Function

' *****************************************************************************
' Purpose       : 최상위 윈도우로 지정하고 해지 한다.
'
' Description   : Enum 상수값이 E_HWND_TOPMOST(-1)면 최상위 윈도우로 설정하고
'                 E_HWND_NOTOPMOST(-2)면 최상위 윈도우 설정을 해제 한다
'
' Inputs        : hWnd - 윈도우 Handle,
'                 pHWAND_MOST - Enum 상수값: E_HWND_TOPMOST  (최상위 윈도우설정)
'                                            E_HWND_NOTOPMOST(최상위 윈도우해지)
' Ouptus        :
' Asserts       : 사용 API - SetWindowPos()
'
' -----------------------------------------------------------------------------
' Developer     Date        Comments
' -----------------------------------------------------------------------------
' 김성환        2002.8.1    최최작성자
' *****************************************************************************
Public Function fnMostWindowPosition(ByVal hwnd As Long, ByVal pHWAND_MOST As udtMOST)
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    
    ' 위도위의 위치, 크기, 순서 변경 API
    SetWindowPos hwnd, pHWAND_MOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Function

Public Function fnForeGroundWindow(ByVal hwnd As Long)
    Dim lngStyle As Long    ' Style값 Variables
    
    ' 지정한 핸들값을 가진 윈도우의 스타일을 반환한다
    lngStyle = GetWindowLong(hwnd, 2)
    
    '* 0이면 활성화 상태가 아니므로 활성화 시킨다.
    If lngStyle = 0 Then
        '* 윈도우를 활성화 시킨다
        Call ShowWindow(hwnd, 3)
    End If
    
    ' 폼을 작업 윈도우로 만든다.
    Call SetForegroundWindow(hwnd)
End Function

Public Function Data2Pict(sPrmData As String, sPrmPict As String) As String

    Dim i As Integer, iDataPos As Integer
    Dim iDataLen As Integer, iPictLen As Integer
    Dim sBufData As String, sPictStr As String, sChar As String

    iDataLen = Len(sPrmData)
    iPictLen = Len(sPrmPict)
    iDataPos = iDataLen
    sBufData = ""
    
    If iDataLen = 0 Or sPrmData = "0" Then
        If Right(sPrmPict, 1) = "0" Then
            Data2Pict = "0"
        Else
            Data2Pict = ""
        End If
        Exit Function
    End If

    For i = iPictLen To 1 Step -1
        sPictStr = ""

        Select Case Mid(sPrmPict, i, 1)
        Case "0", "9"
            sPictStr = Mid(sPrmData, iDataPos, 1)
            If Not IsNumeric(sPictStr) Then
                sPictStr = ""
                i = i + 1
            End If
            iDataPos = iDataPos - 1

        'Case ",", "."
        '    iDataPos = iDataPos - 1

        Case "X"
            sPictStr = Mid(sPrmData, iDataPos, 1)
            iDataPos = iDataPos - 1

        Case Else
            sPictStr = Mid(sPrmPict, i, 1)

        End Select

        sBufData = sPictStr & sBufData

        If iDataPos <= 0 Then
            Exit For
        End If
    Next

    If Left(LTrim(sPrmData), 1) = "-" Then
        sChar = Left(LTrim(sPrmPict), 1)
        Select Case sChar
        Case "-"
            If Left(LTrim(sBufData), 1) = "," Then
                sBufData = sChar & Mid(sBufData, 2)
            Else
                sBufData = sChar & sBufData
            End If

        End Select
    End If

    Data2Pict = sBufData

End Function

Public Function GetDateFull() As String
'Server의 현재 날짜와 시간을 가져온다
'Return = 2000/09/02 10:00:00
'    SQL = "select convert(char(10),getdate(),111) + ' ' + convert(char(8),getdate(),108) "
    
'Oracle
    SQL = "Select To_Char(SysDate, 'mm/dd/yyyy hh24:mi:ss') From Dual"
    db_select_Var SQL, GetDateFull
End Function

Public Function GetDateShort() As String
'Server의 현재 날짜을 가져온다
'Return = 2000/09/02
    SQL = "select convert(char(10),getdate(),111) "

'Oracle
'    SQL = " Select To_Char(SysDate, 'mm/dd/yyyy') From Dual "
    db_select_Var SQL, GetDateShort
End Function

Public Function GetTimeFull() As String
'Server의 현재 시간을 가져온다
'Return = 10:00:00
    SQL = "select convert(char(8),getdate(),108) "

'Oracle
'    SQL = " Select To_Char(SysDate, 'hh24:mi:ss') From Dual "
    db_select_Var SQL, GetTimeFull
End Function

Public Function CR() As String
    CR = Chr(13) & Chr(10)
End Function

Public Function GetText(ByRef vasTable As Object, ByVal vasRow As Integer, ByVal vasCol As Integer) As String
'vsSpread에서 데이타 가져오기
    vasTable.Row = vasRow
    vasTable.col = vasCol
    GetText = vasTable.Text
End Function

Public Function SetText(ByRef vasTable As Object, ByVal SetStr As String, ByVal vasRow As Integer, ByVal vasCol As Integer) As Boolean
'vsSpread에 데이타 넣기
    vasTable.Row = vasRow
    vasTable.col = vasCol
    vasTable.Text = SetStr
End Function

Public Sub ClearSpread(ByRef vasTable As Object)
'vsSpread의 내용을 Clear 한다.
    vasTable.Row = 1
    vasTable.col = 0
    vasTable.Row2 = vasTable.DataRowCnt
    vasTable.Col2 = vasTable.DataColCnt
    vasTable.BlockMode = True
    vasTable.Action = 3
    vasTable.BlockMode = False
End Sub

Public Function vasActiveCell(ByRef vasTable As Object, ByVal vasRow As Integer, ByVal vasCol As Integer) As Boolean
'특정 Cell 지정
    vasTable.Row = vasRow
    vasTable.col = vasCol
    vasTable.Action = 0
End Function

Public Function GetCurRow(ByRef vasTable As Object) As Integer
'현재 Active 된 Row 가져온다
    GetCurRow = vasTable.ActiveRow
End Function

Public Function GetCurCol(ByRef vasTable As Object) As Integer
'현재 Active 된 Col 가져온다
    GetCurCol = vasTable.ActiveCol
End Function

Public Function GetDataRowCnt(ByRef vasTable As Object) As Integer
'SpreadSheet에 들어있는 Data의 RowCount 가져온다
    GetDataRowCnt = vasTable.DataRowCnt
End Function

Public Function GetMaxRow(ByRef vasTable As Object) As Integer
'vaSpread의MaxRow를 가져온다
    GetMaxRow = vasTable.MaxRows
End Function

Public Sub InsertRow(ByVal vasTable As Object, argRow As Integer)
    vasTable.MaxRows = vasTable.MaxRows + 1
    vasTable.Row = argRow
    vasTable.Action = 7
End Sub

Public Sub InsertRow_1(ByVal vasTable As Object, argRow As Integer, addRow As Integer)
    vasTable.MaxRows = vasTable.MaxRows + addRow
    vasTable.Row = argRow
    vasTable.Action = 7
End Sub


Public Sub vasDeleteRow(ByVal vasTable As Object, argRow As Integer)
'
    vasTable.Row = argRow
    vasTable.Action = 5
End Sub

Public Function vasSort(ByRef vasTable As Object, ByVal key1 As Integer, Optional key2 As Integer = 0, Optional key3 As Integer = 0, Optional key4 As Integer = 0, Optional key5 As Integer = 0) As Boolean
'정렬할 부분의 선택
    vasTable.Row = 0
    vasTable.col = 0
    vasTable.Row2 = vasTable.DataRowCnt
    vasTable.Col2 = vasTable.DataColCnt
'정렬을 Row로 실시
    vasTable.SortBy = 2 'SS_SORT_BY_ROW
'정렬 키를 선택
    vasTable.SortKey(1) = key1
    vasTable.SortKeyOrder(1) = 1 'SS_SORT_ORDER_ASCENDING

    vasTable.SortKey(2) = key2
    If (key2 = 0) Then
        vasTable.SortKeyOrder(2) = 0
    Else
        vasTable.SortKeyOrder(2) = 1
    End If

    vasTable.SortKey(3) = key3
    If (key3 = 0) Then
        vasTable.SortKeyOrder(3) = 0
    Else
        vasTable.SortKeyOrder(3) = 1
    End If

    vasTable.SortKey(4) = key4
    If (key4 = 0) Then
        vasTable.SortKeyOrder(4) = 0
    Else
        vasTable.SortKeyOrder(4) = 1
    End If

    vasTable.SortKey(5) = key5
    If (key5 = 0) Then
        vasTable.SortKeyOrder(5) = 0
    Else
        vasTable.SortKeyOrder(5) = 1
    End If
'정렬
    vasTable.Action = 25 'SS_ACTION_SORT

    vasActiveCell vasTable, 1, 1
End Function

Public Function ScanCol(ByRef Obj As Object, ByVal SearchStr As String, _
                        ByVal ColPos As Integer, Optional StartRow = 1) As Integer
'SpreadSheetd의 Col에 있는것과 같은 Text를 찾아낸다.
'Return : 같은 Text가 존재하면 그 Col,
'                     존재하지 않으면 -1 을 반환
    Dim i As Integer
    Dim ChkData As String

    For i = StartRow To Obj.DataRowCnt
        ChkData = GetText(Obj, i, ColPos)
        If Trim(ChkData) = Trim(SearchStr) Then
            ScanCol = i
            Exit Function
        End If
    Next i
    
    ScanCol = -1
End Function

Public Function ScanRow(ByRef Obj As Object, ByVal SearchStr As String, _
                        ByVal RowPos As Integer, Optional StartCol = 1) As Integer
'SpreadSheetd의 Row에 있는것과 같은 Text를 찾아낸다.
'Return : 같은 Text가 존재하면 그 Row,
'                     존재하지 않으면 -1 을 반환
    Dim i As Integer
    Dim ChkData As String
    
    For i = StartCol To Obj.DataColCnt
        ChkData = GetText(Obj, i, RowPos)
        If Trim(ChkData) = Trim(SearchStr) Then
            ScanRow = i
            Exit Function
        End If
    Next i
    
    ScanRow = -1
End Function

Public Sub SelectFocus(ByRef argObj As Object)
'GetFocus 시 Object내의 Text가 전체 선택 되게 한다.
    'argObj.SetFocus
    argObj.SelStart = 0
    argObj.SelLength = Len(argObj.Text)
End Sub

Public Sub CenterForm(frmMe As Form)
'Form을 화면중앙에 위치하도록한다
    frmMe.Left = (Screen.Width - frmMe.Width) / 2
    frmMe.Top = (Screen.Height - frmMe.Height) / 2
End Sub

Public Function SearchErrorCheck(ByVal SchStartDate As String, ByVal SchEndDate As String) As Integer
'조회시 시작일자가 종료일자보다 작거나 같은지 체크
    Dim msg As String
    
    If DateDiff("s", SchStartDate, SchEndDate) < 0 Then
        msg = DateDiff("d", SchEndDate, SchStartDate) & "s 차이 : 조회 할 날짜를 잘못 설정 하셨습니다"
        MsgBox msg, , "알림"
        SearchErrorCheck = -1
        Exit Function
    End If
    SearchErrorCheck = 0
End Function

Public Function SeperatorCls(asStr As String) As String
'숫자외의 구분자를 모두 없앤다
    Dim i       As Integer
    Dim StrLen  As Integer
    Dim RtStr   As String
    
    RtStr = ""

    For i = 1 To Len(asStr)
        If IsNumeric(Mid(asStr, i, 1)) Then
            RtStr = RtStr & Mid(asStr, i, 1)
        End If
    Next i
    
    SeperatorCls = RtStr
End Function

Public Sub SetIME(h As Long, Toggle As Boolean)
'                 h:폼 핸들, Toggle:한/영(true/false)
'====================================================
'   한글로 변환    Call SetIME(Form1.hWnd, True)
'   영어로 변환    Call SetIME(Form1.hWnd, False)
'====================================================
    Dim hIMC As Long
    Dim dwConversion As Long, dwSentence As Long
    Dim Temp As Long '
    
    hIMC = ImmGetContext(h)
    Temp = ImmGetConversionStatus(hIMC, dwConversion, dwSentence)
    If Toggle Then
        dwConversion = dwConversion Or 1
        Temp = ImmSetConversionStatus(hIMC, dwConversion, dwSentence)
    Else
        dwConversion = dwConversion And -2&
        Temp = ImmSetConversionStatus(hIMC, dwConversion, dwSentence)
    End If
End Sub

Public Sub CalAgeSex(ByRef asPNRN As String, ByRef asCurDate As String)
    Dim sBirth As String
    Dim sStart As String
    
    If Mid(asPNRN, 1, 1) = "_" Or Mid(asPNRN, 1, 1) = "" Then
        Exit Sub
    End If
    
    gPatGen.Sex = ""
    gPatGen.Age = ""
    
    asPNRN = SeperatorCls(asPNRN)
    
    sStart = Trim(Mid(Trim(asPNRN), 7, 1))
    sBirth = ""
    
    Select Case sStart
        Case "1", "3", "5", "7"
            gPatGen.Sex = "M"
        Case "2", "4", "6", "8"
            gPatGen.Sex = "F"
    End Select

    Select Case sStart
        Case "1", "2"
            sBirth = "19"
        Case "3", "4"
            sBirth = "20"
        Case "7", "8"
            sBirth = "18"
        Case Else
            sBirth = "19"
    End Select
    
'    sBirth = ""
    sBirth = sBirth & Mid(asPNRN, 1, 2) '& "/" & Mid(asPNRN, 3, 2) & "/" & Mid(asPNRN, 5, 2)
    'If Mid(asPNRN, 3, 2) = "00" Then
        sBirth = sBirth & "/01"
    'Else
    '    sBirth = sBirth & "/" & Mid(asPNRN, 3, 2)
    'End If
    'If Mid(asPNRN, 5, 2) = "00" Then
        sBirth = sBirth & "/01"
    'Else
    '    sBirth = sBirth & "/" & Mid(asPNRN, 5, 2)
    'End If
    
    gPatGen.Age = DateDiff("yyyy", sBirth, asCurDate) + 1
End Sub

Public Function ChangeSex(ByRef asSex As String) As String
    Select Case Trim(asSex)
        Case "남"
            ChangeSex = "M"
        Case "여"
            ChangeSex = "F"
        Case "M"
            ChangeSex = "남"
        Case "F"
            ChangeSex = "여"
        Case Else
            ChangeSex = ""
    End Select
End Function

Public Function NLeftString(ByVal arg_s As String, _
                             ByVal size As Integer)

    Dim i As Integer
    Dim Temp As String
    
    NLeftString = arg_s
    
    If (prnStrlen(arg_s) > size) Then
        For i = Len(arg_s) - 1 To 0 Step -1
            If prnStrlen(Mid$(arg_s, 1, i)) <= size Then
                NLeftString = Mid$(arg_s, 1, i)
                Exit For
            End If
        Next i
    End If
    
    If (Len(NLeftString) = 0) Then
        NLeftString = ""
    End If
    
    NLeftString = NLeftString + Space(size - prnStrlen(NLeftString))
    
End Function


Public Function NMidString(ByVal arg_s As String, _
                             ByVal size As Integer)
    Dim i As Integer
    Dim Temp As String
    Dim addsize As Integer
    Dim h_addsize As Integer
    
    NMidString = arg_s
    
    If (prnStrlen(arg_s) > size) Then
        For i = Len(arg_s) - 1 To 0 Step -1
            If prnStrlen(Mid$(arg_s, 1, i)) <= size Then
                NMidString = Mid$(arg_s, 1, i)
                Exit For
            End If
        Next i
    End If
    
    If (Len(NMidString) = 0) Then
        NMidString = ""
    End If
    
    addsize = size - prnStrlen(NMidString)
    h_addsize = Int(addsize / 2)
    NMidString = Space(addsize - h_addsize) + NMidString + Space(h_addsize)
End Function

Public Function NRightString(ByVal arg_s As String, _
                             ByVal size As Integer)
    Dim i As Integer
    Dim Temp As String
    
    NRightString = arg_s
    
    If (prnStrlen(arg_s) > size) Then
        For i = Len(arg_s) - 1 To 0 Step -1
            If prnStrlen(Mid$(arg_s, 1, i)) <= size Then
                NRightString = Mid$(arg_s, 1, i)
                Exit For
            End If
        Next i
    End If
    
    If (Len(NRightString) = 0) Then
        NRightString = ""
    End If
    
    NRightString = Space(size - prnStrlen(NRightString)) + NRightString
    
End Function


Public Function prnStrlen(ByVal arg_s As String)

    Dim i As Integer
    
    prnStrlen = 0
    For i = 1 To Len(arg_s)
        If (Mid$(arg_s, i, 1) > "가") Then
            prnStrlen = prnStrlen + 2
        Else
            prnStrlen = prnStrlen + 1
        End If
    Next i
    
End Function

Public Function IsolateCode(argAll As String)
    Dim i As Integer
    Dim sCode, sName As String
    
    If argAll = "" Then
        gCode = ""
        gName = ""
        Exit Function
    End If
    
    sCode = ""
    sName = ""
    
    i = InStr(1, argAll, " ")
    
    If i = 0 Then
        gCode = Trim(argAll)
        gName = ""
    Else
        gCode = Trim(Left(argAll, i))
        gName = Trim(Mid(argAll, i))
    End If
End Function

Public Sub SetBackColor_Vas(argSpread As vaSpread, argRow1 As Integer, argRow2 As Integer, _
                            argCol1 As Integer, argCol2 As Integer)

    argSpread.Row = argRow1
    argSpread.Row2 = argRow2
    argSpread.col = argCol1
    argSpread.Col2 = argCol2
    argSpread.BlockMode = True
    argSpread.BackColor = RGB(255, 254, 236)
    argSpread.BlockMode = False
    
End Sub

Sub SetBackColor(asTable As vaSpread, ByVal asRow1 As Long, ByVal asRow2 As Long, asCol1 As Long, asCol2 As Long, asR As Variant, asG As Variant, asB As Variant)
    asTable.Row = asRow1
    asTable.Row2 = asRow2
    asTable.col = asCol1
    asTable.Col2 = asCol2
    asTable.BlockMode = True
    asTable.BackColor = RGB(asR, asG, asB)
    asTable.BlockMode = False
End Sub

Public Function SetChar(asStr As String, asLen As Integer, Optional asPos As Integer = 1, Optional asChar As String = " ") As String
'asPos = 1 : Left 공백
'asPos = 2 : Right 공백 채우기
    Dim sTmp As String
    Dim i As Integer
    
    sTmp = ""
    If Len(asStr) >= asLen Then
        SetChar = Left(asStr, asLen)
        Exit Function
    End If
    
    sTmp = asStr
    For i = 1 To asLen - Len(asStr)
        If asPos = 1 Then
            sTmp = asChar & sTmp
        Else
            sTmp = sTmp & asChar
        End If
    Next i
    
    SetChar = sTmp
End Function

Public Function SetSpace(asStr As String, asLen As Integer, Optional asPos As Integer = 1, Optional asSet As String = " ") As String
'asPos = 1 : Left 공백
'asPos = 2 : Right 공백 채우기
    Dim sTmp As String
    Dim i As Integer
    
    sTmp = ""
    If Len(asStr) >= asLen Then
        SetSpace = Left(asStr, asLen)
        Exit Function
    End If
    
    sTmp = asStr
    For i = 1 To asLen - Len(asStr)
        If asPos = 1 Then
            sTmp = asSet & sTmp
        Else
            sTmp = sTmp & asSet
        End If
    Next i
    
    SetSpace = sTmp
End Function

Sub Save_Raw_Data(asData As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.Path & "\Log")
    End If
    
    sFileName = Format(Date, "yyyymmdd")
    
    Open App.Path & "\Log\" & sFileName & ".txt" For Append As FilNum
'    Print #FilNum, Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:nn:ss") & " " & asData
    Print #FilNum, Format(Time, "hh:nn:ss") & " " & asData
    Close FilNum
End Sub
