Attribute VB_Name = "common"
Option Explicit

Global bOwn100Yno As Integer   '전액본인부담 여부

#Const HOSP_NAME = "UIWANG"

'Declare Sub SetWindowPos Lib "USER" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'--> 사용 Call SetWindowPos(Me.hwnd, -1, Me.Left, Me.Top, Me.Width, Me.Heizght, &H43)
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long


Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

'example by Donavon Kuhn (Donavon.Kuhn@Nextel.com)
Public Const MAX_COMPUTERNAME_LENGTH As Long = 31
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'----------------
'마우스 이동 제한
'----------------
'Public Declare Function ClientToScreen& Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI)
'Public Declare Function ClipCursor& Lib "user32" (lpRect As RECT)
'Public Declare Function ClipCursorBynum& Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long)
'----------------
Public Const WM_LBUTTONUP = &H202

'한영 토글
Public Const IME_HANGUL = &H1
Public Const IME_ENGLISH = &H0
Public Const IME_NONE = &H0

Declare Function ImmGetContext Lib "imm32.dll" (ByVal hWnd As Long) As Long
Declare Function ImmSetConversionStatus Lib "imm32.dll" (ByVal hIMC As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long

'Declare Sub cvtToHan Lib "f:\hnt.prj\공통\CVT_IME.DLL" (ByVal hwnd As Integer)
'Declare Sub cvtToEng Lib "f:\hnt.prj\공통\CVT_IME.DLL" (ByVal hwnd As Integer)


'이건 절대 바꾸면 안됩니다
Global Const ExePath = "C:\Hnt.Cnv\exe\"      '실행화일 path

'New_Path 의 드라이브가 F 로 정해져 있기때문에
'3.0 과 6.0 을 동시에 잡았을때 문제가 생긴다
'따라서 자유롭게 변경 가능한 글로벌 변수 (New_New_Path) 를 선언하고
'LogIn 에서 함수를 실행시켜 값을 담아준다.

Global New_Path As String

'Global Const IconPath = "F:\HNT.PRJ\ICON\"    'Icon path
'Global Const New_Path = "..\exe\"      '실행화일 path
Global Const IconPath = "..\ICON\"    'Icon path
Global Const IConPathNew = "..\Hnt.cnv\ICon\사용중\임시\" 'Icon Path

Public bGblAppSecPowerUpdate As Boolean     'Application 에 대한 Update 권한
Public bGblAppSecPowerRead   As Boolean     'Application 에 대한 Read 권한.

Global Const KEY_HANENG = &H15
Global Const KEY_HANJA = &H19

Type MstHlpInf
    Cod As String
    Nam As String
End Type

Type Mst3HlpInf
    Cod As String * 10
    Data1 As String * 30
    Data2 As String * 30
End Type

'화면의 깜박임을 막기 위해 일괄로 Refresh한다.
Public Const WM_SETREDRAW = &HB

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public sGblImgFlag As String

Type POINTAPI
    x As Long
    Y As Long
End Type

Public Const BI_RGB = 0&
Public Const DIB_RGB_COLORS = 0 '  color table in RGBs

Public Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
Public Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type


Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Public Enum EnumgsPntOnSpd
        Auto
        LEFTTOP
        LeftMiddle
        LiftBottom
        MiddleTop
        Center
        MiddleBottom
        RIGHTTOP
        RightMiddle
        RightBottom
End Enum

Public Enum ENUM_MENU
    배경색변경 = 0
    글자색변경
    기본색으로
End Enum

'[2000년 11월 22일 개발과 김강수]
'다른 부분에 선언된 모든 UidMst을 지우세요...
'여기에 있는 것은 절대로 지우지 마세요...
Public tGblUidData As UidMstRec
Public tGblUidMst As UidMstRec

Public mgl_Result As Long          'Return 값
Public mgu_RetAbsLeft As Long      'Control 의 Left 절대좌표 반환값
Public mgu_RetAbsTop As Long       'Control 의 Top  절대좌표 반환값
Public mgu_SpdAbsCol As Long       '화면 상의 스프레드 좌표 - Col
Public mgu_SpdAbsRow As Long       '화면 상의 스프레드 좌표 - Row
Public mgu_CursPosChnged As Boolean   'KeyDown 다음 일어나는 MouseMove Event 여부
Public mgu_POINTAPI As POINTAPI

Public mgu_MousePointOnSpd As EnumgsPntOnSpd
Public mgu_MsPntOnSpd As EnumgsPntOnSpd

Public mgs_Parameter As String  '전체에서 공통으로 사용가능한 변수

Public sGblAccDte As String     '전체에서 공통으로 사용하는 회계일자 변수
Public sGblOffDuty As String    '전체적으로 사용하는 원무과 근무 구분(1.주간근무 2.당직(12시전) 3. 당직(12시이후)...등
Public aGblRst As ADODB.Recordset      '20040206..HTS..
Public aGblConn As ADODB.Connection

Public msRepeateDay As String

Public Function SystemTimeSec() As String

    mvbFrm.Mvb1.P0 = ""

    mvbFrm.Mvb1.Code = "d ^SystemTimeSec(.P0)"
    mvbFrm.Mvb1.ExecFlag = 1

    SystemTimeSec = mvbFrm.Mvb1.P0

End Function

'Public fbs_Tmp_Data(0 To 20, 0 To 4) As Variant   '원외처방
'Public fbs_InjData(0 To 20, 0 To 4) As String
'Global mgs_Head(HEAD_INSCOD To HEAD_CHTNUM) As String
Public Sub SaveAssCod(ByVal psAssCod As String, ByVal psAssNam As String, ByVal psInsCod As String)

    Dim i As Integer
    Dim sCurKey As String
    Dim sRetVal As String
    
    Dim AssData As AssMstRec
    
    AssData.AssCod = psAssCod
    AssData.AssInsCod = psInsCod
    AssData.AssCodNam = psAssNam
    
    Call AssMstStore(sCurKey, sRetVal, AssData)
    
    i = mWrite("AssMst", sCurKey, sRetVal)

End Sub

'메뉴를 생성할때 항상 인덱스 "0"에는 배경색변경
'                            "1"에는 글자색변경
'                            "2"에는 기본색으로    를 생성한다.
'그리고 mnuColor_Click 이벤트에서 call만 하면...
'Public Sub SaveSetting_SpreadColor(Index As Integer, objSpread As Object, defaultBackColor As Single, defaultForeColor As Single)
'
'    With mvbFrm.cmDialog
'
'On Error GoTo ErrHandler
'
'        If Index = 기본색으로 Then
'            objSpread.ShadowColor = defaultBackColor
'            objSpread.ShadowText = defaultForeColor
'        Else
'            If Index = 배경색변경 Then       '배경색 변경
'                ' 폼의 배경색을 선택한 색으로 설정합니다.
'                .Color = objSpread.ShadowColor
'            Else                   '글자색 변경
'                .Color = objSpread.ShadowText
'            End If
'
'            .DialogTitle = "색변경"
'            ' Flags 속성을 설정합니다.
'            .flags = cdlCCRGBInit
'            ' [색] 대화 상자를 표시합니다.
'            .ShowColor
'
'            If Index = 배경색변경 Then       '배경색 변경
'                ' 폼의 배경색을 선택한 색으로 설정합니다.
'                objSpread.ShadowColor = .Color
'            Else                   '글자색 변경
'                objSpread.ShadowText = .Color
'            End If
'        End If
'
'        SaveSetting "Hnt.Prj", App.EXEName, "ShadowColor", objSpread.ShadowColor
'        SaveSetting "Hnt.Prj", App.EXEName, "ShadowText", objSpread.ShadowText
'
'    End With
'
'    Exit Sub
'
'ErrHandler:
'
'End Sub

Public Sub WriteIctInf_FromispInf(sPrmOcmNum)

    Dim sCurKey As String
    Dim sCmpKey As String
    
    Dim sRetVal As String
    
    Dim sIctCurKey As String
    Dim sIctRetVal As String
    
    Dim IspData As IspInfRec
    Dim IctData As IctInfRec
        
    sCurKey = Format(sPrmOcmNum, "@@@@@@@@@@") & Chr(5)
    sCmpKey = sCurKey
    sCurKey = mSetNext("IspInf", sCurKey)
    Do
        sCurKey = mReadNext("IspInf", sCurKey, sCmpKey, sRetVal)
        If sCurKey <> "" Then
            Call IspInfLoad(sRetVal, IspData)
            
            Call IspInfStore(sIctCurKey, sIctRetVal, IspData)
            If Not mWrite("IctInf", sIctCurKey, sIctRetVal) Then
                If Not mUpdate("IctInf", sIctCurKey, sIctRetVal) Then
                    MsgBox "IctInf Write Error"
                End If
            End If
            
        Else
            Exit Do
        End If
    
    Loop
    
    MsgBox "심사 자료모음이 완료되었습니다.", vbInformation + vbOKOnly
    
End Sub

Public Sub gFormControlEmpty(frm As Form)
    On Error Resume Next
    Dim CurControl As Control
    
    For Each CurControl In frm.Controls
        Select Case TypeName(CurControl)
            Case "ComboBox":
                If CurControl.Style = vbComboDropdownList Then
                    CurControl.ListIndex = 0
                Else
                    CurControl.Text = ""
                End If
            Case "TextBox":
                CurControl.Text = ""
        End Select
    Next CurControl
    
End Sub

Public Sub gFormEnableTrueFalse(frm As Form, bState As Boolean)
' 폼의 ComboBox, TextBox, MaskEdBox, CheckBox, OptionButton의 Enabled를 True/False로
    Dim CurControl As Control
    For Each CurControl In frm.Controls
        If TypeOf CurControl Is ComboBox Or TypeOf CurControl Is TextBox _
            Or TypeOf CurControl Is CheckBox _
                Or TypeOf CurControl Is OptionButton Then
            If bState = False Then
                CurControl.Enabled = False
            Else
                CurControl.Enabled = True
            End If
        End If
    Next CurControl
End Sub

Public Function CRoundingNew(sPrmValue As Variant, sPrmPosition As Variant) As Double

    '-------------------------------------------------
    '파라미터
    '-------------------------------------------------
    'sPrmValue   : 원본 숫자
    'sPrmPosition : 사사오입하는 위치
    '               -----------
    '               방법  : Po
    '               -----------
    '                     (소숫정세째자리) : 0.001  --> Drg금액을 구할때
    '               전미만(소숫점둘째자리) : 0.01
    '               원미만(소숫점첫째자리) : 0.1     -->보통 총진료비구할때.
    '                   10원미만(첫째자리) : 1      -->보통 돈으로 받을때.
    '                  100원미만(둘째자리) : 10
    '-------------------------------------------------
    'ex)
    'Po: 첫째자리 일때
    'CRounding = CLng(((sPrmValue + 5) * 10 \ 10))
    '-------------------------------------------------
    'Return      : 사사오입된 숫자
    '-------------------------------------------------

    Dim lRoundValue As Long
    Dim dInt As Double

    lRoundValue = CLng(sPrmPosition)

On Error GoTo CR_HandlerNew

    dInt = sPrmPosition * 10

    CRoundingNew = Fix((Val(sPrmValue) / dInt) + 0.5) * dInt

    Exit Function

CR_HandlerNew:
    CRoundingNew = 0
    Resume Next


End Function

Public Sub OcmInfChtRead(sChtNum As String, sTmpDte As String, tOcmData As OcmInfRec)

    Dim sCurKey As String, sCmpKey As String, sRetVal As String
    Dim tTmpData As OcmInfRec

    sCmpKey = Format(Trim(sChtNum), "@@@@@@@@") & Chr(5)
    sCurKey = sCmpKey & sTmpDte & "9999"
    sCurKey = mSetPrev("OcmInfChtDtm", sCurKey)
    Do
        sCurKey = mReadPrev("OcmInfChtDtm", sCurKey, sCmpKey, sRetVal)
        If sCurKey = "" Then Exit Do

        OcmInfLoad sRetVal, tTmpData
        If tTmpData.OcmComStt <> "OC" Then
            tOcmData = tTmpData
            Exit Sub
        End If
    Loop
    
End Sub

Public Function FinalNumberSetting_Month(sPrmFnlCod As String, sPrmDate As String) As String

    Dim i As Integer
    Dim FnlData As FnlMstRec
    Dim sBufKey As String
    Dim sBufValue As String
    Dim iTmpSeq As Integer

    Dim sFnlMstCurKey As String
    Dim sFnlMstCmpKey As String
    Dim sPrmRetVal As String

    'Locking Routine (mWrite 의 return값은 True or False)
    For i = 1 To 10000          '10000번 test
        If mWrite("LckMst", sPrmFnlCod, sPrmFnlCod) Then Exit For
    Next

    FnlData.FnlCod = sPrmFnlCod
    FnlMstStore sBufKey, sBufValue, FnlData
    sFnlMstCurKey = sBufKey

    sBufValue = mSetReadEqual("FnlMst", sFnlMstCurKey, sPrmRetVal)

    If sBufValue <> "" Then
        Call FnlMstLoad(sPrmRetVal, FnlData)
        
        Select Case sPrmFnlCod
        Case "DRGEDI"
            If sPrmDate = FnlData.FnlDte Then
                FnlData.FnlNum = CStr(CLong(FnlData.FnlNum) + 1)
                FnlData.FnlDte = sPrmDate
                FinalNumberSetting_Month = FnlData.FnlNum
            Else
                FnlData.FnlNum = "1"
                FnlData.FnlDte = sPrmDate
                FinalNumberSetting_Month = "1"
            End If

         Case Else
            FnlData.FnlNum = CStr(CLong(FnlData.FnlNum) + 1)
            FinalNumberSetting_Month = FnlData.FnlNum
        End Select
    Else
        Select Case sPrmFnlCod
        Case "DRGEDI"
            FnlData.FnlNum = "1"
            FnlData.FnlDte = sPrmDate
            FnlData.FnlCod = sPrmFnlCod
            FinalNumberSetting_Month = "1"
        Case Else
            FnlData.FnlNum = "1"
            FnlData.FnlCod = sPrmFnlCod
            FinalNumberSetting_Month = FnlData.FnlNum
        End Select
    End If

    Call FnlMstStore(sBufKey, sBufValue, FnlData)

    iTmpSeq = mWrite("FnlMst", sBufKey, sBufValue)
    If iTmpSeq = False Then
        iTmpSeq = mUpdate("FnlMst", sBufKey, sBufValue)
    End If

    'Locking 해제
    iTmpSeq = mDelete("LckMst", sPrmFnlCod)

End Function

Public Sub OcmInfChtReadToday(sChtNum As String, sTmpDte As String, tOcmData As OcmInfRec)

    Dim sCurKey As String, sCmpKey As String, sRetVal As String
    Dim tTmpData As OcmInfRec

    sCmpKey = Format(Trim(sChtNum), "@@@@@@@@") & Chr(5)
    sCurKey = sCmpKey & sTmpDte & "9999"
    sCurKey = mSetPrev("OcmInfChtDtm", sCurKey)
    Do
        sCurKey = mReadPrev("OcmInfChtDtm", sCurKey, sCmpKey, sRetVal)
        If sCurKey = "" Then Exit Do

        Call OcmInfLoad(sRetVal, tTmpData)
        If tTmpData.OcmComStt <> "OC" Then
            If Left(tTmpData.OcmAcpDtm, 8) = sTmpDte Then
                tOcmData = tTmpData
                Exit Sub
            End If
        End If
    Loop
    
End Sub

Public Sub Refresh_Path()
    
    '2001/08/30
    'LogIn 이 필요하지 않은 프로그램에선 이 프로시저를 실행 시킨 다음
    '기존의 ExePath 를 New_Path 로 바꾸어 주면 된다.
    
    'Dim sTmp As String
    'sTmp = Left$(CurDir, 1)
    'New_Path = sTmp & Right(ExePath, (Len(ExePath) - 1))
    
    Dim i As Integer
    Dim sTmp As String
    
    If Dir(ExePath) = "" Then
        
        For i = Asc("G") To Asc("Z")
            sTmp = Chr(i) & Mid(ExePath, 2, (Len(ExePath) - 1))
            If Dir(sTmp) <> "" Then
                New_Path = sTmp
                Exit For
            End If
        Next i
    Else
        New_Path = ExePath
        
    End If
End Sub


'vDefaultMaxRows : 화면에 Display 할 값이 없어도 Default값 만큼은 화면에 ""로 보여주자
Public Sub Spread_Clear(Obj As Object, vMaxRows As Variant, Optional vDefaultMaxRows As Variant = 0)

    With Obj
        If vDefaultMaxRows = 0 Then
            .MaxRows = vMaxRows
        Else
            If vDefaultMaxRows < vMaxRows Then
                .MaxRows = vMaxRows
            Else
                .MaxRows = vDefaultMaxRows
            End If
        End If
        
        If vMaxRows < 0 Then
            .Tag = 0
        Else
            .Tag = vMaxRows
        End If
        
        If .MaxRows > 0 Then
            .Row = 1
            .Row2 = .MaxRows
            .col = 1
            .Col2 = .MaxCols
            .BlockMode = True
            .Action = ActionClearText
            .BlockMode = False
        End If
    End With

End Sub

Public Sub Spread_BorderLine_Setting(mpc_Source As Control, fpl_RowCol As Long, Optional fpb_IsSign As Boolean = True, Optional fpb_Vertical As Boolean = True, Optional fps_Color As Single = vbBlack, Optional fpi_CellBorderStyle As Integer = CellBorderStyleSolid, Optional fpb_IsRightorBottom As Boolean = True, Optional fpb_IsAll As Boolean = False)

    Dim i As Integer

    With mpc_Source
        If fpb_IsAll = False Then
            If fpb_Vertical Then
                .Row = 0
                .Row2 = .MaxRows
            Else
                .Row = fpl_RowCol
                .Row2 = fpl_RowCol
            End If

            If fpb_Vertical Then
                .col = fpl_RowCol
                .Col2 = fpl_RowCol
            Else
                .col = 0
                .Col2 = .MaxCols
            End If

            .BlockMode = True
            If fpb_IsSign = True Then
                .CellBorderStyle = fpi_CellBorderStyle

                If fpb_Vertical Then
                    If fpb_IsRightorBottom = True Then
                        .CellBorderType = 2     'Right  Displays the border on the right
                    Else
                        .CellBorderType = 1     'Left
                    End If
                    .CellBorderColor = fps_Color
                Else
                    If fpb_IsRightorBottom = True Then
                        .CellBorderType = 8     'SS_BORDER_TYPE_BOTTOM
                    Else
                        .CellBorderType = 4     'Top
                    End If
                    .CellBorderColor = fps_Color
                End If
            Else
                .CellBorderStyle = CellBorderStyleBlank
                .CellBorderType = 0
            End If

            .Action = ActionSetCellBorder
            .BlockMode = False
        Else
            .Row = 1
            .Row2 = .MaxRows
            .col = 1
            .Col2 = .MaxCols
            .BlockMode = True
            If fpb_IsSign = True Then
                .CellBorderStyle = fpi_CellBorderStyle

                If fpb_Vertical Then
                    If fpb_IsRightorBottom = True Then
                        .CellBorderType = 2     'Right  Displays the border on the right
                    Else
                        .CellBorderType = 1     'Left
                    End If
                    .CellBorderColor = fps_Color
                Else
                    If fpb_IsRightorBottom = True Then
                        .CellBorderType = 8     'SS_BORDER_TYPE_BOTTOM
                    Else
                        .CellBorderType = 4     'Top
                    End If
                    .CellBorderColor = fps_Color
                End If
            Else
                .CellBorderStyle = CellBorderStyleBlank
                .CellBorderType = 0
            End If

            .Action = ActionSetCellBorder
            .BlockMode = False
        End If
    End With

End Sub

Public Sub Spread_BorderLine_Setting2(mpc_Source As Control, fpl_StrRowCol As Long, fpl_RowCol As Long, fpl_RowCol2 As Long, Optional fpb_IsSign As Boolean = True, Optional fpb_Vertical As Boolean = True, Optional fps_Color As Single = vbBlack, Optional fpi_CellBorderStyle As Integer = CellBorderStyleSolid, Optional fpb_IsRightorBottom As Boolean = True, Optional fpb_IsAll As Boolean = False)

    Dim i As Integer

    With mpc_Source
        If fpb_IsAll = False Then
            If fpb_Vertical Then
                .Row = fpl_RowCol
                .Row2 = fpl_RowCol2
            Else
                .Row = fpl_StrRowCol
                .Row2 = fpl_StrRowCol
            End If

            If fpb_Vertical Then
                .col = fpl_StrRowCol
                .Col2 = fpl_StrRowCol
            Else
                .col = fpl_RowCol
                .Col2 = fpl_RowCol2
            End If

            .BlockMode = True
            If fpb_IsSign = True Then
                .CellBorderStyle = fpi_CellBorderStyle

                If fpb_Vertical Then
                    If fpb_IsRightorBottom = True Then
                        .CellBorderType = 2     'Right  Displays the border on the right
                    Else
                        .CellBorderType = 1     'Left
                    End If
                    .CellBorderColor = fps_Color
                Else
                    If fpb_IsRightorBottom = True Then
                        .CellBorderType = 8     'SS_BORDER_TYPE_BOTTOM
                    Else
                        .CellBorderType = 4     'Top
                    End If
                    .CellBorderColor = fps_Color
                End If
            Else
                .CellBorderStyle = CellBorderStyleBlank
                .CellBorderType = 0
            End If

            .Action = ActionSetCellBorder
            .BlockMode = False
        Else
            .Row = 1
            .Row2 = .MaxRows
            .col = 1
            .Col2 = .MaxCols
            .BlockMode = True
            If fpb_IsSign = True Then
                .CellBorderStyle = fpi_CellBorderStyle

                If fpb_Vertical Then
                    If fpb_IsRightorBottom = True Then
                        .CellBorderType = 2     'Right  Displays the border on the right
                    Else
                        .CellBorderType = 1     'Left
                    End If
                    .CellBorderColor = fps_Color
                Else
                    If fpb_IsRightorBottom = True Then
                        .CellBorderType = 8     'SS_BORDER_TYPE_BOTTOM
                    Else
                        .CellBorderType = 4     'Top
                    End If
                    .CellBorderColor = fps_Color
                End If
            Else
                .CellBorderStyle = CellBorderStyleBlank
                .CellBorderType = 0
            End If

            .Action = ActionSetCellBorder
            .BlockMode = False
        End If
    End With

End Sub

'Spread Line의 Porperty를 복사한다.
Public Sub Spread_Property_Copy(mpc_Source As Control, ByVal mpl_Row As Long, ByVal mpl_CopyRow As Long)
    
    Dim fbs_MaxRowHeight As Single
    
    With mpc_Source
        '범위 Check
        If mpl_Row < 1 Or mpl_Row > .MaxRows Then Exit Sub
        
        'fbs_MaxRowHeight = .MaxTextRowHeight(mpl_CopyRow)
        fbs_MaxRowHeight = .RowHeight(mpl_CopyRow)
        
        'Line Height Setting
        .RowHeight(mpl_Row) = fbs_MaxRowHeight
        
        '이동할 범위를 정한다.
        .Row = mpl_CopyRow
        .Row2 = mpl_CopyRow
        
        .col = 1
        .Col2 = .MaxCols
        '해당 범위의 라인을 이동한다.
        .DestCol = 1
        .DestRow = mpl_Row
        .Action = ActionCopyRange
        
        .Row = mpl_Row
        .Row2 = mpl_Row
        .col = 1
        .Col2 = .MaxCols
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With

End Sub

Public Sub Spread_Lock_Protect(mpc_Source As Control, ByVal mpl_Row As Long, ByVal mpl_Row2 As Long, ByVal mpl_Col As Long, ByVal mpl_Col2 As Long, Optional mpb_Lock As Boolean = True)

    With mpc_Source
        If mpl_Row = -1 And mpl_Col = -1 Then
            .Row = mpl_Row
            .col = mpl_Col
            .Lock = mpb_Lock
            .Protect = mpb_Lock
        Else
            .Row = mpl_Row
            .col = mpl_Col
            .Row2 = mpl_Row2
            .Col2 = mpl_Col2
            .BlockMode = True
            .Lock = mpb_Lock
            .Protect = mpb_Lock
            .BlockMode = False
        End If
    End With
    
End Sub

Public Sub Spread_BackColor_Setting(mpc_Source As Control, fpl_Row1 As Long, fpl_Col1 As Long, Optional fpb_IsVertical As Boolean = False, Optional fpl_Row2 As Long = -1, Optional fpl_Col2 As Long = -1, Optional fpl_Color As Long = &HF4F4F4)

    With mpc_Source
        .Row = fpl_Row1
        .col = fpl_Col1
        If fpl_Row2 = -1 Then
            If fpb_IsVertical = True Then
                .Row2 = .MaxRows
            Else
                .Row2 = fpl_Row1
            End If
        Else
            .Row2 = fpl_Row2
        End If
        If fpl_Col2 = -1 Then
            If fpb_IsVertical = True Then
                .Col2 = fpl_Col1
            Else
                .Col2 = .MaxCols
            End If
        Else
            .Col2 = fpl_Col2
        End If
        .BlockMode = True
        .BackColor = fpl_Color
        .BlockMode = False
    End With
    
End Sub

'Spread Line Clear
Public Sub Spread_Line_Clear(mpc_Source As Control, ByVal Row As Long, Optional ByVal col As Long, Optional ByVal Col2 As Long)
    
    With mpc_Source
        .Row = Row
        .Row2 = Row
        
        If col = 0 Then             '없으면
            .col = 1
            .Col2 = .MaxCols
        ElseIf col = -1 Then        '0도 Clear
            .col = 0
            .Col2 = .MaxCols
        Else
            .col = col
            .Col2 = Col2
        End If
        
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
End Sub

'Spread Line Delete
Public Sub Spread_Line_Delete(mpc_Source As Control, ByVal mpl_Row As Long, Optional ByVal mpl_Col As Long = 1)

    With mpc_Source
        '범위 Check
        If mpl_Row < 1 Or mpl_Row > .MaxRows Then Exit Sub
        
        'Line을 Delete한다.
        .Row = mpl_Row
        .Action = ActionDeleteRow
        
        Call Spread_Property_Copy(mpc_Source, .MaxRows, 1)
    End With
    
End Sub

'Spread Line Insert
Public Sub Spread_Line_Insert(mpc_Source As Control, ByVal mpl_Row As Long)

    With mpc_Source
        '범위 Check
        If mpl_Row < 1 Or mpl_Row > .MaxRows Then Exit Sub
        
        'Line을 Delete한다.
        .Row = mpl_Row
        .Action = ActionInsertRow
            
        Call Spread_Property_Copy(mpc_Source, mpl_Row, .MaxRows)
    End With
    
End Sub

'Spread의 Focus 이동
Public Sub Spread_Move_Active_Cell(mpc_Source As Control, Optional ByVal Row As Long = 1, Optional ByVal col As Long = 1)
    
    With mpc_Source
        If Row < -1 Or col < -1 Then Exit Sub
        
        If .MaxRows >= Row Then
        
            .Row = Row
            .col = col
            .Action = ActionActiveCell
        End If
    End With
    
End Sub

'Registry에 저장된 색을 Spread에 설정한다. Form_Load에서 Call한다.
Public Sub GetSetting_SpreadColor(objSpread As Object)

    objSpread.ShadowColor = GetSetting("Hnt.Prj", App.EXEName, "ShadowColor", objSpread.ShadowColor)
    objSpread.ShadowText = GetSetting("Hnt.Prj", App.EXEName, "ShadowText", objSpread.ShadowText)

End Sub


''메뉴를 생성할때 항상 인덱스 "0"에는 배경색변경
''                            "1"에는 글자색변경
''                            "2"에는 기본색으로    를 생성한다.
''그리고 mnuColor_Click 이벤트에서 call만 하면...
'Public Sub SaveSetting_SpreadColor(Index As Integer, objSpread As Object, defaultBackColor As Single, defaultForeColor As Single)
'
'    With mvbFrm.cmDialog
'
'On Error GoTo ErrHandler
'
'        If Index = 기본색으로 Then
'            objSpread.ShadowColor = defaultBackColor
'            objSpread.ShadowText = defaultForeColor
'        Else
'            If Index = 배경색변경 Then       '배경색 변경
'                ' 폼의 배경색을 선택한 색으로 설정합니다.
'                .Color = objSpread.ShadowColor
'            Else                   '글자색 변경
'                .Color = objSpread.ShadowText
'            End If
'
'            .DialogTitle = "색변경"
'            ' Flags 속성을 설정합니다.
'            .Flags = cdlCCRGBInit
'            ' [색] 대화 상자를 표시합니다.
'            .ShowColor
'
'            If Index = 배경색변경 Then       '배경색 변경
'                ' 폼의 배경색을 선택한 색으로 설정합니다.
'                objSpread.ShadowColor = .Color
'            Else                   '글자색 변경
'                objSpread.ShadowText = .Color
'            End If
'        End If
'
'        SaveSetting "Hnt.Prj", App.EXEName, "ShadowColor", objSpread.ShadowColor
'        SaveSetting "Hnt.Prj", App.EXEName, "ShadowText", objSpread.ShadowText
'
'    End With
'
'    Exit Sub
'
'ErrHandler:
'
'End Sub


Public Sub gFillSP(Obj As Object, Str As String, ByVal lCol As Long, ByVal lRow As Long, Optional iopt As Integer)
'vaSpread에 자료를 채운다.
On Error GoTo Handler

    If Obj.MaxCols < lCol Then
        Obj.MaxCols = lCol
    End If


'    Obj.Col = lCol
'    Obj.Row = lRow
    Obj.col = lCol
    Obj.Row = lRow
    
    If Obj.CellType = CellTypePicture Then Obj.CellType = CellTypeEdit
    If iopt = 1 Then
        Obj.Value = Str
    Else
        Obj.Text = Str
    End If
    
    Exit Sub
Handler:
    Resume Next
End Sub

Public Function gfnGetSP(Obj As Object, ByVal col As Long, ByVal Row As Long, Optional iopt As Integer, Optional isNotTrim As Boolean) As String
' vaSpread의 자료를 가져온다.
    On Error GoTo Handler
    With Obj
        .col = col
        .Row = Row
        If iopt = 1 Then
            If isNotTrim = True Then
                gfnGetSP = .Value
            Else
                gfnGetSP = Trim(.Value)
            End If
        Else
            If isNotTrim = True Then
                gfnGetSP = .Text
            Else
                gfnGetSP = Trim(.Text)
            End If
        End If
    End With
    Exit Function
    
Handler:
    gfnGetSP = ""
End Function
'스프레드 Cell 상의 마우스 위치 지정
Public Function Spread_Cnv_MsPntOnCell(mpo_Spd As Object, Optional mpi_RetWhich As Integer = 0, _
                                 Optional mpu_MsPntOnSpd As EnumgsPntOnSpd = 0) As Currency

    Dim mbl_RetValX As Currency     'X 일때 반환값
    Dim mbl_RetValY As Currency     'Y 일때 반환값

   '반환값 X 마우스가 위치하는 Col, Row 값 으로 Cell 상의 마우스 위치 계산
    Select Case mpu_MsPntOnSpd
   'Auto Set
    Case 0
        mgl_Result = mpo_Spd.TypeHAlign
       '정렬모드
        Select Case mgl_Result
        Case 0: mbl_RetValX = 0.01: mbl_RetValY = 0.5    '좌측이면 우중
        Case 1: mbl_RetValX = 0.95: mbl_RetValY = 0.5    '우측이면 좌중
        Case 2: mbl_RetValX = 0.01: mbl_RetValY = 0.5    '가운데면 우중
        End Select
    
   'Manual Set
    Case 1: mbl_RetValX = 0.95: mbl_RetValY = 0.95        '좌상
    Case 2: mbl_RetValX = 0.95: mbl_RetValY = 0.5        '좌중
    Case 3: mbl_RetValX = 0.95: mbl_RetValY = 0.01        '좌하
    Case 4: mbl_RetValX = 0.5: mbl_RetValY = 0.95        '중상
    Case 5: mbl_RetValX = 0.5: mbl_RetValY = 0.5        '중중
    Case 6: mbl_RetValX = 0.5: mbl_RetValY = 0.01        '중하
    Case 7: mbl_RetValX = 0.01: mbl_RetValY = 0.95        '우상
    Case 8: mbl_RetValX = 0.01: mbl_RetValY = 0.5        '우중
    Case 9: mbl_RetValX = 0.01: mbl_RetValY = 0.01        '우하
    End Select
    
   '반환값 지정
    Select Case mpi_RetWhich
    Case 0: Spread_Cnv_MsPntOnCell = mbl_RetValX      'X (Col) 지정
    Case 1:  Spread_Cnv_MsPntOnCell = mbl_RetValY     'Y (Row) 지정
    End Select
    
End Function

Public Sub Get_RetAbsTopLeft(mpo_Spd As Object, mpu_RetAbsLeft As Long, mpu_RetAbsTop As Long)
   
    Dim i As Long
    Dim mbu_RECT As RECT
    Dim mbl_OffSetX As Long
    Dim mbl_OffSetY As Long
    
    Dim mbi_DeepCnt As Integer      '콘테이너의 깊이를 센다.
    Dim mbo_Containers() As Object
        
    On Error GoTo EndOfContainers

    ReDim mbo_Containers(1)
    
    With mpo_Spd
    
    mpu_RetAbsLeft = .Left
    mpu_RetAbsTop = .Top
    
    Set mbo_Containers(1) = mpo_Spd.Container
    
    mpu_RetAbsLeft = mpu_RetAbsLeft + mbo_Containers(1).Left
    mpu_RetAbsTop = mpu_RetAbsTop + mbo_Containers(1).Top
    
   '배열의 다음 위치 지정
    mbi_DeepCnt = 2
    
   '콘테이너의 바닥까지 (에러가 날때까지) 검사
    Do
        ReDim Preserve mbo_Containers(mbi_DeepCnt)
        Set mbo_Containers(mbi_DeepCnt) = mbo_Containers(mbi_DeepCnt - 1).Container
        mpu_RetAbsLeft = mpu_RetAbsLeft + mbo_Containers(mbi_DeepCnt).Left
        mpu_RetAbsTop = mpu_RetAbsTop + mbo_Containers(mbi_DeepCnt).Top
        mbi_DeepCnt = mbi_DeepCnt + 1
    Loop
    
EndOfContainers:
   
    ReDim Preserve mbo_Containers(mbi_DeepCnt - 1)
    
   '폼의 클라이언트 영역 사각크기 구한다.
    mgl_Result = GetClientRect(mbo_Containers(mbi_DeepCnt - 1).hWnd, mbu_RECT)
   
   '타이틀바의 Width , Height 도 포함하여 계산
   
    mpu_RetAbsLeft = mbo_Containers(mbi_DeepCnt - 1).Width + mpu_RetAbsLeft
    mpu_RetAbsTop = mbo_Containers(mbi_DeepCnt - 1).Height + mpu_RetAbsTop
    
   'Pixel 값으로 반환값 계산
    mpu_RetAbsLeft = (mpu_RetAbsLeft / Screen.TwipsPerPixelX) - mbu_RECT.Right
    mpu_RetAbsTop = (mpu_RetAbsTop / Screen.TwipsPerPixelY) - mbu_RECT.Bottom
    
    
    
   '마우스 위치 계산값의 오차에 대한 보상 - BorderStyle 이 있을때에만
    If mbo_Containers(mbi_DeepCnt - 1).BorderStyle <> 0 Then
        mpu_RetAbsLeft = mpu_RetAbsLeft - 5
        mpu_RetAbsTop = mpu_RetAbsTop - 5
    End If
    
   '할당 해제
    For i = 1 To mbi_DeepCnt - 1
        Set mbo_Containers(i) = Nothing
    Next
    
    Erase mbo_Containers
    
    End With
    
End Sub

'2001/10/30 james..... 날짜가 문제다 Picker때문에 이런방법을....
'fps_Value가 NULL이면 fps_CurDte에 값을 채우고 값이 있으면 fpc_dtp에 채운다
Public Sub Setting_DateTimePicker(fpc_dtp As Control, fps_Value As String, Optional fpb_DateType As Boolean = True, Optional fpb_IsDateTimeType As Boolean = False)
    
    On Error Resume Next
    
    If fpb_IsDateTimeType = False Then
        If fpb_DateType Then
            '일자 변경
            If fps_Value = "" Then
                If fpc_dtp.Year = "9999" Then
                    fps_Value = "99999999"
                Else
                    fps_Value = fpc_dtp.Year & Format(fpc_dtp.Month, "00") & Format(fpc_dtp.Day, "00")
                End If
            Else
                If DateValidCheck(fps_Value) Then
                    If fps_Value = "99999999" Then
                        fps_Value = "99991231"
                    End If
                    fpc_dtp.Year = Left(fps_Value, 4)
                    fpc_dtp.Month = Mid(fps_Value, 5, 2)
                    fpc_dtp.Day = Mid(fps_Value, 7, 2)
                End If
            End If
        Else
            '시각변경
            If fps_Value = "" Then
                fps_Value = Format(fpc_dtp.Hour, "00") & Format(fpc_dtp.Minute, "00")
            Else
                If DateValidCheck(fps_Value) Then
                    fpc_dtp.Hour = Left(fps_Value, 2)
                    fpc_dtp.Minute = Right(fps_Value, 2)
                End If
            End If
        End If
    Else
        '날짜 시각 타입일 경우
        If fps_Value = "" Then
            If fpc_dtp.Year = "9999" Then
                fps_Value = "999999999999"
            Else
                fps_Value = fpc_dtp.Year & Format(fpc_dtp.Month, "00") & Format(fpc_dtp.Day, "00")
                fps_Value = fps_Value & Format(fpc_dtp.Hour, "00") & Format(fpc_dtp.Minute, "00")
            End If
        Else
            '200301010 lek edit
            'If DateValidCheck(fps_Value) Then
            If DateTimeValidCheck(fps_Value) Then
                If fps_Value = "999999999999" Then
                    fps_Value = "999912312359"
                End If
                fpc_dtp.Year = Left(fps_Value, 4)
                fpc_dtp.Month = Mid(fps_Value, 5, 2)
                fpc_dtp.Day = Mid(fps_Value, 7, 2)
                fpc_dtp.Hour = Mid(fps_Value, 9, 2)
                fpc_dtp.Minute = Right(fps_Value, 2)
            End If
        End If
    End If
    
End Sub
Public Sub UpdateIcmPreStt(sPrmOcmNum As String, sPrmPreStt As String, Optional sPrmSimDte As String = "")
'Parameter
'[sPrmOcmNum] = 내원번호
'
'[sPrmPreStt]
'1 : 퇴원심사대기중
'2 : 퇴원심사중
'3 : 퇴원심사완료
'4 : 수납대기
'5 : 수납완료

    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sRetVal As String
    Dim tIcmInf As IcmInfRec
        
    sCurKey = Format(sPrmOcmNum, "@@@@@@@@@@") & Chr(5)
    sCurKey = mSetReadEqual("IcmInf", sCurKey, sRetVal)
    If sCurKey <> "" Then
        Call IcmInfLoad(sRetVal, tIcmInf)
        tIcmInf.IcmPreSts = sPrmPreStt
        '당분간 사전심사일자를 막는다.
        'If sPrmSimDte <> "" Then
        '    tIcmInf.IcmSimDtm = sPrmSimDte
        'End If
        If sPrmPreStt = "3" Then
            If Not (tIcmInf.IcmCfmYon = "OT" Or tIcmInf.IcmCfmYon = "OR") Then
                tIcmInf.IcmPreDtm = AddCentury(SystemDate()) & SystemTime()
                tIcmInf.IcmOdrDtm = AddCentury(SystemDate()) & SystemTime()
                tIcmInf.IcmCfmYon = "OM" '장기환자 중간계산할 경우
            End If
        End If
        
        Call IcmInfStore(sCurKey, sRetVal, tIcmInf)
        If Not mUpdate("IcmInf", sCurKey, sRetVal) Then
            MsgBox "IcmInf Wrire Error"
        End If
    End If
    
End Sub

'마우스 포인터를 Spread 의 mpl_ToCol , mpl_ToRow 위치로 옮긴다.
Public Sub Move_SpdMousePointer(mpo_Spd As Object, mpl_ToCol As Long, mpl_ToRow As Long, Optional mpl_Left As Long = -1, Optional mpl_Top As Long = -1)
    
    Dim i As Long
    Dim mbl_ColWidth As Long
    Dim mbl_RowHeight As Long
    Dim mbl_MoveToCol As Long
    Dim mbl_MoveToRow As Long
    
    Dim mbl_GetCol As Long
    Dim mbl_GetRow As Long
    
    Dim mbl_PushColWidth As Long
    Dim mbl_PushRowHeight As Long
    Dim mbl_PushCol As Long
    Dim mbl_PushRow As Long
    
    Dim mbc_RetMsPntOnSpd As Currency
    
    With mpo_Spd
    
   '스프레드의 Col 에 해당하는 Twip 값을 구한다.
    For i = 0 To mpl_ToCol
        mgl_Result = .ColWidthToTwips(.ColWidth(i), mbl_PushColWidth)
        mbl_ColWidth = mbl_ColWidth + mbl_PushColWidth
    Next
    
   'Col 값 + 마우스가 위치할곳 자동계산
    mbc_RetMsPntOnSpd = Spread_Cnv_MsPntOnCell(mpo_Spd, 0, mgu_MousePointOnSpd)
    mbl_ColWidth = mbl_ColWidth - (mbl_PushColWidth * mbc_RetMsPntOnSpd)
        
   '마우스 포인터 옮김여부
    'If IsMissing(mpl_Left) And IsMissing(mpl_Top) Then
    If mpl_Left = -1 And mpl_Top = -1 Then
       '스프레드의 Row 에 해당하는 Twip 값을 구한다.
        For i = 0 To mpl_ToRow
            mgl_Result = .RowHeightToTwips(i, .RowHeight(i), mbl_PushRowHeight)
            mbl_RowHeight = mbl_RowHeight + mbl_PushRowHeight
        Next
    
        'Row 값 + 마우스가 위치할곳 자동계산
         mbc_RetMsPntOnSpd = Spread_Cnv_MsPntOnCell(mpo_Spd, 1, mgu_MousePointOnSpd)
         mbl_RowHeight = mbl_RowHeight - (mbl_PushRowHeight * mbc_RetMsPntOnSpd)
    
       'Return 값은 mgu_RetAbsLeft , mgu_RetAbsTop
        Call Get_RetAbsTopLeft(mpo_Spd, mgu_RetAbsLeft, mgu_RetAbsTop)
    
       'Pixel 값 으로 바꾼다.
        mbl_ColWidth = mbl_ColWidth / Screen.TwipsPerPixelX
        mbl_RowHeight = mbl_RowHeight / Screen.TwipsPerPixelY
        
        mbl_MoveToCol = mgu_RetAbsLeft + mbl_ColWidth
        mbl_MoveToRow = mgu_RetAbsTop + mbl_RowHeight
    
       '마우스 포인터를 옮긴다.
       '주의 - Exit Sub 후 SetCursorPos 에 의해 Mouse Move Event 발생함
        mgl_Result = SetCursorPos(mbl_MoveToCol, mbl_MoveToRow)
    Else
       'Header RowHeight 를 따로 계산
        mgl_Result = .RowHeightToTwips(0, .RowHeight(0), mbl_PushRowHeight)
        mbl_RowHeight = mbl_RowHeight + mbl_PushRowHeight
        
        For i = .TopRow To mpl_ToRow
            mgl_Result = .RowHeightToTwips(i, .RowHeight(i), mbl_PushRowHeight)
            mbl_RowHeight = mbl_RowHeight + mbl_PushRowHeight
        Next
    
        'Row 값 + 마우스가 위치할곳 자동계산
         mbc_RetMsPntOnSpd = Spread_Cnv_MsPntOnCell(mpo_Spd, 1, mgu_MousePointOnSpd)
         mbl_RowHeight = mbl_RowHeight - (mbl_PushRowHeight * mbc_RetMsPntOnSpd)
    
       'Return 값은 mgu_RetAbsLeft , mgu_RetAbsTop
        Call Get_RetAbsTopLeft(mpo_Spd, mgu_RetAbsLeft, mgu_RetAbsTop)
    
       'Pixel 값 으로 바꾼다.
        mbl_ColWidth = mbl_ColWidth / Screen.TwipsPerPixelX
        mbl_RowHeight = mbl_RowHeight / Screen.TwipsPerPixelY
        
        mbl_MoveToCol = mgu_RetAbsLeft + mbl_ColWidth
        mbl_MoveToRow = mgu_RetAbsTop + mbl_RowHeight
    
       '마우스포인터를 옮기지 않고 좌표값만 반환
        mpl_Left = (mbl_ColWidth * Screen.TwipsPerPixelX) + .Left
        mpl_Top = (mbl_RowHeight * Screen.TwipsPerPixelY) + .Top
        Exit Sub
    End If
    

 '----------
    End With
    
End Sub

'---------------------------------------------------------------------------
'   Spread에 여러 라인을 표시할 때 한 라인식 깜박이며 표시되는 것을 막기위해
'       화면 표시전에
'       Call Spread_Redraw(spd_Cod, False)
'       화면 표시후에
'       Call Spread_Redraw(spd_Cod, True)
'
'   fps_Control은 Spread의 이름을 사용한다.
'   (ListBox나 기타 Control에서 사용가능)
'---------------------------------------------------------------------------
Public Sub Spread_Redraw(fps_Control As Control, fpb_Redraw As Boolean)

    If fpb_Redraw Then
        Call SendMessage(fps_Control.hWnd, WM_SETREDRAW, 1, 0)
        
        If TypeOf fps_Control Is vaSpread Then
            Dim mbl_ScrollBars As Long
        
            mbl_ScrollBars = fps_Control.ScrollBars
            fps_Control.ScrollBars = 0
            fps_Control.ScrollBars = mbl_ScrollBars
        End If
        
        DoEvents
        fps_Control.Refresh
    Else
        Call SendMessage(fps_Control.hWnd, WM_SETREDRAW, 0, 0)
    End If

End Sub

Public Sub ActiveCellPosition(grdOacInf As Object, Row As Long, col As Integer)

    With grdOacInf
        .Row = Row
        .col = col
        .Action = ActionActiveCell
    End With
    
End Sub

Public Function AddCenturyLen(sPrmDate As String) As String
    
    'Cache Version의 SystemDate가 Version에 따라 틀리게 들어와서리...이렇게 만듬.
    
    Dim sTmpDate As String
    Dim iTmpYear As Integer
    Dim sTmpTime As String
    Dim iTmpCentury As Integer

    If Len(sPrmDate) = 8 Or Len(sPrmDate) = 12 Then
        AddCenturyLen = sPrmDate
        Exit Function
    ElseIf sPrmDate = "999999" Then
        AddCenturyLen = "99999999"
        Exit Function
    End If
    
    If Len(sPrmDate) = 10 Or Len(sPrmDate) = 12 Then
        sTmpTime = Right(sPrmDate, 4)
    ElseIf Len(sPrmDate) < 10 Then
        sTmpTime = ""
    End If

    If Len(sPrmDate) = 8 Or Len(sPrmDate) = 12 Then
        sTmpDate = sPrmDate
    ElseIf Len(sPrmDate) = 6 Or Len(sPrmDate) = 10 Then
        sTmpDate = Left(sPrmDate, 6)
        iTmpYear = CInteger(Left(sTmpDate, 2))
        
        If iTmpYear > 20 Then
            iTmpCentury = 19
        Else
            iTmpCentury = 20
        End If
        
        sTmpDate = CStr(iTmpCentury) & sTmpDate
    End If

    If IsDate(Format(sTmpDate, "####/##/##")) Then
        AddCenturyLen = sTmpDate & sTmpTime
    Else
        AddCenturyLen = sPrmDate & sTmpTime
    End If

End Function

Public Function AddCentury(sPrmDate As String, Optional bAddCentury As Boolean = True) As String
    
    Dim sTmpDate As String
    Dim iTmpYear As Integer
    Dim iTmpCentury As Integer
    Dim sTmpVal As String
    Dim sBufStr As String
    Dim iCnt As Integer
    Dim sTmpTime As String
    Dim sTmpResNum As String

On Error GoTo ERR_TRAC

    If Len(sPrmDate) = 8 Or Len(sPrmDate) = 12 Then
        AddCentury = sPrmDate
        Exit Function
    ElseIf sPrmDate = "999999" Then
        AddCentury = "99999999"
        Exit Function
    End If
    
    If Len(sPrmDate) >= 10 Then
        sTmpTime = Right(sPrmDate, 4)
        'sTmpVal = sPrmDate
        'Do Until IsDate(Format(Left(sTmpVal, 6), "####/##/##"))
        '    sTmpVal = Mid(sTmpVal, 2, Len(sTmpVal))
        '    If sTmpVal = "" Then
        '        Addcentury = ""
        '        Exit Function
        '    End If
        'Loop
        'sPrmDate = sTmpVal
    Else
        sTmpTime = ""
    End If
    
    sTmpResNum = CInteger(Right(sPrmDate, 1))
    
    sTmpDate = Left(sPrmDate, 6)
    iTmpYear = CInteger(Left(sTmpDate, 2))

    If iTmpYear > 20 Then
        iTmpCentury = 19
    Else
        iTmpCentury = 20
    End If
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '2002/01/16 james
    '1900~1920년과 2000~2020년을 구별못한다... 그래서 이런 방법으로 라도...
    '주민번호라면 앞의 7자리를 넘기고 Optional을 False로....
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If Not bAddCentury Then
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        '이렇게 처리한다고 할 경우에 아래의 값은 맞지 않는 것 같다.. 그래서 수정함
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        Select Case sTmpResNum
        Case "1", "2", "7", "8"    '"0"
            iTmpCentury = "19"
        Case "3", "4"
            iTmpCentury = "20"
        Case "5", "6"
            iTmpCentury = "19"
        Case Else  '0, 9
            iTmpCentury = "18"
        End Select
        'If sTmpResNum = "3" Or sTmpResNum = "4" Then
        '    iTmpCentury = 19
        'End If
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    End If
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    sTmpDate = CStr(iTmpCentury) & sTmpDate

    If IsDate(Format(sTmpDate, "####/##/##")) Then
        AddCentury = sTmpDate & sTmpTime
    Else
        AddCentury = sPrmDate & sTmpTime
    End If

    Exit Function
    
ERR_TRAC:

    sTmpDate = Mid(Date, 3, 2) & sTmpDate

    If IsDate(Format(sTmpDate, "####/##/##")) Then
        AddCentury = sTmpDate & sTmpTime
    Else
        AddCentury = sPrmDate & sTmpTime
    End If

    Exit Function
End Function

Public Sub Add2Collection(oCol As Collection, psdata As String)

    Dim i As Integer
    Dim bSw As Boolean
    
    bSw = True
    For i = 1 To oCol.Count
        If oCol(i) = psdata Then
            bSw = False
            Exit For
        End If
    Next
    
    If bSw Then
        oCol.Add psdata
    End If
    
End Sub

Public Function DataInCollection(oCol As Collection, psdata As String)

    Dim i As Integer
    Dim bSw As Boolean
    
    bSw = False
    For i = 1 To oCol.Count
        If oCol(i) = psdata Then
            bSw = True
            Exit For
        End If
    Next

    DataInCollection = bSw
    
End Function

Public Function AgeCheck(sPrmResNum As String, sPrmToDate As String) As Long

    Dim sTmpFrmDtm As String              '주민등록 년월
    Dim sTmpSex As String                 '성별구분
    Dim sTmpResNum As String
    Dim lTmpYear As Long                  '주민등록 년도
    Dim lTmpMonth As Long                 '주민등록 월
    Dim lTmpDay As Long                   '주민등록 일
    Dim lTmpDate As Long                  '
    Dim sTmpSysDtm As String              'System 년월일
    Dim lTmpSysYear As Long               'System 년도
    Dim lTmpSysMonth As Long              'System 월
    Dim lTmpSysDay As Long                'System 일
    Dim lTmpSysDate As Long               '
    Dim ltmpage As Long                   '나이
    Dim sTmpDur  As String                '기간
    Dim lTmpDur As Long

    If sPrmResNum = "" Then
    AgeCheck = 0
    Exit Function
    End If

    sTmpSex = Mid(sPrmResNum, 7, 1)

    Select Case sTmpSex
    Case "1", "2", "7", "8"    '"0"
        sTmpResNum = "19" & sPrmResNum
    Case "3", "4"
        sTmpResNum = "20" & sPrmResNum
    Case "5", "6"
        sTmpResNum = "19" & Left(sPrmResNum, 2) & "0101" & Right(sPrmResNum, 7)
    Case Else  '0, 9
        sTmpResNum = "18" & sPrmResNum
    End Select

    '주민번호 년,월 Split
    lTmpYear = CLong(Left(sTmpResNum, 4))
    lTmpDate = CLong(Mid(sTmpResNum, 5, 4))
    lTmpMonth = CLong(Mid(sTmpResNum, 5, 2))
    lTmpDay = CLong(Mid(sTmpResNum, 7, 2))
    sTmpFrmDtm = Left(sTmpResNum, 8)

    'System Date Split
    If sPrmToDate = "" Then
        sTmpSysDtm = SystemDate()
        sTmpSysDtm = AddCentury(sTmpSysDtm)
        lTmpSysYear = CLng(Left(sTmpSysDtm, 4))
        lTmpSysDate = CLong(Mid(sTmpSysDtm, 5, 4))
        lTmpSysMonth = CLong(Mid(sTmpSysDtm, 5, 2))
        lTmpSysDay = CLng(Right(sTmpSysDtm, 2))
    Else
        lTmpSysYear = CLng(Left(sPrmToDate, 4))
        lTmpSysDate = CLong(Mid(sPrmToDate, 5, 4))
        lTmpSysMonth = CLong(Mid(sPrmToDate, 5, 2))
        lTmpSysDay = CLng(Right(sPrmToDate, 2))

    End If

    '나이 계산
    ltmpage = lTmpSysYear - lTmpYear
    If lTmpDate >= lTmpSysDate Then
        ltmpage = ltmpage - 1
    End If

    '30일 까지 신생아로 처리(0세)
    If ltmpage <= 1 Then
        '막아야 산다!
        'lTmpDur = AgeCheck(sTmpFrmDtm, "")
        lTmpDur = ToJulian(lTmpSysYear, lTmpSysMonth, lTmpSysDay)
        lTmpDur = lTmpDur - ToJulian(lTmpYear, lTmpMonth, lTmpDay)
        
        If lTmpDur < 365 Then
            ltmpage = 0
        Else
            ltmpage = 1
        End If
    End If
    AgeCheck = ltmpage

End Function

Public Function AgeCheckBaby(sPrmResNum As String, sPrmToDate As String, lPrmAge As Long, lPrmDay As Long) As Integer

    ' 30일 이하이면 AgeCheckBaby = True 이고
    ' lPrmAge는 나이
    ' lPrmDay는 날수


    Dim sTmpFrmDtm As String              '주민등록 년월
    Dim sTmpSex As String                 '성별구분
    Dim sTmpResNum As String
    Dim lTmpYear As Long                  '주민등록 년도
    Dim lTmpMonth As Long                 '주민등록 월
    Dim lTmpDay As Long                   '주민등록 일
    Dim lTmpDate As Long                  '
    Dim sTmpSysDtm As String              'System 년월일
    Dim lTmpSysYear As Long               'System 년도
    Dim lTmpSysMonth As Long              'System 월
    Dim lTmpSysDay As Long                'System 일
    Dim lTmpSysDate As Long               '
    Dim ltmpage As Long                   '나이
    Dim sTmpDur  As String                '기간
    Dim lTmpDur As Long

    If sPrmResNum = "" Then
        lPrmAge = 0
        Exit Function
    End If

    sTmpSex = Mid(sPrmResNum, 7, 1)
    If sTmpSex = "1" Or sTmpSex = "2" Or sTmpSex = "7" Or sTmpSex = "8" Then
        sTmpResNum = "19" & sPrmResNum
    ElseIf sTmpSex = "3" Or sTmpSex = "4" Then
        sTmpResNum = "20" & sPrmResNum
    ElseIf sTmpSex = "5" Or sTmpSex = "6" Then
        sTmpResNum = "19" & Left(sPrmResNum, 2) & "0101" & Right(sPrmResNum, 7)
    Else
        sTmpResNum = "18" & sPrmResNum
    End If

    '주민번호 년,월 Split
    lTmpYear = CLong(Left(sTmpResNum, 4))
    lTmpDate = CLong(Mid(sTmpResNum, 5, 4))
    lTmpMonth = CLong(Mid(sTmpResNum, 5, 2))
    lTmpDay = CLong(Mid(sTmpResNum, 7, 2))
    sTmpFrmDtm = Left(sTmpResNum, 8)

    'System Date Split
    If sPrmToDate = "" Then
    sTmpSysDtm = SystemDate()
    sTmpSysDtm = AddCentury(sTmpSysDtm)
    lTmpSysYear = CLng(Left(sTmpSysDtm, 4))
    lTmpSysDate = CLong(Mid(sTmpSysDtm, 5, 4))
    lTmpSysMonth = CLong(Mid(sTmpSysDtm, 5, 2))
    lTmpSysDay = CLng(Right(sTmpSysDtm, 2))
    Else
    lTmpSysYear = CLng(Left(sPrmToDate, 4))
    lTmpSysDate = CLong(Mid(sPrmToDate, 5, 4))
    lTmpSysMonth = CLong(Mid(sPrmToDate, 5, 2))
    lTmpSysDay = CLng(Right(sPrmToDate, 2))

    End If

    '나이 계산
    ltmpage = lTmpSysYear - lTmpYear
    If lTmpDate >= lTmpSysDate Then
        ltmpage = ltmpage - 1
    End If

    '30일 까지 신생아로 처리(0세)
    If ltmpage <= 1 Then
        '막아야 산다!
        'lTmpDur = AgeCheck(sTmpFrmDtm, "")
        lTmpDur = ToJulian(lTmpSysYear, lTmpSysMonth, lTmpSysDay)
        lTmpDur = lTmpDur - ToJulian(lTmpYear, lTmpMonth, lTmpDay)
        
        If lTmpDur < 365 Then
          ltmpage = 0
        Else
            ltmpage = 1
        End If
    End If
    lPrmAge = ltmpage

    lTmpDur = ToJulian(lTmpSysYear, lTmpSysMonth, lTmpSysDay)
    lPrmDay = lTmpDur - ToJulian(lTmpYear, lTmpMonth, lTmpDay)
    If lPrmDay <= 30 Then AgeCheckBaby = True

End Function

Public Function AgeCheckDay(sPrmResNum As String, sPrmToDate As String) As Long
    
    '<한방> : 일자로 나이 check
    'sPrmToDate :예) 19980101

    Dim sTmpSex As String                 '성별구분
    Dim sTmpResNum As String
    Dim lTmpYear As Long                  '주민등록 년도
    Dim lTmpMonth As Long                 '주민등록 월
    Dim lTmpDay As Long                   '주민등록 일
    Dim lTmpDur As Long

    If sPrmResNum = "" Then
    AgeCheckDay = 0
    Exit Function
    End If
    
    '현재 일자
    lTmpYear = CLong(Left(sPrmToDate, 4))
    lTmpMonth = CLong(Mid(sPrmToDate, 5, 2))
    lTmpDay = CLong(Mid(sPrmToDate, 7, 2))
    
    lTmpDur = ToJulian(lTmpYear, lTmpMonth, lTmpDay)
    

    sTmpSex = Mid(sPrmResNum, 7, 1)
    If sTmpSex = "1" Or sTmpSex = "2" Or sTmpSex = "7" Or sTmpSex = "8" Then
    sTmpResNum = "19" & sPrmResNum
    ElseIf sTmpSex = "3" Or sTmpSex = "4" Then
    sTmpResNum = "20" & sPrmResNum
    Else
    sTmpResNum = "18" & sPrmResNum
    End If

    '주민번호 년,월 Split
    lTmpYear = CLong(Left(sTmpResNum, 4))
    lTmpMonth = CLong(Mid(sTmpResNum, 5, 2))
    lTmpDay = CLong(Mid(sTmpResNum, 7, 2))

    lTmpDur = lTmpDur - ToJulian(lTmpYear, lTmpMonth, lTmpDay)
    AgeCheckDay = lTmpDur

End Function

Public Function AgeCheck_Emb(sPrmResNum As String, sPrmToDate As String) As String

    Dim sTmpFrmDtm As String              '주민등록 년월
    Dim sTmpSex As String                 '성별구분
    Dim sTmpResNum As String
    Dim lTmpYear As Long                  '주민등록 년도
    Dim lTmpMonth As Long                 '주민등록 월
    Dim lTmpDay As Long                   '주민등록 일
    Dim lTmpDate As Long                  '
    Dim sTmpSysDtm As String              'System 년월일
    Dim lTmpSysYear As Long               'System 년도
    Dim lTmpSysMonth As Long              'System 월
    Dim lTmpSysDay As Long                'System 일
    Dim lTmpSysDate As Long               '
    Dim ltmpage As Long                   '나이
    Dim sTmpDur  As String                '기간
    Dim lTmpDur As Long

    If sPrmResNum = "" Then
        AgeCheck_Emb = "0"
        Exit Function
    End If

    sTmpSex = Mid(Pict2Data(sPrmResNum, "9999999999999"), 7, 1)

    Select Case sTmpSex
    Case "1", "2", "7", "8"    '"0"
        sTmpResNum = "19" & sPrmResNum
    Case "3", "4"
        sTmpResNum = "20" & sPrmResNum
    Case "5", "6"
        sTmpResNum = "19" & Left(sPrmResNum, 2) & "0101" & Right(sPrmResNum, 7)
    Case Else  '0, 9
        sTmpResNum = "18" & sPrmResNum
    End Select

    '주민번호 년,월 Split
    lTmpYear = CLong(Left(sTmpResNum, 4))
    lTmpDate = CLong(Mid(sTmpResNum, 5, 4))
    lTmpMonth = CLong(Mid(sTmpResNum, 5, 2))
    lTmpDay = CLong(Mid(sTmpResNum, 7, 2))
    sTmpFrmDtm = Left(sTmpResNum, 8)

    'System Date Split
    If sPrmToDate = "" Then
        sTmpSysDtm = SystemDate()
        sTmpSysDtm = AddCentury(sTmpSysDtm)
        lTmpSysYear = CLng(Left(sTmpSysDtm, 4))
        lTmpSysDate = CLong(Mid(sTmpSysDtm, 5, 4))
        lTmpSysMonth = CLong(Mid(sTmpSysDtm, 5, 2))
        lTmpSysDay = CLng(Right(sTmpSysDtm, 2))
    Else
        lTmpSysYear = CLng(Left(sPrmToDate, 4))
        lTmpSysDate = CLong(Mid(sPrmToDate, 5, 4))
        lTmpSysMonth = CLong(Mid(sPrmToDate, 5, 2))
        lTmpSysDay = CLng(Right(sPrmToDate, 2))

    End If

    '나이 계산
    ltmpage = lTmpSysYear - lTmpYear
    If lTmpDate >= lTmpSysDate Then
        ltmpage = ltmpage - 1
    End If

    '30일 까지 신생아로 처리(0세)
    If ltmpage <= 1 Then
        '막아야 산다!
        'lTmpDur = AgeCheck(sTmpFrmDtm, "")
        lTmpDur = ToJulian(lTmpSysYear, lTmpSysMonth, lTmpSysDay)
        lTmpDur = lTmpDur - ToJulian(lTmpYear, lTmpMonth, lTmpDay)
        
        If lTmpDur < 365 Then
            ltmpage = lTmpDur \ 30
            AgeCheck_Emb = CStr(ltmpage) & " 개월"
        Else
            ltmpage = 1
            AgeCheck_Emb = CStr(ltmpage) & " 세"
        End If
    Else
        AgeCheck_Emb = CStr(ltmpage) & " 세"
    End If
    
    

End Function

Public Function ArgnDate(sDate As String) As String

    If Len(piece(sDate, "/", 1)) = 2 Then
    sDate = Pict2Data(sDate, "99/99/99")
    sDate = AddCentury(sDate)
    
    ElseIf Len(piece(sDate, "/", 1)) = 4 Then
    sDate = Pict2Data(sDate, "9999/99/99")

    ElseIf Len(sDate) = 8 Then

    End If

End Function

Public Sub AssMstRead(sAssCod As String, AssData As AssMstRec)
    
    Dim sAssMstCurKey As String, sAssMstRetVal As String

    sAssMstCurKey = sAssCod & Chr(5)
    sAssMstCurKey = mSetReadEqual("AssMst", sAssMstCurKey, sAssMstRetVal)
    AssMstLoad sAssMstRetVal, AssData

End Sub

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
' 주민등록번호를 가지고 생년월일을 구하는 것이다..
' sPrmResNum -> 주민번호
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Public Function BirthDay(sPrmResNum As String) As String
    
    Dim sYear As String
    Dim sMon As String
    Dim sDay As String
    Dim sTmp As String

    sTmp = AddCentury(Left(sPrmResNum, 6))

    sYear = Left(sTmp, 4)
    sMon = Mid(sTmp, 5, 2)
    sDay = Mid(sTmp, 7, 2)

    sTmp = sYear & "년 "
    sTmp = sTmp & sMon & "월 "
    sTmp = sTmp & sDay & "일 "

    BirthDay = sTmp

End Function

'**************************************************
'   2차원 array sort
'--------------------------------------------------
'   1. iPrmRowCount : Sort할 Row의 갯수
'   2. iPrmColCount : Sort할 Column의 갯수
'   3. iPrmColumn   : Sort의 대상이 되는 Column
'   4. sPrmRowData  : 2차원 array
'**************************************************
Public Sub BubbleSort(iPrmRowCount As Integer, iPrmColCount As Integer, iPrmColumn As Integer, sPrmRowData() As String)
    Dim Row As Integer
    Dim col As Integer
    Dim iCnt As Integer
    Dim sBufValue() As String

    ' 임시 저장 array
    ReDim sBufValue(1 To iPrmColCount)

    If iPrmRowCount = 1 Then Exit Sub

    For iCnt = iPrmRowCount - 1 To 1 Step -1
        For Row = 1 To iCnt
            ' 가장 큰 값을 제일 뒤로 이동
            If sPrmRowData(Row, iPrmColumn) > sPrmRowData(Row + 1, iPrmColumn) Then
            For col = 1 To iPrmColCount
                sBufValue(col) = sPrmRowData(Row + 1, col)
                sPrmRowData(Row + 1, col) = sPrmRowData(Row, col)
                sPrmRowData(Row, col) = sBufValue(col)
            Next
            End If
        Next
    Next

    Erase sBufValue
    
End Sub

Public Function CalcAmount(sPrmInsCod As String, sPrmDgnNfs As String, sPrmDgnDnh As String) As String
'    Dim sTmpAcpDtm As String
'    Dim sTmpCurDtm As String
'    Dim sTmpAdmDur As String
'    Dim sBufvalue  As String
'    Dim sBufDepCod As String
'    Dim sBufGrpCod As String
'    Dim lBufAcpDtm As Long
'    Dim lTmpAdmDur  As Long
'    Dim lTmpSavDate As Long
'    Dim InsData As InsMstRec
'    Dim OcmChtDepAcpData As OcmInfChtDepAcpRec
'    Dim DepGrpData As DepMstGrpRec
'    Dim OcmData As OcmInfRec
    
    '보험유형 Date Base를   Check한다.
    'Read 보험유형 DateBase (InsMst)
'    InsData.InsCod = sPrmInsCod                        '보험유형
'    InsMstKeyStore sBufvalue, InsData
'    sBufvalue = mReadEqual("InsMst", sBufvalue)
'    InsMstLoad sBufvalue, InsData


    'Read 진찰료 테이블 DateBase(FtpMst)
'    FtpData.FtpDgnCod = sPrmDgnNfsFtpData.FtpDgnDnh = sPrmDgnDhn
'    FtpData.FtpAgeCod = sPrmAgeCod
'
'    FtpMstKeyStore sBufvalue, FtpData
'    sBufvalue = mReadEqual("FtpMst", sBufvalue)
'    FtpMstLoad sBufvalue, FtpData
'    sBufFeeCod = FtpData.FtpFeeCod

    'Read 수가마스터 DateBase(FeeMst)
'    FeeData.FeeCod = sBufFeeCod
'    FeeMstKeyStore sBufvalue, FeeData
 '   sBufvalue = mReadEqual("FeeMst", sBufvalue)
 '   FeeMstLoad sBufvalue, FeeData
 '
'    lInsCodType = InsDate.InsCidType
    '진찰료를 자동 산정한다.
'    If InsData.lInsCodType = "11" Or lInsCodType = "12" Then
'        '총진료비
'        lTmpTotAmt = FeeInsAmt
'    End If
'    '환자 청구액
'    lTmpAskAmt = lTmpTotAmt * (InsOpoRat / 100)
'    '기 수납액
'    lTmpOldAmt = FeeInsAmt
'    '수납액
'    lTmpNewAmt = FeeInsAmt
'
'    While sBufvalue <> ""
'    sBufvalue = mReadEqual("OcmInfChrDepAcm", sBufvalue)
'    OcmInfChtDepAcpLoad sBufvalue, OcmChtDepAcpData
'
'    If sBufvalue <> "" Then
'        OcmData.OcmChtNum = sPrmChrNum                    '챠트번호
'                OcmData.OcmDepCod = sBufDepCod                    '진료과 코드
'                OcmData.OcmAcpDtm = sPrmAcpDtm                    '접수일자
'                OcmInfKeyStore sBufvalue, OcmData
'                sBufvalue = mReadLittle("OcmInf", sBufvalue, 3)
'
'                If sBufvalue <> "" Then
'                    lBufAcpDtm = CLong(OcmData.OcmAcpDtm)
'                    If lTmpSavDate < lBufAcpDtm Then
'                        lTmpSavDate = lBufAcpDtm
'                    End If
'                End If
'            End If
'        Wend
'    Wend
'
'
End Function

Public Sub CalDateTime(sPrmPrvDtm As String, sPrmCurDtm As String, iPrmMinute As Integer)

    Dim iErrCod As Integer, sTmpDte As String
    Dim sPrvTim As String, sCurTim As String, sPrvDte As String, sCurDte As String

    If sPrmPrvDtm = "" Then
    sPrvTim = ""
    Else
    sPrvTim = Right(sPrmPrvDtm, 4)
    sPrvDte = Left(sPrmPrvDtm, 8)
    End If

    If sPrmCurDtm = "" Then
    sCurTim = ""
    Else
    sCurTim = Right(sPrmCurDtm, 4)
    sCurDte = Left(sPrmCurDtm, 8)
    End If

    sTmpDte = ""
    iErrCod = CalTime(sPrvTim, sCurTim, iPrmMinute)
    If sPrvTim > sCurTim Then
    If sPrmPrvDtm = "" Then
        Call Caljulian(sTmpDte, sCurDte, "-1")
        sPrmPrvDtm = sTmpDte & sPrvTim
    ElseIf sPrmCurDtm = "" Then
        Call Caljulian(sPrvDte, sTmpDte, "1")
        sPrmCurDtm = sTmpDte & sCurTim
    End If
    Else
    If sPrmPrvDtm = "" Then
        sPrmPrvDtm = sCurDte & sPrvTim
    ElseIf sPrmCurDtm = "" Then
        sPrmCurDtm = sPrvDte & sCurTim
    End If
    End If
    
End Sub

Public Function CalEndDate(sDate As String) As String
    
    Dim i As Integer
    Dim sExpDte As String

    sExpDte = sDate
    For i = 31 To 28 Step -1
    sExpDte = Format(Left(sDate, 6) & CStr(i), "####/##/##")
    If IsDate(sExpDte) Then
        sExpDte = Left(sExpDte, 4) & Mid(sExpDte, 6, 2) & Right(sExpDte, 2)
        Exit For
    End If
    Next

    CalEndDate = sExpDte

End Function

'==========================================================================
'Function Name     : Julian Date  산정
'DesCription       :
'--------------------------------------------------------------------------
'Input ParaMeter   : sPrmFrmDate  시작일시
'                  : sPrmToDate   종료일시
'                  : sPrmDur      기간
'OutPut ParaMeter  :
'참조  Date Base   : None
'Calling  Program  :
'Called   Program  :
'Programe By       :
'Create Date       : 95/09/20
'=========================================================================
Public Sub Caljulian(sPrmFrmDate As String, sPrmToDate As String, sPrmDur As String)

On Error GoTo CalJulianErrorRoutine

    Dim sTmpSwitch As String

    Dim iTmpTotal As Integer
    Dim lTmpYear As Integer
    Dim lTmpDur As Long

    Dim lTmpFrmYear As Long
    Dim lTmpFrmMonth As Long
    Dim lTmpFrmTotal As Long
    Dim lTmpFrmDay As Long

    Dim lTmpToYear As Long
    Dim lTmpToMonth As Long
    Dim lTmpToDay As Long
    Dim lTmpToTotal As Long

    lTmpDur = CLong(sPrmDur)
    '종료일 산정
    If sPrmFrmDate > "0" And lTmpDur > 0 Then
        lTmpFrmYear = CLng(Left(sPrmFrmDate, 4))
        lTmpFrmMonth = CLng(Mid(sPrmFrmDate, 5, 2))
        lTmpFrmDay = CLng(Mid(sPrmFrmDate, 7, 2))
        sTmpSwitch = "3"

    '시작일 산정
    ElseIf sPrmToDate > "0" And lTmpDur < 0 Then
        lTmpToYear = CLng(Left(sPrmToDate, 4))
        lTmpToMonth = CLng(Mid(sPrmToDate, 5, 2))
        lTmpToDay = CLng(Mid(sPrmToDate, 7, 2))
            
        sTmpSwitch = "2"

    '기간 산정
    ElseIf sPrmFrmDate > "0" And sPrmToDate > "0" Then
        lTmpFrmYear = CLng(Left(sPrmFrmDate, 4))
        lTmpFrmMonth = CLng(Mid(sPrmFrmDate, 5, 2))
        lTmpFrmDay = CLng(Mid(sPrmFrmDate, 7, 2))
        lTmpToYear = CLng(Left(sPrmToDate, 4))
        lTmpToMonth = CLng(Mid(sPrmToDate, 5, 2))
        lTmpToDay = CLng(Mid(sPrmToDate, 7, 2))
        
        sTmpSwitch = "1"
    Else
        Exit Sub
    End If

    Select Case sTmpSwitch
    Case "1"
        lTmpFrmTotal = ToJulian(lTmpFrmYear, lTmpFrmMonth, lTmpFrmDay)
        lTmpToTotal = ToJulian(lTmpToYear, lTmpToMonth, lTmpToDay)
        sPrmDur = CStr(lTmpToTotal - lTmpFrmTotal + 1)

    Case "2"
        sPrmFrmDate = CStr(CalToDate(sPrmToDate, sPrmDur))

    Case "3"
        sPrmToDate = CStr(CalToDate(sPrmFrmDate, sPrmDur))

    End Select

    Exit Sub
CalJulianErrorRoutine:
    Resume Next

End Sub

Public Function CalTime(sPrmPrvTim As String, sPrmCurTim As String, iPrmMinute As Integer) As Integer
    ' sPrmPrvTim : 시작 시간
    ' sPrmCurTim : 끝 시간
    ' lPrmMinute : 분 간격
    ' Return Value 0 = 일자 무변환
    '              0 > 일자 수정 (- 일)
    '              0 < 일자 수정 (+ 일)

    Dim iCalFlg As Integer
    ' iCalFlg 0 : 시작 시간
    '         1 : 끝 시간
    '         2 : 분 간격
    Dim iOrgHou As Integer, iOrgMin As Integer, iCalHou As Integer, iCalMin As Integer
    Dim iRetVal As Integer, iTmpMin As Integer

    If iPrmMinute <= 0 Then
    MsgBox "분 간격을 양수로 입력해주세요.", vbModal, "'CalTime()' says..."
    CalTime = 0
    Exit Function
    End If
    
    iCalFlg = 0
    If sPrmPrvTim <> "" And TimeValidCheck(sPrmPrvTim) Then
    If sPrmCurTim <> "" And TimeValidCheck(sPrmCurTim) Then
        iCalFlg = 2
    Else
        iCalFlg = 1
    End If
    Else
    If Not TimeValidCheck(sPrmCurTim) Then
        CalTime = 0
        Exit Function
    End If
    End If

    iRetVal = 0

    Select Case iCalFlg
    Case 0  ' 시작 시간
    iOrgHou = CInteger(Left(sPrmCurTim, 2))
    iOrgMin = CInteger(Right(sPrmCurTim, 2))

    iCalMin = iOrgMin - iPrmMinute
    If iCalMin < 0 Then
        iCalMin = 60 + iCalMin
        If iOrgHou - 1 < 0 Then
        iOrgHou = 24
        End If
    
        iCalHou = iOrgHou - 1
    Else
        iCalHou = iOrgHou
    End If

    sPrmPrvTim = Right("00" & CStr(iCalHou), 2) & Right("00" & CStr(iCalMin), 2)

    Case 1  ' 끝 시간
    iOrgHou = CInteger(Left(sPrmPrvTim, 2))
    iOrgMin = CInteger(Right(sPrmPrvTim, 2))

    iCalMin = iOrgMin + iPrmMinute
    If iCalMin > 59 Then
        iCalMin = iCalMin - 60
        If iOrgHou + 1 >= 24 Then
        iOrgHou = -1
        End If

        iCalHou = iOrgHou + 1
    Else
        iCalHou = iOrgHou
    End If

    sPrmCurTim = Right("00" & CStr(iCalHou), 2) & Right("00" & CStr(iCalMin), 2)

    Case 2  ' 분 간격
    iOrgHou = CInteger(Left(sPrmPrvTim, 2))
    iOrgMin = CInteger(Right(sPrmPrvTim, 2))
    iCalHou = CInteger(Left(sPrmCurTim, 2))
    iCalMin = CInteger(Right(sPrmCurTim, 2))

    iPrmMinute = (iOrgHou - iCalHou) * 60
    iTmpMin = iOrgMin - iCalMin
    ' 분끼리의 계산이 양수일 경우는 더하고, 음수일 경우는 뺀다.
    If iTmpMin > 0 Then
        iPrmMinute = iPrmMinute + iTmpMin
    Else
        iPrmMinute = iPrmMinute - iTmpMin
    End If

    End Select

    CalTime = iRetVal
    
End Function

Public Function CalToDate(sPrmDate As String, sPrmDurDays As String) As String
    
    Dim i As Integer
    Dim lTmpMonthSum() As Long
    Dim lTmpMonthDay() As Long
    Dim lTmpYear As Long
    Dim lTmpMonth As Long
    Dim lTmpDay As Long
    Dim lTmpDays As Long
    Dim lTmpToYear As Long
    Dim lTmpToMonth As Long
    Dim lTmpToDay As Long
    Dim lTmpWrkDays As Long

    Dim lTmpTotalDays As Long
    Dim iTmpSwitch  As Integer

    'Parameter Passing Value
    Dim lPrmDay As Long
    Dim lPrmMonth As Long
    Dim lPrmYear As Long

    '월별 일수
    ReDim lTmpMonthDay(1 To 12)

    lTmpMonthDay(1) = 31
    lTmpMonthDay(2) = 28
    lTmpMonthDay(3) = 31
    lTmpMonthDay(4) = 30
    lTmpMonthDay(5) = 31
    lTmpMonthDay(6) = 30
    lTmpMonthDay(7) = 31
    lTmpMonthDay(8) = 31
    lTmpMonthDay(9) = 30
    lTmpMonthDay(10) = 31
    lTmpMonthDay(11) = 30
    lTmpMonthDay(12) = 31

    '적용일자 Split
    lTmpYear = CLong(Left(sPrmDate, 4))
    lTmpMonth = CLong(Mid(sPrmDate, 5, 2))
    lTmpDay = CLong(Mid(sPrmDate, 7, 2))

    '총일수 계산
    lTmpTotalDays = ToJulian(lTmpYear, lTmpMonth, lTmpDay) + CLong(sPrmDurDays)

    '기준일자 1990년 12월 31일
    lPrmYear = 1990
    lPrmMonth = 12
    lPrmDay = 31
    lTmpToYear = lPrmYear

    '기준일자의 총일수 계산
    lTmpDays = ToJulian(lPrmYear, lPrmMonth, lPrmDay)

    '적용년도 산출
    While lTmpDays < lTmpTotalDays
    lTmpToYear = lTmpToYear + 1
    If IsLeapyear(CInteger(lTmpToYear)) Then
        lTmpDays = lTmpDays + 366
    Else
        lTmpDays = lTmpDays + 365
    End If
    Wend

    '적용월 산출
    If IsLeapyear(CInteger(lTmpToYear)) Then
    lTmpMonthDay(2) = lTmpMonthDay(2) + 1
    End If
    lPrmYear = lTmpToYear
    lTmpDays = ToJulian(lPrmYear - 1, lPrmMonth, lPrmDay)
    lTmpToMonth = 0
    lTmpWrkDays = lTmpDays
    For i = 1 To 12
    If lTmpWrkDays < lTmpTotalDays Then
        lTmpWrkDays = lTmpWrkDays + lTmpMonthDay(i)
        lTmpToMonth = lTmpToMonth + 1
    Else
        Exit For
    End If
    Next i

    If lTmpToMonth < 13 Then
    '    lTmpToMonth = lTmpToMonth - 1
    lPrmMonth = lTmpToMonth
    Else
    lPrmMonth = 12
    lTmpToMonth = 12
    End If
    
    '적용일 산출
    If lTmpToMonth > 1 Then
    lTmpToDay = lTmpTotalDays - ToJulian(lTmpToYear, lTmpToMonth - 1, lTmpMonthDay(lTmpToMonth - 1))
    Else
    lTmpToDay = lTmpTotalDays - lTmpDays
    If lTmpToDay = 0 Then
        lTmpToMonth = 12
        lTmpToDay = 31
        lTmpToYear = lTmpToYear - 1
    End If
    End If
    CalToDate = Format(lTmpToYear, "000#") & Format(lTmpToMonth, "0#") & Format(lTmpToDay, "0#")

End Function

Public Function CCut(sPrmValue As Variant) As Long

    Dim sTmp As String
On Error GoTo CC_Handler

    If sPrmValue = "" Then
    CCut = 0
    Else
    'CCut = CLng(CLng(sPrmValue \ 10) * 10)
     sTmp = piece(sPrmValue, ".", 1)
     CCut = CLng(Mid(sTmp, 1, Len(sTmp) - 1) & "0")
    End If

    Exit Function

CC_Handler:
    CCut = 0
    Resume Next

End Function

Public Function CDouble(sPrmValue As Variant) As Double
    
On Error GoTo DHandler

    If sPrmValue = "" Then
        CDouble = 0
    Else
        CDouble = CDbl(sPrmValue)
    End If

    Exit Function

DHandler:
    CDouble = 0
    Resume Next

End Function

Public Function CenterAlignData2Pict(ByVal sPrmBufStr As String, ByVal sPrmPicStr As String) As String

    Dim iPicLen As Integer, iBufLen As Integer, iTmpLen As Integer
    Dim sRetStr As String, sTmpStr As String
    
    If LenK(sPrmBufStr) > LenK(sPrmPicStr) Then
    CenterAlignData2Pict = sPrmBufStr
    Exit Function
    End If

    sRetStr = Data2Pict(sPrmBufStr, sPrmPicStr)

    iBufLen = LenK(sRetStr)
    iPicLen = LenK(sPrmPicStr)
    iTmpLen = Abs(iPicLen - iBufLen)

    CenterAlignData2Pict = Space(iTmpLen / 2) & sRetStr & Space(iTmpLen / 2)

End Function

Public Function CheckID() As String
    
    Dim CodDataCheck As String
    Dim sBufValue As String
    Dim iBufValue As Long
    
    CodDataCheck = AddCentury(SystemDate())
    iBufValue = ((CLong(CodDataCheck) / 408) + CLong(Mid(CodDataCheck, 5, 4))) + 0.5
    CodDataCheck = piece(CStr(iBufValue), ".", 1)
    CheckID = "H" & CodDataCheck

End Function

Public Function CheckOacInfConYon(sPrmChtNum As String, sPrmRvnTyp As String, sPrmDtm As String, tPrmOacData() As OacInfRec)

    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sRetVal As String

    Dim IcmData As IcmInfRec
    Dim OcmData As OcmInfRec
    Dim TrsData As OcmInfRec

    Dim sOrpCurKey As String
    Dim sOrpCmpKey As String
    Dim sOrpRetVal As String
    Dim OrpData As OrpInfRec

    Dim sIrcCurKey As String
    Dim sIrcCmpKey As String
    Dim sIrcRetVal As String
    Dim IrcData As IrcInfRec

    Dim sIrpCurKey As String
    Dim sIrpCmpKey As String
    Dim sIrpRetVal As String
    Dim IrpData As IrpInfRec

    Dim sOacCurKey As String
    Dim sOacCmpKey As String
    Dim sOacRetVal As String
    Dim OacData As OacInfRec

    Dim iCnt As Integer
    Dim iIcmFlg As Integer, iOcmFlg As Integer, iTrsFlg As Integer
    Dim sRvnTyp As String
    Dim sGlobalName As String


    sCmpKey = Format(sPrmChtNum, "@@@@@@@@") & Chr(5)
    sCurKey = sCmpKey & sPrmDtm & Chr(5)
    sCurKey = mSetPrev("IcmInfChtDtm", sCurKey)
    Do
        sCurKey = mReadPrev("IcmInfChtDtm", sCurKey, sCmpKey, sRetVal)
        If sCurKey <> "" Then
            Call IcmInfLoad(sRetVal, IcmData)
            If Trim(IcmData.IcmAcpStt) = "ID" Then
                Call IcmInfLoad(sRetVal, IcmData)
        
                sCmpKey = Format(IcmData.IcmOcmNum, "@@@@@@@@@@") & Chr(5)
                sCurKey = sCmpKey
                sCurKey = mSetNext("IrpInf", sCurKey)
                Do
                    sCurKey = mReadNext("IrpInf", sCurKey, sCmpKey, sRetVal)
                    If sCurKey <> "" Then
                        Call IrpInfLoad(sRetVal, IrpData)
                            If Trim(IrpData.IrpRcpFlg) = "DISCAL" Then
                                iIcmFlg = True
                                Exit Do
                            End If
                        Else
                            iIcmFlg = False
                        Exit Do
                    End If
                Loop
            End If
        Else
            iIcmFlg = False
            Exit Do
        End If
    Loop

    sCmpKey = Format(sPrmChtNum, "@@@@@@@@") & Chr(5)
    sCurKey = sCmpKey & sPrmDtm & Chr(5)
    sCurKey = mSetPrev("OcmInfChtDtm", sCurKey)
    Do
        sCurKey = mReadPrev("OcmInfChtDtm", sCurKey, sCmpKey, sRetVal)
        If sCurKey <> "" Then
            Call OcmInfLoad(sRetVal, OcmData)
            If OcmData.OcmComStt <> "OC" And OcmData.OcmFreRsn = "" Then
                iOcmFlg = True
                Exit Do
            End If
        Else
            iOcmFlg = False
            Exit Do
        End If
    Loop

    sCmpKey = Format(sPrmChtNum, "@@@@@@@@") & Chr(5)
    sCurKey = sCmpKey & sPrmDtm & Chr(5)
    sCurKey = mSetPrev("OcmTmpChtDtm", sCurKey)
    Do
        sCurKey = mReadPrev("OcmTmpChtDtm", sCurKey, sCmpKey, sRetVal)
        If sCurKey <> "" Then
            Call OcmInfLoad(sRetVal, TrsData)
            If TrsData.OcmComStt <> "OC" And TrsData.OcmFreRsn = "" Then
                iTrsFlg = True
                Exit Do
            End If
        Else
            iTrsFlg = False
            Exit Do
        End If
    Loop

    sGlobalName = "OrpInf"
    '외래와 전환된 자료를 먼저 비교한다!
    If iOcmFlg And iTrsFlg Then
        If OcmData.OcmAcpDtm < TrsData.OcmAcpDtm Then
            OcmData = TrsData
            sGlobalName = "OrpTmp"
        End If
    End If

    If iOcmFlg And iIcmFlg Then
        If OcmData.OcmAcpDtm > IcmData.IcmLevDtm Then
            Select Case sPrmRvnTyp
            Case "O", "M"
                sRvnTyp = sPrmRvnTyp
            Case "I"
                sRvnTyp = "O"
            End Select
        Else
            sRvnTyp = "I"
        End If
    ElseIf iOcmFlg Then
        Select Case sPrmRvnTyp
        Case "O", "M"
            sRvnTyp = sPrmRvnTyp
        Case "I"
            sRvnTyp = "O"
        End Select
    ElseIf iIcmFlg Then
        sRvnTyp = "I"
    Else
        CheckOacInfConYon = False
        Exit Function
    End If
    
    iCnt = 0
    If sRvnTyp = "I" Then
        sIrcCmpKey = IcmData.IcmOcmNum & Chr(5)
        sIrcCurKey = sIrcCmpKey
        sIrcCurKey = mSetNext("IrcInf", sIrcCurKey)
        Do
            sIrcCurKey = mReadNext("IrcInf", sIrcCurKey, sIrcCmpKey, sIrcRetVal)
            If sIrcCurKey = "" Then Exit Do
            Call IrcInfLoad(sIrcRetVal, IrcData)
            If (IrcData.IrcRcpTyp = "D" Or IrcData.IrcRcpTyp = "C") And Trim(IrcData.IrcDupSeq) = "0" Then
                sOacCmpKey = IrcData.IrcIrpNum & Chr(5)
                sOacCurKey = sOacCmpKey
                sOacCurKey = mSetNext("OacInf", sOacCurKey)
                Do
                    sOacCurKey = mReadNext("OacInf", sOacCurKey, sOacCmpKey, sOacRetVal)
                    If sOacCurKey = "" Then Exit Do
                    Call OacInfLoad(sOacRetVal, OacData)
                    If MasterHelpDetail("AccMst", OacData.OacAccCod & Chr(5), "", 14) = "Y" Then
                        iCnt = iCnt + 1
                        tPrmOacData(iCnt) = OacData
                    End If
                Loop
                Exit Do
            End If
        Loop
    Else
        sOrpCurKey = OcmData.OcmNum & Chr(5) & sRvnTyp & Chr(5)
        sOrpCurKey = mSetReadEqual(sGlobalName, sOrpCurKey, sOrpRetVal)
        If sOrpCurKey <> "" Then
            Call OrpInfLoad(sOrpRetVal, OrpData)
            If CLong(OrpData.OrpTotAmt) > 0 Then
                sOacCmpKey = OrpData.OrpRcpNum & Chr(5)
                sOacCurKey = sOacCmpKey
                sOacCurKey = mSetNext("OacInf", sOacCurKey)
                Do
                    sOacCurKey = mReadNext("OacInf", sOacCurKey, sOacCmpKey, sOacRetVal)
                    If sOacCurKey = "" Then Exit Do
                    Call OacInfLoad(sOacRetVal, OacData)
                    If MasterHelpDetail("AccMst", OacData.OacAccCod & Chr(5), "", 14) = "Y" Then
                        iCnt = iCnt + 1
                        tPrmOacData(iCnt) = OacData
                    End If
                Loop
            End If
        End If
    End If

    If iCnt > 0 Then
        CheckOacInfConYon = True
    Else
        CheckOacInfConYon = False
    End If
    
End Function

Public Function CheckSlipIsCancel(iPrmSw As Integer, sPrmOcmNum As String, sPrmFlag As String, sPrmDte As String) As String
    '-------------------------------------------------------------------------'
    ' iPrmSw ==>  True : OspInf  ,   False : IspInf
    ' sPrmOcmNum ==> OcmNum
    ' sPrmFlag ==> XRAY LAB INJ 등 전달부서
    ' sPrmDte  ==> 검색할 날짜
    ' Return Value : 해당 환자의 처방전이 모두 취소되었으면 Y,아니면 N
    '-------------------------------------------------------------------------'

    Dim i As Integer
    Dim sDBName As String
    
    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sRetVal As String

    Dim OspData As OspInfRec
    Dim IspData As IspInfRec

    CheckSlipIsCancel = "Y"  '초기화

    Select Case iPrmSw
    Case True   '외래
        sDBName = "OspInfOcmSlpChkOdr"
        
        sCmpKey = Format(Trim(sPrmOcmNum), "@@@@@@@@@@") & Chr(5) & sPrmFlag & Chr(5)
        sCurKey = sCmpKey
        sCurKey = mSetNext(sDBName, sCurKey)
        Do
            sCurKey = mReadNext(sDBName, sCurKey, sCmpKey, sRetVal)
            If sCurKey = "" Then Exit Do
    
            Call OspInfLoad(sRetVal, OspData)
    
            If Left(OspData.OspOdrDtm, 8) = sPrmDte Or Left(OspData.OspPreDtm, 8) = sPrmDte Then
            If OspData.OspOdrStt <> "OC" Then
                CheckSlipIsCancel = "N"
                Exit Do
            End If
            End If
        Loop

    Case False  '입원
        sDBName = "IspInfOcmSlpChkOdr"
        
        sCmpKey = Format(Trim(sPrmOcmNum), "@@@@@@@@@@") & Chr(5) & sPrmFlag & Chr(5)
        sCurKey = sCmpKey
        sCurKey = mSetNext(sDBName, sCurKey)
        Do
            sCurKey = mReadNext(sDBName, sCurKey, sCmpKey, sRetVal)
            If sCurKey = "" Then Exit Do
    
            Call IspInfLoad(sRetVal, IspData)
            If Left(IspData.IspOdrDtm, 8) = sPrmDte Or Left(IspData.IspPreDtm, 8) = sPrmDte Then
            If IspData.IspOdrStt <> "OC" Then
                CheckSlipIsCancel = "N"
                Exit Do
            End If
            End If
        Loop

    End Select

End Function

Public Function CInteger(sPrmValue As Variant) As Integer
    
    On Error GoTo iHandler
    
    If sPrmValue = "" Then
        CInteger = 0
    Else
        CInteger = CInt(sPrmValue)
    End If
    
    Exit Function
    
iHandler:
    CInteger = 0
    Resume Next
End Function

Public Function CLong(sPrmValue As Variant) As Long
    
On Error GoTo Handler

    'If sPrmValue = "" Then
    '    clong = 0
    'Else
    '    clong = CLng(sPrmValue)
    'End If
    '
    Dim iPos As Integer

    
    If sPrmValue = "" Then
        CLong = 0
    Else
    '/// CLng함수는 0.5 는 0으로 리턴하고 1.5 는 2로 리턴하고 2.5는 2로 리턴하고 3.5는 4로
    '/// 리턴한다. 따라서 0.5를 주어서 0.5를 얻고 싶으면 cDouble함수를 사용해야 한다.
    '/// 싶으면
        iPos = InStr(sPrmValue, ".5")
        If iPos > 0 Then
            If CDbl(sPrmValue) < 0 Then
            CLong = CLng(Left(CStr(sPrmValue), iPos)) - 1
            Else
            CLong = CLng(Left(CStr(sPrmValue), iPos)) + 1
            End If
        Else
            CLong = CLng(CStr(sPrmValue))
        End If
    End If

    Exit Function

Handler:
    CLong = 0
    Resume Next
End Function

Function CRound(sPrmValue As Variant, lPos As Long) As Double
On Error GoTo CR2_Handler
  Dim L As Long, lLog As Long

    If lPos = 0 Then lPos = 1
    lLog = 1
    
    For L = 1 To lPos
         lLog = lLog * 10
    Next L

    If sPrmValue = "" Then
        CRound = 0
    Else
        If CLong((CDbl(sPrmValue) * lLog)) = 0 And Left(piece(sPrmValue, ".", 2), 2) = "00" Then
            CRound = 0.01
        Else
            CRound = CLong((CDbl(sPrmValue) * lLog)) / lLog
        End If
        
    End If

    Exit Function
CR2_Handler:
    CRound = 0
    Resume Next
End Function


Public Function CRounding(sPrmValue As Variant) As Long

On Error GoTo CR_Handler

    If sPrmValue = "" Then
    CRounding = 0
    Else
    CRounding = CLng((sPrmValue * 10 + 5) \ 10)
    End If

    Exit Function

CR_Handler:
    CRounding = 0
    Resume Next

End Function


Public Function cSingle(sPrmValue As Variant) As Single

On Error GoTo CSHandler

    If sPrmValue = "" Then
    cSingle = 0
    Else
    cSingle = CSng(sPrmValue)
    End If

    Exit Function

CSHandler:
    cSingle = 0
    Resume Next

End Function

Public Function CUp(sPrmValue As Variant) As Long

On Error GoTo CUp_Handler

    If sPrmValue = "" Then
        CUp = 0
    Else
        CUp = CLng(CDbl(sPrmValue) * 10 + 9) \ 10
    End If

    Exit Function

CUp_Handler:
    CUp = 0
    Resume Next


End Function

Public Function Data2Format(ByVal sPrmData As Variant, sPrmPict As String) As String
    
    Dim i As Integer
    Dim iDataPos As Integer
    Dim iDataPox As Integer
    Dim iDataLen As Integer
    Dim iPictLen As Integer
    Dim sBufData As String
    Dim sPictStr As String
    Dim sChar As String

    iDataLen = Len(sPrmData)
    iPictLen = Len(sPrmPict)
    
    If Mid(sPrmPict, 1, 1) = "0" Then
    sBufData = Replicate("0", iPictLen - iDataLen) & CStr(sPrmData)
    Else
    sBufData = CStr(sPrmData)
    For i = 1 To iPictLen - iDataLen
        sBufData = sBufData & Space(1)
    Next
    End If
    
    Data2Format = sBufData

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

Public Function DateTimeValidCheck(sPrmDate As String) As Integer
    
    Dim sTmpDate As String

    If Not IsNumeric(sPrmDate) Then
        DateTimeValidCheck = False
        Exit Function
    End If

    If Len(sPrmDate) = 10 Then
        If Not DateValidCheck(Left(sPrmDate, 6)) Or Not TimeValidCheck(Right(sPrmDate, 4)) Then
            DateTimeValidCheck = False
            sPrmDate = ""
            Exit Function
        Else
            sPrmDate = AddCentury(Left(sPrmDate, 6)) & Right(sPrmDate, 4)
        End If
    ElseIf Len(sPrmDate) = 12 Then
        If Not DateValidCheck(Left(sPrmDate, 8)) Or Not TimeValidCheck(Right(sPrmDate, 4)) Then
            DateTimeValidCheck = False
            sPrmDate = ""
            Exit Function
        End If
    Else
        DateTimeValidCheck = False
        sPrmDate = ""
        Exit Function
    End If

    DateTimeValidCheck = True

End Function

Public Function BeforeTime(sPrmTime As String, sPrmDisTime As String) As String
    Dim iTmpHour As Integer
    Dim iTmpMin  As Integer
    
    If (Not IsNumeric(sPrmTime)) Or (Not Len(sPrmTime) = 4) Then
        Exit Function
    End If
    
    iTmpHour = CInteger(Left(sPrmTime, 2))
    iTmpMin = CInteger(Right(sPrmTime, 2))
    
    If (iTmpMin - CInteger(sPrmDisTime)) < 0 Then
        iTmpMin = 60 + (iTmpMin - CInteger(sPrmDisTime))
        If CLong(iTmpHour) = 0 Then
            iTmpHour = 23
        Else
            iTmpHour = iTmpHour - 1
        End If
    Else
        iTmpMin = iTmpMin - CLong(sPrmDisTime)
    End If
    
    BeforeTime = Format(iTmpHour, "00") & Format(iTmpMin, "00")
    
End Function
Public Function DateValidCheck(sPrmDate As String) As Integer
    
    'ex) sPrmDate : 950101

    Dim sTmpDate As String
    
    If Not IsNumeric(sPrmDate) Then
        DateValidCheck = False
        Exit Function
    End If
    
    '자리수 체크도 하자
    If Len(sPrmDate) <> 6 And Len(sPrmDate) <> 8 Then
        DateValidCheck = False
        Exit Function
    End If
    
    '종료일에 사용하는 값은 그냥 True로
    If sPrmDate = "999999" Or sPrmDate = "99999999" Then
        DateValidCheck = True
    Else
        sTmpDate = AddCentury(sPrmDate)
        sTmpDate = Format(sTmpDate, "####/##/##")
        DateValidCheck = IsDate(sTmpDate)
    End If

End Function

Public Sub DeleteLocChtUid(sUidCod As String)
    '담당자가 걸어 놓은 Locking 된 모든 챠트번호를 푼다.
    Dim sCurKey As String, sCmpKey As String, sRetVal As String
    Dim sDelCurkey As String, sDelRetVal  As String
    Dim tLocCht As LocChtRec
    Dim iError As Integer
    
    sCmpKey = sUidCod & Chr(5)
    sCurKey = sCmpKey
    sCurKey = mSetNext("LocChtUid", sCurKey)
    Do
    sCurKey = mReadNext("LocChtUid", sCurKey, sCmpKey, sRetVal)
    If sCurKey = "" Then Exit Do
        Call LocChtLoad(sRetVal, tLocCht)
        Call LocChtStore(sDelCurkey, sDelRetVal, tLocCht)
        iError = mDelete("LocCht", sDelCurkey)
    Loop

End Sub

Public Function DelSpace(ByVal sPrmBufStr As String) As String

    Dim i As Integer, iStrLen As Integer
    Dim sTmpChr As String, sRetStr As String

    iStrLen = Len(sPrmBufStr)
    sRetStr = ""
    For i = 1 To iStrLen
    sTmpChr = Mid(sPrmBufStr, i, 1)
    If sTmpChr <> " " Then
        sRetStr = sRetStr & sTmpChr
    End If
    Next

    DelSpace = sRetStr

End Function

'******************
'   그룹 총괄과
'******************
Public Function DepGrpCode(sPrmDepCod As String, sDate As String) As String
    
    Dim DepData As DepMstRec
    
    Call DepMstRead(sPrmDepCod, sDate, DepData)

    DepGrpCode = DepData.DepGrpCod

End Function

'*************************************************
'   Group총괄과를 읽고 해당과의 과목 코드 return
'*************************************************
Public Function DepGrpNamCode(sEngDepCod As String) As String
    
    Dim DepData As DepMstRec
    Dim sDepMstCurKey As String, sDepMstCmpKey As String, sDepMstRetVal As String

    sDepMstCmpKey = ""
    sDepMstCurKey = sDepMstCmpKey

    sDepMstCurKey = mSetNext("DepMst", sDepMstCurKey)
    Do
    sDepMstCurKey = mReadNext("DepMst", sDepMstCurKey, sDepMstCmpKey, sDepMstRetVal)
    If sDepMstCurKey = "" Then Exit Do

    DepMstLoad sDepMstRetVal, DepData

    If Trim(UCase(DepData.DepEngNam)) = Trim(UCase(sEngDepCod)) Then
        DepGrpNamCode = DepData.DepCod
        Exit Function
    End If
    Loop

    DepGrpNamCode = ""

End Function

Public Function DisplayMessageBox(tPrmMgdData As MsgMstRec, sPrmMsg As String, iPrmFlag As Integer) As Integer
    
    Dim iTmpReturn As Integer
    Dim iMsgType As Integer
    Dim sMsgTitle As String, sMessage As String
    Dim SID As String * 1

    iMsgType = vbOK + vbModal
    
    sMessage = RTrim(tPrmMgdData.MsgCodNam)
    SID = Mid(tPrmMgdData.MsgCod, 4, 1)
    Select Case SID
    Case "I"
    iMsgType = iMsgType + vbInformation
    sMsgTitle = "정보 ["

    Case "W"
    iMsgType = iMsgType + vbExclamation
    sMsgTitle = "경고 ["

    Case "E"
    iMsgType = iMsgType + vbError
    sMsgTitle = "오류 ["

    Case "Q"
    iMsgType = iMsgType - vbOK + vbQuestion + vbYesNo
    sMsgTitle = "선택 ["

    Case "P"
    iMsgType = iMsgType - vbOK + vbQuestion + vbYesNo
    sMsgTitle = "선택 ["

    End Select
    
    iMsgType = iMsgType + iPrmFlag
    
    sMessage = sPrmMsg & sMessage
    sMsgTitle = sMsgTitle & tPrmMgdData.MsgCod & "]"

    DisplayMessageBox = MsgBox(sMessage, iMsgType, sMsgTitle)

End Function

Public Function DisplayMsgBox(tPrmMgdData As MsgMstRec, sPrmMsg As String) As Integer

    Dim iTmpReturn As Integer
    Dim iMsgType As Integer
    Dim sMsgTitle As String, sMessage As String
    Dim SID As String * 1

    iMsgType = vbOKOnly '+ vbModal
    
    sMessage = RTrim(tPrmMgdData.MsgCodNam)
    SID = Mid(tPrmMgdData.MsgCod, 4, 1)
        
    Select Case SID
    Case "I"
        iMsgType = iMsgType + vbInformation
        sMsgTitle = "정보 ["
    
    Case "W"
        iMsgType = iMsgType + vbExclamation
        sMsgTitle = "경고 ["
    
    Case "E"
        iMsgType = iMsgType + vbError
        sMsgTitle = "오류 ["
    
    Case "Q"
        iMsgType = iMsgType - vbOKOnly + vbQuestion + vbYesNo
        sMsgTitle = "선택 ["
    
    Case "P"
        iMsgType = iMsgType - vbOKOnly + vbQuestion + vbYesNo
        sMsgTitle = "선택 ["
    
    End Select

    sMessage = sPrmMsg & sMessage
    sMsgTitle = sMsgTitle & tPrmMgdData.MsgCod & "]"

    DisplayMsgBox = MsgBox(sMessage, iMsgType, sMsgTitle)


End Function

Public Function DnhCheck(sPrmDate As String, sPrmTime As String) As String

    Dim sTmpDate As String
    Dim lTmpTime As Long
    Dim lTmpWeek As Long

    DnhCheck = "D"
    
    If Len(sPrmDate) = 6 Then
        sTmpDate = AddCentury(sPrmDate)
    Else
        sTmpDate = sPrmDate
    End If

    lTmpWeek = IsHoliday(sTmpDate)

    Select Case lTmpWeek
        
        Case 1  '공휴일
           DnhCheck = "H"
        
        Case Else
            lTmpTime = CLong(Left(sPrmTime, 4))
            
            If lTmpWeek = 7 Then
                '2001.07.01...변경안 - 2시간이 늘어났다.
                'If lTmpTime >= 900 And lTmpTime < 1500 Then
                If lTmpTime >= 900 And lTmpTime < 1330 Then  '기존에는 8시25분으로 되어 있었는데
                    DnhCheck = "D"                          '9시로 수정합니다...2000.07.28
                Else
                    DnhCheck = "N"
                End If
                
            Else
                'if lTmpTime >= 900 ad lTmpTime < 2000 then
                If lTmpTime >= 900 And lTmpTime < 1800 Then
                    DnhCheck = "D"
                Else
                    DnhCheck = "N"
                End If
            End If

    End Select

End Function

Public Function DNHCheck2(sPrmAct As String, sPrmDate As String, sPrmTime As String) As String
''
''    Dim sTmpDate As String
''    Dim lTmpTime As Long
''    Dim lTmpWeek As Long
''    Dim lStartTime As Long
''
''    DNHCheck2 = "D"
''
''    If Len(sPrmDate) = 6 Then
''        sTmpDate = AddCentury(sPrmDate)
''    Else
''        sTmpDate = sPrmDate
''    End If
''
''    lTmpWeek = IsHoliday(sTmpDate)
''
''    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
''    '2001/10/29 james
''    '조건부 컴파일을 위하여.... 이런방법을 씁니다...
''    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
''    '상주적십자 병원은 730분 부터 주간시간을 적용한다.
''    #If SANGJU Then
''        lStartTime = 730
''    #Else
''        'lStartTime = 830   '20030109 lek edit
''        '첨단 병원은 아침 8시 부터 외래접수로 하고
''        '오후 7시 부터 야간 당직으로 함
''        lStartTime = 800
''    #End If
''    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
''    Select Case lTmpWeek
''
''    Case 1  '공휴일
''        DNHCheck2 = "H"
''
''    Case Else
''        lTmpTime = CLong(Left(sPrmTime, 4))
''
''        '광주첨단병원 야간 수납..................토요일 13:00, 월~금 : 19:00
''        '토요일이면서 진찰료이면...
''        If lTmpWeek = 7 Then
''            '진찰료...
''            If CLong(sPrmAct) = 1 Then
''                If lTmpTime >= lStartTime And lTmpTime < 1300 Then
''                    DNHCheck2 = "D"
''                Else
''                    DNHCheck2 = "N"
''                End If
''            '기타 처치및 수술
''            Else
''                If lTmpTime >= lStartTime And lTmpTime < 1800 Then
''                    DNHCheck2 = "D"
''                Else
''                    DNHCheck2 = "N"
''                End If
''            End If
''        Else
''            '진찰료...
''            If CLong(sPrmAct) = 1 Then
''
'''                If lTmpTime >= lStartTime And lTmpTime <= 2000 Then          '20021216 이대순 수정 평일(월 ~ 금) 응급접수는 19:00
'''                    DNHCheck2 = "D"                                          '20021228 neverdie 이봉주씨의 요청으로 다시 20:00로변경..나이것참..
'''                Else
'''                    DNHCheck2 = "N"
'''                End If
''                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''                If lTmpTime >= lStartTime And lTmpTime <= 1900 Then          '20021216 이대순 수정 평일(월 ~ 금) 응급접수는 19:00
''                    DNHCheck2 = "D"                                          '20021228 neverdie 이봉주씨의 요청으로 다시 20:00로변경..나이것참..
''                Else
''                    DNHCheck2 = "N"
''                End If
''                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''            Else
''                '기타 처치및 수술
''                If lTmpTime >= lStartTime And lTmpTime <= 1800 Then
''                    DNHCheck2 = "D"
''                Else
''                    DNHCheck2 = "N"
''                End If
''            End If
''        End If
''    End Select
''
''    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
''

    Dim sTmpDate As String
    Dim lTmpTime As Long
    Dim lTmpWeek As Long
    Dim lStartTime As Long

    DNHCheck2 = "D"
    
    If Len(sPrmDate) = 6 Then
        sTmpDate = AddCentury(sPrmDate)
    Else
        sTmpDate = sPrmDate
    End If

    lTmpWeek = IsHoliday(sTmpDate)
    
    #If SungSam = 1 Then
        '대구성삼병원 진료시간셋팅
        lStartTime = 830
    #Else
        lStartTime = 900
    #End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Select Case lTmpWeek
    
    Case 1  '공휴일
        DNHCheck2 = "H"
    
    Case Else
        lTmpTime = CLong(Left(sPrmTime, 4))
    
        '토요일이면서 진찰료이면...
        If lTmpWeek = 7 Then
            '진찰료...
            If CLong(sPrmAct) = 1 Then
                If lTmpTime >= lStartTime And lTmpTime < 1500 Then
                    DNHCheck2 = "D"
                Else
                    DNHCheck2 = "N"
                End If
            '기타 처치및 수술
            Else
                If lTmpTime >= lStartTime And lTmpTime < 1800 Then
                    DNHCheck2 = "D"
                Else
                    DNHCheck2 = "N"
                End If
            End If
        Else
            '진찰료...
            If CLong(sPrmAct) = 1 Then
                If lTmpTime >= lStartTime And lTmpTime <= 2000 Then
                    DNHCheck2 = "D"
                Else
                    DNHCheck2 = "N"
                End If
            Else
                '기타 처치및 수술
                If lTmpTime >= lStartTime And lTmpTime <= 1800 Then
                    DNHCheck2 = "D"
                Else
                    DNHCheck2 = "N"
                End If
            End If
        End If
    End Select
        
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
End Function

'**************************************************
'   2차원 array sort
'--------------------------------------------------
'   1. iPrmRowCount : Sort할 Row의 갯수
'   2. iPrmColCount : Sort할 Column의 갯수
'   3. iPrmKey1     : Sort의 대상이 되는 첫번째 Column
'   4. iPrmKey2     : Sort의 대상이 되는 두번째 Column
'   5. sPrmRowData  : 2차원 array
'**************************************************
Public Sub DoubleSort(iPrmRowCount As Integer, iPrmColCount As Integer, iPrmKey1 As Integer, iPrmKey2 As Integer, sPrmRowData() As String)
    Dim Row As Integer
    Dim col As Integer
    Dim iCnt As Integer
    Dim skey1 As String
    Dim skey2 As String
    Dim sBufValue() As String

    ' 임시 저장 array
    ReDim sBufValue(1 To iPrmColCount)

    If iPrmRowCount = 1 Then Exit Sub

    For iCnt = iPrmRowCount - 1 To 1 Step -1
    For Row = 1 To iCnt
        ' 가장 큰 값을 제일 뒤로 이동
        skey1 = sPrmRowData(Row, iPrmKey1) & Chr(6) & sPrmRowData(Row, iPrmKey2)
        skey2 = sPrmRowData(Row + 1, iPrmKey1) & Chr(6) & sPrmRowData(Row + 1, iPrmKey2)
        If skey1 > skey2 Then
        For col = 1 To iPrmColCount
            sBufValue(col) = sPrmRowData(Row + 1, col)
            sPrmRowData(Row + 1, col) = sPrmRowData(Row, col)
            sPrmRowData(Row, col) = sBufValue(col)
        Next
        End If
    Next
    Next

End Sub

Public Sub DoubleSort3(iPrmRowCount As Integer, iPrmColCount As Integer, iPrmKey1 As Integer, iPrmKey2 As Integer, iPrmKey3 As Integer, sPrmRowData() As String)
    Dim Row As Integer
    Dim col As Integer
    Dim iCnt As Integer
    Dim skey1 As String
    Dim skey2 As String
    Dim sBufValue() As String

    ' 임시 저장 array
    ReDim sBufValue(1 To iPrmColCount)

    If iPrmRowCount = 1 Then Exit Sub

    For iCnt = iPrmRowCount - 1 To 1 Step -1
    For Row = 1 To iCnt
        ' 가장 큰 값을 제일 뒤로 이동
        skey1 = sPrmRowData(Row, iPrmKey1) & Chr(6) & sPrmRowData(Row, iPrmKey2) & Chr(6) & sPrmRowData(Row, iPrmKey3)
        skey2 = sPrmRowData(Row + 1, iPrmKey1) & Chr(6) & sPrmRowData(Row + 1, iPrmKey2) & Chr(6) & sPrmRowData(Row + 1, iPrmKey3)
        If skey1 > skey2 Then
        For col = 1 To iPrmColCount
            sBufValue(col) = sPrmRowData(Row + 1, col)
            sPrmRowData(Row + 1, col) = sPrmRowData(Row, col)
            sPrmRowData(Row, col) = sBufValue(col)
        Next
        End If
    Next
    Next

End Sub

Public Function DptCheck(sPrmMsgCode As String, sPrmChrNum As String, sPrmOcmNum As String, sPrmDepCod As String, sPrmAcpDtm As String) As String
    
    Dim OcmData As OcmInfRec
    Dim StbData As StbInfRec
    Dim IcmData As IcmInfRec
    Dim DepData As DepMstRec
    Dim sTmpAcpDtm As String
    Dim sTmpCurDtm As String
    Dim sTmpAdmDur As String
    Dim sTmpSavDate As String
    Dim lTmpAdmDur  As Long

    Dim sDepMstGrpCurKey As String
    Dim sDepMstGrpCmpKey As String
    Dim sDepMstCurKey As String

    Dim sPrmRetVal As String
    Dim sBufDepValue As String
    Dim sBufOcmValue As String

    Dim sOcmChrDepAcmCurKey As String
    Dim sOcmChrDepAcmCmpKey As String

    Dim sIcmInfChtDtmCurKey As String
    Dim sIcmInfChtDtmCmpKey As String
    Dim sIcmInfChtDtmRetVal As String
    
    Dim sOcmChrAcpCurKey As String
    Dim sOcmChrAcpCmpKey As String
    Dim sBufCurKey As String
    Dim iTmp As Integer
    Dim sGrpDepCod As String            'Group 총괄과

    If Len(Trim(sPrmAcpDtm)) < 12 Then
    sPrmAcpDtm = AddCentury(Left(sPrmAcpDtm, 6)) & Right(sPrmAcpDtm, 4)
    End If

    'Read 진료과 DateBase (DepMst, 진료과)
    Call DepMstRead(sPrmDepCod, Left(sPrmAcpDtm, 8), DepData)

    sDepMstGrpCurKey = DepData.DepGrpCod & Chr(5)
    sDepMstGrpCurKey = mSetNext("DepMstGrp", sDepMstGrpCurKey)
    sDepMstGrpCmpKey = DepData.DepGrpCod
    sGrpDepCod = DepData.DepGrpCod
    sTmpSavDate = ""

    Do
    'Read 진료과 DateBase(DepMstGrp,그룹총괄과 순)
    'sDepMstGrpCurKey = sBufDepValue & Chr(5)
    sDepMstGrpCurKey = mReadNext("DepMstGrp", sDepMstGrpCurKey, sDepMstGrpCmpKey, sPrmRetVal)
    If sDepMstGrpCurKey <> "" Then
        DepMstLoad sPrmRetVal, DepData

        '현재 내원한 진료과와 과거의 내원한 최종 진료과를 비교후 과초진, 재진, 30일 초진을
        '자동 산정한다.

        'Read 내원인적사항 DateBase(OcmChrDepAcm,챠트번호,진료과,내원일시 순)
        sOcmChrDepAcmCurKey = sPrmChrNum & Chr(5) & DepData.DepCod & Chr(5) & "999999999999" & Chr(5)
        sOcmChrDepAcmCurKey = mSetPrev("OcmInfChtDepAcp", sOcmChrDepAcmCurKey)
        sOcmChrDepAcmCmpKey = sPrmChrNum & Chr(5) & DepData.DepCod

        Do
            sOcmChrDepAcmCurKey = mReadPrev("OcmInfChtDepAcp", sOcmChrDepAcmCurKey, sOcmChrDepAcmCmpKey, sPrmRetVal)
            'sOcmChrDepAcmCurKey = sBufOcmValue
    
            If sOcmChrDepAcmCurKey <> "" Then
                OcmInfLoad sPrmRetVal, OcmData
                If OcmData.OcmComStt <> "OC" Then
                If sTmpSavDate < OcmData.OcmAcpDtm And sPrmOcmNum <> OcmData.OcmNum Then
                    sTmpSavDate = OcmData.OcmAcpDtm
                End If
    
                If sPrmOcmNum <> OcmData.OcmNum Then
                    If Left(sPrmAcpDtm, 8) = Left(OcmData.OcmAcpDtm, 8) Then
                    sPrmMsgCode = "HNTE001"
                    sBufCurKey = mSetReadNext("StbInfOcm", OcmData.OcmNum & Chr(5), OcmData.OcmNum & Chr(5), sPrmRetVal)
                    StbInfLoad sPrmRetVal, StbData
                    If StbData.StbEmgYon = "Y" Then
                        DptCheck = "6"
                        sPrmMsgCode = "HNTI063"
                    End If
                    Exit Function
                    End If
                End If
                End If
            Else
                Exit Do
            End If
        Loop
    Else
        Exit Do
    End If
    Loop

    '입원 한것에 관한 일자를 Check한다!
    sIcmInfChtDtmCurKey = sPrmChrNum & Chr(5) & "999999999999"
    sIcmInfChtDtmCmpKey = sPrmChrNum & Chr(5)
    sIcmInfChtDtmCurKey = mSetPrev("IcmInfChtDtm", sIcmInfChtDtmCurKey)
    Do
    sIcmInfChtDtmCurKey = mReadPrev("IcmInfChtDtm", sIcmInfChtDtmCurKey, sIcmInfChtDtmCmpKey, sIcmInfChtDtmRetVal)
    If sIcmInfChtDtmCurKey = "" Then Exit Do
        IcmInfLoad sIcmInfChtDtmRetVal, IcmData

        If IcmData.IcmAcpStt <> "IA" Then
        'Read 진료과 DateBase (DepMst, 진료과)
        Call DepMstRead(IcmData.IcmDepCod, Left(sPrmAcpDtm, 8), DepData)
        'Group 총괄과를 읽는다!
            If sGrpDepCod = DepData.DepGrpCod Then
            '최근것을 할당한다!
            If IcmData.IcmLevDtm <> "999999999999" And CDouble(sTmpSavDate) < CDouble(IcmData.IcmLevDtm) Then
                sTmpSavDate = IcmData.IcmLevDtm
            End If
            End If
        End If
    Loop

    If sTmpSavDate <> "" Then

    'Call Julian Date
    sTmpCurDtm = AddCentury(SystemDate())
    sTmpSavDate = Mid(sTmpSavDate, 1, 8)

    Call Caljulian(sTmpSavDate, sTmpCurDtm, sTmpAdmDur)
    lTmpAdmDur = CLong(sTmpAdmDur)

    '30일 Check 일자는 내원일의 기간 + 투약일수로 한다.
    lTmpAdmDur = lTmpAdmDur '- CLong(OcmData.OcmMdcDay)

    If lTmpAdmDur > 30 Then
        '30일 초진
        DptCheck = "3"
    Else
        '재진
        DptCheck = "4"
    End If
    Else
    sBufCurKey = sPrmChrNum & Chr(5)
    sBufCurKey = mSetReadEqual("PbsInf", sBufCurKey, sPrmRetVal)

    If sBufCurKey <> "" Then
        '과초진
        DptCheck = "2"
    Else
        '초진
        DptCheck = "1"
    End If
    End If

End Function

Public Function DptCheck2(sPrmMsgCode As String, sPrmChrNum As String, sPrmOcmNum As String, sPrmDepCod As String, sPrmAcpDtm As String, sPrmLstDtm As String, DepData As DepMstRec, sPrmMdcDay As String) As String
    
    Dim iTmp    As Integer
    Dim OcmData As OcmInfRec
    Dim StbData As StbInfRec
    Dim IcmData As IcmInfRec
    
    Dim sTmpAcpDtm As String
    Dim sTmpCurDtm As String
    Dim sTmpAdmDur As String
    Dim sTmpSavDate As String
    Dim lTmpAdmDur  As Long

    Dim sDepMstGrpCurKey As String
    Dim sDepMstGrpCmpKey As String
    Dim sDepMstCurKey As String

    Dim sPrmRetVal As String
    Dim sBufDepValue As String
    Dim sBufOcmValue As String

    Dim sOcmChrDepAcmCurKey As String
    Dim sOcmChrDepAcmCmpKey As String

    Dim sIcmInfChtDtmCurKey As String
    Dim sIcmInfChtDtmCmpKey As String
    Dim sIcmInfChtDtmRetVal As String
    
    Dim sOcmChrAcpCurKey As String
    Dim sOcmChrAcpCmpKey As String
    Dim sBufCurKey As String
    Dim sBufCmpKey As String
    Dim sBufRetVal As String
    Dim sGrpDepCod As String            'Group 총괄과

    Dim sOrpCurKey As String
    Dim sOrpCmpKey As String
    Dim sOrpRetVal As String
    Dim OrpData As OrpInfRec
    
    Dim iComp As Integer
    Dim sDBName    As String
    Dim sDepBilCod As String
    
    If Len(Trim(sPrmAcpDtm)) < 12 Then
        sPrmAcpDtm = AddCentury(Left(sPrmAcpDtm, 6)) & Right(sPrmAcpDtm, 4)
    End If

    'Read 진료과 DateBase (DepMst, 진료과)
    Call DepMstRead(sPrmDepCod, Left(sPrmAcpDtm, 8), DepData)

    
    '----------------------------------------------------------------------------
    '2003.04.09 대구성삼병원 심주영
    '----------------------------------------------------------------------------
    '초재진여부구분을 그룹총괄과로 하지 않고 청구코드로 바꾼다...
    '그리고 OCS대기자 띄우는 부분을 그룹총괄과로 한다.
    'OS환자는 외과분야로 같이 띄우지말고 따로 구분해 달라는 요청에 의하여...
    'DepMstBilCod <-- 요거는 인덱스 새로걸었따....
    '----------------------------------------------------------------------------
    #If SungSam = 1 Then
        sDBName = "DepMstBilCod"
        sDepMstGrpCurKey = DepData.DepBilCod & Chr(5)
        sDepMstGrpCurKey = mSetNext(sDBName, sDepMstGrpCurKey)
        sDepMstGrpCmpKey = DepData.DepBilCod
        sDepBilCod = DepData.DepBilCod
    #Else
        sDBName = "DepMstGrp"
        sDepMstGrpCurKey = DepData.DepGrpCod & Chr(5)
        sDepMstGrpCurKey = mSetNext(sDBName, sDepMstGrpCurKey)
        sDepMstGrpCmpKey = DepData.DepGrpCod
        sGrpDepCod = DepData.DepGrpCod
    #End If
    
    
    sTmpSavDate = ""

    Do
        
        'Read 진료과 DateBase(DepMstGrp,그룹총괄과 순)
        'sDepMstGrpCurKey = sBufDepValue & Chr(5)
        'sDepMstGrpCurKey = mReadNext("DepMstGrp", sDepMstGrpCurKey, sDepMstGrpCmpKey, sPrmRetVal)
        sDepMstGrpCurKey = mReadNext(sDBName, sDepMstGrpCurKey, sDepMstGrpCmpKey, sPrmRetVal)
        If sDepMstGrpCurKey <> "" Then
            
            DepMstLoad sPrmRetVal, DepData
            
            '현재 내원한 진료과와 과거의 내원한 최종 진료과를 비교후 과초진, 재진, 30일 초진을
            '자동 산정한다.
    
            'Read 내원인적사항 DateBase(OcmChrDepAcm,챠트번호,진료과,내원일시 순)
            'sOcmChrDepAcmCurKey = sPrmChrNum & Chr(5) & DepData.DepCod & Chr(5) & "999999999999" & Chr(5)
            'sOcmChrDepAcmCurKey = sPrmChrNum & Chr(5) & DepData.DepCod & Chr(5) & sPrmAcpDtm & Chr(5)
            '////같은시간대에 다시 접수하면 이 아이가 제대로 처리를 하지 못한다... 그래서 동일 2과 접수가 된다... 이럴수가....
            sOcmChrDepAcmCurKey = sPrmChrNum & Chr(5) & DepData.DepCod & Chr(5) & Left(sPrmAcpDtm, 8) & "9999" & Chr(5)
            sOcmChrDepAcmCurKey = mSetPrev("OcmInfChtDepAcp", sOcmChrDepAcmCurKey)
            sOcmChrDepAcmCmpKey = sPrmChrNum & Chr(5) & DepData.DepCod
    
            Do
            
                sOcmChrDepAcmCurKey = mReadPrev("OcmInfChtDepAcp", sOcmChrDepAcmCurKey, sOcmChrDepAcmCmpKey, sPrmRetVal)
                'sOcmChrDepAcmCurKey = sBufOcmValue
        
                If sOcmChrDepAcmCurKey <> "" Then
                    
                    OcmInfLoad sPrmRetVal, OcmData
                    
                    '''기타수납으로 당일 금액이 들어온 사람은...동일과 내원으로 뜨면 안된다.
                    iComp = True
                    
                    sOrpCurKey = OcmData.OcmChtNum & Chr(5) & OcmData.OcmNum & Chr(5)
                    sOrpCmpKey = sOrpCurKey
                    sOrpCurKey = mSetNext("OrpInfCht", sOrpCurKey)
                    Do
                        sOrpCurKey = mReadNext("OrpInfCht", sOrpCurKey, sOrpCmpKey, sOrpRetVal)
                        If sOrpCurKey = "" Then Exit Do
                            
                        Call OrpInfLoad(sOrpRetVal, OrpData)
                                                            
                        If Left(OcmData.OcmAcpDtm, 8) = Left(OrpData.OrpUpdDtm, 8) Then
                            If OrpData.OrpRvnTyp = "E" Or OrpData.OrpRvnTyp = "T" Then
                                iComp = False
                            End If
                        End If
                    Loop
                                            
                    If iComp = True Then
                    
                        If OcmData.OcmComStt <> "OC" Then
                            
                            If sTmpSavDate < OcmData.OcmAcpDtm And sPrmOcmNum <> OcmData.OcmNum Then
                                sTmpSavDate = OcmData.OcmAcpDtm     '내원일자
                                sPrmMdcDay = OcmData.OcmMdcDay      '투약일수
                            End If
            
                            If sPrmOcmNum <> OcmData.OcmNum Then
                                
                                If Left(sPrmAcpDtm, 8) = Left(OcmData.OcmAcpDtm, 8) Then
                                    
                                    sPrmMsgCode = "HNTE001"
                                    sBufCurKey = mSetReadNext("StbInfOcm", OcmData.OcmNum & Chr(5), OcmData.OcmNum & Chr(5), sPrmRetVal)
                                    StbInfLoad sPrmRetVal, StbData
                                    
                                    If StbData.StbEmgYon = "Y" Then
                                        DptCheck2 = "6"
                                        sPrmMsgCode = "HNTI063"
                                    End If
                                    
                                    Exit Function
                                
                                End If
                            
                            End If
                        
                        End If
                    End If
                Else
                    
                    Exit Do
                
                End If
                
            Loop
        
        Else
            
            Exit Do
        
        End If
    
    Loop

    '입원 한것에 관한 일자를 Check한다!
    'sIcmInfChtDtmCurKey = sPrmChrNum & Chr(5) & "999999999999"
    sIcmInfChtDtmCurKey = sPrmChrNum & Chr(5) & sPrmAcpDtm
    sIcmInfChtDtmCmpKey = sPrmChrNum & Chr(5)
    sIcmInfChtDtmCurKey = mSetPrev("IcmInfChtDtm", sIcmInfChtDtmCurKey)

    Do
        sIcmInfChtDtmCurKey = mReadPrev("IcmInfChtDtm", sIcmInfChtDtmCurKey, sIcmInfChtDtmCmpKey, sIcmInfChtDtmRetVal)
        If sIcmInfChtDtmCurKey = "" Then Exit Do
        
        IcmInfLoad sIcmInfChtDtmRetVal, IcmData
        If IcmData.IcmAcpStt <> "IA" Then
            'Read 진료과 DateBase (DepMst, 진료과)
            Call DepMstRead(IcmData.IcmDepCod, Left(sPrmAcpDtm, 8), DepData) '
'            'Group 총괄과를 읽는다!
'            If sGrpDepCod = DepData.DepGrpCod Then'
            '------------------------------------------------------------------
            '대구성삼병원 심주영
            '------------------------------------------------------------------
                #If SungSam = 1 Then
                    If sDepBilCod = DepData.DepBilCod Then  '청구코드를 읽는다!
                #Else
                    If sGrpDepCod = DepData.DepGrpCod Then  'Group 총괄과를 읽는다!
                #End If
                
            '최근것을 할당한다!
            If IcmData.IcmLevDtm <> "999999999999" And CDouble(sTmpSavDate) < CDouble(IcmData.IcmLevDtm) Then
                sTmpSavDate = IcmData.IcmLevDtm
                sPrmMdcDay = ""
            End If
        End If
    End If
    Loop

    If sTmpSavDate <> "" Then

        'Call Julian Date
        'sTmpCurDtm = AddCentury(SystemDate())
        sTmpCurDtm = sPrmAcpDtm
        '-------------------------------------------------
        '이 한줄만 추가했다!!!!!(1996년12월12일)
        '-------------------------------------------------
        sPrmLstDtm = sTmpSavDate
        sTmpSavDate = Mid(sTmpSavDate, 1, 8)
    
        Call Caljulian(sTmpSavDate, sTmpCurDtm, sTmpAdmDur)
        lTmpAdmDur = CLong(sTmpAdmDur)
    
        '30일 Check 일자는 내원일의 기간 - 투약일수로 한다.
        lTmpAdmDur = lTmpAdmDur '- CLong(OcmData.OcmMdcDay)
    
        '----------------------------------------------------------------
        '2003.06.13 대구성삼병원 심쭈
        '성삼병원에서는 90일 초진 사용하지 않음.
        '----------------------------------------------------------------
        #If SungSam = 1 Then
                If lTmpAdmDur < 0 Then
                    DptCheck2 = "2"
                Else
                    DptCheck2 = "4"         '재진
                End If
        #Else
                If lTmpAdmDur < 0 Then
                    DptCheck2 = "2"
                ElseIf lTmpAdmDur > 90 Then
                    DptCheck2 = "3"         '90일 초진
                Else
                    DptCheck2 = "4"         '재진
                End If
        
        #End If
        '----------------------------------------------------------------
    Else
        '신환체크 루틴을 변경 (내원기준__MarsMan__990107)
        '>>>기존루틴
        'sBufCurKey = sPrmChrNum & Chr(5)
        'sBufCurKey = mSetReadEqual("PbsInf", sBufCurKey, sPrmRetVal)
        '>>>신규루틴
        sBufCmpKey = sPrmChrNum & Chr(5)
        sBufCurKey = sBufCmpKey & CStr(Val(sPrmAcpDtm) - 1) & Chr(5)
        sBufCurKey = mSetPrev("OcmInfChtDtm", sBufCurKey)
        sBufCurKey = mReadPrev("OcmInfChtDtm", sBufCurKey, sBufCmpKey, sBufRetVal)
        
        If sBufCurKey <> "" Then
            DptCheck2 = "2"         '과초진
        Else
            DptCheck2 = "1"         '신환
        End If
    
    End If

    '2001/09/06 의료정보사업부 개발과 서효기
    '덕산병원에서 소아과는 신환환자를 제외하고
    '모두 재진으로 처리한다.
    'If sGrpDepCod = "PED" Then
    '    DptCheck2 = "4"         '재진
    'End If
    
End Function

'---------------------------------------------------------------------------------------------------
'   해당과의 최종내원일을 전달한다!
'---------------------------------------------------------------------------------------------------
Public Function DptCheck2Old(sPrmMsgCode As String, sPrmChrNum As String, sPrmOcmNum As String, sPrmDepCod As String, sPrmAcpDtm As String, sPrmLstDtm As String, DepData As DepMstRec, sPrmMdcDay As String) As String
    
    Dim iTmp    As Integer
    Dim OcmData As OcmInfRec
    Dim StbData As StbInfRec
    Dim IcmData As IcmInfRec
    
    Dim sTmpAcpDtm As String
    Dim sTmpCurDtm As String
    Dim sTmpAdmDur As String
    Dim sTmpSavDate As String
    Dim lTmpAdmDur  As Long

    Dim sDepMstGrpCurKey As String
    Dim sDepMstGrpCmpKey As String
    Dim sDepMstCurKey As String

    Dim sPrmRetVal As String
    Dim sBufDepValue As String
    Dim sBufOcmValue As String

    Dim sOcmChrDepAcmCurKey As String
    Dim sOcmChrDepAcmCmpKey As String

    Dim sIcmInfChtDtmCurKey As String
    Dim sIcmInfChtDtmCmpKey As String
    Dim sIcmInfChtDtmRetVal As String
    
    Dim sOcmChrAcpCurKey As String
    Dim sOcmChrAcpCmpKey As String
    Dim sBufCurKey As String
    Dim sGrpDepCod As String            'Group 총괄과

    If Len(Trim(sPrmAcpDtm)) < 12 Then
    sPrmAcpDtm = AddCentury(Left(sPrmAcpDtm, 6)) & Right(sPrmAcpDtm, 4)
    End If

    'Read 진료과 DateBase (DepMst, 진료과)
    Call DepMstRead(sPrmDepCod, Left(sPrmAcpDtm, 8), DepData)

    sDepMstGrpCurKey = DepData.DepGrpCod & Chr(5)
    sDepMstGrpCurKey = mSetNext("DepMstGrp", sDepMstGrpCurKey)
    sDepMstGrpCmpKey = DepData.DepGrpCod
    sGrpDepCod = DepData.DepGrpCod
    sTmpSavDate = ""

    Do
    'Read 진료과 DateBase(DepMstGrp,그룹총괄과 순)
    'sDepMstGrpCurKey = sBufDepValue & Chr(5)
    sDepMstGrpCurKey = mReadNext("DepMstGrp", sDepMstGrpCurKey, sDepMstGrpCmpKey, sPrmRetVal)
    If sDepMstGrpCurKey <> "" Then
        DepMstLoad sPrmRetVal, DepData
        
        '현재 내원한 진료과와 과거의 내원한 최종 진료과를 비교후 과초진, 재진, 30일 초진을
        '자동 산정한다.

        'Read 내원인적사항 DateBase(OcmChrDepAcm,챠트번호,진료과,내원일시 순)
        sOcmChrDepAcmCurKey = sPrmChrNum & Chr(5) & DepData.DepCod & Chr(5) & "999999999999" & Chr(5)
        sOcmChrDepAcmCurKey = mSetPrev("OcmInfChtDepAcp", sOcmChrDepAcmCurKey)
        sOcmChrDepAcmCmpKey = sPrmChrNum & Chr(5) & DepData.DepCod

        Do
        sOcmChrDepAcmCurKey = mReadPrev("OcmInfChtDepAcp", sOcmChrDepAcmCurKey, sOcmChrDepAcmCmpKey, sPrmRetVal)
        'sOcmChrDepAcmCurKey = sBufOcmValue

        If sOcmChrDepAcmCurKey <> "" Then
            OcmInfLoad sPrmRetVal, OcmData
            If OcmData.OcmComStt <> "OC" Then
            If sTmpSavDate < OcmData.OcmAcpDtm And sPrmOcmNum <> OcmData.OcmNum Then
                sTmpSavDate = OcmData.OcmAcpDtm     '내원일자
                sPrmMdcDay = OcmData.OcmMdcDay      '투약일수
            End If

            If sPrmOcmNum <> OcmData.OcmNum Then
                If Left(sPrmAcpDtm, 8) = Left(OcmData.OcmAcpDtm, 8) Then
                sPrmMsgCode = "HNTE001"
                sBufCurKey = mSetReadNext("StbInfOcm", OcmData.OcmNum & Chr(5), OcmData.OcmNum & Chr(5), sPrmRetVal)
                StbInfLoad sPrmRetVal, StbData
                If StbData.StbEmgYon = "Y" Then
                    DptCheck2Old = "6"
                    sPrmMsgCode = "HNTI063"
                End If
                Exit Function
                End If
            End If
            End If
        Else
            Exit Do
        End If
        Loop
    Else
        Exit Do
    End If
    Loop

    '입원 한것에 관한 일자를 Check한다!
    sIcmInfChtDtmCurKey = sPrmChrNum & Chr(5) & "999999999999"
    sIcmInfChtDtmCmpKey = sPrmChrNum & Chr(5)
    sIcmInfChtDtmCurKey = mSetPrev("IcmInfChtDtm", sIcmInfChtDtmCurKey)
    Do
    sIcmInfChtDtmCurKey = mReadPrev("IcmInfChtDtm", sIcmInfChtDtmCurKey, sIcmInfChtDtmCmpKey, sIcmInfChtDtmRetVal)
    If sIcmInfChtDtmCurKey = "" Then Exit Do
    
    IcmInfLoad sIcmInfChtDtmRetVal, IcmData
    If IcmData.IcmAcpStt <> "IA" Then
        'Read 진료과 DateBase (DepMst, 진료과)
        Call DepMstRead(IcmData.IcmDepCod, Left(sPrmAcpDtm, 8), DepData)
        'Group 총괄과를 읽는다!
        If sGrpDepCod = DepData.DepGrpCod Then
            '최근것을 할당한다!
            If IcmData.IcmLevDtm <> "999999999999" And CDouble(sTmpSavDate) < CDouble(IcmData.IcmLevDtm) Then
            sTmpSavDate = IcmData.IcmLevDtm
            sPrmMdcDay = ""
            End If
        End If
    End If
    Loop

    If sTmpSavDate <> "" Then

    'Call Julian Date
    sTmpCurDtm = AddCentury(SystemDate())
    '-------------------------------------------------
    '이 한줄만 추가했다!!!!!(1996년12월12일)
    '-------------------------------------------------
    sPrmLstDtm = sTmpSavDate
    sTmpSavDate = Mid(sTmpSavDate, 1, 8)

    Call Caljulian(sTmpSavDate, sTmpCurDtm, sTmpAdmDur)
    lTmpAdmDur = CLong(sTmpAdmDur)

    '30일 Check 일자는 내원일의 기간 - 투약일수로 한다.
    lTmpAdmDur = lTmpAdmDur '- CLong(OcmData.OcmMdcDay)

    If lTmpAdmDur > 30 Then
        DptCheck2Old = "3"         '30일 초진
    Else
        DptCheck2Old = "4"         '재진
    End If
    Else
    sBufCurKey = sPrmChrNum & Chr(5)
    sBufCurKey = mSetReadEqual("PbsInf", sBufCurKey, sPrmRetVal)
    If sBufCurKey <> "" Then
        DptCheck2Old = "2"         '과초진
    Else
        DptCheck2Old = "1"         '초진
    End If
    End If

End Function


Public Function Dtm2DateAndTime(sPrmDtm As String, sPrmDte As String, sPrmTim As String) As String
' YYYYMMDDHHMM 형식의 일시타입을 YYYY년 M월 D일 오전[오후] H시 (M분) 으로 리턴
' sPrmDte 에 YYYY년 M월 D일
' sPrmTim 에 오전(오후) H시 (M분) 을 넘겨준다.

    Dim sBufDte As String, sBufTim As String
    Dim sFmtDte As String, sFmtTim As String
    Dim iTmpHour As Integer, iTmpMin As Integer

    sBufDte = Left(sPrmDtm, 8)
    sBufTim = Right(sPrmDtm, 4)

    sFmtDte = Format(Left(sBufDte, 4), "####") & "년 "
    sFmtDte = sFmtDte & CStr(CLong(Mid(sBufDte, 5, 2))) & "월 "
    sFmtDte = sFmtDte & CStr(CLong(Right(sBufDte, 2))) & "일"

    iTmpHour = CInteger(Left(sBufTim, 2))
    If iTmpHour > 12 Then
    sFmtTim = "오후 " & CStr(iTmpHour - 12) & "시 "
    Else
    sFmtTim = "오전 " & CStr(iTmpHour) & "시 "
    End If
    
    iTmpMin = CInteger(Right(sBufTim, 2))
    If iTmpMin <> 0 Then
    sFmtTim = sFmtTim & CStr(iTmpMin) & "분"
    Else
    sFmtTim = Trim(sFmtTim)
    End If

    sPrmDte = sFmtDte
    sPrmTim = sFmtTim
    Dtm2DateAndTime = sFmtDte & " " & sFmtTim

End Function

Public Function EmrAutoCheck(sSysDtm As String, tDepData As DepMstRec)

    Dim sEmrFg   As String
    Dim sBufDate As String
    Dim sBufTime As String

    sBufDate = Left(sSysDtm, 8)
    sBufTime = Right(sSysDtm, 4)
    
    sEmrFg = DNHCheck2("1", sBufDate, sBufTime)

    Select Case sEmrFg
    Case "D"
        EmrAutoCheck = False
    Case "N", "H"
        EmrAutoCheck = True
    Case Else
        EmrAutoCheck = False
    End Select

    'If tMthData.MthCod <> "" Then
    '    Select Case tMthData.MthGrpCod
    '    Case "NB"                       '신생아는 응급실로 접수 하지 않는다.
    '        EmrAutoCheck = False
    '    Case Else
    '        EmrAutoCheck = 2            '설정된
    '        'No action
    '    End Select
    'Else
    '    EmrAutoCheck = 2            '설정된
    'End If

End Function

Public Sub HanOn(Src As Object)
    '한글 IME Mode
    
    On Error Resume Next
    
    Dim hIME As Long

    hIME = ImmGetContext(Src.hWnd)
    ImmSetConversionStatus hIME, IME_HANGUL, IME_NONE
    DoEvents
    Src.SetFocus
    
End Sub

Public Sub EngOn(Src As Object)
    '영문 IME Mode
    On Error Resume Next
    
    Dim hIME As Long

    hIME = ImmGetContext(Src.hWnd)
    ImmSetConversionStatus hIME, IME_ENGLISH, IME_NONE
    DoEvents
    Src.SetFocus
    
End Sub

Public Function FinalNumberSetting(sPrmFnlCod As String, Optional sPrmDate As String = "") As String

    Dim i As Long
    Dim FnlData As FnlMstRec
    Dim sBufKey As String
    Dim sBufValue As String
    Dim iTmpSeq As Long
    Dim sFnlNum As String

    Dim sFnlMstCurKey As String
    Dim sFnlMstCmpKey As String
    Dim sPrmRetVal As String

    Dim sBufDate As String
    
FinalNumberSetting_DASI:

    'Locking Routine (mWrite 의 return값은 True or False)
    For i = 1 To 60000          '10000번 test  후 30000으로 수정 했으나 요즘은 Client PC 사양이 좋아져서 횟수를 늘려야 한다. ---꼽냐?
        If mWrite("LckMst", sPrmFnlCod, sPrmFnlCod) Then
            Exit For
        End If
    Next

    FnlData.FnlCod = sPrmFnlCod
    Call FnlMstStore(sBufKey, sBufValue, FnlData)
    sFnlMstCurKey = sBufKey

    sBufValue = mSetReadEqual("FnlMst", sFnlMstCurKey, sPrmRetVal)

    If sBufValue <> "" Then
        FnlMstLoad sPrmRetVal, FnlData
        'If sPrmFnlCod = "MEDNUM" Or sPrmFnlCod = "INJNUM" Or sPrmFnlCod = "PHYNUM" Then
        Select Case sPrmFnlCod
        Case "MEDNUM", "INJNUM", "PHYNUM", "BOFNUM", "GASNUM", "CYTNUM", "PARNUM", _
             "SSTNUM", "COANUM", "ICHNUM", "ISENUM", "OUTNUM", "PACSNUM", "XRYNUM", _
             "BNKNUM", "CHENUM", "FLUNUM", "HEMNUM", "IMMNUM", "MICNUM", "REFNUM", "SERNUM", "URINUM"           '검사실 검체번호
            
            If sPrmDate = "" Then
                sBufDate = SystemDate()
            Else
                sBufDate = sPrmDate
            End If
            
            If sBufDate = FnlData.FnlDte Then
                FnlData.FnlNum = CStr(CLong(FnlData.FnlNum) + 1)
                FnlData.FnlDte = sBufDate
                sFnlNum = FnlData.FnlNum
            Else
                FnlData.FnlNum = "1"
                FnlData.FnlDte = sBufDate
                sFnlNum = "1"
            End If
'---------------------------------------------> 추가
'        '20040102..HTS..add
'        Case "RCPNUM"
'            If sPrmDate = "" Then
'                    sBufDate = SystemDate()
'                Else
'                    sBufDate = sPrmDate
'            End If
'
'            If Left(sBufDate, 4) = FnlData.FnlDte Then
'                FnlData.FnlNum = Format(CStr(CDouble(FnlData.FnlNum) + 1), "0#########")
'                FnlData.FnlDte = Left(sBufDate, 4)
'                sFnlNum = FnlData.FnlNum
'            Else
'                FnlData.FnlNum = Left(sBufDate, 4) & "000001"
'                FnlData.FnlDte = Left(sBufDate, 4)
'                sFnlNum = FnlData.FnlNum
'            End If
'---------------------------------------------> 추가
        Case Else
            FnlData.FnlNum = CStr(CLong(FnlData.FnlNum) + 1)
            sFnlNum = FnlData.FnlNum
            
        End Select
    Else
        Select Case sPrmFnlCod
        Case "MEDNUM", "INJNUM", "PHYNUM", "BOFNUM", "GASNUM", "CYTNUM", "PARNUM", _
             "SSTNUM", "COANUM", "ICHNUM", "ISENUM", "OUTNUM", "PARNUM_", "XRUNUM", _
             "BNKNUM", "CHENUM", "FLUNUM", "HEMNUM", "IMMNUM", "MICNUM", "REFNUM", "SERNUM", "URINUM"           '검사실 검체번호
        
            If sPrmDate = "" Then
                sBufDate = AddCenturyLen(SystemDate())
            Else
                sBufDate = sPrmDate
            End If
            
            FnlData.FnlNum = "1"
            FnlData.FnlDte = sBufDate
            FnlData.FnlCod = sPrmFnlCod
            sFnlNum = "1"
            
        Case Else
            FnlData.FnlNum = "1"
            FnlData.FnlCod = sPrmFnlCod
            sFnlNum = FnlData.FnlNum
            
        End Select
    End If

    Call FnlMstStore(sBufKey, sBufValue, FnlData)

    iTmpSeq = mWrite("FnlMst", sBufKey, sBufValue)
    If iTmpSeq = False Then
        iTmpSeq = mUpdate("FnlMst", sBufKey, sBufValue)
    End If

    'Locking 해제
    iTmpSeq = mDelete("LckMst", sPrmFnlCod)

    If CDouble(sFnlNum) = 0 Then
        MsgBox "최종번호 생성에 실패 했습니다. 확인 하시면 다시 시도 합니다."
        GoTo FinalNumberSetting_DASI
    End If

    FinalNumberSetting = sFnlNum

End Function
Public Function FinalOutNumberSetting(Optional sPrmDate As String = "") As String

    Dim i As Integer
    Dim iTmpSeq As Integer
    
    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sRetVal As String
    
    Dim sUpdCurKey As String
    Dim sUpdRetVal As String
    Dim sFnlNum As String
    Dim OutData As OutMstRec
    
    'Locking Routine (mWrite 의 return값은 True or False)
    For i = 1 To 60000          '10000번 test  후 30000으로 수정 했으나 요즘은 Client PC 사양이 좋아져서 횟수를 늘려야 한다. ---꼽냐?
        If mWrite("LckMst", "OutMst", "OutMst") Then
            Exit For
        End If
    Next
    
    sCurKey = sPrmDate & Chr(5)
    sCurKey = mSetReadEqual("OutMst", sCurKey, sRetVal)
    
    If sCurKey <> "" Then
        OutMstLoad sRetVal, OutData
        
        OutData.OutNum = CStr(CLong(OutData.OutNum) + 1)
        OutData.OutUpdDtm = SystemDate() & SystemTime()
        sFnlNum = OutData.OutNum
        Call OutMstStore(sUpdCurKey, sUpdRetVal, OutData)
        
        iTmpSeq = mWrite("OutMst", sUpdCurKey, sUpdRetVal)
        If iTmpSeq = False Then
            iTmpSeq = mUpdate("OutMst", sCurKey, sRetVal)
        End If
    Else
        OutData.OutOdrDte = sPrmDate
        OutData.OutNum = "1"
        OutData.OutUpdDtm = SystemDate() & SystemTime()
        
        sFnlNum = OutData.OutNum
    End If

    'Locking 해제
    iTmpSeq = mDelete("LckMst", "OutMst")
    
    FinalOutNumberSetting = sFnlNum

End Function
Public Function FinalNumberSettingOld(sPrmFnlCod) As String

    Dim FnlData As FnlMstRec
    Dim sBufKey As String
    Dim sBufValue As String
    Dim iTmpSeq As Integer

    Dim sFnlMstCurKey As String
    Dim sFnlMstCmpKey As String
    Dim sPrmRetVal As String

    FnlData.FnlCod = sPrmFnlCod
    FnlMstStore sBufKey, sBufValue, FnlData
    sFnlMstCurKey = sBufKey

    sBufValue = mSetReadEqual("FnlMst", sFnlMstCurKey, sPrmRetVal)

    If sBufValue <> "" Then
    FnlMstLoad sPrmRetVal, FnlData
    If sPrmFnlCod = "MEDNUM" Or sPrmFnlCod = "INJNUM" Or sPrmFnlCod = "PHYNUM" Then
        Dim sBufDate As String
        sBufDate = SystemDate()
        If sBufDate = FnlData.FnlDte Then
        FnlData.FnlNum = CStr(CLong(FnlData.FnlNum) + 1)
        FnlData.FnlDte = sBufDate
        FinalNumberSettingOld = FnlData.FnlNum
        Else
        FnlData.FnlNum = "1"
        FnlData.FnlDte = sBufDate
        FinalNumberSettingOld = "1"
        End If
    Else
        FnlData.FnlNum = CStr(CLong(FnlData.FnlNum) + 1)
        FinalNumberSettingOld = FnlData.FnlNum
    End If
    Else
    If sPrmFnlCod = "MEDNUM" Or sPrmFnlCod = "INJNUM" Or sPrmFnlCod = "PHYNUM" Then
        FnlData.FnlNum = "1"
        FnlData.FnlDte = SystemDate()
        FnlData.FnlCod = sPrmFnlCod
        FinalNumberSettingOld = "1"
    Else
        FnlData.FnlNum = "1"
        FnlData.FnlCod = sPrmFnlCod
        FinalNumberSettingOld = FnlData.FnlNum
    End If
    End If

    FnlMstStore sBufKey, sBufValue, FnlData

    iTmpSeq = mWrite("FnlMst", sBufKey, sBufValue)
    If iTmpSeq = False Then
    iTmpSeq = mUpdate("FnlMst", sBufKey, sBufValue)
    End If

End Function

Public Sub FinalNumberUndo(sPrmFnlCod As String, sPrmFnlNum As String)

    Dim FnlData As FnlMstRec
    Dim sBufKey As String
    Dim sBufValue As String
    Dim iTmpSeq As Integer

    Dim sFnlMstCurKey As String
    Dim sFnlMstCmpKey As String
    Dim sPrmRetVal As String

    FnlData.FnlCod = sPrmFnlCod             ' CHTNUM, ETCNUM, OCMNUM ...etc
    FnlMstStore sBufKey, sBufValue, FnlData
    sFnlMstCurKey = sBufKey

    sBufValue = mSetReadEqual("FnlMst", sFnlMstCurKey, sPrmRetVal)

    '최종번호와 부여받은 번호가 같을때 만 최종번호를 -1 한다.
    If sBufValue <> "" Then
    
    FnlMstLoad sPrmRetVal, FnlData
    If FnlData.FnlNum = Trim(sPrmFnlNum) Then
        FnlData.FnlNum = CStr(CLong(FnlData.FnlNum) - 1)
        FnlMstStore sBufKey, sBufValue, FnlData

        iTmpSeq = mWrite("FnlMst", sBufKey, sBufValue)
        If iTmpSeq = False Then
        iTmpSeq = mUpdate("FnlMst", sBufKey, sBufValue)
        End If
    End If
    End If

End Sub

'Record,시작위치,항목값
Public Function funItemLoad1(ByVal sPrmBuf As String, sPrmCnt As Integer) As String
    
    Static sOldBuf As String
    'Static pfromOld As Integer
    Dim pto As Integer, Length As Integer ', pfrom As Integer

    If sPrmBuf = "" Then Exit Function
    
    If sPrmCnt <> 1 Then
    'pfrom = pfromOld + 1
    'pto = InStr(pfrom, sPrmBuf, Chr$(5))'"")  '찾은위치
    pto = InStr(1, sOldBuf, "")  '찾은위치
    
    Else
    'pfrom = 1
    'pto = InStr(pfrom, sPrmBuf, Chr$(5))'"")  '찾은위치
    sOldBuf = sPrmBuf
    pto = InStr(1, sOldBuf, "")  '찾은위치
    
    End If
    
    If pto = 0 Then
    funItemLoad1 = ""   'pfrom
    Exit Function
    End If

    'length = pto - 1'pfrom
    'funItemLoad1 = IIf(length > 0, Mid$(sPrmBuf, pfrom, length), "")   '항목값
    'funItemLoad1 = Mid$(sPrmBuf, pfrom, length)                         '항목값
    funItemLoad1 = Mid$(sOldBuf, 1, pto - 1)                        '항목값
    
    'pfromOld = pto             '다음 탐색시작위치
    sOldBuf = Right$(sOldBuf, Len(sOldBuf) - pto)

End Function

Public Function GetDetailItem(sPrmTabKey As String, sPrmDtlKey As String) As String
    Dim tDtlData As DtlMstRec
    Dim sDtlMstCurKey As String, sDtlMstCmpKey As String, sDtlMstRetVal As String

    sDtlMstCurKey = sPrmTabKey & Chr(5) & sPrmDtlKey & Chr(5)
    sDtlMstCurKey = mSetReadEqual("DtlMst", sDtlMstCurKey, sDtlMstRetVal)
    If sDtlMstCurKey <> "" Then
        DtlMstLoad sDtlMstRetVal, tDtlData
        GetDetailItem = RTrim(tDtlData.DtlCodNam)
    Else
        GetDetailItem = ""
    End If
End Function

Public Function GetHolidayNameAtMonth(sPrmDate As String, iPrmIsAddSunday As Integer, sPrmHolArr() As String) As Integer
    ' 이달에 있는 전체 공휴일을 가져온다.
    ' 년월일 (년월만 있어도 상관없다.), 일요일 포함, 공휴일 2차원 배열(1 ~ 31, 1 ~ 2) : 공휴일 갯수가 리턴
    
    Dim i As Integer, iHolCnt As Integer, iDayCnt As Integer, iIsDupe As Integer
    Dim sTmpYear As String, sTmpMonth As String, sTmpDate As String
    Dim sCurKey As String, sCmpKey As String, sRetVal As String
    Dim tHolMst As HolMstRec

    sTmpYear = Left(sPrmDate, 4)
    sTmpMonth = Mid(sPrmDate, 5, 2)

    iHolCnt = 0
    sCmpKey = sTmpYear & sTmpMonth
    sCurKey = sCmpKey
    sCurKey = mSetNext("HolMst", sCurKey)
    Do
        sCurKey = mReadNext("HolMst", sCurKey, sCmpKey, sRetVal)
        If sCurKey = "" Then Exit Do
    
        iHolCnt = iHolCnt + 1
        Call HolMstStore(sCurKey, sRetVal, tHolMst)
        sPrmHolArr(iHolCnt, 1) = tHolMst.HolDte
        sPrmHolArr(iHolCnt, 2) = tHolMst.HolDteNam
    Loop

    GetHolidayNameAtMonth = iHolCnt
    If Not iPrmIsAddSunday Then Exit Function

    iDayCnt = 1
    
    Do
        iIsDupe = False
        sTmpDate = sTmpYear & sTmpMonth & Format(iDayCnt, "0#")
        sTmpDate = Format(sTmpDate, "####/##/##")
        If Not IsDate(sTmpDate) Then Exit Do
        ' 일요일
        If Weekday(sTmpDate) = 1 Then
            sTmpDate = Pict2Data(sTmpDate, "9999/99/99")
            For i = 1 To iHolCnt
            If sTmpDate = sPrmHolArr(i, 1) Then
                iIsDupe = True
                Exit For
            End If
            Next
    
            If Not iIsDupe Then
            iHolCnt = iHolCnt + 1
            sPrmHolArr(iHolCnt, 1) = sTmpDate
            sPrmHolArr(iHolCnt, 2) = "일요일"
            End If
        End If
    
        iDayCnt = iDayCnt + 1
        If iDayCnt > 31 Then Exit Do
    Loop

    Call BubbleSort(iHolCnt, 2, 1, sPrmHolArr())
    GetHolidayNameAtMonth = iHolCnt

End Function

Public Function GetPopupItem(sPrmItem As String) As String

    Dim sTagStr As String
    
    sTagStr = mvbFrm.Tag
    If Not sTagStr = "" Then
    GetPopupItem = piece(sTagStr, Chr(5), 1)
    sPrmItem = RTrim(piece(sTagStr, Chr(5), 2))
    End If

    Unload mvbFrm

End Function

Public Sub GetSexAge(sResNum As String, sSex As String, sAge As String)
    
    Dim iAge As Integer
    Dim sBuf As String
    Dim sYear As String
    Dim sMonDay As String
    Dim iPlus As Integer
    Dim sTmpSex  As String

    If Trim(sResNum) = "" Or Len(sResNum) < 8 Then Exit Sub

    sAge = CStr(AgeCheck(sResNum, ""))
    
    sTmpSex = Mid(sResNum, 7, 1)
    
    sTmpSex = CLong(sTmpSex) Mod 2
    If CLong(sTmpSex) = 0 Then
        sSex = "F"
    Else
        sSex = "M"
    End If
    'If sTmpSex = "1" Or sTmpSex = "3" Or sTmpSex = "7" Then
    '    sSex = "M"
    'ElseIf sTmpSex = "2" Or sTmpSex = "4" Or sTmpSex = "8" Then
    '    sSex = "F"
    'End If

End Sub

'*******************************************
'   주야공휴로 시각을 입력하는 코드 check
'*******************************************
Public Function GrdDnhOk(sPrmAddCod As String) As Integer

    '수술, 마취
    If sPrmAddCod = "SUR" Or sPrmAddCod = "ANS" Or sPrmAddCod = "TRS" Or sPrmAddCod = "CAS" Or sPrmAddCod = "ETR" Then
        GrdDnhOk = True
    Else
        GrdDnhOk = False
    End If

End Function

'***************************************************************
'   방사선, 수술, 케스트등의 기타가산이 있는 코드여부 check
'***************************************************************
Public Function GrdEtcOk(sPrmAddCod As String) As Integer

    Select Case sPrmAddCod
    '방사선, 수술, 케스트, 한방경혈(하-1)
    Case "RAD", "SUR", "CAS", "HUL"
        GrdEtcOk = True
    Case Else
        GrdEtcOk = False
    End Select

End Function

Public Sub HangulOff()
    'Call cvtToEng(1)
End Sub

Public Sub HangulOn()
    'Call cvtToHan(1)
End Sub

Public Sub IcmInfChtRead(sChtNum As String, sTmpDte As String, tIcmData As IcmInfRec)

    Dim sCurKey As String, sCmpKey As String, sRetVal As String
    Dim tTmpData As IcmInfRec

    sCmpKey = Format(Trim(sChtNum), "@@@@@@@@") & Chr(5)
    sCurKey = sCmpKey & sTmpDte & "9999"
    sCurKey = mSetPrev("IcmInfChtDtm", sCurKey)
    Do
    sCurKey = mReadPrev("IcmInfChtDtm", sCurKey, sCmpKey, sRetVal)
    
    If sCurKey = "" Then Exit Do

    IcmInfLoad sRetVal, tTmpData
    If tTmpData.IcmAcpStt <> "IC" Then
        tIcmData = tTmpData
        Exit Sub
    End If
    Loop

End Sub

Public Sub IcmInfRead(sIcmNum As String, IcmData As IcmInfRec)
    Dim sIcmInfCurKey As String, sIcmInfRetVal As String
    Dim sIcmInfCmpKey As String
    
    sIcmInfCmpKey = Format(sIcmNum, "@@@@@@@@@@") & Chr(5)
    sIcmInfCurKey = mSetReadEqual("IcmInf", sIcmInfCmpKey, sIcmInfRetVal)
    IcmInfLoad sIcmInfRetVal, IcmData

End Sub

Public Function IsExistDtlMst(sPrmKey As String) As Integer
    Dim tDtlData As DtlMstRec
    Dim sBufValue As String

    sPrmKey = mSetReadNext("DtlMst", sPrmKey, piece(sPrmKey, Chr(5), 1), sBufValue)
    If sPrmKey = "" Then
    IsExistDtlMst = False
    Else
    DtlMstLoad sBufValue, tDtlData
    IsExistDtlMst = IIf(tDtlData.DtlTblCod = Left(sPrmKey, 6), True, False)
    End If
End Function

Public Function IsHoliday(sPrmDate As String) As Long
    
    ' 1 Sunday
    ' 7 SaturDay
    
    On Error Resume Next
    
    Dim sTmpDate As String, sTmpName As String
    Dim iTmpYear As Integer
    
    If ReturnHolidayName(sPrmDate, sTmpName) Then
        IsHoliday = 1
        Exit Function
    End If

    sTmpDate = Format(sPrmDate, "####/##/##")
    IsHoliday = CLong(Weekday(sTmpDate))
  
End Function

Public Function IsLeapyear(iPrmYear As Integer) As Integer
    IsLeapyear = False
    If (iPrmYear Mod 4 = 0) And (iPrmYear Mod 100 <> 0) Or (iPrmYear Mod 400 = 0) Then
    IsLeapyear = True
    End If
End Function

Public Function IsSex(sPrmResNum As String) As String

    Dim sTmpSexNum As String
    Dim sResNum    As String
    Dim iSex    As Integer
    
    sResNum = Pict2Data(sPrmResNum, "9999999999999")
    sTmpSexNum = Mid(sResNum, 7, 1)
    
    iSex = CLong(sTmpSexNum) Mod 2
    If iSex = 0 Then
        IsSex = "F"
    Else
        IsSex = "M"
    End If
    
End Function

Public Function ItrInfFinalData(sPrmOcmNum As String, sPrmItrData As ItrInfRec) As Integer

    Dim sItrInfCurKey As String, sItrInfCmpKey As String, sItrInfRetVal As String
    
    'ItrInf에 내역이 없으면

    ItrInfFinalData = False

    sItrInfCmpKey = sPrmOcmNum & Chr(5)
    sItrInfCurKey = sItrInfCmpKey & "999999999999"
    sItrInfCurKey = mSetPrev("ItrInf", sItrInfCurKey)
    sItrInfCurKey = mReadPrev("ItrInf", sItrInfCurKey, sItrInfCmpKey, sItrInfRetVal)
    If sItrInfCurKey <> "" Then
    ItrInfLoad sItrInfRetVal, sPrmItrData
    ItrInfFinalData = True
    End If

End Function

Public Sub ItrInfRead(sOcm As String, sDte As String, ItrData As ItrInfRec)

    Dim sCurKey  As String
    Dim sCmpKey As String
    Dim sRetVal As String
    
    sCmpKey = Format(sOcm, "@@@@@@@@@@") & Chr(5)
    sCurKey = sCmpKey & Pict2Data(sDte, "99999999")
    sCurKey = mSetPrev("ItrInf", sCurKey)
    sCurKey = mReadPrev("ItrInf", sCurKey, sCmpKey, sRetVal)
    Call ItrInfLoad(sRetVal, ItrData)
    
End Sub

Public Function LeftAlignData2Pict(ByVal sPrmBufStr As String, ByVal sPrmPicStr As String) As String

    Dim iPicLen As Integer, iBufLen As Integer, iTmpLen As Integer
    Dim sRetStr As String
    
    sRetStr = Data2Pict(sPrmBufStr, sPrmPicStr)

    iBufLen = LenK(sRetStr)
    iPicLen = LenK(sPrmPicStr)
    iTmpLen = Abs(iPicLen - iBufLen)
    
    LeftAlignData2Pict = Left(sRetStr & Space(iTmpLen), iPicLen)

End Function

Public Sub LetmeCentering(cPrmForm As Form)

    cPrmForm.Left = (Screen.Width - cPrmForm.Width) / 2
    cPrmForm.Top = (Screen.Height - cPrmForm.Height) / 2

End Sub

'*****************************
'   상병명 display routine
'*****************************
Public Sub LoadOicInf(sPrmOcmNum As String, tPrmOicData() As OicInfRec)
    
    Dim i As Integer
    Dim Index As Integer
    Dim sBufTxt As String
    Dim OicData As OicInfRec
    Dim sOicInfCurKey As String, sOicInfCmpKey As String, sOicInfRetVal As String
    
    ' tPrmOicData Clear
    For i = 1 To 10 '5(20개로 조정함.990818)
    tPrmOicData(i) = OicData
    Next i

    sOicInfCurKey = sPrmOcmNum & Chr(5)
    sOicInfCmpKey = sPrmOcmNum
        
    sOicInfCurKey = mSetNext("OicInf", sOicInfCurKey)
    
    For i = 1 To 10 '5(20개로 조정함.990818)
    sOicInfCurKey = mReadNext("OicInf", sOicInfCurKey, sOicInfCmpKey, sOicInfRetVal)
    If sOicInfCurKey = "" Then Exit For

    OicInfLoad sOicInfRetVal, tPrmOicData(i)
    Next i

End Sub

Public Function Master3Help(sPrmMstName As String, sPrmFindKey As String, sPrmCompKey As String, iPrmCodPos As Integer, iPrmDatPos1 As Integer, iPrmDatPos2 As Integer, sPrmDataVal As String) As String

    Dim sMstPara As String, sTmpMstKey As String
    Dim iErrCod As Integer

    '여기서 체크
    sTmpMstKey = sPrmFindKey & Chr(5)
    sTmpMstKey = mSetReadNext(sPrmMstName, sTmpMstKey, sPrmCompKey, sMstPara)
    If sTmpMstKey = "" Then
    sPrmDataVal = ""
    Master3Help = ""
    iErrCod = Message("HNTI007")
    Exit Function
    End If

    sMstPara = sPrmMstName & Chr(6)
    sMstPara = sMstPara & sPrmFindKey & Chr(5) & Chr(6)
    sMstPara = sMstPara & sPrmCompKey & Chr(6)
    sMstPara = sMstPara & iPrmCodPos & Chr(6)
    sMstPara = sMstPara & iPrmDatPos1 & Chr(6)
    sMstPara = sMstPara & iPrmDatPos2
    
    'Mvb3Frm.Tag = sMstPara
    'Mvb3Frm.Show 1
    
    'Master3Help = GetPopupItem(sPrmDataVal)
    
    'Master3Help = Trim(Piece(Mvb3Frm.Tag, Chr(5), 1))
    
    'If Piece(Mvb3Frm.Tag, Chr(5), 2) = "" Then
    '    sPrmDataVal = Piece(Mvb3Frm.Tag, Chr(5), 3)
    'Else
    '    sPrmDataVal = Piece(Mvb3Frm.Tag, Chr(5), 2)
    'End If

End Function

Public Function MasterHelp(sPrmMstName As String, sPrmFindKey As String, sPrmCompKey As String, iPrmCodePos As Integer, iPrmDataPos As Integer, sPrmDataVal As String) As String

    Dim sMstPara As String, sTmpMstKey As String
    Dim iErrCod As Integer

    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '2001/10/29 james
    'VB6.0에서는 인덱스로 한번에 찾을 수 없다...
    '그랴서 두번 M Routin을 쓴다...
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '여기서 체크
    sTmpMstKey = sPrmFindKey & Chr(5)
    'sTmpMstKey = mSetReadNext(sPrmMstName, sTmpMstKey, sPrmCompKey, sMstPara)
    sTmpMstKey = mSetNext(sPrmMstName, sTmpMstKey)
    sTmpMstKey = mReadNext(sPrmMstName, sTmpMstKey, sPrmCompKey, sMstPara)
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    If sTmpMstKey = "" Then
        sPrmDataVal = ""
        MasterHelp = ""
        iErrCod = Message("HNTI007")
        Exit Function
    End If
    
    sMstPara = sPrmMstName & Chr(6)
   'sMstPara = sMstPara & sPrmFindKey & Chr(5) & Chr(6)
    sMstPara = sMstPara & sPrmFindKey & Chr(6)
    sMstPara = sMstPara & sPrmCompKey & Chr(6)
    sMstPara = sMstPara & iPrmCodePos & Chr(6)
    sMstPara = sMstPara & iPrmDataPos

    mvbFrm.Tag = sMstPara
    mvbFrm.ZOrder 0
    mvbFrm.Show 1
    
    MasterHelp = GetPopupItem(sPrmDataVal)

End Function

Public Function MasterHelpDetail(sPrmMstName As String, sPrmFindKey As String, sPrmCompKey As String, iPrmDataPos As Integer) As String

    Dim sBufCurKey As String, sBufRetVal As String

    sBufCurKey = sPrmFindKey
    sBufCurKey = mSetReadEqual(sPrmMstName, sBufCurKey, sBufRetVal)

    If (sBufCurKey = "") And (Not sPrmCompKey = "") Then
        sBufCurKey = sPrmFindKey
        sBufCurKey = mSetReadNext(sPrmMstName, sBufCurKey, sPrmCompKey, sBufRetVal)
        If sBufCurKey = "" Then
            MasterHelpDetail = ""
        Else
            sPrmFindKey = sBufCurKey
            MasterHelpDetail = piece(sBufRetVal, Chr(5), iPrmDataPos)
        End If
    Else
        sPrmFindKey = sBufCurKey
        MasterHelpDetail = piece(sBufRetVal, Chr(5), iPrmDataPos)
    End If

End Function

Public Function MaxMgdMst() As Integer

    Dim sMgdCurKey As String, sMgdCmpKey As String, sMgdRetVal As String
    Dim iCnt As Integer

    iCnt = 0
    sMgdCmpKey = ""
    sMgdCurKey = sMgdCmpKey & Chr(5)

    sMgdCurKey = mSetNext("MgdMst", sMgdCurKey)
    Do
    sMgdCurKey = mReadNext("MgdMst", sMgdCurKey, sMgdCmpKey, sMgdRetVal)
    If sMgdCurKey = "" Then Exit Do

    iCnt = iCnt + 1
    Loop

    MaxMgdMst = iCnt

End Function
'MsgMst에 등록하지 않고 메시지를 사용하고 싶은 경우에 사용한다.
'sPrmMsg가 있는 경우는 MsgMst를 읽지 않는다.
'                      sPrmMsgCode를 4자리만 입력한다. 예) HNTQ, HNTI, HNTW ...
'                      Message는 sPrmMsg를 표시하고, 코드 유형은 sPrmMsgCode를 따른다.
Public Function Message(sPrmMsgCode As String, Optional sPrmMsg As String = "") As Integer

    Dim sTmpStr As String
    Dim tMgdData As MsgMstRec
    Dim sCode As String, sMsg As String

    If sPrmMsg = "" Then
        If InStr(sPrmMsgCode, Chr(5)) <> 0 Then
            sCode = piece(sPrmMsgCode, Chr(5), 1)
            sMsg = piece(sPrmMsgCode, Chr(5), 2)
        Else
            sCode = sPrmMsgCode
            sMsg = ""
        End If

        Message = vbYes
        sPrmMsgCode = mSetReadEqual("MsgMst", sCode & Chr(5), sTmpStr)
        If sPrmMsgCode = "" Then Exit Function

        MsgMstLoad sTmpStr, tMgdData
    Else
        tMgdData.MsgCod = sPrmMsgCode
        tMgdData.MsgCodNam = sPrmMsg
    End If

    Message = DisplayMsgBox(tMgdData, sMsg)

End Function

Public Function MessageBoxNew(sPrmMsgCode As String, iPrmFlag As Integer) As Integer
    
    Dim sTmpStr As String
    Dim tMgdData As MsgMstRec
    Dim sCode As String, sMsg As String

    If InStr(sPrmMsgCode, Chr(5)) <> 0 Then
    sCode = piece(sPrmMsgCode, Chr(5), 1)
    sMsg = piece(sPrmMsgCode, Chr(5), 2)
    Else
    sCode = sPrmMsgCode
    sMsg = ""
    End If
    MessageBoxNew = vbOK
    sPrmMsgCode = mSetReadEqual("MsgMst", sCode & Chr(5), sTmpStr)
    If sPrmMsgCode = "" Then Exit Function

    MsgMstLoad sTmpStr, tMgdData
    MessageBoxNew = DisplayMessageBox(tMgdData, sMsg, iPrmFlag)

End Function

Public Function NewAddCentury(sOldDate As String) As String
    
    If IsDate(Format(sOldDate, "@@-@@-@@")) = False Then
    NewAddCentury = ""
    Exit Function
    Else
    sOldDate = Format(Right$(sOldDate, 4) & Left$(sOldDate, 2), "@@/@@/@@")
    End If

    mvbFrm.Mvb1.P0 = sOldDate
    mvbFrm.Mvb1.P1 = ""

    mvbFrm.Mvb1.Code = "d ^AddCentury(P0,.P1)"
    mvbFrm.Mvb1.ExecFlag = 1

    NewAddCentury = mvbFrm.Mvb1.P1

End Function

Public Sub OcmInfRead(sOcmNum As String, OcmData As OcmInfRec)
    
    Dim sOcmInfCurKey As String, sOcmInfRetVal As String
    Dim sOcmInfCmpKey As String
    
    sOcmInfCmpKey = Format(sOcmNum, "@@@@@@@@@@") & Chr(5)
    sOcmInfCurKey = mSetReadEqual("OcmInf", sOcmInfCmpKey, sOcmInfRetVal)
    OcmInfLoad sOcmInfRetVal, OcmData

End Sub

Public Sub OrpInfRead(sOcm As String, sFg As String, OrpData As OrpInfRec)

    Dim sOrpInfCurKey As String, sOrpInfRetVal As String
    
    sOrpInfCurKey = Format(sOcm, "@@@@@@@@@@") & Chr(5) & sFg & Chr(5)
    sOrpInfCurKey = mSetReadEqual("OrpInf", sOrpInfCurKey, sOrpInfRetVal)
    OrpInfLoad sOrpInfRetVal, OrpData

End Sub

Public Sub Outlines(formname As Form)

    Dim drkgray As Long, fullwhite As Long
    Dim i As Integer
    Dim ctop As Integer, cleft As Integer, cright As Integer, cbottom As Integer

    ' Outline a form's controls for 3D look unless control's TAG
    ' property is set to "skip".

    Dim cName As Control
    drkgray = RGB(128, 128, 128)
    fullwhite = RGB(255, 255, 255)

    For i = 0 To (formname.Controls.Count - 1)
    Set cName = formname.Controls(i)
    If TypeOf cName Is Menu Then

    ElseIf (UCase(cName.Tag) = "OL") Then
        ctop = cName.Top - Screen.TwipsPerPixelY
        cleft = cName.Left - Screen.TwipsPerPixelX
        cright = cName.Left + cName.Width
        cbottom = cName.Top + cName.Height
        formname.Line (cleft, ctop)-(cright, ctop), drkgray
        formname.Line (cleft, ctop)-(cleft, cbottom), drkgray
        formname.Line (cleft, cbottom)-(cright, cbottom), fullwhite
        formname.Line (cright, ctop)-(cright, cbottom), fullwhite
    End If
    Next i
End Sub

'**************************
'   개인 인적 사항 읽기
'**************************
Public Sub PbsInfRead(sChart As String, PbsData As PbsInfRec)
    
    Dim sPbsInfCurKey As String, sPbsInfRetVal As String
    Dim tTmpPbsData As PbsInfRec

    PbsData.PbsChtNum = Format(CDouble(sChart), "@@@@@@@@")
    'PbsData.PbsChtNum = Format(RTrim(sChart), "@@@@@@@@")
    sPbsInfCurKey = PbsData.PbsChtNum & Chr(5)
    'PbsInfStore sPbsInfCurKey, sPbsInfRetVal, PbsData

    sPbsInfCurKey = mSetReadEqual("PbsInf", sPbsInfCurKey, sPbsInfRetVal)
    'If sPbsInfCurKey <> "" Then
    PbsInfLoad sPbsInfRetVal, PbsData
    'Else
    '    PbsData = tTmpPbsData
    'End If

End Sub

Public Sub PicOutlines(pic As Control, Ctl As Control)

    Dim drkgray As Long, fullwhite As Long
    Dim ctop As Integer, cleft As Integer, cright As Integer, cbottom As Integer

    ' Outline a form's controls for 3D look unless control's TAG
    ' property is set to "skip".

    Dim cName As Control
    drkgray = RGB(128, 128, 128)
    fullwhite = RGB(255, 255, 255)

    ctop = Ctl.Top - Screen.TwipsPerPixelY
    cleft = Ctl.Left - Screen.TwipsPerPixelX
    cright = Ctl.Left + Ctl.Width
    cbottom = Ctl.Top + Ctl.Height
    pic.Line (cleft, ctop)-(cright, ctop), drkgray
    pic.Line (cleft, ctop)-(cleft, cbottom), drkgray
    pic.Line (cleft, cbottom)-(cright, cbottom), fullwhite
    pic.Line (cright, ctop)-(cright, cbottom), fullwhite
End Sub

Public Function Pict2Data(sPrmData As String, sPrmPict As String) As String

    Dim i As Integer, iPictPos As Integer
    Dim iDataLen As Integer, iPictLen As Integer
    Dim sBufData As String, sPictStr As String, sChar As String

    iDataLen = Len(sPrmData)
    iPictLen = Len(sPrmPict)
    iPictPos = 1
    sBufData = ""
    
    For i = 1 To iPictLen
    sPictStr = ""

    Select Case Mid(sPrmPict, i, 1)
    Case "X"
        sPictStr = Mid(sPrmData, iPictPos, 1)
        iPictPos = iPictPos + 1

    Case "9"
        sPictStr = Mid(sPrmData, iPictPos, 1)
        If Not IsNumeric(sPictStr) Then
        sPictStr = ""
        i = i - 1
        End If
        iPictPos = iPictPos + 1

    End Select

    sBufData = sBufData & sPictStr

    If iPictPos > iDataLen Then
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

    Pict2Data = sBufData

End Function

Public Function pieceNew(ByVal sPrmBuf As String, Delimeter As String, Cnt As Integer) As String

    Dim sVal() As String
    
    sVal = Split(sPrmBuf, Delimeter)
    pieceNew = sVal(Cnt - 1)

End Function

Public Function piece(ByVal sPrmBuf As String, Delimeter As String, Cnt As Integer) As String

    Dim i As Integer, Length As Integer, pto As Integer, pfrom As Integer

    Dim Hit As Integer
    Dim sCurBuf As String
    Static OldBuf As String
    Static OldCnt As Integer
    Static OldPFrom As Integer
    Static OldPTo As Integer

    piece = ""

    sCurBuf = sPrmBuf
    sCurBuf = sCurBuf & Replicate(Delimeter, 10)

    If (OldBuf = sCurBuf) Then
        Hit = True
    Else
        Hit = False
    End If
    
    If ((Hit = False) Or (OldCnt > Cnt)) Then
        pto = 1 - Len(Delimeter)  ' ????????????????????
        For i = 1 To Cnt
            pfrom = pto + Len(Delimeter)
            pto = InStr(pfrom, sCurBuf, Delimeter)
        Next
    Else
        pto = OldPTo
        pfrom = OldPFrom
        For i = 1 To Cnt - OldCnt
            pfrom = pto + Len(Delimeter)
            pto = InStr(pfrom, sCurBuf, Delimeter)
        Next
    End If

    OldCnt = Cnt

    If pto = 0 Then Exit Function
    
    Length = pto - pfrom
    piece = Mid$(sCurBuf, pfrom, Length)

    If (Hit = False) Then
        OldBuf = sCurBuf
    End If
    
    OldPFrom = pfrom
    OldPTo = pto

End Function

Public Function Piece1(ByVal buffer As String, Delimeter As String, Cnt As Integer) As String
    
    Dim Index As Integer
    Dim ePos As Long
    Dim Length As Long
    Dim Value As String
    Dim sValue As String
    
    If buffer = "" Then
    Piece1 = ""
    Exit Function
    End If

    For Index = 1 To Cnt
    
    Length = Len(buffer)
    
    ePos = InStr(1, buffer, Delimeter)
    Value = ""
    If (ePos = 0) Or (Length <= 0) Then
        Value = buffer
        Exit For
    Else
        Value = Left(buffer, ePos - 1)
    End If
    
    If Length < ePos Then
        buffer = ""
    Else
        buffer = Right(buffer, Length - ePos)
    End If
    Next Index

    Piece1 = Value
End Function

Public Function Piece2(ByVal sPrmBuf As String, Delimeter As String, Cnt As Integer)
    
    ' Piece=>?조각?
    ' sPrmBuf = 날짜들의 모음, Delimeter = Chr(5), cnt = 1 - 31(Value i)
    Dim sBufValue As String
    Dim i As Integer, Length As Integer, pto As Integer, pfrom As Integer

    pto = 1 - Len(Delimeter)  ' ????????????????????
    sBufValue = sPrmBuf & Replicate(Delimeter, 10)
    Piece2 = ""
    For i = 1 To Cnt
    pfrom = pto + Len(Delimeter)
    pto = InStr(pfrom, sBufValue, Delimeter)
    Next
    If pto = 0 Then Exit Function
    Length = pto - pfrom
    Piece2 = Mid$(sBufValue, pfrom, Length)

End Function

Public Function PieceChange(sPrmSrc As String, sPrmSepStr As String, iPrmSepStrPos As Integer, sPrmRepDst As String) As String

    Dim i As Integer
    Dim sRetVal As String, sBufStr As String
    Dim iSrcStrLen As Integer, iSepStrLen As Integer, iSepStrCnt As Integer, iSepStrPos As Integer

    iSrcStrLen = Len(sPrmSrc)
    iSepStrLen = Len(sPrmSepStr)

    iSepStrCnt = 0
    sBufStr = sPrmSrc
    For i = 1 To iSrcStrLen
    iSepStrPos = InStr(sBufStr, sPrmSepStr)
    If iSepStrPos = 0 Then
        If i = 1 Then
        PieceChange = sPrmSrc
        Exit Function
        Else
        Exit For
        End If
    End If
    iSepStrCnt = iSepStrCnt + 1
    sBufStr = Mid(sBufStr, iSepStrPos + iSepStrLen)
    Next
    iSepStrCnt = iSepStrCnt + 1

    sRetVal = ""
    For i = 1 To iSepStrCnt
    sBufStr = piece(sPrmSrc, sPrmSepStr, i)
    If i = iPrmSepStrPos Then
        sBufStr = sPrmRepDst
    End If
    sRetVal = sRetVal & sBufStr & sPrmSepStr
    Next

    PieceChange = Left(sRetVal, Len(sRetVal) - iSepStrLen)

End Function

Public Function PopupItemList(sPrmItemKey As String, sPrmRetKey As String, sPrmRetVal As String) As Integer

    If Not IsExistDtlMst(sPrmItemKey) Then
    PopupItemList = False
    Exit Function
    End If
    Load mvbFrm
    mvbFrm.Tag = sPrmItemKey
    mvbFrm.Show 1

    sPrmRetKey = GetPopupItem(sPrmRetVal)
    PopupItemList = True

End Function

'---------------------------------------------------------------------------
'   하나의 EXE화일이 하나의 PC에서 두개이상 실행되지 않게..
'       1996년 6월 12일
'   공통함수내에서 "End"문을 사용하지 않기 위해 Function으로 내용을 바꾼다.
'        "End" 사용시 DLL에서 사용할 수 없음
'---------------------------------------------------------------------------
'Sub PrevInstanceCheck()
'
'    Dim iTmp As Integer
'
'    If App.PrevInstance Then
'    '메시지마스터에 심어 놨다가 혹시 병원마다 틀려질것을 의심해서 ...
'        Call MsgBox("이미 프로그램이 실행중입니다!")
'        End
'    End If
'
'End Sub
Public Function PrevInstanceCheck() As Boolean

    Dim iTmp As Integer

    If App.PrevInstance Then
    '메시지마스터에 심어 놨다가 혹시 병원마다 틀려질것을 의심해서 ...
        Call MsgBox("이미 프로그램이 실행중입니다!")
        PrevInstanceCheck = True
    Else
        PrevInstanceCheck = False
    End If

End Function


Public Sub PspInfRead(sChart As String, PspData As PspInfRec)
    Dim sPspInfCurKey As String, sPspInfRetVal As String
    Dim tTmpPspData As PspInfRec

    PspData.PspChtNum = Format(CDouble(sChart), "@@@@@@@@")
    sPspInfCurKey = PspData.PspChtNum & Chr(5)
    'PspInfStore sPspInfCurKey, sPspInfRetVal, PspData

    sPspInfCurKey = mSetReadEqual("PspInf", sPspInfCurKey, sPspInfRetVal)
    'If sPspInfCurKey <> "" Then
    PspInfLoad sPspInfRetVal, PspData
    'Else
    '    PspData = tTmpPspData
    'End If

End Sub

Public Function Replicate(ByVal sPrmChr As String, ByVal iPrmLen As Integer) As String

    Dim i As Integer
    Dim sRetBuf As String

    sRetBuf = ""
    For i = 1 To iPrmLen
        sRetBuf = sRetBuf & sPrmChr
    Next

    Replicate = sRetBuf

End Function

Public Function ResNumValidCheck(sPrmResNum As String)

    Dim i As Integer
    Dim iSum As Long
    Dim sTmp As String
    Dim iVal As Integer

    If LenK(sPrmResNum) <> 13 Then
        ResNumValidCheck = False
        Exit Function
    End If

    For i = 1 To 8
        sTmp = MidK(sPrmResNum, i, 1)
        iSum = iSum + CInteger(sTmp) * (i + 1)
    Next

    For i = 1 To 4
        sTmp = MidK(sPrmResNum, i + 8, 1)
        iSum = iSum + CLong(sTmp) * (i + 1)
    Next
    
    'iVal = iSum - CInt(iSum / 11) * 11

    '981222민경
    iVal = iSum Mod 11
    
'    If iVal < 2 Then
'        'iVal = iVal + 11
'        '981222민경(나머지가 1과0인 경우만 값을 10으로 Setting한다.)
'        iVal = 10
'    End If
'
'    If 11 - iVal <> CInteger(RightK(sPrmResNum, 1)) Then
'        ResNumValidCheck = False
'    Else
'        ResNumValidCheck = True
'    End If

    '추가
    Dim ChkDigit As Integer
    
    '''이부분을 막고 새롭게 넣어주세요.
    'If iVal < 2 Then
    ' 'iVal = iVal + 11
    ' '981222민경(나머지가 1과0인 경우만 값을 10으로 Setting한다.)
    ' iVal = 10
    'End If
    
    If iVal = 0 Then
        iVal = 10
    ElseIf iVal = 1 Then
        iVal = 11
    End If
    
    ChkDigit = 11 - iVal
    
    If ChkDigit <> CInteger(RightK(sPrmResNum, 1)) Then
        ResNumValidCheck = False
    Else
        ResNumValidCheck = True
    End If
End Function

Public Function ReturnHolidayName(sPrmDate As String, sPrmName As String) As Integer

    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sRetVal As String
    
    Dim tHolMst As HolMstRec

    sCmpKey = sPrmDate & Chr(5)
    sCurKey = sCmpKey
    sCurKey = mSetNext("HolMst", sCurKey)
    sCurKey = mReadNext("HolMst", sCurKey, sCmpKey, sRetVal)
    If sCurKey <> "" Then
        Call HolMstLoad(sRetVal, tHolMst)
        sPrmName = tHolMst.HolDteNam
        ReturnHolidayName = True
    Else
        sPrmName = ""
        ReturnHolidayName = False
    End If

End Function

Public Function RightAlignData2Pict(ByVal sPrmBufStr As String, ByVal sPrmPicStr As String) As String

    Dim iPicLen As Integer, iBufLen As Integer, iTmpLen As Integer
    Dim sRetStr As String
    
    sRetStr = Data2Pict(sPrmBufStr, sPrmPicStr)

    iBufLen = LenK(sRetStr)
    iPicLen = LenK(sPrmPicStr)
    iTmpLen = Abs(iPicLen - iBufLen)
    
    RightAlignData2Pict = RightK(Space(iTmpLen) & sRetStr, iPicLen)

End Function

Public Function SystemDate() As String

    mvbFrm.Mvb1.P0 = ""

    mvbFrm.Mvb1.Code = "d ^SystemDate(.P0)"
    mvbFrm.Mvb1.ExecFlag = 1

    SystemDate = mvbFrm.Mvb1.P0

End Function

Public Function SystemLongDate() As String

    'MvbFrm.MVB1.P0 = ""

    'MvbFrm.MVB1.Code = "d ^SystemLongDate(.P0)"
    'MvbFrm.MVB1.ExecFlag = 1

    'SystemLongDate = MvbFrm.MVB1.P0
    SystemLongDate = AddCentury(SystemDate())

End Function

Public Function SystemTime() As String
    
    mvbFrm.Mvb1.P0 = ""

    mvbFrm.Mvb1.Code = "d ^SystemTime(.P0)"
    mvbFrm.Mvb1.ExecFlag = 1

    SystemTime = mvbFrm.Mvb1.P0

End Function

Public Function TimeValidCheck(sPrmTime As String) As Integer
    
    Dim iTmpHour As Integer
    Dim iTmpMin As Integer

    If (Not IsNumeric(sPrmTime)) Or (Not Len(sPrmTime) = 4) Then
        TimeValidCheck = False
        Exit Function
    End If

    iTmpHour = CInteger(Left(sPrmTime, 2))
    iTmpMin = CInteger(Right(sPrmTime, 2))

    '종료일에 사용하는 값은 그냥 True로
    If sPrmTime <> "9999" Then
    '24시 까지 허용해서는 않된다!..1997년 11월 12일
    'If iTmpHour < 0 Or iTmpHour > 24 Or iTmpMin < 0 Or iTmpMin > 59 Then
    If iTmpHour < 0 Or iTmpHour > 23 Or iTmpMin < 0 Or iTmpMin > 59 Then
        TimeValidCheck = False
        Exit Function
    End If
    End If

    TimeValidCheck = True
    
End Function

Public Function ToJulian(lPrmYear As Long, lPrmMonth As Long, lPrmDay As Long) As Long

On Error GoTo ToJulianErrorTrap

    Dim i As Integer
    Dim lTmpMonthSum() As Long
    Dim lTmpTotal As Long, lTmpYear As Long

    ReDim lTmpMonthSum(0 To 12)

    lTmpMonthSum(0) = 0
    lTmpMonthSum(1) = 31
    lTmpMonthSum(2) = 59
    lTmpMonthSum(3) = 90
    lTmpMonthSum(4) = 120
    lTmpMonthSum(5) = 151
    lTmpMonthSum(6) = 181
    lTmpMonthSum(7) = 212
    lTmpMonthSum(8) = 243
    lTmpMonthSum(9) = 273
    lTmpMonthSum(10) = 304
    lTmpMonthSum(11) = 334
    lTmpMonthSum(12) = 365

    lTmpYear = CLong(lPrmYear) - 1
    lTmpTotal = lTmpYear * 365
    lTmpTotal = lTmpTotal + (lTmpYear \ 4)
    lTmpTotal = lTmpTotal + (lTmpYear \ 400)
    lTmpTotal = lTmpTotal - (lTmpYear \ 100)

    lTmpTotal = lTmpTotal + lTmpMonthSum(lPrmMonth - 1) + lPrmDay

    If lPrmMonth > 2 Then
    If IsLeapyear(CInteger(lPrmYear)) Then
        lTmpTotal = lTmpTotal + 1
    End If
    End If
    
    ToJulian = lTmpTotal

    Exit Function

ToJulianErrorTrap:
    Resume Next

End Function

Public Function Translate(sPrmSrc As String, sPrmRepSrc As String, iPrmRepSrcPos As Integer, sPrmRepDst As String) As String

    Dim i As Integer
    Dim iSrcStrPos As Integer, iSrcStrLen As Integer
    Dim iRepSrcStrLen As Integer, iOldStrLen As Integer, iBufStrLen As Integer
    Dim sBufStr As String, sRetVal As String

    Translate = ""
    iSrcStrLen = Len(sPrmSrc)
    iRepSrcStrLen = Len(sPrmRepSrc)
    
    iOldStrLen = 0
    sBufStr = sPrmSrc
    iBufStrLen = Len(sBufStr)
    For i = 1 To iSrcStrLen
    iSrcStrPos = InStr(sBufStr, sPrmRepSrc)
    If (i = 1) And (iSrcStrPos = 0) Then
        Exit Function
    ElseIf (iPrmRepSrcPos = i) And (Not iSrcStrPos = 0) Then
        iOldStrLen = iOldStrLen + iSrcStrPos
        Exit For
    End If
    sBufStr = Right(sBufStr, iBufStrLen - (iSrcStrPos + iRepSrcStrLen) + 1)
    iOldStrLen = iOldStrLen + (iBufStrLen - Len(sBufStr))
    iBufStrLen = Len(sBufStr)
    Next
    iSrcStrPos = iOldStrLen

    sRetVal = Left(sPrmSrc, iSrcStrPos - 1)
    Translate = sRetVal & sPrmRepDst & Right(sPrmSrc, iSrcStrLen - (iSrcStrPos + iRepSrcStrLen) + 1)
    
End Function


'**************************************************
'   ItmMst 화일 읽기
'   History 관리 로직이 추가되면서 sPrmDte가 추가됨
'**************************************************
Public Sub ItmMstRead(sCode As String, ItmData As ItmMstRec, Optional ByVal sPrmDte As String)

    Dim tItmHst As ItmHstRec
    Dim sCurKey As String, sCmpKey As String, sRetVal As String
    
    If sPrmDte = "" Then sPrmDte = SystemLongDate()
    
    sCurKey = sCode
    If Right(sCode, 1) <> Chr(5) Then
        sCurKey = sCode & Chr(5)
    End If
    sCurKey = mSetReadEqual("ItmMst", sCurKey, sRetVal)
    
    ItmMstLoad sRetVal, ItmData
    
    'If ItmData.ItmAdpDte <= sPrmDte And ItmData.ItmExpDte >= sPrmDte Then
    '    Exit Sub
    'End If
    
    'sCmpKey = sCode & Chr(5)
    'sCurKey = sCmpKey & sPrmDte & Chr(5)
    'sCurKey = mSetReadPrev("ItmHst", sCurKey, sCmpKey, sRetVal)

    'If sCurKey = "" Then
    '    sRetVal = ""
    '    ItmMstLoad sRetVal, ItmData
    'Else
    '    ItmHstLoad sRetVal, tItmHst
    '    If tItmHst.ItmAdpDte <= sPrmDte And tItmHst.ItmExpDte >= sPrmDte Then
    '        ItmHstStore sCurKey, sRetVal, tItmHst
    '        sRetVal = sCode & Chr(5) & sRetVal
    '        ItmMstLoad sRetVal, ItmData
    '    Else
    '        sRetVal = ""
    '        ItmMstLoad sRetVal, ItmData
    '    End If
    'End If
    
End Sub

Public Sub BedMstIcmRead2Wrd(sIcmNum As String, WrdMst As WrdMstRec)

    Dim sCurKey As String, sCmpKey As String, sRetVal As String
    Dim BedData As BedMstRec
    Dim WrdData As WrdMstRec
    Dim sWrdCurKey As String, sWrdCmpKey As String, sWrdRetVal As String
    
    sCmpKey = sIcmNum & Chr(5)
    sCurKey = sCmpKey
    
    sCurKey = mSetNext("BedMstOcm", sCurKey)
    Do
        sCurKey = mReadNext("BedMstOcm", sCurKey, sCmpKey, sRetVal)
    
        If sCurKey = "" Then Exit Sub
        
        Call BedMstLoad(sRetVal, BedData)
            
        If Trim(sIcmNum) = Trim(BedData.BedOcmNum) Then
            sWrdCurKey = BedData.BedWrdCod & Chr(5)
            sWrdCurKey = mSetReadEqual("WrdMst", sWrdCurKey, sWrdRetVal)
            Call WrdMstLoad(sWrdRetVal, WrdData)
            WrdMst = WrdData
            
            DoEvents
            Exit Sub
        End If
    Loop
    
End Sub

'**************************************************
'   담당자 화일 읽기
'   History 관리 로직이 추가되면서 sPrmDte가 추가됨
'**************************************************
Public Sub UidMstRead(sCode As String, UidData As UidMstRec, Optional ByVal sPrmDte As String)

    Dim tUidHst As UidHstRec
    Dim sCurKey As String, sCmpKey As String, sRetVal As String
    
    ''If IsMissing(sPrmDte) Then sPrmDte = SystemLongDate()
    If sPrmDte = "" Then sPrmDte = SystemLongDate()
    
    sCurKey = sCode & Chr(5)
    sCurKey = mSetReadEqual("UidMst", sCurKey, sRetVal)
    
    UidMstLoad sRetVal, UidData
    
    tGblUidMst = UidData
    ''If UidData.UidAdpDte >= sPrmDte And UidData.UidEndDte <= sPrmDte Then
    If UidData.UidAdpDte <= sPrmDte And UidData.UidEndDte >= sPrmDte Then
        Exit Sub
    End If
    
    sCmpKey = sCode & Chr(5)
    sCurKey = sCmpKey & sPrmDte & Chr(5)
    sCurKey = mSetPrev("UidHst", sCurKey)
    sCurKey = mReadPrev("UidHst", sCurKey, sCmpKey, sRetVal)

    If sCurKey = "" Then
        sRetVal = ""
    Else
        UidHstLoad sRetVal, tUidHst
        If tUidHst.UidAdpDte <= sPrmDte And tUidHst.UidEndDte >= sPrmDte Then
            UidHstStore sCurKey, sRetVal, tUidHst
            sRetVal = sCode & Chr(5) & sRetVal
            UidMstLoad sRetVal, UidData
        Else
            sRetVal = ""
            'UidMstLoad sRetVal, UidData
        End If
    End If
    
End Sub

Public Sub UsgMstRead(sUsgCod As String, UsgData As UsgMstRec)

    Dim sCurKey As String, sRetVal As String
    
    sCurKey = Trim(sUsgCod) & Chr(5)
    sCurKey = mSetReadEqual("UsgMst", sCurKey, sRetVal)
    Call UsgMstLoad(sRetVal, UsgData)

End Sub


'Dll에서 End를 사용할 수 없어서 Function 형태로 변경한다.
'Public Sub UnloadOfProgram(iPrmFlag As Integer)
'
'    Dim iTmp As Integer
'
'    iTmp = Message("HNTQ003")
'    If iTmp = vbYes Then
'        End
'    Else
'        iPrmFlag = True
'    End If
'
'End Sub

Public Function UnLockChtNum(sChtNum As String, sLevCod As String, sExeName As String, sUidCod As String, sIPAddr As String) As Integer

    '챠트번호 Locking을 지운다.
    Dim tLocData As LocChtRec
    Dim sCmpKey  As String, sRetVal As String
    Dim i As Integer

    UnLockChtNum = False
    sCmpKey = sChtNum & Chr(5) & sLevCod & Chr(5)
    If mSetReadEqual("LocCht", sCmpKey, sRetVal) <> "" Then
    LocChtLoad sRetVal, tLocData
    If CStr(sExeName) = tLocData.LocExeNam And sIPAddr = tLocData.LocIpAddr Then
        sCmpKey = sChtNum & Chr(5) & sLevCod & Chr(5)
        i = mDelete("LocCht", sCmpKey)
        UnLockChtNum = True
    End If
    End If

End Function

Public Function LockingChtNum(sChtNum As String, sLevCod As String, sExeName As String, sUidCod As String, sIPAddr As String, Optional pbDisplayMsg As Boolean = True) As Integer
    '챠트번호,실행파일명,화일Level,UidCod,IPAddress
    Dim tLocData As LocChtRec
    Dim sCurKey As String, sCmpKey As String, sRetVal As String
    Dim sPrtCod As String, sMsg As String, i As Integer
    
    Call DeleteLocChtUid(sUidCod)
    '------------------------------------------------------------------------------------------------
    sCmpKey = sChtNum & Chr(5) & sLevCod & Chr(5)
    If mSetReadEqual("LocCht", sCmpKey, sRetVal) <> "" Then
        Call LocChtLoad(sRetVal, tLocData)
        sMsg = MasterHelpDetail("PbsInf", tLocData.LocChtNum & Chr(5), tLocData.LocChtNum & Chr(5), 2) & "님은 "
        sMsg = sMsg & MasterHelpDetail("UidMst", tLocData.LocUidCod & Chr(5), tLocData.LocUidCod & Chr(5), 2) & "씨가 "
        '사용부서
        sPrtCod = MasterHelpDetail("UidMst", tLocData.LocUidCod & Chr(5), tLocData.LocUidCod & Chr(5), 4)
        sMsg = sMsg & MasterHelpDetail("DtlMst", "UDPTBL" & Chr(5) & sPrtCod & Chr(5), "UDPTBL" & Chr(5) & sPrtCod & Chr(5), 3) & "에서 뛰워놓은 상태입니다. 조회모드로 전환합니다."
        sMsg = sMsg & vbCrLf & "조회모드 상태에서도 결과조회나 처방조회, 챠트조회는 가능합니다."
        sMsg = sMsg & "  IPAddress ☞ " & tLocData.LocIpAddr
        If pbDisplayMsg Then
            MsgBox sMsg
        End If
        LockingChtNum = False
    Else
        tLocData.LocChtNum = sChtNum
        tLocData.LocLevCod = sLevCod
        tLocData.LocExeNam = sExeName
        tLocData.LocUidCod = sUidCod
        tLocData.LocIpAddr = sIPAddr
        tLocData.LocChtDtm = AddCentury(SystemDate()) & SystemTime()
        LocChtStore sCmpKey, sRetVal, tLocData
        i = mWrite("LocCht", sCmpKey, sRetVal)
        LockingChtNum = True
    End If
End Function

''=======================================================================================
''outslp.1
''------------
''01. 보험유형 (기타일경우는 기타내용)   : 보험, 보호, 산재, 자보, 일반등 한글로 표시
''02. 요양기관기호                       :
''03. 교부년월일 및 번호                 : YYYYMMDD-12345
''04. 환자성명                           :
''05. 환자주민번호                       : 13자리
''06. 의료기관명칭                       :
''07. 의료기관전화번호                   :
''08. 의료기관팩스번호                   :
''09. 의료기관e-mail주소                 :
''10. 질병분류기호1                      :
''11. 질병분류기호2                      :
''12. 처방의료인의 성명                  :
''13. 처방의료인의 서명                  : 그림 파일이 있는곳의 경로와 파일명
''14. 처방의료인의 면허종별              :
''15. 처방의료인의 면허번호              :
''16. 사용기간                           :
''17. 의약품 총갯수
''18. 재처방여부
''19. 진료과명
''20. 의사전화번호
''21. 의사EMAIL
''22. 보훈번호
''=======================================================================================
'
'Public Sub WriteOutInf2OutSlp(sPrmOutDte As String, sPrmOutNum As String, tPrmOcmData As OcmInfRec, sPrmRePrint As String)     '*^^*
'
'    Dim sBufCurKey As String
'    Dim sBufCmpKey As String
'    Dim sBufRetVal As String
'
'    Dim sCurKey As String
'    Dim sRetVal As String
'
'    Dim tOspData As OspInfRec
'    Dim tIspData As IspInfRec
'    Dim tOutData As OutInfRec
'    Dim tHspData As HspMstRec
'    Dim tOicData As OicInfRec
'    Dim tUidData As UidMstRec
'    Dim tPmdData As PmdInfRec
'
'    ReDim tBufOutData(1 To 50) As OutInfRec
'
'    Dim sOdrNam As String
'    Dim iCount As Integer
'    Dim iCountMed As Integer
'    Dim i As Integer
'    Dim iOicCount As Integer
'    Dim iName As Integer
'    Dim sAssNumber As String        '보훈대상자번호
'    Dim sTmp As String
'    Dim sPatNam As String
'    Dim sElcCod As String           '청구코드
'
'
'    Dim iarMed As Integer
'    Dim iarInj As Integer
'
'    '원외처방전이 있는지 읽어본다!
'    sBufCmpKey = sPrmOutDte & Chr(5) & Format(Trim(sPrmOutNum), "@@@@@") & Chr(5)
'    sBufCurKey = sBufCmpKey
'    sBufCurKey = mSetNext("OutInf", sBufCurKey)
'    Do
'        sBufCurKey = mReadNext("OutInf", sBufCurKey, sBufCmpKey, sBufRetVal)
'        If sBufCurKey = "" Then Exit Do
'        Call OutInfLoad(sBufRetVal, tOutData)
'
'        '---------------------------------------------------------------------------
'        '[2000년 12월 12일 월요일 개발과 김강수]
'        '외래에서 나온 원외처방전 외에 입원에서 나온 원외처방전도 출력하게 해준다!!!
'        '---------------------------------------------------------------------------
'        sCurKey = tOutData.OutOcmNum & Chr(5)
'        sCurKey = mSetReadEqual("OcmInf", sCurKey, sRetVal)
'        If sCurKey <> "" Then
'            '원외처방전이 있는지 읽어본다!
'            sCurKey = tOutData.OutOcmNum & Chr(5) & tOutData.OutOdrNum & Chr(5) & tOutData.OutOdrSeq & Chr(5)
'            sCurKey = mSetReadEqual("OspInf", sCurKey, sRetVal)
'            Call OspInfLoad(sRetVal, tOspData)
'        Else
'            '원외처방전이 있는지 읽어본다!
'            sCurKey = tOutData.OutOcmNum & Chr(5) & tOutData.OutOdrNum & Chr(5) & tOutData.OutOdrSeq & Chr(5)
'            sCurKey = mSetReadEqual("IspInf", sCurKey, sRetVal)
'            Call IspInfLoad(sRetVal, tIspData)
'            tOspData.OspItmCod = tIspData.IspItmCod
'        End If
'
'        '약
'        'If Left(tOspData.OspItmCod, 2) = "03" Then
'        If Left(tOspData.OspItmCod, 2) = "03" And Trim(tOspData.OspOdrStt <> "OC") And (tOspData.OspInsYon = "6" Or tOspData.OspInsYon = "8" Or tOspData.OspInsYon = "9") Then
'            iCountMed = iCountMed + 1
'
'        '주사
'        'ElseIf Left(tOspData.OspItmCod, 2) = "04" Then
'        ElseIf Left(tOspData.OspItmCod, 2) = "04" And Trim(tOspData.OspOdrStt <> "OC") And (tOspData.OspInsYon = "6" Or tOspData.OspInsYon = "8" Or tOspData.OspInsYon = "9") Then
'            iCount = iCount + 1
'        End If
'
'        i = i + 1
'        '일단은 배열에 자료를 담아놓는다!
'        tBufOutData(i) = tOutData
'    Loop
'
'    i = 0
'
'    '원외처방전이 없으면 그냥 나간다.1
'    If iCount = 0 And iCountMed = 0 Then Exit Sub
'
'    '2003/03/26 원외처방을 Txt 화일로 안쓰고 그냥 하자....
'
'    '2001/11/20 neverdie 원외처방폴더가 없을 경우에는 만들어줘야한다.
'    '안만들어주게 되면 프로그램이 종료되어 버린다.
''    If Dir("C:\원외처방", vbDirectory) = "" Then
''        Call MkDir("C:\원외처방")
''    End If
''
''    Open "C:\원외처방\outslp.1" For Output As #1
'
'    sBufCurKey = tPrmOcmData.OcmChtNum & Chr(5) & tPrmOcmData.OcmInsCod & Chr(5) & tPrmOcmData.OcmInsSeq & Chr(5)
'    sBufCurKey = mSetReadEqual("PmdInf", sBufCurKey, sBufRetVal)
'    If sBufCurKey = "" Then
'        sAssNumber = ""
'    Else
'        Call PmdInfLoad(sBufRetVal, tPmdData)
'        If tPmdData.PmdXplNum <> "" Then
'            sAssNumber = tPmdData.PmdXplNum
'        End If
'    End If
'
'    '보험유형
'    Select Case tPrmOcmData.OcmInsCod
'    Case "31", "32", "33"
'        mgs_Head(HEAD_INSCOD) = "보험"
'    Case "21" To "29"
'        mgs_Head(HEAD_INSCOD) = "자보"
'    Case "41"
'        mgs_Head(HEAD_INSCOD) = "산재"
'    Case "51" To "59"
'        mgs_Head(HEAD_INSCOD) = "보호"
'    Case Else
'        If sAssNumber <> "" Then
'            mgs_Head(HEAD_INSCOD) = "보훈"
'        ElseIf tPrmOcmData.OcmInsCod = "11" Then
'            mgs_Head(HEAD_INSCOD) = "일반"
'        Else
'            sTmp = MasterHelpDetail("DtlMst", "INSTBL" & Chr(5) & tPrmOcmData.OcmInsCod & Chr(5), "INSTBL" & Chr(5) & tPrmOcmData.OcmInsCod & Chr(5), 3)
'            If sTmp = "" Then
'                sTmp = "일반"
'            End If
'
'            mgs_Head(HEAD_INSCOD) = sTmp
'
'        End If
'    End Select
'
'    '요양기관번호
'    sBufCurKey = "HNT001" & Chr(5)
'    sBufCurKey = mSetReadEqual("HspMst", sBufCurKey, sBufRetVal)
'    If sBufCurKey <> "" Then
'        Call HspMstLoad(sBufRetVal, tHspData)
'        mgs_Head(HEAD_HSPCOD) = tHspData.HspInsNum
'    End If
'
'    '교부번호
'    mgs_Head(HEAD_OUTNUM) = sPrmOutDte & sPrmOutNum
'
'    '환자성명
'    sPatNam = MasterHelpDetail("PbsInf", tPrmOcmData.OcmChtNum & Chr(5), tPrmOcmData.OcmChtNum & Chr(5), 2)
'    mgs_Head(HEAD_PATNAM) = sPatNam
'
'    '환자주민번호
'    mgs_Head(HEAD_RESNUM) = MasterHelpDetail("PbsInf", tPrmOcmData.OcmChtNum & Chr(5), tPrmOcmData.OcmChtNum & Chr(5), 3)
'
'    '의료기관명칭
'    mgs_Head(HEAD_HSPNAM) = tHspData.HspNam
'    'Print #1, ""
'
'    '의료기관전화번호
'    mgs_Head(HEAD_TELNUM) = MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "TELNUM" & Chr(5), "OUTCFG" & Chr(5) & "TELNUM" & Chr(5), 3)
'    'Print #1, ""
'
'    '의료기관FAX번호
'    mgs_Head(HEAD_FAXNUM) = MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "FAXNUM" & Chr(5), "OUTCFG" & Chr(5) & "FAXNUM" & Chr(5), 3)
'        'Print #1, ""
'
'    '의료기관e-mail
'    mgs_Head(HEAD_EMAIL) = MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "EMAIL" & Chr(5), "OUTCFG" & Chr(5) & "EMAIL" & Chr(5), 3)
'    'Print #1, ""
'
'    '질병출력
'    sBufCurKey = tPrmOcmData.OcmNum & Chr(5)
'    sBufCmpKey = sBufCurKey
'    sBufCurKey = mSetNext("OicInf", sBufCurKey)
'    Do
'        sBufCurKey = mReadNext("OicInf", sBufCurKey, sBufCmpKey, sBufRetVal)
'        If sBufCurKey = "" Then
'            mgs_Head(HEAD_ICD1) = ""
'            mgs_Head(HEAD_ICD2) = ""
'            Exit Do
'        End If
'
'        Call OicInfLoad(sBufRetVal, tOicData)
'        iOicCount = iOicCount + 1
'        If iOicCount > 2 Then Exit Do
'        Select Case iOicCount
'            Case 1
'                mgs_Head(HEAD_ICD1) = tOicData.OicIcdCod
'            Case 2
'                mgs_Head(HEAD_ICD2) = tOicData.OicIcdCod
'        End Select
'    Loop
'
'
''    If iOicCount = 0 Then
''        Print #1, ""
''        Print #1, ""
''    ElseIf iOicCount = 1 Then
''        Print #1, ""
''    End If
'
'    If tPrmOcmData.OcmDtrCod = "" Then
'        tPrmOcmData.OcmDtrCod = MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "DTRCOD" & Chr(5), "OUTCFG" & Chr(5) & "DTRCOD" & Chr(5), 3)
'    End If
'
'    sBufCurKey = tPrmOcmData.OcmDtrCod & Chr(5)
'    sBufCmpKey = sBufCurKey
'    sBufCurKey = mSetReadEqual("UidMst", sBufCurKey, sBufRetVal)
'    If sBufCurKey <> "" Then
'        Call UidMstLoad(sBufRetVal, tUidData)
'        '의사이름
'        mgs_Head(HEAD_DTRNAM) = tUidData.UidNam
'
'        '의사sign
'        mgs_Head(HEAD_DTRSGN) = tUidData.UidSgnDir & tUidData.UidSgnFle
'
'        '면허종별
'        mgs_Head(HEAD_DTRTYP) = "의사"
'
'        '면허번호
'        mgs_Head(HEAD_DTRNUM) = tUidData.UidLicNum
'
'        '사용기간
'        mgs_Head(HEAD_DURDAY) = MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "OUTDAY" & Chr(5), "OUTCFG" & Chr(5) & "OUTDAY" & Chr(5), 3)
'
'        '의약품갯수
'        mgi_Total = iCountMed
'        mgs_Head(HEAD_TOTCNT) = iCountMed
'
'        '재출력여부
'        mgs_Head(HEAD_REPRINT) = sPrmRePrint
'
'        '19. 진료과명
'        'Print #1, MasterHelpDetail("DtlMst", "DEPTBL" & Chr(5) & tUidData.UidDepCod & Chr(5), "DEPTBL" & Chr(5) & tUidData.UidDepCod & Chr(5), 3)
'        'Print #1, "(" & tPrmOcmData.OcmInsCod & "-" & tPrmOcmData.OcmChtNum & ")" & Space(2) & MasterHelpDetail("DtlMst", "DEPTBL" & Chr(5) & tUidData.UidDepCod & Chr(5), "DEPTBL" & Chr(5) & tUidData.UidDepCod & Chr(5), 3); "  조합:" & tPmdData.PmdAssCod & " 증:" & tPmdData.PmdInsNum
'        'Print #1, MasterHelpDetail("DtlMst", "DEPTBL" & Chr(5) & tUidData.UidDepCod & Chr(5), "DEPTBL" & Chr(5) & tUidData.UidDepCod & Chr(5), 3)
'        mgs_Head(HEAD_DEPNAM) = MasterHelpDetail("DtlMst", "DEPTBL" & Chr(5) & tPrmOcmData.OcmDepCod & Chr(5), "DEPTBL" & Chr(5) & tPrmOcmData.OcmDepCod & Chr(5), 3)
'
'        '20. 의사전화번호
'        mgs_Head(HEAD_DTRTELNUM) = tUidData.UidTelNum
'
'        '21. 의사EMAIL
'        'Print #1, tUidData.UidMalAdd
'        mgs_Head(HEAD_DTREMAIL) = MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "EMAIL" & Chr(5), "OUTCFG" & Chr(5) & "EMAIL" & Chr(5), 3)
'
'        '22. 보훈번호
'        mgs_Head(HEAD_NATNUM) = sAssNumber
'
'        '환자 차트번호
'        mgs_Head(HEAD_CHTNUM) = tBufOutData(1).OutChtNum
'
'
'    End If
'
''    Close #1
'
'    '===========================
'    'outslp.2
'    '===========================
'    '01. 명칭
'    '02. 1회투약량
'    '03. 1일투여횟수
'    '04. 총투약일수
'    '05. 용법
'    '===========================
''    Open "c:\원외처방\outslp.2" For Output As #2
'
'        For i = 1 To iCount + iCountMed
'
'            '원외처방전이 있는지 읽어본다!
'            sBufCurKey = tBufOutData(i).OutOcmNum & Chr(5) & tBufOutData(i).OutOdrNum & Chr(5) & tBufOutData(i).OutOdrSeq & Chr(5)
'            sBufCurKey = mSetReadEqual("OspInf", sBufCurKey, sBufRetVal)
'            Call OspInfLoad(sBufRetVal, tOspData)
'
'            If Left(tOspData.OspItmCod, 2) = "03" Then
'
'                iarMed = iarMed + 1
'
'                '처방명칭 앞에 청구코드를 넣어준다.
'                sElcCod = "[" & MasterHelpDetail("SeeMst", tOspData.OspOdrCod & Chr(5), tOspData.OspOdrCod & Chr(5), 8) & "]"
'
'                '6:영문명칭 7:한글명칭, 31:성분명칭
'                Select Case MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "PRTNAM" & Chr(5), "OUTCFG" & Chr(5) & "PRTNAM" & Chr(5), 3)
'                Case "1"
'                    iName = 6
'                Case "2"
'                    iName = 7
'                Case "3"
'                    iName = 31
'                End Select
'
'                '오더명칭 출력
'                If tOspData.OspInsYon = "9" Or (tOspData.OspInsYon = "9" And Right(tOspData.OspItmCod, 1) = "2") Then
'                    sOdrNam = MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "PRTNON" & Chr(5), "OUTCFG" & Chr(5) & "PRTNON" & Chr(5), 3) & sElcCod & " " & MasterHelpDetail("SeeMst", tOspData.OspOdrCod & Chr(5), tOspData.OspOdrCod & Chr(5), iName)
'                ElseIf tOspData.OspInsYon = "8" Or (tOspData.OspInsYon = "8" And Right(tOspData.OspItmCod, 1) = "2") Then
'                    sOdrNam = MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "PRT100" & Chr(5), "OUTCFG" & Chr(5) & "PRT100" & Chr(5), 3) & sElcCod & " " & MasterHelpDetail("SeeMst", tOspData.OspOdrCod & Chr(5), tOspData.OspOdrCod & Chr(5), iName)
'                Else
'                    sOdrNam = sElcCod & " " & MasterHelpDetail("SeeMst", tOspData.OspOdrCod & Chr(5), tOspData.OspOdrCod & Chr(5), iName)
'                End If
'
'                '특기사항이 있는 경우 오더명에 기록하고 출력시 다음 라인에 출력합니다.
'                If tOspData.OspSplYon = "Y" And Trim(Replace(tOspData.OspSplCmt, vbCrLf, " ")) <> "" Then
'                    sOdrNam = sOdrNam & Chr(5) & Trim(Replace(tOspData.OspSplCmt, vbCrLf, " "))
'                End If
'
'                fbs_Tmp_Data(iarMed, 0) = sOdrNam
'
''                Print #2, sOdrNam
'
'                If CDouble(tOspData.OspOdrTms) = 0 Then
'                    sTmp = "처방전의 횟수가 0으로 입력되어 처방전 출력에 오류가 있읍니다."
'                    sTmp = sTmp & Chr(13) & Chr(10) & "아래 환자의 처방전을 지우고 다시 수납하십시요."
'                    sTmp = sTmp & Chr(13) & Chr(10)
'                    sTmp = sTmp & Chr(13) & Chr(10) & "Chart No : " & tPrmOcmData.OcmChtNum
'                    sTmp = sTmp & Chr(13) & Chr(10) & "환자성명 : " & sPatNam
'
'                    MsgBox sTmp
'                    fbs_Tmp_Data(iarMed, 1) = "Error"
'                Else
'                    'Print #2, CStr(CDouble(tOspData.OspOdrQty) / CDouble(tOspData.OspOdrTms))
'
'                    '2001년 12월 05일 상주 적십자 병원 ..
'                    '.. 함량처방이 아리송~~ 아리송~~ 하다.
'                    '약국에서 수납에서 보여주는 데로 표시해 달라고 한다.
'                    '이래서는 안되는것 같지마 일단은 원하는데로 해주고... 좀 지켜보자...
'                    If CDouble(tOspData.OspBasUnt) = CDouble(tOspData.OspOdrQty) Or tOspData.OspBasUnt = "" Then
'                        fbs_Tmp_Data(iarMed, 1) = CStr(CDouble(tOspData.OspOdrQty) / CDouble(tOspData.OspOdrTms))
'                    Else
'                        fbs_Tmp_Data(iarMed, 1) = CStr(CDouble(tOspData.OspBasUnt) / CDouble(tOspData.OspOdrTms))
'                    End If
'                End If
'
'                fbs_Tmp_Data(iarMed, 2) = tOspData.OspOdrTms
'                fbs_Tmp_Data(iarMed, 3) = tOspData.OspOdrDay
'
'                If tOspData.OspUsgCod = "" Then
'                    fbs_Tmp_Data(iarMed, 4) = tOspData.OspSplCmt
'                Else
'                    fbs_Tmp_Data(iarMed, 4) = MasterHelpDetail("UsgMst", tOspData.OspUsgCod & Chr(5), tOspData.OspUsgCod & Chr(5), 3)
'                End If
''            End If
'
'
'
'            ElseIf Left(tOspData.OspItmCod, 2) = "04" Then
'
'                iarInj = iarInj + 1
'
'                sElcCod = "[" & MasterHelpDetail("SeeMst", tOspData.OspOdrCod & Chr(5), tOspData.OspOdrCod & Chr(5), 8) & "]"
'
'                '6:영문명칭 7:한글명칭, 31:성분명칭
'                Select Case MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "PRTNAM" & Chr(5), "OUTCFG" & Chr(5) & "PRTNAM" & Chr(5), 3)
'                Case "1"
'                    iName = 6
'                Case "2"
'                    iName = 7
'                Case "3"
'                    iName = 31
'                End Select
'
'                If tOspData.OspInsYon = "9" Or (tOspData.OspInsYon = "9" And Right(tOspData.OspItmCod, 1) = "2") Then
'                    fbs_InjData(iarInj, 0) = MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "PRTNON" & Chr(5), "OUTCFG" & Chr(5) & "PRTNON" & Chr(5), 3) & sElcCod & " " & MasterHelpDetail("SeeMst", tOspData.OspOdrCod & Chr(5), tOspData.OspOdrCod & Chr(5), iName) & " [" & MasterHelpDetail("SeeMst", tOspData.OspOdrCod & Chr(5), tOspData.OspOdrCod & Chr(5), 24) & "]"
'                ElseIf tOspData.OspInsYon = "8" Or (tOspData.OspInsYon = "8" And Right(tOspData.OspItmCod, 1) = "2") Then
'                    fbs_InjData(iarInj, 0) = MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "PRT100" & Chr(5), "OUTCFG" & Chr(5) & "PRT100" & Chr(5), 3) & sElcCod & " " & " " & MasterHelpDetail("SeeMst", tOspData.OspOdrCod & Chr(5), tOspData.OspOdrCod & Chr(5), iName) & " [" & MasterHelpDetail("SeeMst", tOspData.OspOdrCod & Chr(5), tOspData.OspOdrCod & Chr(5), 24) & "]"
'                Else
'                    fbs_InjData(iarInj, 0) = sElcCod & " " & MasterHelpDetail("SeeMst", tOspData.OspOdrCod & Chr(5), tOspData.OspOdrCod & Chr(5), iName) & " [" & MasterHelpDetail("SeeMst", tOspData.OspOdrCod & Chr(5), tOspData.OspOdrCod & Chr(5), 24) & "]"
'                End If
'
'                If CDouble(tOspData.OspOdrTms) = 0 Then
'                    sTmp = "처방전의 횟수가 0으로 입력되어 처방전 출력에 오류가 있읍니다."
'                    sTmp = sTmp & Chr(13) & Chr(10) & "아래 환자의 처방전을 지우고 다시 수납하십시요."
'                    sTmp = sTmp & Chr(13) & Chr(10)
'                    sTmp = sTmp & Chr(13) & Chr(10) & "Chart No : " & tPrmOcmData.OcmChtNum
'                    sTmp = sTmp & Chr(13) & Chr(10) & "환자성명 : " & sPatNam
'
'                    MsgBox sTmp
'                    fbs_InjData(iarInj, 1) = "Error"
'                Else
'                    fbs_InjData(iarInj, 1) = CStr(CDouble(tOspData.OspOdrQty) / CDouble(tOspData.OspOdrTms))
'                End If
'
'                fbs_InjData(iarInj, 2) = tOspData.OspOdrTms
'                fbs_InjData(iarInj, 3) = tOspData.OspOdrDay
'                If CInteger(tOspData.OspInsYon) < 6 Then
'                    fbs_InjData(iarInj, 4) = "원내"
'                Else
'                    fbs_InjData(iarInj, 4) = "원외"
'                End If
'            End If
'
'
'
'        Next
'
''    Close #2
'
'    '===========================
'    'outslp.3 (주사제)
'    '===========================
'    '01. 명칭
'    '02. 1회투약량
'    '03. 1일투여횟수
'    '04. 총투약일수
'    '05. 용법
'    '===========================
''    Open "c:\원외처방\outslp.3" For Output As #2
'
''        For i = 1 To iCount + iCountMed
''
''            '원외처방전이 있는지 읽어본다!
''            sBufCurKey = tBufOutData(i).OutOcmNum & Chr(5) & tBufOutData(i).OutOdrNum & Chr(5) & tBufOutData(i).OutOdrSeq & Chr(5)
''            sBufCurKey = mSetReadEqual("OspInf", sBufCurKey, sBufRetVal)
''            Call OspInfLoad(sBufRetVal, tOspData)
''
''            If Left(tOspData.OspItmCod, 2) = "04" Then
''
''                sElcCod = "[" & MasterHelpDetail("SeeMst", tOspData.OspOdrCod & Chr(5), tOspData.OspOdrCod & Chr(5), 8) & "]"
''
''                '6:영문명칭 7:한글명칭, 31:성분명칭
''                Select Case MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "PRTNAM" & Chr(5), "OUTCFG" & Chr(5) & "PRTNAM" & Chr(5), 3)
''                Case "1"
''                    iName = 6
''                Case "2"
''                    iName = 7
''                Case "3"
''                    iName = 31
''                End Select
''
''                If tOspData.OspInsYon = "9" Or (tOspData.OspInsYon = "9" And Right(tOspData.OspItmCod, 1) = "2") Then
''                    Print #2, MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "PRTNON" & Chr(5), "OUTCFG" & Chr(5) & "PRTNON" & Chr(5), 3) & sElcCod & " " & MasterHelpDetail("SeeMst", tOspData.OspOdrCod & Chr(5), tOspData.OspOdrCod & Chr(5), iName) & " [" & MasterHelpDetail("SeeMst", tOspData.OspOdrCod & Chr(5), tOspData.OspOdrCod & Chr(5), 24) & "]"
''                ElseIf tOspData.OspInsYon = "8" Or (tOspData.OspInsYon = "8" And Right(tOspData.OspItmCod, 1) = "2") Then
''                    Print #2, MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "PRT100" & Chr(5), "OUTCFG" & Chr(5) & "PRT100" & Chr(5), 3) & sElcCod & " " & " " & MasterHelpDetail("SeeMst", tOspData.OspOdrCod & Chr(5), tOspData.OspOdrCod & Chr(5), iName) & " [" & MasterHelpDetail("SeeMst", tOspData.OspOdrCod & Chr(5), tOspData.OspOdrCod & Chr(5), 24) & "]"
''                Else
''                    Print #2, sElcCod & " " & MasterHelpDetail("SeeMst", tOspData.OspOdrCod & Chr(5), tOspData.OspOdrCod & Chr(5), iName) & " [" & MasterHelpDetail("SeeMst", tOspData.OspOdrCod & Chr(5), tOspData.OspOdrCod & Chr(5), 24) & "]"
''                End If
''
''                If CDouble(tOspData.OspOdrTms) = 0 Then
''                    sTmp = "처방전의 횟수가 0으로 입력되어 처방전 출력에 오류가 있읍니다."
''                    sTmp = sTmp & Chr(13) & Chr(10) & "아래 환자의 처방전을 지우고 다시 수납하십시요."
''                    sTmp = sTmp & Chr(13) & Chr(10)
''                    sTmp = sTmp & Chr(13) & Chr(10) & "Chart No : " & tPrmOcmData.OcmChtNum
''                    sTmp = sTmp & Chr(13) & Chr(10) & "환자성명 : " & sPatNam
''
''                    MsgBox sTmp
''                    Print #2, "Error"
''                Else
''                    Print #2, CStr(CDouble(tOspData.OspOdrQty) / CDouble(tOspData.OspOdrTms))
''                End If
''
''                Print #2, tOspData.OspOdrTms
''                Print #2, tOspData.OspOdrDay
''                If CInteger(tOspData.OspInsYon) < 6 Then
''                    Print #2, "원내"
''                Else
''                    Print #2, "원외"
''                End If
''            End If
''
''        Next
''
''    Close #2
'
'End Sub
'
'

'Sub WriteOutInf2OutSlpIcm(sPrmOutDte As String, sPrmOutNum As String, tPrmIcmData As IcmInfRec, sPrmRePrint As String)
'
'    Dim sBufCurKey As String
'    Dim sBufCmpKey As String
'    Dim sBufRetVal As String
'
'    Dim sCurKey As String
'    Dim sRetVal As String
'
'    Dim tIspData As IspInfRec
'    Dim tOutData As OutInfRec
'    Dim tHspData As HspMstRec
'    Dim tOicData As OicInfRec
'    Dim tUidData As UidMstRec
'    Dim tPmdData As PmdInfRec
'
'    ReDim tBufOutData(1 To 50) As OutInfRec
'
'    Dim iCount As Integer
'    Dim iCountMed As Integer
'    Dim i As Integer, j As Integer
'    Dim iOicCount As Integer
'    Dim iName As Integer
'    Dim sAssNumber As String        '보훈대상자번호
'
'    Dim sCanOutNum() As String
'    Dim iCanCnt As Integer
'    Dim sTmpStr As String
'    Dim sTmp As String
'    Dim sElcCod As String
'
'    Dim iarMed As Integer
'    Dim iarInj As Integer
'
'
'    iCanCnt = 0
'    '원외처방전이 있는지 읽어본다!
'    sBufCmpKey = sPrmOutDte & Chr(5) & Format(Trim(sPrmOutNum), "@@@@@") & Chr(5)
'    sBufCurKey = sBufCmpKey
'    sBufCurKey = mSetNext("OutInf", sBufCurKey)
'    Do
'        sBufCurKey = mReadNext("OutInf", sBufCurKey, sBufCmpKey, sBufRetVal)
'        If sBufCurKey = "" Then Exit Do
'        Call OutInfLoad(sBufRetVal, tOutData)
'
'        '원외처방전이 있는지 읽어본다!
'        sCurKey = tOutData.OutOcmNum & Chr(5) & tOutData.OutOdrNum & Chr(5) & tOutData.OutOdrSeq & Chr(5)
'        sCurKey = mSetReadEqual("IspInf", sCurKey, sRetVal)
'        Call IspInfLoad(sRetVal, tIspData)
'
'        '약
'        'If Left(tIspData.IspItmCod, 2) = "03" Then
'        If Left(tIspData.IspItmCod, 2) = "03" And tIspData.IspOdrStt <> "OC" And (tIspData.IspInsYon = "6" Or tIspData.IspInsYon = "8" Or tIspData.IspInsYon = "9") Then
'
'            iCountMed = iCountMed + 1
'        '주사
'        'ElseIf Left(tIspData.IspItmCod, 2) = "04" Then
'        ElseIf Left(tIspData.IspItmCod, 2) = "04" And tIspData.IspOdrStt <> "OC" And (tIspData.IspInsYon = "6" Or tIspData.IspInsYon = "8" Or tIspData.IspInsYon = "9") Then
'
'            iCount = iCount + 1
'
'        End If
'        i = i + 1
'        '일단은 배열에 자료를 담아놓는다!
'        tBufOutData(i) = tOutData
'
'        If Trim(tOutData.OutCanNum) <> "" Then
'            Do
'                iCanCnt = iCanCnt + 1
'                sTmpStr = piece(tOutData.OutCanNum, Chr(6), iCanCnt)
'                If sTmpStr = "" Then
'                    iCanCnt = iCanCnt - 1
'                    Exit Do
'                End If
'                For j = 1 To iCanCnt - 1
'                    If tOutData.OutCanNum = sCanOutNum(j) Then Exit For
'                Next
'
'                ReDim Preserve sCanOutNum(1 To iCanCnt)
'                sCanOutNum(iCanCnt) = sTmpStr
'            Loop
'        End If
'
'    Loop
'
'    i = 0
'
'    '원외처방전이 없으면 그냥 나간다.1
'    If iCount = 0 And iCountMed = 0 Then Exit Sub
'
'
'    '2003/03/26 EverSky 원외처방을 Txt에 쓰지 않고 바로 출력한다.
'    '2001/11/20 neverdie 원외처방폴더가 없을 경우에는 만들어줘야한다.
'    '안만들어주게 되면 프로그램이 종료되어 버린다.
'
''    If Dir("C:\원외처방", vbDirectory) = "" Then
''        Call MkDir("C:\원외처방")
''    End If
''
''    Open "C:\원외처방\outslp.1" For Output As #1
'
'    sBufCurKey = tPrmIcmData.IcmChtNum & Chr(5) & tPrmIcmData.IcmInsCod & Chr(5) & tPrmIcmData.IcmInsSeq & Chr(5)
'    sBufCurKey = mSetReadEqual("PmdInf", sBufCurKey, sBufRetVal)
'    If sBufCurKey = "" Then
'        sAssNumber = ""
'    Else
'        Call PmdInfLoad(sBufRetVal, tPmdData)
'        If tPmdData.PmdXplNum <> "" Then
'            sAssNumber = tPmdData.PmdXplNum
'        End If
'    End If
'
'        '보험유형
'        Select Case tPrmIcmData.IcmInsCod
'            Case "31", "32", "33"
'                mgs_Head(HEAD_INSCOD) = "보험"
'            Case "21" To "29"
'                mgs_Head(HEAD_INSCOD) = "자보"
'            Case "41"
'                mgs_Head(HEAD_INSCOD) = "산재"
'            Case "51" To "59"
'                mgs_Head(HEAD_INSCOD) = "보호"
'            Case Else
'                If sAssNumber <> "" Then
'                    mgs_Head(HEAD_INSCOD) = "보훈"
'                ElseIf tPrmIcmData.IcmInsCod = "11" Then
'                    mgs_Head(HEAD_INSCOD) = "일반"
'                Else
'                    sTmp = MasterHelpDetail("DtlMst", "INSTBL" & Chr(5) & tPrmIcmData.IcmInsCod & Chr(5), "INSTBL" & Chr(5) & tPrmIcmData.IcmInsCod & Chr(5), 3)
'                    If sTmp = "" Then
'                        sTmp = "일반"
'                    End If
'
'                    mgs_Head(HEAD_INSCOD) = sTmp
'
'                End If
'        End Select
'
'        '요양기관번호
'        sBufCurKey = "HNT001" & Chr(5)
'        sBufCurKey = mSetReadEqual("HspMst", sBufCurKey, sBufRetVal)
'        If sBufCurKey <> "" Then
'            Call HspMstLoad(sBufRetVal, tHspData)
'            mgs_Head(HEAD_HSPCOD) = tHspData.HspInsNum
'        End If
'
'        '교부번호
'        mgs_Head(HEAD_OUTNUM) = sPrmOutDte & sPrmOutNum
'
'        '환자성명
'        mgs_Head(HEAD_PATNAM) = tBufOutData(1).OutPatNam
'
'        '환자주민번호
'        mgs_Head(HEAD_RESNUM) = tBufOutData(1).OutResNum
'
'        '의료기관명칭
'        mgs_Head(HEAD_HSPNAM) = tHspData.HspNam
'
'        '의료기관전화번호
'        mgs_Head(HEAD_TELNUM) = MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "TELNUM" & Chr(5), "OUTCFG" & Chr(5) & "TELNUM" & Chr(5), 3)
'
'        '의료기관FAX번호
'        mgs_Head(HEAD_FAXNUM) = MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "FAXNUM" & Chr(5), "OUTCFG" & Chr(5) & "FAXNUM" & Chr(5), 3)
'
'        '의료기관e-mail
'        mgs_Head(HEAD_EMAIL) = MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "EMAIL" & Chr(5), "OUTCFG" & Chr(5) & "EMAIL" & Chr(5), 3)
'
'        '질병출력
'        sBufCurKey = tPrmIcmData.IcmOcmNum & Chr(5)
'        sBufCmpKey = sBufCurKey
'        sBufCurKey = mSetNext("OicInf", sBufCurKey)
'        Do
''            sBufCurKey = mReadNext("OicInf", sBufCurKey, sBufCmpKey, sBufRetVal)
''            If sBufCurKey = "" Then Exit Do
''            Call OicInfLoad(sBufRetVal, tOicData)
''            iOicCount = iOicCount + 1
''            If iOicCount > 2 Then Exit Do
''            Print #1, tOicData.OicIcdCod
'
'            sBufCurKey = mReadNext("OicInf", sBufCurKey, sBufCmpKey, sBufRetVal)
'            If sBufCurKey = "" Then
'                mgs_Head(HEAD_ICD1) = ""
'                mgs_Head(HEAD_ICD2) = ""
'                Exit Do
'            End If
'
'            Call OicInfLoad(sBufRetVal, tOicData)
'            iOicCount = iOicCount + 1
'            If iOicCount > 2 Then Exit Do
'            Select Case iOicCount
'                Case 1
'                    mgs_Head(HEAD_ICD1) = tOicData.OicIcdCod
'                Case 2
'                    mgs_Head(HEAD_ICD2) = tOicData.OicIcdCod
'            End Select
'
'        Loop
'
''        If iOicCount = 0 Then
''            Print #1, ""
''            Print #1, ""
''        ElseIf iOicCount = 1 Then
''            Print #1, ""
''        End If
'
'
'
'        sBufCurKey = tPrmIcmData.IcmDtrCod & Chr(5)
'        sBufCmpKey = sBufCurKey
'        sBufCurKey = mSetReadEqual("UidMst", sBufCurKey, sBufRetVal)
'        If sBufCurKey <> "" Then
'            Call UidMstLoad(sBufRetVal, tUidData)
'            '의사이름
'            mgs_Head(HEAD_DTRNAM) = tUidData.UidNam
'
'            '의사sign
'            mgs_Head(HEAD_DTRSGN) = tUidData.UidSgnDir & tUidData.UidSgnFle
'
'            '면허종별
'            mgs_Head(HEAD_DTRTYP) = "의사"
'
'            '면허번호
'            mgs_Head(HEAD_DTRNUM) = tUidData.UidLicNum
'
'            '사용기간
'            mgs_Head(HEAD_DURDAY) = MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "OUTDAY" & Chr(5), "OUTCFG" & Chr(5) & "OUTDAY" & Chr(5), 3)
'
'            '의약품갯수
'            mgs_Head(HEAD_TOTCNT) = iCountMed
'
'            '재출력여부
'            mgs_Head(HEAD_REPRINT) = sPrmRePrint
'
'            '19. 진료과명
'            '@-@;;
'            'Print #1, tPrmOcmData.OcmInsCod & "-" & tPrmOcmData.OcmChtNum & Space(2) & MasterHelpDetail("DtlMst", "DEPTBL" & Chr(5) & tUidData.UidDepCod & Chr(5), "DEPTBL" & Chr(5) & tUidData.UidDepCod & Chr(5), 3)
'            mgs_Head(HEAD_DEPNAM) = "(" & tPrmIcmData.IcmInsCod & "-" & tPrmIcmData.IcmChtNum & ")" & Space(2) & MasterHelpDetail("DtlMst", "DEPTBL" & Chr(5) & tUidData.UidDepCod & Chr(5), "DEPTBL" & Chr(5) & tUidData.UidDepCod & Chr(5), 3)
'            '@-@;;
'
'            '20. 의사전화번호
'            mgs_Head(HEAD_DTRTELNUM) = tUidData.UidTelNum
'
'            '21. 의사EMAIL
'            mgs_Head(HEAD_DTREMAIL) = tUidData.UidMalAdd
'
'            '22. 보훈번호
'            For i = 1 To iCanCnt
'                If i = 1 Then
'                    sAssNumber = sAssNumber & Space(2) & "["
'                End If
'                sAssNumber = sAssNumber & Space(1) & Trim(sCanOutNum(i))
'                If i <> iCanCnt Then
'                    sAssNumber = sAssNumber & ","
'                Else
'                    sAssNumber = sAssNumber & " 회수]"
'                End If
'            Next
'
'            mgs_Head(HEAD_NATNUM) = sAssNumber
'
'        End If
'
''    Close #1
'
'    '===========================
'    'outslp.2
'    '===========================
'    '01. 명칭
'    '02. 1회투약량
'    '03. 1일투여횟수
'    '04. 총투약일수
'    '05. 용법
'    '===========================
''    Open "c:\원외처방\outslp.2" For Output As #2
'
'        For i = 1 To iCount + iCountMed
'
'            '원외처방전이 있는지 읽어본다!
'            sBufCurKey = tBufOutData(i).OutOcmNum & Chr(5) & tBufOutData(i).OutOdrNum & Chr(5) & tBufOutData(i).OutOdrSeq & Chr(5)
'            sBufCurKey = mSetReadEqual("IspInf", sBufCurKey, sBufRetVal)
'            Call IspInfLoad(sBufRetVal, tIspData)
'
'            If Left(tIspData.IspItmCod, 2) = "03" Then
'
'                iarMed = iarMed + 1
'
'                '처방명칭 앞에 청구코드를 넣어준다.
'                sElcCod = "[" & MasterHelpDetail("SeeMst", tIspData.IspOdrCod & Chr(5), tIspData.IspOdrCod & Chr(5), 8) & "]"
'
'                '6:영문명칭 7:한글명칭, 31:성분명칭
'                Select Case MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "PRTNAM" & Chr(5), "OUTCFG" & Chr(5) & "PRTNAM" & Chr(5), 3)
'                Case "1"
'                    iName = 6
'                Case "2"
'                    iName = 7
'                Case "3"
'                    iName = 31
'                End Select
'
'                '[의약분업(비급여)]...원래가 비급여이거나 강제로 비급여를 만들었으면...
'                'Print #2, MasterHelpDetail("SeeMst", tIspData.IspOdrCod & Chr(5), tIspData.IspOdrCod & Chr(5), iName)
'                If tIspData.IspInsYon = "9" Or (tIspData.IspInsYon = "9" And Right(tIspData.IspItmCod, 1) = "2") Then
'                    fbs_Tmp_Data(iarMed, 0) = MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "PRTNON" & Chr(5), "OUTCFG" & Chr(5) & "PRTNON" & Chr(5), 3) & MasterHelpDetail("SeeMst", tIspData.IspOdrCod & Chr(5), tIspData.IspOdrCod & Chr(5), iName)
'                ElseIf tIspData.IspInsYon = "8" Or (tIspData.IspInsYon = "8" And Right(tIspData.IspItmCod, 1) = "2") Then
'                    fbs_Tmp_Data(iarMed, 0) = MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "PRT100" & Chr(5), "OUTCFG" & Chr(5) & "PRT100" & Chr(5), 3) & MasterHelpDetail("SeeMst", tIspData.IspOdrCod & Chr(5), tIspData.IspOdrCod & Chr(5), iName)
'                Else
'                    fbs_Tmp_Data(iarMed, 0) = sElcCod & "  " & MasterHelpDetail("SeeMst", tIspData.IspOdrCod & Chr(5), tIspData.IspOdrCod & Chr(5), iName)
'                End If
'                '[의약분업(비급여)]
'                fbs_Tmp_Data(iarMed, 1) = CStr(CDouble(tIspData.IspOdrQty) / CDouble(tIspData.IspOdrTms))
'                fbs_Tmp_Data(iarMed, 2) = tIspData.IspOdrTms
'                fbs_Tmp_Data(iarMed, 3) = tIspData.IspOdrDay
'                fbs_Tmp_Data(iarMed, 4) = MasterHelpDetail("UsgMst", tIspData.IspUsgCod & Chr(5), tIspData.IspUsgCod & Chr(5), 3)
'
'
''            End If
'
'
'            ElseIf Left(tIspData.IspItmCod, 2) = "04" Then
'
'                iarInj = iarInj + 1
'                '6:영문명칭 7:한글명칭, 31:성분명칭
'                Select Case MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "PRTNAM" & Chr(5), "OUTCFG" & Chr(5) & "PRTNAM" & Chr(5), 3)
'                Case "1"
'                    iName = 6
'                Case "2"
'                    iName = 7
'                Case "3"
'                    iName = 31
'                End Select
'
'                If Trim(tIspData.IspMthCod) = "" Then
'                    fbs_InjData(iarInj, 0) = MasterHelpDetail("SeeMst", tIspData.IspOdrCod & Chr(5), tIspData.IspOdrCod & Chr(5), iName)
'                Else
'                    fbs_InjData(iarInj, 0) = "(" & tIspData.IspMthCod & ")" & MasterHelpDetail("SeeMst", tIspData.IspOdrCod & Chr(5), tIspData.IspOdrCod & Chr(5), iName)
'                End If
'
'                fbs_InjData(iarInj, 1) = CStr(CDouble(tIspData.IspOdrQty) / CDouble(tIspData.IspOdrTms))
'                fbs_InjData(iarInj, 2) = tIspData.IspOdrTms
'                fbs_InjData(iarInj, 3) = tIspData.IspOdrDay
'                If CInteger(tIspData.IspInsYon) < 6 Then
'                    fbs_InjData(iarInj, 4) = "원내"
'                Else
'                    fbs_InjData(iarInj, 4) = "원외"
'                End If
'
'
'            End If
'
'
'
'        Next
'
''    Close #2
'
'    '===========================
'    'outslp.3 (주사제)
'    '===========================
'    '01. 명칭
'    '02. 1회투약량
'    '03. 1일투여횟수
'    '04. 총투약일수
'    '05. 용법
'    '===========================
''    Open "c:\원외처방\outslp.3" For Output As #2
''
''        For i = 1 To iCount + iCountMed
''
''            '원외처방전이 있는지 읽어본다!
''            sBufCurKey = tBufOutData(i).OutOcmNum & Chr(5) & tBufOutData(i).OutOdrNum & Chr(5) & tBufOutData(i).OutOdrSeq & Chr(5)
''            sBufCurKey = mSetReadEqual("IspInf", sBufCurKey, sBufRetVal)
''            Call IspInfLoad(sBufRetVal, tIspData)
''
''            If Left(tIspData.IspItmCod, 2) = "04" Then
''
''                '6:영문명칭 7:한글명칭, 31:성분명칭
''                Select Case MasterHelpDetail("DtlMst", "OUTCFG" & Chr(5) & "PRTNAM" & Chr(5), "OUTCFG" & Chr(5) & "PRTNAM" & Chr(5), 3)
''                Case "1"
''                    iName = 6
''                Case "2"
''                    iName = 7
''                Case "3"
''                    iName = 31
''                End Select
''
''                If Trim(tIspData.IspMthCod) = "" Then
''                    Print #2, MasterHelpDetail("SeeMst", tIspData.IspOdrCod & Chr(5), tIspData.IspOdrCod & Chr(5), iName)
''                Else
''                    Print #2, "(" & tIspData.IspMthCod & ")" & MasterHelpDetail("SeeMst", tIspData.IspOdrCod & Chr(5), tIspData.IspOdrCod & Chr(5), iName)
''                End If
''
''                Print #2, CStr(CDouble(tIspData.IspOdrQty) / CDouble(tIspData.IspOdrTms))
''                Print #2, tIspData.IspOdrTms
''                Print #2, tIspData.IspOdrDay
''                If CInteger(tIspData.IspInsYon) < 6 Then
''                    Print #2, "원내"
''                Else
''                    Print #2, "원외"
''                End If
''            End If
''
''        Next
''
''    Close #2
'
'End Sub
'
'
Public Sub ZipMstRead(sPrmZipCod As String, ZipData As ZipMstRec)

    Dim sZipCurKey As String, sZipRetValue As String

    sZipCurKey = sPrmZipCod & Chr(5)
    sZipCurKey = mSetReadEqual("ZipMst", sZipCurKey, sZipRetValue)

    Call ZipMstLoad(sZipRetValue, ZipData)

End Sub

Public Sub BabInfRead(sPrmChtNum As String, BabData As BabInfRec)

    Dim sCurKey As String, sRetVal As String

    sCurKey = sPrmChtNum & Chr(5)
    sCurKey = mSetReadEqual("BabInf", sCurKey, sRetVal)

    Call BabInfLoad(sRetVal, BabData)

End Sub

Public Function LeftK(mps_Source As String, mpi_Length As Integer) As String
    '한글을 2바이트로 영어및 숫자는 1바이트로 계산하여
    '왼쪽에서 지정한 길이만큼 잘라내는 함수
    
    Dim mbi_Len As Integer, mbi_LenK As Integer
    
    mbi_Len = Len(mps_Source)
    mbi_LenK = LenK(mps_Source)
    
    If mbi_Len = mbi_LenK Then
        LeftK = Left(mps_Source, mpi_Length)
    Else
        LeftK = StrConv(LeftB(StrConv(mps_Source, vbFromUnicode), mpi_Length), vbUnicode)
    End If
    
End Function

Public Function RightK(mps_Source As String, mpi_Length As Integer) As String
    '한글을 2바이트로 영어및 숫자는 1바이트로 계산하여
    '오른쪽에서 지정한 길이만큼 잘라내는 함수

    Dim mbi_Len As Integer, mbi_LenK As Integer
    
    mbi_Len = Len(mps_Source)
    mbi_LenK = LenK(mps_Source)
    
    If mbi_Len = mbi_LenK Then
        RightK = Right(mps_Source, mpi_Length)
    Else
        RightK = StrConv(RightB(StrConv(mps_Source, vbFromUnicode), mpi_Length), vbUnicode)
    End If
    
End Function

Public Function LenK(mps_Str As String) As Integer
' 한글이 포함된 string의 길이 계산시 정확한 길이를 return한다.

    Dim mbi_Result As Integer
    
    mbi_Result = LenB(StrConv(mps_Str, vbFromUnicode))
    LenK = mbi_Result

End Function

Public Function MidK(mps_Source As String, mpi_Start As Integer, Optional mpi_Length As Integer = 0) As String
    '한글을 2바이트로 영어및 숫자는 1바이트로 계산하여
    '중간에서 시작위치부터 지정한 길이만큼 잘라내는 함수
    '길이를 지정하지 않으면 시작위치부터 끝까지 잘라낸다.

    Dim mbi_Len As Integer, mbi_LenK As Integer
    
    mbi_Len = Len(mps_Source)
    mbi_LenK = LenK(mps_Source)
    
    If mbi_Len = mbi_LenK Then
        If mpi_Length = 0 Then
            MidK = Mid(mps_Source, mpi_Start)
        Else
            MidK = Mid(mps_Source, mpi_Start, mpi_Length)
        End If
    Else
        If mpi_Length = 0 Then
            MidK = StrConv(MidB(StrConv(mps_Source, vbFromUnicode), mpi_Start), vbUnicode)
        Else
            MidK = StrConv(MidB(StrConv(mps_Source, vbFromUnicode), mpi_Start, mpi_Length), vbUnicode)
        End If
    End If
    
End Function

Public Sub SpreadSort(spdObj As Object, Row As Long, col As Long, Row2 As Long, Col2 As Long, bSortSw As Boolean, lSortKey_1 As Long, Optional lSortKey_2 As Long = 1)

    With spdObj

        .Row = Row
        .col = col
        .Row2 = Row2
        .Col2 = Col2
        .BlockMode = True
        ' Set sort definition for key 1
        .SortBy = SortByRow
        .SortKey(1) = lSortKey_1
        
        ' Set sort definition for key 2
        .SortKey(2) = lSortKey_2
        
        If bSortSw Then
            .SortKeyOrder(1) = SortKeyOrderAscending
            .SortKeyOrder(2) = SortKeyOrderAscending
        Else
            .SortKeyOrder(1) = SortKeyOrderDescending
            .SortKeyOrder(2) = SortKeyOrderDescending
        End If
                
        .Action = ActionSort
        .BlockMode = False
        
    End With

End Sub

'큰값구하기
Public Function MaxValue(mpv_Value1 As Variant, mpv_Value2 As Variant, Optional mpv_Value3 As Variant = 0) As Currency

    On Error GoTo Error_MaxValue

    If CLong(mpv_Value1) >= CLong(mpv_Value2) Then
        If CLong(mpv_Value1) >= CLong(mpv_Value3) Then
            MaxValue = CLong(mpv_Value1)
        Else
            MaxValue = CLong(mpv_Value3)
        End If
    Else
        If CLong(mpv_Value2) >= CLong(mpv_Value3) Then
            MaxValue = CLong(mpv_Value2)
        Else
            MaxValue = CLong(mpv_Value3)
        End If
    End If

    Exit Function
    
Error_MaxValue:
    MaxValue = 0
    
End Function


'※ 그룹 History를 위해 적용해야 한다.
Public Function GrpConditionCheck(tGrpMst As GrpMstRec, sPrmOdrDte As String, lPrmAge As Long) As Boolean

    Dim sBufAnti As String
    
    GrpConditionCheck = False
    
        ' sBufAnti에 적용구분을 반대로 setting
    If lPrmAge >= 8 Then
        sBufAnti = "C"                           ' 소아 (8세 미만)
    Else
        sBufAnti = "A"                           ' 성인
    End If
    

    If tGrpMst.GrpAdpTyp <> sBufAnti Then
        If sPrmOdrDte >= "20010101" Then
            '예전에 발생한 코드에 대해서는 모두 적용하자.
            If Trim(tGrpMst.GrpAdpDte) = "" And Trim(tGrpMst.GrpExpDte) = "" Then
                GrpConditionCheck = True
            End If
            
            '기간이 만료되었는지 확인하자
            If sPrmOdrDte >= tGrpMst.GrpAdpDte And sPrmOdrDte <= tGrpMst.GrpExpDte Then
                GrpConditionCheck = True
            End If
        Else
            GrpConditionCheck = True
        End If
    End If
    
    
End Function

Public Sub Grid_Clear(Obj As Object, Optional fCol As Integer, Optional eCol As Integer)

    With Obj
        If fCol > 0 Then
            .Row = 1
            .Row2 = .MaxRows
            .col = fCol
            .Col2 = eCol
            .BlockMode = True
            .Action = ActionClear
            .BlockMode = False
        Else
            .Row = -1
            .col = -1
            .Action = ActionClearText
        End If
        
    End With

End Sub

Public Sub Clear_SpreadSheet(mpc_Form As Form, Optional mpc_Control As Control = Nothing, Optional mpc_Graphic As Boolean = False)

    Dim i As Integer
    Dim mbo_Obj As Control
    
    On Error GoTo Clear_SpreadSheet_ERR:
    
    
    If mpc_Control Is Nothing Then
        For Each mbo_Obj In mpc_Form.Controls
        
            If TypeName(mpc_Control) = "fpSpread" Or TypeOf mpc_Control Is vaSpread Then
                If mbo_Obj.Tag = "Column0" Then
                    mbo_Obj.col = 0
                Else
                    mbo_Obj.col = 1
                End If
                mbo_Obj.Col2 = mbo_Obj.MaxCols
                mbo_Obj.Row = 1
                mbo_Obj.Row2 = mbo_Obj.MaxRows
                mbo_Obj.BlockMode = True
                mbo_Obj.Action = ActionClearText  'SPD_ACTION_CLEAR_TEXT
                mbo_Obj.BlockMode = False
                If mpc_Graphic Then
                    For i = 1 To mbo_Obj.MaxCols
                        mbo_Obj.col = i
                        If mbo_Obj.CellType = 9 Then        'Picture Type
                            mbo_Obj.Col2 = i
                            mbo_Obj.Row = 1
                            mbo_Obj.Row2 = mbo_Obj.MaxRows
                            mbo_Obj.BlockMode = True
                            mbo_Obj.Action = ActionClear  'SPD_ACTION_CLEAR
                            mbo_Obj.BlockMode = False
                        End If
                    Next
                End If
            End If
        Next
    Else
        'If TypeOf mpc_Control Is vaSpread Then
        If TypeName(mpc_Control) = "fpSpread" Or TypeOf mpc_Control Is vaSpread Then
            With mpc_Control
                
                If .Tag = "Column0" Then
                    .col = 0
                Else
                    .col = 1
                End If
                .Col2 = .MaxCols
                .Row = 1
                .Row2 = .MaxRows
                .BlockMode = True
                .Action = ActionClearText  'SPD_ACTION_CLEAR_TEXT
                .BlockMode = False
                If mpc_Graphic Then
                    For i = 1 To .MaxCols
                        .col = i
                        If .CellType = 9 Then        'Picture Type
                            .Col2 = i
                            .Row = 1
                            .Row2 = .MaxRows
                            .BlockMode = True
                            .Action = ActionClear  'SPD_ACTION_CLEAR
                            .BlockMode = False
                        End If
                    Next
                End If
            End With
        End If
    End If
    
Clear_SpreadSheet_ERR:
    
End Sub

'마우스가 위치하는 곳을 활성화 시킨다.
Public Sub Spread_Get_SpdColRow(mpo_Spd As Object, mpl_TwipsX As Single, mpl_TwipsY As Single, Optional mpl_Row As Long, Optional mpl_Col As Long)
        
    Dim mbl_GetCol As Long
    Dim mbl_GetRow As Long
    Dim mbl_RowHi As Long       '스프레드 화면상의 마지막 Row
    Dim mbl_RowHi_1 As Long     '위 Column - 1 Row
    Dim mbl_ColWid As Long      '스프레드 화면상의 마지막 Column
    Dim mbl_ColWid_1 As Long    '위 Column - 1 Column
    Dim mbl_Top As Long
    Dim mbl_Right As Long
    Dim mbl_Bottom As Long
        
    Dim i As Long, j As Long, k As Long, L As Long, m As Long, n As Long
    
    With mpo_Spd
    
        '스프레드에 표시사항이 없으면
        If .MaxRows = 0 Then Exit Sub
                   
        .SetFocus
    
       '현 마우스 위치에서 스프레드의 Col , Row 를 구한다.
        
        'Row, Col Position Check
        If Not IsMissing(mpl_Row) Then
            Call .GetCellFromScreenCoord(mpl_Col, mpl_Row, CLng(mpl_TwipsX), CLng(mpl_TwipsY))
            Exit Sub
        End If

        mpl_TwipsX = mpl_TwipsX / Screen.TwipsPerPixelX
        mpl_TwipsY = mpl_TwipsY / Screen.TwipsPerPixelY
    
        mgl_Result = .GetCellFromScreenCoord(mbl_GetCol, mbl_GetRow, CLng(mpl_TwipsX), CLng(mpl_TwipsY))
       

       '스프레드 화면상의 상단 Row 값을 구한다.
        mbl_Top = .TopRow
        
       '현재 마우스가 있는 위치
        For mbl_GetCol = mbl_GetCol To .MaxCols
            .col = mbl_GetCol
            If .ColHidden = False Then Exit For
        Next
    
        .Row = mbl_GetRow
            
        If mbl_GetCol < 1 Or mbl_GetCol > .MaxCols Then Exit Sub
        If mbl_GetRow < 1 Or mbl_GetRow > .MaxRows Then Exit Sub
             
       'VisibleRows, VisibleRows 범위를 벋어나면 Box 를 옮기지 않는다.
        If .VisibleRows < ((mbl_GetRow - mbl_Top) + 1) Then .Row = mbl_GetRow - 1
        If .VisibleCols < ((mbl_GetCol - .LeftCol) + 1) Then .col = mbl_GetCol - 1
        
        mgu_SpdAbsCol = (.col - .LeftCol) + 1
        mgu_SpdAbsRow = (.Row - mbl_Top) + 1
        
        .Action = 0     ' Activate
        
       'SetCursorPos 에 의한 Mouse Move Event 발생이면 TypeHAlign
       '속성에 맞춰 마우스 포인터를 조정 한다.
        If mgu_CursPosChnged = True Then
            Call Move_SpdMousePointer(mpo_Spd, mgu_SpdAbsCol, mgu_SpdAbsRow)
            mgu_CursPosChnged = False
        End If
        
    End With
    
End Sub

'해당 Window를 가장 위에 보여준다!
Public Function Display_Top_Most(mpl_Handle As Long, mpl_Left As Long, mpl_Top As Long, mpl_Width As Long, mpl_Height As Long) As Long

    Display_Top_Most = SetWindowPos(mpl_Handle, -1, mpl_Left, mpl_Top, mpl_Width, mpl_Height, &H43)
    
End Function

'가장 위에 보여준 Window를 다시 표준으로..
Public Function Display_Top_Most_Cancel(mpl_Handle As Long, mpl_Left As Long, mpl_Top As Long, mpl_Width As Long, mpl_Height As Long) As Long

    Display_Top_Most_Cancel = SetWindowPos(mpl_Handle, -2, mpl_Left, mpl_Top, mpl_Width, mpl_Height, &H43)
    
End Function

Public Function Get_App_Caption() As String

    Get_App_Caption = App.Title & " " & App.EXEName & " Ver " & App.Major & "." & App.Minor & " (Build : " & App.Revision & ")"
    
End Function

Public Sub frmAlignRightBottom(prmFrm As Form, lprmMgn As Long)
    
    prmFrm.Top = Screen.Height - prmFrm.Height - lprmMgn
    prmFrm.Left = Screen.Width - prmFrm.Width - lprmMgn

End Sub

Public Function GetWeekday(sDate As String, Optional bIsKor As Boolean = True) As String

    Dim sTempDate As String
    Dim sWeek As String
    
    sTempDate = Format(Left(sDate, 8), "####/##/##")
    
    If bIsKor Then
        Select Case Weekday(sTempDate)
        Case vbSunday
            sWeek = "일"
        Case vbMonday
            sWeek = "월"
        Case vbTuesday
            sWeek = "화"
        Case vbWednesday
            sWeek = "수"
        Case vbThursday
            sWeek = "목"
        Case vbFriday
            sWeek = "금"
        Case vbSaturday
            sWeek = "토"
        End Select
    Else
        Select Case Weekday(sTempDate)
        Case vbSunday
            sWeek = "Sun"
        Case vbMonday
            sWeek = "Mon"
        Case vbTuesday
            sWeek = "Tue"
        Case vbWednesday
            sWeek = "Wen"
        Case vbThursday
            sWeek = "Thu"
        Case vbFriday
            sWeek = "Fri"
        Case vbSaturday
            sWeek = "Sat"
        End Select
    End If
    
    GetWeekday = sWeek
    
End Function

Public Function Spread_FindColumn(spdObj As vaSpread, ByVal txt As String) As Integer



    
    Dim i As Integer

    Spread_FindColumn = 0
    txt = UCase$(txt)

    ' Search fields first

    With spdObj
        .Row = 0
        
        For i = 1 To spdObj.MaxCols
            .col = i
            If txt = UCase$(.Value) Then
                Spread_FindColumn = i
                Exit Function
            End If
        Next
    
        ' Now search for headings
    
        For i = 1 To spdObj.MaxCols
            .col = i
            If txt = UCase$(.Value) Then
                Spread_FindColumn = i
                Exit Function
            End If
        Next
    End With
    
End Function

Public Sub Spread_FindData(oSpread As Object, psSearchData As String, Row As Long, col As Long)

    Dim lRow As Long
    Dim lCol As Long
    
    With oSpread
        For lRow = 1 To .MaxRows
            For lCol = 1 To .MaxCols
                .Row = lRow
                .col = lCol
                If UCase(.Value) = UCase(psSearchData) Then
                    Row = lRow
                    col = lCol
                    Exit Sub
                End If
            Next
        Next
    End With
    
End Sub

Public Function TgFindColumn(spdObj As vaSpread, ByVal txt As String) As Integer

    Dim iCnt As Integer
    
    On Error GoTo err:

    With spdObj
    .Row = 0
    For iCnt = 1 To .MaxCols
        .col = iCnt
        If .Value = txt Then
            TgFindColumn = iCnt
        End If
    Next
    End With
    
err:

End Function


Public Function GetIPAddress() As String

    GetIPAddress = mvbFrm.Winsock.LocalIP
    
'    Dim IPAddrAndPort As String
'    Dim IPAddr As String
'    Dim i As Integer

'    mvbFrm.Mvb1.Code = "s P0=$$IP^%PPSS($J)"
'    'mvbFrm.Mvb1.Code = "s P0=^%CDServer(""alive"",$ZU(110),$J,""clientid"")"
'    mvbFrm.Mvb1.ExecFlag = 1
'
'    IPAddrAndPort = mvbFrm.Mvb1.P0
'
'    IPAddr = ""
'    For i = 1 To 4
'        IPAddr = IPAddr & piece(IPAddrAndPort, ".", i)
'        If (i <> 4) Then
'            IPAddr = IPAddr & "."
'       End If
'    Next i
'
'    GetIPAddress = IPAddr

    
End Function

Public Function GetCodeName(sDBName As String, sCode As String, iPos As Integer) As String
'--------------------------------------------------------
'다른 모듈에서 GetCodeName이 존재하면 모두 삭제할것...
'이제부턴 Common.bas에 있는 GetCodeName함수만 사용할 것임
'--------------------------------------------------------

'// 각종 코드 명칭을 읽어서 들고 온다.
'// sDbName : Global Name
'// sCode   : Key Code
'// iPos    : Name Position
'

    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sRetValue As String
    
    sCmpKey = ""
    sCurKey = sCode & Chr(5)

    sCurKey = mSetReadEqual(sDBName, sCurKey, sRetValue)
    If sCurKey = "" Then Exit Function
    GetCodeName = piece(sRetValue, Chr(5), iPos)
    
End Function

Public Sub SpdSet(spd As vaSpread, ByRef iRow As Long, ByRef iCol As Integer, sVal As String)

    With spd
    
        .col = iCol
        .Row = iRow
        .Value = sVal
    
    End With


End Sub

Public Function SpdGet(spd As vaSpread, ByRef iRow As Long, ByRef iCol As Integer) As String

    With spd
    
        .col = iCol
        .Row = iRow
        SpdGet = .Value
    
    End With


End Function

Public Function ReturnOutComeDeptNam(ByVal sPrmChtNum As String, ByVal sPrmOcmDte As String, ByVal sPrmDepCod As String) As String

    Dim i As Integer
    Dim bSw As Boolean
    Dim sStr As String
    Dim oCol As Collection

    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sRetVal As String
    
    Dim OcmData As OcmInfRec
    Dim DepData As DepMstRec
    
    Set oCol = New Collection
    
    sCmpKey = sPrmChtNum & Chr(5) & sPrmOcmDte
    sCurKey = sCmpKey
    sCurKey = mSetNext("OcmInfChtDtm", sCurKey)
    Do
        sCurKey = mReadNext("OcmInfChtDtm", sCurKey, sCmpKey, sRetVal)
        If sCurKey = "" Then Exit Do
        
        Call OcmInfLoad(sRetVal, OcmData)
    
        If OcmData.OcmComStt <> "OC" Then
        
            bSw = True
            For i = 1 To oCol.Count
                If oCol.Item(i) = OcmData.OcmDepCod Then
                    bSw = False
                    Exit For
                End If
            Next
            
            If bSw Then
                oCol.Add OcmData.OcmDepCod, OcmData.OcmDepCod
                
                If OcmData.OcmDepCod <> sPrmDepCod Then
                    Call DepMstRead(OcmData.OcmDepCod, Left(OcmData.OcmAcpDtm, 8), DepData)
                    sStr = sStr & "  " & Trim(DepData.DepKorNam)
                End If
            End If
        End If
    Loop
            
    ReturnOutComeDeptNam = Trim(sStr)
    Set oCol = Nothing

End Function

Public Sub PAUSE(ByVal nSecond As Single)
   
   Dim t0 As Single
   t0 = Timer
   Do While Timer - t0 < nSecond
      Dim Dummy As Integer
      Dummy = DoEvents()
      ' if we cross midnight, back up one day
      If Timer < t0 Then
         t0 = t0 - CLng(24) * CLng(60) * CLng(60)
      End If
   Loop

End Sub

Public Function ZipCodeAction(sPrmTxt As String) As String

    ' Zip Master 에서 명칭 조회
    Dim ZipData As ZipMstRec
    Dim sBufValue As String
    Dim sBufKey As String
    Dim sBufSize As String * 40

    Dim sZipMstCurKey As String
    Dim sZipMstCmpKey As String
    Dim sZipRetVal As String

    ZipData.ZipCod = sPrmTxt
    
    Call ZipMstStore(sBufKey, sBufValue, ZipData)
    
    sZipMstCurKey = sBufKey
    sBufValue = mSetReadEqual("ZipMst", sZipMstCurKey, sZipRetVal)
    
    If sBufValue <> "" Then
        Call ZipMstLoad(sZipRetVal, ZipData)
        ZipCodeAction = ZipData.ZipLrgNam
    Else
        ZipCodeAction = ""
    End If

End Function

Public Sub ZipStrAction(sPrmTxt As String, sPrmStr As String)

    Dim iCount As Integer
    Dim sZipCurKey As String, sZipCmpKey As String, sZipRetVal As String
    Dim ZipData As ZipMstRec
    Dim sTmpRetCod As String
    
    '문자 즉, 동이름을 입력한경우의 처리
    iCount = 0
    
    sZipCurKey = Trim(sPrmTxt)
    sZipCmpKey = sZipCurKey
    sZipCurKey = mSetNext("ZipMstSml", sZipCurKey)
    Do
        sZipCurKey = mReadNext("ZipMstSml", sZipCurKey, sZipCmpKey, sZipRetVal)
        
        If sZipCurKey = "" Then Exit Do
        
        iCount = iCount + 1
        
        Call ZipMstLoad(sZipRetVal, ZipData)
    Loop
    
    If iCount = 0 Then
        sPrmTxt = ""
        sPrmStr = ""
    ElseIf iCount = 1 Then
        '오직 하나밖에 없다면 바로 우편번호를 가져오자
        sPrmTxt = ZipData.ZipCod
        sPrmStr = ZipData.ZipLrgNam
    Else
        '두개 이상이면 도움말을 보여주자
        '--------------------------------------
        'ZipMstSml이 Index에 있는 경우 D-3 사용
        sZipCurKey = Trim(sPrmTxt)
        sZipCmpKey = sZipCurKey
        sTmpRetCod = MasterHelp("ZipMstSml", sZipCurKey, sZipCmpKey, 1, 2, sZipRetVal)
        '--------------------------------------
        If sTmpRetCod <> "" Then
            sPrmTxt = sTmpRetCod
            sPrmStr = sZipRetVal
        Else
            sPrmTxt = ""
            sPrmStr = ""
        End If
    End If

End Sub

Public Function GetDataFromMsg(ByVal psSource As String, ByVal psStartChr As String, ByVal psEndChr As String, Optional ByVal piPosition As String = 1) As String
'-----------------------------------------------------------------------------------------------'
' 이 함수는 psSource의 String중에서 psStartChr 과 psEndChr사이의 String을 구해 주는 함수입니다.
' piPosition은 조건에 맞는 자료가 여러개일 경우 몇번째 자료를 전달받을 것인지를 정해주는 위치값.
' 작성일자 : 2002년 7월 3일
' 작성자   : 인화성
'-----------------------------------------------------------------------------------------------'

    Dim vData As Variant
    Dim sStr As String
    
    '초기화
    GetDataFromMsg = ""
    
    '해당 Data에 비교대상 string이 없으면 그냥 나간다.
    If InStr(psSource, psStartChr) = 0 Then Exit Function

    '해당 Data에 비교종료 String이 없으면 전체String을 전달한다.
    If InStr(psSource, psEndChr) = 0 Then
        GetDataFromMsg = psSource
        Exit Function
    End If
    
    vData = Split(psSource, psStartChr)
    
    '원하는 데이터의 위치에 자료가 없는경우는 Null을 전달한다.
    If UBound(vData) < piPosition Then
        GetDataFromMsg = ""
        Exit Function
    End If
    
    If piPosition < 1 Then piPosition = 1
    sStr = vData(piPosition)
    
    vData = Split(sStr, psEndChr)
    
    GetDataFromMsg = vData(0)

End Function

Public Sub IisInfRead(sIcmNum As String, sOcmSeq As String, sDupSeq As String, IisData As IisInfRec)
    
    Dim sCurKey As String
    Dim sRetVal As String
        
    sCurKey = Format(Trim(sIcmNum), "@@@@@@@@@@") & Chr(5) & Format(Trim(sOcmSeq), "@@") & Chr(5) & Format(Trim(sDupSeq), "@@") & Chr(5)
    sCurKey = mSetReadEqual("IisInf", sCurKey, sRetVal)
    IisInfLoad sRetVal, IisData

End Sub


Public Sub DefineSeeCode(sPrmSeeCod As String, sPrmDate As String, SeeData As SeeMstRec)
    
    Dim SeeHstData As SeeHstRec
    Dim sSeeMstCurKey As String, sSeeMstCmpKey As String, sSeeMstRetVal As String

    sSeeMstCmpKey = sPrmSeeCod & Chr(5)
    sSeeMstCurKey = sSeeMstCmpKey
    sSeeMstCurKey = mSetReadNext("SeeMst", sSeeMstCurKey, sSeeMstCmpKey, sSeeMstRetVal)
    
    If sSeeMstCurKey = "" Then
        Call SeeMstLoad(sSeeMstRetVal, SeeData)
        Exit Sub
    End If
        
    
    Call SeeMstLoad(sSeeMstRetVal, SeeData)
    
    '------------------------------------------------------------
    '- 적용일에 부합되면 히스토리를 읽을 필요 없이 Exit 한다.
    '------------------------------------------------------------
    If Left(sPrmDate, 8) >= Left((SeeData.SeeAdpDte), 8) And Left(sPrmDate, 8) <= Left((SeeData.SeeExpDte), 8) Then
        Exit Sub
    End If
'    Debug.Print SeeData.SeeElcCod
    If Trim(SeeData.SeeRelCod) <> "" Then
        MsgBox SeeData.SeeKorNam & "(" & SeeData.SeeOdrCod & ")는 " & _
               SeeData.SeeRelCod & "로 대체 되었습니다."
    End If
    
    '-------------------------------------------------------
    '- 적용일범위를 벗어나면 수가정보의 History를 읽는다.
    '-------------------------------------------------------
    sSeeMstCmpKey = sPrmSeeCod & Chr(5)
    sSeeMstCurKey = sSeeMstCmpKey & sPrmDate & Chr(5)
    sSeeMstCurKey = mSetPrev("SeeHst", sSeeMstCurKey)
    sSeeMstCurKey = mReadPrev("SeeHst", sSeeMstCurKey, sSeeMstCmpKey, sSeeMstRetVal)
            
    'Bug가 맞는데 일단은 그냥 둔다.
    'If sSeeMstCurKey = "" Then Exit Sub
    
    Call SeeHstLoad(sSeeMstRetVal, SeeHstData)
    Call SeeHstStore(sSeeMstCurKey, sSeeMstRetVal, SeeHstData)

    '------------------------------------------------------
    '- History도 적용일에 부합되는지 check한다(970918)
    '------------------------------------------------------
    If Left(sPrmDate, 8) >= Left((SeeHstData.SeeAdpDte), 8) And Left(sPrmDate, 8) <= Left((SeeHstData.SeeExpDte), 8) Then
        sSeeMstRetVal = sPrmSeeCod & Chr(5) & sSeeMstRetVal
        Call SeeMstLoad(sSeeMstRetVal, SeeData)
    Else
        sSeeMstRetVal = ""
        Call SeeMstLoad(sSeeMstRetVal, SeeData)
    End If

End Sub

Public Sub CutMstReadPrev(sCutGub As String, sCutCod As String, sAdpDte As String, CutData As CutMstRec)

    Dim sCurKey As String, sCmpKey As String, sRetVal As String

    sCmpKey = sCutGub & Chr(5) & sCutCod & Chr(5)
    sCurKey = sCmpKey & sAdpDte & Chr(5)
    sCurKey = mSetPrev("CutMst", sCurKey)
    sCurKey = mReadPrev("CutMst", sCurKey, sCmpKey, sRetVal)
    If sCurKey <> "" Then
        Call CutMstLoad(sRetVal, CutData)
        '적용기간이 아니면
        If CutData.CutAdpDte > sAdpDte And CutData.CutExpDte < sAdpDte Then
            CutMstLoad "", CutData
        End If
    Else
        CutMstLoad "", CutData
    End If

End Sub

Public Function IsInjection(sItmCod As String) As Integer
    
    '주사 여부
    If CInteger(Left(sItmCod, 2)) = 4 Then
        IsInjection = True
    Else
        IsInjection = False
    End If


End Function

Public Function IsMaterial(sItmCod As String) As Integer
    
    Dim iVal As Integer

    iVal = CInteger(Mid(sItmCod, 5, 1))
    '재료 여부
    If iVal = 0 Or iVal = 2 Then
        IsMaterial = True
    Else
        IsMaterial = False
    End If

End Function

Public Function IsMeal(sItmCod As String) As Integer
    
    '식대
    If Left(sItmCod, 4) = "0207" Then
        IsMeal = True
    Else
        IsMeal = False
    End If

End Function

Public Function IsMedication(sItmCod As String) As Integer
    
    '투약 여부
    If CInteger(Left(sItmCod, 2)) = 3 Then
        IsMedication = True
    Else
        IsMedication = False
    End If

End Function

Public Function IsPhysical(sItmCod As String) As Integer
    
    '물리치료 여부
    If CInteger(Left(sItmCod, 2)) = 6 Then
        IsPhysical = True
    Else
        IsPhysical = False
    End If


End Function

Public Function CheckVersion(ByVal psSourceFile As String, ByVal psDestFile As String) As Boolean
    
'2004/04/06 이상은 - 주석처리
'    Dim i As Integer
'    Dim vDestVer As Variant
'    Dim vSourceVer As Variant
'    Dim sDestVer As String
'    Dim sSourceVer As String
'
'    Dim sDestDtm As String
'    Dim sSourceDtm As String
'
'    Dim oFSO As FileSystemObject
'    Dim oFile As File
'
'    Set oFSO = New FileSystemObject
'
'    If oFSO.GetFileVersion(psDestFile) <> "" Then                   '20021018 이대순 수정 버전 값을 얻지못하는 것은 넘어가도록
'        vDestVer = Split(oFSO.GetFileVersion(psDestFile), ".")
'        sDestVer = vDestVer(0) & vDestVer(1) & vDestVer(2) & Format(vDestVer(3), "0#")
'
'        vSourceVer = Split(oFSO.GetFileVersion(psSourceFile), ".")
'        sSourceVer = vSourceVer(0) & vSourceVer(1) & vSourceVer(2) & Format(vSourceVer(3), "0#")
'
'        Set oFile = oFSO.GetFile(psDestFile)
'        sDestDtm = Format(oFile.DateLastModified, "yyyymmddhhmm")
'
'        Set oFile = oFSO.GetFile(psSourceFile)
'        sSourceDtm = Format(oFile.DateLastModified, "yyyymmddhhmm")
'
'        If CLng(sDestVer) <> CLng(sSourceVer) Or CDbl(sDestDtm) < CDbl(sSourceDtm) Then
'            CheckVersion = True
'        Else
'            CheckVersion = False
'        End If
'    Else
'        '파일 버젼을 알수 없는것은 무조건 복사한다.
'        CheckVersion = True
'    End If
'
'    Set oFSO = Nothing
'    Set oFile = Nothing

End Function


Public Function VersionCheck_CaseByCase(psFullPathExeName As String) As Boolean

'2004/04/06 이상은 - 주석처리
'    Dim sExeName As String
'    Dim sFullName As String
'    Dim sServerPath As String
'    Dim oFSO As FileSystemObject
'    Dim oFile As File
'
'On Error Resume Next
'
'    VersionCheck_CaseByCase = True
'    sServerPath = GetSetting("HNT.CNV", "Server", "Path")
'
'    If Dir(sServerPath, vbDirectory) = "" Then
'        VersionCheck_CaseByCase = False
'        MsgBox "서버경로가 정확하지 않읍니다. 메뉴의 환경설정을 확인하십시요."
'        Exit Function
'    End If
'
'    '실행파일의 이름과 경로를 읽는다.
'    Set oFSO = New FileSystemObject
'    Set oFile = oFSO.GetFile(psFullPathExeName)
'    sExeName = oFile.NAME
'    sFullName = oFile.Path
'
'    '해당 파일이 없다면 서버에서 복사해 온다.
'    If Dir(sFullName) = "" Then
'        ' Source --> destination
'        FileCopy sServerPath & "exe\" & sExeName, sFullName
'    Else
'        '---------------------------------------------------------------------------
'        '먼저 version을 check해서 기존 버전과 틀리면 client의 하드디스크로 복사한다.
'        '---------------------------------------------------------------------------
'        'If Check_Dll_OCX_Version_And_Copy(sServerPath & "exe\" & sExeName, sFullName) Then
'        If CheckVersion(sServerPath & "exe\" & sExeName, sFullName) Then
'            FileCopy sServerPath & "exe\" & sExeName, sFullName
'        End If
'    End If
'
'    Set oFile = Nothing
'    Set oFSO = Nothing

End Function

Public Sub AppShell(sExeName As String, Optional sCommand As String = "", Optional iWinStyle As Integer = 1)

    Dim i As Double
        
On Error GoTo err_1
    If VersionCheck_CaseByCase(sExeName) Then
        If sCommand = "" Then
            i = Shell(sExeName, iWinStyle)
        Else
            i = Shell(sExeName & " " & sCommand, iWinStyle)
        End If
    End If

    Exit Sub

err_1:
    MsgBox "파일을 찾을수 없읍니다."
    Exit Sub
    
End Sub

Public Sub FormUnloadAction(oPrmFrm As Form)
'------------------------------------------------------------------------------------------------
'20021029 lek Add
'모든 프로그램의 폼 Unlaod 부분에서 이 함수를 호출함. 현재 프로그램에서 로드된 모든 폼을 종료시킴
'mvbForm이 종료될때 DeleteTcpInf 함수를 호출.
' oPrmFrm :FormUnloadAction 함수를 호출한 main Form Name
'------------------------------------------------------------------------------------------------
    Dim oTmp As Form
    
    For Each oTmp In Forms
        If oTmp.Name <> oPrmFrm.Name Or oTmp.Name = "mvbFrm" Then
            Unload oTmp
            DoEvents
        End If
    Next
    
End Sub

Public Sub DefineFeeCode(sPrmSeeCod As String, sPrmDate As String, SeeData As SeeMstRec, FeeData As FeeMstRec)
    
    Dim sSeeMstCurKey As String, sSeeMstCmpKey As String, sSeeMstRetVal As String

    Call DefineSeeCode(sPrmSeeCod, sPrmDate, SeeData)

    If SeeData.SeeElcCod <> "" Then
        Call ReadFeeCode((SeeData.SeeElcCod), sPrmDate, FeeData)
        Exit Sub
    Else
        Call FeeMstLoad("", FeeData)
    End If

End Sub

'-------------------------------------------------------
'- 기간에 해당되는 금액을 수가정보에서 가져옴
'-------------------------------------------------------
Public Sub ReadFeeCode(sPrmCod As String, sPrmDate As String, FeeData As FeeMstRec)
    
    Dim FeeHstData As FeeHstRec
    Dim sFeeMstCurKey As String, sFeeMstCmpKey As String, sFeeMstRetVal As String

    '-----------------------------------------------------------------------------------
    '- 일단 최종내역(FeeMst)을 읽고 해당 기간이 아니면 History(FeeHst)를 읽는다
    '-----------------------------------------------------------------------------------
    sFeeMstCurKey = sPrmCod & Chr(5)
    sFeeMstCurKey = mSetReadEqual("FeeMst", sFeeMstCurKey, sFeeMstRetVal)
    
    Call FeeMstLoad(sFeeMstRetVal, FeeData)
    If Len(FeeData.FeeAdpDte) = 6 Then
        FeeData.FeeAdpDte = AddCenturyLen(FeeData.FeeAdpDte)
    End If

    If Left(sPrmDate, 8) >= FeeData.FeeAdpDte And Left(sPrmDate, 8) <= FeeData.FeeExpDte Then Exit Sub

    sFeeMstCmpKey = sPrmCod & Chr(5)
    sFeeMstCurKey = sFeeMstCmpKey & Left(sPrmDate, 8) & Chr(5)
    sFeeMstCurKey = mSetPrev("FeeHst", sFeeMstCurKey)
    sFeeMstCurKey = mReadPrev("FeeHst", sFeeMstCurKey, sFeeMstCmpKey, sFeeMstRetVal)
    
    Call FeeHstLoad(sFeeMstRetVal, FeeHstData)
    Call FeeHstStore(sFeeMstCurKey, sFeeMstRetVal, FeeHstData)
    
    If Left(sPrmDate, 8) >= FeeHstData.FeeAdpDte And Left(sPrmDate, 8) <= FeeHstData.FeeExpDte Then
        sFeeMstRetVal = sPrmCod & Chr(5) & sFeeMstRetVal
        Call FeeMstLoad(sFeeMstRetVal, FeeData)
    Else
        sFeeMstRetVal = ""
        Call FeeMstLoad(sFeeMstRetVal, FeeData)
    End If

End Sub

Public Function TextGeneratorSpaceSplit(ByVal psdata As String, ByVal piLinePerCharator As Integer) As String

'-------------------------------------------------------------------------------------------'
'- psData의 String을 piLinePerCharator만큼씩 잘라서 vbcrlf를 붙여주는 함수입니다.(Mars-Man)-'
'-------------------------------------------------------------------------------------------'
    Dim i     As Integer
    Dim sTmp  As String
    Dim vData As Variant
    Dim sVal  As String
    
    psdata = Replace(psdata, vbCrLf, " ")
    vData = Split(psdata, " ")
    
    For i = 0 To UBound(vData)
        If sTmp = "" Then
            sTmp = vData(i)
        Else
            sTmp = sTmp & " " & vData(i)
        End If
            
        If LenK(sTmp) > piLinePerCharator Then
            If sVal = "" Then
                sVal = sTmp
            Else
                sVal = sVal & vbCrLf & sTmp
            End If
            sTmp = ""
        End If
    Next
    
    If sVal = "" Then
        sVal = sTmp
    Else
        sVal = sVal & vbCrLf & sTmp
    End If
    
    TextGeneratorSpaceSplit = sVal
    
End Function

Public Function TextCountInstr(ByVal psdata As String, ByVal psSearchChar As String) As Integer

    Dim vData As Variant
    
    vData = Split(psdata, psSearchChar)
    
    TextCountInstr = UBound(vData) + 1
    
End Function

Public Function Text_GetTextBetweenSeparator(ByVal psText As String, ByVal psStartSeparator As String, ByVal psEndSeparator As String)

    '2003-07-02 marsman
    '두개의 분리자 사이의 값을 구함니다.

    Dim iStrPos As Integer
    Dim iEndPos As Integer
    Dim sStr    As String
    
    Text_GetTextBetweenSeparator = ""
    
    iStrPos = InStr(psText, psStartSeparator)
    If iStrPos < 1 Then Exit Function

    iEndPos = InStr(psText, psEndSeparator)
    If iEndPos < 1 Then Exit Function
    If iEndPos <= iStrPos Then Exit Function
    
    sStr = Mid(psText, iStrPos + 1, iEndPos - iStrPos - 1)
    Text_GetTextBetweenSeparator = sStr

End Function

Public Sub SetDTPicker(poDTP As Object, psDte As String)

    If psDte = "" Then Exit Sub
    
    '초기화
    poDTP.Year = "2003"
    poDTP.Month = "01"
    poDTP.Day = "01"
    
    '날자 설정.
    poDTP.Year = Left(psDte, 4)
    poDTP.Day = Right(psDte, 2)
    poDTP.Month = Format(Mid(psDte, 5, 2), "##")
    
    

End Sub

Public Sub GetDTPicker(poDTP As Object, psDte As String)

    psDte = poDTP.Year
    psDte = psDte & Format(poDTP.Month, "0#")
    psDte = psDte & Format(poDTP.Day, "0#")

End Sub

Public Sub GetFinalControlArrayIndex(poCtlArray As Object, piSmallIndex As Integer, piBigIndex As Integer)

    '컨트롤 배열의 최소 인덱스 와 최대 인덱스를 구한다.
    Dim iSmall As Integer
    Dim iBig As Integer
    Dim oCtl As Object

    iSmall = 999
    For Each oCtl In poCtlArray
        If oCtl.Index > iBig Then iBig = oCtl.Index
        
        If oCtl.Index < iSmall Then iSmall = oCtl.Index
    Next
    
    piSmallIndex = iSmall
    piBigIndex = iBig
    
End Sub

Public Function ControlArrayErrorTrap(poCtlArray As Object, piValue As Integer, piDefaultValue As Integer, Optional pbIfMAX As Boolean = True) As Integer

    Dim iSmall As Integer
    Dim iBig As Integer
    
    '전달받은 값이 컨트롤 배열에 있는 Index인지 체크하여 오류가 있다면, 전달받은 기본값을 전달해 준다.
    Call GetFinalControlArrayIndex(poCtlArray, iSmall, iBig)
    
    '최대값 비교라면..
    If pbIfMAX Then
        If piValue > iBig Then
            ControlArrayErrorTrap = piDefaultValue
        Else
            ControlArrayErrorTrap = piValue
        End If
        
    Else
    
        If piValue < iSmall Then
            ControlArrayErrorTrap = piDefaultValue
        Else
            ControlArrayErrorTrap = piValue
        End If

    End If

End Function

Public Sub GiveFullPower()

    bGblAppSecPowerUpdate = True
    bGblAppSecPowerRead = True

End Sub

Public Sub GetUsersPowerLimit(sUidCod As String)

    Dim sCurKey As String
    Dim sRetVal As String
    Dim Secdata As SecMstRec
    
    If sUidCod = "BIT" Then
        Call GiveFullPower
        Exit Sub
    End If
    
    sCurKey = sUidCod & Chr(5) & App.EXEName
    sCurKey = mSetReadEqual("SecMst", sCurKey, sRetVal)
    
    Call SecMstLoad(sRetVal, Secdata)
    
    If Secdata.SecAllPwr = "1" Then
        bGblAppSecPowerUpdate = True
        bGblAppSecPowerRead = True
    ElseIf Secdata.SecRedOny = "1" Then
        bGblAppSecPowerUpdate = False
        bGblAppSecPowerRead = True
    Else
        bGblAppSecPowerUpdate = False
        bGblAppSecPowerRead = False
    End If
    
End Sub

Function ReturnMeal(sPrmOcmNum As String, sPrmOdrDte As String) As String

    Dim sTmp As String
    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sRetVal As String
    
    Dim ImlData As ImlInfRec
    Dim MgdData As MgdMstRec
    
    sCmpKey = sPrmOcmNum & Chr(5)
    sCurKey = sCmpKey & "99999999" & Chr(5)
    sCurKey = mSetPrev("ImlInf", sCurKey)
    Do
        sCurKey = mReadPrev("ImlInf", sCurKey, sCmpKey, sRetVal)
        If sCurKey = "" Then Exit Do
        
        Call ImlInfLoad(sRetVal, ImlData)
        
        If CLong(sPrmOdrDte) >= CLong(ImlData.ImlAdpDte) And CLong(sPrmOdrDte) <= CLong(ImlData.ImlExpDte) Then
            
            Call MgdMstRead(ImlData.ImlBrfCod, sPrmOdrDte, MgdData)
            sTmp = MgdData.MgdNam
        
            Call MgdMstRead(ImlData.ImlLnhCod, sPrmOdrDte, MgdData)
            sTmp = sTmp & Chr(6) & MgdData.MgdNam
            
            Call MgdMstRead(ImlData.ImlDnrCod, sPrmOdrDte, MgdData)
            sTmp = sTmp & Chr(6) & MgdData.MgdNam
            
            '특이사항
            sTmp = sTmp & Chr(6) & MasterHelpDetail("DtlMst", "ADDCOD" & Chr(5) & ImlData.ImlSplCmt & Chr(5), "ADDCOD" & Chr(5) & ImlData.ImlSplCmt & Chr(5), 3)
            
            '변경사유
            sTmp = sTmp & Chr(6) & MasterHelpDetail("DtlMst", "WHYCOD" & Chr(5) & ImlData.ImlWhyCod & Chr(5), "WHYCOD" & Chr(5) & ImlData.ImlWhyCod & Chr(5), 3)
            
            '기타사유
            sTmp = sTmp & Chr(6) & ImlData.ImlEtcCmt
            
            ReturnMeal = sTmp
            Exit Function
        End If
    Loop

End Function

Public Function GetLogoPath() As String

    Dim sLogoPath As String
    
    sLogoPath = MasterHelpDetail("DtlMst", "LOGO" & Chr(5) & "LOGO" & Chr(5), "LOGO" & Chr(5) & "LOGO" & Chr(5), 3)
    
    If Dir(sLogoPath) = "" Then
        GetLogoPath = ""
    Else
        GetLogoPath = sLogoPath
    End If
    
    If Dir(sLogoPath) = "" Then
        GetLogoPath = ""
    End If
    
    
End Function

Public Function GetStringCount(ByVal psSource As String, ByVal psFindString As String) As Long
    
    Dim lCnt  As Long
    Dim vData As Variant
    
    If InStr(psSource, psFindString) = 0 Then
        lCnt = 0
    Else
        vData = Split(psSource, psFindString)
        
        lCnt = UBound(vData)
    End If

    GetStringCount = lCnt
    
End Function

Public Function ReCalcInjectionQty(ByVal psOdrQty As String, ByVal psOdrTms As String, ByVal psDivYon As String) As String

    Dim sTmpQty As String

    '분할여부에 따른... 처리...
    Select Case psDivYon
    Case "N"
        sTmpQty = CStr(CDouble(psOdrQty) / CDouble(psOdrTms))
        sTmpQty = CStr(CUp(sTmpQty))
        sTmpQty = CStr(CLong(sTmpQty) * CLong(psOdrTms))
    
    Case Else
        sTmpQty = psOdrQty
    
    End Select
    
    ReCalcInjectionQty = sTmpQty
    
End Function

Public Function GetEodQty(ByVal psOcmNum As String, ByVal psEodNum As String, ByVal psEodSeq As String) As String

End Function

Public Sub Check_Kg_Qty(ByVal psKG As String, ByVal psOdrCod As String, psReturnQty As String, psReturnDtrQtyYon As String)

    Dim sQty    As String
    Dim sDtrQty As String
    Dim sCurKey As String
    Dim sRetVal As String
    Dim KgoData As KgoMstRec
    
    sCurKey = psOdrCod & Chr(5)
    sCurKey = mSetReadEqual("KgoMst", sCurKey, sRetVal)
    If sCurKey = "" Or Trim(psKG) = "" Then
        sQty = "1"
        sDtrQty = "N"
    Else
        Call KgoMstLoad(sRetVal, KgoData)
            
        sQty = CRounding(CDouble(psKG) / CDouble(KgoData.KgoUntCod))
        sQty = CStr(sQty * CDouble(KgoData.KgoOdrQty))
        If sQty = "0" Then sQty = KgoData.KgoOdrQty
        sDtrQty = KgoData.KgoDtrQty
    End If
    
    psReturnQty = sQty
    psReturnDtrQtyYon = sDtrQty
        
End Sub

