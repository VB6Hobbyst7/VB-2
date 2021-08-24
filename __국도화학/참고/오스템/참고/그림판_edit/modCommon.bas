Attribute VB_Name = "modCommon"
'===============================================================================
'  프로그램 : 오스템 임플란트 모듈
'  파 일 명 : modCommon.bas
'  작 성 일 : 2011.09.21
'  작 성 자 : 오세원
'  홈페이지 : http://www.didiminfoinfo.co.kr
'  설    명 :
'  수정이력 :
'===============================================================================
Option Explicit

'==== 객체이동[MouseMove]관련 구조체
Public Type POINTAPI
    obj     As Variant
    fromx   As Long
    fromy   As Long
    x       As Long
    y       As Long
End Type

Public LMousePos   As POINTAPI     ' X,Y 좌표

'==== 인쇄관련 상수
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_PAINT = &HF
Public Const WM_PRINT = &H317


'==== 설정 Read/Wright [ostem.ini] 함스
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                                          ByVal lpKeyName As Any, ByVal lpDefault As String, _
                                                                                          ByVal lpReturnedString As String, ByVal nSize As Long, _
                                                                                          ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                                              ByVal lpKeyName As Any, ByVal lpString As Any, _
                                                                                              ByVal lplFileName As String) As Long

'==== 속성설정 구조체
Type Config
    Image   As String
    Layout  As String
    Logo    As String
    Scan    As String
    Work    As String
    Log     As String
End Type
Public gSetup As Config

'==== 경로속성 전역변수[CONFIG Set]
Public gImage   As String
Public gLayOut  As String
Public gLogo    As String
Public gScan    As String
Public gWork    As String
Public gLog     As String

'==== 경로속성 전역변수[MODE Set]
Public gScaleMode     As String
Public gScaleCal      As String
Public gDevide        As String

'==== 용지레이아웃 전역변수[LAYOUT Set]
Public gLayOutValue() As String
Public gLayOutUse     As String

'==== 메인메뉴 관련 상수
Public Const TLBKEY_NEW        As String = "NEW"
Public Const TLBKEY_OPEN       As String = "OPEN"
Public Const TLBKEY_SAVE       As String = "SAVE"
Public Const TLBKEY_MAKE       As String = "MAKE"
Public Const TLBKEY_VIEW       As String = "VIEW"
Public Const TLBKEY_EDIT       As String = "EDIT"
Public Const TLBKEY_EXIT       As String = "EXIT"

'==== LOF 파일열기 관련 상수
Public Const SEP       As String = "^"
Public Const CP_UTF8 = 65001

'==== LOF 파일읽기/쓰기 관련 함수
Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal codepage As Long, _
                                                             ByVal dwFlags As Long, _
                                                             ByVal lpMultiByteStr As Long, _
                                                             ByVal cchMultiByte As Long, _
                                                             ByVal lpWideCharStr As Long, _
                                                             ByVal cchWideChar As Long) As Long

Public Declare Function WideCharToMultiByteArray Lib "kernel32" Alias "WideCharToMultiByte" _
                                                            (ByVal codepage As Long, _
                                                             ByVal dwFlags As Long, _
                                                             ByRef lpWideCharStr As Byte, _
                                                             ByVal cchWideChar As Long, _
                                                             ByRef lpMultiByteStr As Byte, _
                                                             ByVal cchMultiByte As Long, _
                                                             ByVal lpDefaultChar As Long, _
                                                             ByVal lpUsedDefaultChar As Long) As Long

Public Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" _
                                                            (ByVal lpAppName As String, _
                                                             ByVal lpKeyName As String, _
                                                             ByVal lpDefault As String, _
                                                             ByVal lpReturnedString As String, _
                                                             ByVal nSize As Long) As Long

'==== LOF 파일읽기/쓰기 관련 Sub
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Byte, _
                                                                    Source As Byte, _
                                                                    ByVal Length As Long)

'==== 픽처박스 드래그드롭 값
Public DrawX   As Long
Public DrawY   As Long
Public Ot_X     As Long
Public Ot_Y     As Long

''Dim drageMode As Boolean

'==== 전체좌표 이동 인덱스 값 [0:Left, 1:Right, 2:Top, 3:Bottom]
Public intMoveIdx As Integer

'==== Mode Set [0:로드,1:적용,2:이동/크기조정,3:생성]
Public intMode As Integer

'==== 바코드 이미지명
Public strBarImgName As String

'==== 센치 to 트윕
Public Const CM_TOTWIP = 37.7952


'==== 레지스트리 키 ROOT 형식...
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
'Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

'==== 레지스트리 데이터 형식...
Public Const REG_NONE = 0                       ' No value type
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Public Const REG_BINARY = 3                     ' Free form binary
Public Const REG_DWORD = 4                      ' 32-bit number
Public Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Public Const REG_LINK = 6                       ' Symbolic Link (unicode)
Public Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings

'==== 반환값...
Public Const ERROR_NONE = 0
Public Const ERROR_BADKEY = 2
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_SUCCESS = 0


Global Const REG_POSITION   As String = "Software\VB and VBA Program Settings\DIDIM Info"
Global Const REG_USER_ID    As String = "USERID"
Global Const REG_PASSWD     As String = "PASSWD"
Global Const REG_PWD        As String = "20990101"
Global Const REG_UID        As String = "admin"

'---------------------------------------------------------------
'- 레지스트리 API 선언...
'---------------------------------------------------------------
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Private r           As Long
Private lValueType  As Long


'-- 세로출력관련
Public Const Pi = 3.14159265358979
Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 33
End Type

'Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
'Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long

'==== 컨트롤 명
Public gblCtrlNm    As String
Public gblCtrlIdx   As Integer

Private m_ColCommandButton              As Collection               ' 동적 생성 컨트롤 저장을 위한 컬렉션

Public ClsEventMonitor       As ClassEventMonitor        ' 이벤트 전달을 위한 클래스

''Public Sub RotateControl(ctl As Control, intAngle As Integer)
''    Dim lnghNewFont As Long
''    Dim lnghOriginalFonrt As Long
''    Dim lngHeight As Long
''    Dim lngWidth As Long
''    Dim ClsEventObject      As ClassEventObject
''    Dim obj             As Object
''
''
''    With frmLabelDesign.Picture1
''
''    Set ClsEventObject = New ClassEventObject
''
''
''        Set obj = ClsEventObject.CreateObject(frmLabelDesign, ClsEventMonitor, EventObjectSLabel, ctl.Name)
''
''        .ScaleMode = vbPixels
''        .AutoRedraw = True
''        lngHeight = .TextHeight(ctl)
''        lngWidth = 0
''
''
''        With .Font
''            lnghNewFont = CreateFont(lngHeight, lngWidth, intAngle * 10, intAngle * 10, .Weight, .Italic, .Underline, .Strikethrough, .Charset, 0, 0, 0, 0, .Name)
''        End With
''        ctl.Font = lnghNewFont
''
''        lnghOriginalFonrt = SelectObject(.hdc, lnghNewFont)
''        .CurrentX = ctl.Left
''        .CurrentY = ctl.Top
''        frmLabelDesign.Picture1.Print ctl
''
''        lnghNewFont = SelectObject(.hdc, lnghOriginalFonrt)
''        .AutoRedraw = False
''    End With
''    DeleteObject lnghNewFont
''    ctl.Visible = False
''
''    Set ctl.Container = frmLabelDesign.Picture1
''
'''    m_ColCommandButton.Add ClsEventObject
''
'''    Set ClsEventObject = Nothing
''
''End Sub

'obj======그리는곳
'X ====좌표
'Y ====좌표
'Txt======글자
'TxtGag===글자의 기울기
'H========글자의 높이(1에 대한 배율)
'W========글자의 너비(1에 대한 배율)
'LineSpace ====줄간격(1에 대한 배율)
'''
'''Public Sub FontStuff(obj As Object, X As Single, Y As Single, Txt As String, TxtGag As Integer, H As Single, W As Single, LineSpace As Single)
'''        On Error GoTo GetOut
'''        Dim F As LOGFONT, hPrevFont As Long, hFont As Long
'''        Dim str() As String
'''        Dim I As Long
'''
'''        '필요한건 알아서 입력요
'''        '================================
'''        Dim iFontName As String
'''        Dim iFontSize As Integer
'''
'''        iFontSize = 9
'''        iFontName = "굴림"
'''        '================================
'''
'''
'''        F.lfEscapement = 10 * Val(TxtGag) 'rotation angle, in tenths
'''        F.lfFacename = iFontName + Chr$(0)
'''        F.lfHeight = (iFontSize * -20) / 15
'''        F.lfWidth = (iFontSize * 10) / 15
'''
'''        F.lfHeight = F.lfHeight * H
'''        F.lfWidth = F.lfWidth * W
'''
'''
'''        hFont = CreateFontIndirect(F)
'''        hPrevFont = SelectObject(frmLabelDesign.Picture1.hdc, hFont)
'''
'''        str() = Split(Txt, Chr(13) & Chr(10)) '문자열을 줄단위로 자른다.
'''        For I = 0 To UBound(str)
'''                frmLabelDesign.Picture1.CurrentX = Cos((TxtGag * Pi / 180) - Pi / 2) * Abs(F.lfHeight * LineSpace) * I + X
'''                frmLabelDesign.Picture1.CurrentY = Sin((TxtGag * Pi / 180) + Pi / 2) * Abs(F.lfHeight * LineSpace) * I + Y
'''                frmLabelDesign.Picture1.Print str(I)
'''        Next I
'''
'''        hFont = SelectObject(frmLabelDesign.Picture1.hdc, hPrevFont)
'''        DeleteObject hFont
'''
'''
'''        Exit Sub
'''GetOut:
'''  Exit Sub
'''
'''End Sub

'레지스 트리에 문자열값 가져오기
Public Function GetString(hKey As Long, strPath As String, strValue As String)

    Dim keyhand As Long
    Dim DataType As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    
    r = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function

'-- 설정파일[ostem.ini] 읽어오기
Function GetSetup() As Boolean
Dim strFileName As String
Dim strReturnedString As String
Dim i As Integer
Dim intTotCnt As String
'Dim intUseCnt As String

    GetSetup = False
    strFileName = App.Path & "\ostem.ini"
    
    '=== [CONFIG Set] =========================================================================================
    strReturnedString = String(1024, " ")
    GetPrivateProfileString "CONFIG", "ImagePath", "", strReturnedString, Len(strReturnedString), strFileName
    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
    gImage = strReturnedString
    
    strReturnedString = String(1024, " ")
    GetPrivateProfileString "CONFIG", "LayoutPath", "", strReturnedString, Len(strReturnedString), strFileName
    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
    gLayOut = strReturnedString
        
    strReturnedString = String(1024, " ")
    GetPrivateProfileString "CONFIG", "LogoPath", "", strReturnedString, Len(strReturnedString), strFileName
    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
    gLogo = strReturnedString
        
    strReturnedString = String(1024, " ")
    GetPrivateProfileString "CONFIG", "ScanPath", "", strReturnedString, Len(strReturnedString), strFileName
    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
    gScan = strReturnedString
        
    strReturnedString = String(1024, " ")
    GetPrivateProfileString "CONFIG", "WorkPath", "", strReturnedString, Len(strReturnedString), strFileName
    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
    gWork = strReturnedString
        
    strReturnedString = String(1024, " ")
    GetPrivateProfileString "CONFIG", "LogPath", "", strReturnedString, Len(strReturnedString), strFileName
    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
    gLog = strReturnedString
    
    '=== [MODE Set] =========================================================================================
    strReturnedString = String(1024, " ")
    GetPrivateProfileString "MODE", "ScaleMode", "", strReturnedString, Len(strReturnedString), strFileName
    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
    gScaleMode = strReturnedString
    
    strReturnedString = String(1024, " ")
    GetPrivateProfileString "MODE", "ScaleCal", "", strReturnedString, Len(strReturnedString), strFileName
    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
    gScaleCal = strReturnedString
    
    strReturnedString = String(1024, " ")
    GetPrivateProfileString "MODE", "Devide", "", strReturnedString, Len(strReturnedString), strFileName
    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
    gDevide = strReturnedString
    
    '=== [LAYOUT Set] =========================================================================================
    strReturnedString = String(1024, " ")
    GetPrivateProfileString "LAYOUT", "Cnt", "", strReturnedString, Len(strReturnedString), strFileName
    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
    intTotCnt = strReturnedString
    
    ReDim Preserve gLayOutValue(intTotCnt) As String
    
    For i = 1 To intTotCnt
        strReturnedString = String(1024, " ")
        GetPrivateProfileString "LAYOUT", CStr(i), "", strReturnedString, Len(strReturnedString), strFileName
        strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
        gLayOutValue(i) = strReturnedString
    Next
    
    strReturnedString = String(1024, " ")
    GetPrivateProfileString "LAYOUT", "Use", "", strReturnedString, Len(strReturnedString), strFileName
    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
    gLayOutUse = strReturnedString
    
    GetSetup = True

End Function

'-- 설정파일[ostem.ini]에 쓰기
Function PutSetup(strIpKeyNm As String, strIpKey As String, strIpData As String) As Boolean
Dim strFileName As String
Dim strReturnedString As String

    PutSetup = False
    strFileName = App.Path & "\ostem.ini"
    
    strReturnedString = String(1024, " ")
    WritePrivateProfileString strIpKeyNm, strIpKey, strIpData, strFileName
    
    PutSetup = True

End Function

'-- 동적개체 클릭 이벤트
Public Sub obj_Click(obj As Object, objtype As Integer)
    Dim strImsiNm As String

    '-- Mode Set [적용가능]
    intMode = 1
        
    With frmPaint
        .sstType.Tab = objtype
        Select Case objtype
            Case 0
                .txtTitle.Text = obj.Tag
                .txtTag.Text = obj.Name
                
                .txtFontName(0).Text = obj.Font
                .txtFontSize(0).Text = Round(obj.FontSize / gDevide, 0)
                .chkFontBold(0).Value = IIf(obj.FontBold = True, "1", "0")
                .chkFontItalic(0).Value = IIf(obj.FontItalic = True, "1", "0")
                .chkFontUnder(0).Value = IIf(obj.FontUnderline = True, "1", "0")
                .txtYpos.Text = Round(obj.Top / gDevide, 0)
                .txtXpos.Text = Round(obj.Left / gDevide, 0)
                .txtContent(0).Text = obj.Caption
                .chkTStatic.Value = obj.DataMember                      '-- 무조건고정
                .chkPrint.Value = IIf(obj.DataField = "1", "0", "1")   '-- 출력안함
                                    
            Case 1
                .txtTitle.Text = obj.Tag
                .txtTag.Text = obj.Name
                
                .txtFontName(1).Text = obj.Font
                .txtFontSize(1).Text = Round(obj.FontSize / gDevide, 0)
                .chkFontBold(1).Value = IIf(obj.FontBold = True, "1", "0")
                .chkFontItalic(1).Value = IIf(obj.FontItalic = True, "1", "0")
                .chkFontUnder(1).Value = IIf(obj.FontUnderline = True, "1", "0")
                .txtYpos.Text = Round(obj.Top / gDevide, 0)
                .txtXpos.Text = Round(obj.Left / gDevide, 0)
                .txtContent(1).Text = obj.Caption
                .chkPrint.Value = IIf(obj.DataMember = "1", "0", "1")   '-- 출력안함
            Case 2
                .txtTitle.Text = obj.Tag
                .txtTag.Text = obj.Name
                
                .txtYpos.Text = Round(obj.Top / gDevide, 0)
                .txtXpos.Text = Round(obj.Left / gDevide, 0)
                .txtImageWSize(0).Text = Round(obj.Width / gDevide, 0)
                .txtImageHSize(0).Text = Round(obj.Height / gDevide, 0)
                .txtImageName(0).Text = obj.DataMember          '-- 이미지경로
                .chkIStatic.Value = obj.ToolTipText             '-- 무조건고정
                .chkPrint.Value = IIf(obj.DataField = "1", "0", "1")   '-- 출력안함
            
            Case 3
                .txtTitle.Text = obj.Tag
                .txtTag.Text = obj.Name
                .txtYpos.Text = Round(obj.Top / gDevide, 0)
                .txtXpos.Text = Round(obj.Left / gDevide, 0)
                .txtImageWSize(1).Text = Round(obj.Width / gDevide, 0)
                .txtImageHSize(1).Text = Round(obj.Height / gDevide, 0)
                .txtImageName(1).Text = obj.DataMember
                .chkPrint.Value = IIf(obj.DataField = "1", "0", "1")   '-- 출력안함
                
            Case 4
                .txtTitle.Text = obj.Tag
                .txtTag.Text = obj.Name
                '-- 이미지 컨트롤로 바코드를 대체하여 ToolTipText 에 바코드타입을 저장하여 사용한다.
                .cboBarType.ListIndex = obj.ToolTipText
                
                .txtYpos.Text = Round(obj.Top / gDevide, 0)
                .txtXpos.Text = Round(obj.Left / gDevide, 0)
                .txtBarWSize.Text = Round(obj.Width / gDevide, 0)
                .txtBarHSize.Text = Round(obj.Height / gDevide, 0)
                .txtYpos.Text = Round(obj.Top / gDevide, 0)
                .txtXpos.Text = Round(obj.Left / gDevide, 0)
                
                .chkPrint.Value = IIf(obj.DataField = "1", "0", "1")   '-- 출력안함
                
            Case 5
                .txtTitle.Text = obj.Tag
                .txtTag.Text = obj.Name
                
                .chkPrint.Value = IIf(obj.ToolTipText = "1", "0", "1")   '-- 출력안함
                If obj.DataMember = "0" Then                '-- Rotate
                    .txtLineWSize = Round(obj.Width / gDevide, 0)
                    .txtLineHSize = Round(obj.Height / gDevide, 0)
                    .txtYpos.Text = Round(obj.Top / gDevide, 0)
                    .txtXpos.Text = Round(obj.Left / gDevide, 0)
                    .chkLineRotate.Value = "0"
                Else
                    .txtLineHSize = Round(obj.Width / gDevide, 0)
                    .txtLineWSize = Round(obj.Height / gDevide, 0)
                    .txtYpos.Text = Round(obj.Top / gDevide, 0)
                    .txtXpos.Text = Round(obj.Left / gDevide, 0)
                    .chkLineRotate.Value = "1"
                End If
        End Select
    End With
    
    Call frmPaint.cmdSet_Click

End Sub

'-- 동적개체 마우스다운 이벤트
Public Sub obj_MouseDown(obj As Object, Button As Integer, Shift As Integer, x As Single, y As Single)
    '-- Mode Set [적용가능]
    intMode = 1

    LMousePos.obj = obj.Name
    LMousePos.fromx = obj.Left
    LMousePos.fromy = obj.Top
    
    LMousePos.x = Round(x / 15, 0) 'pixel to twip
    LMousePos.y = Round(y / 15, 0) 'pixel to twip
        
    
End Sub

'-- 동적개체 마우스무브 이벤트
Public Sub obj_MouseMove(obj As Object, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim LPanPos As POINTAPI
    Dim i As Integer
    
    '-- Mode Set [적용가능]
    intMode = 1

    If Button = vbLeftButton Or Button = vbRightButton Then
        x = x / 15 'pixel to twip
        y = y / 15 'pixel to twip

        LPanPos.x = (obj.Left + x - LMousePos.x)
        LPanPos.y = (obj.Top + y - LMousePos.y)

        LPanPos.x = IIf(LPanPos.x < 0, 0, LPanPos.x)
        LPanPos.y = IIf(LPanPos.y < 0, 0, LPanPos.y)

        obj.Move LPanPos.x, LPanPos.y

        frmPaint.txtXpos.Text = LPanPos.x / gDevide
        frmPaint.txtYpos.Text = LPanPos.y / gDevide

        '-- X,Y 좌표 Spread 적용
        With frmPaint.spdList
            For i = 1 To .MaxRows
                .Row = i
                .Col = 29
                If Trim(.Text) = obj.Name Then
                    Call .SetText(4, i, frmPaint.txtXpos.Text)
                    Call .SetText(6, i, frmPaint.txtYpos.Text)
                    Exit For
                End If
            Next
        End With
        
    End If
 

End Sub

