Option Strict Off
Option Explicit On
Module modCommon
	'===============================================================================
	'  프로그램 : 오스템 임플란트 모듈
	'  파 일 명 : modCommon.bas
	'  작 성 일 : 2011.09.21
	'  작 성 자 : 오세원
	'  홈페이지 : http://www.didiminfoinfo.co.kr
	'  설    명 :
	'  수정이력 :
	'===============================================================================
	
	'==== 객체이동[MouseMove]관련 구조체
	Public Structure POINTAPI
		Dim obj As Object
		Dim fromx As Integer
		Dim fromy As Integer
		Dim x As Integer
		Dim y As Integer
	End Structure
	
	Public LMousePos As POINTAPI ' X,Y 좌표
	
	'==== 인쇄관련 상수
	'FIXIT: As Any는 Visual Basic .NET에서 지원되지 않습니다. 형식을 지정하여 사용하십시오.                            FixIT90210ae-R5608-H1984
	'UPGRADE_ISSUE: 매개 변수를 'As Any'로 선언할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Function SendMessage Lib "user32"  Alias "SendMessageA"(ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Any) As Integer
	Public Const WM_PAINT As Integer = &HF
	Public Const WM_PRINT As Integer = &H317
	
	
	'==== 설정 Read/Wright [ostem.ini] 함스
	'FIXIT: As Any는 Visual Basic .NET에서 지원되지 않습니다. 형식을 지정하여 사용하십시오.                            FixIT90210ae-R5608-H1984
	'UPGRADE_ISSUE: 매개 변수를 'As Any'로 선언할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Declare Function GetPrivateProfileString Lib "kernel32"  Alias "GetPrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	
	'FIXIT: As Any는 Visual Basic .NET에서 지원되지 않습니다. 형식을 지정하여 사용하십시오.                            FixIT90210ae-R5608-H1984
	'UPGRADE_ISSUE: 매개 변수를 'As Any'로 선언할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: 매개 변수를 'As Any'로 선언할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As VariantType, ByVal lpString As VariantType, ByVal lplFileName As String) As Integer
	
	'==== 속성설정 구조체
	Structure Config
		'FIXIT: Image property 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.                FixIT90210ae-R7593-R67265
		Dim Image As String
		Dim Layout As String
		Dim Logo As String
		Dim Scan As String
		Dim Work As String
		Dim Log As String
	End Structure
	Public gSetup As Config
	
	'==== 경로속성 전역변수[CONFIG Set]
	Public gImage As String
	Public gLayOut As String
	Public gLogo As String
	Public gScan As String
	Public gWork As String
	Public gLog As String
	
	'==== 경로속성 전역변수[MODE Set]
	Public gScaleMode As String
	Public gScaleCal As String
	Public gDevide As String
	Public gBojung As String
	
	'==== 용지레이아웃 전역변수[LAYOUT Set]
	Public gLayOutValue() As String
	Public gLayOutUse As String
	
	'==== 메인메뉴 관련 상수
	Public Const TLBKEY_NEW As String = "NEW"
	Public Const TLBKEY_OPEN As String = "OPEN"
	Public Const TLBKEY_SAVE As String = "SAVE"
	Public Const TLBKEY_MAKE As String = "MAKE"
	Public Const TLBKEY_VIEW As String = "VIEW"
	Public Const TLBKEY_EDIT As String = "EDIT"
	Public Const TLBKEY_EXIT As String = "EXIT"
	
	'==== LOF 파일열기 관련 상수
	Public Const SEP As String = "^"
	Public Const CP_UTF8 As Integer = 65001
	Public Const CP_ACP As Short = 0
	
	Public gOpenFileNm As String
	
	'==== LOF 파일읽기/쓰기 관련 함수
	''MultiByteToWideChar -> 멀티바이트에서 unicode로
	''WideCharToMultiByte -> unicode에서 멀티바이트로
	''위에 두 함수는 API에서 제공해주는 함수 입니다.
	''MSDN에 있다는데 선천적으로 영어에 이질감을 느끼는 저는 네이버에서 찾았데요 ㅠ
	''여튼 이 두놈 때문에 겨우 해결 할 수 있을 것 같군요 아직 테스트는 안해 봤지만..
	''멀티바이트에서 유니코드 변환 방법
	''  // sTime이란 ANSI 무낮열을 bstr이란 이름의 유니코드(BSTR타입) 변수로 변환
	''  char sTime[] = '유니코드 변환 예제';
	''  BSTR bstr;
	''  // sTime을 유니코드로 변환하기에 앞서 먼저 그것의 유니코드에서의 길이를 알아야 한다.
	''  int nLen = MultiByteToWideChar(CP_ACP, 0, sTime, lstrlen(sTime), NULL, NULL)
	''  // 얻어낸 길이만큼 메모리를 할당한다.
	''  bstr = SysAllocStringLen(NULL, nLen);
	''  // 이제 변환을 수행한다.
	''  MultiByteToWideChar(CP_ACP, 0, sTime, lstrlen(sTime), bstr, nLen);
	''
	''유니코드에서 멀티바이트로 변환 방법
	''   // newVal이란 BSTR 타입에 있는 유니코드 문자열을 sTime이라는 ANSI 문자열로 변환
	''   char sTime[128];
	''   WideCharToMultiByte(CP_ACP, 0, newVal, -1, sTime, 128, NULL, NULL);
	''[출처] MultiByteToWideChar // WideCharToMultiByte|작성자 인간비행
	
	
	
	Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal codepage As Integer, ByVal dwFlags As Integer, ByVal lpMultiByteStr As Integer, ByVal cchMultiByte As Integer, ByVal lpWideCharStr As Integer, ByVal cchWideChar As Integer) As Integer
	
	Public Declare Function WideCharToMultiByteArray Lib "kernel32"  Alias "WideCharToMultiByte"(ByVal codepage As Integer, ByVal dwFlags As Integer, ByRef lpWideCharStr As Byte, ByVal cchWideChar As Integer, ByRef lpMultiByteStr As Byte, ByVal cchMultiByte As Integer, ByVal lpDefaultChar As Integer, ByVal lpUsedDefaultChar As Integer) As Integer
	
	Public Declare Function GetProfileString Lib "kernel32"  Alias "GetProfileStringA"(ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer
	
	Public Declare Function TextOutW Lib "gdi32" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal lpString As Integer, ByVal nCount As Integer) As Integer
	
	
	'==== LOF 파일읽기/쓰기 관련 Sub
	Public Declare Sub CopyMemory Lib "kernel32"  Alias "RtlMoveMemory"(ByRef Destination As Byte, ByRef Source As Byte, ByVal Length As Integer)
	
	'==== 픽처박스 드래그드롭 값
	Public DrawX As Integer
	Public DrawY As Integer
	Public Ot_X As Integer
	Public Ot_Y As Integer
	
	''Dim drageMode As Boolean
	
	'==== 전체좌표 이동 인덱스 값 [0:Left, 1:Right, 2:Top, 3:Bottom]
	Public intMoveIdx As Short
	
	'==== Mode Set [0:로드,1:적용,2:이동/크기조정,3:생성]
	Public intMode As Short
	
	'==== 바코드 이미지명
	Public strBarImgName As String
	
	'==== 센치 to 트윕
	Public Const CM_TOTWIP As Double = 37.7952
	
	
	'==== 레지스트리 키 ROOT 형식...
	Public Const HKEY_CLASSES_ROOT As Integer = &H80000000
	Public Const HKEY_CURRENT_USER As Integer = &H80000001
	'Public Const HKEY_LOCAL_MACHINE = &H80000002
	Public Const HKEY_USERS As Integer = &H80000003
	Public Const HKEY_PERFORMANCE_DATA As Integer = &H80000004
	
	'==== 레지스트리 데이터 형식...
	Public Const REG_NONE As Short = 0 ' No value type
	Public Const REG_SZ As Short = 1 ' Unicode nul terminated string
	Public Const REG_EXPAND_SZ As Short = 2 ' Unicode nul terminated string
	Public Const REG_BINARY As Short = 3 ' Free form binary
	Public Const REG_DWORD As Short = 4 ' 32-bit number
	Public Const REG_DWORD_BIG_ENDIAN As Short = 5 ' 32-bit number
	Public Const REG_LINK As Short = 6 ' Symbolic Link (unicode)
	Public Const REG_MULTI_SZ As Short = 7 ' Multiple Unicode strings
	
	'==== 반환값...
	Public Const ERROR_NONE As Short = 0
	Public Const ERROR_BADKEY As Short = 2
	Public Const ERROR_ACCESS_DENIED As Short = 8
	Public Const ERROR_SUCCESS As Short = 0
	
	
	Public Const REG_POSITION As String = "Software\VB and VBA Program Settings\DIDIM Info"
	Public Const REG_USER_ID As String = "USERID"
	Public Const REG_PASSWD As String = "PASSWD"
	Public Const REG_PWD As String = "20990101"
	Public Const REG_UID As String = "admin"
	
	'---------------------------------------------------------------
	'- 레지스트리 API 선언...
	'---------------------------------------------------------------
	Public Declare Function RegCreateKey Lib "advapi32.dll"  Alias "RegCreateKeyA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByRef phkResult As Integer) As Integer
	Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Integer) As Integer
	Public Declare Function RegDeleteKey Lib "advapi32.dll"  Alias "RegDeleteKeyA"(ByVal hKey As Integer, ByVal lpSubKey As String) As Integer
	Public Declare Function RegDeleteValue Lib "advapi32.dll"  Alias "RegDeleteValueA"(ByVal hKey As Integer, ByVal lpValueName As String) As Integer
	Public Declare Function RegOpenKey Lib "advapi32.dll"  Alias "RegOpenKeyA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByRef phkResult As Integer) As Integer
	Public Declare Function RegOpenKeyEx Lib "advapi32.dll"  Alias "RegOpenKeyExA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
	Public Declare Function RegQueryValueEx Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As String, ByRef lpcbData As Integer) As Integer
	Public Declare Function RegSetValueEx Lib "advapi32.dll"  Alias "RegSetValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal Reserved As Integer, ByVal dwType As Integer, ByVal lpData As String, ByVal cbData As Integer) As Integer
	
	Private r As Integer
	Private lValueType As Integer
	
	
	
	
	'-- 세로출력관련
	Public Const Pi As Double = 3.14159265358979
	Public Structure LOGFONT
		Dim lfHeight As Integer
		Dim lfWidth As Integer
		Dim lfEscapement As Integer
		Dim lfOrientation As Integer
		Dim lfWeight As Integer
		Dim lfItalic As Byte
		Dim lfUnderline As Byte
		Dim lfStrikeOut As Byte
		Dim lfCharSet As Byte
		Dim lfOutPrecision As Byte
		Dim lfClipPrecision As Byte
		Dim lfQuality As Byte
		Dim lfPitchAndFamily As Byte
		'UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(33),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=33)> Public lfFaceName() As Char
	End Structure
	
	
	'Private Type LOGFONT
	'lfHeight As Long
	'lfWidth As Long
	'lfEscapement As Long
	'lfOrientation As Long
	'lfWeight As Long
	'lfItalic As Byte
	'lfUnderline As Byte
	'lfStrikeOut As Byte
	'lfCharSet As Byte
	'lfOutPrecision As Byte
	'lfClipPrecision As Byte
	'lfQuality As Byte
	'lfPitchAndFamily As Byte
	'End Type
	
	
	
	'Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
	'Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
	'Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
	Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer
	Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Integer) As Integer
	Public Declare Function CreateFont Lib "gdi32"  Alias "CreateFontA"(ByVal H As Integer, ByVal W As Integer, ByVal E As Integer, ByVal O As Integer, ByVal W As Integer, ByVal i As Integer, ByVal u As Integer, ByVal S As Integer, ByVal C As Integer, ByVal OP As Integer, ByVal CP As Integer, ByVal Q As Integer, ByVal PAF As Integer, ByVal F As String) As Integer
	
	'==== 컨트롤 명
	Public gblCtrlNm As String
	Public gblCtrlIdx As Short
	
	Private m_ColCommandButton As Collection ' 동적 생성 컨트롤 저장을 위한 컬렉션
	
	Public ClsEventMonitor As ClassEventMonitor ' 이벤트 전달을 위한 클래스
	
	
	
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
	'FIXIT: 'GetString'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                         FixIT90210ae-R1672-R1B8ZE
	Public Function GetString(ByRef hKey As Integer, ByRef strPath As String, ByRef strValue As String) As Object
		
		Dim keyhand As Integer
		Dim DataType As Integer
		Dim lResult As Integer
		Dim strBuf As String
		Dim lDataBufSize As Integer
		Dim intZeroPos As Short
		
		r = RegOpenKey(hKey, strPath, keyhand)
		lResult = RegQueryValueEx(keyhand, strValue, 0, lValueType, CStr(0), lDataBufSize)
		If lValueType = REG_SZ Then
			'FIXIT: 'String' 함수를 'String$' 함수로 바꾸십시오.                                                  FixIT90210ae-R9757-R1B8ZE
			strBuf = New String(" ", lDataBufSize)
			lResult = RegQueryValueEx(keyhand, strValue, 0, 0, strBuf, lDataBufSize)
			If lResult = ERROR_SUCCESS Then
				intZeroPos = InStr(strBuf, Chr(0))
				If intZeroPos > 0 Then
					GetString = Left(strBuf, intZeroPos - 1)
				Else
					'UPGRADE_WARNING: GetString 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					GetString = strBuf
				End If
			End If
		End If
	End Function
	
	'-- 설정파일[ostem.ini] 읽어오기
	Function GetSetup() As Boolean
		Dim strFileName As String
		Dim strReturnedString As String
		Dim i As Short
		Dim intTotCnt As String
		'Dim intUseCnt As String
		
		GetSetup = False
		strFileName = My.Application.Info.DirectoryPath & "\ostem.ini"
		
		'=== [CONFIG Set] =========================================================================================
		'FIXIT: 'String' 함수를 'String$' 함수로 바꾸십시오.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("CONFIG", "ImagePath", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gImage = strReturnedString
		
		'FIXIT: 'String' 함수를 'String$' 함수로 바꾸십시오.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("CONFIG", "LayoutPath", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gLayOut = strReturnedString
		
		'FIXIT: 'String' 함수를 'String$' 함수로 바꾸십시오.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("CONFIG", "LogoPath", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gLogo = strReturnedString
		
		'FIXIT: 'String' 함수를 'String$' 함수로 바꾸십시오.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("CONFIG", "ScanPath", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gScan = strReturnedString
		
		'FIXIT: 'String' 함수를 'String$' 함수로 바꾸십시오.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("CONFIG", "WorkPath", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gWork = strReturnedString
		
		'FIXIT: 'String' 함수를 'String$' 함수로 바꾸십시오.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("CONFIG", "LogPath", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gLog = strReturnedString
		
		'=== [MODE Set] =========================================================================================
		'FIXIT: 'String' 함수를 'String$' 함수로 바꾸십시오.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("MODE", "ScaleMode", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gScaleMode = strReturnedString
		
		'FIXIT: 'String' 함수를 'String$' 함수로 바꾸십시오.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("MODE", "ScaleCal", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gScaleCal = strReturnedString
		
		'FIXIT: 'String' 함수를 'String$' 함수로 바꾸십시오.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("MODE", "Devide", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gDevide = strReturnedString
		
		'-- 인쇄보정값
		'FIXIT: 'String' 함수를 'String$' 함수로 바꾸십시오.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("MODE", "Bojung", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gBojung = strReturnedString
		
		'=== [LAYOUT Set] =========================================================================================
		'FIXIT: 'String' 함수를 'String$' 함수로 바꾸십시오.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("LAYOUT", "Cnt", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		intTotCnt = strReturnedString
		
		ReDim Preserve gLayOutValue(intTotCnt)
		
		For i = 1 To CInt(intTotCnt)
			'FIXIT: 'String' 함수를 'String$' 함수로 바꾸십시오.                                                  FixIT90210ae-R9757-R1B8ZE
			strReturnedString = New String(" ", 1024)
			GetPrivateProfileString("LAYOUT", CStr(i), "", strReturnedString, Len(strReturnedString), strFileName)
			'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
			strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
			gLayOutValue(i) = strReturnedString
		Next 
		
		'FIXIT: 'String' 함수를 'String$' 함수로 바꾸십시오.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("LAYOUT", "Use", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gLayOutUse = strReturnedString
		
		GetSetup = True
		
	End Function
	
	'-- 설정파일[ostem.ini]에 쓰기
	Function PutSetup(ByRef strIpKeyNm As String, ByRef strIpKey As String, ByRef strIpData As String) As Boolean
		Dim strFileName As String
		Dim strReturnedString As String
		
		PutSetup = False
		strFileName = My.Application.Info.DirectoryPath & "\ostem.ini"
		
		'FIXIT: 'String' 함수를 'String$' 함수로 바꾸십시오.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		WritePrivateProfileString(strIpKeyNm, strIpKey, strIpData, strFileName)
		
		PutSetup = True
		
	End Function
	
	'-- 동적개체 클릭 이벤트
	'FIXIT: 'obj'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                               FixIT90210ae-R1672-R1B8ZE
	Public Sub obj_Click(ByRef obj As Object, ByRef objtype As Short)
		Dim strImsiNm As String
		
		'-- Mode Set [적용가능]
		intMode = 1
		
		With frmLabelDesign
			.sstType.SelectedIndex = objtype
			Select Case objtype
				Case 0
					'UPGRADE_WARNING: obj.Tag 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTitle.Text = obj.Tag
					'UPGRADE_WARNING: obj.Name 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTag.Text = obj.Name
					
					'UPGRADE_WARNING: obj.Font 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtFontName(0).Text = obj.Font
					'UPGRADE_WARNING: obj.FontSize 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtFontSize(0).Text = CStr(System.Math.Round(obj.FontSize / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.FontBold 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkFontBold(0).CheckState = IIf(obj.FontBold = True, "1", "0")
					'UPGRADE_WARNING: obj.FontItalic 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkFontItalic(0).CheckState = IIf(obj.FontItalic = True, "1", "0")
					'UPGRADE_WARNING: obj.FontUnderline 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkFontUnder(0).CheckState = IIf(obj.FontUnderline = True, "1", "0")
					'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtYpos.Text = CStr(System.Math.Round(obj.Top / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtXpos.Text = CStr(System.Math.Round(obj.Left / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtContent(0).Text = obj.Caption
					'.txtContent1.Text = obj.Caption
					'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkTStatic.CheckState = obj.DataMember '-- 무조건고정
					'UPGRADE_WARNING: obj.DataField 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkPrint.CheckState = IIf(obj.DataField = "1", "0", "1") '-- 출력안함
					
				Case 1
					'UPGRADE_WARNING: obj.Tag 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTitle.Text = obj.Tag
					'UPGRADE_WARNING: obj.Name 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTag.Text = obj.Name
					
					'UPGRADE_WARNING: obj.Font 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtFontName(1).Text = obj.Font
					'UPGRADE_WARNING: obj.FontSize 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtFontSize(1).Text = CStr(System.Math.Round(obj.FontSize / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.FontBold 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkFontBold(1).CheckState = IIf(obj.FontBold = True, "1", "0")
					'UPGRADE_WARNING: obj.FontItalic 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkFontItalic(1).CheckState = IIf(obj.FontItalic = True, "1", "0")
					'UPGRADE_WARNING: obj.FontUnderline 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkFontUnder(1).CheckState = IIf(obj.FontUnderline = True, "1", "0")
					'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtYpos.Text = CStr(System.Math.Round(obj.Top / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtXpos.Text = CStr(System.Math.Round(obj.Left / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtContent(1).Text = obj.Caption
					'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkPrint.CheckState = IIf(obj.DataMember = "1", "0", "1") '-- 출력안함
				Case 2
					'UPGRADE_WARNING: obj.Tag 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTitle.Text = obj.Tag
					'UPGRADE_WARNING: obj.Name 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTag.Text = obj.Name
					
					'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtYpos.Text = CStr(System.Math.Round(obj.Top / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtXpos.Text = CStr(System.Math.Round(obj.Left / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtImageWSize(0).Text = CStr(System.Math.Round(obj.Width / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtImageHSize(0).Text = CStr(System.Math.Round(obj.Height / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtImageName(0).Text = obj.DataMember '-- 이미지경로
					'UPGRADE_ISSUE: Object 속성 obj.ToolTipText이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_WARNING: obj.ToolTipText 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkIStatic.CheckState = obj.ToolTipText '-- 무조건고정
					'UPGRADE_WARNING: obj.DataField 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkPrint.CheckState = IIf(obj.DataField = "1", "0", "1") '-- 출력안함
					
				Case 3
					'UPGRADE_WARNING: obj.Tag 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTitle.Text = obj.Tag
					'UPGRADE_WARNING: obj.Name 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTag.Text = obj.Name
					'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtYpos.Text = CStr(System.Math.Round(obj.Top / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtXpos.Text = CStr(System.Math.Round(obj.Left / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtImageWSize(1).Text = CStr(System.Math.Round(obj.Width / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtImageHSize(1).Text = CStr(System.Math.Round(obj.Height / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtImageName(1).Text = obj.DataMember
					'UPGRADE_WARNING: obj.DataField 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkPrint.CheckState = IIf(obj.DataField = "1", "0", "1") '-- 출력안함
					
				Case 4
					'UPGRADE_WARNING: obj.Tag 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTitle.Text = obj.Tag
					'UPGRADE_WARNING: obj.Name 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTag.Text = obj.Name
					'-- 이미지 컨트롤로 바코드를 대체하여 ToolTipText 에 바코드타입을 저장하여 사용한다.
					'UPGRADE_ISSUE: Object 속성 obj.ToolTipText이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_WARNING: obj.ToolTipText 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.cboBarType.SelectedIndex = obj.ToolTipText
					
					'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtYpos.Text = CStr(System.Math.Round(obj.Top / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtXpos.Text = CStr(System.Math.Round(obj.Left / CDbl(gDevide), 0))
					
					If frmLabelDesign.chkBarRotate.CheckState = CDbl("0") Then
						'UPGRADE_WARNING: obj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtBarWSize.Text = CStr(System.Math.Round(obj.Width / CDbl(gDevide), 0))
						'UPGRADE_WARNING: obj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtBarHSize.Text = CStr(System.Math.Round(obj.Height / CDbl(gDevide), 0))
					Else
						'UPGRADE_WARNING: obj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtBarHSize.Text = CStr(System.Math.Round(obj.Width / CDbl(gDevide), 0))
						'UPGRADE_WARNING: obj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtBarWSize.Text = CStr(System.Math.Round(obj.Height / CDbl(gDevide), 0))
					End If
					
					'UPGRADE_WARNING: obj.DataField 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkPrint.CheckState = IIf(obj.DataField = "1", "0", "1") '-- 출력안함
					
				Case 5
					'UPGRADE_WARNING: obj.Tag 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTitle.Text = obj.Tag
					'UPGRADE_WARNING: obj.Name 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTag.Text = obj.Name
					
					'UPGRADE_ISSUE: Object 속성 obj.ToolTipText이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_WARNING: obj.ToolTipText 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkPrint.CheckState = IIf(obj.ToolTipText = "1", "0", "1") '-- 출력안함
					'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If obj.DataMember = "0" Then '-- Rotate
						'UPGRADE_WARNING: obj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtLineWSize.Text = CStr(System.Math.Round(obj.Width / CDbl(gDevide), 0))
						'UPGRADE_WARNING: obj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtLineHSize.Text = CStr(System.Math.Round(obj.Height / CDbl(gDevide), 0))
						'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtYpos.Text = CStr(System.Math.Round(obj.Top / CDbl(gDevide), 0))
						'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtXpos.Text = CStr(System.Math.Round(obj.Left / CDbl(gDevide), 0))
						.chkLineRotate.CheckState = CShort("0")
					Else
						'UPGRADE_WARNING: obj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtLineHSize.Text = CStr(System.Math.Round(obj.Width / CDbl(gDevide), 0))
						'UPGRADE_WARNING: obj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtLineWSize.Text = CStr(System.Math.Round(obj.Height / CDbl(gDevide), 0))
						'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtYpos.Text = CStr(System.Math.Round(obj.Top / CDbl(gDevide), 0))
						'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtXpos.Text = CStr(System.Math.Round(obj.Left / CDbl(gDevide), 0))
						.chkLineRotate.CheckState = CShort("1")
					End If
			End Select
		End With
		
		Call frmLabelDesign.cmdSet_Click(Nothing, New System.EventArgs())
		
	End Sub
	
	'-- 동적개체 마우스다운 이벤트
	'FIXIT: 'obj'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                               FixIT90210ae-R1672-R1B8ZE
	Public Sub obj_MouseDown(ByRef obj As Object, ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
		'-- Mode Set [적용가능]
		intMode = 1
		
		'UPGRADE_WARNING: obj.Name 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: LMousePos.obj 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		LMousePos.obj = obj.Name
		'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		LMousePos.fromx = obj.Left
		'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		LMousePos.fromy = obj.Top
		
		LMousePos.x = System.Math.Round(x / 15, 0) 'pixel to twip
		LMousePos.y = System.Math.Round(y / 15, 0) 'pixel to twip
		
		
	End Sub
	
	'-- 동적개체 마우스무브 이벤트
	'FIXIT: 'obj'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                               FixIT90210ae-R1672-R1B8ZE
	Public Sub obj_MouseMove(ByRef obj As Object, ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
		Dim LPanPos As POINTAPI
		Dim i As Short
		
		'-- Mode Set [적용가능]
		intMode = 1
		
		If Button = VB6.MouseButtonConstants.LeftButton Or Button = VB6.MouseButtonConstants.RightButton Then
			x = x / 15 'pixel to twip
			y = y / 15 'pixel to twip
			
			'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LPanPos.x = (obj.Left + x - LMousePos.x)
			'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LPanPos.y = (obj.Top + y - LMousePos.y)
			
			LPanPos.x = IIf(LPanPos.x < 0, 0, LPanPos.x)
			LPanPos.y = IIf(LPanPos.y < 0, 0, LPanPos.y)
			
			'UPGRADE_WARNING: obj.Move 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			obj.Move(LPanPos.x, LPanPos.y)
			
			frmLabelDesign.txtXpos.Text = CStr(LPanPos.x / CDbl(gDevide))
			frmLabelDesign.txtYpos.Text = CStr(LPanPos.y / CDbl(gDevide))
			
			'-- X,Y 좌표 Spread 적용
			With frmLabelDesign.spdList
				For i = 1 To .MaxRows
					.Row = i
					.Col = 29
					'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
					'UPGRADE_WARNING: obj.Name 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Trim(.Text) = obj.Name Then
						Call .SetText(4, i, frmLabelDesign.txtXpos.Text)
						Call .SetText(6, i, frmLabelDesign.txtYpos.Text)
						Exit For
					End If
				Next 
			End With
			
		End If
		
		
	End Sub
End Module