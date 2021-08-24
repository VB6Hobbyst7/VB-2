Option Strict Off
Option Explicit On
Module modCommon
	'===============================================================================
	'  ���α׷� : ������ ���ö�Ʈ ���
	'  �� �� �� : modCommon.bas
	'  �� �� �� : 2011.09.21
	'  �� �� �� : ������
	'  Ȩ������ : http://www.didiminfoinfo.co.kr
	'  ��    �� :
	'  �����̷� :
	'===============================================================================
	
	'==== ��ü�̵�[MouseMove]���� ����ü
	Public Structure POINTAPI
		Dim obj As Object
		Dim fromx As Integer
		Dim fromy As Integer
		Dim x As Integer
		Dim y As Integer
	End Structure
	
	Public LMousePos As POINTAPI ' X,Y ��ǥ
	
	'==== �μ���� ���
	'FIXIT: As Any�� Visual Basic .NET���� �������� �ʽ��ϴ�. ������ �����Ͽ� ����Ͻʽÿ�.                            FixIT90210ae-R5608-H1984
	'UPGRADE_ISSUE: �Ű� ������ 'As Any'�� ������ �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Function SendMessage Lib "user32"  Alias "SendMessageA"(ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Any) As Integer
	Public Const WM_PAINT As Integer = &HF
	Public Const WM_PRINT As Integer = &H317
	
	
	'==== ���� Read/Wright [ostem.ini] �Խ�
	'FIXIT: As Any�� Visual Basic .NET���� �������� �ʽ��ϴ�. ������ �����Ͽ� ����Ͻʽÿ�.                            FixIT90210ae-R5608-H1984
	'UPGRADE_ISSUE: �Ű� ������ 'As Any'�� ������ �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Declare Function GetPrivateProfileString Lib "kernel32"  Alias "GetPrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	
	'FIXIT: As Any�� Visual Basic .NET���� �������� �ʽ��ϴ�. ������ �����Ͽ� ����Ͻʽÿ�.                            FixIT90210ae-R5608-H1984
	'UPGRADE_ISSUE: �Ű� ������ 'As Any'�� ������ �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: �Ű� ������ 'As Any'�� ������ �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As VariantType, ByVal lpString As VariantType, ByVal lplFileName As String) As Integer
	
	'==== �Ӽ����� ����ü
	Structure Config
		'FIXIT: Image property ��(��) Visual Basic .NET���� �ش�Ǵ� �׸��� �����Ƿ� ���׷��̵���� �ʽ��ϴ�.                FixIT90210ae-R7593-R67265
		Dim Image As String
		Dim Layout As String
		Dim Logo As String
		Dim Scan As String
		Dim Work As String
		Dim Log As String
	End Structure
	Public gSetup As Config
	
	'==== ��μӼ� ��������[CONFIG Set]
	Public gImage As String
	Public gLayOut As String
	Public gLogo As String
	Public gScan As String
	Public gWork As String
	Public gLog As String
	
	'==== ��μӼ� ��������[MODE Set]
	Public gScaleMode As String
	Public gScaleCal As String
	Public gDevide As String
	Public gBojung As String
	
	'==== �������̾ƿ� ��������[LAYOUT Set]
	Public gLayOutValue() As String
	Public gLayOutUse As String
	
	'==== ���θ޴� ���� ���
	Public Const TLBKEY_NEW As String = "NEW"
	Public Const TLBKEY_OPEN As String = "OPEN"
	Public Const TLBKEY_SAVE As String = "SAVE"
	Public Const TLBKEY_MAKE As String = "MAKE"
	Public Const TLBKEY_VIEW As String = "VIEW"
	Public Const TLBKEY_EDIT As String = "EDIT"
	Public Const TLBKEY_EXIT As String = "EXIT"
	
	'==== LOF ���Ͽ��� ���� ���
	Public Const SEP As String = "^"
	Public Const CP_UTF8 As Integer = 65001
	Public Const CP_ACP As Short = 0
	
	Public gOpenFileNm As String
	
	'==== LOF �����б�/���� ���� �Լ�
	''MultiByteToWideChar -> ��Ƽ����Ʈ���� unicode��
	''WideCharToMultiByte -> unicode���� ��Ƽ����Ʈ��
	''���� �� �Լ��� API���� �������ִ� �Լ� �Դϴ�.
	''MSDN�� �ִٴµ� ��õ������ ��� �������� ������ ���� ���̹����� ã�ҵ��� ��
	''��ư �� �γ� ������ �ܿ� �ذ� �� �� ���� �� ������ ���� �׽�Ʈ�� ���� ������..
	''��Ƽ����Ʈ���� �����ڵ� ��ȯ ���
	''  // sTime�̶� ANSI �������� bstr�̶� �̸��� �����ڵ�(BSTRŸ��) ������ ��ȯ
	''  char sTime[] = '�����ڵ� ��ȯ ����';
	''  BSTR bstr;
	''  // sTime�� �����ڵ�� ��ȯ�ϱ⿡ �ռ� ���� �װ��� �����ڵ忡���� ���̸� �˾ƾ� �Ѵ�.
	''  int nLen = MultiByteToWideChar(CP_ACP, 0, sTime, lstrlen(sTime), NULL, NULL)
	''  // �� ���̸�ŭ �޸𸮸� �Ҵ��Ѵ�.
	''  bstr = SysAllocStringLen(NULL, nLen);
	''  // ���� ��ȯ�� �����Ѵ�.
	''  MultiByteToWideChar(CP_ACP, 0, sTime, lstrlen(sTime), bstr, nLen);
	''
	''�����ڵ忡�� ��Ƽ����Ʈ�� ��ȯ ���
	''   // newVal�̶� BSTR Ÿ�Կ� �ִ� �����ڵ� ���ڿ��� sTime�̶�� ANSI ���ڿ��� ��ȯ
	''   char sTime[128];
	''   WideCharToMultiByte(CP_ACP, 0, newVal, -1, sTime, 128, NULL, NULL);
	''[��ó] MultiByteToWideChar // WideCharToMultiByte|�ۼ��� �ΰ�����
	
	
	
	Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal codepage As Integer, ByVal dwFlags As Integer, ByVal lpMultiByteStr As Integer, ByVal cchMultiByte As Integer, ByVal lpWideCharStr As Integer, ByVal cchWideChar As Integer) As Integer
	
	Public Declare Function WideCharToMultiByteArray Lib "kernel32"  Alias "WideCharToMultiByte"(ByVal codepage As Integer, ByVal dwFlags As Integer, ByRef lpWideCharStr As Byte, ByVal cchWideChar As Integer, ByRef lpMultiByteStr As Byte, ByVal cchMultiByte As Integer, ByVal lpDefaultChar As Integer, ByVal lpUsedDefaultChar As Integer) As Integer
	
	Public Declare Function GetProfileString Lib "kernel32"  Alias "GetProfileStringA"(ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer
	
	Public Declare Function TextOutW Lib "gdi32" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal lpString As Integer, ByVal nCount As Integer) As Integer
	
	
	'==== LOF �����б�/���� ���� Sub
	Public Declare Sub CopyMemory Lib "kernel32"  Alias "RtlMoveMemory"(ByRef Destination As Byte, ByRef Source As Byte, ByVal Length As Integer)
	
	'==== ��ó�ڽ� �巡�׵�� ��
	Public DrawX As Integer
	Public DrawY As Integer
	Public Ot_X As Integer
	Public Ot_Y As Integer
	
	''Dim drageMode As Boolean
	
	'==== ��ü��ǥ �̵� �ε��� �� [0:Left, 1:Right, 2:Top, 3:Bottom]
	Public intMoveIdx As Short
	
	'==== Mode Set [0:�ε�,1:����,2:�̵�/ũ������,3:����]
	Public intMode As Short
	
	'==== ���ڵ� �̹�����
	Public strBarImgName As String
	
	'==== ��ġ to Ʈ��
	Public Const CM_TOTWIP As Double = 37.7952
	
	
	'==== ������Ʈ�� Ű ROOT ����...
	Public Const HKEY_CLASSES_ROOT As Integer = &H80000000
	Public Const HKEY_CURRENT_USER As Integer = &H80000001
	'Public Const HKEY_LOCAL_MACHINE = &H80000002
	Public Const HKEY_USERS As Integer = &H80000003
	Public Const HKEY_PERFORMANCE_DATA As Integer = &H80000004
	
	'==== ������Ʈ�� ������ ����...
	Public Const REG_NONE As Short = 0 ' No value type
	Public Const REG_SZ As Short = 1 ' Unicode nul terminated string
	Public Const REG_EXPAND_SZ As Short = 2 ' Unicode nul terminated string
	Public Const REG_BINARY As Short = 3 ' Free form binary
	Public Const REG_DWORD As Short = 4 ' 32-bit number
	Public Const REG_DWORD_BIG_ENDIAN As Short = 5 ' 32-bit number
	Public Const REG_LINK As Short = 6 ' Symbolic Link (unicode)
	Public Const REG_MULTI_SZ As Short = 7 ' Multiple Unicode strings
	
	'==== ��ȯ��...
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
	'- ������Ʈ�� API ����...
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
	
	
	
	
	'-- ������°���
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
		'UPGRADE_WARNING: ���� ���� ���ڿ� ũ�Ⱑ ���ۿ� �¾ƾ� �մϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
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
	
	'==== ��Ʈ�� ��
	Public gblCtrlNm As String
	Public gblCtrlIdx As Short
	
	Private m_ColCommandButton As Collection ' ���� ���� ��Ʈ�� ������ ���� �÷���
	
	Public ClsEventMonitor As ClassEventMonitor ' �̺�Ʈ ������ ���� Ŭ����
	
	
	
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
	
	'obj======�׸��°�
	'X ====��ǥ
	'Y ====��ǥ
	'Txt======����
	'TxtGag===������ ����
	'H========������ ����(1�� ���� ����)
	'W========������ �ʺ�(1�� ���� ����)
	'LineSpace ====�ٰ���(1�� ���� ����)
	'''
	'''Public Sub FontStuff(obj As Object, X As Single, Y As Single, Txt As String, TxtGag As Integer, H As Single, W As Single, LineSpace As Single)
	'''        On Error GoTo GetOut
	'''        Dim F As LOGFONT, hPrevFont As Long, hFont As Long
	'''        Dim str() As String
	'''        Dim I As Long
	'''
	'''        '�ʿ��Ѱ� �˾Ƽ� �Է¿�
	'''        '================================
	'''        Dim iFontName As String
	'''        Dim iFontSize As Integer
	'''
	'''        iFontSize = 9
	'''        iFontName = "����"
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
	'''        str() = Split(Txt, Chr(13) & Chr(10)) '���ڿ��� �ٴ����� �ڸ���.
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
	
	'������ Ʈ���� ���ڿ��� ��������
	'FIXIT: 'GetString'��(��) �ʱ⿡ ���ε��Ǵ� ������ �������� �����Ͻʽÿ�.                                         FixIT90210ae-R1672-R1B8ZE
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
			'FIXIT: 'String' �Լ��� 'String$' �Լ��� �ٲٽʽÿ�.                                                  FixIT90210ae-R9757-R1B8ZE
			strBuf = New String(" ", lDataBufSize)
			lResult = RegQueryValueEx(keyhand, strValue, 0, 0, strBuf, lDataBufSize)
			If lResult = ERROR_SUCCESS Then
				intZeroPos = InStr(strBuf, Chr(0))
				If intZeroPos > 0 Then
					GetString = Left(strBuf, intZeroPos - 1)
				Else
					'UPGRADE_WARNING: GetString ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					GetString = strBuf
				End If
			End If
		End If
	End Function
	
	'-- ��������[ostem.ini] �о����
	Function GetSetup() As Boolean
		Dim strFileName As String
		Dim strReturnedString As String
		Dim i As Short
		Dim intTotCnt As String
		'Dim intUseCnt As String
		
		GetSetup = False
		strFileName = My.Application.Info.DirectoryPath & "\ostem.ini"
		
		'=== [CONFIG Set] =========================================================================================
		'FIXIT: 'String' �Լ��� 'String$' �Լ��� �ٲٽʽÿ�.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("CONFIG", "ImagePath", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gImage = strReturnedString
		
		'FIXIT: 'String' �Լ��� 'String$' �Լ��� �ٲٽʽÿ�.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("CONFIG", "LayoutPath", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gLayOut = strReturnedString
		
		'FIXIT: 'String' �Լ��� 'String$' �Լ��� �ٲٽʽÿ�.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("CONFIG", "LogoPath", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gLogo = strReturnedString
		
		'FIXIT: 'String' �Լ��� 'String$' �Լ��� �ٲٽʽÿ�.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("CONFIG", "ScanPath", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gScan = strReturnedString
		
		'FIXIT: 'String' �Լ��� 'String$' �Լ��� �ٲٽʽÿ�.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("CONFIG", "WorkPath", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gWork = strReturnedString
		
		'FIXIT: 'String' �Լ��� 'String$' �Լ��� �ٲٽʽÿ�.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("CONFIG", "LogPath", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gLog = strReturnedString
		
		'=== [MODE Set] =========================================================================================
		'FIXIT: 'String' �Լ��� 'String$' �Լ��� �ٲٽʽÿ�.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("MODE", "ScaleMode", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gScaleMode = strReturnedString
		
		'FIXIT: 'String' �Լ��� 'String$' �Լ��� �ٲٽʽÿ�.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("MODE", "ScaleCal", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gScaleCal = strReturnedString
		
		'FIXIT: 'String' �Լ��� 'String$' �Լ��� �ٲٽʽÿ�.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("MODE", "Devide", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gDevide = strReturnedString
		
		'-- �μ⺸����
		'FIXIT: 'String' �Լ��� 'String$' �Լ��� �ٲٽʽÿ�.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("MODE", "Bojung", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gBojung = strReturnedString
		
		'=== [LAYOUT Set] =========================================================================================
		'FIXIT: 'String' �Լ��� 'String$' �Լ��� �ٲٽʽÿ�.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("LAYOUT", "Cnt", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		intTotCnt = strReturnedString
		
		ReDim Preserve gLayOutValue(intTotCnt)
		
		For i = 1 To CInt(intTotCnt)
			'FIXIT: 'String' �Լ��� 'String$' �Լ��� �ٲٽʽÿ�.                                                  FixIT90210ae-R9757-R1B8ZE
			strReturnedString = New String(" ", 1024)
			GetPrivateProfileString("LAYOUT", CStr(i), "", strReturnedString, Len(strReturnedString), strFileName)
			'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
			strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
			gLayOutValue(i) = strReturnedString
		Next 
		
		'FIXIT: 'String' �Լ��� 'String$' �Լ��� �ٲٽʽÿ�.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		GetPrivateProfileString("LAYOUT", "Use", "", strReturnedString, Len(strReturnedString), strFileName)
		'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
		strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, CompareMethod.Binary))
		gLayOutUse = strReturnedString
		
		GetSetup = True
		
	End Function
	
	'-- ��������[ostem.ini]�� ����
	Function PutSetup(ByRef strIpKeyNm As String, ByRef strIpKey As String, ByRef strIpData As String) As Boolean
		Dim strFileName As String
		Dim strReturnedString As String
		
		PutSetup = False
		strFileName = My.Application.Info.DirectoryPath & "\ostem.ini"
		
		'FIXIT: 'String' �Լ��� 'String$' �Լ��� �ٲٽʽÿ�.                                                  FixIT90210ae-R9757-R1B8ZE
		strReturnedString = New String(" ", 1024)
		WritePrivateProfileString(strIpKeyNm, strIpKey, strIpData, strFileName)
		
		PutSetup = True
		
	End Function
	
	'-- ������ü Ŭ�� �̺�Ʈ
	'FIXIT: 'obj'��(��) �ʱ⿡ ���ε��Ǵ� ������ �������� �����Ͻʽÿ�.                                               FixIT90210ae-R1672-R1B8ZE
	Public Sub obj_Click(ByRef obj As Object, ByRef objtype As Short)
		Dim strImsiNm As String
		
		'-- Mode Set [���밡��]
		intMode = 1
		
		With frmLabelDesign
			.sstType.SelectedIndex = objtype
			Select Case objtype
				Case 0
					'UPGRADE_WARNING: obj.Tag ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTitle.Text = obj.Tag
					'UPGRADE_WARNING: obj.Name ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTag.Text = obj.Name
					
					'UPGRADE_WARNING: obj.Font ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtFontName(0).Text = obj.Font
					'UPGRADE_WARNING: obj.FontSize ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtFontSize(0).Text = CStr(System.Math.Round(obj.FontSize / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.FontBold ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkFontBold(0).CheckState = IIf(obj.FontBold = True, "1", "0")
					'UPGRADE_WARNING: obj.FontItalic ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkFontItalic(0).CheckState = IIf(obj.FontItalic = True, "1", "0")
					'UPGRADE_WARNING: obj.FontUnderline ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkFontUnder(0).CheckState = IIf(obj.FontUnderline = True, "1", "0")
					'UPGRADE_WARNING: obj.Top ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtYpos.Text = CStr(System.Math.Round(obj.Top / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Left ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtXpos.Text = CStr(System.Math.Round(obj.Left / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Caption ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtContent(0).Text = obj.Caption
					'.txtContent1.Text = obj.Caption
					'UPGRADE_WARNING: obj.DataMember ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkTStatic.CheckState = obj.DataMember '-- �����ǰ���
					'UPGRADE_WARNING: obj.DataField ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkPrint.CheckState = IIf(obj.DataField = "1", "0", "1") '-- ��¾���
					
				Case 1
					'UPGRADE_WARNING: obj.Tag ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTitle.Text = obj.Tag
					'UPGRADE_WARNING: obj.Name ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTag.Text = obj.Name
					
					'UPGRADE_WARNING: obj.Font ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtFontName(1).Text = obj.Font
					'UPGRADE_WARNING: obj.FontSize ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtFontSize(1).Text = CStr(System.Math.Round(obj.FontSize / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.FontBold ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkFontBold(1).CheckState = IIf(obj.FontBold = True, "1", "0")
					'UPGRADE_WARNING: obj.FontItalic ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkFontItalic(1).CheckState = IIf(obj.FontItalic = True, "1", "0")
					'UPGRADE_WARNING: obj.FontUnderline ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkFontUnder(1).CheckState = IIf(obj.FontUnderline = True, "1", "0")
					'UPGRADE_WARNING: obj.Top ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtYpos.Text = CStr(System.Math.Round(obj.Top / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Left ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtXpos.Text = CStr(System.Math.Round(obj.Left / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Caption ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtContent(1).Text = obj.Caption
					'UPGRADE_WARNING: obj.DataMember ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkPrint.CheckState = IIf(obj.DataMember = "1", "0", "1") '-- ��¾���
				Case 2
					'UPGRADE_WARNING: obj.Tag ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTitle.Text = obj.Tag
					'UPGRADE_WARNING: obj.Name ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTag.Text = obj.Name
					
					'UPGRADE_WARNING: obj.Top ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtYpos.Text = CStr(System.Math.Round(obj.Top / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Left ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtXpos.Text = CStr(System.Math.Round(obj.Left / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Width ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtImageWSize(0).Text = CStr(System.Math.Round(obj.Width / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Height ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtImageHSize(0).Text = CStr(System.Math.Round(obj.Height / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.DataMember ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtImageName(0).Text = obj.DataMember '-- �̹������
					'UPGRADE_ISSUE: Object �Ӽ� obj.ToolTipText��(��) ���׷��̵���� �ʾҽ��ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_WARNING: obj.ToolTipText ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkIStatic.CheckState = obj.ToolTipText '-- �����ǰ���
					'UPGRADE_WARNING: obj.DataField ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkPrint.CheckState = IIf(obj.DataField = "1", "0", "1") '-- ��¾���
					
				Case 3
					'UPGRADE_WARNING: obj.Tag ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTitle.Text = obj.Tag
					'UPGRADE_WARNING: obj.Name ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTag.Text = obj.Name
					'UPGRADE_WARNING: obj.Top ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtYpos.Text = CStr(System.Math.Round(obj.Top / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Left ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtXpos.Text = CStr(System.Math.Round(obj.Left / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Width ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtImageWSize(1).Text = CStr(System.Math.Round(obj.Width / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Height ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtImageHSize(1).Text = CStr(System.Math.Round(obj.Height / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.DataMember ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtImageName(1).Text = obj.DataMember
					'UPGRADE_WARNING: obj.DataField ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkPrint.CheckState = IIf(obj.DataField = "1", "0", "1") '-- ��¾���
					
				Case 4
					'UPGRADE_WARNING: obj.Tag ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTitle.Text = obj.Tag
					'UPGRADE_WARNING: obj.Name ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTag.Text = obj.Name
					'-- �̹��� ��Ʈ�ѷ� ���ڵ带 ��ü�Ͽ� ToolTipText �� ���ڵ�Ÿ���� �����Ͽ� ����Ѵ�.
					'UPGRADE_ISSUE: Object �Ӽ� obj.ToolTipText��(��) ���׷��̵���� �ʾҽ��ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_WARNING: obj.ToolTipText ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.cboBarType.SelectedIndex = obj.ToolTipText
					
					'UPGRADE_WARNING: obj.Top ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtYpos.Text = CStr(System.Math.Round(obj.Top / CDbl(gDevide), 0))
					'UPGRADE_WARNING: obj.Left ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtXpos.Text = CStr(System.Math.Round(obj.Left / CDbl(gDevide), 0))
					
					If frmLabelDesign.chkBarRotate.CheckState = CDbl("0") Then
						'UPGRADE_WARNING: obj.Width ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtBarWSize.Text = CStr(System.Math.Round(obj.Width / CDbl(gDevide), 0))
						'UPGRADE_WARNING: obj.Height ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtBarHSize.Text = CStr(System.Math.Round(obj.Height / CDbl(gDevide), 0))
					Else
						'UPGRADE_WARNING: obj.Width ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtBarHSize.Text = CStr(System.Math.Round(obj.Width / CDbl(gDevide), 0))
						'UPGRADE_WARNING: obj.Height ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtBarWSize.Text = CStr(System.Math.Round(obj.Height / CDbl(gDevide), 0))
					End If
					
					'UPGRADE_WARNING: obj.DataField ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkPrint.CheckState = IIf(obj.DataField = "1", "0", "1") '-- ��¾���
					
				Case 5
					'UPGRADE_WARNING: obj.Tag ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTitle.Text = obj.Tag
					'UPGRADE_WARNING: obj.Name ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.txtTag.Text = obj.Name
					
					'UPGRADE_ISSUE: Object �Ӽ� obj.ToolTipText��(��) ���׷��̵���� �ʾҽ��ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_WARNING: obj.ToolTipText ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.chkPrint.CheckState = IIf(obj.ToolTipText = "1", "0", "1") '-- ��¾���
					'UPGRADE_WARNING: obj.DataMember ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If obj.DataMember = "0" Then '-- Rotate
						'UPGRADE_WARNING: obj.Width ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtLineWSize.Text = CStr(System.Math.Round(obj.Width / CDbl(gDevide), 0))
						'UPGRADE_WARNING: obj.Height ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtLineHSize.Text = CStr(System.Math.Round(obj.Height / CDbl(gDevide), 0))
						'UPGRADE_WARNING: obj.Top ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtYpos.Text = CStr(System.Math.Round(obj.Top / CDbl(gDevide), 0))
						'UPGRADE_WARNING: obj.Left ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtXpos.Text = CStr(System.Math.Round(obj.Left / CDbl(gDevide), 0))
						.chkLineRotate.CheckState = CShort("0")
					Else
						'UPGRADE_WARNING: obj.Width ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtLineHSize.Text = CStr(System.Math.Round(obj.Width / CDbl(gDevide), 0))
						'UPGRADE_WARNING: obj.Height ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtLineWSize.Text = CStr(System.Math.Round(obj.Height / CDbl(gDevide), 0))
						'UPGRADE_WARNING: obj.Top ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtYpos.Text = CStr(System.Math.Round(obj.Top / CDbl(gDevide), 0))
						'UPGRADE_WARNING: obj.Left ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.txtXpos.Text = CStr(System.Math.Round(obj.Left / CDbl(gDevide), 0))
						.chkLineRotate.CheckState = CShort("1")
					End If
			End Select
		End With
		
		Call frmLabelDesign.cmdSet_Click(Nothing, New System.EventArgs())
		
	End Sub
	
	'-- ������ü ���콺�ٿ� �̺�Ʈ
	'FIXIT: 'obj'��(��) �ʱ⿡ ���ε��Ǵ� ������ �������� �����Ͻʽÿ�.                                               FixIT90210ae-R1672-R1B8ZE
	Public Sub obj_MouseDown(ByRef obj As Object, ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
		'-- Mode Set [���밡��]
		intMode = 1
		
		'UPGRADE_WARNING: obj.Name ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: LMousePos.obj ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		LMousePos.obj = obj.Name
		'UPGRADE_WARNING: obj.Left ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		LMousePos.fromx = obj.Left
		'UPGRADE_WARNING: obj.Top ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		LMousePos.fromy = obj.Top
		
		LMousePos.x = System.Math.Round(x / 15, 0) 'pixel to twip
		LMousePos.y = System.Math.Round(y / 15, 0) 'pixel to twip
		
		
	End Sub
	
	'-- ������ü ���콺���� �̺�Ʈ
	'FIXIT: 'obj'��(��) �ʱ⿡ ���ε��Ǵ� ������ �������� �����Ͻʽÿ�.                                               FixIT90210ae-R1672-R1B8ZE
	Public Sub obj_MouseMove(ByRef obj As Object, ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
		Dim LPanPos As POINTAPI
		Dim i As Short
		
		'-- Mode Set [���밡��]
		intMode = 1
		
		If Button = VB6.MouseButtonConstants.LeftButton Or Button = VB6.MouseButtonConstants.RightButton Then
			x = x / 15 'pixel to twip
			y = y / 15 'pixel to twip
			
			'UPGRADE_WARNING: obj.Left ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LPanPos.x = (obj.Left + x - LMousePos.x)
			'UPGRADE_WARNING: obj.Top ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LPanPos.y = (obj.Top + y - LMousePos.y)
			
			LPanPos.x = IIf(LPanPos.x < 0, 0, LPanPos.x)
			LPanPos.y = IIf(LPanPos.y < 0, 0, LPanPos.y)
			
			'UPGRADE_WARNING: obj.Move ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			obj.Move(LPanPos.x, LPanPos.y)
			
			frmLabelDesign.txtXpos.Text = CStr(LPanPos.x / CDbl(gDevide))
			frmLabelDesign.txtYpos.Text = CStr(LPanPos.y / CDbl(gDevide))
			
			'-- X,Y ��ǥ Spread ����
			With frmLabelDesign.spdList
				For i = 1 To .MaxRows
					.Row = i
					.Col = 29
					'FIXIT: 'Trim' �Լ��� 'Trim$' �Լ��� �ٲٽʽÿ�.                                                      FixIT90210ae-R9757-R1B8ZE
					'UPGRADE_WARNING: obj.Name ��ü�� �⺻ �Ӽ��� Ȯ���� �� �����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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